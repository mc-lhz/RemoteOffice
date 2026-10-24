// 幂等补丁：让 x2t_helper.js 的 WASM 下载改为"在线优先（仅上游 chaxus
// Cloudflare Pages 镜像），失败回退本地相对路径"。
//
// 在线源走 XHR 而非 fetch：当前 Edge/Chromium 在 Worker 里对跨源 brotli
// 响应做 fetch 的 body 读取会永久挂起（header 正常、主线程正常），而
// XHR responseType=arraybuffer 完全正常——x2t 转换恰好在 Worker 中执行。
//
// 用法: node tools/patch_x2t_online.js --dist <dist目录>
// 幂等: 检测到 WASM_SOURCES 标记即跳过。
const fs = require('fs');
const path = require('path');

function parseArgs(argv) {
    const out = { dist: '' };
    for (let i = 0; i < argv.length; i++) {
        if (argv[i] === '--dist') {
            out.dist = argv[++i] || '';
        }
    }
    return out;
}

const args = parseArgs(process.argv.slice(2));
const distDir = path.resolve(args.dist || path.join(__dirname, '..', '..', '..', 'office-dist'));
const helperPath = path.join(distDir, 'sdkjs', 'common', 'wasm', 'x2t', 'x2t_helper.js');
if (!fs.existsSync(helperPath)) {
    console.error('x2t_helper.js not found: ' + helperPath);
    process.exit(1);
}

let text = fs.readFileSync(helperPath, 'utf8');
if (text.indexOf('WASM_SOURCES') >= 0) {
    console.log('already patched: ' + helperPath);
    process.exit(0);
}

// ① 在 WASM_PATH 常量后插入候选链
const pathAnchor = "        this.WASM_PATH = '../../../../sdkjs/common/wasm/x2t/x2t.wasm.br'\n";
if (text.indexOf(pathAnchor) < 0) {
    console.error('WASM_PATH anchor not found (dist version drift?)');
    process.exit(1);
}
const sourcesBlock = pathAnchor +
    "        // Online-first chain: the upstream Cloudflare Pages mirror serves the\n" +
    "        // brotli variant with ACAO:* (works even from plain-HTTP pages, and the\n" +
    "        // https hop makes the browser advertise br so the encoding negotiates\n" +
    "        // itself). Falls back to the local relative path, where office_bp does\n" +
    "        // its own Accept-Encoding negotiation (plain HTTP -> .gz variant).\n" +
    "        this.WASM_SOURCES = [\n" +
    "            'https://edit.chaxus.com/sdkjs/common/wasm/x2t/x2t.wasm.br',\n" +
    "            '../../../../sdkjs/common/wasm/x2t/x2t.wasm.br',\n" +
    "        ];\n";
text = text.replace(pathAnchor, sourcesBlock);

// ② fetchWasmResponse 改为逐源尝试，源列表替换单一 wasmPath
const oldHead = '    X2TConverter.prototype.fetchWasmResponse = function () {\n' +
    '        var wasmPath = this.WASM_PATH;\n' +
    '        var attempts = this.WASM_FETCH_ATTEMPTS;\n' +
    '        var backoffMs = this.WASM_FETCH_BACKOFF_MS;\n';
const newHead = '    X2TConverter.prototype.fetchWasmResponse = function () {\n' +
    '        var sources = (this.WASM_SOURCES && this.WASM_SOURCES.length)\n' +
    '            ? this.WASM_SOURCES.slice()\n' +
    '            : [this.WASM_PATH];\n' +
    '        var attempts = this.WASM_FETCH_ATTEMPTS;\n' +
    '        var backoffMs = this.WASM_FETCH_BACKOFF_MS;\n';
if (text.indexOf(oldHead) < 0) {
    console.error('fetchWasmResponse head anchor not found (dist version drift?)');
    process.exit(1);
}
text = text.replace(oldHead, newHead);

// ③ attempt 前插入 loadSource（在线 http(s) 源走 XHR，本地相对路径保持 fetch）
const attemptAnchor = '        var attempt = function (n) {';
if (text.indexOf(attemptAnchor) < 0) {
    console.error('attempt anchor not found (dist version drift?)');
    process.exit(1);
}
const loadSourceBlock =
    "        // Cross-origin fetch() inside a worker hangs on this brotli body in\n" +
    "        // current Edge/Chromium (headers arrive, arrayBuffer() never settles),\n" +
    "        // while XHR with responseType arraybuffer streams it fine. Route\n" +
    "        // absolute http(s) sources through XHR; keep native fetch for the\n" +
    "        // local relative path (preserves its caching semantics).\n" +
    "        var loadSource = function (wasmPath) {\n" +
    "            if (!/^https?:/i.test(wasmPath)) {\n" +
    "                return fetch(wasmPath);\n" +
    "            }\n" +
    "            return new Promise(function (resolve, reject) {\n" +
    "                var xhr = new XMLHttpRequest();\n" +
    "                xhr.open('GET', wasmPath);\n" +
    "                xhr.responseType = 'arraybuffer';\n" +
    "                xhr.timeout = 15000;\n" +
    "                xhr.onload = function () {\n" +
    "                    if (xhr.status >= 200 && xhr.status < 300) {\n" +
    "                        resolve(new Response(xhr.response, { status: xhr.status }));\n" +
    "                    } else {\n" +
    "                        reject(new Error(\"Failed to fetch x2t WASM at '\" + wasmPath + \"' (\" + xhr.status + ')'));\n" +
    "                    }\n" +
    "                };\n" +
    "                xhr.onerror = function () {\n" +
    "                    reject(new Error('WASM XHR network error: ' + wasmPath));\n" +
    "                };\n" +
    "                xhr.ontimeout = function () {\n" +
    "                    reject(new Error('WASM XHR timeout (15s): ' + wasmPath));\n" +
    "                };\n" +
    "                xhr.send();\n" +
    "            });\n" +
    "        };\n\n" +
    '        var attempt = function (wasmPath, n) {';
text = text.replace(attemptAnchor, loadSourceBlock);

// ④ fetch 换 loadSource、重试传参带 wasmPath
const fetchAnchor = '            return fetch(wasmPath).then(';
if (text.indexOf(fetchAnchor) < 0) {
    console.error('fetch anchor not found (dist version drift?)');
    process.exit(1);
}
text = text.replace(fetchAnchor, '            return loadSource(wasmPath).then(');
text = text.replace(
    '                    return attempt(n + 1);',
    '                    return attempt(wasmPath, n + 1);');

// ⑤ 尾部逐源回退
const tailAnchor = '        return attempt(1);\n    };';
if (text.indexOf(tailAnchor) < 0) {
    console.error('tail anchor not found (dist version drift?)');
    process.exit(1);
}
text = text.replace(tailAnchor,
    '        var trySource = function (index) {\n' +
    '            var wasmPath = sources[index];\n' +
    '            return attempt(wasmPath, 1).catch(function (error) {\n' +
    '                if (index + 1 >= sources.length) { throw error; }\n' +
    "                console.warn('[x2t] WASM source failed, falling back to the next source:', wasmPath);\n" +
    '                return trySource(index + 1);\n' +
    '            });\n' +
    '        };\n\n' +
    '        return trySource(0);\n    };');

fs.writeFileSync(helperPath, text);
console.log('patched: ' + helperPath);
