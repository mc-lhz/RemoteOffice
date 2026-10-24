# Office 插件（RemoteOffice）

RemoteConnecter 配套插件：基于 [ranuts/document](https://github.com/ranuts/document)（OnlyOffice 系）构建的浏览器端本地 Office 编辑器。
无需安装 Office、无需外部服务，docx / xlsx / pptx / pdf 等文档在浏览器内完成打开、编辑与保存，文件不离开本机。

## 支持格式

`.docx` `.doc` `.odt` `.rtf` `.txt` `.xlsx` `.xls` `.ods` `.csv` `.pptx` `.ppt` `.odp` `.pdf`（保存上限 200MB）

## 安装

1. 将本目录（`office/`）放入 RemoteConnecter 的插件目录，通常是主程序目录下的 `plugins/office/`。
2. 部署编辑器静态资源（见下节）。
3. 启动主程序后访问 `/office?path=<文件绝对路径>&lang=zh-CN`，主程序文件列表的"在线编辑"入口也会跳转到该地址。

无需修改主程序代码；插件通过 `plugin.json`（id: `office`）自动加载。

## 静态资源部署

编辑器资源目录（`office-dist`，约 530MB）不入 Git，按下述顺序探测（`office_bp.resolveDistDir`）：

1. 环境变量 `RC_OFFICE_DIST`
2. `%PROGRAMDATA%\RemoteConnecter\office\dist`
3. `plugins/office/office-dist`
4. 仓库根 `office-dist`

目录下需含 `editor.html`（ranuts/document 构建产物，含 `sdkjs/`、`web-apps/`、`dictionaries/` 等）。

### x2t 资产准备（重要）

`x2t.wasm.br` 默认以 `Content-Encoding: br` 下发，但 Chrome/Edge 只在安全上下文（HTTPS 或 localhost）接受 Brotli；
通过局域网 IP 走 HTTP 访问时会报 `ERR_CONTENT_DECODING_FAILED`。部署后请生成 gzip 变体：

```bash
node tools/prepare_office_assets.js            # 在默认 office-dist 下生成 x2t.wasm.gz（约 10MB）
node tools/prepare_office_assets.js --plain    # 额外生成未压缩 x2t.wasm（约 42MB，兜底用）
node tools/prepare_office_assets.js --dist "C:\ProgramData\RemoteConnecter\office\dist"
```

服务端按下述顺序响应 `*.br` 请求：同名 `.gz`（gzip）→ 去后缀原文件（无压缩）→ 原始 `.br`。
`GET /office/health` 的 `x2t.variant` 字段显示当前实际生效的变体（gzip / plain / brotli）。

### x2t 在线优先下载（可选）

对 dist 打补丁后，编辑器的 x2t WASM 会**优先从上游 Cloudflare Pages 镜像下载**（仅此一个在线源，
约 6.9MB brotli，`Access-Control-Allow-Origin: *`，从纯 HTTP 页面跨源拉取也合法），
失败自动回退本地相对路径（走上面的编码协商链）。公网/低上行带宽部署可省下约 10MB 的服务器出口流量，
机房断网不受影响：

```bash
node tools/patch_x2t_online.js --dist "C:\ProgramData\RemoteConnecter\office\dist"   # 幂等，可重复执行
```

在线源地址硬编码于补丁内（`https://edit.chaxus.com/sdkjs/common/wasm/x2t/x2t.wasm.br`）；
不需要在线加速时，不执行本补丁即保持纯本地行为。

## 路由一览

| 路由 | 说明 |
|---|---|
| `GET /office?path=...&lang=&theme=` | 编辑器页面（zh-CN / en-US，dark / light） |
| `GET /office/api/open?path=` | 读取文档（带 `X-Office-MTime` / `X-Office-Size` 头） |
| `POST /office/api/save` | 保存（multipart：path/mtime/file）；mtime 冲突返回 409 |
| `GET /office/api/download?path=` | 以附件形式下载当前文件 |
| `GET /office/health` | 就绪状态、资源目录、x2t 变体、词典信息 |
| `GET /sw.js` | no-op Service Worker（消除 404） |
| `GET /{sdkjs,web-apps,fonts,...}/...` | 编辑器静态资源别名（自动带 COOP/COEP/CORP 头） |

保存语义：原子替换（先写 `.rc-office.tmp`），替换前备份为 `<文件名>.rc-office.bak`，ETag/mtime 冲突检测防止覆盖他人改动。

## 访问方式与浏览器限制

| 访问方式 | Service Worker / COOP 隔离 | 编辑功能 |
|---|---|---|
| `http://localhost:8765` / `127.0.0.1` | 可用 | 完整 |
| `http://<局域网IP>:8765`（已部署 .gz） | 不可用（浏览器安全模型） | 完整（控制台会有 COOP 警告，可忽略） |
| 客户端 hosts 把 `office.localhost` 指向服务器 IP | 可用 | 完整 |
| HTTPS | 可用 | 完整 |

局域网多机场景推荐在客户端 hosts 添加 `<服务器IP> office.localhost` 后用 `http://office.localhost:8765` 访问，可零改动恢复全部能力。

## 测试

主仓库内可用 `run_office_test.py` 启动独立测试服务（默认 `0.0.0.0:8765`），冒烟检查：

```bash
curl http://127.0.0.1:8765/office/health
```
