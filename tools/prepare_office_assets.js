const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

function parseArgs(argv) {
    const out = { plain: false, dist: '' };
    for (let i = 0; i < argv.length; i++) {
        if (argv[i] === '--plain') {
            out.plain = true;
        } else if (argv[i] === '--dist') {
            out.dist = argv[++i] || '';
        }
    }
    return out;
}

function findBrotliFiles(dir, found) {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) {
            findBrotliFiles(full, found);
        } else if (entry.name.toLowerCase().endsWith('.br')) {
            found.push(full);
        }
    }
}

const args = parseArgs(process.argv.slice(2));
const distDir = path.resolve(args.dist || path.join(__dirname, '..', '..', '..', 'office-dist'));
if (!fs.existsSync(distDir)) {
    console.error('dist dir not found: ' + distDir);
    process.exit(1);
}

const files = [];
findBrotliFiles(distDir, files);
if (!files.length) {
    console.log('no .br assets under ' + distDir);
    process.exit(0);
}

for (const brFile of files) {
    const baseFile = brFile.slice(0, -3);
    const gzFile = baseFile + '.gz';
    const plain = zlib.brotliDecompressSync(fs.readFileSync(brFile));
    fs.writeFileSync(gzFile, zlib.gzipSync(plain, { level: 9 }));
    const parts = [
        path.relative(distDir, gzFile),
        'gz=' + fs.statSync(gzFile).size,
        'plain=' + plain.length
    ];
    if (args.plain) {
        fs.writeFileSync(baseFile, plain);
        parts.push('plainFile=yes');
    }
    console.log(parts.join(' '));
}
