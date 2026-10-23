# -*- coding: utf-8 -*-
import json
import mimetypes
import os
import shutil
import threading

from flask import Blueprint, current_app, jsonify, redirect, render_template, request, send_file, send_from_directory

pluginDir = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(pluginDir, 'plugin.json'), encoding='utf-8-sig') as f:
    pluginID = json.load(f)['id']

office_bp = Blueprint(
    f'{pluginID}_bp',
    __name__,
    template_folder='templates',
    static_folder='static',
    static_url_path=f'/{pluginID}/static',
)

ALLOWED_EXTENSIONS = {
    '.docx', '.doc', '.odt', '.rtf', '.txt',
    '.xlsx', '.xls', '.ods', '.csv',
    '.pptx', '.ppt', '.odp',
    '.pdf',
}
MAX_SAVE_BYTES = 200 * 1024 * 1024
ASSET_PREFIXES = (
    'sdkjs', 'web-apps', 'fonts', 'libs', 'icons', 'assets',
    'convert', 'open', 'social', 'ran-fonts',
)
_fileLocks = {}
_fileLocksGuard = threading.Lock()

for ext, mime in {
    '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    '.doc': 'application/msword',
    '.xls': 'application/vnd.ms-excel',
    '.ppt': 'application/vnd.ms-powerpoint',
    '.odt': 'application/vnd.oasis.opendocument.text',
    '.ods': 'application/vnd.oasis.opendocument.spreadsheet',
    '.odp': 'application/vnd.oasis.opendocument.presentation',
    '.pdf': 'application/pdf',
    '.csv': 'text/csv',
}.items():
    mimetypes.add_type(mime, ext)


def resolveDistDir():
    candidates = []
    envDir = os.environ.get('RC_OFFICE_DIST')
    if envDir:
        candidates.append(envDir)
    programData = os.environ.get('PROGRAMDATA')
    if programData:
        candidates.append(os.path.join(programData, 'RemoteConnecter', 'office', 'dist'))
    candidates.append(os.path.join(pluginDir, 'office-dist'))
    candidates.append(os.path.abspath(os.path.join(pluginDir, '..', '..', 'office-dist')))
    for candidate in candidates:
        if candidate and os.path.isfile(os.path.join(candidate, 'editor.html')):
            return candidate
    return None


def getFileLock(path):
    key = os.path.normcase(os.path.abspath(path))
    with _fileLocksGuard:
        return _fileLocks.setdefault(key, threading.Lock())


def isAllowedFile(path):
    return os.path.isfile(path) and os.path.splitext(path)[1].lower() in ALLOWED_EXTENSIONS


def normalizeLang(rawLang):
    value = (rawLang or '').strip().lower()
    if value.startswith('zh'):
        return 'zh-CN'
    return 'en-US'


@office_bp.after_request
def addOfficeHeaders(response):
    response.headers['Cross-Origin-Opener-Policy'] = 'same-origin'
    response.headers['Cross-Origin-Embedder-Policy'] = 'require-corp'
    response.headers['Cross-Origin-Resource-Policy'] = 'same-origin'
    return response


@office_bp.route(f'/{pluginID}')
@office_bp.route(f'/{pluginID}/')
def officeHome():
    path = request.args.get('path', '').strip()
    lang = normalizeLang(request.args.get('lang') or request.headers.get('Accept-Language'))
    theme = 'light' if request.args.get('theme') == 'light' else 'dark'
    error = None
    fileName = ''
    mtime = 0
    if path:
        path = os.path.abspath(path)
        if not isAllowedFile(path):
            error = '文件不存在或类型不受支持'
        else:
            fileName = os.path.basename(path)
            mtime = os.path.getmtime(path)
    config = {
        'pluginId': pluginID,
        'filePath': path if not error else '',
        'fileName': fileName,
        'mtime': mtime,
        'lang': lang,
        'theme': theme,
        'error': error,
        'distReady': resolveDistDir() is not None,
    }
    return render_template('office.html', config=config)


@office_bp.route(f'/{pluginID}/app')
def officeAppRoot():
    return redirect(f'/{pluginID}/app/editor.html')


@office_bp.route(f'/{pluginID}/app/')
@office_bp.route(f'/{pluginID}/app/<path:filename>')
def officeApp(filename='editor.html'):
    distDir = resolveDistDir()
    if not distDir:
        return 'Office assets not found', 503
    if not filename:
        filename = 'editor.html'
    response = send_from_directory(distDir, filename, conditional=True)
    if filename.lower().endswith('.html'):
        response.headers['Cache-Control'] = 'no-cache'
    else:
        response.headers['Cache-Control'] = 'public, max-age=3600'
    return response


def serveDistFile(filename):
    distDir = resolveDistDir()
    if not distDir:
        return 'Office assets not found', 503
    serveName = filename
    encoding = None
    baseName = None
    if filename.lower().endswith('.br'):
        baseName = filename[:-3]
        gzName = baseName + '.gz'
        if os.path.isfile(os.path.join(distDir, gzName)):
            serveName = gzName
            encoding = 'gzip'
        elif os.path.isfile(os.path.join(distDir, baseName)):
            serveName = baseName
    response = send_from_directory(distDir, serveName, conditional=True)
    if serveName.lower().endswith('.br'):
        try:
            with open(os.path.join(distDir, serveName), 'rb') as f:
                magic = f.read(2)
            response.headers['Content-Encoding'] = 'gzip' if magic == b'\x1f\x8b' else 'br'
        except OSError:
            pass
    elif encoding:
        response.headers['Content-Encoding'] = encoding
    if baseName and baseName.lower().endswith('.wasm'):
        response.headers['Content-Type'] = 'application/wasm'
    response.headers['Cache-Control'] = 'public, max-age=3600'
    return response


def makeAssetView(prefix):
    def view(filename):
        return serveDistFile(prefix + '/' + filename)
    view.__name__ = 'officeAsset_' + prefix
    return view


for _prefix in ASSET_PREFIXES:
    office_bp.add_url_rule(
        f'/{_prefix}/<path:filename>',
        endpoint=f'officeAsset_{_prefix}',
        view_func=makeAssetView(_prefix),
    )


@office_bp.route(f'/{pluginID}/api/open')
def officeOpen():
    path = request.args.get('path', '').strip()
    if not path:
        return jsonify({'error': 'missing path'}), 400
    path = os.path.abspath(path)
    if not isAllowedFile(path):
        return jsonify({'error': 'file not found or unsupported'}), 404
    size = os.path.getsize(path)
    if size > MAX_SAVE_BYTES:
        return jsonify({'error': 'file too large'}), 413
    mimeType = mimetypes.guess_type(path)[0] or 'application/octet-stream'
    response = send_file(path, mimetype=mimeType, as_attachment=False, conditional=True)
    response.headers['Cache-Control'] = 'no-store'
    response.headers['X-Office-MTime'] = str(os.path.getmtime(path))
    response.headers['X-Office-Size'] = str(size)
    return response


@office_bp.route(f'/{pluginID}/api/save', methods=['POST'])
def officeSave():
    path = request.form.get('path', '').strip()
    expectedMtime = request.form.get('mtime', '').strip()
    upload = request.files.get('file')
    if not path or not isAllowedFile(path):
        return jsonify({'error': 'file not found or unsupported'}), 404
    if upload is None:
        return jsonify({'error': 'missing file'}), 400
    path = os.path.abspath(path)
    with getFileLock(path):
        currentMtime = os.path.getmtime(path)
        if expectedMtime:
            try:
                expectedValue = float(expectedMtime)
            except ValueError:
                return jsonify({'error': 'invalid mtime'}), 400
            if abs(currentMtime - expectedValue) > 0.5:
                return jsonify({'error': 'file changed on disk', 'mtime': currentMtime}), 409
        tmpPath = path + '.rc-office.tmp'
        try:
            upload.save(tmpPath)
            size = os.path.getsize(tmpPath)
            if size > MAX_SAVE_BYTES:
                os.remove(tmpPath)
                return jsonify({'error': 'file too large'}), 413
            backupPath = path + '.rc-office.bak'
            shutil.copy2(path, backupPath)
            os.replace(tmpPath, path)
        except Exception as exc:
            if os.path.exists(tmpPath):
                try:
                    os.remove(tmpPath)
                except OSError:
                    pass
            return jsonify({'error': str(exc)}), 500
    return jsonify({
        'ok': True,
        'path': path,
        'size': size,
        'mtime': os.path.getmtime(path),
        'backup': path + '.rc-office.bak',
    })


@office_bp.route('/sw.js')
def officeServiceWorker():
    script = (
        "self.addEventListener('install',function(){self.skipWaiting();});"
        "self.addEventListener('activate',function(e){e.waitUntil(self.clients.claim());});"
    )
    response = current_app.response_class(script, mimetype='application/javascript')
    response.headers['Service-Worker-Allowed'] = '/'
    response.headers['Cache-Control'] = 'no-cache'
    return response


@office_bp.route(f'/{pluginID}/health')
def officeHealth():
    distDir = resolveDistDir()
    return jsonify({
        'plugin': pluginID,
        'dist': distDir,
        'ready': bool(distDir),
    })
