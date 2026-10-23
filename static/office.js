(function () {
  'use strict';

  var config = window.OFFICE_CONFIG || {};
  var frame = document.getElementById('editorFrame');
  var saveButton = document.getElementById('saveButton');
  var reloadButton = document.getElementById('reloadButton');
  var dirtyBadge = document.getElementById('dirtyBadge');
  var statusBox = document.querySelector('.officeStatus');
  var statusText = document.getElementById('statusText');
  var statusTimer = 0;
  var emptyState = document.getElementById('emptyState');
  var pathInput = document.getElementById('pathInput');
  var openPathButton = document.getElementById('openPathButton');
  var emptyError = document.getElementById('emptyError');
  var requestId = 0;
  var pending = new Map();
  var currentMtime = String(config.mtime || '');
  var fileOpened = false;
  var dirty = false;
  var saving = false;
  var shortcutTimer = 0;

  function setStatus(text, isError) {
    statusText.textContent = text;
    statusText.style.color = isError ? '#ff8a80' : '';
    statusBox.hidden = false;
    window.clearTimeout(statusTimer);
    if (!isError) {
      statusTimer = window.setTimeout(function () {
        statusBox.hidden = true;
      }, 4000);
    }
  }

  function updateDirty(value) {
    dirty = !!value;
    dirtyBadge.hidden = !dirty;
    saveButton.disabled = !dirty || !fileOpened || saving;
  }

  function saveShortcutHandler(event) {
    if (!(event.ctrlKey || event.metaKey) || event.altKey) {
      return;
    }
    if (String(event.key || '').toLowerCase() !== 's') {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    if (event.stopImmediatePropagation) {
      event.stopImmediatePropagation();
    }
    if (!fileOpened) {
      return;
    }
    if (!dirty) {
      setStatus('没有未保存的更改');
      return;
    }
    saveDocument();
  }

  function attachSaveShortcut() {
    if (!frame.contentWindow) {
      return;
    }
    try {
      var wrapperDoc = frame.contentWindow.document;
      if (!wrapperDoc.__rcSaveShortcut) {
        wrapperDoc.__rcSaveShortcut = true;
        wrapperDoc.addEventListener('keydown', saveShortcutHandler, true);
      }
      var nestedFrames = wrapperDoc.querySelectorAll('iframe');
      for (var i = 0; i < nestedFrames.length; i++) {
        var nestedWindow = nestedFrames[i].contentWindow;
        if (nestedWindow && !nestedWindow.__rcSaveShortcut) {
          nestedWindow.__rcSaveShortcut = true;
          nestedWindow.addEventListener('keydown', saveShortcutHandler, true);
        }
      }
    } catch (e) {}
  }

  function startSaveShortcutWatcher() {
    attachSaveShortcut();
    if (shortcutTimer) {
      return;
    }
    shortcutTimer = window.setInterval(attachSaveShortcut, 1000);
    window.setTimeout(function () {
      if (shortcutTimer) {
        window.clearInterval(shortcutTimer);
        shortcutTimer = 0;
      }
    }, 30000);
  }

  function post(type, payload) {
    var id = 'office-' + (++requestId);
    return new Promise(function (resolve, reject) {
      pending.set(id, { resolve: resolve, reject: reject });
      frame.contentWindow.postMessage({ id: id, type: type, payload: payload || {} }, window.location.origin);
      window.setTimeout(function () {
        if (pending.has(id)) {
          pending.delete(id);
          reject(new Error(type + ' timeout'));
        }
      }, 200000);
    });
  }

  function handleMessage(event) {
    if (event.origin !== window.location.origin || event.source !== frame.contentWindow) {
      return;
    }
    var data = event.data || {};
    var type = data.type;
    var id = data.id;
    var payload = data.payload || {};
    if (!type || type.indexOf('document:') !== 0) {
      return;
    }
    if (type === 'document:ready') {
      setStatus('编辑器已就绪，正在打开文档...');
      startSaveShortcutWatcher();
      if (config.filePath) {
        openDocument();
      }
      return;
    }
    if (type === 'document:error') {
      var message = payload.message || '编辑器返回错误';
      setStatus(message, true);
      if (id && pending.has(id)) {
        pending.get(id).reject(new Error(message));
        pending.delete(id);
      }
      return;
    }
    if (type === 'document:dirty-changed') {
      updateDirty(payload.dirty);
      return;
    }
    if (type === 'document:opened') {
      fileOpened = true;
      updateDirty(false);
      setStatus('文档已打开');
      startSaveShortcutWatcher();
      return;
    }
    if (id && pending.has(id)) {
      pending.get(id).resolve(payload);
      pending.delete(id);
    }
  }

  async function openDocument() {
    if (!config.filePath || !config.distReady) {
      return;
    }
    setStatus('正在读取文件...');
    try {
      var response = await fetch('/office/api/open?path=' + encodeURIComponent(config.filePath), {
        cache: 'no-store',
      });
      if (!response.ok) {
        var errorBody = await response.text();
        throw new Error(errorBody || ('HTTP ' + response.status));
      }
      var mtime = response.headers.get('X-Office-MTime');
      if (mtime) {
        currentMtime = mtime;
      }
      var buffer = await response.arrayBuffer();
      await post('document:open-buffer', {
        fileName: config.fileName,
        buffer: buffer,
        mimeType: response.headers.get('Content-Type') || '',
        readonly: false,
      });
    } catch (error) {
      setStatus('打开失败: ' + error.message, true);
    }
  }

  async function saveDocument() {
    if (!fileOpened || !config.filePath || saving) {
      return;
    }
    saving = true;
    saveButton.disabled = true;
    setStatus('正在导出并保存...');
    try {
      var extension = (config.fileName.split('.').pop() || '').toLowerCase();
      var extMap = { docx: 'DOCX', xlsx: 'XLSX', pptx: 'PPTX', csv: 'CSV' };
      var savePayload = extMap[extension] ? { targetExt: extMap[extension] } : {};
      var result = await post('document:save', savePayload);
      if (!result.file) {
        throw new Error('编辑器没有返回文件');
      }
      var formData = new FormData();
      formData.append('path', config.filePath);
      formData.append('mtime', currentMtime);
      formData.append('file', result.file, config.fileName);
      var response = await fetch('/office/api/save', {
        method: 'POST',
        body: formData,
      });
      var body = await response.json().catch(function () { return {}; });
      if (!response.ok) {
        throw new Error(body.error || ('HTTP ' + response.status));
      }
      currentMtime = String(body.mtime || '');
      updateDirty(!!result.dirty);
      setStatus('已保存，备份文件: ' + (body.backup || ''));
    } catch (error) {
      saveButton.disabled = !dirty;
      setStatus('保存失败: ' + error.message, true);
    } finally {
      saving = false;
    }
  }

  function openTypedPath() {
    var value = pathInput.value.trim();
    if (!value) {
      emptyError.textContent = '请输入文件完整路径';
      return;
    }
    window.location.href = '/office?path=' + encodeURIComponent(value) +
      '&lang=' + encodeURIComponent(config.lang) +
      '&theme=' + encodeURIComponent(config.theme || 'dark');
  }

  saveButton.addEventListener('click', saveDocument);
  reloadButton.addEventListener('click', function () {
    window.location.reload();
  });
  openPathButton.addEventListener('click', openTypedPath);
  pathInput.addEventListener('keydown', function (event) {
    if (event.key === 'Enter') {
      openTypedPath();
    }
  });
  window.addEventListener('message', handleMessage);
  window.addEventListener('keydown', saveShortcutHandler, true);

  if (!config.filePath) {
    emptyState.hidden = false;
    frame.hidden = true;
    return;
  }

  if (!config.distReady) {
    emptyState.hidden = false;
    frame.hidden = true;
    emptyError.textContent = '未找到 Office 静态资源，请先部署 office-dist。';
    return;
  }

  emptyState.hidden = true;
  frame.hidden = false;
  try {
    window.localStorage.setItem('ran-theme', config.theme || 'dark');
    window.localStorage.removeItem('ui-theme-id');
  } catch (e) {}
  frame.src = '/office/app/editor.html?embed=1&embedOrigin=' +
    encodeURIComponent(window.location.origin) +
    '&locale=' + encodeURIComponent(config.lang) +
    '&lang=' + encodeURIComponent(config.lang);
  setStatus('正在加载编辑器...');
})();
