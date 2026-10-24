(function () {
  'use strict';

  var config = window.OFFICE_CONFIG || {};
  var frame = document.getElementById('editorFrame');
  var saveButton = document.getElementById('saveButton');
  var downloadButton = document.getElementById('downloadButton');
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
  var statePollTimer = 0;
  var isPdf = /\.pdf$/i.test(String(config.filePath || ''));
  var RECENT_KEY = 'rc-office-recent';
  var recentBox = document.getElementById('recentFiles');
  var recentList = document.getElementById('recentList');
  var recentClear = document.getElementById('recentClear');

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

  function showHint(text, ms) {
    statusText.textContent = text;
    statusText.style.color = '';
    statusBox.hidden = false;
    window.clearTimeout(statusTimer);
    statusTimer = window.setTimeout(function () {
      statusBox.hidden = true;
    }, ms || 8000);
  }

  function maybeShowInsecureHint() {
    if (window.isSecureContext) {
      return;
    }
    var text = config.lang === 'zh-CN'
      ? '非安全源访问：Service Worker 与隔离特性不可用，编辑功能不受影响；如需完整能力请用 localhost 访问'
      : 'Insecure origin: Service Worker and isolation features unavailable; editing still works. Use localhost for full capability';
    window.setTimeout(function () {
      showHint(text, 8000);
    }, 1200);
  }

  function updateDirty(value) {
    dirty = !!value;
    dirtyBadge.hidden = !dirty;
    saveButton.disabled = !fileOpened || saving || (!dirty && !isPdf);
  }

  function requestState() {
    if (!fileOpened || !frame.contentWindow) {
      return;
    }
    post('document:get-state').then(function (payload) {
      if (payload && typeof payload.dirty === 'boolean' && payload.dirty !== dirty) {
        updateDirty(payload.dirty);
      }
    }).catch(function () {});
  }

  function startStatePolling() {
    if (statePollTimer || !fileOpened) {
      return;
    }
    statePollTimer = window.setInterval(requestState, 2000);
  }

  function loadRecent() {
    try {
      var raw = window.localStorage.getItem(RECENT_KEY);
      var list = raw ? JSON.parse(raw) : [];
      return Array.isArray(list) ? list : [];
    } catch (e) {
      return [];
    }
  }

  function saveRecentList(list) {
    try {
      window.localStorage.setItem(RECENT_KEY, JSON.stringify(list));
    } catch (e) {}
  }

  function rememberRecent(path, name) {
    if (!path) {
      return;
    }
    var list = loadRecent().filter(function (item) {
      return item && item.path !== path;
    });
    list.unshift({ path: path, name: name || path, time: Date.now() });
    saveRecentList(list.slice(0, 8));
  }

  function officeUrl(path) {
    return '/office?path=' + encodeURIComponent(path) +
      '&lang=' + encodeURIComponent(config.lang) +
      '&theme=' + encodeURIComponent(config.theme || 'dark');
  }

  function renderRecent() {
    if (!recentBox || !recentList) {
      return;
    }
    var list = loadRecent();
    recentList.textContent = '';
    list.forEach(function (item) {
      if (!item || !item.path) {
        return;
      }
      var li = document.createElement('li');
      var button = document.createElement('button');
      button.type = 'button';
      button.className = 'recentItem';
      var name = document.createElement('span');
      name.className = 'recentName';
      name.textContent = item.name || item.path;
      var meta = document.createElement('span');
      meta.className = 'recentMeta';
      meta.textContent = (item.time ? new Date(item.time).toLocaleString() : '') + ' · ' + item.path;
      button.appendChild(name);
      button.appendChild(meta);
      button.addEventListener('click', function () {
        window.location.href = officeUrl(item.path);
      });
      li.appendChild(button);
      recentList.appendChild(li);
    });
    recentBox.hidden = list.length === 0;
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
    if (!dirty && !isPdf) {
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
      rememberRecent(config.filePath, config.fileName);
      if (downloadButton) {
        downloadButton.disabled = !config.filePath;
      }
      setStatus('文档已打开');
      maybeShowInsecureHint();
      startSaveShortcutWatcher();
      startStatePolling();
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
      saveButton.disabled = !dirty && !isPdf;
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
    // 目录 → 就地浏览；文件 → 打开编辑器
    fetch('/office/api/browse?path=' + encodeURIComponent(value))
      .then(function (r) { return r.ok ? r.json() : null; })
      .then(function (data) {
        if (data && !data.error) {
          renderBrowse(data);
          return;
        }
        window.location.href = officeUrl(value);
      })
      .catch(function () {
        window.location.href = officeUrl(value);
      });
  }

  // ---------------- 文件管理器 ----------------

  var fmList = document.getElementById('fmList');
  var fmPath = document.getElementById('fmPath');
  var fmUp = document.getElementById('fmUp');
  var fmRefresh = document.getElementById('fmRefresh');
  var fmFoot = document.getElementById('fmFoot');
  var fmCurrent = null;
  var fmParent = null;
  var fmToken = 0;

  var ICON_KINDS = {
    word: { exts: ['doc', 'docx', 'odt', 'rtf'], cls: 'word', label: 'W' },
    excel: { exts: ['xls', 'xlsx', 'ods', 'csv'], cls: 'excel', label: 'X' },
    ppt: { exts: ['ppt', 'pptx', 'odp'], cls: 'ppt', label: 'P' },
    pdf: { exts: ['pdf'], cls: 'pdf', label: 'PDF' },
    text: { exts: ['txt'], cls: 'text', label: 'TXT' }
  };

  function iconFor(name) {
    var ext = (name.split('.').pop() || '').toLowerCase();
    for (var kind in ICON_KINDS) {
      if (ICON_KINDS[kind].exts.indexOf(ext) >= 0) {
        return ICON_KINDS[kind];
      }
    }
    return null;
  }

  function fileIconNode(name, dimmed) {
    var icon = document.createElement('span');
    var kind = iconFor(name);
    if (kind) {
      icon.className = 'fileIcon ' + kind.cls;
      icon.textContent = kind.label;
    } else {
      icon.className = 'fileIcon other';
      icon.textContent = '·';
    }
    if (dimmed) { icon.className += ' dim'; }
    return icon;
  }

  function fmRow(classes, iconNode, name, title, onClick) {
    var row = document.createElement('button');
    row.type = 'button';
    row.className = 'fmRow ' + classes;
    row.title = title || name;
    row.appendChild(iconNode);
    var label = document.createElement('span');
    label.className = 'fmName';
    label.textContent = name;
    row.appendChild(label);
    if (onClick) {
      row.addEventListener('click', onClick);
    }
    return row;
  }

  function renderBrowse(data) {
    if (!fmList) { return; }
    fmCurrent = data.path || '';
    fmParent = data.parent || '';
    fmPath.textContent = data.path || '此电脑';
    fmPath.title = data.path || '';
    if (fmUp) { fmUp.disabled = !data.path; }
    fmList.textContent = '';

    (data.dirs || []).forEach(function (dir) {
      var folderIcon = document.createElement('span');
      folderIcon.className = 'folderIcon';
      folderIcon.textContent = '▸';
      fmList.appendChild(fmRow('isDir', folderIcon, dir.name, dir.path, function () {
        browse(dir.path);
      }));
    });

    (data.documents || []).forEach(function (doc) {
      var row = fmRow('isDoc', fileIconNode(doc.name, false), doc.name, doc.path + ' — 点击在线编辑', function () {
        window.location.href = officeUrl(doc.path);
      });
      fmList.appendChild(row);
    });

    (data.others || []).forEach(function (file) {
      var row = fmRow('isOther', fileIconNode(file.name, true), file.name,
        file.path + ' — 该类型不支持在线编辑');
      row.disabled = true;
      fmList.appendChild(row);
    });

    if (!fmList.children.length) {
      var empty = document.createElement('div');
      empty.className = 'fmEmpty';
      empty.textContent = '空文件夹';
      fmList.appendChild(empty);
    }

    if (fmFoot) {
      var parts = [];
      var docCount = (data.documents || []).length;
      var otherCount = (data.others || []).length;
      if (docCount) { parts.push(docCount + ' 个可编辑文档'); }
      if (otherCount) { parts.push(otherCount + ' 个其他文件'); }
      if (data.truncated) { parts.push('列表过长已截断'); }
      fmFoot.textContent = parts.join(' · ');
      fmFoot.hidden = parts.length === 0;
    }
  }

  function browse(path) {
    if (!fmList) { return; }
    var token = ++fmToken;
    fmList.textContent = '';
    var loading = document.createElement('div');
    loading.className = 'fmEmpty';
    loading.textContent = '读取中...';
    fmList.appendChild(loading);
    if (emptyError) { emptyError.textContent = ''; }
    fetch('/office/api/browse?path=' + encodeURIComponent(path || ''))
      .then(function (r) {
        return r.json().then(function (data) { return { ok: r.ok, data: data }; });
      })
      .then(function (result) {
        if (token !== fmToken) { return; }
        if (!result.ok || result.data.error) {
          fmList.textContent = '';
          var err = document.createElement('div');
          err.className = 'fmEmpty';
          err.textContent = result.data.error || '读取失败';
          fmList.appendChild(err);
          return;
        }
        renderBrowse(result.data);
      })
      .catch(function () {
        if (token !== fmToken) { return; }
        fmList.textContent = '';
        var err = document.createElement('div');
        err.className = 'fmEmpty';
        err.textContent = '网络错误';
        fmList.appendChild(err);
      });
  }

  if (fmUp) {
    fmUp.addEventListener('click', function () {
      browse(fmParent || '');
    });
  }
  if (fmRefresh) {
    fmRefresh.addEventListener('click', function () {
      browse(fmCurrent);
    });
  }

  saveButton.addEventListener('click', saveDocument);
  if (downloadButton) {
    downloadButton.addEventListener('click', function () {
      if (!config.filePath) {
        return;
      }
      window.location.href = '/office/api/download?path=' + encodeURIComponent(config.filePath);
    });
  }
  reloadButton.addEventListener('click', function () {
    window.location.reload();
  });
  openPathButton.addEventListener('click', openTypedPath);
  pathInput.addEventListener('keydown', function (event) {
    if (event.key === 'Enter') {
      openTypedPath();
    }
  });
  if (recentClear) {
    recentClear.addEventListener('click', function () {
      saveRecentList([]);
      renderRecent();
    });
  }
  window.addEventListener('message', handleMessage);
  window.addEventListener('keydown', saveShortcutHandler, true);

  if (!config.filePath) {
    emptyState.hidden = false;
    frame.hidden = true;
    renderRecent();
    browse('');
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
