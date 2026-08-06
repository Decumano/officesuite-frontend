// Platform adapter: isolates every native file-system call behind one small
// interface, so main.js never talks to window.__TAURI__ (or the web API)
// directly. On Tauri, calls go through window.__TAURI__.core.invoke. In this
// web build, when Tauri isn't present, the same six work-folder methods are
// backed by fetch() calls to this server's /api/workspace/* endpoints instead
// of rejecting as "unavailable" — the server derives the workspace root from
// the logged-in session, so `root` is never actually sent to it.

var Platform = (function () {
  // The split pane embeds the app in a same-origin iframe; Tauri only injects
  // __TAURI__ into top-level windows, so child frames borrow the parent's.
  var tauriGlobal = window.__TAURI__ || (function () {
    try { return window.parent !== window ? window.parent.__TAURI__ : null; }
    catch (e) { return null; }
  })();
  var tauri = tauriGlobal && tauriGlobal.core;

  // Sentinel returned by pickWorkFolder() once authenticated: there's no
  // folder to pick on the web, just an implicit per-user workspace.
  var CLOUD_ROOT = 'cloud';

  // ── Cloud workspaces (web only) ──
  // A named workspace on the web is just a top-level folder of the user's one
  // server-side space, addressed with the root sentinel 'cloud:<folder>'.
  // Scoping is applied here, in the single funnel every API call goes
  // through: work-folder methods derive the folder from their `root`
  // argument; share/comment methods don't receive a root, so main.js tells
  // us the active folder once via setCloudWorkspace. The server only ever
  // sees full root-relative paths, so sync, sharing and deletion tombstones
  // work unchanged.
  var cloudWorkspaceFolder = '';

  function setCloudWorkspace(folder) {
    cloudWorkspaceFolder = folder || '';
  }

  function cloudFolderOf(root) {
    return (typeof root === 'string' && root.indexOf('cloud:') === 0) ? root.slice(6) : '';
  }

  // Workspace-relative -> root-relative, for methods that take `root`.
  function cloudRel(root, relPath) {
    var folder = cloudFolderOf(root);
    return folder ? folder + '/' + relPath : relPath;
  }

  // Same, for the share/comment methods scoped via setCloudWorkspace.
  function activeRel(relPath) {
    return cloudWorkspaceFolder ? cloudWorkspaceFolder + '/' + relPath : relPath;
  }

  function api(path, options) {
    options = options || {};
    options.credentials = 'include';
    return fetch('/api' + path, options).then(function (res) {
      if (!res.ok) {
        return res.json().catch(function () { return {}; }).then(function (body) {
          throw new Error(body.error || ('Request failed: ' + res.status));
        });
      }
      return res;
    });
  }

  function apiJson(path, options) {
    return api(path, options).then(function (res) { return res.json(); });
  }

  function pickWorkFolder() {
    if (tauri) return tauri.invoke('pick_work_folder');
    return currentUser().then(function (user) {
      return user ? CLOUD_ROOT : null;
    });
  }

  function listWorkFolder(root) {
    if (tauri) return tauri.invoke('list_work_folder', { root: root });
    return apiJson('/workspace').then(function (tree) {
      var folder = cloudFolderOf(root);
      if (!folder) return tree;

      // Scoped workspace: present the named top-level folder's contents as
      // the whole tree, with workspace-relative paths.
      var node = null;
      (tree || []).forEach(function (e) {
        if (e.isDir && e.name === folder) node = e;
      });
      if (!node) return [];

      var strip = folder.length + 1;
      function stripRel(list) {
        list.forEach(function (e) {
          e.relPath = e.relPath.slice(strip);
          if (e.children && e.children.length) stripRel(e.children);
        });
        return list;
      }
      return stripRel(node.children || []);
    });
  }

  function readWorkFile(root, relPath) {
    if (tauri) return tauri.invoke('read_work_file', { root: root, relPath: relPath });
    return api('/workspace/file?path=' + encodeURIComponent(cloudRel(root, relPath))).then(function (res) { return res.text(); });
  }

  function writeWorkFile(root, relPath, content) {
    if (tauri) return tauri.invoke('write_work_file', { root: root, relPath: relPath, content: content });
    return api('/workspace/file?path=' + encodeURIComponent(cloudRel(root, relPath)), {
      method: 'PUT',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: content
    }).then(function () { return undefined; });
  }

  // Cheap change-detection: {modified, hash} for one file, without the
  // content. Resolves null on Tauri, where live sync doesn't apply (the
  // desktop app has its own explicit cloud-sync flow).
  function statWorkFile(root, relPath) {
    if (tauri) return Promise.resolve(null);
    return apiJson('/workspace/stat?path=' + encodeURIComponent(cloudRel(root, relPath)));
  }

  function createWorkFolder(root, relPath) {
    if (tauri) return tauri.invoke('create_work_folder', { root: root, relPath: relPath });
    return api('/workspace/folder', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ relPath: cloudRel(root, relPath) })
    }).then(function () { return undefined; });
  }

  function deleteWorkEntry(root, relPath, isDir) {
    if (tauri) return tauri.invoke('delete_work_entry', { root: root, relPath: relPath, isDir: isDir });
    return api('/workspace/entry?path=' + encodeURIComponent(cloudRel(root, relPath)) + '&isDir=' + !!isDir, {
      method: 'DELETE'
    }).then(function () { return undefined; });
  }

  function moveWorkEntry(root, fromRelPath, toRelPath) {
    if (tauri) return tauri.invoke('move_work_entry', { root: root, fromRelPath: fromRelPath, toRelPath: toRelPath });
    return api('/workspace/move', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ from: cloudRel(root, fromRelPath), to: cloudRel(root, toRelPath) })
    }).then(function () { return undefined; });
  }

  // ── Auth (web only; no-ops don't apply to Tauri, which has no accounts) ──

  function currentUser() {
    if (tauri) return Promise.resolve(null);
    return fetch('/api/auth/me', { credentials: 'include' }).then(function (res) {
      return res.ok ? res.json() : null;
    });
  }

  function register(email, password) {
    return apiJson('/auth/register', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ email: email, password: password })
    });
  }

  function login(email, password) {
    return apiJson('/auth/login', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ email: email, password: password })
    });
  }

  function logout() {
    return api('/auth/logout', { method: 'POST' }).then(function () { return undefined; });
  }

  // ── Cloud sync (Tauri only; the web build already talks to the cloud
  // natively through the calls above and has no separate "connect" step) ──

  function cloudConnect(serverUrl, email, password, root) {
    if (!tauri) return Promise.reject(new Error('cloud sync is desktop-only'));
    return tauri.invoke('cloud_connect', { serverUrl: serverUrl, email: email, password: password, root: root });
  }

  function cloudDisconnect() {
    if (!tauri) return Promise.reject(new Error('cloud sync is desktop-only'));
    return tauri.invoke('cloud_disconnect');
  }

  function cloudStatus() {
    if (!tauri) return Promise.resolve({ connected: false, mode: 'disconnected' });
    return tauri.invoke('cloud_status');
  }

  function cloudSyncNow() {
    if (!tauri) return Promise.reject(new Error('cloud sync is desktop-only'));
    return tauri.invoke('cloud_sync_now');
  }

  function cloudListConflicts() {
    if (!tauri) return Promise.resolve([]);
    return tauri.invoke('cloud_list_conflicts');
  }

  function cloudResolveConflict(relPath, choice) {
    if (!tauri) return Promise.reject(new Error('cloud sync is desktop-only'));
    return tauri.invoke('cloud_resolve_conflict', { relPath: relPath, choice: choice });
  }

  // ── Sharing & comments (web only; the desktop app works on a local folder
  // and reaches shared content through the website instead) ──

  function webOnly() { return Promise.reject(new Error('sharing is available on the web app')); }

  function shareEntry(relPath, isDir, email, permission) {
    if (tauri) return webOnly();
    return apiJson('/shares', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ relPath: activeRel(relPath), isDir: isDir, email: email, permission: permission })
    });
  }

  function listSharesFor(relPath) {
    if (tauri) return Promise.resolve([]);
    return apiJson('/shares?path=' + encodeURIComponent(activeRel(relPath)));
  }

  function revokeShare(shareId) {
    if (tauri) return webOnly();
    return api('/shares/' + encodeURIComponent(shareId), { method: 'DELETE' }).then(function () { return undefined; });
  }

  function listSharedWithMe() {
    if (tauri) return Promise.resolve([]);
    return apiJson('/shared');
  }

  function listSharedFolder(shareId, subPath) {
    if (tauri) return webOnly();
    return apiJson('/shared/list?share=' + encodeURIComponent(shareId) + '&path=' + encodeURIComponent(subPath || ''));
  }

  function readSharedFile(shareId, subPath) {
    if (tauri) return webOnly();
    return api('/shared/file?share=' + encodeURIComponent(shareId) + '&path=' + encodeURIComponent(subPath || ''))
      .then(function (res) { return res.text(); });
  }

  function statSharedFile(shareId, subPath) {
    if (tauri) return Promise.resolve(null);
    return apiJson('/shared/stat?share=' + encodeURIComponent(shareId) + '&path=' + encodeURIComponent(subPath || ''));
  }

  function writeSharedFile(shareId, subPath, content) {
    if (tauri) return webOnly();
    return api('/shared/file?share=' + encodeURIComponent(shareId) + '&path=' + encodeURIComponent(subPath || ''), {
      method: 'PUT',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: content
    }).then(function () { return undefined; });
  }

  // ── Link shares ("anyone with the link"; the token is the authorization) ──

  function createShareLink(relPath, isDir, permission) {
    if (tauri) return webOnly();
    return apiJson('/links', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ relPath: activeRel(relPath), isDir: isDir, permission: permission })
    });
  }

  function listShareLinks(relPath) {
    if (tauri) return Promise.resolve([]);
    return apiJson('/links?path=' + encodeURIComponent(activeRel(relPath)));
  }

  function revokeShareLink(linkId) {
    if (tauri) return webOnly();
    return api('/links/' + encodeURIComponent(linkId), { method: 'DELETE' }).then(function () { return undefined; });
  }

  function linkMeta(token) {
    if (tauri) return webOnly();
    return apiJson('/link/' + encodeURIComponent(token));
  }

  function listLinkFolder(token, subPath) {
    if (tauri) return webOnly();
    return apiJson('/link/' + encodeURIComponent(token) + '/list?path=' + encodeURIComponent(subPath || ''));
  }

  function readLinkFile(token, subPath) {
    if (tauri) return webOnly();
    return api('/link/' + encodeURIComponent(token) + '/file?path=' + encodeURIComponent(subPath || ''))
      .then(function (res) { return res.text(); });
  }

  function statLinkFile(token, subPath) {
    if (tauri) return Promise.resolve(null);
    return apiJson('/link/' + encodeURIComponent(token) + '/stat?path=' + encodeURIComponent(subPath || ''));
  }

  function writeLinkFile(token, subPath, content) {
    if (tauri) return webOnly();
    return api('/link/' + encodeURIComponent(token) + '/file?path=' + encodeURIComponent(subPath || ''), {
      method: 'PUT',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: content
    }).then(function () { return undefined; });
  }

  // target: { path } for my own file, { share, subPath } for an account
  // share, or { link, subPath } for a share link (works logged out).
  function commentQs(target) {
    if (target.link) return 'link=' + encodeURIComponent(target.link) + '&subPath=' + encodeURIComponent(target.subPath || '');
    if (target.share) return 'share=' + encodeURIComponent(target.share) + '&subPath=' + encodeURIComponent(target.subPath || '');
    return 'path=' + encodeURIComponent(activeRel(target.path));
  }

  function listComments(target) {
    if (tauri) return Promise.resolve([]);
    return apiJson('/comments?' + commentQs(target));
  }

  function addComment(target, body, anchor) {
    if (tauri) return webOnly();
    var payload = target.link
      ? { link: target.link, subPath: target.subPath || '', body: body }
      : target.share
        ? { share: target.share, subPath: target.subPath || '', body: body }
        : { path: activeRel(target.path), body: body };
    if (anchor) payload.anchor = anchor;
    return apiJson('/comments', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
  }

  function deleteComment(commentId) {
    if (tauri) return webOnly();
    return api('/comments/' + encodeURIComponent(commentId), { method: 'DELETE' }).then(function () { return undefined; });
  }

  // ── Custom fonts (account-level; see src/fonts.rs) ──
  // Web: served as real HTTP assets, so @font-face just works in any browser.
  // Tauri: fetched through the cloud connection (if any) as base64 data: URIs
  // via Rust commands, since the desktop app has no session/account of its
  // own to authenticate a direct fetch with.

  function listCustomFonts() {
    if (tauri) return tauri.invoke('cloud_list_fonts').catch(function () { return []; });
    return apiJson('/fonts').catch(function () { return []; });
  }

  function uploadCustomFont(familyName, filename, bytes) {
    if (tauri) return webOnly();
    var qs = '?familyName=' + encodeURIComponent(familyName) + '&filename=' + encodeURIComponent(filename);
    return apiJson('/fonts' + qs, {
      method: 'POST',
      headers: { 'Content-Type': 'application/octet-stream' },
      body: bytes
    });
  }

  function deleteCustomFont(id) {
    if (tauri) return webOnly();
    return api('/fonts/' + encodeURIComponent(id), { method: 'DELETE' }).then(function () { return undefined; });
  }

  // Returns a data: URI (base64) for the font, playing the same role on
  // Tauri that a direct /api/fonts/:id fetch plays on the web — a same-origin
  // <link>/@font-face url() isn't available to an unauthenticated webview,
  // but an inlined data: URI needs no request at all.
  function customFontDataUrl(id) {
    if (tauri) return tauri.invoke('cloud_font_data_url', { fontId: id });
    return Promise.resolve('/api/fonts/' + encodeURIComponent(id));
  }

  var DEFAULT_FILE_CONTENT = [
    "# Welcome to Lore Keep\n\n",
    "Lore Keep is a lightweight office suite that stores everything in **Markdown**.\n\n",
    "## Features\n\n",
    "**Documents** rich markdown editing with live preview\n",
    "**Spreadsheets** formula-capable grid stored as cell=value pairs\n",
    "## Markdown Quick Reference\n",
    "| Element | Syntax |\n",
    "|---------|--------|\n",
    "| Bold | **text** |\n",
    "| Italic | *text* |\n",
    "| Heading | # H1 ## H2 |\n",
    "| List | - item |\n",
    "| Blockquote | > text |\n",
    "| Code | ``` code ``` |\n",
    "| Link | [text](url) |\n\n",
    "## Getting Started\n\n",
    "1. Click **New** in the sidebar to create a file\n",
    "2. Write in Markdown, use the toolbar for formatting\n",
    "3. Toggle the **split** view button to preview alongside your writing\n",
    "4. Click the **download** button to export your file\n\n",
    "> All files are saved automatically to your browser's local storage.\n\n",
    "Happy writing!"
  ].join('');

  // Returns canned welcome content for a brand-new install. On Tauri this
  // matches the Rust-side `defaultFile` command; elsewhere it's generated
  // here so first-run works without a native backend.
  function defaultFile(name) {
    if (tauri) return tauri.invoke('defaultFile', { name: name });
    return Promise.resolve({
      name: 'Welcome to Lore Keep ' + name,
      docType: 'doc',
      content: DEFAULT_FILE_CONTENT
    });
  }

  // Triggers a standard browser download of a Blob under the given name.
  function downloadBlob(name, blob) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // Saves a one-off exported file. On Tauri this opens a native save dialog;
  // elsewhere it triggers a standard browser download of the same content.
  function saveFile(name, content) {
    if (tauri) return tauri.invoke('save_file', { name: name, content: content });
    try {
      downloadBlob(name, new Blob([content], { type: 'text/plain;charset=utf-8' }));
      return Promise.resolve(true);
    } catch (e) {
      return Promise.reject(e);
    }
  }

  // Renders `html` to a PDF server-side (headless Chromium, headers/footers
  // disabled - see src/export.rs) and downloads the result. Not available on
  // Tauri, which renders/prints PDFs locally via the print-preview flow instead.
  function exportPdf(name, html) {
    if (tauri) return Promise.reject(new Error('exportPdf is web-only'));
    return api('/export/pdf', {
      method: 'POST',
      headers: { 'Content-Type': 'text/html;charset=utf-8' },
      body: html
    }).then(function (res) {
      return res.blob();
    }).then(function (blob) {
      downloadBlob(name, blob);
      return true;
    });
  }

  return {
    isNative: !!tauri,
    setCloudWorkspace: setCloudWorkspace,
    pickWorkFolder: pickWorkFolder,
    listWorkFolder: listWorkFolder,
    readWorkFile: readWorkFile,
    writeWorkFile: writeWorkFile,
    statWorkFile: statWorkFile,
    createWorkFolder: createWorkFolder,
    deleteWorkEntry: deleteWorkEntry,
    moveWorkEntry: moveWorkEntry,
    defaultFile: defaultFile,
    saveFile: saveFile,
    exportPdf: exportPdf,
    currentUser: currentUser,
    register: register,
    login: login,
    logout: logout,
    cloudConnect: cloudConnect,
    cloudDisconnect: cloudDisconnect,
    cloudStatus: cloudStatus,
    cloudSyncNow: cloudSyncNow,
    cloudListConflicts: cloudListConflicts,
    cloudResolveConflict: cloudResolveConflict,
    shareEntry: shareEntry,
    listSharesFor: listSharesFor,
    revokeShare: revokeShare,
    listSharedWithMe: listSharedWithMe,
    listSharedFolder: listSharedFolder,
    readSharedFile: readSharedFile,
    writeSharedFile: writeSharedFile,
    statSharedFile: statSharedFile,
    listComments: listComments,
    addComment: addComment,
    deleteComment: deleteComment,
    createShareLink: createShareLink,
    listShareLinks: listShareLinks,
    revokeShareLink: revokeShareLink,
    linkMeta: linkMeta,
    listLinkFolder: listLinkFolder,
    readLinkFile: readLinkFile,
    writeLinkFile: writeLinkFile,
    statLinkFile: statLinkFile,
    listCustomFonts: listCustomFonts,
    uploadCustomFont: uploadCustomFont,
    deleteCustomFont: deleteCustomFont,
    customFontDataUrl: customFontDataUrl
  };
})();
