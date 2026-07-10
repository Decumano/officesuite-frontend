// Platform adapter: isolates every native file-system call behind one small
// interface, so main.js never talks to window.__TAURI__ directly. That keeps
// this frontend portable to hosts other than Tauri — a host without a native
// backend just won't get pickWorkFolder() to return a folder, and the app's
// existing workFolderRoot-gated logic falls back to localStorage-only mode.

var Platform = (function () {
  var tauri = window.__TAURI__ && window.__TAURI__.core;

  function unavailable() {
    return Promise.reject(new Error('No native file-system backend available'));
  }

  function pickWorkFolder() {
    if (!tauri) return Promise.resolve(null);
    return tauri.invoke('pick_work_folder');
  }

  function listWorkFolder(root) {
    if (!tauri) return unavailable();
    return tauri.invoke('list_work_folder', { root: root });
  }

  function readWorkFile(root, relPath) {
    if (!tauri) return unavailable();
    return tauri.invoke('read_work_file', { root: root, relPath: relPath });
  }

  function writeWorkFile(root, relPath, content) {
    if (!tauri) return unavailable();
    return tauri.invoke('write_work_file', { root: root, relPath: relPath, content: content });
  }

  function createWorkFolder(root, relPath) {
    if (!tauri) return unavailable();
    return tauri.invoke('create_work_folder', { root: root, relPath: relPath });
  }

  function deleteWorkEntry(root, relPath, isDir) {
    if (!tauri) return unavailable();
    return tauri.invoke('delete_work_entry', { root: root, relPath: relPath, isDir: isDir });
  }

  function moveWorkEntry(root, fromRelPath, toRelPath) {
    if (!tauri) return unavailable();
    return tauri.invoke('move_work_entry', { root: root, fromRelPath: fromRelPath, toRelPath: toRelPath });
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

  // Saves a one-off exported file. On Tauri this opens a native save dialog;
  // elsewhere it triggers a standard browser download of the same content.
  function saveFile(name, content) {
    if (tauri) return tauri.invoke('save_file', { name: name, content: content });
    try {
      var blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
      var url = URL.createObjectURL(blob);
      var a = document.createElement('a');
      a.href = url;
      a.download = name;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      return Promise.resolve(true);
    } catch (e) {
      return Promise.reject(e);
    }
  }

  return {
    isNative: !!tauri,
    pickWorkFolder: pickWorkFolder,
    listWorkFolder: listWorkFolder,
    readWorkFile: readWorkFile,
    writeWorkFile: writeWorkFile,
    createWorkFolder: createWorkFolder,
    deleteWorkEntry: deleteWorkEntry,
    moveWorkEntry: moveWorkEntry,
    defaultFile: defaultFile,
    saveFile: saveFile
  };
})();
