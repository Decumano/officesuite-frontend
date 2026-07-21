// ── STATE ──
let files = {};
let currentFileId = null;
let currentAppType = 'doc';
let newFileType = 'doc';
let editorView = 'write';
let saveTimer = null;
let sheetData = {};
// Cell decorations, aliased per page just like sheetData below:
// sheetColors = {ref: '#rrggbb'} paints a cell's background, serialized as a
// `;#hex` suffix on the cell's line (B1=5;#FAFAFA); sheetTextColors paints a
// cell's text, serialized as a `;txt#hex` suffix (B1=5;txt#FF0000);
// sheetMerges = {anchorRef: [coveredRefs]} joins cells into one bigger cell,
// serialized as a `:B2,B3,B4` suffix (B1=5:B2,B3,B4). The suffixes may appear
// on one line in any order.
let sheetColors = {};
let sheetTextColors = {};
let sheetMerges = {};
// Multi-tab spreadsheets: each page owns a {ref: value} map; sheetData is
// always an alias to the active page's map, so the whole grid/formula engine
// keeps working unchanged. Serialized with a `# page name` heading per page;
// formulas reference other tabs with the link-like syntax [Tab name](A1).
let sheetPages = [{ name: 'Sheet 1', data: sheetData, colors: sheetColors, textColors: sheetTextColors, merges: sheetMerges }];
let sheetPageIdx = 0;
// Set when the user switches tabs mid-formula: the next cell clicked on the
// other tab is inserted into that pending formula as [Tab name](REF) and the
// editor jumps back to where they were typing.
let formulaPickCtx = null;
let activeCell =
{
  row: 1,
  col: 1
};

let isDragging = false;
let selectionAnchor = null;
let selectionEnd = null;

let editingCell = null;
let isRefDragging = false;
// True while the ref-drag above targets the formula bar instead of an
// in-cell input (typing "=" in the bar, then clicking/dragging cells).
let refDragViaBar = false;
let formulaInsertStart = null;
let formulaInsertLength = 0;

let isFillDragging = false;
let fillSourceBox = null;
let fillRange = null;
let incrementLiteralsOnFill = false;

let functionSuggestState = null;

var sidebarSearch = '';
var newDocTemplate = 'blank';
var sidebarCollapsed = false;

// Shared undo/redo controller factory for the two plain-textarea editors
// (Docs, Graphs) - both need the exact same thing: batch a burst of typing
// into one undo step (capturing the "before" state on the first keystroke of
// a burst, via keydown, since oninput only ever sees the "after" state), plus
// a forced single-step snapshot for toolbar actions that rewrite the value
// programmatically (bold/italic wrapping, inserting a table, etc).
function makeTextUndo(getEditor)
{
  var undoStack = [],
      redoStack = [],
      pending = null,
      debounceTimer = null;

  function snapshotOf(editor)
  {
    return { value: editor.value, start: editor.selectionStart, end: editor.selectionEnd };
  }

  function flushPending()
  {
    if (!pending)
      return;

    undoStack.push(pending);
    redoStack = [];
    pending = null;

    if (undoStack.length > 100)
      undoStack.shift();
  }

  function noteKeydown(e)
  {
    if (e.ctrlKey || e.metaKey || e.altKey)
      return;

    var navigationKeys = ['ArrowLeft','ArrowRight','ArrowUp','ArrowDown','Home','End','PageUp','PageDown','Shift','Control','Alt','Meta','Escape'];

    if (navigationKeys.indexOf(e.key) !== -1)
      return;

    if (!pending)
      pending = snapshotOf(getEditor());

    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(flushPending, 700);
  }

  function forceSnapshot()
  {
    flushPending();
    undoStack.push(snapshotOf(getEditor()));
    redoStack = [];

    if (undoStack.length > 100)
      undoStack.shift();
  }

  function restore(from, to)
  {
    flushPending();

    if (!from.length)
      return false;

    var editor = getEditor();
    to.push(snapshotOf(editor));

    var snap = from.pop();
    editor.value = snap.value;
    editor.selectionStart = snap.start;
    editor.selectionEnd = snap.end;

    return true;
  }

  return  {
            noteKeydown: noteKeydown,
            forceSnapshot: forceSnapshot,
            undo: function(){ return restore(undoStack, redoStack); },
            redo: function(){ return restore(redoStack, undoStack); },
            reset: function(){ undoStack = []; redoStack = []; pending = null; clearTimeout(debounceTimer); }
          };
}

// Chart definitions persisted with the sheet; live Chart.js instances are
// kept separately (sheetChartInstances, keyed by chart id) since they aren't
// serializable data - they're rebuilt from sheetCharts on every render.
let sheetCharts = [];
let sheetChartInstances = {};
let sheetChartDrag = null;
let newChartType = 'bar';

// Sheet-wide undo: a snapshot is the whole {data, charts} pair rather than a
// single cell, since one undo step (a paste, a fill, a chart move) can touch
// many cells/charts at once - same "snapshot the whole relevant state" idea
// used for Notebook pages, just at the sheet level instead of the page level.
let sheetUndoStack = [];
let sheetRedoStack = [];

let historyDiffIndex = null;

let graphEditorView = 'split';

let workFolderRoot = null;
let folders = {};
let expandedFolders = new Set();
let draggedEntryId = null;

let ROWS = 40,
    COLS = 26;

let documentHeadings = [];
let headingRenderCursor = 0;
let slugCounts = {};

async function createDefaultFile()
{
  let fileContext = await Platform.defaultFile("Decumano");

  const file = createFile
  (
    fileContext.name,
    fileContext.docType,
    fileContext.content
  );

  openFile(file);
}

async function init()
{
  loadSettings();

  // #link=<token> in the URL is a share link (see initLinkAccess); it works
  // with or without a logged-in account.
  var linkToken = null;
  if (!Platform.isNative)
  {
    var lm = (window.location.hash || '').match(/^#link=([0-9a-fA-F-]{36})/);
    if (lm) linkToken = lm[1];
  }

  // Web only: surface the desktop-app download page. On Tauri you're already
  // in the desktop app, so the item stays hidden.
  if (!Platform.isNative)
  {
    var dlItem = document.getElementById('file-menu-download');
    if (dlItem) dlItem.style.display = '';
    refreshSharedSection();
  }

  // Anonymous link visitor: skip the personal-workspace bootstrap entirely —
  // there is no account, just the linked file/folder.
  if (linkToken && !(await Platform.currentUser()))
  {
    initTheme();
    applyEnabledTypes();
    applySidebarCollapse();
    buildSheet();
    renderFileList();
    await initLinkAccess(linkToken);
    return;
  }

  await refreshCloudStatus();
  initTheme();
  applyEnabledTypes();
  applySidebarCollapse();
  buildSheet();

  if (workFolderRoot)
  {
    await loadWorkFolderTree();
    await loadBacklinksIndex();
    loadOpenTabs();
    injectCustomFontsStyle(); // fire-and-forget: doesn't block first paint
    if (linkToken) await initLinkAccess(linkToken);
    return;
  }

  loadFromStorage();
  await loadBacklinksIndex();
  loadOpenTabs();
  renderFileList();

  if (Object.keys(files).length === 0 && !linkToken)
  {
    createDefaultFile();
  }

  if (linkToken) await initLinkAccess(linkToken);
}

function saveToStorage()
{
  try
  {
    localStorage.setItem('lore_keep_v2', JSON.stringify(files));
  }
  catch(e)
  {
    console.warn('Storage error', e);
  }
}

function loadFromStorage()
{
  try
  {
    const d = localStorage.getItem('lore_keep_v2');
    if (d)
      files = JSON.parse(d);
  }
  catch(e)
  {
    files = {};
  }
}

function loadSettings()
{
  try
  {
    const raw = localStorage.getItem('lore_keep_settings'),
          settings = raw ? JSON.parse(raw) : {};

    workFolderRoot = settings.workFolder || null;
  }
  catch(e)
  {
    workFolderRoot = null;
  }
}

function saveSettings()
{
  try
  {
    localStorage.setItem('lore_keep_settings', JSON.stringify({ workFolder: workFolderRoot }));
  }
  catch(e)
  {
    console.warn('Settings storage error', e);
  }
}

var DOC_TEMPLATES =
{
  character: '{type: Character, status: Active, faction: Unknown, race: Unknown}\n\n#character\n\n## Background\n\n## Personality\n\n## Relationships\n\n## Notes\n',
  faction:   '{type: Faction, status: Active, alignment: Neutral, size: Unknown}\n\n#faction\n\n## Overview\n\n## Goals\n\n## Members\n\n## Rivals\n',
  location:  '{type: Location, region: Unknown, climate: Unknown, status: Active}\n\n#location\n\n## Description\n\n## Points of Interest\n\n## Inhabitants\n\n## History\n',
  event:     '{type: Event, era: Unknown, date: Unknown, outcome: Unknown}\n\n#event\n\n## Summary\n\n## Participants\n\n## Consequences\n\n## Notes\n',
  region:    '{type: Region, climate: Unknown, culture: Unknown, status: Active, ruler: Unknown}\n\n#region\n\n## Overview\n\n## Climate & Geography\n\n## Culture & Society\n\n## Political Status\n\n## Key NPCs\n\n## Linked Events\n\n## Notes\n'
};

function defaultContentForType(type, name)
{
  if (type === 'sheet')
    return '---\ntype: spreadsheet\n---\n\n';

  if (type === 'graph')
    return '<!-- type: graph -->\n\n```mermaid\ngraph TD\n    A[Start] --> B{Decision}\n    B -->|Yes| C[Continue]\n    B -->|No| D[Stop]\n```\n';

  if (type === 'notebook')
    return serializeNotebookData(defaultNotebookData());

  // The JSON-backed apps embed a display title; seed it with the name the
  // user gave the file so the first save doesn't rename the file to a
  // canned default.
  if (type === 'glossary')
    return JSON.stringify({ name: name || 'My Glossary', entries: [], roots: [] }, null, 2);

  if (type === 'calendar')
    return JSON.stringify({
      name: name || 'My Calendar',
      daysPerYear: 365,
      epochFictionYear: 1,
      epochRealDate: new Date().toISOString().slice(0, 10),
      seasons: [
        { id: 's1', name: 'Spring', color: '#7a9e6a', monthIds: [] },
        { id: 's2', name: 'Summer', color: '#e8a832', monthIds: [] },
        { id: 's3', name: 'Autumn', color: '#c4621a', monthIds: [] },
        { id: 's4', name: 'Winter', color: '#4a9fc8', monthIds: [] }
      ],
      months: [],
      holidays: []
    }, null, 2);

  if (type === 'economy')
    return JSON.stringify({ name: name || 'Economy Notes', currencies: [], exchangeRates: [], tradeGoods: [], regions: [] }, null, 2);

  if (type === 'bestiary')
    return JSON.stringify({ name: name || 'My Bestiary', beasts: [] }, null, 2);

  return '';
}

function createFile(name, type, content)
{
  const fileId = 'f_' + Date.now();

  if (content === undefined)
    content = defaultContentForType(type, name);

  files[fileId] =
  {
    name, type, content, modified: Date.now()
  };

  saveToStorage();

  return fileId;
}

// Strips characters that are illegal in filenames on Windows (a superset of
// what's illegal on macOS/Linux), so a display name can double as a real
// on-disk filename without surprising the user when they open the folder.
function sanitizeFileName(name)
{
  return String(name || '')
          .trim()
          .replace(/[\\/:*?"<>|]/g, '_')
          .replace(/\.+$/, '')
          .trim();
}

// Finds a rel_path that isn't already taken by another file, suffixing " (2)",
// " (3)", etc. on collision - except a candidate matching excludeRelPath itself
// is always accepted, so re-saving a file under its own current name is a no-op.
function uniqueRelPath(folder, baseName, ext, excludeRelPath)
{
  var prefix = folder ? folder + '/' : '',
      candidate = prefix + baseName + '.' + ext,
      n = 2;

  while (files[candidate] && candidate !== excludeRelPath)
  {
    candidate = prefix + baseName + ' (' + n + ').' + ext;
    n++;
  }

  return candidate;
}

function uniqueFolderRelPath(parent, baseName, excludeRelPath)
{
  var prefix = parent ? parent + '/' : '',
      candidate = prefix + baseName,
      n = 2;

  while (folders[candidate] && candidate !== excludeRelPath)
  {
    candidate = prefix + baseName + ' (' + n + ')';
    n++;
  }

  return candidate;
}

function fileTypeIcon(type)
{
  if (type === 'sheet')
    return  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">'+
              '<rect x="3" y="3" width="18" height="18" rx="2"/>'+
              '<line x1="3" y1="9" x2="21" y2="9"/>'+
              '<line x1="3" y1="15" x2="21" y2="15"/>'+
              '<line x1="9" y1="3" x2="9" y2="21"/>'+
              '<line x1="15" y1="3" x2="15" y2="21"/>'+
            '</svg>';

  if (type === 'graph')
    return  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">'+
              '<circle cx="6" cy="6" r="3"/>'+
              '<circle cx="6" cy="18" r="3"/>'+
              '<circle cx="18" cy="12" r="3"/>'+
              '<line x1="8.5" y1="7.5" x2="15.5" y2="10.5"/>'+
              '<line x1="8.5" y1="16.5" x2="15.5" y2="13.5"/>'+
            '</svg>';

  if (type === 'notebook')
    return  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">'+
              '<rect x="4" y="3" width="16" height="18" rx="2"/>'+
              '<line x1="8" y1="3" x2="8" y2="21"/>'+
              '<line x1="12" y1="8" x2="17" y2="8"/>'+
              '<line x1="12" y1="12" x2="17" y2="12"/>'+
              '<line x1="12" y1="16" x2="15" y2="16"/>'+
            '</svg>';

  if (type === 'glossary')
    return  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">'+
              '<path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"/>'+
              '<path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"/>'+
              '<line x1="9" y1="8" x2="16" y2="8"/>'+
              '<line x1="9" y1="12" x2="14" y2="12"/>'+
            '</svg>';

  if (type === 'bestiary')
    return  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">'+
              '<circle cx="5.5" cy="10" r="1.6"/>'+
              '<circle cx="9.5" cy="6.5" r="1.6"/>'+
              '<circle cx="14.5" cy="6.5" r="1.6"/>'+
              '<circle cx="18.5" cy="10" r="1.6"/>'+
              '<path d="M12 11c-2.6 0-5.5 2.6-5.5 5a3 3 0 0 0 3 3c1 0 1.7-.5 2.5-.5s1.5.5 2.5.5a3 3 0 0 0 3-3c0-2.4-2.9-5-5.5-5z"/>'+
            '</svg>';

  if (type === 'calendar')
    return  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">'+
              '<rect x="3" y="4" width="18" height="18" rx="2"/>'+
              '<line x1="16" y1="2" x2="16" y2="6"/>'+
              '<line x1="8" y1="2" x2="8" y2="6"/>'+
              '<line x1="3" y1="10" x2="21" y2="10"/>'+
            '</svg>';

  if (type === 'economy')
    return  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">'+
              '<circle cx="12" cy="12" r="9"/>'+
              '<path d="M14.8 9A2 2 0 0 0 13 8h-2a2 2 0 0 0 0 4h2a2 2 0 0 1 0 4h-2a2 2 0 0 1-1.8-1"/>'+
              '<line x1="12" y1="6" x2="12" y2="8"/>'+
              '<line x1="12" y1="16" x2="12" y2="18"/>'+
            '</svg>';

  return  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">'+
            '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>'+
            '<polyline points="14,2 14,8 20,8"/>'+
          '</svg>';
}

function fileExtensionFor(type)
{
  if (type === 'sheet')    return 'mds';
  if (type === 'notebook') return 'mdn';
  if (type === 'graph')    return 'mdg';
  if (type === 'glossary') return 'mdl';
  if (type === 'calendar') return 'mdc';
  if (type === 'economy')  return 'mde';
  if (type === 'bestiary') return 'mdb';
  return 'mdp';
}

var FOLDER_ICON =  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">'+
                      '<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V7z"/>'+
                    '</svg>';

var EMPTY_FILE_LIST_HTML =  '<div style="padding: 20px 12px; text-align: center; color: var(--text3); font-size: 12px;">' +
                              'No files yet.<br>Click <strong style="color:var(--accent)">+ New</strong> to start.' +
                            '</div>';

var CHEVRON_ICON =  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">'+
                       '<polyline points="9,6 15,12 9,18"/>'+
                     '</svg>';

function fileRowHtml(id, f, depth)
{
  const active = (id === currentFileId),
        icon = fileTypeIcon(f.type),
        date = new Date(f.modified).toLocaleDateString(undefined,{month:'short',day:'numeric'}),
        indent = depth ? (' style="padding-left:' + (12 + depth * 16) + 'px;"') : '',
        safeId = escAttr(id);

  const tagChips = fileTagsOf(f).map(function(t){
    return '<span class="file-tag" onclick="event.stopPropagation();sidebarTagFilter(\'' + escAttr(t) + '\')">#' + escHtml(t) + '</span>';
  }).join('');

  return  '<div class="file-item ' + (active?'active':'') + '"' + indent + ' data-id="' + escHtml(id) + '"' +
            ' draggable="true" ondragstart="handleDragStart(event,\'' + safeId + '\')" ondragend="handleDragEnd(event)"' +
            ' onclick="openFile(\'' + safeId + '\')" oncontextmenu="openContextMenu(event,\'file\',\'' + safeId + '\')">' +
            '<span class="file-icon">' +
              icon +
            '</span>' +
            '<div class="file-info">'+
              '<div class="file-name" ondblclick="startInlineRename(event,this,\'' + safeId + '\')">' +
                escHtml(f.name) +
              '</div>'+
              '<div class="file-meta">' +
                date + tagChips +
              '</div>'+
            '</div>'+
            '<span class="file-del" onclick="deleteFile(event,\'' + safeId + '\')">'+
              '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">'+
                '<line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>'+
              '</svg>'+
            '</span>'+
          '</div>';
}

function folderRowHtml(path, depth)
{
  const expanded = expandedFolders.has(path),
        indent = 12 + depth * 16,
        safePath = escAttr(path);

  return  '<div class="folder-item" style="padding-left:' + indent + 'px;" data-id="' + escHtml(path) + '"' +
            ' draggable="true" ondragstart="handleDragStart(event,\'' + safePath + '\')" ondragend="handleDragEnd(event)"' +
            ' ondragenter="handleDragEnter(event)" ondragover="handleDragOver(event)" ondragleave="handleDragLeave(event)" ondrop="handleDrop(event,\'' + safePath + '\')"' +
            ' onclick="toggleFolderExpanded(\'' + safePath + '\')" oncontextmenu="openContextMenu(event,\'folder\',\'' + safePath + '\')">' +
            '<span class="folder-caret ' + (expanded ? 'expanded' : '') + '">' + CHEVRON_ICON + '</span>' +
            '<span class="file-icon">' + FOLDER_ICON + '</span>' +
            '<div class="file-info">'+
              '<div class="file-name" ondblclick="startInlineRename(event,this,\'' + safePath + '\')">' +
                escHtml(folders[path].name) +
              '</div>'+
            '</div>'+
            '<span class="file-del" onclick="deleteFolderEntry(event,\'' + safePath + '\')">'+
              '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">'+
                '<line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>'+
              '</svg>'+
            '</span>'+
          '</div>';
}

function renderFileList()
{
  const fileList = document.getElementById('file-list');

  var _searchQ = sidebarSearch.trim().toLowerCase();

  // "#tag" filters by tag: file-level tags on the JSON data apps, and inline
  // #hashtags in Documents (which carry them in the markdown itself).
  var _tagQ = null;
  if (_searchQ.charAt(0) === '#')
  {
    _tagQ = _searchQ.slice(1).trim();
    _searchQ = '';
  }

  const filesOfType = Object.entries(files).filter(function(entry)
  {
    var f = entry[1];
    if (entry[0] === SHARED_TMP_ID) return false; // shared files live in their own sidebar section
    if (f.type !== currentAppType) return false;
    if (_tagQ)
    {
      if (fileTagsOf(f).some(function(t){ return t.toLowerCase() === _tagQ; })) return true;
      return !!(f.content && f.content.toLowerCase().indexOf('#' + _tagQ) !== -1);
    }
    if (!_searchQ) return true;
    if (f.name.toLowerCase().indexOf(_searchQ) !== -1) return true;
    if (fileTagsOf(f).some(function(t){ return t.toLowerCase().indexOf(_searchQ) !== -1; })) return true;
    if (f.content && f.content.toLowerCase().indexOf(_searchQ) !== -1) return true;
    return false;
  });

  renderOpenTabs(); // names/active state shown in the tab bar change on the same events

  if (Object.keys(folders).length === 0)
  {
    filesOfType.sort((a, b) => b[1].modified - a[1].modified);

    fileList.innerHTML = filesOfType.length
                          ?
                            filesOfType.map(([id, f]) => fileRowHtml(id, f, 0)).join('')
                          :
                            EMPTY_FILE_LIST_HTML;
    return;
  }

  const byFolder = {};

  filesOfType.forEach(function(entry)
  {
    const folder = entry[1].folder || '';
    (byFolder[folder] = byFolder[folder] || []).push(entry);
  });

  const childFolders = {};

  Object.keys(folders).forEach(function(path)
  {
    const parent = folders[path].parent || '';
    (childFolders[parent] = childFolders[parent] || []).push(path);
  });

  function renderLevel(parentPath, depth)
  {
    let html = '';

    const subFolders = (childFolders[parentPath] || []).slice().sort(function(a, b)
    {
      return folders[a].name.localeCompare(folders[b].name);
    });

    subFolders.forEach(function(path)
    {
      html += folderRowHtml(path, depth);

      if (expandedFolders.has(path))
        html += renderLevel(path, depth + 1);
    });

    const filesHere = (byFolder[parentPath] || []).slice().sort(function(a, b)
    {
      return b[1].modified - a[1].modified;
    });

    filesHere.forEach(function(entry)
    {
      html += fileRowHtml(entry[0], entry[1], depth);
    });

    return html;
  }

  fileList.innerHTML = renderLevel('', 0) || EMPTY_FILE_LIST_HTML;
}

function onSidebarSearch(value)
{
  sidebarSearch = value;
  onGlobalSearch(value);
}

function filterByTag(tag)
{
  if (currentAppType !== 'doc')
    switchAppType('doc');
  sidebarSearch = '#' + tag;
  var inp = document.getElementById('sidebar-search');
  if (inp) inp.value = sidebarSearch;
  renderFileList();
}

// Tag chip on a sidebar file row: filter the current app's list by that tag.
function sidebarTagFilter(tag)
{
  sidebarSearch = '#' + tag;
  var inp = document.getElementById('sidebar-search');
  if (inp) inp.value = sidebarSearch;
  onGlobalSearch(sidebarSearch);
}

// ── GLOBAL SEARCH ──────────────────────────────────────────

var globalSearchActive  = false;
var globalSearchLoadTmr = null;

function extractSearchableText(id)
{
  var f = files[id];
  if (!f || !f.content) return '';
  var c = f.content;

  if (f.type === 'bestiary')
  {
    try
    {
      var bd = JSON.parse(c), bparts = (bd.tags || []).slice();
      (bd.beasts || []).forEach(function(b) { bparts.push(b.name, b.commonName, b.category, b.habitat, b.description); if (b.abilities) bparts = bparts.concat(b.abilities); if (b.tags) bparts = bparts.concat(b.tags); });
      return bparts.filter(Boolean).join(' ');
    }
    catch(e) { return ''; }
  }

  if (f.type === 'glossary')
  {
    try
    {
      var d = JSON.parse(c), parts = (d.tags || []).slice();
      (d.entries || []).forEach(function(e) { parts.push(e.word, e.definition, e.example, e.language); if (e.tags) parts = parts.concat(e.tags); });
      (d.roots   || []).forEach(function(r) { parts.push(r.form, r.meaning); if (r.examples) parts = parts.concat(r.examples); if (r.tags) parts = parts.concat(r.tags); });
      return parts.filter(Boolean).join(' ');
    }
    catch(e) { return c; }
  }

  if (f.type === 'calendar')
  {
    try
    {
      var d = JSON.parse(c), parts = (d.tags || []).slice();
      (d.months   || []).forEach(function(m) { parts.push(m.name, m.season, m.description); if (m.tags) parts = parts.concat(m.tags); });
      (d.seasons  || []).forEach(function(s) { parts.push(s.name, s.description); });
      (d.holidays || []).forEach(function(h) { parts.push(h.name, h.description); if (h.tags) parts = parts.concat(h.tags); });
      return parts.filter(Boolean).join(' ');
    }
    catch(e) { return c; }
  }

  if (f.type === 'economy')
  {
    try
    {
      var d = JSON.parse(c), parts = (d.tags || []).slice();
      (d.currencies || []).forEach(function(x) { parts.push(x.name, x.code, x.symbol, x.region, x.description); if (x.tags) parts = parts.concat(x.tags); });
      (d.tradeGoods || []).forEach(function(x) { parts.push(x.name, x.category, x.description, x.origin); if (x.tags) parts = parts.concat(x.tags); });
      (d.regions    || []).forEach(function(x) { parts.push(x.name, x.economicStatus, x.notes); if (x.tags) parts = parts.concat(x.tags); });
      return parts.filter(Boolean).join(' ');
    }
    catch(e) { return c; }
  }

  if (f.type === 'sheet')
  {
    // "REF=value" or "REF=formula=result" lines — extract value parts
    return c.split('\n').map(function(l)
    {
      var eq = l.indexOf('=');
      return eq > 0 ? l.slice(eq + 1) : '';
    }).join(' ');
  }

  return c; // doc: raw markdown; graph/notebook: JSON that contains label/text strings
}

function gsEscRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

function gsHighlight(text, query)
{
  if (!text || !query) return escHtml(text || '');
  var re = new RegExp('(' + gsEscRe(query) + ')', 'gi');
  return escHtml(text).replace(re, '<mark class="gs-hl">$1</mark>');
}

function gsSnippet(raw, query)
{
  if (!raw) return '';
  var lc  = raw.toLowerCase();
  var qlo = query.toLowerCase();
  var idx = lc.indexOf(qlo);
  if (idx < 0) return '';
  var s   = Math.max(0, idx - 50);
  var e   = Math.min(raw.length, idx + query.length + 90);
  var txt = (s > 0 ? '…' : '') + raw.slice(s, e) + (e < raw.length ? '…' : '');
  return gsHighlight(txt, query);
}

function runGlobalSearch(query)
{
  var q   = query.trim();
  var qlo = q.toLowerCase();
  if (!q) { hideGlobalSearch(); return; }

  var results = [];
  Object.keys(files).forEach(function(id)
  {
    var f = files[id];
    if (!f) return;
    var nameHit = f.name.toLowerCase().indexOf(qlo) >= 0;
    var text    = extractSearchableText(id);
    var snippet = gsSnippet(text, q);
    if (nameHit || snippet)
      results.push({ id: id, file: f, nameHit: nameHit, snippet: snippet });
  });

  results.sort(function(a, b)
  {
    if (a.nameHit !== b.nameHit) return a.nameHit ? -1 : 1;
    return a.file.name.localeCompare(b.file.name);
  });

  globalSearchActive = true;
  document.getElementById('file-list').style.display           = 'none';
  document.getElementById('sidebar-search-clear').style.display = '';

  var panel = document.getElementById('global-search-results');
  panel.style.display = '';

  if (!results.length)
  {
    panel.innerHTML = '<div class="gs-empty">No results for <strong>' + escHtml(q) + '</strong></div>';
    return;
  }

  panel.innerHTML =
    '<div class="gs-count">' + results.length + ' result' + (results.length !== 1 ? 's' : '') + '</div>' +
    results.map(function(r)
    {
      return '<div class="gs-result" onclick="openFile(\'' + escAttr(r.id) + '\');clearGlobalSearch();">' +
        '<div class="gs-result-name">' +
          '<span class="gs-file-icon">' + fileTypeIcon(r.file.type) + '</span>' +
          '<span class="gs-name">' + gsHighlight(r.file.name, q) + '</span>' +
        '</div>' +
        (r.snippet ? '<div class="gs-snippet">' + r.snippet + '</div>' : '') +
      '</div>';
    }).join('');
}

function hideGlobalSearch()
{
  globalSearchActive = false;
  document.getElementById('global-search-results').style.display    = 'none';
  document.getElementById('file-list').style.display                = '';
  document.getElementById('sidebar-search-clear').style.display     = 'none';
  renderFileList();
}

// Work-folder mode: background-load any unread files, then call `done` so
// the caller can re-run its filter over the full content set.
function scheduleUnloadedContentLoad(done)
{
  if (!workFolderRoot) return;
  globalSearchLoadTmr = setTimeout(async function()
  {
    var unloaded = Object.keys(files).filter(function(id) { return files[id] && !files[id].contentLoaded; });
    if (!unloaded.length) return;
    await Promise.all(unloaded.map(async function(id)
    {
      try
      {
        var _r = await Platform.readWorkFile(workFolderRoot, id);
        if (files[id]) { files[id].content = _r; files[id].contentLoaded = true; }
      }
      catch(err) {}
    }));
    done();
  }, 150);
}

function onGlobalSearch(value)
{
  clearTimeout(globalSearchLoadTmr);
  var q = value.trim();
  if (!q) { hideGlobalSearch(); return; }

  // '#tag' queries filter the sidebar file list in place (renderFileList
  // understands them) instead of opening the global results panel.
  if (q.charAt(0) === '#')
  {
    globalSearchActive = false;
    document.getElementById('global-search-results').style.display    = 'none';
    document.getElementById('file-list').style.display                = '';
    document.getElementById('sidebar-search-clear').style.display     = '';
    renderFileList();
    scheduleUnloadedContentLoad(function(){ if (sidebarSearch.trim() === q) renderFileList(); });
    return;
  }

  runGlobalSearch(q); // immediate pass — searches already-loaded content

  scheduleUnloadedContentLoad(function(){ if (sidebarSearch.trim() === q) runGlobalSearch(q); }); // re-render with full content
}

function clearGlobalSearch()
{
  var inp = document.getElementById('sidebar-search');
  if (inp) inp.value = '';
  sidebarSearch = '';
  hideGlobalSearch();
}

// ── SIDEBAR COLLAPSE ───────────────────────────────────────

function toggleSidebarCollapse()
{
  sidebarCollapsed = !sidebarCollapsed;
  document.getElementById('sidebar').classList.toggle('sidebar--rail', sidebarCollapsed);
  try { localStorage.setItem('lk_sidebar_rail', sidebarCollapsed ? '1' : '0'); } catch(e) {}
}

function applySidebarCollapse()
{
  try { sidebarCollapsed = localStorage.getItem('lk_sidebar_rail') === '1'; } catch(e) {}
  document.getElementById('sidebar').classList.toggle('sidebar--rail', sidebarCollapsed);
}

// ── MOBILE SIDEBAR DRAWER ──
// On narrow screens the sidebar is off-canvas; a floating button slides it
// in over a backdrop, and picking a file slides it back out.

function toggleMobileSidebar()
{
  var sidebar = document.getElementById('sidebar'),
      open = !sidebar.classList.contains('mobile-open');

  // The rail (icons-only) state is a desktop concept — a drawer opens full.
  if (open && sidebarCollapsed)
    toggleSidebarCollapse();

  sidebar.classList.toggle('mobile-open', open);
  document.getElementById('mobile-backdrop').classList.toggle('open', open);
}

function closeMobileSidebar()
{
  var sidebar = document.getElementById('sidebar'),
      backdrop = document.getElementById('mobile-backdrop');

  if (sidebar) sidebar.classList.remove('mobile-open');
  if (backdrop) backdrop.classList.remove('open');
}

// ── BACKLINKS INDEX ────────────────────────────────────────
//
// Stored in _lkbl.json at the work-folder root (or localStorage key "lk_bl").
// Shape: { targetFileId: [sourceFileId, ...] }
// Files themselves are never modified for backlink data.

var backlinksIndex = {};

async function loadBacklinksIndex()
{
  if (workFolderRoot)
  {
    try
    {
      var raw = await Platform.readWorkFile(workFolderRoot, '_lkbl.json');
      backlinksIndex = JSON.parse(raw) || {};
    }
    catch(e) { backlinksIndex = {}; }
  }
  else
  {
    try { var r = localStorage.getItem('lk_bl'); backlinksIndex = r ? JSON.parse(r) : {}; }
    catch(e) { backlinksIndex = {}; }
  }
}

async function saveBacklinksIndex()
{
  if (workFolderRoot)
  {
    try { await Platform.writeWorkFile(workFolderRoot, '_lkbl.json', JSON.stringify(backlinksIndex, null, 2)); }
    catch(e) {}
  }
  else
  {
    try { localStorage.setItem('lk_bl', JSON.stringify(backlinksIndex)); } catch(e) {}
  }
}

// Returns list of file ids that sourceId currently mentions by name.
// Outside Work Folder mode content always lives in memory, so only the
// lazily-loaded work-folder files can be "not loaded yet".
function computeMentionsListFor(sourceId)
{
  if (!files[sourceId] || (workFolderRoot && !files[sourceId].contentLoaded)) return [];
  var text = extractSearchableText(sourceId).toLowerCase();
  var out  = [];
  Object.keys(files).forEach(function(targetId)
  {
    if (targetId === sourceId || !files[targetId]) return;
    var name = files[targetId].name.toLowerCase();
    if (name && text.indexOf(name) >= 0) out.push(targetId);
  });
  return out;
}

// Called when a file is opened or saved — updates only changed entries in the index.
async function updateBacklinksForFile(sourceId)
{
  if (!files[sourceId] || (workFolderRoot && !files[sourceId].contentLoaded)) return;

  var newMentions = computeMentionsListFor(sourceId);
  var oldMentions = files[sourceId].lkMentions || [];
  var removed     = oldMentions.filter(function(t) { return newMentions.indexOf(t) < 0; });
  var added       = newMentions.filter(function(t) { return oldMentions.indexOf(t) < 0; });
  files[sourceId].lkMentions = newMentions;

  if (!removed.length && !added.length) return;

  removed.forEach(function(tid)
  {
    if (!backlinksIndex[tid]) return;
    var i = backlinksIndex[tid].indexOf(sourceId);
    if (i >= 0) backlinksIndex[tid].splice(i, 1);
    if (!backlinksIndex[tid].length) delete backlinksIndex[tid];
  });

  added.forEach(function(tid)
  {
    if (!backlinksIndex[tid]) backlinksIndex[tid] = [];
    if (backlinksIndex[tid].indexOf(sourceId) < 0) backlinksIndex[tid].push(sourceId);
  });

  await saveBacklinksIndex();
}

// ── BACKLINKS (live scan) ───────────────────────────────────

function computeBacklinks(fileId)
{
  if (!fileId || !files[fileId]) return [];
  var targetName = files[fileId].name.toLowerCase();
  var results    = [];
  Object.keys(files).forEach(function(id)
  {
    if (id === fileId || !files[id]) return;
    var f    = files[id];
    var text = f.name + ' ' + extractSearchableText(id);
    if (text.toLowerCase().indexOf(targetName) >= 0)
      results.push({ id: id, file: f });
  });
  return results;
}


function toggleFolderExpanded(path)
{
  if (expandedFolders.has(path))
    expandedFolders.delete(path);
  else
    expandedFolders.add(path);

  renderFileList();
}

// ── OPEN-FILE TABS (desktop) ───────────────────────────────
//
// Every file opened gets a browser-style tab above the editor, so switching
// between a handful of working files doesn't mean hunting the sidebar. The
// bar is hidden on phones via CSS (the sidebar drawer plays that role there).
// Tab entries are file ids, so Work Folder renames/moves must remap them
// (remapOpenTab) and structural refreshes must drop dead ids (pruneOpenTabs).

var openTabs = [];

function loadOpenTabs()
{
  try
  {
    var raw = localStorage.getItem('lk_tabs');
    openTabs = raw ? JSON.parse(raw).filter(function(id) { return !!files[id]; }) : [];
  }
  catch(e) { openTabs = []; }
  renderOpenTabs();
}

function persistOpenTabs()
{
  try { localStorage.setItem('lk_tabs', JSON.stringify(openTabs)); } catch(e) {}
}

function noteTabOpened(id)
{
  if (id === SHARED_TMP_ID || !files[id])
    return;

  if (openTabs.indexOf(id) < 0)
  {
    openTabs.push(id);
    persistOpenTabs();
  }

  renderOpenTabs();
}

function remapOpenTab(oldId, newId)
{
  var i = openTabs.indexOf(oldId);

  if (i < 0)
    return;

  if (openTabs.indexOf(newId) >= 0)
    openTabs.splice(i, 1); // already open under the new id - drop the stale one
  else
    openTabs[i] = newId;

  persistOpenTabs();
  renderOpenTabs();
}

function pruneOpenTabs()
{
  var before = openTabs.length;
  openTabs = openTabs.filter(function(id) { return !!files[id]; });

  if (openTabs.length !== before)
    persistOpenTabs();

  renderOpenTabs();
}

function closeOpenTab(e, id)
{
  e.stopPropagation();

  var i = openTabs.indexOf(id);

  if (i >= 0)
  {
    openTabs.splice(i, 1);
    persistOpenTabs();
  }

  // Closing the active tab moves to a neighbor; closing the last one leaves
  // the editors empty, same as after deleting the open file.
  if (currentFileId === id)
  {
    var next = openTabs[Math.min(i, openTabs.length - 1)];

    if (next)
    {
      openFile(next);
      return;
    }

    currentFileId = null;
    clearActiveEditors();
    renderFileList();
  }

  renderOpenTabs();
}

function renderOpenTabs()
{
  var bar = document.getElementById('open-tabs-bar');

  if (!bar)
    return;

  var tabs = openTabs.filter(function(id) { return !!files[id]; });

  if (!tabs.length)
  {
    bar.innerHTML = '';
    bar.style.display = 'none';
    return;
  }

  bar.style.display = '';
  bar.innerHTML = tabs.map(function(id)
  {
    var f = files[id];

    return '<div class="open-tab' + (id === currentFileId ? ' active' : '') + '" onclick="openFile(\'' + escAttr(id) + '\')" title="' + escAttr(f.name) + '">' +
             '<span class="open-tab-icon">' + fileTypeIcon(f.type) + '</span>' +
             '<span class="open-tab-name">' + escHtml(f.name) + '</span>' +
             '<button class="open-tab-close" onclick="closeOpenTab(event,\'' + escAttr(id) + '\')" title="Close tab">×</button>' +
           '</div>';
  }).join('');

  var active = bar.querySelector('.open-tab.active');

  if (active && active.scrollIntoView)
    active.scrollIntoView({ block: 'nearest', inline: 'nearest' });
}

// ── RELATIONSHIP GRAPH (desktop) ───────────────────────────
//
// Force-directed view of how files relate: solid edges are explicit
// cross-file links ([text](lorekeep://file/<id>) in Documents), faint edges
// are name mentions (the same signal that feeds the backlinks index). Plain
// canvas + a tiny simulation - no vendored graph library needed at the scale
// of a personal archive.

var relState = null; // { nodes, edges, raf, ... } while the modal is open

// User-tunable layout: spacing = spring rest length in px (how far apart
// connected nodes sit; repulsion scales with it so density feels uniform),
// size = node radius percentage. Persisted so the graph opens how you left it.
var relPrefs = { spacing: 110, size: 100 };

function loadRelPrefs()
{
  try
  {
    var raw = localStorage.getItem('lk_rel_prefs');
    if (raw)
    {
      var p = JSON.parse(raw);
      if (p.spacing >= 40 && p.spacing <= 300) relPrefs.spacing = p.spacing;
      if (p.size >= 50 && p.size <= 250)       relPrefs.size = p.size;
    }
  }
  catch(e) {}
}

function saveRelPrefs()
{
  try { localStorage.setItem('lk_rel_prefs', JSON.stringify(relPrefs)); } catch(e) {}
}

function relApplyPrefs()
{
  if (!relState)
    return;

  relState.nodes.forEach(function(n) { n.r = n.baseR * relPrefs.size / 100; });
  relState.alpha = Math.max(relState.alpha, 0.6); // reheat so the layout re-settles
}

function onRelSpacingChange(v)
{
  relPrefs.spacing = parseInt(v, 10) || 110;
  saveRelPrefs();
  relApplyPrefs();
}

function onRelSizeChange(v)
{
  relPrefs.size = parseInt(v, 10) || 100;
  saveRelPrefs();
  relApplyPrefs();
}

var REL_TYPE_COLORS =
{
  doc:      '#3498db',
  sheet:    '#2ecc71',
  graph:    '#9b59b6',
  notebook: '#e67e22',
  glossary: '#f1c40f',
  calendar: '#4a9fc8',
  economy:  '#e8a832',
  bestiary: '#e74c3c'
};

var REL_TYPE_LABELS =
{
  doc: 'Document', sheet: 'Spreadsheet', graph: 'Diagram', notebook: 'Notebook',
  glossary: 'Glossary', calendar: 'Calendar', economy: 'Economy', bestiary: 'Bestiary'
};

// Explicit doc links: any lorekeep://file/<url-encoded id> occurrence.
function relExplicitTargets(content)
{
  var out = [], re = /lorekeep:\/\/file\/([^)\s"'<>]+)/g, m;

  while ((m = re.exec(content || '')))
  {
    try { out.push(decodeURIComponent(m[1])); }
    catch(e) {}
  }

  return out;
}

async function relEnsureContentsLoaded()
{
  if (!workFolderRoot)
    return;

  var pending = Object.keys(files).filter(function(id)
  {
    return id !== SHARED_TMP_ID && !files[id].contentLoaded;
  });

  await Promise.all(pending.map(function(id)
  {
    return Platform.readWorkFile(workFolderRoot, id)
            .then(function(c) { files[id].content = c; files[id].contentLoaded = true; })
            .catch(function() {});
  }));
}

function relBuildGraph()
{
  var ids = Object.keys(files).filter(function(id) { return id !== SHARED_TMP_ID; });

  var nodes = ids.map(function(id, i)
  {
    var angle = (i / Math.max(ids.length, 1)) * Math.PI * 2;

    return {
      id: id,
      name: files[id].name,
      type: files[id].type,
      x: Math.cos(angle) * 160 + (Math.random() - 0.5) * 40,
      y: Math.sin(angle) * 160 + (Math.random() - 0.5) * 40,
      vx: 0, vy: 0,
      degree: 0
    };
  });

  var nodeById = {};
  nodes.forEach(function(n) { nodeById[n.id] = n; });

  // Edge key "src|tgt" (undirected: smaller id first) → { a, b, kind }.
  // Explicit links win over mere mentions when both exist for a pair.
  var edgeMap = {};

  function addEdge(src, tgt, kind)
  {
    if (src === tgt || !nodeById[src] || !nodeById[tgt])
      return;

    var key = src < tgt ? src + '|' + tgt : tgt + '|' + src;

    if (!edgeMap[key])
      edgeMap[key] = { a: nodeById[src], b: nodeById[tgt], kind: kind };
    else if (kind === 'link')
      edgeMap[key].kind = 'link';
  }

  ids.forEach(function(id)
  {
    relExplicitTargets(files[id].content).forEach(function(tgt) { addEdge(id, tgt, 'link'); });
    computeMentionsListFor(id).forEach(function(tgt) { addEdge(id, tgt, 'mention'); });
  });

  var edges = Object.keys(edgeMap).map(function(k) { return edgeMap[k]; });

  edges.forEach(function(e) { e.a.degree++; e.b.degree++; });
  nodes.forEach(function(n)
  {
    n.baseR = Math.min(7 + Math.sqrt(n.degree) * 3.2, 18);
    n.r = n.baseR * relPrefs.size / 100;
  });

  return { nodes: nodes, edges: edges };
}

async function openRelationsGraph()
{
  document.getElementById('relations-modal').style.display = 'flex';

  loadRelPrefs();
  document.getElementById('rel-spacing').value = relPrefs.spacing;
  document.getElementById('rel-size').value    = relPrefs.size;

  await relEnsureContentsLoaded();

  var graph  = relBuildGraph(),
      wrap   = document.querySelector('.relations-canvas-wrap'),
      canvas = document.getElementById('relations-canvas');

  document.getElementById('relations-empty').style.display = graph.edges.length ? 'none' : '';

  // Legend: file types actually present, plus the two edge kinds
  var typesPresent = {};
  graph.nodes.forEach(function(n) { typesPresent[n.type] = true; });

  document.getElementById('relations-legend').innerHTML =
    Object.keys(typesPresent).map(function(t)
    {
      return '<span class="rel-legend-item"><span class="rel-legend-dot" style="background:' + (REL_TYPE_COLORS[t] || '#9ca3af') + '"></span>' + escHtml(REL_TYPE_LABELS[t] || t) + '</span>';
    }).join('') +
    '<span class="rel-legend-item"><span class="rel-legend-line"></span>Linked</span>' +
    '<span class="rel-legend-item"><span class="rel-legend-line rel-legend-line-dim"></span>Mentions</span>';

  var css = getComputedStyle(document.documentElement);

  relState =
  {
    nodes: graph.nodes,
    edges: graph.edges,
    canvas: canvas,
    ctx: canvas.getContext('2d'),
    wrap: wrap,
    scale: 1, tx: 0, ty: 0,
    alpha: 1,
    dragNode: null,
    panning: false,
    pointerDown: null,
    hovered: null,
    raf: 0,
    textColor:  (css.getPropertyValue('--text')  || '#ddd').trim(),
    edgeColor:  (css.getPropertyValue('--text3') || '#888').trim(),
    haloColor:  (css.getPropertyValue('--bg')    || '#16181d').trim(),
    resize: function() { relResizeCanvas(); }
  };

  window.addEventListener('resize', relState.resize);
  relAttachPointerHandlers();
  relResizeCanvas();
  relState.raf = requestAnimationFrame(relTick);
}

function closeRelationsGraph()
{
  document.getElementById('relations-modal').style.display = 'none';

  if (!relState)
    return;

  cancelAnimationFrame(relState.raf);
  window.removeEventListener('resize', relState.resize);
  relState = null;
}

function relResizeCanvas()
{
  if (!relState)
    return;

  var dpr = window.devicePixelRatio || 1,
      w = relState.wrap.clientWidth,
      h = relState.wrap.clientHeight;

  relState.canvas.width  = w * dpr;
  relState.canvas.height = h * dpr;
  relState.canvas.style.width  = w + 'px';
  relState.canvas.style.height = h + 'px';
  relState.dpr = dpr;
  relState.w = w;
  relState.h = h;
}

// One simulation step: pairwise repulsion, spring along edges, weak pull to
// the center. Standard velocity-Verlet-ish integration with damping; alpha
// cools the whole system down so the layout settles instead of jittering.
function relStep()
{
  var s = relState,
      nodes = s.nodes,
      edges = s.edges,
      alpha = s.alpha,
      spacing = relPrefs.spacing,
      spacingF = spacing / 110,
      repelK = 1600 * spacingF * spacingF,
      cutoff2 = Math.pow(spacing * 3.6, 2), // beyond ~3.6× spacing repulsion is negligible
      i, j, a, b, dx, dy, d2, d, f;

  for (i = 0; i < nodes.length; i++)
    for (j = i + 1; j < nodes.length; j++)
    {
      a = nodes[i]; b = nodes[j];
      dx = b.x - a.x; dy = b.y - a.y;
      d2 = dx * dx + dy * dy || 1;

      if (d2 > cutoff2) continue;

      f = (repelK / d2) * alpha;
      d = Math.sqrt(d2);
      dx /= d; dy /= d;
      a.vx -= dx * f; a.vy -= dy * f;
      b.vx += dx * f; b.vy += dy * f;
    }

  for (i = 0; i < edges.length; i++)
  {
    a = edges[i].a; b = edges[i].b;
    dx = b.x - a.x; dy = b.y - a.y;
    d = Math.sqrt(dx * dx + dy * dy) || 1;
    f = (d - spacing) * 0.02 * alpha * (edges[i].kind === 'link' ? 1.6 : 1);
    dx /= d; dy /= d;
    a.vx += dx * f; a.vy += dy * f;
    b.vx -= dx * f; b.vy -= dy * f;
  }

  for (i = 0; i < nodes.length; i++)
  {
    a = nodes[i];
    a.vx -= a.x * 0.004 * alpha;
    a.vy -= a.y * 0.004 * alpha;

    if (a === s.dragNode) { a.vx = 0; a.vy = 0; continue; }

    a.vx *= 0.85; a.vy *= 0.85;
    a.x += a.vx; a.y += a.vy;
  }

  s.alpha = Math.max(alpha * 0.985, 0.02);
}

function relDraw()
{
  var s = relState,
      ctx = s.ctx;

  ctx.setTransform(s.dpr, 0, 0, s.dpr, 0, 0);
  ctx.clearRect(0, 0, s.w, s.h);
  ctx.translate(s.w / 2 + s.tx, s.h / 2 + s.ty);
  ctx.scale(s.scale, s.scale);

  var i, e, n,
      hover = s.hovered;

  for (i = 0; i < s.edges.length; i++)
  {
    e = s.edges[i];

    var touched = hover && (e.a === hover || e.b === hover);

    ctx.globalAlpha = e.kind === 'link' ? (touched ? 0.95 : 0.55) : (touched ? 0.7 : 0.22);
    ctx.strokeStyle = touched ? (REL_TYPE_COLORS[hover.type] || s.edgeColor) : s.edgeColor;
    ctx.lineWidth = (e.kind === 'link' ? 1.8 : 1) / s.scale + (e.kind === 'link' ? 0.6 : 0.2);
    ctx.beginPath();
    ctx.moveTo(e.a.x, e.a.y);
    ctx.lineTo(e.b.x, e.b.y);
    ctx.stroke();
  }

  ctx.globalAlpha = 1;

  for (i = 0; i < s.nodes.length; i++)
  {
    n = s.nodes[i];

    var dimmed = hover && n !== hover && !s.edges.some(function(e) { return (e.a === hover && e.b === n) || (e.b === hover && e.a === n); });

    ctx.globalAlpha = dimmed ? 0.25 : 1;
    ctx.beginPath();
    ctx.arc(n.x, n.y, n.r, 0, Math.PI * 2);
    ctx.fillStyle = REL_TYPE_COLORS[n.type] || '#9ca3af';
    ctx.fill();

    if (n === hover)
    {
      ctx.lineWidth = 2 / s.scale;
      ctx.strokeStyle = s.textColor;
      ctx.stroke();
    }

    // Labels stay a constant on-screen size regardless of zoom (divide by
    // scale since the canvas is scaled), grow a little with the node-size
    // slider, and get a halo in the background color so they stay readable
    // when they cross edges or other labels.
    var fontPx = Math.min(11 * (0.75 + (relPrefs.size / 100) * 0.35), 17) / s.scale,
        label = n.name.length > 24 ? n.name.slice(0, 22) + '…' : n.name;

    ctx.font = (n === hover ? 'bold ' : '') + fontPx.toFixed(2) + 'px sans-serif';
    ctx.textAlign = 'center';
    ctx.lineJoin = 'round';
    ctx.lineWidth = 3.5 / s.scale;
    ctx.strokeStyle = s.haloColor;
    ctx.strokeText(label, n.x, n.y + n.r + 14 / s.scale);
    ctx.fillStyle = s.textColor;
    ctx.fillText(label, n.x, n.y + n.r + 14 / s.scale);
  }

  ctx.globalAlpha = 1;
}

function relTick()
{
  if (!relState)
    return;

  if (relState.alpha > 0.021 || relState.dragNode)
    relStep();

  relDraw();
  relState.raf = requestAnimationFrame(relTick);
}

// Screen px → graph coordinates
function relToGraph(px, py)
{
  var s = relState,
      rect = s.canvas.getBoundingClientRect();

  return {
    x: (px - rect.left - s.w / 2 - s.tx) / s.scale,
    y: (py - rect.top  - s.h / 2 - s.ty) / s.scale
  };
}

function relNodeAt(px, py)
{
  var p = relToGraph(px, py),
      nodes = relState.nodes;

  for (var i = nodes.length - 1; i >= 0; i--)
  {
    var dx = p.x - nodes[i].x,
        dy = p.y - nodes[i].y,
        hit = nodes[i].r + 4 / relState.scale;

    if (dx * dx + dy * dy <= hit * hit)
      return nodes[i];
  }

  return null;
}

function relAttachPointerHandlers()
{
  var canvas = relState.canvas;

  canvas.onpointerdown = function(e)
  {
    if (!relState) return;
    canvas.setPointerCapture(e.pointerId);
    var n = relNodeAt(e.clientX, e.clientY);
    relState.pointerDown = { x: e.clientX, y: e.clientY, moved: false, node: n };

    if (n) { relState.dragNode = n; relState.alpha = Math.max(relState.alpha, 0.3); }
    else   relState.panning = true;
  };

  canvas.onpointermove = function(e)
  {
    if (!relState) return;

    var pd = relState.pointerDown;

    if (pd)
    {
      if (Math.abs(e.clientX - pd.x) + Math.abs(e.clientY - pd.y) > 4)
        pd.moved = true;

      if (relState.dragNode)
      {
        var p = relToGraph(e.clientX, e.clientY);
        relState.dragNode.x = p.x;
        relState.dragNode.y = p.y;
        relState.alpha = Math.max(relState.alpha, 0.25);
      }
      else if (relState.panning)
      {
        relState.tx += e.movementX;
        relState.ty += e.movementY;
      }
    }
    else
    {
      relState.hovered = relNodeAt(e.clientX, e.clientY);
      canvas.style.cursor = relState.hovered ? 'pointer' : 'grab';
    }
  };

  canvas.onpointerup = function(e)
  {
    if (!relState) return;

    var pd = relState.pointerDown;

    // A press-and-release on a node without dragging opens that file
    if (pd && pd.node && !pd.moved)
    {
      var id = pd.node.id;
      closeRelationsGraph();
      openFile(id);
      return;
    }

    relState.pointerDown = null;
    relState.dragNode = null;
    relState.panning = false;
  };

  canvas.onwheel = function(e)
  {
    if (!relState) return;
    e.preventDefault();

    var factor = e.deltaY < 0 ? 1.12 : 1 / 1.12,
        next = Math.min(Math.max(relState.scale * factor, 0.25), 4);

    // Zoom around the cursor: keep the graph point under it fixed
    var rect = relState.canvas.getBoundingClientRect(),
        cx = e.clientX - rect.left - relState.w / 2,
        cy = e.clientY - rect.top  - relState.h / 2;

    relState.tx = cx - (cx - relState.tx) * (next / relState.scale);
    relState.ty = cy - (cy - relState.ty) * (next / relState.scale);
    relState.scale = next;
  };
}

// ── FILE HISTORY (one entry per "close" - switching away, or app exit) ──
//
// Each entry stores only what changed since the previous entry, not a full
// copy of the document - a list of "hunks", each one { line, type, removed,
// added }, where line is the 1-based starting line in the OLD content, type
// is 'added'/'removed'/'modified', and removed/added are the actual old/new
// lines (so "what changed" is readable straight off the stored entry, not
// just its line numbers). A version's full text is reconstructed by
// replaying every entry's hunks in order, starting from ''.

function linesOf(text)
{
  return (text || '').split('\n');
}

// Same LCS-based approach as computeWordDiff, but over lines instead of
// whitespace tokens, then grouped into contiguous add/remove/modify hunks.
function computeLineHunks(oldText, newText)
{
  var a = linesOf(oldText),
      b = linesOf(newText),
      n = a.length,
      m = b.length;

  var dp = new Array(n + 1);

  for (var i = 0; i <= n; i++)
    dp[i] = new Uint32Array(m + 1);

  for (i = n - 1; i >= 0; i--)
    for (var j = m - 1; j >= 0; j--)
      dp[i][j] = (a[i] === b[j]) ? (dp[i+1][j+1] + 1) : Math.max(dp[i+1][j], dp[i][j+1]);

  var ops = [];

  i = 0;
  var j = 0;

  while (i < n && j < m)
  {
    if (a[i] === b[j])
    {
      ops.push({ type: 'same', value: a[i] });
      i++; j++;
    }
    else if (dp[i+1][j] >= dp[i][j+1])
    {
      ops.push({ type: 'removed', value: a[i] });
      i++;
    }
    else
    {
      ops.push({ type: 'added', value: b[j] });
      j++;
    }
  }

  while (i < n) { ops.push({ type: 'removed', value: a[i] }); i++; }
  while (j < m) { ops.push({ type: 'added', value: b[j] }); j++; }

  var hunks = [],
      oldLineNo = 1,
      k = 0;

  while (k < ops.length)
  {
    if (ops[k].type === 'same')
    {
      oldLineNo++;
      k++;
      continue;
    }

    var startLine = oldLineNo,
        removed = [],
        added = [];

    while (k < ops.length && ops[k].type !== 'same')
    {
      if (ops[k].type === 'removed')
      {
        removed.push(ops[k].value);
        oldLineNo++;
      }
      else
        added.push(ops[k].value);

      k++;
    }

    hunks.push({ line: startLine, type: (removed.length && added.length) ? 'modified' : (added.length ? 'added' : 'removed'), removed: removed, added: added });
  }

  return hunks;
}

// Replays one entry's hunks against a base text to produce the next version.
function applyLineHunks(content, hunks)
{
  var lines = linesOf(content),
      shift = 0;

  hunks.forEach
  (
    function(h)
    {
      var idx = (h.line - 1) + shift;

      lines.splice.apply(lines, [idx, h.removed.length].concat(h.added));
      shift += h.added.length - h.removed.length;
    }
  );

  return lines.join('\n');
}

// Rebuilds the full text of file.history[uptoIndex] by replaying every
// entry from the start. Entries from before this format change (which
// stored a full snapshot as entry.content) are handled as a hard reset of
// the running text, so old and new-format entries can coexist in one chain.
function reconstructHistoryContent(file, uptoIndex)
{
  var content = '';

  for (var i = 0; i <= uptoIndex; i++)
  {
    var entry = file.history[i];

    content = entry.hunks ? applyLineHunks(content, entry.hunks) : (entry.content || '');
  }

  return content;
}

// Records what changed since the last entry, skipping it if nothing did (so
// just opening and closing a file without editing it doesn't pollute the
// history with no-op entries).
function recordFileHistory(id)
{
  var file = files[id];

  if (!file)
    return;

  // Shared files aren't ours: no history sidecar in someone else's workspace.
  if (id === SHARED_TMP_ID)
    return;

  var content = file.content || '';

  if (!file.history)
    file.history = [];

  var lastContent = reconstructHistoryContent(file, file.history.length - 1);

  if (lastContent === content)
    return;

  file.history.push({ timestamp: Date.now(), hunks: computeLineHunks(lastContent, content) });

  if (file.history.length > 50)
  {
    // Drop the two oldest entries and replace them with a single re-based
    // "diff from empty" entry covering the same ground, so the replay chain
    // stays intact instead of losing its anchor at index 0.
    var mergedContent = reconstructHistoryContent(file, 1);

    file.history.splice(0, 2, { timestamp: file.history[1].timestamp, hunks: computeLineHunks('', mergedContent) });
  }

  if (workFolderRoot)
  {
    Platform.writeWorkFile(workFolderRoot, id + '.history.json', JSON.stringify(file.history))
      .catch(function(e){ console.warn('History sidecar write error', e); });
  }
  else
    saveToStorage();
}

window.addEventListener
(
  'beforeunload',
  function()
  {
    if (currentFileId)
      recordFileHistory(currentFileId);
  }
);

document.addEventListener('click', function(e)
{
  var tPop = document.getElementById('travel-popover');
  if (tPop && tPop.style.display !== 'none')
  {
    if (!tPop.contains(e.target) && e.target.id !== 'map-travel-result')
      closeTravelPopover();
  }
  var pPop = document.getElementById('pin-popover');
  if (pPop && pPop.style.display !== 'none')
  {
    if (!pPop.contains(e.target) && !e.target.closest('.map-pin'))
      closePinPopover();
  }
});

async function openHistoryModal()
{
  if (!currentFileId || !files[currentFileId])
    return;

  var file = files[currentFileId];

  if (workFolderRoot && !file.historyLoaded)
  {
    try
    {
      var raw = await Platform.readWorkFile(workFolderRoot, currentFileId + '.history.json');
      file.history = JSON.parse(raw) || [];
    }
    catch(e)
    {
      file.history = file.history || [];
    }

    file.historyLoaded = true;
  }

  renderHistoryModal();
  document.getElementById('history-modal').classList.add('open');
}

function closeHistoryModal()
{
  document.getElementById('history-modal').classList.remove('open');
}

document.getElementById('history-modal').addEventListener
(
  'click',
  function(e)
  {
    if (e.target === document.getElementById('history-modal'))
      closeHistoryModal();
  }
);

// Sheets/Notebook don't read as plain text, so their preview is a short
// structural summary instead of a text snippet.
function historyPreviewFor(type, content)
{
  content = content || '';

  if (type === 'sheet')
  {
    var cellCount = content.split('\n').filter(function(line){ return /^[A-Za-z]+\d+=/.test(line); }).length;
    return cellCount + ' cell' + (cellCount === 1 ? '' : 's') + ' with data';
  }

  if (type === 'notebook')
  {
    try
    {
      var pageCount = parseNotebookContent(content).pages.length;
      return pageCount + ' page' + (pageCount === 1 ? '' : 's');
    }
    catch(e)
    {
      return 'Notebook snapshot';
    }
  }

  var flat = content.replace(/\s+/g, ' ').trim();
  return flat ? (flat.length > 120 ? flat.slice(0, 120) + '…' : flat) : '(empty)';
}

function renderHistoryModal()
{
  var file = files[currentFileId],
      list = document.getElementById('history-list'),
      history = file.history || [];

  if (!history.length)
  {
    list.innerHTML = '<div class="history-empty">No history yet — a version is saved each time you switch away from this document or close the app.</div>';
    return;
  }

  var html = '',
      reconstructed = new Array(history.length),
      running = '';

  for (var i = 0; i < history.length; i++)
  {
    var entry = history[i];

    running = entry.hunks ? applyLineHunks(running, entry.hunks) : (entry.content || '');
    reconstructed[i] = running;
  }

  for (i = history.length - 1; i >= 0; i--)
  {
    entry = history[i];

    var when = new Date(entry.timestamp).toLocaleString(),
        preview = historyPreviewFor(file.type, reconstructed[i]),
        changeSummary = historyChangeSummary(entry);

    html +=  '<div class="history-item">' +
                '<div class="history-item-info">' +
                  '<div class="history-item-time">' + escHtml(when) + (changeSummary ? ' <span class="history-item-changes">' + changeSummary + '</span>' : '') + '</div>' +
                  '<div class="history-item-preview">' + escHtml(preview) + '</div>' +
                '</div>' +
                '<div class="history-item-actions">' +
                  '<button type="button" class="btn-cancel history-compare-btn" onclick="compareHistoryEntry(' + i + ')">Compare</button>' +
                  '<button type="button" class="btn-cancel history-restore-btn" onclick="restoreHistoryEntry(' + i + ')">Restore</button>' +
                '</div>' +
              '</div>';
  }

  list.innerHTML = html;
}

// A compact "+N -M line(s)" badge built straight from the stored hunks -
// entries from before this format change (full-snapshot entries) have no
// hunks, so they just don't get a badge.
function historyChangeSummary(entry)
{
  if (!entry.hunks)
    return '';

  var added = 0,
      removed = 0;

  entry.hunks.forEach(function(h){ added += h.added.length; removed += h.removed.length; });

  if (!added && !removed)
    return '';

  var parts = [];

  if (added)
    parts.push('<span class="diff-added">+' + added + '</span>');

  if (removed)
    parts.push('<span class="diff-removed">−' + removed + '</span>');

  return parts.join(' ');
}

// Splits on whitespace runs while keeping them as their own tokens, so
// joining the tokens back together reconstructs the text exactly - the
// standard tokenization for a word-level (rather than line-level) diff,
// which reads much better than line diffing on wrapped markdown paragraphs.
function tokenizeForDiff(text)
{
  return (text || '').split(/(\s+)/).filter(function(t){ return t.length > 0; });
}

// Classic LCS-based diff over the token arrays. Returns a flat op list
// (same/removed/added) in document order, or null if the documents are too
// large for an O(n*m) table to be worth computing.
function computeWordDiff(oldText, newText)
{
  var a = tokenizeForDiff(oldText),
      b = tokenizeForDiff(newText),
      n = a.length,
      m = b.length;

  if (n * m > 4000000)
    return null;

  var dp = new Array(n + 1);

  for (var i = 0; i <= n; i++)
    dp[i] = new Uint32Array(m + 1);

  for (i = n - 1; i >= 0; i--)
    for (var j = m - 1; j >= 0; j--)
      dp[i][j] = (a[i] === b[j]) ? (dp[i+1][j+1] + 1) : Math.max(dp[i+1][j], dp[i][j+1]);

  var ops = [];

  i = 0;
  var j = 0;

  while (i < n && j < m)
  {
    if (a[i] === b[j])
    {
      ops.push({ type: 'same', value: a[i] });
      i++; j++;
    }
    else if (dp[i+1][j] >= dp[i][j+1])
    {
      ops.push({ type: 'removed', value: a[i] });
      i++;
    }
    else
    {
      ops.push({ type: 'added', value: b[j] });
      j++;
    }
  }

  while (i < n) { ops.push({ type: 'removed', value: a[i] }); i++; }
  while (j < m) { ops.push({ type: 'added', value: b[j] }); j++; }

  return ops;
}

function renderWordDiff(oldText, newText, oldPaneId, newPaneId)
{
  var ops = computeWordDiff(oldText, newText),
      oldPane = document.getElementById(oldPaneId || 'diff-old-pane'),
      newPane = document.getElementById(newPaneId || 'diff-new-pane');

  if (!ops)
  {
    oldPane.textContent = oldText;
    newPane.textContent = newText;
    return;
  }

  var oldHtml = '',
      newHtml = '';

  ops.forEach(function(op)
  {
    var escaped = escHtml(op.value);

    if (op.type === 'same')
    {
      oldHtml += escaped;
      newHtml += escaped;
    }
    else if (op.type === 'removed')
      oldHtml += '<span class="diff-removed">' + escaped + '</span>';
    else
      newHtml += '<span class="diff-added">' + escaped + '</span>';
  });

  oldPane.innerHTML = oldHtml || '<em>(empty)</em>';
  newPane.innerHTML = newHtml || '<em>(empty)</em>';
}

function compareHistoryEntry(index)
{
  var file = files[currentFileId],
      entry = file && file.history && file.history[index];

  if (!entry)
    return;

  historyDiffIndex = index;

  var oldContent = index > 0 ? reconstructHistoryContent(file, index - 1) : '';
  var newContent = reconstructHistoryContent(file, index);

  var oldLabel = index > 0
    ? new Date(file.history[index - 1].timestamp).toLocaleString()
    : 'Before';

  document.getElementById('diff-old-header').textContent = oldLabel;
  document.getElementById('diff-new-header').textContent = new Date(entry.timestamp).toLocaleString();

  renderWordDiff(oldContent, newContent);

  closeHistoryModal();
  document.getElementById('history-diff-modal').classList.add('open');
}

function closeHistoryDiffModal()
{
  document.getElementById('history-diff-modal').classList.remove('open');
}

document.getElementById('history-diff-modal').addEventListener
(
  'click',
  function(e)
  {
    if (e.target === document.getElementById('history-diff-modal'))
      closeHistoryDiffModal();
  }
);

async function restoreFromDiff()
{
  var index = historyDiffIndex;

  closeHistoryDiffModal();
  await restoreHistoryEntry(index);
}

async function restoreHistoryEntry(index)
{
  var file = files[currentFileId],
      entry = file && file.history && file.history[index];

  if (!entry)
    return;

  if (!confirm('Restore this version? Your current content will be saved to history first, so you can still get back to it.'))
    return;

  var restoredContent = reconstructHistoryContent(file, index);

  recordFileHistory(currentFileId);
  file.content = restoredContent;

  closeHistoryModal();

  if (file.type === 'doc')
    loadDocFile(file);
  else if (file.type === 'graph')
    loadGraphFile(file);
  else if (file.type === 'notebook')
    loadNotebookFile(file);
  else
    loadSheetFile(file);

  if (workFolderRoot)
  {
    try { await Platform.writeWorkFile(workFolderRoot, currentFileId, file.content || ''); }
    catch(e) { console.warn('Work folder write error', e); }
  }
  else
    saveToStorage();

  renderFileList();
}

async function openFile(id)
{
  if (!files[id])
    return;

  if (id !== SHARED_TMP_ID)
    exitSharedMode();

  if (currentFileId && currentFileId !== id && currentFileId !== SHARED_TMP_ID)
    recordFileHistory(currentFileId);

  currentFileId = id;

  if (workFolderRoot && !files[id].contentLoaded)
  {
    try
    {
      files[id].content       = await Platform.readWorkFile(workFolderRoot, id);
      files[id].contentLoaded = true;
    }
    catch(e)
    {
      console.warn('Work folder read error', e);
      alert('Could not open this file.');
      return;
    }
  }

  const file = files[id];

  switchAppType(file.type, false);
  renderFileList();
  noteTabOpened(id);
  closeMobileSidebar(); // picking a file dismisses the phone drawer

  if (file.type === 'doc')
    loadDocFile(file);

  else if (file.type === 'graph')
    loadGraphFile(file);

  else if (file.type === 'notebook')
    loadNotebookFile(file);

  else if (file.type === 'glossary')
    loadGlossaryFile(file);

  else if (file.type === 'calendar')
    loadCalendarFile(file);

  else if (file.type === 'economy')
    loadEconomyFile(file);

  else if (file.type === 'bestiary')
    loadBestiaryFile(file);

  else
    loadSheetFile(file);

  updateCommentsUI();
  updateLiveSync();

  // Fire-and-forget: update the shared backlinks index whenever a file is opened
  updateBacklinksForFile(id);
}

function clearActiveEditors()
{
  document.getElementById('doc-title-input').value = '';
  document.getElementById('editor').value = '';
  document.getElementById('toolbar-title').innerHTML = '<span>No document open</span>';
  updatePreview();
  updateStatus();
}

async function deleteFile(e, id)
{
  e.stopPropagation();

  if (!confirm('Delete "' + files[id].name + '"?'))
    return;

  if (workFolderRoot)
  {
    try
    {
      await Platform.deleteWorkEntry(workFolderRoot, id, false);
      Platform.deleteWorkEntry(workFolderRoot, id + '.history.json', false).catch(function(){});
    }
    catch(err)
    {
      console.warn('Work folder delete error', err);
      alert('Could not delete the file.');
      return;
    }

    if (currentFileId === id)
    {
      currentFileId = null;
      clearActiveEditors();
    }

    await loadWorkFolderTree();
    return;
  }

  delete files[id];

  if (currentFileId === id)
  {
    currentFileId = null;
    clearActiveEditors();
  }

  saveToStorage();
  pruneOpenTabs();
  renderFileList();
}

async function deleteFolderEntry(e, path)
{
  e.stopPropagation();

  if (!confirm('Delete folder "' + folders[path].name + '" and everything inside it?'))
    return;

  try
  {
    await Platform.deleteWorkEntry(workFolderRoot, path, true);
  }
  catch(err)
  {
    console.warn('Work folder delete error', err);
    alert('Could not delete the folder.');
    return;
  }

  if (currentFileId && (currentFileId === path || currentFileId.indexOf(path + '/') === 0))
  {
    currentFileId = null;
    clearActiveEditors();
  }

  await loadWorkFolderTree();
}

// ── WORK FOLDER (on-disk storage mirrored into the sidebar) ──

// Rebuilds `files`/`folders` from disk. Used as the single source of truth
// refresh after any structural change (create/delete/move) rather than
// hand-patching the in-memory tree - simpler and safe at the scale of a
// person clicking buttons rather than a hot path.
async function loadWorkFolderTree()
{
  if (!workFolderRoot)
    return;

  const preserved = {};

  Object.keys(files).forEach(function(id)
  {
    if (files[id].contentLoaded)
      preserved[id] = files[id];
  });

  let entries;

  try
  {
    entries = await Platform.listWorkFolder(workFolderRoot);
  }
  catch(e)
  {
    console.warn('Work folder read error', e);
    alert('Could not read the work folder. It may have been moved or deleted.');
    return;
  }

  const newFiles = {},
        newFolders = {};

  function walk(list, parent)
  {
    list.forEach(function(entry)
    {
      if (entry.isDir)
      {
        newFolders[entry.relPath] = { name: entry.name, parent: parent };
        walk(entry.children, entry.relPath);
        return;
      }

      // Sync sidecars (e.g. the _lktpl.json custom-templates store) are in
      // the server listing so the desktop app can sync them, but they aren't
      // documents — keep them out of the file tree.
      if (/\.json$/i.test(entry.name))
        return;

      const dot = entry.name.lastIndexOf('.'),
            ext = (dot === -1) ? '' : entry.name.slice(dot + 1).toLowerCase(),
            baseName = (dot === -1) ? entry.name : entry.name.slice(0, dot),
            type = ext === 'mds' ? 'sheet' : ext === 'mdg' ? 'graph' : ext === 'mdn' ? 'notebook' : ext === 'mdl' ? 'glossary' : ext === 'mdc' ? 'calendar' : ext === 'mde' ? 'economy' : ext === 'mdb' ? 'bestiary' : 'doc',
            kept = preserved[entry.relPath];

      newFiles[entry.relPath] =
      {
        name: baseName,
        type: type,
        folder: parent,
        modified: entry.modified,
        content: kept ? kept.content : undefined,
        contentLoaded: !!kept,
        lkMentions:    kept ? (kept.lkMentions || []) : []
      };
    });
  }

  walk(entries, '');

  files = newFiles;
  folders = newFolders;

  if (currentFileId && !files[currentFileId])
    currentFileId = null;

  pruneOpenTabs(); // ids are rel paths; anything renamed/moved/deleted on disk is gone
  renderFileList();
}

async function createWorkFile(name, type, content)
{
  const ext = fileExtensionFor(type),
        base = sanitizeFileName(name) || 'Untitled',
        relPath = uniqueRelPath('', base, ext);

  try
  {
    await Platform.writeWorkFile(workFolderRoot, relPath, content !== undefined ? content : defaultContentForType(type, name));
  }
  catch(e)
  {
    console.warn('Work folder create error', e);
    alert('Could not create the file.');
    return null;
  }

  await loadWorkFolderTree();

  return relPath;
}

// Single funnel for persisting a file's content and/or its on-disk name, used
// both by the autosave debounce (acting on currentFileId) and by sidebar
// inline-rename (which may target a file that isn't the open one). A rename
// is implemented as a move within the same parent folder.
async function persistFileEntry(id)
{
  if (!files[id])
    return;

  // Files opened from "Shared with me" never touch my own workspace: they're
  // written back through the share (edit permission) or not at all.
  if (id === SHARED_TMP_ID)
  {
    if (sharedCtx && sharedCtx.permission === 'edit')
    {
      try
      {
        var sharedContent = files[id].content || '';
        if (sharedCtx.link)
          await Platform.writeLinkFile(sharedCtx.link, sharedCtx.subPath, sharedContent);
        else
          await Platform.writeSharedFile(sharedCtx.shareId, sharedCtx.subPath, sharedContent);
        if (id === liveSyncFileId) { liveSyncBase = sharedContent; liveSyncServerHash = null; }
      }
      catch(e) { console.warn('Shared write error', e); }
    }
    return;
  }

  if (!workFolderRoot)
  {
    await updateBacklinksForFile(id);
    saveToStorage();
    renderFileList();
    return;
  }

  const file = files[id],
        ext = fileExtensionFor(file.type),
        desiredRelPath = uniqueRelPath(file.folder || '', sanitizeFileName(file.name) || 'Untitled', ext, id);

  let activeId = id;

  if (desiredRelPath !== id)
  {
    try
    {
      await Platform.moveWorkEntry(workFolderRoot, id, desiredRelPath);

      files[desiredRelPath] = file;
      delete files[id];

      if (currentFileId === id)
        currentFileId = desiredRelPath;

      remapOpenTab(id, desiredRelPath);

      activeId = desiredRelPath;

      // Best-effort: keep the version-history sidecar attached to the file
      // it documents. Fails quietly if there's no history yet.
      Platform.moveWorkEntry(workFolderRoot, id + '.history.json', desiredRelPath + '.history.json').catch(function(){});
    }
    catch(e)
    {
      console.warn('Work folder rename error', e);
    }
  }

  try
  {
    await Platform.writeWorkFile(workFolderRoot, activeId, files[activeId].content || '');
    if (activeId === liveSyncFileId) { liveSyncBase = files[activeId].content || ''; liveSyncServerHash = null; }
    await updateBacklinksForFile(activeId);
  }
  catch(e)
  {
    console.warn('Work folder write error', e);
  }

  renderFileList();

  if (currentFileId === activeId)
  {
    if (file.type === 'doc')
      document.getElementById('doc-title-input').value = file.name;

    else if (file.type === 'graph')
      document.getElementById('graph-title-input').value = file.name;
  }
}

async function renameFolderEntry(path, newName)
{
  const folder = folders[path],
        base = sanitizeFileName(newName) || folder.name,
        desiredRelPath = uniqueFolderRelPath(folder.parent, base, path);

  if (desiredRelPath === path)
  {
    renderFileList();
    return;
  }

  try
  {
    await Platform.moveWorkEntry(workFolderRoot, path, desiredRelPath);
  }
  catch(e)
  {
    console.warn('Work folder rename error', e);
  }

  expandedFolders.delete(path);
  expandedFolders.add(desiredRelPath);

  await loadWorkFolderTree();
}

function startInlineRename(e, nameEl, id)
{
  e.stopPropagation();

  const current = files[id] ? files[id].name : (folders[id] ? folders[id].name : '');

  nameEl.innerHTML = '';

  const input = document.createElement('input');
  input.className = 'inline-rename-input';
  input.value = current;

  nameEl.appendChild(input);
  input.focus();
  input.select();

  let committed = false;

  function commit()
  {
    if (committed)
      return;

    committed = true;
    finishInlineRename(id, input.value.trim());
  }

  input.addEventListener('blur', commit);
  input.addEventListener('keydown', function(ke)
  {
    if (ke.key === 'Enter')
    {
      ke.preventDefault();
      input.blur();
    }
    else if (ke.key === 'Escape')
    {
      committed = true;
      renderFileList();
    }
  });
}

function finishInlineRename(id, newName)
{
  if (files[id])
  {
    if (newName && newName !== files[id].name)
    {
      files[id].name = newName;
      files[id].modified = Date.now();
      persistFileEntry(id);
      return;
    }

    renderFileList();
    return;
  }

  if (folders[id])
  {
    if (newName && newName !== folders[id].name)
    {
      renameFolderEntry(id, newName);
      return;
    }

    renderFileList();
  }
}

function handleDragStart(e, id)
{
  draggedEntryId = id;
  e.dataTransfer.effectAllowed = 'move';
  e.dataTransfer.setData('text/plain', id);
  e.currentTarget.classList.add('dragging');
}

function handleDragEnd(e)
{
  e.currentTarget.classList.remove('dragging');
  draggedEntryId = null;

  document.querySelectorAll('.drag-over').forEach(function(el)
  {
    el.classList.remove('drag-over');
  });
}

// Both dragenter and dragover need preventDefault() for a drop to be allowed -
// some engines only honor the dragenter on the very first frame a drag enters
// an element, so it's handled defensively alongside dragover rather than relying
// on dragover alone.
function handleDragEnter(e)
{
  if (!workFolderRoot || !draggedEntryId)
    return;

  e.preventDefault();
  e.stopPropagation();
  e.currentTarget.classList.add('drag-over');
}

function handleDragOver(e)
{
  if (!workFolderRoot || !draggedEntryId)
    return;

  e.preventDefault();
  e.stopPropagation();
  e.dataTransfer.dropEffect = 'move';
  e.currentTarget.classList.add('drag-over');
}

function handleDragLeave(e)
{
  e.stopPropagation();
  e.currentTarget.classList.remove('drag-over');
}

// Core move used by both drag-and-drop and the right-click "Move to" menu.
async function moveEntryToFolder(id, targetFolderPath)
{
  const isDir = !!folders[id];

  if (!isDir && !files[id])
    return false;

  if (isDir && (targetFolderPath === id || targetFolderPath.indexOf(id + '/') === 0))
    return false; // can't move a folder into itself or its own descendant

  const currentParent = isDir ? folders[id].parent : files[id].folder;

  if (currentParent === targetFolderPath)
    return false;

  const toRelPath = isDir
                      ?
                        uniqueFolderRelPath(targetFolderPath, folders[id].name)
                      :
                        uniqueRelPath(targetFolderPath, sanitizeFileName(files[id].name) || 'Untitled', fileExtensionFor(files[id].type));

  try
  {
    await Platform.moveWorkEntry(workFolderRoot, id, toRelPath);

    if (!isDir)
      Platform.moveWorkEntry(workFolderRoot, id + '.history.json', toRelPath + '.history.json').catch(function(){});
  }
  catch(err)
  {
    console.warn('Work folder move error', err);
    alert('Could not move that item.');
    return false;
  }

  if (currentFileId === id)
    currentFileId = toRelPath;

  else if (currentFileId && isDir && currentFileId.indexOf(id + '/') === 0)
    currentFileId = toRelPath + currentFileId.slice(id.length);

  if (isDir)
  {
    // Keep tabs for files inside the moved folder alive under their new paths
    openTabs = openTabs.map(function(t)
    {
      return t.indexOf(id + '/') === 0 ? toRelPath + t.slice(id.length) : t;
    });
    persistOpenTabs();

    expandedFolders.delete(id);
    expandedFolders.add(toRelPath);
  }
  else
    remapOpenTab(id, toRelPath);

  await loadWorkFolderTree();
  return true;
}

async function handleDrop(e, targetFolderPath)
{
  e.preventDefault();
  e.stopPropagation();
  e.currentTarget.classList.remove('drag-over');

  const id = draggedEntryId;
  draggedEntryId = null;

  if (!id || !workFolderRoot)
    return;

  await moveEntryToFolder(id, targetFolderPath);
}

// ── RIGHT-CLICK CONTEXT MENU (folder creation + "Move to" as a drag-and-drop alternative) ──
let contextMenuTarget = null; // { kind: 'file'|'folder'|'root', id }

function openContextMenu(e, kind, id)
{
  e.preventDefault();
  e.stopPropagation();

  if (!workFolderRoot)
    return;

  contextMenuTarget = { kind: kind, id: id };
  renderContextMenu();

  const menu = document.getElementById('context-menu');

  menu.classList.add('open');

  const maxLeft = window.innerWidth - menu.offsetWidth - 8,
        maxTop = window.innerHeight - menu.offsetHeight - 8;

  menu.style.left = Math.max(4, Math.min(e.clientX, maxLeft)) + 'px';
  menu.style.top = Math.max(4, Math.min(e.clientY, maxTop)) + 'px';
}

function closeContextMenu()
{
  document.getElementById('context-menu').classList.remove('open');
  contextMenuTarget = null;
}

document.addEventListener
(
  'click',
  function(e)
  {
    const menu = document.getElementById('context-menu');

    if (menu.classList.contains('open') && !menu.contains(e.target))
      closeContextMenu();
  }
);

function renderContextMenu()
{
  const menu = document.getElementById('context-menu'),
        t = contextMenuTarget;

  let html = '<button class="file-menu-item" onclick="contextMenuNewFolder()">New Folder</button>';

  if (t.kind !== 'root' && !Platform.isNative)
    html += '<div class="file-menu-divider"></div>' +
            '<button class="file-menu-item" onclick="openShareModal()">Share&hellip;</button>';

  if (t.kind !== 'root')
  {
    const currentParent = (t.kind === 'folder') ? folders[t.id].parent : files[t.id].folder;

    html += '<div class="file-menu-divider"></div>';

    if (currentParent !== '')
      html += '<button class="file-menu-item" onclick="contextMenuMoveTo(\'\')">Move to Root</button>';

    Object.keys(folders)
      .filter(function(p)
      {
        if (p === currentParent)
          return false;

        if (t.kind === 'folder' && (p === t.id || p.indexOf(t.id + '/') === 0))
          return false; // can't move a folder into itself or its own descendant

        return true;
      })
      .sort()
      .forEach(function(p)
      {
        html += '<button class="file-menu-item" onclick="contextMenuMoveTo(\'' + escAttr(p) + '\')">Move to &ldquo;' + escHtml(folders[p].name) + '&rdquo;</button>';
      });
  }

  menu.innerHTML = html;
}

async function contextMenuNewFolder()
{
  const t = contextMenuTarget,
        parent = (t.kind === 'folder') ? t.id : (t.kind === 'file' ? (files[t.id].folder || '') : '');

  closeContextMenu();

  const relPath = uniqueFolderRelPath(parent, 'New Folder');

  try
  {
    await Platform.createWorkFolder(workFolderRoot, relPath);
  }
  catch(e)
  {
    console.warn('Work folder create-folder error', e);
    alert('Could not create the folder.');
    return;
  }

  if (parent)
    expandedFolders.add(parent);

  await loadWorkFolderTree();

  const nameEl = document.querySelector('.folder-item[data-id="' + relPath + '"] .file-name');

  if (nameEl)
    startInlineRename({ stopPropagation: function(){} }, nameEl, relPath);
}

async function contextMenuMoveTo(targetFolderPath)
{
  const id = contextMenuTarget.id;

  closeContextMenu();
  await moveEntryToFolder(id, targetFolderPath);
}

function switchAppType(type, rerender)
{
  if (rerender === undefined)
    rerender = true;

  currentAppType = type;
  
  document.getElementById('doc-app').style.display      = (type === 'doc')      ? 'flex' : 'none';
  document.getElementById('sheet-app').style.display    = (type === 'sheet')    ? 'flex' : 'none';
  document.getElementById('graph-app').style.display    = (type === 'graph')    ? 'flex' : 'none';
  document.getElementById('notebook-app').style.display = (type === 'notebook') ? 'flex' : 'none';
  document.getElementById('glossary-app').style.display = (type === 'glossary') ? 'flex' : 'none';
  document.getElementById('calendar-app').style.display = (type === 'calendar') ? 'flex' : 'none';
  document.getElementById('economy-app').style.display  = (type === 'economy')  ? 'flex' : 'none';
  document.getElementById('bestiary-app').style.display = (type === 'bestiary') ? 'flex' : 'none';

  document.querySelectorAll('.app-tab').forEach
  (
    function(appTabs)
    {
        appTabs.classList.remove('active');
    }
  );

  document.getElementById('tab-' + type).classList.add('active');

  if (rerender)
    renderFileList();
}

// ── SHARED DATA-EDIT MODAL ──

var demSaveFn = null;

function openDataModal(title, fieldsHtml, saveFn)
{
  document.getElementById('dem-title').textContent = title;
  document.getElementById('dem-fields').innerHTML  = fieldsHtml;
  demSaveFn = saveFn;
  document.getElementById('data-edit-modal').style.display = 'flex';
  var first = document.getElementById('dem-fields').querySelector('input,select,textarea');
  if (first) setTimeout(function(){ first.focus(); }, 40);
}

function closeDataModal()
{
  document.getElementById('data-edit-modal').style.display = 'none';
  demSaveFn = null;
}

function demSave()
{
  if (demSaveFn) demSaveFn();
}

function genId()
{
  return 'i' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5);
}

// ── TAGS & FACET FILTERS (shared by the JSON data apps) ──
// Every Glossary / Bestiary / Calendar / Economy item can carry free-form
// tags, and each list view gets two facet dropdowns: one for the item's
// natural "type" field (language, category, season, status…) and one for
// tags. The dropdowns are rebuilt from the data on every render so they
// always offer exactly the values that exist.

var DATA_FILE_TYPES = { glossary: 1, bestiary: 1, calendar: 1, economy: 1 };

function parseTagsInput(value)
{
  return String(value || '').split(',').map(function(t){ return t.trim(); }).filter(Boolean);
}

function tagsField(inputId, tags)
{
  return '<label class="field-label" style="margin-top:8px">Tags <span style="color:var(--text3);font-weight:400">(comma separated)</span>' +
         '<input class="modal-input" id="' + inputId + '" value="' + escAttr((tags || []).join(', ')) + '" placeholder="e.g. draft, important…"></label>';
}

// Clickable #tag chips on an item card; clicking one sets that app's tag facet.
function tagChipsHtml(tags, filterFnName)
{
  if (!tags || !tags.length) return '';
  return '<div class="item-tags">' + tags.map(function(t){
    return '<span class="item-tag" onclick="event.stopPropagation();' + filterFnName + '(\'' + escAttr(t) + '\')">#' + escHtml(t) + '</span>';
  }).join('') + '</div>';
}

function collectItemTags(items)
{
  var out = [];
  (items || []).forEach(function(it){
    (it.tags || []).forEach(function(t){ if (out.indexOf(t) === -1) out.push(t); });
  });
  return out;
}

// Rebuilds a facet <select> from the values present in the data, keeping the
// current selection. Returns the effective value: '' when the previously
// selected value no longer exists (so a stale filter can't hide everything).
function fillFacetSelect(selId, values, current, allLabel)
{
  var sel = document.getElementById(selId);
  if (!sel) return current || '';
  var uniq = [];
  (values || []).forEach(function(v){ if (v && uniq.indexOf(v) === -1) uniq.push(v); });
  uniq.sort(function(a, b){ return String(a).localeCompare(String(b)); });
  if (current && uniq.indexOf(current) === -1) current = '';
  sel.innerHTML = '<option value="">' + escHtml(allLabel) + '</option>' +
    uniq.map(function(v){
      return '<option value="' + escAttr(v) + '"' + (v === current ? ' selected' : '') + '>' + escHtml(v) + '</option>';
    }).join('');
  return current;
}

function itemMatchesFacets(item, typeVal, typeKey, tagVal)
{
  if (typeVal && String(item[typeKey] || '') !== typeVal) return false;
  if (tagVal && (item.tags || []).indexOf(tagVal) === -1) return false;
  return true;
}

// File-level tags for the data apps live in the JSON root ({ tags: [...] })
// and are cached on the sidebar entry for filtering without a reparse.
function fileTagsOf(f)
{
  if (!f || !DATA_FILE_TYPES[f.type]) return [];
  if (f.fileTags) return f.fileTags;
  if (f.content === undefined) return [];
  try { f.fileTags = JSON.parse(f.content || '{}').tags || []; }
  catch(e) { f.fileTags = []; }
  return f.fileTags;
}

// ── GLOSSARY ──

var glsData  = null;
var glsTab   = 'words';
var glsQuery = '';
var glsFilter = { words: { type: '', tag: '' }, roots: { type: '', tag: '' } };

function loadGlossaryFile(file)
{
  try { glsData = JSON.parse(file.content || '{}'); }
  catch(e) { glsData = {}; }
  if (!glsData.entries) glsData.entries = [];
  if (!glsData.roots)   glsData.roots   = [];
  // The file's own (user-given) name wins over the embedded title, so a
  // canned default inside the JSON can never rename the file on save.
  document.getElementById('glossary-title-input').value = file.name || glsData.name || '';
  var ft = document.getElementById('gls-file-tags');
  if (ft) ft.value = (glsData.tags || []).join(', ');
  glsQuery = '';
  glsFilter = { words: { type: '', tag: '' }, roots: { type: '', tag: '' } };
  document.getElementById('gls-search').value = '';
  switchGlossaryTab('words');
  renderGlossary();
}

function saveGlossaryData()
{
  if (!currentFileId || !files[currentFileId] || files[currentFileId].type !== 'glossary') return;
  glsData.name = document.getElementById('glossary-title-input').value.trim() || glsData.name || 'Glossary';
  var ft = document.getElementById('gls-file-tags');
  if (ft) glsData.tags = parseTagsInput(ft.value);
  files[currentFileId].name     = glsData.name;
  files[currentFileId].content  = JSON.stringify(glsData, null, 2);
  files[currentFileId].fileTags = glsData.tags || [];
  files[currentFileId].modified = Date.now();
  scheduleSave();
  renderFileList();
}

function onGlossaryTitleChange() { saveGlossaryData(); }

function switchGlossaryTab(tab)
{
  glsTab = tab;
  document.getElementById('gls-words-panel').style.display  = (tab === 'words') ? '' : 'none';
  document.getElementById('gls-roots-panel').style.display  = (tab === 'roots') ? '' : 'none';
  document.getElementById('gls-tab-words').classList.toggle('active', tab === 'words');
  document.getElementById('gls-tab-roots').classList.toggle('active', tab === 'roots');
  renderGlossary();
}

function onGlossarySearch(q) { glsQuery = q; renderGlossary(); }
function onGlsFilterType(v)  { glsFilter[glsTab].type = v; renderGlossary(); }
function onGlsFilterTag(v)   { glsFilter[glsTab].tag  = v; renderGlossary(); }
function glsFilterByTag(t)   { glsFilter[glsTab].tag  = t; renderGlossary(); }
function onGlossaryFileTagsChange() { saveGlossaryData(); }

function renderGlossary()
{
  if (glsTab === 'words') renderGlossaryWords();
  else renderGlossaryRoots();
}

var GLS_LANG_PALETTE = ['#e74c3c','#e67e22','#f1c40f','#2ecc71','#3498db','#9b59b6','#1abc9c','#e91e63','#00b894','#fd79a8'];
function glsLangColor(lang)
{
  if (!lang) return '#9ca3af';
  var h = 0; for (var i = 0; i < lang.length; i++) h = ((h << 5) - h) + lang.charCodeAt(i);
  return GLS_LANG_PALETTE[Math.abs(h) % GLS_LANG_PALETTE.length];
}

var ROOT_TYPE_COLORS = { prefix:'#3498db', suffix:'#e67e22', root:'#2ecc71', infix:'#9b59b6' };

function renderGlossaryWords()
{
  var q   = glsQuery.toLowerCase();
  var all = glsData.entries || [];
  var flt = glsFilter.words;
  flt.type = fillFacetSelect('gls-filter-type', all.map(function(e){ return e.language; }), flt.type, 'All languages');
  flt.tag  = fillFacetSelect('gls-filter-tag', collectItemTags(all), flt.tag, 'All tags');

  var entries = all
    .filter(function(e){ return itemMatchesFacets(e, flt.type, 'language', flt.tag); })
    .filter(function(e){
      return !q || (e.word||'').toLowerCase().includes(q) ||
             (e.definition||'').toLowerCase().includes(q) ||
             (e.language||'').toLowerCase().includes(q) ||
             (e.tags||[]).some(function(t){ return t.toLowerCase().includes(q); });
    })
    .sort(function(a, b){ return (a.word||'').localeCompare(b.word||''); });

  document.getElementById('gls-words-empty').style.display = entries.length ? 'none' : '';

  var groups = {};
  entries.forEach(function(e){
    var letter = (e.word||'').charAt(0).toUpperCase() || '#';
    (groups[letter] = groups[letter] || []).push(e);
  });

  document.getElementById('gls-words-body').innerHTML = Object.keys(groups).sort().map(function(letter){
    return '<div class="gls-letter-group">' +
      '<div class="gls-letter-divider">' + escHtml(letter) + '</div>' +
      groups[letter].map(function(e){
        var lc  = glsLangColor(e.language);
        return '<div class="gls-entry-card" style="--lang-color:' + lc + '" onclick="openGlossaryEntryModal(\'' + e.id + '\')">' +
          '<div class="gls-entry-left">' +
            '<div class="gls-entry-word">' + escHtml(e.word||'') + '</div>' +
            (e.language ? '<span class="gls-lang-badge" style="background:' + lc + '22;color:' + lc + ';border-color:' + lc + '44">' + escHtml(e.language) + '</span>' : '') +
          '</div>' +
          '<div class="gls-entry-right">' +
            '<div class="gls-entry-def">' + escHtml(e.definition||'') + '</div>' +
            (e.example ? '<div class="gls-entry-ex">"' + escHtml(e.example) + '"</div>' : '') +
            tagChipsHtml(e.tags, 'glsFilterByTag') +
          '</div>' +
          '<button class="gls-card-del" onclick="deleteGlossaryEntry(event,\'' + e.id + '\')">×</button>' +
        '</div>';
      }).join('') +
    '</div>';
  }).join('');
}

function renderGlossaryRoots()
{
  var q   = glsQuery.toLowerCase();
  var all = glsData.roots || [];
  var flt = glsFilter.roots;
  flt.type = fillFacetSelect('gls-filter-type', all.map(function(r){ return r.type; }), flt.type, 'All types');
  flt.tag  = fillFacetSelect('gls-filter-tag', collectItemTags(all), flt.tag, 'All tags');

  var rows = all
    .filter(function(r){ return itemMatchesFacets(r, flt.type, 'type', flt.tag); })
    .filter(function(r){
      return !q || (r.form||'').toLowerCase().includes(q) ||
             (r.meaning||'').toLowerCase().includes(q) ||
             (r.tags||[]).some(function(t){ return t.toLowerCase().includes(q); });
    })
    .sort(function(a, b){ return (a.form||'').localeCompare(b.form||''); });

  document.getElementById('gls-roots-empty').style.display = rows.length ? 'none' : '';
  document.getElementById('gls-roots-body').innerHTML = rows.map(function(r){
    var tc = ROOT_TYPE_COLORS[r.type] || '#9ca3af';
    var exs = (r.examples||[]).map(function(ex){ return '<span class="gls-tag">' + escHtml(ex) + '</span>'; }).join('');
    return '<div class="gls-root-card" onclick="openGlossaryRootModal(\'' + r.id + '\')">' +
      '<div class="gls-root-form">' + escHtml(r.form||'') + '</div>' +
      '<span class="gls-root-type" style="background:' + tc + '22;color:' + tc + ';border-color:' + tc + '55">' + escHtml(r.type||'root') + '</span>' +
      '<div class="gls-root-meaning">' + escHtml(r.meaning||'') + '</div>' +
      (exs ? '<div class="gls-root-examples">' + exs + '</div>' : '') +
      tagChipsHtml(r.tags, 'glsFilterByTag') +
      '<button class="gls-card-del" onclick="deleteGlossaryRoot(event,\'' + r.id + '\')">×</button>' +
    '</div>';
  }).join('');
}

function glsAddCurrent()
{
  if (glsTab === 'words') openGlossaryEntryModal(null);
  else openGlossaryRootModal(null);
}

function openGlossaryEntryModal(id)
{
  var e = id ? (glsData.entries||[]).find(function(x){ return x.id === id; }) : null;
  if (e) noteItemAnchor('glossary-entry', e.id, e.word);
  var f =
    '<div class="dem-grid">' +
      '<label class="field-label">Word<input class="modal-input" id="gef-word" value="' + escAttr((e||{}).word||'') + '" placeholder="Word or phrase"></label>' +
      '<label class="field-label">Language<input class="modal-input" id="gef-lang" value="' + escAttr((e||{}).language||'') + '" placeholder="Language or dialect"></label>' +
    '</div>' +
    '<label class="field-label" style="margin-top:8px">Definition<textarea class="modal-input" id="gef-def" rows="2" placeholder="Meaning…">' + escHtml((e||{}).definition||'') + '</textarea></label>' +
    '<label class="field-label" style="margin-top:8px">Example<textarea class="modal-input" id="gef-ex" rows="2" placeholder="Usage example…">' + escHtml((e||{}).example||'') + '</textarea></label>' +
    tagsField('gef-tags', (e||{}).tags);

  openDataModal(id ? 'Edit entry' : 'Add entry', f, function()
  {
    var entry = {
      id:         id || genId(),
      word:       document.getElementById('gef-word').value.trim(),
      language:   document.getElementById('gef-lang').value.trim(),
      definition: document.getElementById('gef-def').value.trim(),
      example:    document.getElementById('gef-ex').value.trim(),
      tags:       parseTagsInput(document.getElementById('gef-tags').value)
    };
    if (!entry.word) return;
    if (id)
      glsData.entries = glsData.entries.map(function(x){ return x.id === id ? entry : x; });
    else
      glsData.entries.push(entry);
    closeDataModal();
    saveGlossaryData();
    renderGlossaryWords();
  });
}

function openGlossaryRootModal(id)
{
  var r = id ? (glsData.roots||[]).find(function(x){ return x.id === id; }) : null;
  if (r) noteItemAnchor('glossary-root', r.id, r.form);
  var f =
    '<div class="dem-grid">' +
      '<label class="field-label">Form<input class="modal-input" id="grf-form" value="' + escAttr((r||{}).form||'') + '" placeholder="e.g. -ael, kron-"></label>' +
      '<label class="field-label">Type<select class="modal-input tb-select" id="grf-type">' +
        '<option value="prefix"' + ((r||{}).type==='prefix'?' selected':'') + '>Prefix</option>' +
        '<option value="suffix"' + ((r||{}).type==='suffix'?' selected':'') + '>Suffix</option>' +
        '<option value="root"'  + ((r||{}).type==='root'  ?' selected':'') + '>Root</option>' +
        '<option value="infix"' + ((r||{}).type==='infix' ?' selected':'') + '>Infix</option>' +
      '</select></label>' +
    '</div>' +
    '<label class="field-label" style="margin-top:8px">Meaning<input class="modal-input" id="grf-mean" value="' + escAttr((r||{}).meaning||'') + '" placeholder="What it means…"></label>' +
    '<label class="field-label" style="margin-top:8px">Examples <span style="color:var(--text3);font-weight:400">(comma separated)</span><input class="modal-input" id="grf-ex" value="' + escAttr(((r||{}).examples||[]).join(', ')) + '"></label>' +
    tagsField('grf-tags', (r||{}).tags);

  openDataModal(id ? 'Edit root / affix' : 'Add root / affix', f, function()
  {
    var root = {
      id:       id || genId(),
      form:     document.getElementById('grf-form').value.trim(),
      type:     document.getElementById('grf-type').value,
      meaning:  document.getElementById('grf-mean').value.trim(),
      examples: document.getElementById('grf-ex').value.split(',').map(function(t){ return t.trim(); }).filter(Boolean),
      tags:     parseTagsInput(document.getElementById('grf-tags').value)
    };
    if (!root.form) return;
    if (id)
      glsData.roots = glsData.roots.map(function(x){ return x.id === id ? root : x; });
    else
      glsData.roots.push(root);
    closeDataModal();
    saveGlossaryData();
    renderGlossaryRoots();
  });
}

function deleteGlossaryEntry(e, id)
{
  e.stopPropagation();
  glsData.entries = (glsData.entries||[]).filter(function(x){ return x.id !== id; });
  saveGlossaryData(); renderGlossaryWords();
}

function deleteGlossaryRoot(e, id)
{
  e.stopPropagation();
  glsData.roots = (glsData.roots||[]).filter(function(x){ return x.id !== id; });
  saveGlossaryData(); renderGlossaryRoots();
}

// ── BESTIARY ──

var bstData  = null;
var bstQuery = '';
var bstFilter = { type: '', tag: '' };

var BST_DANGER_LABELS = ['Harmless', 'Low', 'Moderate', 'High', 'Deadly'];
var BST_DANGER_COLORS = ['#2ecc71', '#a3c94a', '#f1c40f', '#e67e22', '#e74c3c'];

function loadBestiaryFile(file)
{
  try { bstData = JSON.parse(file.content || '{}'); }
  catch(e) { bstData = {}; }
  if (!bstData.beasts) bstData.beasts = [];
  document.getElementById('bestiary-title-input').value = file.name || bstData.name || '';
  var ft = document.getElementById('bst-file-tags');
  if (ft) ft.value = (bstData.tags || []).join(', ');
  bstQuery  = '';
  bstFilter = { type: '', tag: '' };
  document.getElementById('bst-search').value = '';
  renderBestiary();
}

function saveBestiaryData()
{
  if (!currentFileId || !files[currentFileId] || files[currentFileId].type !== 'bestiary') return;
  bstData.name = document.getElementById('bestiary-title-input').value.trim() || bstData.name || 'Bestiary';
  var ft = document.getElementById('bst-file-tags');
  if (ft) bstData.tags = parseTagsInput(ft.value);
  files[currentFileId].name     = bstData.name;
  files[currentFileId].content  = JSON.stringify(bstData, null, 2);
  files[currentFileId].fileTags = bstData.tags || [];
  files[currentFileId].modified = Date.now();
  scheduleSave();
  renderFileList();
}

function onBestiaryTitleChange() { saveBestiaryData(); }
function onBestiaryFileTagsChange() { saveBestiaryData(); }

function onBestiarySearch(q) { bstQuery = q; renderBestiary(); }
function onBstFilterType(v) { bstFilter.type = v; renderBestiary(); }
function onBstFilterTag(v)  { bstFilter.tag  = v; renderBestiary(); }
function bstFilterByTag(t)  { bstFilter.tag  = t; renderBestiary(); }

function bstDangerBadge(level)
{
  var idx = Math.min(Math.max((level|0) - 1, 0), 4);
  var color = BST_DANGER_COLORS[idx];
  return '<span class="bst-danger" style="background:' + color + '22;color:' + color + ';border-color:' + color + '55" title="Danger: ' + BST_DANGER_LABELS[idx] + '">' +
    '&#9679;'.repeat(idx + 1) + '<span class="bst-danger-label">' + BST_DANGER_LABELS[idx] + '</span></span>';
}

function renderBestiary()
{
  var q   = bstQuery.toLowerCase();
  var all = bstData.beasts || [];
  bstFilter.type = fillFacetSelect('bst-filter-type', all.map(function(b){ return b.category; }), bstFilter.type, 'All categories');
  bstFilter.tag  = fillFacetSelect('bst-filter-tag', collectItemTags(all), bstFilter.tag, 'All tags');

  var beasts = all
    .filter(function(b){ return itemMatchesFacets(b, bstFilter.type, 'category', bstFilter.tag); })
    .filter(function(b){
      return !q || (b.name||'').toLowerCase().includes(q) ||
             (b.commonName||'').toLowerCase().includes(q) ||
             (b.category||'').toLowerCase().includes(q) ||
             (b.habitat||'').toLowerCase().includes(q) ||
             (b.description||'').toLowerCase().includes(q) ||
             (b.tags||[]).some(function(t){ return t.toLowerCase().includes(q); }) ||
             (b.abilities||[]).some(function(a){ return a.toLowerCase().includes(q); });
    })
    .sort(function(a, b){ return (a.commonName||a.name||'').localeCompare(b.commonName||b.name||''); });

  document.getElementById('bst-empty').style.display = beasts.length ? 'none' : '';

  // Group by category so related creatures sit together.
  var groups = {};
  beasts.forEach(function(b){
    var cat = (b.category||'').trim() || 'Uncategorised';
    (groups[cat] = groups[cat] || []).push(b);
  });

  document.getElementById('bst-body').innerHTML = Object.keys(groups).sort().map(function(cat){
    var cc = glsLangColor(cat);
    return '<div class="gls-letter-group">' +
      '<div class="gls-letter-divider" style="color:' + cc + '">' + escHtml(cat) + '</div>' +
      groups[cat].map(function(b){
        var abilities = (b.abilities||[]).map(function(a){ return '<span class="gls-tag">' + escHtml(a) + '</span>'; }).join('');
        return '<div class="bst-card" onclick="openBeastModal(\'' + b.id + '\')">' +
          '<div class="bst-card-head">' +
            '<div class="bst-name">' + escHtml(b.commonName || b.name || '') +
              (b.commonName && b.name ? ' <span class="bst-sci-name">' + escHtml(b.name) + '</span>' : '') +
            '</div>' +
            bstDangerBadge(b.danger || 1) +
            (b.habitat ? '<span class="bst-habitat" title="Habitat">' + escHtml(b.habitat) + '</span>' : '') +
          '</div>' +
          (b.description ? '<div class="bst-desc">' + escHtml(b.description) + '</div>' : '') +
          (abilities ? '<div class="bst-abilities">' + abilities + '</div>' : '') +
          tagChipsHtml(b.tags, 'bstFilterByTag') +
          '<button class="gls-card-del" onclick="deleteBeast(event,\'' + b.id + '\')">×</button>' +
        '</div>';
      }).join('') +
    '</div>';
  }).join('');
}

function openBeastModal(id)
{
  var b = id ? (bstData.beasts||[]).find(function(x){ return x.id === id; }) : null;
  if (b) noteItemAnchor('beast', b.id, b.commonName || b.name);
  var danger = (b||{}).danger || 1;
  var f =
    '<div class="dem-grid">' +
      '<label class="field-label">Common name<input class="modal-input" id="bef-common" value="' + escAttr((b||{}).commonName||'') + '" placeholder="Everyday name…"></label>' +
      '<label class="field-label">Scientific name<input class="modal-input" id="bef-name" value="' + escAttr((b||{}).name||'') + '" placeholder="Formal / scholarly name…"></label>' +
    '</div>' +
    '<div class="dem-grid" style="margin-top:8px">' +
      '<label class="field-label">Category<input class="modal-input" id="bef-cat" value="' + escAttr((b||{}).category||'') + '" placeholder="e.g. Dragon, Undead, Fey…"></label>' +
      '<label class="field-label">Danger level<select class="modal-input tb-select" id="bef-danger">' +
        BST_DANGER_LABELS.map(function(lbl, i){
          return '<option value="' + (i + 1) + '"' + (danger === i + 1 ? ' selected' : '') + '>' + (i + 1) + ' — ' + lbl + '</option>';
        }).join('') +
      '</select></label>' +
    '</div>' +
    '<label class="field-label" style="margin-top:8px">Habitat<input class="modal-input" id="bef-habitat" value="' + escAttr((b||{}).habitat||'') + '" placeholder="Where it lives…"></label>' +
    '<label class="field-label" style="margin-top:8px">Description<textarea class="modal-input" id="bef-desc" rows="4" placeholder="Appearance, behaviour, lore…">' + escHtml((b||{}).description||'') + '</textarea></label>' +
    '<label class="field-label" style="margin-top:8px">Abilities <span style="color:var(--text3);font-weight:400">(comma separated)</span><input class="modal-input" id="bef-abilities" value="' + escAttr(((b||{}).abilities||[]).join(', ')) + '" placeholder="e.g. Fire breath, Flight…"></label>' +
    tagsField('bef-tags', (b||{}).tags);

  openDataModal(id ? 'Edit beast' : 'Add beast', f, function()
  {
    var beast = {
      id:          id || genId(),
      name:        document.getElementById('bef-name').value.trim(),
      commonName:  document.getElementById('bef-common').value.trim(),
      category:    document.getElementById('bef-cat').value.trim(),
      danger:      parseInt(document.getElementById('bef-danger').value, 10) || 1,
      habitat:     document.getElementById('bef-habitat').value.trim(),
      description: document.getElementById('bef-desc').value.trim(),
      abilities:   document.getElementById('bef-abilities').value.split(',').map(function(t){ return t.trim(); }).filter(Boolean),
      tags:        parseTagsInput(document.getElementById('bef-tags').value)
    };
    if (!beast.name && !beast.commonName) return;
    if (id)
      bstData.beasts = bstData.beasts.map(function(x){ return x.id === id ? beast : x; });
    else
      bstData.beasts.push(beast);
    closeDataModal();
    saveBestiaryData();
    renderBestiary();
  });
}

function deleteBeast(e, id)
{
  e.stopPropagation();
  bstData.beasts = (bstData.beasts||[]).filter(function(x){ return x.id !== id; });
  saveBestiaryData(); renderBestiary();
}

// ── CALENDAR ──

var calData = null;
var calTab  = 'months';
var calFilter = { months: { type: '', tag: '' }, holidays: { type: '', tag: '' } };

function loadCalendarFile(file)
{
  calFilter = { months: { type: '', tag: '' }, holidays: { type: '', tag: '' } };
  try { calData = JSON.parse(file.content || '{}'); }
  catch(e) { calData = {}; }
  if (!calData.months)   calData.months   = [];
  if (!calData.seasons)  calData.seasons  = [];
  if (!calData.holidays) calData.holidays = [];
  if (!calData.epochRealDate)    calData.epochRealDate    = new Date().toISOString().slice(0,10);
  if (!calData.epochFictionYear) calData.epochFictionYear = 1;
  if (!calData.daysPerYear)      calData.daysPerYear      = 365;

  document.getElementById('calendar-title-input').value = file.name || calData.name || '';
  document.getElementById('cal-days-per-year').value    = calData.daysPerYear;
  document.getElementById('cal-epoch-real').value       = calData.epochRealDate;
  var ft = document.getElementById('cal-file-tags');
  if (ft) ft.value = (calData.tags || []).join(', ');
  switchCalendarTab('months');
  renderCalendar();
}

function saveCalendarData()
{
  if (!currentFileId || !files[currentFileId] || files[currentFileId].type !== 'calendar') return;
  calData.name = document.getElementById('calendar-title-input').value.trim() || calData.name || 'Calendar';
  var ft = document.getElementById('cal-file-tags');
  if (ft) calData.tags = parseTagsInput(ft.value);
  files[currentFileId].name     = calData.name;
  files[currentFileId].content  = JSON.stringify(calData, null, 2);
  files[currentFileId].fileTags = calData.tags || [];
  files[currentFileId].modified = Date.now();
  scheduleSave();
  renderFileList();
}

function onCalendarTitleChange() { saveCalendarData(); }
function onCalendarFileTagsChange() { saveCalendarData(); }
function onCalFilterType(v) { calFilter[calTab === 'holidays' ? 'holidays' : 'months'].type = v; renderCalendar(); }
function onCalFilterTag(v)  { calFilter[calTab === 'holidays' ? 'holidays' : 'months'].tag  = v; renderCalendar(); }
function calFilterByTag(t)  { calFilter[calTab === 'holidays' ? 'holidays' : 'months'].tag  = t; renderCalendar(); }
function onCalDaysChange(v)  { calData.daysPerYear = parseInt(v,10)||365; saveCalendarData(); }
function onCalEpochChange()  { calData.epochRealDate = document.getElementById('cal-epoch-real').value; saveCalendarData(); }

function switchCalendarTab(tab)
{
  calTab = tab;
  ['months','holidays','convert'].forEach(function(t){
    document.getElementById('cal-' + t + '-panel').style.display = (t === tab) ? '' : 'none';
    document.getElementById('cal-tab-' + t).classList.toggle('active', t === tab);
  });
  var addBtn = document.getElementById('cal-add-btn');
  if (addBtn) addBtn.style.display = (tab === 'convert') ? 'none' : '';
  // The date converter has nothing to filter.
  ['cal-filter-type', 'cal-filter-tag'].forEach(function(fid){
    var el = document.getElementById(fid);
    if (el) el.style.display = (tab === 'convert') ? 'none' : '';
  });
  renderCalendar();
}

function calAddCurrent()
{
  if (calTab === 'months')   { openCalendarAddModal(); }
  else if (calTab === 'holidays') { openCalendarHolidayModal(null); }
}

function renderCalendar()
{
  if (calTab === 'months')        renderCalendarMonths();
  else if (calTab === 'holidays') renderCalendarHolidays();
  else if (calTab === 'convert')  renderCalendarConverter();
}

function renderCalendarMonths()
{
  var seasons  = calData.seasons  || [];
  var months   = calData.months   || [];
  var hasData  = seasons.length + months.length > 0;
  document.getElementById('cal-months-empty').style.display = hasData ? 'none' : '';

  var flt = calFilter.months;
  flt.type = fillFacetSelect('cal-filter-type', seasons.map(function(s){ return s.name; }), flt.type, 'All seasons');
  flt.tag  = fillFacetSelect('cal-filter-tag', collectItemTags(months), flt.tag, 'All tags');

  // Year timeline
  var totalDays = months.reduce(function(sum, m){ return sum + (m.days||0); }, 0);
  var timeline  = document.getElementById('cal-year-timeline');
  if (timeline)
  {
    if (months.length && totalDays > 0)
    {
      timeline.style.display = '';
      timeline.innerHTML = '<div class="cal-timeline-label">Year overview</div>' +
        '<div class="cal-timeline-bar">' +
        months.map(function(m){
          var season = seasons.find(function(s){ return (s.monthIds||[]).includes(m.id); });
          var color  = season ? season.color : '#555';
          var pct    = ((m.days||0) / totalDays * 100).toFixed(2);
          var hols   = (calData.holidays||[]).filter(function(h){ return h.monthId === m.id; });
          var dots   = hols.map(function(h){ return '<span class="cal-tl-hol" title="' + escAttr(h.name||'') + '" style="left:' + ((h.day||1)/(m.days||30)*100).toFixed(1) + '%"></span>'; }).join('');
          return '<div class="cal-tl-seg" style="width:' + pct + '%;background:' + escAttr(color) + '" title="' + escAttr(m.name||'') + ' · ' + (m.days||0) + ' days" onclick="openCalendarMonthModal(\'' + m.id + '\')">' +
            '<span class="cal-tl-name">' + escHtml(m.name||'') + '</span>' +
            dots +
          '</div>';
        }).join('') +
        '</div>';
    }
    else timeline.style.display = 'none';
  }

  // Season strip
  var strip = document.getElementById('cal-seasons-strip');
  strip.innerHTML = seasons.map(function(s){
    var days = calSeasonDays(s);
    var mcount = (s.monthIds||[]).length;
    return '<div class="cal-season-chip" style="--season-color:' + escAttr(s.color||'#888') + '" onclick="openCalendarSeasonModal(\'' + s.id + '\')">' +
      '<span class="cal-season-dot" style="background:' + escAttr(s.color||'#888') + '"></span>' +
      '<span class="cal-season-name">' + escHtml(s.name||'') + '</span>' +
      '<span class="cal-season-meta">' + mcount + (mcount===1?' month':' months') + (days ? ' · ' + days + ' days' : '') + '</span>' +
      '<button class="gls-card-del" style="margin-left:4px" onclick="deleteCalendarSeason(event,\'' + s.id + '\')">×</button>' +
    '</div>';
  }).join('');

  // Month cards. The year timeline and season strip above stay unfiltered
  // (they show the whole year by design); the facets narrow the cards only.
  var shownMonths = months.filter(function(m){
    var season = seasons.find(function(s){ return (s.monthIds||[]).includes(m.id); });
    if (flt.type && (!season || season.name !== flt.type)) return false;
    if (flt.tag && (m.tags||[]).indexOf(flt.tag) === -1) return false;
    return true;
  });

  var grid = document.getElementById('cal-months-grid');
  grid.innerHTML = shownMonths.map(function(m){
    var season = seasons.find(function(s){ return (s.monthIds||[]).includes(m.id); });
    var sColor = season ? season.color : '#555';
    var sName  = season ? season.name  : '';
    var hols   = (calData.holidays||[]).filter(function(h){ return h.monthId === m.id; });
    var days   = m.days || 0;
    // mini day-dot grid (max 42 dots)
    var dotHtml = '';
    if (days > 0 && days <= 42)
    {
      var holDays = {};
      hols.forEach(function(h){ if (h.day) holDays[h.day] = h; });
      var dotsArr = [];
      for (var d = 1; d <= days; d++)
        dotsArr.push('<span class="cal-day-dot' + (holDays[d] ? ' cal-day-hol-dot' : '') + '"' +
          (holDays[d] ? ' title="' + escAttr(holDays[d].name||'') + '"' : '') +
          ' style="' + (holDays[d] ? 'background:' + escAttr(sColor) + '' : '') + '"></span>');
      dotHtml = '<div class="cal-day-dots">' + dotsArr.join('') + '</div>';
    }
    var holHtml = hols.map(function(h){
      return '<div class="cal-month-hol"><span class="cal-hol-dot" style="background:' + escAttr(sColor) + '"></span>' + escHtml(h.name||'') + ' <span class="cal-hol-day">day ' + (h.day||'?') + '</span></div>';
    }).join('');
    return '<div class="cal-month-card" style="--season-color:' + escAttr(sColor) + '" onclick="openCalendarMonthModal(\'' + m.id + '\')">' +
      '<div class="cal-month-header">' +
        '<span class="cal-month-name">' + escHtml(m.name||'') + '</span>' +
        '<span class="cal-month-days-badge">' + (m.days||'?') + '</span>' +
      '</div>' +
      (sName ? '<div class="cal-month-season-label">' + escHtml(sName) + '</div>' : '') +
      (m.description ? '<div class="cal-month-desc">' + escHtml(m.description) + '</div>' : '') +
      dotHtml +
      tagChipsHtml(m.tags, 'calFilterByTag') +
      (holHtml ? '<div class="cal-month-hols">' + holHtml + '</div>' : '') +
      '<button class="gls-card-del cal-card-del" onclick="deleteCalendarMonth(event,\'' + m.id + '\')">×</button>' +
    '</div>';
  }).join('');
}

function calSeasonDays(season)
{
  var total = 0;
  (season.monthIds||[]).forEach(function(mid){
    var m = (calData.months||[]).find(function(x){ return x.id === mid; });
    if (m) total += m.days || 0;
  });
  return total;
}

var HOL_TYPE_ICONS = { festival:'🎉', observance:'🔔', memorial:'🕯', harvest:'🌾', solstice:'☀', other:'◆' };
var HOL_TYPE_COLORS = { festival:'#e67e22', observance:'#3498db', memorial:'#9b59b6', harvest:'#2ecc71', solstice:'#f1c40f', other:'#9ca3af' };

function renderCalendarHolidays()
{
  var all = calData.holidays || [];
  var flt = calFilter.holidays;
  flt.type = fillFacetSelect('cal-filter-type', all.map(function(h){ return h.type; }), flt.type, 'All types');
  flt.tag  = fillFacetSelect('cal-filter-tag', collectItemTags(all), flt.tag, 'All tags');

  var holidays = all.filter(function(h){ return itemMatchesFacets(h, flt.type, 'type', flt.tag); });
  document.getElementById('cal-holidays-empty').style.display = holidays.length ? 'none' : '';
  var sorted = holidays.slice().sort(function(a, b){
    var ai = (calData.months||[]).findIndex(function(m){ return m.id === a.monthId; });
    var bi = (calData.months||[]).findIndex(function(m){ return m.id === b.monthId; });
    return ai - bi || (a.day||0) - (b.day||0);
  });
  document.getElementById('cal-holidays-body').innerHTML = sorted.map(function(h){
    var month  = (calData.months||[]).find(function(m){ return m.id === h.monthId; });
    var season = month ? (calData.seasons||[]).find(function(s){ return (s.monthIds||[]).includes(month.id); }) : null;
    var sc     = season ? season.color : '#9ca3af';
    var tc     = HOL_TYPE_COLORS[h.type] || '#9ca3af';
    return '<div class="cal-hol-card" onclick="openCalendarHolidayModal(\'' + h.id + '\')">' +
      '<div class="cal-hol-marker" style="background:' + escAttr(sc) + '"></div>' +
      '<div class="cal-hol-body">' +
        '<div class="cal-hol-title">' +
          '<span class="cal-hol-name">' + escHtml(h.name||'') + '</span>' +
          '<span class="cal-hol-type-badge" style="background:' + tc + '22;color:' + tc + ';border-color:' + tc + '55">' + escHtml(h.type||'festival') + '</span>' +
        '</div>' +
        '<div class="cal-hol-when">' +
          (month ? escHtml(month.name) + ', day ' + (h.day||'?') : 'day ' + (h.day||'?')) +
        '</div>' +
        (h.description ? '<div class="cal-hol-desc">' + escHtml(h.description) + '</div>' : '') +
        tagChipsHtml(h.tags, 'calFilterByTag') +
      '</div>' +
      '<button class="gls-card-del" onclick="deleteCalendarHoliday(event,\'' + h.id + '\')">×</button>' +
    '</div>';
  }).join('');
}

function renderCalendarConverter()
{
  var months = calData.months || [];
  var sel = document.getElementById('cal-conv-fm');
  if (sel) sel.innerHTML = months.map(function(m){ return '<option value="' + escAttr(m.id) + '">' + escHtml(m.name||'') + ' (' + (m.days||'?') + ' days)</option>'; }).join('');
}

function convertFictionToReal()
{
  var fyear = parseInt(document.getElementById('cal-conv-fy').value, 10);
  var fmid  = document.getElementById('cal-conv-fm').value;
  var fday  = parseInt(document.getElementById('cal-conv-fd').value, 10) || 1;
  var res   = document.getElementById('cal-conv-f2r');

  if (!calData.epochRealDate || !fyear) { res.textContent = 'Set epoch and enter a year.'; return; }

  var months = calData.months || [];
  var mIdx   = months.findIndex(function(m){ return m.id === fmid; });
  if (mIdx < 0 && months.length) mIdx = 0;

  var dayOfYear = 0;
  for (var i = 0; i < mIdx; i++) dayOfYear += months[i].days || 0;
  dayOfYear += fday - 1;

  var totalDays = (fyear - (calData.epochFictionYear || 1)) * (calData.daysPerYear || 365) + dayOfYear;
  var real = new Date(calData.epochRealDate);
  real.setDate(real.getDate() + totalDays);
  res.textContent = real.toDateString();
}

function convertRealToFiction()
{
  var rd  = document.getElementById('cal-conv-rd').value;
  var res = document.getElementById('cal-conv-r2f');
  if (!rd || !calData.epochRealDate) { res.textContent = 'Set epoch and enter a real date.'; return; }

  var diff = Math.round((new Date(rd) - new Date(calData.epochRealDate)) / 86400000);
  var dpy  = calData.daysPerYear || 365;
  var fyear = (calData.epochFictionYear || 1) + Math.floor(diff / dpy);
  var rem   = ((diff % dpy) + dpy) % dpy;

  var months = calData.months || [];
  var mName  = '(no months defined)';
  var day    = rem + 1;
  for (var i = 0; i < months.length; i++)
  {
    if (rem < (months[i].days||0)) { mName = months[i].name; day = rem + 1; break; }
    rem -= months[i].days || 0;
  }

  res.textContent = 'Year ' + fyear + ', ' + mName + ' day ' + day;
}

function openCalendarAddModal()
{
  var seasonOpts = (calData.seasons||[]).map(function(s){
    return '<option value="' + escAttr(s.id) + '">' + escHtml(s.name||'Season') + '</option>';
  }).join('');

  var f =
    '<p style="font-size:12px;color:var(--text2);margin-bottom:10px">Add a month or a season</p>' +
    '<div style="display:flex;gap:10px;margin-bottom:12px">' +
      '<button class="btn-create" style="flex:1" onclick="closeDataModal();openCalendarMonthModal(null)">+ Month</button>' +
      '<button class="btn-create" style="flex:1" onclick="closeDataModal();openCalendarSeasonModal(null)">+ Season</button>' +
    '</div>';

  openDataModal('Add to calendar', f, null);
  document.getElementById('dem-save-btn').style.display = 'none';
  document.getElementById('data-edit-modal').querySelector('.modal-actions .btn-cancel').textContent = 'Close';
}

function openCalendarMonthModal(id)
{
  document.getElementById('dem-save-btn').style.display = '';
  document.getElementById('data-edit-modal').querySelector('.modal-actions .btn-cancel').textContent = 'Cancel';
  var m = id ? (calData.months||[]).find(function(x){ return x.id === id; }) : null;
  var seasonOpts = (calData.seasons||[]).map(function(s){
    var selMonth = m && (s.monthIds||[]).includes(m.id);
    return '<option value="' + escAttr(s.id) + '"' + (selMonth?' selected':'') + '>' + escHtml(s.name||'') + '</option>';
  }).join('');

  var f =
    '<div class="dem-grid">' +
      '<label class="field-label">Month name<input class="modal-input" id="cmf-name" value="' + escAttr((m||{}).name||'') + '" placeholder="e.g. Frostmoon"></label>' +
      '<label class="field-label">Days<input class="modal-input" id="cmf-days" type="number" min="1" value="' + ((m||{}).days||30) + '"></label>' +
    '</div>' +
    '<label class="field-label" style="margin-top:8px">Season<select class="modal-input tb-select" id="cmf-season"><option value="">— None —</option>' + seasonOpts + '</select></label>' +
    '<label class="field-label" style="margin-top:8px">Description<input class="modal-input" id="cmf-desc" value="' + escAttr((m||{}).description||'') + '" placeholder="Brief description…"></label>' +
    tagsField('cmf-tags', (m||{}).tags);

  openDataModal(id ? 'Edit month' : 'Add month', f, function()
  {
    var name     = document.getElementById('cmf-name').value.trim();
    var days     = parseInt(document.getElementById('cmf-days').value,10) || 30;
    var seasonId = document.getElementById('cmf-season').value;
    var desc     = document.getElementById('cmf-desc').value.trim();
    if (!name) return;

    var mid = id || genId();
    var month = { id: mid, name: name, days: days, description: desc, tags: parseTagsInput(document.getElementById('cmf-tags').value) };

    if (id)
      calData.months = calData.months.map(function(x){ return x.id === id ? month : x; });
    else
      calData.months.push(month);

    // Update season membership
    calData.seasons.forEach(function(s){
      s.monthIds = (s.monthIds||[]).filter(function(x){ return x !== mid; });
      if (seasonId && s.id === seasonId) s.monthIds.push(mid);
    });

    closeDataModal();
    saveCalendarData();
    renderCalendarMonths();
  });
}

function openCalendarSeasonModal(id)
{
  document.getElementById('dem-save-btn').style.display = '';
  document.getElementById('data-edit-modal').querySelector('.modal-actions .btn-cancel').textContent = 'Cancel';
  var s = id ? (calData.seasons||[]).find(function(x){ return x.id === id; }) : null;
  var monthOpts = (calData.months||[]).map(function(m){
    var checked = s && (s.monthIds||[]).includes(m.id) ? ' checked' : '';
    return '<label style="display:flex;align-items:center;gap:6px;font-size:12px;cursor:pointer">' +
      '<input type="checkbox" value="' + escAttr(m.id) + '"' + checked + '> ' + escHtml(m.name||'') +
    '</label>';
  }).join('');

  var f =
    '<div class="dem-grid">' +
      '<label class="field-label">Season name<input class="modal-input" id="csf-name" value="' + escAttr((s||{}).name||'') + '" placeholder="e.g. Winter"></label>' +
      '<label class="field-label">Color<input class="modal-input" id="csf-color" type="color" value="' + escAttr((s||{}).color||'#4a9fc8') + '"></label>' +
    '</div>' +
    (monthOpts ? '<div class="field-label" style="margin-top:8px">Months in this season<div style="display:flex;flex-wrap:wrap;gap:8px;margin-top:6px">' + monthOpts + '</div></div>' : '');

  openDataModal(id ? 'Edit season' : 'Add season', f, function()
  {
    var name  = document.getElementById('csf-name').value.trim();
    var color = document.getElementById('csf-color').value;
    if (!name) return;

    var monthIds = [];
    document.querySelectorAll('#dem-fields input[type=checkbox]:checked').forEach(function(cb){ monthIds.push(cb.value); });

    var season = { id: id || genId(), name: name, color: color, monthIds: monthIds };
    if (id)
      calData.seasons = calData.seasons.map(function(x){ return x.id === id ? season : x; });
    else
      calData.seasons.push(season);

    closeDataModal();
    saveCalendarData();
    renderCalendarMonths();
  });
}

function openCalendarHolidayModal(id)
{
  document.getElementById('dem-save-btn').style.display = '';
  document.getElementById('data-edit-modal').querySelector('.modal-actions .btn-cancel').textContent = 'Cancel';
  var h = id ? (calData.holidays||[]).find(function(x){ return x.id === id; }) : null;
  var monthOpts = (calData.months||[]).map(function(m){
    return '<option value="' + escAttr(m.id) + '"' + ((h&&h.monthId===m.id)?' selected':'') + '>' + escHtml(m.name||'') + '</option>';
  }).join('');

  var typeOpts = ['festival','observance','memorial','harvest','solstice','other'].map(function(t){
    return '<option value="' + t + '"' + ((h&&h.type===t)?' selected':'') + '>' + t.charAt(0).toUpperCase() + t.slice(1) + '</option>';
  }).join('');

  var f =
    '<label class="field-label">Name<input class="modal-input" id="chf-name" value="' + escAttr((h||{}).name||'') + '" placeholder="Holiday name…"></label>' +
    '<div class="dem-grid" style="margin-top:8px">' +
      '<label class="field-label">Month<select class="modal-input tb-select" id="chf-month"><option value="">— select —</option>' + monthOpts + '</select></label>' +
      '<label class="field-label">Day<input class="modal-input" id="chf-day" type="number" min="1" value="' + ((h||{}).day||1) + '"></label>' +
      '<label class="field-label">Type<select class="modal-input tb-select" id="chf-type">' + typeOpts + '</select></label>' +
    '</div>' +
    '<label class="field-label" style="margin-top:8px">Description<textarea class="modal-input" id="chf-desc" rows="2">' + escHtml((h||{}).description||'') + '</textarea></label>' +
    tagsField('chf-tags', (h||{}).tags);

  openDataModal(id ? 'Edit holiday' : 'Add holiday', f, function()
  {
    var name = document.getElementById('chf-name').value.trim();
    if (!name) return;
    var holiday = {
      id:          id || genId(),
      name:        name,
      monthId:     document.getElementById('chf-month').value,
      day:         parseInt(document.getElementById('chf-day').value, 10) || 1,
      type:        document.getElementById('chf-type').value,
      description: document.getElementById('chf-desc').value.trim(),
      tags:        parseTagsInput(document.getElementById('chf-tags').value)
    };
    if (id)
      calData.holidays = calData.holidays.map(function(x){ return x.id === id ? holiday : x; });
    else
      calData.holidays.push(holiday);
    closeDataModal();
    saveCalendarData();
    renderCalendarHolidays();
  });
}

function deleteCalendarSeason(e, id)
{
  e.stopPropagation();
  calData.seasons = (calData.seasons||[]).filter(function(x){ return x.id !== id; });
  saveCalendarData(); renderCalendarMonths();
}

function deleteCalendarMonth(e, id)
{
  e.stopPropagation();
  calData.months = (calData.months||[]).filter(function(x){ return x.id !== id; });
  calData.seasons.forEach(function(s){ s.monthIds = (s.monthIds||[]).filter(function(x){ return x !== id; }); });
  saveCalendarData(); renderCalendarMonths();
}

function deleteCalendarHoliday(e, id)
{
  e.stopPropagation();
  calData.holidays = (calData.holidays||[]).filter(function(x){ return x.id !== id; });
  saveCalendarData(); renderCalendarHolidays();
}

// ── ECONOMY ──

var ecoData = null;
var ecoTab  = 'currencies';
var ecoFilter = { currencies: { type: '', tag: '' }, goods: { type: '', tag: '' }, regions: { type: '', tag: '' } };

function loadEconomyFile(file)
{
  try { ecoData = JSON.parse(file.content || '{}'); }
  catch(e) { ecoData = {}; }
  if (!ecoData.currencies)    ecoData.currencies    = [];
  if (!ecoData.exchangeRates) ecoData.exchangeRates = [];
  if (!ecoData.tradeGoods)    ecoData.tradeGoods    = [];
  if (!ecoData.regions)       ecoData.regions       = [];
  document.getElementById('economy-title-input').value = file.name || ecoData.name || '';
  var ft = document.getElementById('eco-file-tags');
  if (ft) ft.value = (ecoData.tags || []).join(', ');
  ecoFilter = { currencies: { type: '', tag: '' }, goods: { type: '', tag: '' }, regions: { type: '', tag: '' } };
  switchEconomyTab('currencies');
  renderEconomy();
}

function saveEconomyData()
{
  if (!currentFileId || !files[currentFileId] || files[currentFileId].type !== 'economy') return;
  ecoData.name = document.getElementById('economy-title-input').value.trim() || ecoData.name || 'Economy';
  var ft = document.getElementById('eco-file-tags');
  if (ft) ecoData.tags = parseTagsInput(ft.value);
  files[currentFileId].name     = ecoData.name;
  files[currentFileId].content  = JSON.stringify(ecoData, null, 2);
  files[currentFileId].fileTags = ecoData.tags || [];
  files[currentFileId].modified = Date.now();
  scheduleSave();
  renderFileList();
}

function onEconomyTitleChange() { saveEconomyData(); }
function onEconomyFileTagsChange() { saveEconomyData(); }
function onEcoFilterType(v) { ecoFilter[ecoTab].type = v; renderEconomy(); }
function onEcoFilterTag(v)  { ecoFilter[ecoTab].tag  = v; renderEconomy(); }
function ecoFilterByTag(t)  { ecoFilter[ecoTab].tag  = t; renderEconomy(); }

function switchEconomyTab(tab)
{
  ecoTab = tab;
  ['currencies','goods','regions'].forEach(function(t){
    document.getElementById('eco-' + t + '-panel').style.display = (t === tab) ? '' : 'none';
    document.getElementById('eco-tab-' + t).classList.toggle('active', t === tab);
  });
  renderEconomy();
}

function ecoAddCurrent()
{
  if (ecoTab === 'currencies') openEcoCurrencyModal(null);
  else if (ecoTab === 'goods') openEcoGoodModal(null);
  else if (ecoTab === 'regions') openEcoRegionModal(null);
}

function renderEconomy()
{
  if (ecoTab === 'currencies')    renderEcoCurrencies();
  else if (ecoTab === 'goods')    renderEcoGoods();
  else if (ecoTab === 'regions')  renderEcoRegions();
}

var ECO_STATUS_COLORS = { Prosperous:'#2ecc71', Stable:'#3498db', Struggling:'#e67e22', Declining:'#e74c3c', Isolated:'#9b59b6', 'War-torn':'#c0392b', Unknown:'#9ca3af' };
var ECO_CURR_PALETTE  = ['#f39c12','#3498db','#2ecc71','#e74c3c','#9b59b6','#1abc9c','#e67e22','#e91e63'];

function ecoCurrColor(idx)
{
  return ECO_CURR_PALETTE[idx % ECO_CURR_PALETTE.length];
}

function renderEcoCurrencies()
{
  var currencies = ecoData.currencies || [];
  var flt = ecoFilter.currencies;
  flt.type = fillFacetSelect('eco-filter-type', currencies.map(function(c){ return c.region; }), flt.type, 'All regions');
  flt.tag  = fillFacetSelect('eco-filter-tag', collectItemTags(currencies), flt.tag, 'All tags');

  document.getElementById('eco-currencies-empty').style.display = currencies.length ? 'none' : '';
  document.getElementById('eco-exchange-calc').style.display    = currencies.length >= 2 ? '' : 'none';

  // Color and "base currency" status follow the position in the full list,
  // so filtering can't change which currency reads as the base.
  document.getElementById('eco-currencies-body').innerHTML = currencies.map(function(c, i){
    return { c: c, i: i };
  }).filter(function(x){
    return itemMatchesFacets(x.c, flt.type, 'region', flt.tag);
  }).map(function(x){
    var c = x.c;
    var color = ecoCurrColor(x.i);
    var isBase = (x.i === 0);
    return '<div class="eco-curr-card" onclick="openEcoCurrencyModal(\'' + c.id + '\')">' +
      '<div class="eco-curr-symbol" style="color:' + color + ';border-color:' + color + '44">' + escHtml(c.symbol||'?') + '</div>' +
      '<div class="eco-curr-name">' + escHtml(c.name||'') + '</div>' +
      (c.region ? '<div class="eco-curr-region">' + escHtml(c.region) + '</div>' : '') +
      '<div class="eco-curr-value">' + (isBase ? 'Base currency' : '×' + (c.baseValue||1)) + '</div>' +
      (c.description ? '<div class="eco-curr-desc">' + escHtml(c.description) + '</div>' : '') +
      tagChipsHtml(c.tags, 'ecoFilterByTag') +
      '<button class="gls-card-del" onclick="deleteEcoCurrency(event,\'' + c.id + '\')">×</button>' +
    '</div>';
  }).join('');

  var opts = currencies.map(function(c){ return '<option value="' + escAttr(c.id) + '">' + escHtml(c.name||'') + ' (' + escHtml(c.symbol||'') + ')</option>'; }).join('');
  var fromSel = document.getElementById('eco-calc-from');
  var toSel   = document.getElementById('eco-calc-to');
  if (fromSel) fromSel.innerHTML = opts;
  if (toSel)   { toSel.innerHTML = opts; if (currencies.length >= 2) toSel.selectedIndex = 1; }
}

var CAT_PALETTE = ['#e74c3c','#e67e22','#f1c40f','#2ecc71','#3498db','#9b59b6','#1abc9c','#e91e63','#00cec9','#fd79a8'];
function ecoCatColor(cat)
{
  if (!cat) return '#9ca3af';
  var h = 0; for (var i = 0; i < cat.length; i++) h = ((h << 5) - h) + cat.charCodeAt(i);
  return CAT_PALETTE[Math.abs(h) % CAT_PALETTE.length];
}

function renderEcoGoods()
{
  var goods = ecoData.tradeGoods || [];
  var flt = ecoFilter.goods;
  flt.type = fillFacetSelect('eco-filter-type', goods.map(function(g){ return g.category; }), flt.type, 'All categories');
  flt.tag  = fillFacetSelect('eco-filter-tag', collectItemTags(goods), flt.tag, 'All tags');

  var shown = goods.filter(function(g){ return itemMatchesFacets(g, flt.type, 'category', flt.tag); });
  document.getElementById('eco-goods-empty').style.display = shown.length ? 'none' : '';
  var sorted = shown.slice().sort(function(a,b){ return (a.name||'').localeCompare(b.name||''); });
  document.getElementById('eco-goods-body').innerHTML = sorted.map(function(g){
    var cur = (ecoData.currencies||[]).find(function(c){ return c.id === g.priceCurrencyId; });
    var price = g.priceAmount != null ? g.priceAmount + (cur ? ' ' + escHtml(cur.symbol) : '') : '';
    var cc = ecoCatColor(g.category);
    return '<tr class="data-row" onclick="openEcoGoodModal(\'' + g.id + '\')">' +
      '<td class="eco-good-name">' + escHtml(g.name||'') + tagChipsHtml(g.tags, 'ecoFilterByTag') + '</td>' +
      '<td>' + (g.category ? '<span class="eco-cat-badge" style="background:' + cc + '22;color:' + cc + ';border-color:' + cc + '44">' + escHtml(g.category) + '</span>' : '') + '</td>' +
      '<td class="eco-price-cell">' + escHtml(price) + '</td>' +
      '<td class="eco-origin-cell">' + escHtml(g.origin||'') + '</td>' +
      '<td class="eco-desc-cell">' + escHtml(g.description||'') + '</td>' +
      '<td><button class="gls-card-del" onclick="deleteEcoGood(event,\'' + g.id + '\')">×</button></td>' +
    '</tr>';
  }).join('');
}

function renderEcoRegions()
{
  var regions = ecoData.regions || [];
  var flt = ecoFilter.regions;
  flt.type = fillFacetSelect('eco-filter-type', regions.map(function(r){ return r.economicStatus; }), flt.type, 'All statuses');
  flt.tag  = fillFacetSelect('eco-filter-tag', collectItemTags(regions), flt.tag, 'All tags');

  var shown = regions.filter(function(r){ return itemMatchesFacets(r, flt.type, 'economicStatus', flt.tag); });
  document.getElementById('eco-regions-empty').style.display = shown.length ? 'none' : '';
  document.getElementById('eco-regions-body').innerHTML = shown.map(function(r){
    var cur  = (ecoData.currencies||[]).find(function(c){ return c.id === r.primaryCurrencyId; });
    var sc   = ECO_STATUS_COLORS[r.economicStatus] || '#9ca3af';
    var exps = (r.exports||[]).map(function(e){ return '<span class="eco-trade-chip eco-export-chip">' + escHtml(e) + '</span>'; }).join('');
    var imps = (r.imports||[]).map(function(e){ return '<span class="eco-trade-chip eco-import-chip">' + escHtml(e) + '</span>'; }).join('');
    return '<div class="eco-region-card" onclick="openEcoRegionModal(\'' + r.id + '\')">' +
      '<div class="eco-region-header">' +
        '<span class="eco-region-name">' + escHtml(r.name||'') + '</span>' +
        (r.economicStatus ? '<span class="eco-status-badge" style="background:' + sc + '22;color:' + sc + ';border-color:' + sc + '44">' + escHtml(r.economicStatus) + '</span>' : '') +
      '</div>' +
      (cur ? '<div class="eco-region-currency">Currency: ' + escHtml(cur.name||'') + (cur.symbol ? ' (' + escHtml(cur.symbol) + ')' : '') + '</div>' : '') +
      (exps ? '<div class="eco-region-trade"><span class="eco-trade-label">Exports</span>' + exps + '</div>' : '') +
      (imps ? '<div class="eco-region-trade"><span class="eco-trade-label">Imports</span>' + imps + '</div>' : '') +
      (r.notes ? '<div class="eco-region-notes">' + escHtml(r.notes) + '</div>' : '') +
      tagChipsHtml(r.tags, 'ecoFilterByTag') +
      '<button class="gls-card-del" onclick="deleteEcoRegion(event,\'' + r.id + '\')">×</button>' +
    '</div>';
  }).join('');
}

function calcEcoExchange()
{
  var amount   = parseFloat(document.getElementById('eco-calc-amount').value) || 1;
  var fromId   = document.getElementById('eco-calc-from').value;
  var toId     = document.getElementById('eco-calc-to').value;
  var result   = document.getElementById('eco-calc-result');
  var currencies = ecoData.currencies || [];
  var from = currencies.find(function(c){ return c.id === fromId; });
  var to   = currencies.find(function(c){ return c.id === toId; });
  if (!from || !to) { result.textContent = 'Select two currencies.'; return; }

  // Check stored exchange rates first (from→to and to→from)
  var rate = null;
  var stored = (ecoData.exchangeRates||[]).find(function(r){ return r.fromId === fromId && r.toId === toId; });
  if (stored) { rate = stored.rate; }
  else {
    var rev = (ecoData.exchangeRates||[]).find(function(r){ return r.fromId === toId && r.toId === fromId; });
    if (rev && rev.rate) rate = 1 / rev.rate;
  }
  // Fall back to base values
  if (rate === null && from.baseValue && to.baseValue) rate = to.baseValue / from.baseValue;

  if (rate === null) { result.textContent = 'No exchange rate defined between these currencies.'; return; }
  var converted = amount * rate;
  result.textContent = amount + ' ' + (from.symbol||from.name) + ' = ' + converted.toFixed(4).replace(/\.?0+$/, '') + ' ' + (to.symbol||to.name);
}

function openEcoCurrencyModal(id)
{
  var c = id ? (ecoData.currencies||[]).find(function(x){ return x.id === id; }) : null;
  if (c) noteItemAnchor('eco-currency', c.id, c.name);
  // Build exchange rate rows for this currency
  var xrHtml = '';
  if (id && ecoData.currencies.length > 1)
  {
    var others = ecoData.currencies.filter(function(x){ return x.id !== id; });
    xrHtml = '<div class="field-label" style="margin-top:12px">Exchange rates (1 ' + escHtml((c||{}).symbol||'?') + ' =)</div>' +
      '<div style="display:flex;flex-direction:column;gap:4px;margin-top:4px">' +
      others.map(function(o){
        var stored = (ecoData.exchangeRates||[]).find(function(r){ return r.fromId === id && r.toId === o.id; });
        var rev    = (ecoData.exchangeRates||[]).find(function(r){ return r.fromId === o.id && r.toId === id; });
        var val    = stored ? stored.rate : (rev && rev.rate ? (1/rev.rate).toFixed(4) : '');
        return '<div style="display:flex;align-items:center;gap:8px;font-size:12px">' +
          '<span style="flex:1;color:var(--text2)">' + escHtml(o.name||'') + ' (' + escHtml(o.symbol||'') + ')</span>' +
          '<input class="modal-input tb-input-sm" data-xr-to="' + escAttr(o.id) + '" value="' + escAttr(String(val)) + '" type="number" step="any" style="width:100px" placeholder="rate">' +
        '</div>';
      }).join('') +
      '</div>';
  }

  var f =
    '<div class="dem-grid">' +
      '<label class="field-label">Name<input class="modal-input" id="ecf-name" value="' + escAttr((c||{}).name||'') + '" placeholder="e.g. Gold Crown"></label>' +
      '<label class="field-label">Symbol<input class="modal-input" id="ecf-sym" value="' + escAttr((c||{}).symbol||'') + '" placeholder="GC"></label>' +
    '</div>' +
    '<div class="dem-grid" style="margin-top:8px">' +
      '<label class="field-label">Region / Faction<input class="modal-input" id="ecf-reg" value="' + escAttr((c||{}).region||'') + '"></label>' +
      '<label class="field-label">Base value<input class="modal-input" id="ecf-base" type="number" step="any" value="' + ((c||{}).baseValue||1) + '" placeholder="relative value"></label>' +
    '</div>' +
    '<label class="field-label" style="margin-top:8px">Notes<input class="modal-input" id="ecf-desc" value="' + escAttr((c||{}).description||'') + '"></label>' +
    tagsField('ecf-tags', (c||{}).tags) +
    xrHtml;

  openDataModal(id ? 'Edit currency' : 'Add currency', f, function()
  {
    var name = document.getElementById('ecf-name').value.trim();
    if (!name) return;
    var cid = id || genId();
    var currency = {
      id:          cid,
      name:        name,
      symbol:      document.getElementById('ecf-sym').value.trim(),
      region:      document.getElementById('ecf-reg').value.trim(),
      baseValue:   parseFloat(document.getElementById('ecf-base').value) || 1,
      description: document.getElementById('ecf-desc').value.trim(),
      tags:        parseTagsInput(document.getElementById('ecf-tags').value)
    };

    if (id)
      ecoData.currencies = ecoData.currencies.map(function(x){ return x.id === id ? currency : x; });
    else
      ecoData.currencies.push(currency);

    // Save exchange rates
    document.querySelectorAll('#dem-fields input[data-xr-to]').forEach(function(inp){
      var toId = inp.dataset.xrTo;
      var val  = parseFloat(inp.value);
      ecoData.exchangeRates = (ecoData.exchangeRates||[]).filter(function(r){ return !(r.fromId === cid && r.toId === toId); });
      if (!isNaN(val) && val > 0) ecoData.exchangeRates.push({ id: genId(), fromId: cid, toId: toId, rate: val, notes: '' });
    });

    closeDataModal();
    saveEconomyData();
    renderEcoCurrencies();
  });
}

function openEcoGoodModal(id)
{
  var g = id ? (ecoData.tradeGoods||[]).find(function(x){ return x.id === id; }) : null;
  if (g) noteItemAnchor('eco-good', g.id, g.name);
  var curOpts = (ecoData.currencies||[]).map(function(c){
    return '<option value="' + escAttr(c.id) + '"' + ((g&&g.priceCurrencyId===c.id)?' selected':'') + '>' + escHtml(c.name||'') + ' (' + escHtml(c.symbol||'') + ')</option>';
  }).join('');

  var f =
    '<div class="dem-grid">' +
      '<label class="field-label">Name<input class="modal-input" id="egf-name" value="' + escAttr((g||{}).name||'') + '" placeholder="e.g. Ironwood"></label>' +
      '<label class="field-label">Category<input class="modal-input" id="egf-cat" value="' + escAttr((g||{}).category||'') + '" placeholder="Material, Food, Luxury…"></label>' +
    '</div>' +
    '<div class="dem-grid" style="margin-top:8px">' +
      '<label class="field-label">Base price<input class="modal-input" id="egf-price" type="number" step="any" value="' + ((g||{}).priceAmount||'') + '" placeholder="Amount"></label>' +
      '<label class="field-label">Currency<select class="modal-input tb-select" id="egf-cur"><option value="">— select —</option>' + curOpts + '</select></label>' +
    '</div>' +
    '<label class="field-label" style="margin-top:8px">Origin region<input class="modal-input" id="egf-origin" value="' + escAttr((g||{}).origin||'') + '"></label>' +
    '<label class="field-label" style="margin-top:8px">Description<textarea class="modal-input" id="egf-desc" rows="2">' + escHtml((g||{}).description||'') + '</textarea></label>' +
    tagsField('egf-tags', (g||{}).tags);

  openDataModal(id ? 'Edit trade good' : 'Add trade good', f, function()
  {
    var name = document.getElementById('egf-name').value.trim();
    if (!name) return;
    var good = {
      id:               id || genId(),
      name:             name,
      category:         document.getElementById('egf-cat').value.trim(),
      priceAmount:      parseFloat(document.getElementById('egf-price').value) || null,
      priceCurrencyId:  document.getElementById('egf-cur').value,
      origin:           document.getElementById('egf-origin').value.trim(),
      description:      document.getElementById('egf-desc').value.trim(),
      tags:             parseTagsInput(document.getElementById('egf-tags').value)
    };
    if (id)
      ecoData.tradeGoods = ecoData.tradeGoods.map(function(x){ return x.id === id ? good : x; });
    else
      ecoData.tradeGoods.push(good);
    closeDataModal();
    saveEconomyData();
    renderEcoGoods();
  });
}

function openEcoRegionModal(id)
{
  var r = id ? (ecoData.regions||[]).find(function(x){ return x.id === id; }) : null;
  if (r) noteItemAnchor('eco-region', r.id, r.name);
  var curOpts = (ecoData.currencies||[]).map(function(c){
    return '<option value="' + escAttr(c.id) + '"' + ((r&&r.primaryCurrencyId===c.id)?' selected':'') + '>' + escHtml(c.name||'') + '</option>';
  }).join('');
  var statusOpts = ['Prosperous','Stable','Struggling','Declining','Isolated','War-torn','Unknown'].map(function(s){
    return '<option value="' + s + '"' + ((r&&r.economicStatus===s)?' selected':'') + '>' + s + '</option>';
  }).join('');

  var f =
    '<div class="dem-grid">' +
      '<label class="field-label">Region / Faction<input class="modal-input" id="erf-name" value="' + escAttr((r||{}).name||'') + '" placeholder="Name…"></label>' +
      '<label class="field-label">Economic status<select class="modal-input tb-select" id="erf-status"><option value="">— select —</option>' + statusOpts + '</select></label>' +
    '</div>' +
    '<label class="field-label" style="margin-top:8px">Primary currency<select class="modal-input tb-select" id="erf-cur"><option value="">— none —</option>' + curOpts + '</select></label>' +
    '<div class="dem-grid" style="margin-top:8px">' +
      '<label class="field-label">Main exports <span style="color:var(--text3);font-weight:400">(comma separated)</span><input class="modal-input" id="erf-exp" value="' + escAttr(((r||{}).exports||[]).join(', ')) + '"></label>' +
      '<label class="field-label">Main imports <span style="color:var(--text3);font-weight:400">(comma separated)</span><input class="modal-input" id="erf-imp" value="' + escAttr(((r||{}).imports||[]).join(', ')) + '"></label>' +
    '</div>' +
    '<label class="field-label" style="margin-top:8px">Notes<textarea class="modal-input" id="erf-notes" rows="2">' + escHtml((r||{}).notes||'') + '</textarea></label>' +
    tagsField('erf-tags', (r||{}).tags);

  openDataModal(id ? 'Edit region' : 'Add region', f, function()
  {
    var name = document.getElementById('erf-name').value.trim();
    if (!name) return;
    var region = {
      id:                id || genId(),
      name:              name,
      economicStatus:    document.getElementById('erf-status').value,
      primaryCurrencyId: document.getElementById('erf-cur').value,
      exports:           document.getElementById('erf-exp').value.split(',').map(function(t){ return t.trim(); }).filter(Boolean),
      imports:           document.getElementById('erf-imp').value.split(',').map(function(t){ return t.trim(); }).filter(Boolean),
      notes:             document.getElementById('erf-notes').value.trim(),
      tags:              parseTagsInput(document.getElementById('erf-tags').value)
    };
    if (id)
      ecoData.regions = ecoData.regions.map(function(x){ return x.id === id ? region : x; });
    else
      ecoData.regions.push(region);
    closeDataModal();
    saveEconomyData();
    renderEcoRegions();
  });
}

function deleteEcoCurrency(e, id)
{
  e.stopPropagation();
  ecoData.currencies    = (ecoData.currencies||[]).filter(function(x){ return x.id !== id; });
  ecoData.exchangeRates = (ecoData.exchangeRates||[]).filter(function(r){ return r.fromId !== id && r.toId !== id; });
  saveEconomyData(); renderEcoCurrencies();
}

function deleteEcoGood(e, id)
{
  e.stopPropagation();
  ecoData.tradeGoods = (ecoData.tradeGoods||[]).filter(function(x){ return x.id !== id; });
  saveEconomyData(); renderEcoGoods();
}

function deleteEcoRegion(e, id)
{
  e.stopPropagation();
  ecoData.regions = (ecoData.regions||[]).filter(function(x){ return x.id !== id; });
  saveEconomyData(); renderEcoRegions();
}

// ── DOC ──
var docTextUndo = makeTextUndo(function(){ return document.getElementById('editor'); });

function loadDocFile(file)
{
  docTextUndo.reset();
  document.getElementById('doc-title-input').value = file.name;
  document.getElementById('editor').value = file.content || '';
  document.getElementById('toolbar-title').innerHTML = '<span>' + escHtml(file.name) + '</span>';
  updatePreview();
  updateStatus();
}

function onDocTitleChange()
{
  if ((!currentFileId) || (files[currentFileId].type !== 'doc'))
    return;

  files[currentFileId].name = (document.getElementById('doc-title-input').value || files[currentFileId].name);
  document.getElementById('toolbar-title').innerHTML = '<span>' + escHtml(files[currentFileId].name) + '</span>';
  scheduleSave();
  renderFileList();
}

function onEditorChange()
{
  if (!currentFileId || !files[currentFileId] || files[currentFileId].type !== 'doc')
    return;

  files[currentFileId].content = document.getElementById('editor').value;

  updatePreview();
  updateStatus();
  scheduleSave();
}

function scheduleSave()
{
  clearTimeout(saveTimer);
  saveTimer = setTimeout
  (
    function()
    {
        if (currentFileId)
        {
          files[currentFileId].modified = Date.now();
          persistFileEntry(currentFileId);
        }
    },
    800
  );
}

// ── MARKDOWN EXTENSIONS (styled blocks, font spans, table of contents) ──
// Adds Pandoc-style commands on top of standard markdown:
//   :: ... ::                                               wraps a block, left-aligned
//   ::: ... :::                                             wraps a block, right-aligned
//   ;; ... ;;                                                wraps a block, centered
//   ;;; ... ;;;                                              wraps a block, justified
//   ||Georgia ... ||                                        wraps a block in a font-family div
//   :::font="Georgia" ... :::                              older keyword form, still parses
//   [text]{font="Georgia"}                                 changes the font of an inline passage
//   [TOC]                                                  expands into a generated table of contents
// Implemented as marked.js extensions (not a post-processing pass) so headings nested inside
// styled blocks are still visited in true document order, keeping TOC anchors in sync.
// Generic CSS family keywords must stay unquoted to mean what they say -
// font-family:'monospace' looks up a font literally NAMED "monospace" (and
// fails), while font-family:monospace is the actual generic fallback.
var CSS_GENERIC_FONT_FAMILIES = ['serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui', 'ui-serif', 'ui-sans-serif', 'ui-monospace', 'ui-rounded', 'math', 'emoji', 'fangsong'];

function cssFontFamilyValue(name)
{
  name = (name || '').trim();

  if (CSS_GENERIC_FONT_FAMILIES.indexOf(name.toLowerCase()) !== -1)
    return name.toLowerCase();

  return "'" + name.replace(/\\/g, '\\\\').replace(/'/g, "\\'") + "'";
}

function parseBlockDirectives(line)
{
  var style = '',
      re = /([A-Za-z]+)(?:=("([^"]*)"|'([^']*)'|[^\s]+))?/g,
      match;

  while ((match = re.exec(line)))
  {
    var key = match[1].toLowerCase(),
        hasValue = match[2] !== undefined,
        value = match[3] !== undefined ? match[3] : (match[4] !== undefined ? match[4] : match[2]);

    if (!hasValue && (key === 'left' || key === 'center' || key === 'right' || key === 'justify'))
      style += 'text-align:' + key + ';';
    else if (hasValue && key === 'font' && value)
      style += 'font-family:' + escHtml(cssFontFamilyValue(value)) + ';';
  }

  return style;
}

var checkboxRenderIndex = 0;

// ── MATH (KaTeX) ──
// LaTeX-style formulas in Documents: $inline$ and $$display$$ blocks render
// through the vendored KaTeX build. Registered as marked tokenizers (not a
// post-pass) so markdown emphasis rules never mangle the TeX source.

function katexRenderTex(tex, displayMode)
{
  if (typeof katex === 'undefined')
    return '<code>' + escHtml(tex) + '</code>';
  try { return katex.renderToString(tex, { displayMode: displayMode, throwOnError: false }); }
  catch(e) { return '<code>' + escHtml(tex) + '</code>'; }
}

var mathBlockExtension =
{
  name: 'mathBlock',
  level: 'block',
  start: function(src) { var i = src.indexOf('$$'); return i === -1 ? undefined : i; },
  tokenizer: function(src)
  {
    var m = /^\$\$([\s\S]+?)\$\$/.exec(src);
    if (m) return { type: 'mathBlock', raw: m[0], text: m[1].trim() };
  },
  renderer: function(token) { return '<div class="math-block">' + katexRenderTex(token.text, true) + '</div>\n'; }
};

var mathInlineExtension =
{
  name: 'mathInline',
  level: 'inline',
  start: function(src) { var i = src.indexOf('$'); return i === -1 ? undefined : i; },
  tokenizer: function(src)
  {
    // Pandoc-style guard against currency ("$5 and $10"): the TeX must not
    // start or end with whitespace, and the closing $ can't precede a digit.
    var m = /^\$(?!\s)((?:\\\$|[^$\n])+?)\$(?!\d)/.exec(src);
    if (m && !/\s$/.test(m[1])) return { type: 'mathInline', raw: m[0], text: m[1].replace(/\\\$/g, '$') };
  },
  renderer: function(token) { return katexRenderTex(token.text, false); }
};

function configureMarkedExtensions()
{
  if (typeof marked === 'undefined')
    return;

  marked.use({ extensions: [mathBlockExtension, mathInlineExtension] });

  marked.use
  (
    {
      renderer:
      {
        heading: function(text, level)
        {
          var entry = documentHeadings[headingRenderCursor],
              id = '';

          if (entry)
          {
            id = entry.slug;
            headingRenderCursor++;
          }

          return '<h' + level + (id ? (' id="' + id + '"') : '') + '>' + text + '</h' + level + '>\n';
        },
        code: function(code, infostring)
        {
          var lang = (infostring || '').match(/^\S*/);
          lang = lang ? lang[0].toLowerCase() : '';

          // Mermaid blocks render to a live diagram instead of a literal code
          // block; returning false for any other language falls back to
          // marked's default <pre><code> rendering.
          if (lang !== 'mermaid')
            return false;

          return '<div class="mermaid">' + escHtml(code.replace(/\n$/, '')) + '\n</div>\n';
        },
        checkbox: function(checked)
        {
          var idx = checkboxRenderIndex++;
          return '<input type="checkbox" class="preview-checkbox" data-idx="' + idx + '"' + (checked ? ' checked' : '') + '>';
        },
        image: function(href, title, text)
        {
          if (!href || href.indexOf(FILE_LINK_SCHEME) !== 0)
            return false;

          var id = decodeURIComponent(href.slice(FILE_LINK_SCHEME.length)),
              file = files[id];

          if (!file)
            return '<div class="embed-missing">' + escHtml('[File not found: ' + (text || id) + ']') + '</div>';

          if (file.type === 'graph')
          {
            var mermaidMatch = /```mermaid\n([\s\S]*?)```/i.exec(file.content || '');
            if (!mermaidMatch)
              return '<div class="embed-missing">' + escHtml('[No diagram in: ' + file.name + ']') + '</div>';
            return  '<figure class="embed-graph">' +
                      '<div class="mermaid">' + escHtml(mermaidMatch[1].replace(/\n$/, '')) + '</div>' +
                      (text ? '<figcaption class="embed-caption">' + escHtml(text) + '</figcaption>' : '') +
                    '</figure>';
          }

          if (file.type === 'sheet')
          {
            var defs = parseSheetChartsFromContent(file.content || '');
            if (!defs.length)
              return '<div class="embed-missing">' + escHtml('[No charts in: ' + file.name + ']') + '</div>';
            var cells = parseSheetCellsForEmbed(file.content || '');
            return  '<figure class="embed-sheet">' +
                      defs.map(function(def)
                      {
                        try
                        {
                          var config = buildChartJsConfigFromCells(def, cells);
                          return '<div class="embed-sheet-chart" data-chart-config="' + escAttr(JSON.stringify(config)) + '"><canvas></canvas></div>';
                        }
                        catch(e) { return ''; }
                      }).join('') +
                      (text ? '<figcaption class="embed-caption">' + escHtml(text) + '</figcaption>' : '') +
                    '</figure>';
          }

          return false;
        }
      },
      extensions:
      [
        {
          name: 'styledBlock',
          level: 'block',
          start: function(src)
          {
            var match = /^:::[ \t]*\S[^\n]*$/m.exec(src);
            return match ? match.index : undefined;
          },
          tokenizer: function(src)
          {
            var match = /^:::[ \t]*(\S[^\n]*)\n([\s\S]*?)\n:::(?:\n+|$)/.exec(src),
                style = match ? parseBlockDirectives(match[1]) : '';

            if (!match || !style)
              return undefined;

            return  {
                      type: 'styledBlock',
                      raw: match[0],
                      style: style,
                      tokens: this.lexer.blockTokens(match[2], [])
                    };
          },
          renderer: function(token)
          {
            return '<div style="' + token.style + '">\n' + this.parser.parse(token.tokens) + '</div>\n';
          }
        },
        {
          // Bare-symbol alignment blocks (see ALIGN_BLOCK_SYMBOLS near wrapBlock,
          // which is what writes these). The symbol must be alone on its line -
          // no trailing text - and the closer must repeat the exact same symbol
          // (matched via the \1 backreference), so this never collides with the
          // ":::keyword" form above (which always requires extra text after :::).
          name: 'alignSymbolBlock',
          level: 'block',
          start: function(src)
          {
            var match = /^(?:::(?!:)|:::(?!:)|;;(?!;)|;;;(?!;))[ \t]*$/m.exec(src);
            return match ? match.index : undefined;
          },
          tokenizer: function(src)
          {
            var match = /^(::(?!:)|:::(?!:)|;;(?!;)|;;;(?!;))[ \t]*\n([\s\S]*?)\n\1[ \t]*(?:\n+|$)/.exec(src);

            if (!match)
              return undefined;

            var align = { '::': 'left', ':::': 'right', ';;': 'center', ';;;': 'justify' }[match[1]];

            return  {
                      type: 'alignSymbolBlock',
                      raw: match[0],
                      style: 'text-align:' + align + ';',
                      tokens: this.lexer.blockTokens(match[2], [])
                    };
          },
          renderer: function(token)
          {
            return '<div style="' + token.style + '">\n' + this.parser.parse(token.tokens) + '</div>\n';
          }
        },
        {
          // Bare-symbol font blocks (see wrapFontBlock, which is what writes
          // these): || followed directly by the font name - no "font=" keyword,
          // no quotes - closed by a bare ||. The older ":::font="..." ... :::"
          // form (parseBlockDirectives, above) still parses correctly for
          // documents written before this existed.
          name: 'fontSymbolBlock',
          level: 'block',
          start: function(src)
          {
            var match = /^\|\|[ \t]*\S[^\n]*$/m.exec(src);
            return match ? match.index : undefined;
          },
          tokenizer: function(src)
          {
            var match = /^\|\|[ \t]*(\S[^\n]*)\n([\s\S]*?)\n\|\|[ \t]*(?:\n+|$)/.exec(src);

            if (!match)
              return undefined;

            return  {
                      type: 'fontSymbolBlock',
                      raw: match[0],
                      style: 'font-family:' + escHtml(cssFontFamilyValue(match[1].trim())) + ';',
                      tokens: this.lexer.blockTokens(match[2], [])
                    };
          },
          renderer: function(token)
          {
            return '<div style="' + token.style + '">\n' + this.parser.parse(token.tokens) + '</div>\n';
          }
        },
        {
          name: 'fontSpan',
          level: 'inline',
          start: function(src)
          {
            var match = /\[[^\]\n]+\]\{font=/.exec(src);
            return match ? match.index : undefined;
          },
          tokenizer: function(src)
          {
            var match = /^\[([^\]\n]+)\]\{font=("([^"]*)"|'([^']*)'|[^\s}]+)\}/.exec(src);

            if (!match)
              return undefined;

            var fontName = match[3] !== undefined ? match[3] : (match[4] !== undefined ? match[4] : match[2]);

            return  {
                      type: 'fontSpan',
                      raw: match[0],
                      fontName: fontName,
                      tokens: this.lexer.inlineTokens(match[1])
                    };
          },
          renderer: function(token)
          {
            return '<span style="font-family:' + escHtml(cssFontFamilyValue(token.fontName)) + ';">' +
                   this.parser.parseInline(token.tokens) + '</span>';
          }
        },
        {
          name: 'tocBlock',
          level: 'block',
          start: function(src)
          {
            var match = /^\[TOC\][ \t]*$/im.exec(src);
            return match ? match.index : undefined;
          },
          tokenizer: function(src)
          {
            var match = /^\[TOC\][ \t]*(?:\n+|$)/i.exec(src);
            return match ? { type: 'tocBlock', raw: match[0] } : undefined;
          },
          renderer: function()
          {
            return renderTOC();
          }
        },
        {
          // Roman-numeral ordered lists: consecutive lines like "i. item" or
          // "II) item" render as a real <ol> with roman markers (lowercase
          // and uppercase are separate lists). Arabic "1." lists are standard
          // Markdown and unchanged. A lone item only counts when it starts
          // at i/I, so abbreviations like "c. 1450" stay plain text.
          name: 'romanList',
          level: 'block',
          start: function(src)
          {
            var m = /^ {0,3}[ivxlcdmIVXLCDM]+[.)][ \t]/m.exec(src);
            return m ? m.index : undefined;
          },
          tokenizer: function(src)
          {
            var itemRe = /^ {0,3}([ivxlcdm]+|[IVXLCDM]+)[.)][ \t]+([^\n]*)(?:\n|$)/,
                rest = src,
                raw = '',
                items = [],
                lower = null,
                startNum = 0;

            while (true)
            {
              var m = itemRe.exec(rest);
              if (!m) break;

              var isLower = m[1] === m[1].toLowerCase();
              if (lower === null) lower = isLower;
              else if (lower !== isLower) break; // case switch = new list

              var value = romanToInt(m[1]);
              if (!value) break;
              if (!items.length) startNum = value;

              raw += m[0];
              items.push(m[2]);
              rest = rest.slice(m[0].length);
            }

            if (!items.length) return undefined;
            if (items.length === 1 && startNum !== 1) return undefined;

            var lexer = this.lexer;
            return {
              type: 'romanList',
              raw: raw,
              lower: lower,
              startNum: startNum,
              items: items.map(function(text){ return lexer.inlineTokens(text); })
            };
          },
          renderer: function(token)
          {
            var parser = this.parser;
            return '<ol style="list-style-type:' + (token.lower ? 'lower-roman' : 'upper-roman') + ';" start="' + token.startNum + '">' +
                   token.items.map(function(tokens){ return '<li>' + parser.parseInline(tokens) + '</li>'; }).join('') +
                   '</ol>\n';
          }
        },
        {
          // Doc metadata fields: key: value pairs in braces, rendered as a
          // styled card. One line — {type: Character, status: Active} — lays
          // the fields out side by side; putting the braces on their own
          // lines with one "key: value," per line renders the same card
          // vertically, one field per row.
          name: 'docFields',
          level: 'block',
          start: function(src)
          {
            var m = /^\{/m.exec(src);
            return m ? m.index : undefined;
          },
          tokenizer: function(src)
          {
            // Single-line form: everything inside one pair of braces.
            var match = /^\{([^{}\n]+)\}[ \t]*(?:\n|$)/.exec(src);
            if (match)
            {
              var pairs = [];
              match[1].split(',').forEach(function(part)
              {
                var colon = part.indexOf(':');
                if (colon > -1)
                {
                  var key = part.slice(0, colon).trim(),
                      val = part.slice(colon + 1).trim();
                  if (key) pairs.push({ key: key, value: val });
                }
              });

              if (!pairs.length) return undefined;

              return { type: 'docFields', raw: match[0], pairs: pairs, vertical: false };
            }

            // Multi-line form: '{' alone on its line, one pair per line, '}'
            // alone on its line. Every non-empty inner line must look like
            // "key: value" (trailing comma optional) — anything else is left
            // for other tokenizers, so stray braced text stays untouched.
            match = /^\{[ \t]*\r?\n([\s\S]*?)\r?\n[ \t]*\}[ \t]*(?:\r?\n|$)/.exec(src);
            if (!match) return undefined;

            var vPairs = [], valid = true;

            match[1].split('\n').forEach(function(line)
            {
              var text = line.trim().replace(/,+$/, '');
              if (!text) return;

              var colon = text.indexOf(':');
              if (colon < 1) { valid = false; return; }

              vPairs.push({ key: text.slice(0, colon).trim(), value: text.slice(colon + 1).trim() });
            });

            if (!valid || !vPairs.length) return undefined;

            return { type: 'docFields', raw: match[0], pairs: vPairs, vertical: true };
          },
          renderer: function(token)
          {
            if (token.vertical)
            {
              return '<div class="doc-fields doc-fields-vertical">' +
                     token.pairs.map(function(p)
                     {
                       return '<b class="doc-field-key">' + escHtml(p.key) + '</b><span class="doc-field-value">' + escHtml(p.value) + '</span>';
                     }).join('') +
                     '</div>\n';
            }

            return '<div class="doc-fields">' +
                   token.pairs.map(function(p)
                   {
                     return '<span class="doc-field"><b>' + escHtml(p.key) + '</b> ' + escHtml(p.value) + '</span>';
                   }).join('') +
                   '</div>\n';
          }
        },
        {
          // Inline doc tag: #tagname renders as a clickable pill that filters the sidebar.
          // Only matches #word (no space after #), so it never collides with headings.
          name: 'docTag',
          level: 'inline',
          start: function(src) { return src.indexOf('#'); },
          tokenizer: function(src)
          {
            var match = /^#([\w-]+)/.exec(src);
            return match ? { type: 'docTag', raw: match[0], tag: match[1] } : undefined;
          },
          renderer: function(token)
          {
            return '<span class="doc-tag" onclick="filterByTag(\'' + escAttr(token.tag) + '\')">#' + escHtml(token.tag) + '</span>';
          }
        }
      ]
    }
  );
}

configureMarkedExtensions();

// Loose roman-numeral parser for the romanList extension: subtractive
// notation handled, 0 returned for anything that isn't roman.
function romanToInt(text)
{
  var values = { i: 1, v: 5, x: 10, l: 50, c: 100, d: 500, m: 1000 },
      total = 0,
      prev = 0;

  text = text.toLowerCase();

  for (var k = text.length - 1; k >= 0; k--)
  {
    var v = values[text[k]];
    if (!v) return 0;

    if (v < prev)
      total -= v;
    else
    {
      total += v;
      prev = v;
    }
  }

  return (total > 0 && total < 4000) ? total : 0;
}

function configureMermaid()
{
  if (typeof mermaid === 'undefined')
    return;

  mermaid.initialize({ startOnLoad: false });
}

configureMermaid();

function extractPlainText(tokens)
{
  if (!tokens || !tokens.length)
    return '';

  return tokens.map
  (
    function(token)
    {
      return token.tokens ? extractPlainText(token.tokens) : (token.text || token.raw || '');
    }
  ).join('');
}

function makeSlug(text)
{
  var base = text.toLowerCase()
                  .trim()
                  .replace(/[^a-z0-9\s-]/g, '')
                  .replace(/\s+/g, '-')
                  .replace(/-+/g, '-')
                  .replace(/^-|-$/g, '') || 'section';

  var count = slugCounts[base] || 0;
  slugCounts[base] = count + 1;

  return count === 0 ? base : (base + '-' + count);
}

function collectHeadings(tokens)
{
  tokens.forEach
  (
    function(token)
    {
      if (token.type === 'heading')
      {
        var text = extractPlainText(token.tokens);

        documentHeadings.push
        (
          {
            depth: token.depth,
            text: text,
            slug: makeSlug(text)
          }
        );
      }
      else if (token.type === 'list')
      {
        token.items.forEach
        (
          function(item)
          {
            collectHeadings(item.tokens);
          }
        );
      }
      else if (token.tokens)
      {
        collectHeadings(token.tokens);
      }
    }
  );
}

function renderTOC()
{
  if (!documentHeadings.length)
    return '<div class="toc toc-empty">No headings yet.</div>';

  var minDepth = documentHeadings.reduce
  (
    function(min, h)
    {
      return Math.min(min, h.depth);
    },
    documentHeadings[0].depth
  );

  var html = '<nav class="toc">',
      openDepth = minDepth - 1;

  documentHeadings.forEach
  (
    function(h)
    {
      while (openDepth < h.depth)
      {
        html += '<ul>';
        openDepth++;
      }

      while (openDepth > h.depth)
      {
        html += '</ul>';
        openDepth--;
      }

      html += '<li><a href="#' + h.slug + '">' + escHtml(h.text) + '</a></li>';
    }
  );

  while (openDepth >= minDepth)
  {
    html += '</ul>';
    openDepth--;
  }

  html += '</nav>';

  return html;
}

function updatePreview()
{
  document.getElementById('preview-title').textContent = document.getElementById('doc-title-input').value;

  var source = document.getElementById('editor').value,
      html = source;

  if (typeof marked !== 'undefined')
  {
    documentHeadings = [];
    slugCounts = {};
    headingRenderCursor = 0;
    checkboxRenderIndex = 0;

    var tokens = marked.lexer(source);
    collectHeadings(tokens);

    html = marked.parser(tokens);
  }

  var previewContent = document.getElementById('preview-content');
  previewContent.innerHTML = html;
  renderMermaidDiagrams(previewContent);
  renderEmbeddedSheetCharts(previewContent);
  renderPreviewBacklinks();
}

function renderPreviewBacklinks()
{
  var panel = document.getElementById('preview-backlinks');
  if (!panel) return;
  if (!currentFileId) { panel.style.display = 'none'; return; }

  // Merge persisted index with live scan of loaded content
  var indexed = (backlinksIndex[currentFileId] || []).filter(function(id) { return !!files[id]; });
  var live    = computeBacklinks(currentFileId).map(function(r) { return r.id; });
  var allIds  = indexed.slice();
  live.forEach(function(id) { if (allIds.indexOf(id) < 0) allIds.push(id); });

  if (!allIds.length) { panel.style.display = 'none'; return; }

  panel.style.display = '';
  panel.innerHTML =
    '<div class="pbl-header">' +
      '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M6 3H3a1 1 0 0 0-1 1v9a1 1 0 0 0 1 1h9a1 1 0 0 0 1-1v-3"/><path d="M10 2h4v4"/><line x1="14" y1="2" x2="7" y2="9"/></svg>' +
      'Mentioned in' +
    '</div>' +
    '<div class="pbl-list">' +
      allIds.map(function(id)
      {
        var f = files[id];
        return '<div class="pbl-item" onclick="openFile(\'' + escAttr(id) + '\')">' +
          '<span class="pbl-icon">' + fileTypeIcon(f.type) + '</span>' +
          '<span class="pbl-name">' + escHtml(f.name) + '</span>' +
        '</div>';
      }).join('') +
    '</div>';
}

function updateStatus()
{
  var text = document.getElementById('editor').value;
  var words = text.trim() ? text.trim().split(/\s+/).length : 0;

  document.getElementById('word-count').textContent = words;
  document.getElementById('char-count').textContent = text.length;
}

document.getElementById('editor').addEventListener
(
  'keyup',
  function(target)
  {
    var editor = target.target;
    var lines = editor.value.substr(0, editor.selectionStart).split('\n');
    document.getElementById('cursor-pos').textContent = 'Ln ' + lines.length + ', Col ' + (lines[lines.length - 1].length + 1);
  }
);

function handleEditorKey(keyEvent)
{
  docTextUndo.noteKeydown(keyEvent);

  if (keyEvent.ctrlKey || keyEvent.metaKey)
  {
    if (!keyEvent.shiftKey && keyEvent.key.toLowerCase() === 'z')
    {
      keyEvent.preventDefault();

      if (docTextUndo.undo())
        onEditorChange();

      return;
    }

    if (keyEvent.key.toLowerCase() === 'y' || (keyEvent.shiftKey && keyEvent.key.toLowerCase() === 'z'))
    {
      keyEvent.preventDefault();

      if (docTextUndo.redo())
        onEditorChange();

      return;
    }

    if (keyEvent.key === 'b')
    {
      keyEvent.preventDefault();
      wrapSelectedText('**','**');
    }

    if (keyEvent.key === 'i')
    {
      keyEvent.preventDefault();
      wrapSelectedText('*','*');
    }

    if (keyEvent.key === 's')
    {
      keyEvent.preventDefault();
      saveToStorage();
    }
  }

  if (keyEvent.key === 'Tab')
  {
    keyEvent.preventDefault();

    var target = keyEvent.target,
        startSelection = target.selectionStart;

    target.value = target.value.slice(0, startSelection) + '  ' + target.value.slice(target.selectionEnd);
    target.selectionStart = target.selectionEnd = startSelection + 2;
    onEditorChange();
  }
}

document.addEventListener
(
  'keydown',
  function(keyEvent)
  {
    if ((keyEvent.ctrlKey || keyEvent.metaKey) && keyEvent.key==='n')
    {
      keyEvent.preventDefault();
      openNewModal();
    }

    if (keyEvent.key === 'Escape')
    {
      var functionHelpPanel = document.getElementById('function-help-panel');

      if (functionHelpPanel)
        functionHelpPanel.classList.remove('open');
    }
  }
);

// ── DOCS: FIND & REPLACE ──
// A small panel over the editor: find (with live match count and optional
// case sensitivity), step through matches, replace one or all. Replacements
// are single undo steps via the shared textarea undo controller.

function frPanel() { return document.getElementById('find-replace-panel'); }

function openFindReplace()
{
  var panel = frPanel(),
      editor = document.getElementById('editor');

  if (!panel || !editor)
    return;

  panel.classList.add('open');

  // A single-line selection in the editor becomes the initial query.
  var findInp = document.getElementById('fr-find'),
      selected = editor.value.slice(editor.selectionStart, editor.selectionEnd);

  if (selected && selected.indexOf('\n') === -1)
    findInp.value = selected;

  findInp.focus();
  findInp.select();
  frUpdateCount();
}

function closeFindReplace()
{
  var panel = frPanel();
  if (panel) panel.classList.remove('open');
  document.getElementById('editor').focus();
}

function toggleFindReplace()
{
  if (frPanel().classList.contains('open'))
    closeFindReplace();
  else
    openFindReplace();
}

function frQuery()
{
  return  {
            find: document.getElementById('fr-find').value,
            cs: document.getElementById('fr-case').checked
          };
}

function frCountMatches()
{
  var q = frQuery();
  if (!q.find) return 0;

  var editor = document.getElementById('editor'),
      hay = q.cs ? editor.value : editor.value.toLowerCase(),
      needle = q.cs ? q.find : q.find.toLowerCase(),
      count = 0,
      at = 0;

  while ((at = hay.indexOf(needle, at)) !== -1)
  {
    count++;
    at += needle.length;
  }

  return count;
}

function frUpdateCount()
{
  var q = frQuery();
  document.getElementById('fr-count').textContent = q.find ? (frCountMatches() + ' found') : '';
}

// Selects a range and scrolls the textarea so it's visible (textareas don't
// scroll to a programmatic selection on their own).
function frSelectRange(editor, start, end)
{
  editor.focus();
  editor.setSelectionRange(start, end);

  var lineHeight = parseFloat(getComputedStyle(editor).lineHeight) || 20,
      lineIdx = editor.value.slice(0, start).split('\n').length - 1;

  editor.scrollTop = Math.max(0, lineIdx * lineHeight - editor.clientHeight / 2);
}

function frFindNext()
{
  var q = frQuery();
  if (!q.find) return;

  var editor = document.getElementById('editor'),
      hay = q.cs ? editor.value : editor.value.toLowerCase(),
      needle = q.cs ? q.find : q.find.toLowerCase(),
      at = hay.indexOf(needle, editor.selectionEnd || 0);

  if (at === -1)
    at = hay.indexOf(needle); // wrap back to the top

  if (at === -1)
    return;

  frSelectRange(editor, at, at + q.find.length);
}

function frReplaceOne()
{
  var q = frQuery();
  if (!q.find) return;

  var editor = document.getElementById('editor'),
      selText = editor.value.slice(editor.selectionStart, editor.selectionEnd),
      onMatch = q.cs ? (selText === q.find) : (selText.toLowerCase() === q.find.toLowerCase());

  // First invocation just lands on the next match; the next one replaces it.
  if (!onMatch)
  {
    frFindNext();
    return;
  }

  var replaceWith = document.getElementById('fr-replace').value,
      start = editor.selectionStart;

  docTextUndo.forceSnapshot();
  editor.value = editor.value.slice(0, start) + replaceWith + editor.value.slice(editor.selectionEnd);
  editor.setSelectionRange(start + replaceWith.length, start + replaceWith.length);
  onEditorChange();
  frUpdateCount();
  frFindNext();
}

function frReplaceAll()
{
  var q = frQuery();
  if (!q.find || !frCountMatches()) return;

  var editor = document.getElementById('editor'),
      replaceWith = document.getElementById('fr-replace').value;

  docTextUndo.forceSnapshot();

  if (q.cs)
    editor.value = editor.value.split(q.find).join(replaceWith);
  else
  {
    var pattern = new RegExp(q.find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
    // Function replacement so '$' in the replacement text stays literal.
    editor.value = editor.value.replace(pattern, function(){ return replaceWith; });
  }

  onEditorChange();
  frUpdateCount();
}

function frKey(keyEvent)
{
  if (keyEvent.key === 'Enter')
  {
    keyEvent.preventDefault();
    if (keyEvent.target.id === 'fr-replace') frReplaceOne();
    else frFindNext();
  }

  if (keyEvent.key === 'Escape')
    closeFindReplace();
}

// Ctrl/Cmd+F and Ctrl/Cmd+H inside the doc editor open the panel.
(function()
{
  var editor = document.getElementById('editor');
  if (!editor) return;

  editor.addEventListener('keydown', function(e)
  {
    if ((e.ctrlKey || e.metaKey) && (e.key === 'f' || e.key === 'F' || e.key === 'h' || e.key === 'H'))
    {
      e.preventDefault();
      openFindReplace();
    }
  });
})();

function wrapSelectedText(before, after)
{
  docTextUndo.forceSnapshot();

  var editor = document.getElementById('editor'),
      selectionStart = editor.selectionStart,
      selectionEnd = editor.selectionEnd,
      selection = editor.value.slice(selectionStart, selectionEnd);
  
  editor.value = editor.value.slice(0, selectionStart) + before + selection + after + editor.value.slice(selectionEnd);
  editor.selectionStart = selectionStart + before.length;
  editor.selectionEnd = selectionEnd + after.length;

  editor.focus();
  onEditorChange();
}

function insertLine(prefix)
{
  docTextUndo.forceSnapshot();

  var editor = document.getElementById('editor'),
      selectionStart = editor.selectionStart,
      lineStart = editor.value.lastIndexOf('\n', selectionStart - 1) + 1;
  
  editor.value = editor.value.slice(0, lineStart) + prefix + editor.value.slice(lineStart);
  editor.selectionStart = editor.selectionEnd = lineStart + prefix.length + (selectionStart - lineStart);

  editor.focus();
  onEditorChange();
}

function applyHeading(prefix)
{
    if (!prefix)
      return;

    insertLine(prefix);
    document.getElementById('heading-sel').value = '';
}

function insertCodeBlock()
{
  docTextUndo.forceSnapshot();

  var editor = document.getElementById('editor'),
      selectionStart = editor.selectionStart,
      selectionEnd = editor.selectionEnd;

  editor.value = editor.value.slice(0, selectionStart) + '\n```\n' + editor.value.slice(selectionStart, selectionEnd) + '\n```\n' + editor.value.slice(selectionEnd);
  editor.selectionStart = editor.selectionEnd = selectionStart + 5;
  editor.focus();
  onEditorChange();
}

function openTableModal()
{
  document.getElementById('table-modal').classList.add('open');
  document.getElementById('table-rows').value = 2;
  document.getElementById('table-cols').value = 3;

  setTimeout
  (
    function()
    {
      document.getElementById('table-rows').focus();
    },
    50
  );
}

function closeTableModal()
{
  document.getElementById('table-modal').classList.remove('open');
}

function createTable()
{
  docTextUndo.forceSnapshot();

  var rows = Math.max(1, Math.min(50, parseInt(document.getElementById('table-rows').value, 10) || 1)),
      cols = Math.max(1, Math.min(20, parseInt(document.getElementById('table-cols').value, 10) || 1)),
      headerCells = [],
      dividerCells = [],
      i, j;

  for (j = 0; j < cols; j++)
  {
    headerCells.push('Column ' + (j + 1));
    dividerCells.push('----------');
  }

  var table = '\n| ' + headerCells.join(' | ') + ' |\n' +
              '| ' + dividerCells.join(' | ') + ' |\n';

  for (i = 0; i < rows; i++)
  {
    var rowCells = [];

    for (j = 0; j < cols; j++)
      rowCells.push('Cell');

    table += '| ' + rowCells.join(' | ') + ' |\n';
  }

  table += '\n';

  var editor = document.getElementById('editor'),
      selectionStart = editor.selectionStart;

  editor.value = editor.value.slice(0, selectionStart) + table + editor.value.slice(selectionStart);
  editor.selectionStart = editor.selectionEnd = selectionStart + table.length;

  closeTableModal();
  editor.focus();
  onEditorChange();
}

document.getElementById('table-modal').addEventListener
(
  'click',
  function(e)
  {
    if (e.target === document.getElementById('table-modal'))
      closeTableModal();
  }
);

function insertLink()
{
  var url = prompt('URL:', 'https://');

  if (!url)
    return;

  wrapSelectedText('[', ']('+url+')');
}

// ── CROSS-FILE LINKS (Docs only) ──
// Written as ordinary markdown links whose target uses a custom scheme -
// [text](lorekeep://file/<id>) - so a click handler in the preview can tell
// them apart from real URLs and route to openFile() instead of navigating.
// The id is URL-encoded since Work Folder ids are relative paths (can
// contain "/") and could in principle contain other URL-unsafe characters.
var FILE_LINK_SCHEME = 'lorekeep://file/';

function openFileLinkModal()
{
  document.getElementById('file-link-search').value = '';
  renderFileLinkList();
  document.getElementById('file-link-modal').classList.add('open');
  document.getElementById('file-link-search').focus();

  // In Work Folder mode, eagerly pre-load content for embeddable files so the
  // Embed button works immediately when clicked (content is otherwise only
  // loaded on explicit openFile).
  if (workFolderRoot)
  {
    Object.keys(files).forEach(function(id)
    {
      var f = files[id];
      if ((f.type === 'sheet' || f.type === 'graph') && !f.contentLoaded)
      {
        Platform.readWorkFile(workFolderRoot, id)
          .then(function(c){ f.content = c; f.contentLoaded = true; })
          .catch(function(){});
      }
    });
  }
}

function closeFileLinkModal()
{
  document.getElementById('file-link-modal').classList.remove('open');
}

document.getElementById('file-link-modal').addEventListener
(
  'click',
  function(e)
  {
    if (e.target === document.getElementById('file-link-modal'))
      closeFileLinkModal();
  }
);

// Folders the user collapsed inside the Link-file modal (session-scoped).
// Everything starts expanded: the point of the tree is browsing by location.
var fileLinkCollapsed = new Set();

function toggleFileLinkFolder(path)
{
  if (fileLinkCollapsed.has(path))
    fileLinkCollapsed.delete(path);
  else
    fileLinkCollapsed.add(path);

  renderFileLinkList();
}

function fileLinkRowHtml(id, f, depth, subtitle)
{
  var canEmbed = (f.type === 'graph' || f.type === 'sheet');

  return  '<div class="file-link-item" style="margin-left:' + (depth * 16) + 'px" onclick="chooseFileLink(\'' + escAttr(id) + '\')">' +
            '<span class="file-icon">' + fileTypeIcon(f.type) + '</span>' +
            '<span class="file-link-item-name">' + escHtml(f.name) +
              (subtitle ? '<span class="file-link-item-path">' + escHtml(subtitle) + '</span>' : '') +
            '</span>' +
            (canEmbed ? '<button class="file-link-embed-btn" onclick="event.stopPropagation();chooseFileEmbed(\'' + escAttr(id) + '\')">Embed</button>' : '') +
          '</div>';
}

// Mirrors the sidebar's folder tree so files can be found by where they
// live, not just by name. Searching flattens the tree (matches from any
// folder) but keeps each hit's folder path visible as a subtitle.
function renderFileLinkList()
{
  var query = document.getElementById('file-link-search').value.trim().toLowerCase(),
      list = document.getElementById('file-link-list'),
      ids = Object.keys(files).filter(function(id)
      {
        return id !== SHARED_TMP_ID && id !== currentFileId; // no self-links
      });

  function byName(a, b) { return files[a].name.localeCompare(files[b].name); }

  if (query)
  {
    var matches = ids.filter(function(id) { return files[id].name.toLowerCase().indexOf(query) !== -1; })
                     .sort(byName);

    list.innerHTML = matches.length
                      ?
                        matches.map(function(id)
                        {
                          return fileLinkRowHtml(id, files[id], 0, files[id].folder ? files[id].folder + '/' : '');
                        }).join('')
                      :
                        '<div class="history-empty">No matching files.</div>';
    return;
  }

  // No folders (e.g. browser-storage mode): flat list, most recent first
  if (Object.keys(folders).length === 0)
  {
    var flat = ids.sort(function(a, b) { return (files[b].modified || 0) - (files[a].modified || 0); });

    list.innerHTML = flat.length
                      ? flat.map(function(id) { return fileLinkRowHtml(id, files[id], 0, ''); }).join('')
                      : '<div class="history-empty">No other files yet.</div>';
    return;
  }

  var byFolder = {},
      childFolders = {};

  ids.forEach(function(id)
  {
    var parent = files[id].folder || '';
    (byFolder[parent] = byFolder[parent] || []).push(id);
  });

  Object.keys(folders).forEach(function(path)
  {
    var parent = folders[path].parent || '';
    (childFolders[parent] = childFolders[parent] || []).push(path);
  });

  function renderLevel(parentPath, depth)
  {
    var html = '';

    (childFolders[parentPath] || []).slice().sort(function(a, b)
    {
      return folders[a].name.localeCompare(folders[b].name);
    })
    .forEach(function(path)
    {
      var collapsed = fileLinkCollapsed.has(path);

      html += '<div class="file-link-folder" style="margin-left:' + (depth * 16) + 'px" onclick="toggleFileLinkFolder(\'' + escAttr(path) + '\')">' +
                '<span class="folder-caret' + (collapsed ? '' : ' expanded') + '">' + CHEVRON_ICON + '</span>' +
                '<span class="file-icon">' + FOLDER_ICON + '</span>' +
                '<span class="file-link-item-name">' + escHtml(folders[path].name) + '</span>' +
              '</div>';

      if (!collapsed)
        html += renderLevel(path, depth + 1);
    });

    (byFolder[parentPath] || []).slice().sort(byName).forEach(function(id)
    {
      html += fileLinkRowHtml(id, files[id], depth, '');
    });

    return html;
  }

  list.innerHTML = renderLevel('', 0) || '<div class="history-empty">No other files yet.</div>';
}

function chooseFileLink(fileId)
{
  closeFileLinkModal();
  insertFileLink(fileId);
}

function insertFileLink(fileId)
{
  var target = files[fileId];

  if (!target)
    return;

  docTextUndo.forceSnapshot();

  var editor = document.getElementById('editor'),
      selectionStart = editor.selectionStart,
      selectionEnd = editor.selectionEnd,
      selection = editor.value.slice(selectionStart, selectionEnd),
      linkText = selection || target.name,
      markdown = '[' + linkText + '](' + FILE_LINK_SCHEME + encodeURIComponent(fileId) + ')';

  editor.value = editor.value.slice(0, selectionStart) + markdown + editor.value.slice(selectionEnd);
  editor.selectionStart = editor.selectionEnd = selectionStart + markdown.length;

  editor.focus();
  onEditorChange();
}

async function chooseFileEmbed(fileId)
{
  closeFileLinkModal();

  // In Work Folder mode, ensure content is loaded before embedding — file
  // content is lazy-loaded only on openFile, so an unvisited sheet would
  // otherwise embed with undefined content and show "No charts in: …".
  var f = files[fileId];
  if (f && workFolderRoot && !f.contentLoaded)
  {
    try
    {
      f.content = await Platform.readWorkFile(workFolderRoot, fileId);
      f.contentLoaded = true;
    }
    catch(e) { console.warn('Could not load file content for embed', e); }
  }

  insertFileEmbed(fileId);
}

function insertFileEmbed(fileId)
{
  var target = files[fileId];

  if (!target)
    return;

  docTextUndo.forceSnapshot();

  var editor = document.getElementById('editor'),
      selectionStart = editor.selectionStart,
      markdown = '![' + (target.name || '') + '](' + FILE_LINK_SCHEME + encodeURIComponent(fileId) + ')';

  editor.value = editor.value.slice(0, selectionStart) + markdown + editor.value.slice(selectionStart);
  editor.selectionStart = editor.selectionEnd = selectionStart + markdown.length;

  editor.focus();
  onEditorChange();
}

function parseSheetChartsFromContent(content)
{
  var lines = (content || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n'),
      dataStart = 0;

  if (lines[0] === '---')
  {
    var end = lines.indexOf('---', 1);
    dataStart = end + 2;
  }

  var cellLines = lines.slice(dataStart),
      chartsMarker = cellLines.indexOf('```charts');

  if (chartsMarker === -1) return [];

  var chartsEnd = cellLines.indexOf('```', chartsMarker + 1);

  if (chartsEnd === -1) return [];

  try { return JSON.parse(cellLines.slice(chartsMarker + 1, chartsEnd).join('\n')) || []; }
  catch(e) { return []; }
}

function parseSheetCellsForEmbed(content)
{
  var cells = {},
      lines = (content || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n'),
      dataStart = 0;

  if (lines[0] === '---')
  {
    var end = lines.indexOf('---', 1);
    dataStart = end + 2;
  }

  var cellLines = lines.slice(dataStart),
      chartsMarker = cellLines.indexOf('```charts');

  if (chartsMarker !== -1) cellLines = cellLines.slice(0, chartsMarker);

  // Multi-tab sheets: embeds show the FIRST page; a second `# name` heading
  // ends it (the first heading, if any, is just that page's name).
  var headingsSeen = 0, firstPageEnd = cellLines.length;
  for (var li = 0; li < cellLines.length; li++)
  {
    if (cellLines[li].charAt(0) === '#' && ++headingsSeen === 2) { firstPageEnd = li; break; }
  }
  cellLines = cellLines.slice(0, firstPageEnd);

  cellLines.forEach(function(line)
  {
    if (!line.trim()) return;
    var refEnd = line.indexOf('=');
    if (refEnd === -1) return;
    var ref = line.slice(0, refEnd),
        rest = line.slice(refEnd + 1),
        formulaEnd = rest.indexOf('=');
    cells[ref] = (formulaEnd === -1) ? rest : rest.slice(formulaEnd + 1);
  });

  return cells;
}

function getChartRangeGridFromCells(range, cells)
{
  var dims = parseChartRange(range);
  if (!dims) return null;
  var grid = [];
  for (var r = dims.rowStart; r <= dims.rowEnd; r++)
  {
    var row = [];
    for (var c = dims.colStart; c <= dims.colEnd; c++)
      row.push(cells[colName(c) + r] || '');
    grid.push(row);
  }
  return grid;
}

function buildChartDataFromCells(chartDef, cells)
{
  var grid = getChartRangeGridFromCells(chartDef.range, cells);
  if (!grid || !grid.length) return { labels: [], datasets: [] };
  if (chartDef.type === 'scatter') return buildScatterChartData(chartDef, grid);
  var data = buildCategoricalChartData(chartDef, grid);
  if (chartDef.type === 'pie' || chartDef.type === 'doughnut') data.datasets = data.datasets.slice(0, 1);
  if (chartDef.type === 'percent')
    data = percentStackDatasets(data, chartDef);
  return data;
}

function buildChartJsConfigFromCells(chartDef, cells)
{
  var data = buildChartDataFromCells(chartDef, cells),
      isPie = (chartDef.type === 'pie' || chartDef.type === 'doughnut'),
      baseType = (chartDef.type === 'combo' || chartDef.type === 'percent') ? 'bar' : chartDef.type;

  var datasets = data.datasets.map(function(ds, i)
  {
    var color = CHART_COLORS[i % CHART_COLORS.length],
        styled =
        {
          label: ds.label,
          data: ds.data,
          backgroundColor: isPie ? ds.data.map(function(_, j){ return CHART_COLORS[j % CHART_COLORS.length]; }) : color,
          borderColor: color,
          borderWidth: (chartDef.type === 'line' || chartDef.type === 'radar') ? 2 : 1
        };

    if (chartDef.type === 'combo') styled.type = (i === 0) ? 'bar' : 'line';

    if (chartDef.type === 'line' || chartDef.type === 'radar' || (chartDef.type === 'combo' && i > 0))
    {
      styled.fill = false;
      styled.tension = 0.25;
    }

    return styled;
  });

  return  {
            type: baseType,
            data: { labels: data.labels, datasets: datasets },
            options:
            {
              responsive: true,
              maintainAspectRatio: false,
              animation: false,
              plugins:
              {
                title: { display: !!chartDef.title, text: chartDef.title, color: '#f0ede6' },
                legend: { display: datasets.length > 1 || isPie, labels: { color: '#9e9b94' } }
              },
              scales: isPie ? {} : buildChartScales(chartDef)
            }
          };
}

function renderEmbeddedSheetCharts(container)
{
  if (typeof Chart === 'undefined') return;

  var embeds = container.querySelectorAll('.embed-sheet-chart[data-chart-config]');

  embeds.forEach(function(div)
  {
    try
    {
      var config = JSON.parse(div.getAttribute('data-chart-config')),
          canvas = div.querySelector('canvas');

      if (canvas) new Chart(canvas.getContext('2d'), config);
    }
    catch(e) { console.warn('Embedded chart render error', e); }
  });
}

document.getElementById('preview-content').addEventListener
(
  'click',
  function(e)
  {
    var link = e.target.closest('a');

    if (!link)
      return;

    var href = link.getAttribute('href') || '';

    if (href.indexOf(FILE_LINK_SCHEME) !== 0)
      return;

    e.preventDefault();

    var id = decodeURIComponent(href.slice(FILE_LINK_SCHEME.length));

    if (!files[id])
    {
      alert('That file no longer exists, or was renamed/moved.');
      return;
    }

    openFile(id);
  }
);

function toggleDocCheckbox(idx)
{
  var editor = document.getElementById('editor'),
      count = 0,
      newText = editor.value.replace(/^(\s*[-*+]\s+)\[([ xX])\]/gm, function(m, prefix, state)
      {
        if (count === idx)
        {
          count++;
          return prefix + '[' + (state.trim() === '' ? 'x' : ' ') + ']';
        }
        count++;
        return m;
      });

  if (newText !== editor.value)
  {
    docTextUndo.forceSnapshot();
    editor.value = newText;
    onEditorChange();
  }
}

document.getElementById('preview-content').addEventListener
(
  'change',
  function(e)
  {
    if (!e.target.classList.contains('preview-checkbox'))
      return;
    toggleDocCheckbox(parseInt(e.target.getAttribute('data-idx'), 10));
  }
);

// Alignment blocks are written with a bare symbol (no keyword) repeated on
// the opening and closing line - :: for left, ::: for right, ;; for center,
// ;;; for justify. Anything else (currently just font="...") still uses the
// older ":::keyword ... :::" form, unchanged.
var ALIGN_BLOCK_SYMBOLS = { left: '::', right: ':::', center: ';;', justify: ';;;' };

function wrapBlock(align)
{
  docTextUndo.forceSnapshot();

  var editor = document.getElementById('editor'),
      selectionStart = editor.selectionStart,
      selectionEnd = editor.selectionEnd,
      selection = editor.value.slice(selectionStart, selectionEnd) || 'Text',
      symbol = ALIGN_BLOCK_SYMBOLS[align];

  var block = '\n' + symbol + '\n' + selection + '\n' + symbol + '\n';

  editor.value = editor.value.slice(0, selectionStart) + block + editor.value.slice(selectionEnd);
  editor.selectionStart = editor.selectionEnd = selectionStart + block.length;

  editor.focus();
  onEditorChange();
}

// Font blocks use their own bare-symbol scheme (see fontSymbolBlock above) -
// || followed directly by the font name, no "font=" keyword, no quotes,
// closed by a bare ||.
function wrapFontBlock(fontName)
{
  docTextUndo.forceSnapshot();

  var editor = document.getElementById('editor'),
      selectionStart = editor.selectionStart,
      selectionEnd = editor.selectionEnd,
      selection = editor.value.slice(selectionStart, selectionEnd) || 'Text';

  var block = '\n||' + fontName + '\n' + selection + '\n||\n';

  editor.value = editor.value.slice(0, selectionStart) + block + editor.value.slice(selectionEnd);
  editor.selectionStart = editor.selectionEnd = selectionStart + block.length;

  editor.focus();
  onEditorChange();
}

function insertTOC()
{
  docTextUndo.forceSnapshot();

  var editor = document.getElementById('editor'),
      selectionStart = editor.selectionStart,
      lineStart = editor.value.lastIndexOf('\n', selectionStart - 1) + 1,
      marker = '[TOC]\n\n';

  editor.value = editor.value.slice(0, lineStart) + marker + editor.value.slice(lineStart);
  editor.selectionStart = editor.selectionEnd = lineStart + marker.length;

  editor.focus();
  onEditorChange();
}

// ── FONT PICKER ──
// Prefers the real fonts installed on the user's machine via the Local Font
// Access API (window.queryLocalFonts - Chromium/WebView2, requires a user
// gesture and a one-time permission grant), falling back to a static list
// of the fonts that ship with Windows when that API is unavailable, denied,
// or fails for any other reason.
var FALLBACK_FONT_LIST =
[
  'Arial', 'Arial Black', 'Bahnschrift', 'Calibri', 'Cambria', 'Candara', 'Comic Sans MS',
  'Consolas', 'Constantia', 'Corbel', 'Courier New', 'Ebrima', 'Franklin Gothic Medium',
  'Gabriola', 'Gadugi', 'Georgia', 'Impact', 'Ink Free', 'Javanese Text', 'Leelawadee UI',
  'Lucida Console', 'Lucida Sans Unicode', 'Malgun Gothic', 'Microsoft Himalaya',
  'Microsoft JhengHei', 'Microsoft New Tai Lue', 'Microsoft PhagsPa', 'Microsoft Sans Serif',
  'Microsoft Tai Le', 'Microsoft Yahei', 'Microsoft Yi Baiti', 'MingLiU-ExtB', 'Mongolian Baiti',
  'MS Gothic', 'MV Boli', 'Myanmar Text', 'Nirmala UI', 'Palatino Linotype', 'Segoe Print',
  'Segoe Script', 'Segoe UI', 'Segoe UI Historic', 'SimSun', 'Sitka', 'Sylfaen', 'Tahoma',
  'Times New Roman', 'Trebuchet MS', 'Verdana', 'Yu Gothic',
  'serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui'
];

var cachedFontList = null,
    fontPickerSelection = null;

async function getAvailableFonts()
{
  if (cachedFontList)
    return cachedFontList;

  var names = null;

  // queryLocalFonts() only exists in a secure context (HTTPS, or literally
  // "localhost") in Chromium browsers, so on a plain-http deployment - or
  // any non-Chromium browser - this silently has nothing to offer; the
  // account's uploaded fonts (merged in below) are the reliable path that
  // works the same way everywhere.
  if (typeof window.queryLocalFonts === 'function')
  {
    try
    {
      var fonts = await window.queryLocalFonts(),
          seen = {};

      names = [];

      fonts.forEach(function(f)
      {
        if (!seen[f.family])
        {
          seen[f.family] = true;
          names.push(f.family);
        }
      });

      names.sort(function(a, b){ return a.localeCompare(b); });
    }
    catch(e)
    {
      console.warn('Local Font Access unavailable, falling back to the standard font list', e);
      names = null;
    }
  }

  var baseNames = (names && names.length) ? names : FALLBACK_FONT_LIST.slice();

  if (customFontsCache === null) await refreshCustomFonts();
  var customNames = (customFontsCache || []).map(function(f){ return f.familyName; });

  // Custom (uploaded, always-renders-correctly) fonts first, then whatever
  // the OS/fallback list contributed, deduped.
  var seenNames = {}, merged = [];
  customNames.concat(baseNames).forEach(function(n)
  {
    var key = n.toLowerCase();
    if (!seenNames[key]) { seenNames[key] = true; merged.push(n); }
  });

  cachedFontList = merged;
  return cachedFontList;
}

// Changes the font of the current selection, akin to LaTeX's \fontspec/\setfont commands.
// A selection spanning multiple paragraphs becomes a ":::" block (own font for the whole passage);
// a single-line/inline selection becomes a "[text]{font="..."}" span (own font for that phrase only).
async function applyFont()
{
  var editor = document.getElementById('editor');

  fontPickerSelection = { selectionStart: editor.selectionStart, selectionEnd: editor.selectionEnd };

  document.getElementById('font-picker-search').value = '';
  document.getElementById('font-picker-list').innerHTML = '<div class="history-empty">Loading fonts…</div>';
  document.getElementById('font-picker-modal').classList.add('open');
  document.getElementById('font-picker-search').focus();

  fontPickerSelection.fonts = await getAvailableFonts();
  renderFontPickerList();
}

function closeFontPickerModal()
{
  document.getElementById('font-picker-modal').classList.remove('open');
}

document.getElementById('font-picker-modal').addEventListener
(
  'click',
  function(e)
  {
    if (e.target === document.getElementById('font-picker-modal'))
      closeFontPickerModal();
  }
);

function renderFontPickerList()
{
  var query = document.getElementById('font-picker-search').value.trim().toLowerCase(),
      list = document.getElementById('font-picker-list'),
      fonts = (fontPickerSelection && fontPickerSelection.fonts) || [],
      matches = fonts.filter(function(f){ return !query || f.toLowerCase().indexOf(query) !== -1; });

  if (!matches.length)
  {
    list.innerHTML = '<div class="history-empty">No matching fonts.</div>';
    return;
  }

  list.innerHTML = matches.map(function(f)
  {
    return  '<div class="font-picker-item" style="font-family:' + escHtml(cssFontFamilyValue(f)) + ';" onclick="chooseFontFromPicker(\'' + escAttr(f) + '\')">' +
              escHtml(f) +
            '</div>';
  }).join('');
}

function chooseFontFromPicker(fontName)
{
  closeFontPickerModal();

  var editor = document.getElementById('editor'),
      selectionStart = fontPickerSelection.selectionStart,
      selectionEnd = fontPickerSelection.selectionEnd,
      selection = editor.value.slice(selectionStart, selectionEnd);

  editor.selectionStart = selectionStart;
  editor.selectionEnd = selectionEnd;

  if (selection && !/\n[ \t]*\n/.test(selection))
    wrapSelectedText('[', ']{font="' + fontName + '"}');

  else
    wrapFontBlock(fontName);
}

function setView(newView)
{
  editorView = newView;

  document.getElementById('editor-area').className = (newView === 'write')
                                                      ?
                                                        ''
                                                      :
                                                        newView;

  ['write', 'split', 'preview'].forEach(
    function(view)
    {
      document.getElementById('vbtn-'+ view).classList.toggle('active', view === newView);
    }
  );

  if (newView !== 'write')
    updatePreview();
}

document.getElementById('editor').addEventListener('dblclick', function()
{
  if (editorView !== 'split') return;
  var text = this.value,
      cursorLine = text.slice(0, this.selectionStart).split('\n').length,
      totalLines = text.split('\n').length || 1,
      pp = document.getElementById('preview-pane');
  pp.scrollTop = (cursorLine / totalLines) * (pp.scrollHeight - pp.clientHeight);
});

document.getElementById('preview-pane').addEventListener('dblclick', function(e)
{
  if (editorView !== 'split') return;
  var pp = this,
      ratio = (pp.scrollTop + e.clientY - pp.getBoundingClientRect().top) / pp.scrollHeight,
      editor = document.getElementById('editor');
  editor.scrollTop = ratio * (editor.scrollHeight - editor.clientHeight);
});

// ── GRAPH (Mermaid diagrams, stored as Markdown) ──
var graphTextUndo = makeTextUndo(function(){ return document.getElementById('graph-editor'); });

function loadGraphFile(file)
{
  graphTextUndo.reset();
  document.getElementById('graph-title-input').value = file.name;
  document.getElementById('graph-editor').value = file.content || '';

  updateGraphPreview();
}

function handleGraphEditorKey(keyEvent)
{
  graphTextUndo.noteKeydown(keyEvent);

  if (!(keyEvent.ctrlKey || keyEvent.metaKey))
    return;

  if (!keyEvent.shiftKey && keyEvent.key.toLowerCase() === 'z')
  {
    keyEvent.preventDefault();

    if (graphTextUndo.undo())
      onGraphEditorChange();
  }
  else if (keyEvent.key.toLowerCase() === 'y' || (keyEvent.shiftKey && keyEvent.key.toLowerCase() === 'z'))
  {
    keyEvent.preventDefault();

    if (graphTextUndo.redo())
      onGraphEditorChange();
  }
}

function onGraphTitleChange()
{
  if ((!currentFileId) || (files[currentFileId].type !== 'graph'))
    return;

  files[currentFileId].name = (document.getElementById('graph-title-input').value || files[currentFileId].name);
  scheduleSave();
  renderFileList();
}

function onGraphEditorChange()
{
  if (!currentFileId || !files[currentFileId] || files[currentFileId].type !== 'graph')
    return;

  files[currentFileId].content = document.getElementById('graph-editor').value;

  updateGraphPreview();
  scheduleSave();
}

function updateGraphPreview()
{
  var source = document.getElementById('graph-editor').value,
      html = escHtml(source);

  if (typeof marked !== 'undefined')
    html = marked.parser(marked.lexer(source));

  var container = document.getElementById('graph-preview-content');

  container.innerHTML = html;

  renderMermaidDiagrams(container);
}

// Renders any `.mermaid` placeholder divs produced by the marked code-block
// override into actual diagrams. Re-running on freshly-built markup is safe
// since each call works on brand-new, unprocessed nodes.
function renderMermaidDiagrams(container)
{
  if ((typeof mermaid === 'undefined') || !container)
    return;

  var nodes = container.querySelectorAll('.mermaid');

  if (!nodes.length)
    return;

  try
  {
    mermaid.run({ nodes: nodes, suppressErrors: true });
  }
  catch(e)
  {
    console.warn('Mermaid render error', e);
  }
}

function setGraphView(newView)
{
  graphEditorView = newView;

  document.getElementById('graph-editor-area').className = (newView === 'write')
                                                            ?
                                                              ''
                                                            :
                                                              newView;

  ['write', 'split', 'preview'].forEach(
    function(view)
    {
      document.getElementById('gvbtn-'+ view).classList.toggle('active', view === newView);
    }
  );

  if (newView !== 'write')
    updateGraphPreview();
}

var MERMAID_SNIPPETS =
[
  {
    name: 'Flowchart',
    desc: 'Boxes and arrows for a process or decision.',
    snippet: '```mermaid\ngraph TD\n    A[Start] --> B{Decision}\n    B -->|Yes| C[Continue]\n    B -->|No| D[Stop]\n```'
  },
  {
    name: 'Sequence diagram',
    desc: 'Messages exchanged between participants over time.',
    snippet: '```mermaid\nsequenceDiagram\n    Alice->>Bob: Hello Bob, how are you?\n    Bob-->>Alice: I am good, thanks!\n```'
  },
  {
    name: 'Class diagram',
    desc: 'Classes, members and relationships.',
    snippet: '```mermaid\nclassDiagram\n    Animal <|-- Dog\n    Animal : +String name\n    Animal : +makeSound()\n```'
  },
  {
    name: 'State diagram',
    desc: 'States and the transitions between them.',
    snippet: '```mermaid\nstateDiagram-v2\n    [*] --> Idle\n    Idle --> Running : start\n    Running --> Idle : stop\n    Running --> [*]\n```'
  },
  {
    name: 'Entity relationship',
    desc: 'Entities and how they relate to each other.',
    snippet: '```mermaid\nerDiagram\n    CUSTOMER ||--o{ ORDER : places\n    ORDER ||--|{ LINE-ITEM : contains\n```'
  },
  {
    name: 'Gantt chart',
    desc: 'Tasks laid out against a timeline.',
    snippet: '```mermaid\ngantt\n    title Project Plan\n    section Phase 1\n    Design : 2024-01-01, 5d\n    Build  : 5d\n```'
  },
  {
    name: 'Pie chart',
    desc: 'Proportions of a whole.',
    snippet: '```mermaid\npie title Distribution\n    "A" : 40\n    "B" : 35\n    "C" : 25\n```'
  },
  {
    name: 'Relationship map',
    desc: 'Characters, factions and places connected by relationships.',
    snippet: '```mermaid\ngraph LR\n    %% Characters\n    Aria((Aria))\n    Kael((Kael))\n\n    %% Factions\n    Order[Order of Dawn]\n    Shadow[Shadow Guild]\n\n    %% Places\n    City([Ironhaven])\n\n    Aria -->|member of| Order\n    Kael -->|member of| Shadow\n    Aria ---|rivals| Kael\n    Order ---|based in| City\n\n    classDef character fill:#4a7fa8,stroke:#3a6f98,color:#fff\n    classDef faction fill:#8b6f3e,stroke:#7b5f2e,color:#fff\n    classDef place fill:#4a8b60,stroke:#3a7b50,color:#fff\n    class Aria,Kael character\n    class Order,Shadow faction\n    class City place\n```'
  },
  {
    name: 'Timeline',
    desc: 'Horizontal timeline for events, eras and character arcs.',
    snippet: '```mermaid\ntimeline\n    title Chronicle of Aethermoor\n    section Age of Founding\n        Year 1   : Kingdom established\n                 : First Conclave formed\n        Year 50  : Great Schism\n    section Age of Conflict\n        Year 100 : War of Three Crowns begins\n        Year 130 : Peace of Whitehall signed\n    section Current Era\n        Year 200 : Present day\n```'
  },
  {
    name: 'Hierarchy tree',
    desc: 'Org-chart style tree for lineages, power structures or systems.',
    snippet: '```mermaid\ngraph TD\n    A[The High Crown] --> B[Northern Realm]\n    A --> C[Southern Realm]\n    A --> D[Eastern Realm]\n    B --> E[House Ironwood]\n    B --> F[House Frostholm]\n    C --> G[House Sundagger]\n    C --> H[House Ashveil]\n    D --> I[House Embercroft]\n\n    classDef royal fill:#8b6f3e,stroke:#7b5f2e,color:#fff\n    classDef noble fill:#4a7fa8,stroke:#3a6f98,color:#fff\n    class A royal\n    class B,C,D,E,F,G,H,I noble\n```'
  }
];

function renderMermaidHelp()
{
  var panel = document.getElementById('mermaid-help-panel');

  if (!panel)
    return;

  var intro = '<div class="fx-help-intro">'+
                'Diagrams are stored as Markdown with a <code>```mermaid</code> code block. '+
                'Click a diagram type below to insert a starter template at the cursor.'+
              '</div>';

  var rows = MERMAID_SNIPPETS.map
  (
    function(s, i)
    {
      return  '<div class="fx-help-row mermaid-help-row" onclick="insertMermaidSnippet(' + i + ')">'+
                '<div class="fx-help-name">' + escHtml(s.name) + '</div>'+
                '<div class="fx-help-desc">' + escHtml(s.desc) + '</div>'+
              '</div>';
    }
  ).join('');

  panel.innerHTML = intro + rows;
}

function toggleMermaidHelp(clickEvent)
{
  if (clickEvent)
    clickEvent.stopPropagation();

  var panel = document.getElementById('mermaid-help-panel');

  if (!panel)
    return;

  if (!panel.classList.contains('open'))
    renderMermaidHelp();

  panel.classList.toggle('open');
}

document.addEventListener
(
  'click',
  function(clickEvent)
  {
    var panel = document.getElementById('mermaid-help-panel');

    if (!panel || !panel.classList.contains('open'))
      return;

    if (panel.contains(clickEvent.target) || clickEvent.target.id === 'mermaid-help-btn')
      return;

    panel.classList.remove('open');
  }
);

function insertMermaidSnippet(index)
{
  var snippet = MERMAID_SNIPPETS[index];

  if (!snippet)
    return;

  graphTextUndo.forceSnapshot();

  var editor = document.getElementById('graph-editor'),
      start = editor.selectionStart,
      end = editor.selectionEnd,
      before = editor.value.slice(0, start),
      after = editor.value.slice(end),
      lead = (!before || /\n\n$/.test(before)) ? '' : (/\n$/.test(before) ? '\n' : '\n\n'),
      trail = (!after || /^\n/.test(after)) ? '' : '\n\n';

  editor.value = before + lead + snippet.snippet + trail + after;

  var caretPos = (before + lead + snippet.snippet).length;
  editor.setSelectionRange(caretPos, caretPos);
  editor.focus();

  onGraphEditorChange();

  document.getElementById('mermaid-help-panel').classList.remove('open');
}

// ── NOTEBOOK (Xournal++-style pen/highlighter/eraser/shapes/text/PDF annotation) ──

// Transient editing state - not persisted; reset whenever a notebook file is opened.
let nbData = null;
let nbTool = 'pen';
let nbColor = '#1a1a1a';
let nbWidth = 3;
let nbCurrentPage = 0;
let nbUndoStack = [];
let nbRedoStack = [];
let nbSelection = [];
let nbActive = null;
let nbPdfDocCache = null;
let nbPdfDocCacheKey = null;
let nbPageObserver = null;
let nbZoom = 1;

// Pending SVG vector import awaiting interactive placement: strokes are held
// normalized (bounds shifted to origin, unscaled) and only baked into the
// page's stroke list when the user accepts the position/size.
let nbImport = null;

// Per-page offscreen bitmap of every *committed* stroke/shape (i.e. everything
// except whatever's currently being drawn), keyed by page index. Letting an
// active pen/highlighter/shape draw on top of a cached blit instead of
// replaying the page's entire history on every pointermove is what keeps fast
// stylus input feeling responsive once a page has built up real content -
// redrawing everything every move event scales with total ink on the page,
// which is the opposite of what you want from a 100+Hz pointer stream.
let nbCommittedCache = {};

const NB_COLORS = ['#1a1a1a', '#c0574a', '#2f6fd1', '#2f9e44', '#e8830f', '#7048b8'];

// The ink layer is rendered at least 2x even on a plain 1x display - canvas
// path strokes are always antialiased by the browser, but only against
// whatever resolution the buffer actually has, so deliberately oversampling
// (supersampling) gives visibly smoother edges than native resolution alone,
// the same idea as MSAA. devicePixelRatio is still respected on top of that
// floor, so a HiDPI display gets whichever of the two is sharper.
function notebookCanvasScale()
{
  return Math.max(window.devicePixelRatio || 1, 2);
}

function rebuildNotebookCommittedCache(pageIndex)
{
  const page = nbData && nbData.pages[pageIndex];

  if (!page)
    return;

  const scale = notebookCanvasScale(),
        rawWidth = Math.round(page.width * scale),
        rawHeight = Math.round(page.height * scale);

  let cached = nbCommittedCache[pageIndex];

  if (!cached || cached.width !== rawWidth || cached.height !== rawHeight)
  {
    cached = document.createElement('canvas');
    cached.width = rawWidth;
    cached.height = rawHeight;
    nbCommittedCache[pageIndex] = cached;
  }

  const ctx = cached.getContext('2d');

  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.clearRect(0, 0, rawWidth, rawHeight);
  ctx.scale(scale, scale);

  page.strokes.forEach(function(s){ drawNotebookStroke(ctx, s); });
  page.shapes.forEach(function(s){ drawNotebookShape(ctx, s); });
}

function defaultNotebookPage()
{
  return {
    width: 816,
    height: 1056,
    background: { type: 'ruled' },
    strokes: [],
    shapes: [],
    texts: []
  };
}

function defaultNotebookData()
{
  return { version: 1, pdfData: null, pages: [defaultNotebookPage()] };
}

function serializeNotebookData(data)
{
  return '<!-- type: notebook -->\n\n```notebook\n' + JSON.stringify(data) + '\n```\n';
}

function parseNotebookContent(content)
{
  const match = /```notebook\n([\s\S]*?)\n```/.exec(content || '');

  if (!match)
    return defaultNotebookData();

  try
  {
    const data = JSON.parse(match[1]);

    if (!data.pages || !data.pages.length)
      data.pages = [defaultNotebookPage()];

    return data;
  }
  catch(e)
  {
    console.warn('Notebook parse error', e);
    return defaultNotebookData();
  }
}

function loadNotebookFile(file)
{
  document.getElementById('notebook-title-input').value = file.name;

  nbData = parseNotebookContent(file.content || '');
  nbCurrentPage = 0;
  nbUndoStack = [];
  nbRedoStack = [];
  nbSelection = [];
  nbActive = null;
  nbImport = null;
  document.getElementById('nb-import-bar').style.display = 'none';
  nbPdfDocCache = null;
  nbPdfDocCacheKey = null;
  nbZoom = 1;

  renderNotebookColorSwatches();
  renderNotebookPages();
  updateNotebookZoomIndicator();
  updateMapToolbarVisibility();
}

function onNotebookTitleChange()
{
  if ((!currentFileId) || (files[currentFileId].type !== 'notebook'))
    return;

  files[currentFileId].name = (document.getElementById('notebook-title-input').value || files[currentFileId].name);
  scheduleSave();
  renderFileList();
}

function commitNotebookChange()
{
  if (!currentFileId || !nbData || !files[currentFileId] || files[currentFileId].type !== 'notebook')
    return;

  files[currentFileId].content = serializeNotebookData(nbData);
  scheduleSave();
}

// ── Tool / color / width selection ──

function setNotebookTool(tool)
{
  nbTool = tool;
  nbSelection = [];

  document.querySelectorAll('.nb-tool-btn').forEach(function(btn)
  {
    btn.classList.toggle('active', btn.dataset.tool === tool);
  });

  const pages = document.getElementById('notebook-pages');

  if (pages)
  {
    pages.className = pages.className.replace(/\btool-\S+/g, '').trim();
    pages.classList.add('tool-' + tool);
  }

  if (nbData)
    redrawNotebookPage(nbCurrentPage);
}

function setNotebookColor(color)
{
  nbColor = color;

  document.querySelectorAll('.nb-swatch').forEach(function(sw)
  {
    sw.classList.toggle('active', sw.dataset.color === color);
  });
}

function setNotebookWidth(width)
{
  nbWidth = parseFloat(width);
}

function renderNotebookColorSwatches()
{
  const group = document.getElementById('notebook-color-group');

  if (!group)
    return;

  group.innerHTML = NB_COLORS.map(function(c)
  {
    return  '<button type="button" class="nb-swatch' + (c === nbColor ? ' active' : '') + '"' +
              ' data-color="' + c + '" style="background:' + c + '" onclick="setNotebookColor(\'' + c + '\')" title="' + c + '"></button>';
  }).join('');
}

// ── Page rendering ──

function renderNotebookPages()
{
  const container = document.getElementById('notebook-pages');

  container.innerHTML = '';

  // Page indices can shift (a page added/removed/reordered), so the whole
  // cache is rebuilt fresh below rather than trying to patch it in place.
  nbCommittedCache = {};

  nbData.pages.forEach(function(page, i)
  {
    const pageEl = document.createElement('div');
    pageEl.className = 'notebook-page' + (page.type === 'map' ? ' map-notebook-page' : '');
    pageEl.style.width = (page.width * nbZoom) + 'px';
    pageEl.style.height = (page.height * nbZoom) + 'px';
    pageEl.dataset.pageIndex = i;

    // Canvases are sized well above CSS resolution so strokes render crisp
    // rather than blurry, on both HiDPI displays and plain 1x ones (see
    // notebookCanvasScale). The ink canvas's context is scaled once here and
    // never touched again, so every other ink-canvas function keeps drawing
    // in plain page-space coordinates (0..page.width/height) regardless of
    // the actual buffer size.
    const scale = notebookCanvasScale();

    const bgCanvas = document.createElement('canvas');
    bgCanvas.className = 'notebook-bg-canvas';
    bgCanvas.width = Math.round(page.width * scale);
    bgCanvas.height = Math.round(page.height * scale);

    const inkCanvas = document.createElement('canvas');
    inkCanvas.className = 'notebook-ink-canvas';
    inkCanvas.width = Math.round(page.width * scale);
    inkCanvas.height = Math.round(page.height * scale);
    inkCanvas.getContext('2d').scale(scale, scale);

    const textLayer = document.createElement('div');
    textLayer.className = 'notebook-text-layer';

    pageEl.appendChild(bgCanvas);
    pageEl.appendChild(inkCanvas);
    pageEl.appendChild(textLayer);

    if (page.type === 'map')
    {
      // Scale/travel overlay canvases (not scaled by context, raw pixel match)
      const scaleCanvas = document.createElement('canvas');
      scaleCanvas.className = 'nb-scale-canvas';
      scaleCanvas.width = Math.round(page.width * scale);
      scaleCanvas.height = Math.round(page.height * scale);
      pageEl.appendChild(scaleCanvas);

      const travelCanvas = document.createElement('canvas');
      travelCanvas.className = 'nb-travel-canvas';
      travelCanvas.width = Math.round(page.width * scale);
      travelCanvas.height = Math.round(page.height * scale);
      pageEl.appendChild(travelCanvas);

      const pinLayer = document.createElement('div');
      pinLayer.className = 'nb-pin-layer';
      pageEl.appendChild(pinLayer);
    }

    container.appendChild(pageEl);

    renderNotebookPageBackground(page, bgCanvas);
    wireNotebookPointerEvents(inkCanvas, i);
    rebuildNotebookCommittedCache(i);
    redrawNotebookPage(i);
    renderNotebookTextLayerForPage(i);

    if (page.type === 'map')
      renderMapPinLayer(i);
  });

  updateNotebookPageIndicator();
  setupNotebookPageObserver();
  updateMapToolbarVisibility();
}

function notebookSetZoom(zoom)
{
  nbZoom = Math.max(0.25, Math.min(3, zoom));

  document.querySelectorAll('.notebook-page').forEach(function(pageEl)
  {
    const i = parseInt(pageEl.dataset.pageIndex, 10),
          page = nbData.pages[i];

    pageEl.style.width = (page.width * nbZoom) + 'px';
    pageEl.style.height = (page.height * nbZoom) + 'px';

    if (page.type === 'map') renderMapPinLayer(i);
  });

  updateNotebookZoomIndicator();
}

function notebookZoomIn()
{
  notebookSetZoom(nbZoom + 0.1);
}

function notebookZoomOut()
{
  notebookSetZoom(nbZoom - 0.1);
}

function notebookZoomReset()
{
  notebookSetZoom(1);
}

function updateNotebookZoomIndicator()
{
  const el = document.getElementById('notebook-zoom-indicator');

  if (el)
    el.textContent = Math.round(nbZoom * 100) + '%';
}

function renderNotebookPageBackground(page, canvas)
{
  const ctx = canvas.getContext('2d');

  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);

  if (page.type === 'map' && page.imageData)
  {
    var mapImg = new Image();
    mapImg.onload = function()
    {
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(mapImg, 0, 0, canvas.width, canvas.height);
    };
    mapImg.src = page.imageData;
    return;
  }

  if (page.background.type === 'pdf')
  {
    if (nbData.pdfData)
      renderNotebookPdfBackground(page, canvas);

    return;
  }

  if (page.background.type === 'ruled' || page.background.type === 'grid')
  {
    // Scoped to just this section (save/restore) rather than scaling the whole
    // context, since the PDF branch above computes its own scale to fill the
    // raw buffer directly - a persistent transform here would double-scale it.
    const scale = notebookCanvasScale();

    ctx.save();
    ctx.scale(scale, scale);

    if (page.background.type === 'ruled')
    {
      ctx.strokeStyle = '#cdd6e0';
      ctx.lineWidth = 1;

      for (let y = 40; y < page.height; y += 32)
      {
        ctx.beginPath();
        ctx.moveTo(0, y + 0.5);
        ctx.lineTo(page.width, y + 0.5);
        ctx.stroke();
      }
    }
    else
    {
      ctx.strokeStyle = '#dde3ea';
      ctx.lineWidth = 1;

      for (let gx = 0; gx < page.width; gx += 24)
      {
        ctx.beginPath();
        ctx.moveTo(gx + 0.5, 0);
        ctx.lineTo(gx + 0.5, page.height);
        ctx.stroke();
      }

      for (let gy = 0; gy < page.height; gy += 24)
      {
        ctx.beginPath();
        ctx.moveTo(0, gy + 0.5);
        ctx.lineTo(page.width, gy + 0.5);
        ctx.stroke();
      }
    }

    ctx.restore();
  }
}

async function getNotebookPdfDocument()
{
  if (nbPdfDocCache && nbPdfDocCacheKey === nbData.pdfData)
    return nbPdfDocCache;

  const base64 = nbData.pdfData.split(',').pop(),
        binary = atob(base64),
        bytes = new Uint8Array(binary.length);

  for (let i = 0; i < binary.length; i++)
    bytes[i] = binary.charCodeAt(i);

  nbPdfDocCache = await pdfjsLib.getDocument({ data: bytes }).promise;
  nbPdfDocCacheKey = nbData.pdfData;

  return nbPdfDocCache;
}

async function renderNotebookPdfBackground(page, canvas)
{
  if (typeof pdfjsLib === 'undefined')
    return;

  try
  {
    const pdfDoc = await getNotebookPdfDocument(),
          pdfPage = await pdfDoc.getPage(page.background.pdfPageIndex + 1),
          scale = canvas.width / pdfPage.getViewport({ scale: 1 }).width,
          viewport = pdfPage.getViewport({ scale: scale });

    await pdfPage.render({ canvasContext: canvas.getContext('2d'), viewport: viewport }).promise;
  }
  catch(e)
  {
    console.warn('Notebook PDF background render error', e);
  }
}

function drawNotebookStroke(ctx, stroke, alphaScale)
{
  const pts = stroke.points;

  ctx.globalAlpha = ((stroke.tool === 'highlighter') ? 0.35 : 1) * (alphaScale || 1);
  ctx.strokeStyle = stroke.color;
  ctx.fillStyle = stroke.color;
  ctx.lineCap = 'round';
  ctx.lineJoin = 'round';

  if (stroke.fill)
  {
    // Filled stroke (SVG import): the points trace the shape's outline(s),
    // with `breaks` marking indices where a new subpath starts. Nonzero
    // winding (the SVG default) keeps self-intersecting shapes solid -
    // Excalidraw freedraw ribbons overlap themselves constantly - while
    // still leaving holes (letter counters, rings) empty via subpath
    // winding direction; evenodd is honored when the source asked for it.
    const breaks = stroke.breaks || [];
    let bi = 0;

    ctx.beginPath();

    for (let i = 0; i < pts.length; i++)
    {
      if (i === 0)
      {
        ctx.moveTo(pts[i].x, pts[i].y);
      }
      else if (bi < breaks.length && breaks[bi] === i)
      {
        ctx.closePath();
        ctx.moveTo(pts[i].x, pts[i].y);
        bi++;
      }
      else
      {
        ctx.lineTo(pts[i].x, pts[i].y);
      }
    }

    ctx.closePath();
    ctx.fill(stroke.fillRule === 'evenodd' ? 'evenodd' : 'nonzero');
    ctx.globalAlpha = 1;
    return;
  }

  if (pts.length === 1)
  {
    ctx.beginPath();
    ctx.arc(pts[0].x, pts[0].y, (pts[0].w || stroke.baseWidth) / 2, 0, Math.PI * 2);
    ctx.fill();
  }
  else if (pts.length === 2)
  {
    ctx.beginPath();
    ctx.lineWidth = ((pts[0].w || stroke.baseWidth) + (pts[1].w || stroke.baseWidth)) / 2;
    ctx.moveTo(pts[0].x, pts[0].y);
    ctx.lineTo(pts[1].x, pts[1].y);
    ctx.stroke();
  }
  else
  {
    // Quadratic curve through the midpoint of each consecutive pair, using the
    // shared point as the control point - the standard cheap smoothing trick
    // for freehand input, since raw polylines look visibly faceted/jagged at
    // normal sampling rates rather than smooth handwriting strokes. Chained
    // from the actual first point through each midpoint to the actual last
    // point, so the smoothing doesn't clip the stroke's start/end tips.
    let prevMid = pts[0];

    for (let i = 1; i < pts.length - 1; i++)
    {
      const p1 = pts[i], p2 = pts[i + 1],
            mid = { x: (p1.x + p2.x) / 2, y: (p1.y + p2.y) / 2 };

      ctx.beginPath();
      ctx.lineWidth = p1.w || stroke.baseWidth;
      ctx.moveTo(prevMid.x, prevMid.y);
      ctx.quadraticCurveTo(p1.x, p1.y, mid.x, mid.y);
      ctx.stroke();

      prevMid = mid;
    }

    const last = pts[pts.length - 1];

    ctx.beginPath();
    ctx.lineWidth = last.w || stroke.baseWidth;
    ctx.moveTo(prevMid.x, prevMid.y);
    ctx.lineTo(last.x, last.y);
    ctx.stroke();
  }

  ctx.globalAlpha = 1;
}

function drawNotebookShape(ctx, shape)
{
  ctx.strokeStyle = shape.color;
  ctx.lineWidth = shape.width;
  ctx.lineCap = 'round';
  ctx.beginPath();

  if (shape.tool === 'line')
  {
    ctx.moveTo(shape.x1, shape.y1);
    ctx.lineTo(shape.x2, shape.y2);
  }
  else if (shape.tool === 'rectangle')
  {
    ctx.rect(Math.min(shape.x1, shape.x2), Math.min(shape.y1, shape.y2), Math.abs(shape.x2 - shape.x1), Math.abs(shape.y2 - shape.y1));
  }
  else if (shape.tool === 'ellipse')
  {
    const cx = (shape.x1 + shape.x2) / 2,
          cy = (shape.y1 + shape.y2) / 2,
          rx = Math.abs(shape.x2 - shape.x1) / 2,
          ry = Math.abs(shape.y2 - shape.y1) / 2;

    ctx.ellipse(cx, cy, rx, ry, 0, 0, Math.PI * 2);
  }

  ctx.stroke();
}

function notebookStrokeBounds(stroke)
{
  const xs = stroke.points.map(function(p){ return p.x; }),
        ys = stroke.points.map(function(p){ return p.y; });

  return { x1: Math.min.apply(null, xs), y1: Math.min.apply(null, ys), x2: Math.max.apply(null, xs), y2: Math.max.apply(null, ys) };
}

function notebookShapeBounds(shape)
{
  return { x1: Math.min(shape.x1, shape.x2), y1: Math.min(shape.y1, shape.y2), x2: Math.max(shape.x1, shape.x2), y2: Math.max(shape.y1, shape.y2) };
}

function notebookTextBounds(text)
{
  return { x1: text.x, y1: text.y, x2: text.x + text.width, y2: text.y + text.height };
}

function notebookItemBounds(page, sel)
{
  if (sel.kind === 'stroke')
    return notebookStrokeBounds(page.strokes[sel.index]);

  if (sel.kind === 'shape')
    return notebookShapeBounds(page.shapes[sel.index]);

  if (sel.kind === 'text')
    return notebookTextBounds(page.texts[sel.index]);

  return null;
}

function redrawNotebookPage(pageIndex)
{
  if (!nbData || !nbData.pages[pageIndex])
    return;

  const pageEl = document.querySelector('.notebook-page[data-page-index="' + pageIndex + '"]');

  if (!pageEl)
    return;

  const canvas = pageEl.querySelector('.notebook-ink-canvas'),
        ctx = canvas.getContext('2d'),
        page = nbData.pages[pageIndex];

  // The context carries a persistent devicePixelRatio scale (applied once when
  // the canvas was created), so clears/draws use logical page coordinates -
  // canvas.width/height here would be the larger raw buffer size and over-clear.
  ctx.clearRect(0, 0, page.width, page.height);

  const activeHere = nbActive && nbActive.pageIndex === pageIndex;

  if (activeHere && (nbActive.kind === 'erase' || nbActive.kind === 'move'))
  {
    // These mutate already-committed strokes/shapes in place rather than adding
    // a new one, so the cached bitmap (which reflects the *previous* commit)
    // can't be reused for them - has to be a real redraw from current data.
    page.strokes.forEach(function(s){ drawNotebookStroke(ctx, s); });
    page.shapes.forEach(function(s){ drawNotebookShape(ctx, s); });
  }
  else
  {
    const cached = nbCommittedCache[pageIndex];

    if (cached)
      ctx.drawImage(cached, 0, 0, page.width, page.height);
    else
    {
      page.strokes.forEach(function(s){ drawNotebookStroke(ctx, s); });
      page.shapes.forEach(function(s){ drawNotebookShape(ctx, s); });
    }

    if (activeHere && nbActive.kind === 'stroke')
      drawNotebookStroke(ctx, nbActive.stroke);
    else if (activeHere && nbActive.kind === 'shape')
      drawNotebookShape(ctx, nbActive.shape);
  }

  if (nbImport && nbImport.pageIndex === pageIndex)
  {
    const iw = nbImport.w * nbImport.scale,
          ih = nbImport.h * nbImport.scale;

    // Ghost preview of the pending vector, rendered through the exact same
    // stroke pipeline that committed ink uses, just translated/scaled and
    // slightly transparent until the user accepts the placement.
    ctx.save();
    ctx.translate(nbImport.x, nbImport.y);
    ctx.scale(nbImport.scale, nbImport.scale);
    nbImport.strokes.forEach(function(s){ drawNotebookStroke(ctx, s, 0.75); });
    ctx.restore();

    ctx.save();
    ctx.strokeStyle = '#d4a843';
    ctx.lineWidth = 1.5;
    ctx.setLineDash([6, 4]);
    ctx.strokeRect(nbImport.x - 6, nbImport.y - 6, iw + 12, ih + 12);
    ctx.setLineDash([]);
    ctx.fillStyle = '#d4a843';
    ctx.fillRect(nbImport.x + iw + 1, nbImport.y + ih + 1, 10, 10);
    ctx.restore();
  }

  if (nbActive && nbActive.kind === 'marquee' && nbActive.pageIndex === pageIndex)
  {
    const mx = Math.min(nbActive.x1, nbActive.x2),
          my = Math.min(nbActive.y1, nbActive.y2),
          mw = Math.abs(nbActive.x2 - nbActive.x1),
          mh = Math.abs(nbActive.y2 - nbActive.y1);

    ctx.save();
    ctx.fillStyle = 'rgba(212,168,67,0.12)';
    ctx.strokeStyle = 'rgba(212,168,67,0.8)';
    ctx.lineWidth = 1;
    ctx.fillRect(mx, my, mw, mh);
    ctx.strokeRect(mx, my, mw, mh);
    ctx.restore();
  }

  if (pageIndex === nbCurrentPage)
  {
    nbSelection.forEach(function(sel)
    {
      const b = notebookItemBounds(page, sel);

      if (!b)
        return;

      ctx.save();
      ctx.strokeStyle = '#d4a843';
      ctx.lineWidth = 1.5;
      ctx.setLineDash([5, 4]);
      ctx.strokeRect(b.x1 - 4, b.y1 - 4, (b.x2 - b.x1) + 8, (b.y2 - b.y1) + 8);
      ctx.restore();
    });
  }
}

function updateNotebookPageIndicator()
{
  const el = document.getElementById('notebook-page-indicator');

  if (el && nbData)
    el.textContent = (nbCurrentPage + 1) + ' / ' + nbData.pages.length;
}

// While the user free-scrolls, the observer keeps nbCurrentPage in sync with
// whatever's actually on screen. But programmatic navigation (add/delete/prev/
// next page) sets nbCurrentPage deliberately and then scrolls to match - and an
// observer callback queued from the *previous* scroll position can fire just
// after, stomping the deliberate value before the new scroll position is even
// reflected. Suppressing the observer for a short window after any
// programmatic jump avoids that race.
let nbSuppressObserverUntil = 0;

function setupNotebookPageObserver()
{
  if (nbPageObserver)
    nbPageObserver.disconnect();

  const container = document.getElementById('notebook-pages');

  nbPageObserver = new IntersectionObserver(function(entries)
  {
    if (Date.now() < nbSuppressObserverUntil)
      return;

    let best = null;

    entries.forEach(function(entry)
    {
      if (entry.isIntersecting && (!best || entry.intersectionRatio > best.intersectionRatio))
        best = entry;
    });

    if (best)
    {
      nbCurrentPage = parseInt(best.target.dataset.pageIndex, 10);
      updateNotebookPageIndicator();
    }
  }, { root: container, threshold: [0.5] });

  document.querySelectorAll('.notebook-page').forEach(function(el)
  {
    nbPageObserver.observe(el);
  });
}

function notebookScrollToPage(index)
{
  const pageEl = document.querySelector('.notebook-page[data-page-index="' + index + '"]');

  if (pageEl)
    pageEl.scrollIntoView({ behavior: 'auto', block: 'start' });
}

function notebookGoToPage(index)
{
  nbCurrentPage = index;
  nbSuppressObserverUntil = Date.now() + 400;
  updateNotebookPageIndicator();
  notebookScrollToPage(index);
  updateMapToolbarVisibility();
}

function notebookPrevPage()
{
  if (!nbData || nbCurrentPage <= 0)
    return;

  notebookGoToPage(nbCurrentPage - 1);
}

function notebookNextPage()
{
  if (!nbData || nbCurrentPage >= nbData.pages.length - 1)
    return;

  notebookGoToPage(nbCurrentPage + 1);
}

function notebookAddPage()
{
  if (!nbData)
    return;

  const lastPage = nbData.pages[nbData.pages.length - 1],
        newPage =
        {
          width: lastPage.width,
          height: lastPage.height,
          background: { type: (lastPage.background.type === 'pdf') ? 'blank' : lastPage.background.type },
          strokes: [],
          shapes: [],
          texts: []
        };

  nbData.pages.push(newPage);

  renderNotebookPages();
  commitNotebookChange();
  notebookGoToPage(nbData.pages.length - 1);
}

function notebookDeletePage()
{
  if (!nbData)
    return;

  if (nbData.pages.length <= 1)
  {
    alert('A notebook needs at least one page.');
    return;
  }

  if (!confirm('Delete this page?'))
    return;

  const targetPage = nbCurrentPage;

  nbData.pages.splice(targetPage, 1);
  nbSelection = [];

  // Page indices shift under a pending vector import - drop it rather than
  // letting it land on the wrong page.
  if (nbImport)
    cancelNotebookVectorImport();

  renderNotebookPages();
  commitNotebookChange();
  notebookGoToPage(Math.min(targetPage, nbData.pages.length - 1));
}

// ── Undo / redo ──

function notebookSnapshotForUndo(pageIndex)
{
  const page = nbData.pages[pageIndex];

  var snapFields = { strokes: page.strokes, shapes: page.shapes, texts: page.texts };
  if (page.type === 'map') snapFields.pins = page.pins || [];

  nbUndoStack.push({ pageIndex: pageIndex, snapshot: JSON.parse(JSON.stringify(snapFields)) });
  nbRedoStack = [];

  if (nbUndoStack.length > 50)
    nbUndoStack.shift();
}

function notebookRestoreSnapshot(entry, stack)
{
  const page = nbData.pages[entry.pageIndex],
        inverse = JSON.parse(JSON.stringify({ strokes: page.strokes, shapes: page.shapes, texts: page.texts }));

  stack.push({ pageIndex: entry.pageIndex, snapshot: inverse });

  page.strokes = entry.snapshot.strokes;
  page.shapes = entry.snapshot.shapes;
  page.texts = entry.snapshot.texts;

  nbSelection = [];
  nbCurrentPage = entry.pageIndex;

  rebuildNotebookCommittedCache(entry.pageIndex);
  renderNotebookTextLayerForPage(entry.pageIndex);
  redrawNotebookPage(entry.pageIndex);
  commitNotebookChange();
}

function notebookUndo()
{
  if (!nbUndoStack.length)
    return;

  notebookRestoreSnapshot(nbUndoStack.pop(), nbRedoStack);
}

function notebookRedo()
{
  if (!nbRedoStack.length)
    return;

  notebookRestoreSnapshot(nbRedoStack.pop(), nbUndoStack);
}

document.addEventListener
(
  'keydown',
  function(e)
  {
    if (currentAppType !== 'notebook' || !nbData)
      return;

    const tag = (document.activeElement && document.activeElement.tagName) || '';

    if (tag === 'INPUT' || tag === 'TEXTAREA' || (document.activeElement && document.activeElement.isContentEditable))
      return;

    if (nbImport)
    {
      if (e.key === 'Enter')
      {
        e.preventDefault();
        acceptNotebookVectorImport();
      }
      else if (e.key === 'Escape')
      {
        e.preventDefault();
        cancelNotebookVectorImport();
      }
      else if (e.key === 'ArrowLeft' || e.key === 'ArrowRight' || e.key === 'ArrowUp' || e.key === 'ArrowDown')
      {
        e.preventDefault();

        const nudge = e.shiftKey ? 10 : 1;

        if (e.key === 'ArrowLeft')  nbImport.x -= nudge;
        if (e.key === 'ArrowRight') nbImport.x += nudge;
        if (e.key === 'ArrowUp')    nbImport.y -= nudge;
        if (e.key === 'ArrowDown')  nbImport.y += nudge;

        redrawNotebookPage(nbImport.pageIndex);
      }

      return;
    }

    if ((e.key === 'Delete' || e.key === 'Backspace') && nbSelection.length)
    {
      notebookSnapshotForUndo(nbCurrentPage);
      notebookDeleteSelection();
      rebuildNotebookCommittedCache(nbCurrentPage);
      commitNotebookChange();
      renderNotebookTextLayerForPage(nbCurrentPage);
      redrawNotebookPage(nbCurrentPage);
    }
    else if ((e.ctrlKey || e.metaKey) && (e.code === 'Equal' || e.code === 'NumpadAdd'))
    {
      // e.code (the physical key) rather than e.key, since Ctrl+Plus is
      // usually pressed as Ctrl+Shift+= - e.key would be '+' there, not '=',
      // and checking only '=' silently missed that case.
      e.preventDefault();
      notebookZoomIn();
    }
    else if ((e.ctrlKey || e.metaKey) && (e.code === 'Minus' || e.code === 'NumpadSubtract'))
    {
      e.preventDefault();
      notebookZoomOut();
    }
    else if ((e.ctrlKey || e.metaKey) && (e.code === 'Digit0' || e.code === 'Numpad0'))
    {
      e.preventDefault();
      notebookZoomReset();
    }
    else if ((e.ctrlKey || e.metaKey) && !e.shiftKey && e.key.toLowerCase() === 'z')
    {
      e.preventDefault();
      notebookUndo();
    }
    else if ((e.ctrlKey || e.metaKey) && (e.key.toLowerCase() === 'y' || (e.shiftKey && e.key.toLowerCase() === 'z')))
    {
      e.preventDefault();
      notebookRedo();
    }
  }
);

// Ctrl+wheel zoom (the standard convention - browsers, Figma, Maps, etc.).
// Must be a non-passive listener for preventDefault() to actually suppress
// the page's own scroll instead of zooming and scrolling at the same time.
document.getElementById('notebook-pages').addEventListener
(
  'wheel',
  function(e)
  {
    if (!e.ctrlKey)
      return;

    e.preventDefault();

    if (e.deltaY < 0)
      notebookZoomIn();
    else if (e.deltaY > 0)
      notebookZoomOut();
  },
  { passive: false }
);

function notebookDeleteSelection()
{
  const page = nbData.pages[nbCurrentPage],
        byKind = { stroke: [], shape: [], text: [] };

  nbSelection.forEach(function(s){ byKind[s.kind].push(s.index); });

  [['stroke','strokes'], ['shape','shapes'], ['text','texts']].forEach(function(pair)
  {
    const indices = byKind[pair[0]].sort(function(a, b){ return b - a; });

    indices.forEach(function(idx)
    {
      page[pair[1]].splice(idx, 1);
    });
  });

  nbSelection = [];
}

function notebookHitTest(page, pt)
{
  for (let i = page.texts.length - 1; i >= 0; i--)
  {
    const b = notebookTextBounds(page.texts[i]);

    if (pt.x >= b.x1 && pt.x <= b.x2 && pt.y >= b.y1 && pt.y <= b.y2)
      return { kind: 'text', index: i };
  }

  for (let i = page.shapes.length - 1; i >= 0; i--)
  {
    const b = notebookShapeBounds(page.shapes[i]);

    if (pt.x >= b.x1 - 6 && pt.x <= b.x2 + 6 && pt.y >= b.y1 - 6 && pt.y <= b.y2 + 6)
      return { kind: 'shape', index: i };
  }

  for (let i = page.strokes.length - 1; i >= 0; i--)
  {
    const pts = page.strokes[i].points;

    for (let j = 0; j < pts.length; j++)
    {
      if (Math.hypot(pts[j].x - pt.x, pts[j].y - pt.y) <= 10)
        return { kind: 'stroke', index: i };
    }
  }

  return null;
}

function notebookItemsInRect(page, rect)
{
  const rx1 = Math.min(rect.x1, rect.x2),
        ry1 = Math.min(rect.y1, rect.y2),
        rx2 = Math.max(rect.x1, rect.x2),
        ry2 = Math.max(rect.y1, rect.y2),
        result = [];

  function intersects(b)
  {
    return (b.x2 >= rx1) && (b.x1 <= rx2) && (b.y2 >= ry1) && (b.y1 <= ry2);
  }

  page.strokes.forEach(function(s, i){ if (intersects(notebookStrokeBounds(s))) result.push({ kind: 'stroke', index: i }); });
  page.shapes.forEach(function(s, i){ if (intersects(notebookShapeBounds(s))) result.push({ kind: 'shape', index: i }); });
  page.texts.forEach(function(s, i){ if (intersects(notebookTextBounds(s))) result.push({ kind: 'text', index: i }); });

  return result;
}

function notebookMoveSelection(page, dx, dy)
{
  nbSelection.forEach(function(sel)
  {
    if (sel.kind === 'stroke')
    {
      page.strokes[sel.index].points.forEach(function(p){ p.x += dx; p.y += dy; });
    }
    else if (sel.kind === 'shape')
    {
      const s = page.shapes[sel.index];
      s.x1 += dx; s.y1 += dy; s.x2 += dx; s.y2 += dy;
    }
    else if (sel.kind === 'text')
    {
      page.texts[sel.index].x += dx;
      page.texts[sel.index].y += dy;
    }
  });
}

// ── Text boxes (DOM overlay, not canvas-drawn, so editing stays native) ──

function renderNotebookTextLayerForPage(pageIndex)
{
  const pageEl = document.querySelector('.notebook-page[data-page-index="' + pageIndex + '"]');

  if (!pageEl || !nbData)
    return;

  const layer = pageEl.querySelector('.notebook-text-layer'),
        page = nbData.pages[pageIndex];

  layer.innerHTML = '';

  page.texts.forEach(function(t, i)
  {
    const div = document.createElement('div');
    div.className = 'notebook-text-box';
    div.contentEditable = 'true';
    div.dataset.textIndex = i;
    div.style.left = t.x + 'px';
    div.style.top = t.y + 'px';
    div.style.width = t.width + 'px';
    div.style.minHeight = t.height + 'px';
    div.style.color = t.color;
    div.style.fontSize = t.fontSize + 'px';
    div.textContent = t.text;

    div.addEventListener('focus', function()
    {
      notebookSnapshotForUndo(pageIndex);
    });

    div.addEventListener('blur', function()
    {
      t.text = div.textContent;
      t.height = Math.max(40, div.offsetHeight);
      commitNotebookChange();
    });

    div.addEventListener('pointerdown', function(e)
    {
      e.stopPropagation();
    });

    layer.appendChild(div);
  });
}

function notebookCreateTextBoxAt(pageIndex, x, y)
{
  notebookSnapshotForUndo(pageIndex);

  nbData.pages[pageIndex].texts.push({ x: x, y: y, width: 220, height: 40, color: nbColor, fontSize: 16, text: '' });

  renderNotebookTextLayerForPage(pageIndex);

  const idx = nbData.pages[pageIndex].texts.length - 1,
        pageEl = document.querySelector('.notebook-page[data-page-index="' + pageIndex + '"]'),
        el = pageEl && pageEl.querySelector('.notebook-text-box[data-text-index="' + idx + '"]');

  // Focusing synchronously inside the same pointerdown that created this element
  // gets clobbered by the browser's own default mousedown focus-handling; deferring
  // it past that default action (a 0ms timeout is the standard fix) makes it stick.
  if (el)
    setTimeout(function(){ el.focus(); }, 0);

  setNotebookTool('select');
}

// ── MAP PAGES ──
// A map page is a regular notebook page with type:'map'. It adds an image
// background, location pins linked to Doc files, a scale reference, and a
// travel-time calculator. Drawing tools still work on top of the map image.

var nbMapTool = null;        // null | 'pin' | 'travel' | 'scale'
var mapTravelPath = [];      // [{x,y}] in page coordinates
var mapTravelPageIndex = -1;
var mapScalePhase = 0;       // 0=idle, 1=waiting for second click
var mapScaleP1 = null;       // {x,y}
var mapScalePageIndex = -1;
var mapActivePinEdit = null; // {pageIndex, pinId}

var MAP_TERRAIN = [
  { id: 'road',  label: 'Road',          color: '#9ca3af' },
  { id: 'trail', label: 'Trail',         color: '#c4a26a' },
  { id: 'rough', label: 'Rough terrain', color: '#7a9e6a' },
  { id: 'sea',   label: 'Sea / River',   color: '#4a9fc8' }
];

function terrainColor(id)
{
  var t = MAP_TERRAIN.find(function(t){ return t.id === id; });
  return t ? t.color : '#9ca3af';
}

function onTerrainSelectChange(id)
{
  var dot = document.getElementById('map-terrain-dot');
  if (dot) dot.style.background = terrainColor(id);
}

// km/day per transport mode, keyed by terrain id.
// null means this mode is not applicable for that terrain.
var TRAVEL_MODES = [
  { id: 'walk',  label: 'Walking',       byTerrain: { road: 35,   trail: 20,   rough: 12,   sea: null } },
  { id: 'horse', label: 'Horse',         byTerrain: { road: 65,   trail: 40,   rough: 22,   sea: null } },
  { id: 'car',   label: 'Car',           byTerrain: { road: 700,  trail: 300,  rough: 100,  sea: null } },
  { id: 'train', label: 'Train',         byTerrain: { road: 900,  trail: null, rough: null, sea: null } },
  { id: 'sail',  label: 'Sailing ship',  byTerrain: { road: null, trail: null, rough: null, sea: 220  } },
  { id: 'motor', label: 'Motor ship',    byTerrain: { road: null, trail: null, rough: null, sea: 550  } },
  { id: 'plane', label: 'Plane',         byTerrain: { road: 8000, trail: 8000, rough: 8000, sea: 8000 } }
];

var mapLastTravelKm       = 0;
var mapLastTravelWpts     = 0;
var mapLastTravelSegments = []; // [{km, terrain}] one per leg

var PIN_COLORS = ['#e74c3c','#e67e22','#f1c40f','#2ecc71','#3498db','#9b59b6','#1abc9c'];

function defaultMapPage()
{
  return {
    type: 'map',
    width: 1600,
    height: 1200,
    background: { type: 'blank' },
    strokes: [],
    shapes: [],
    texts: [],
    imageData: null,
    kmPerPx: null,
    pins: []
  };
}

function notebookAddMapPage()
{
  if (!nbData) return;

  nbData.pages.push(defaultMapPage());
  renderNotebookPages();
  commitNotebookChange();
  notebookGoToPage(nbData.pages.length - 1);
}

function setMapTool(tool)
{
  var prevTool = nbMapTool;
  nbMapTool = (nbMapTool === tool) ? null : tool;

  if (prevTool === 'travel')
  {
    // Toggled off via travel button = user wants to calculate the accumulated path
    var toggled = (prevTool === tool);
    if (toggled && mapTravelPath.length >= 2 && mapTravelPageIndex !== -1)
      finishTravelCalc(mapTravelPageIndex);
    else
      clearTravelOverlay();

    mapTravelPath = [];
    mapTravelPageIndex = -1;
  }
  else if (nbMapTool !== 'travel')
  {
    clearTravelOverlay();
  }

  if (nbMapTool !== 'scale')
  {
    mapScalePhase = 0;
    mapScaleP1 = null;
    clearScaleOverlay();
  }

  updateMapToolbarState();
  updateMapCursor();

  if (nbMapTool === 'scale')
    showMapHint('Click the first reference point on the map.');
  else if (nbMapTool === 'travel')
    showMapHint('Click to add waypoints. Click the travel button again to calculate.');
  else if (nbMapTool === 'pin')
    showMapHint('Click on the map to place a pin.');
  else
    showMapHint('');
}

function showMapHint(text)
{
  var el = document.getElementById('map-hint');
  if (el) el.textContent = text;
}

function updateMapToolbarState()
{
  document.querySelectorAll('.map-tool-btn').forEach(function(btn)
  {
    btn.classList.toggle('active', btn.dataset.mapTool === nbMapTool);
  });

  var terrainGroup = document.getElementById('map-terrain-group');
  if (terrainGroup)
  {
    terrainGroup.style.display = (nbMapTool === 'travel') ? 'flex' : 'none';
    if (nbMapTool === 'travel')
    {
      var sel = document.getElementById('map-terrain-select');
      onTerrainSelectChange(sel ? sel.value : 'road');
    }
  }

  // Block pin clicks while travel tool is active
  document.querySelectorAll('.nb-pin-layer').forEach(function(pl)
  {
    pl.classList.toggle('travel-active', nbMapTool === 'travel');
  });
}

function updateMapCursor()
{
  var pages = document.getElementById('notebook-pages');
  if (!pages) return;

  pages.className = pages.className.replace(/\btool-map-\S+/g, '').trim();

  if (nbMapTool === 'pin')    pages.classList.add('tool-map-pin');
  if (nbMapTool === 'travel') pages.classList.add('tool-map-travel');
  if (nbMapTool === 'scale')  pages.classList.add('tool-map-scale');
}

function switchNbTab(tab)
{
  var draw = document.getElementById('nb-draw-tools');
  var map  = document.getElementById('nb-map-tools');
  if (draw) draw.style.display = (tab === 'draw') ? 'contents' : 'none';
  if (map)  map.style.display  = (tab === 'map')  ? 'flex'     : 'none';

  var btnDraw = document.getElementById('nb-tab-draw');
  var btnMap  = document.getElementById('nb-tab-map');
  if (btnDraw) btnDraw.classList.toggle('active', tab === 'draw');
  if (btnMap)  btnMap.classList.toggle('active',  tab === 'map');
}

function updateMapToolbarVisibility()
{
  var page = nbData && nbData.pages[nbCurrentPage];
  var isMap = !!(page && page.type === 'map');

  switchNbTab(isMap ? 'map' : 'draw');

  if (!isMap)
  {
    nbMapTool = null;
    mapTravelPath = [];
    mapTravelPageIndex = -1;
    clearTravelOverlay();
    clearScaleOverlay();
    showMapHint('');
  }
  else
  {
    updateMapCursor();
    updateMapToolbarState();
    refreshMapScaleDisplay();
  }
}

function refreshMapScaleDisplay()
{
  var page = nbData && nbData.pages[nbCurrentPage];
  var el = document.getElementById('map-scale-display');
  if (!el || !page) return;

  if (page.kmPerPx)
  {
    var pxPerKm = 1 / page.kmPerPx;
    el.textContent = '1 px = ' + page.kmPerPx.toFixed(3) + ' km';
    el.style.display = '';
  }
  else
  {
    el.textContent = 'Scale not set';
    el.style.display = '';
  }
}

// ── Pin layer ──

function renderMapPinLayer(pageIndex)
{
  var pinLayer = document.querySelector('.notebook-page[data-page-index="' + pageIndex + '"] .nb-pin-layer');
  if (!pinLayer || !nbData) return;

  var page = nbData.pages[pageIndex];
  if (!page || page.type !== 'map') return;

  var pins = page.pins || [];

  pinLayer.innerHTML = pins.map(function(pin)
  {
    var lp = ((pin.x / page.width) * 100).toFixed(4) + '%';
    var tp = ((pin.y / page.height) * 100).toFixed(4) + '%';
    var color = escAttr(pin.color || PIN_COLORS[0]);
    var label = escHtml(pin.label || '');

    var sw = Math.round(18 * nbZoom), sh = Math.round(22 * nbZoom);
    return '<div class="map-pin" style="left:' + lp + ';top:' + tp + ';--pin-color:' + color + '"' +
           ' onclick="openPinPopover(' + pageIndex + ',\'' + escAttr(pin.id) + '\',event)" title="' + label + '">' +
             '<svg width="' + sw + '" height="' + sh + '" viewBox="0 0 24 24" fill="' + color + '" stroke="#0005" stroke-width="0.5"><path d="M12 2C8.13 2 5 5.13 5 9c0 5.25 7 13 7 13s7-7.75 7-13c0-3.87-3.13-7-7-7zm0 9.5c-1.38 0-2.5-1.12-2.5-2.5s1.12-2.5 2.5-2.5 2.5 1.12 2.5 2.5-1.12 2.5-2.5 2.5z"/></svg>' +
             (label ? '<span class="map-pin-label">' + label + '</span>' : '') +
           '</div>';
  }).join('');
}

// ── Pin popover ──

function openPinPopover(pageIndex, pinId, event)
{
  event.stopPropagation();

  var page = nbData && nbData.pages[pageIndex];
  if (!page) return;

  var pin = (page.pins || []).find(function(p){ return p.id === pinId; });
  if (!pin) return;

  mapActivePinEdit = { pageIndex: pageIndex, pinId: pinId };

  var pop = document.getElementById('pin-popover');
  if (!pop) return;

  document.getElementById('pin-label-input').value = pin.label || '';

  var docSel = document.getElementById('pin-doc-select');
  docSel.innerHTML = '<option value="">— No linked document —</option>' +
    Object.entries(files)
      .filter(function(e){ return e[1].type === 'doc'; })
      .sort(function(a, b){ return a[1].name.localeCompare(b[1].name); })
      .map(function(e){ return '<option value="' + escAttr(e[0]) + '"' + (e[0] === pin.docId ? ' selected' : '') + '>' + escHtml(e[1].name) + '</option>'; })
      .join('');

  var swatches = document.getElementById('pin-color-swatches');
  swatches.innerHTML = PIN_COLORS.map(function(c)
  {
    return '<span class="pin-swatch' + (c === (pin.color || PIN_COLORS[0]) ? ' active' : '') + '"' +
           ' style="background:' + c + '" onclick="selectPinColor(\'' + escAttr(c) + '\',this)"></span>';
  }).join('');

  pop.style.display = '';
  var rect = event.target.closest('.map-pin').getBoundingClientRect();
  pop.style.left = Math.min(rect.right + 8, window.innerWidth - 280) + 'px';
  pop.style.top  = Math.max(rect.top - 10, 10) + 'px';
}

function selectPinColor(color, el)
{
  document.querySelectorAll('#pin-color-swatches .pin-swatch').forEach(function(s){ s.classList.remove('active'); });
  el.classList.add('active');
}

function savePinEdits()
{
  if (!mapActivePinEdit) return;

  var pageIndex = mapActivePinEdit.pageIndex,
      pinId = mapActivePinEdit.pinId,
      page = nbData && nbData.pages[pageIndex];
  if (!page) return;

  var pin = (page.pins || []).find(function(p){ return p.id === pinId; });
  if (!pin) return;

  pin.label = document.getElementById('pin-label-input').value.trim();
  pin.docId = document.getElementById('pin-doc-select').value || null;

  var activeSwatchEl = document.querySelector('#pin-color-swatches .pin-swatch.active');
  if (activeSwatchEl)
    pin.color = activeSwatchEl.style.background;

  closePinPopover();
  renderMapPinLayer(pageIndex);
  commitNotebookChange();
}

function closePinPopover()
{
  var pop = document.getElementById('pin-popover');
  if (pop) pop.style.display = 'none';
  mapActivePinEdit = null;
}

function openPinDoc()
{
  if (!mapActivePinEdit) return;

  var pageIndex = mapActivePinEdit.pageIndex,
      pinId = mapActivePinEdit.pinId,
      page = nbData && nbData.pages[pageIndex];
  if (!page) return;

  var pin = (page.pins || []).find(function(p){ return p.id === pinId; });
  if (!pin || !pin.docId) { alert('No document linked to this pin.'); return; }

  closePinPopover();
  switchAppType('doc');
  openFile(pin.docId);
}

function deleteCurrentPin()
{
  if (!mapActivePinEdit) return;

  var pageIndex = mapActivePinEdit.pageIndex,
      pinId = mapActivePinEdit.pinId,
      page = nbData && nbData.pages[pageIndex];
  if (!page) return;

  page.pins = (page.pins || []).filter(function(p){ return p.id !== pinId; });

  closePinPopover();
  renderMapPinLayer(pageIndex);
  commitNotebookChange();
}

function addMapPin(pageIndex, x, y)
{
  var page = nbData && nbData.pages[pageIndex];
  if (!page) return;

  var pin = {
    id: 'pin_' + Date.now(),
    x: x,
    y: y,
    label: '',
    docId: null,
    color: PIN_COLORS[0]
  };

  if (!page.pins) page.pins = [];
  page.pins.push(pin);
  renderMapPinLayer(pageIndex);
  commitNotebookChange();

  // Open popover for immediate naming
  var pinEl = document.querySelector(
    '.notebook-page[data-page-index="' + pageIndex + '"] .map-pin:last-child'
  );
  if (pinEl)
  {
    // Fake a click event on the pin element
    var fakeEvent = { stopPropagation: function(){}, target: pinEl };
    openPinPopover(pageIndex, pin.id, fakeEvent);
  }
}

// ── Scale tool ──

function clearScaleOverlay()
{
  document.querySelectorAll('.nb-scale-canvas').forEach(function(c)
  {
    var ctx = c.getContext('2d');
    ctx.clearRect(0, 0, c.width, c.height);
  });
}

function handleScaleClick(pageIndex, x, y, clientX, clientY)
{
  if (mapScalePhase === 0)
  {
    mapScalePhase = 1;
    mapScaleP1 = { x: x, y: y };
    mapScalePageIndex = pageIndex;
    showMapHint('Click the second reference point.');
    drawScalePoint(pageIndex, x, y, null, null);
  }
  else if (mapScalePhase === 1 && mapScalePageIndex === pageIndex)
  {
    var dx = x - mapScaleP1.x,
        dy = y - mapScaleP1.y,
        px = Math.sqrt(dx * dx + dy * dy);

    mapScalePhase = 0;
    clearScaleOverlay();

    if (px < 5) { showMapHint('Points too close — try again.'); return; }

    var km = parseFloat(prompt('Distance between the two points in km:'));
    if (!km || km <= 0) { showMapHint(''); return; }

    var page = nbData.pages[pageIndex];
    page.kmPerPx = km / px;

    commitNotebookChange();
    refreshMapScaleDisplay();
    showMapHint('Scale set: ' + page.kmPerPx.toFixed(4) + ' km/px.');
    nbMapTool = null;
    updateMapToolbarState();
    updateMapCursor();
  }
}

function drawScalePoint(pageIndex, x, y, x2, y2)
{
  var canvas = document.querySelector('.notebook-page[data-page-index="' + pageIndex + '"] .nb-scale-canvas');
  if (!canvas || !nbData) return;

  var page = nbData.pages[pageIndex];
  var ctx = canvas.getContext('2d');
  ctx.clearRect(0, 0, canvas.width, canvas.height);

  var sx = canvas.width  / page.width;
  var sy = canvas.height / page.height;

  ctx.strokeStyle = '#3498db';
  ctx.fillStyle   = '#3498db';
  ctx.lineWidth   = 2;

  ctx.beginPath();
  ctx.arc(x * sx, y * sy, 6, 0, Math.PI * 2);
  ctx.fill();

  if (x2 !== null)
  {
    ctx.beginPath();
    ctx.arc(x2 * sx, y2 * sy, 6, 0, Math.PI * 2);
    ctx.fill();
    ctx.beginPath();
    ctx.moveTo(x * sx, y * sy);
    ctx.lineTo(x2 * sx, y2 * sy);
    ctx.stroke();
  }
}

// ── Travel tool ──

function clearTravelOverlay()
{
  document.querySelectorAll('.nb-travel-canvas').forEach(function(c)
  {
    var ctx = c.getContext('2d');
    ctx.clearRect(0, 0, c.width, c.height);
  });
  mapLastTravelKm = 0; mapLastTravelWpts = 0; mapLastTravelSegments = [];
  closeTravelPopover();
  var el = document.getElementById('map-travel-result');
  if (el) el.textContent = '';
}

function addTravelWaypoint(pageIndex, x, y)
{
  if (mapTravelPageIndex === -1) mapTravelPageIndex = pageIndex;
  if (mapTravelPageIndex !== pageIndex) return;

  var terrain = (document.getElementById('map-terrain-select') || {}).value || 'road';
  mapTravelPath.push({ x: x, y: y, terrain: terrain });
  renderTravelPath(pageIndex);

  var n = mapTravelPath.length;
  var tLabel = (MAP_TERRAIN.find(function(t){ return t.id === terrain; }) || {}).label || terrain;
  showMapHint(n + ' waypoint' + (n === 1 ? '' : 's') + ' · next leg: ' + tLabel + ' — change terrain or keep clicking.');
}

function formatDays(days)
{
  if (days < 1 / 24) return '< 1 hr';
  if (days < 1)
  {
    var hrs = Math.round(days * 24);
    return '~' + hrs + ' hr' + (hrs === 1 ? '' : 's');
  }
  var d = Math.floor(days);
  var h = Math.round((days - d) * 24);
  if (h === 0 || d >= 14) return '~' + d + (d === 1 ? ' day' : ' days');
  return '~' + d + 'd ' + h + 'h';
}

function formatTravelTime(km, kmPerDay)
{
  return formatDays(km / kmPerDay);
}

function finishTravelCalc(pageIndex)
{
  var page = nbData && nbData.pages[pageIndex];
  if (!page) return;

  if (mapTravelPath.length < 2)
  {
    showMapHint('Add at least 2 waypoints before calculating.');
    return;
  }

  if (!page.kmPerPx)
  {
    showMapHint('Scale not set — use the scale tool first.');
    clearTravelOverlay();
    return;
  }

  mapLastTravelSegments = [];
  var totalKm = 0;
  for (var i = 1; i < mapTravelPath.length; i++)
  {
    var dx = mapTravelPath[i].x - mapTravelPath[i-1].x,
        dy = mapTravelPath[i].y - mapTravelPath[i-1].y;
    var segKm = Math.sqrt(dx*dx + dy*dy) * page.kmPerPx;
    // terrain of a segment is the terrain recorded at the FROM waypoint
    mapLastTravelSegments.push({ km: segKm, terrain: mapTravelPath[i-1].terrain || 'road' });
    totalKm += segKm;
  }

  mapLastTravelKm   = totalKm;
  mapLastTravelWpts = mapTravelPath.length;

  var el = document.getElementById('map-travel-result');
  if (el) el.textContent = mapLastTravelWpts + ' pts · ' + mapLastTravelKm.toFixed(1) + ' km';

  renderTravelPath(pageIndex);
  showMapHint('Click the result chip for travel times.');
}

function openTravelPopover(e)
{
  if (e) e.stopPropagation();
  if (!mapLastTravelKm) return;

  var pop = document.getElementById('travel-popover');
  var header = document.getElementById('travel-pop-header');
  var modesEl = document.getElementById('travel-pop-modes');
  if (!pop || !header || !modesEl) return;

  header.textContent = mapLastTravelKm.toFixed(1) + ' km · ' + mapLastTravelWpts + ' waypoints';

  // Build a summary of which terrains appear in the route
  var terrainSeen = {};
  mapLastTravelSegments.forEach(function(s){ terrainSeen[s.terrain] = true; });

  modesEl.innerHTML = TRAVEL_MODES.map(function(mode)
  {
    // Sum days per segment using each segment's own terrain speed
    var totalDays = 0;
    var hasImpassable = false;
    mapLastTravelSegments.forEach(function(seg)
    {
      var speed = mode.byTerrain[seg.terrain];
      if (!speed) { hasImpassable = true; }
      else { totalDays += seg.km / speed; }
    });

    var na = hasImpassable && totalDays === 0;  // every segment impassable
    var partial = hasImpassable && totalDays > 0; // some segments impassable

    var timeStr, speedStr;
    if (na)
    {
      timeStr  = '—';
      speedStr = 'route has impassable terrain';
    }
    else if (partial)
    {
      timeStr  = formatDays(totalDays) + '*';
      speedStr = '* partial — some terrain N/A';
    }
    else
    {
      // All segments passable: show min–max speed range if mixed terrain
      var speeds = Object.keys(terrainSeen).map(function(id){ return mode.byTerrain[id]; }).filter(Boolean);
      var minSpd = Math.min.apply(null, speeds), maxSpd = Math.max.apply(null, speeds);
      speedStr = minSpd === maxSpd ? minSpd + ' km/day' : minSpd + '–' + maxSpd + ' km/day';
      timeStr  = formatDays(totalDays);
    }

    return '<div class="travel-mode-row' + (na ? ' travel-mode-na' : '') + '">' +
      '<span class="travel-mode-name">' + mode.label + '</span>' +
      '<span class="travel-mode-dur">' + timeStr + '</span>' +
      '<span class="travel-mode-rate">' + speedStr + '</span>' +
    '</div>';
  }).join('');

  var chip = document.getElementById('map-travel-result');
  if (chip)
  {
    var rect = chip.getBoundingClientRect();
    pop.style.left = rect.left + 'px';
    pop.style.top  = (rect.bottom + 6) + 'px';
  }
  pop.style.display = 'block';
}

function closeTravelPopover()
{
  var pop = document.getElementById('travel-popover');
  if (pop) pop.style.display = 'none';
}

function renderTravelPath(pageIndex)
{
  var canvas = document.querySelector('.notebook-page[data-page-index="' + pageIndex + '"] .nb-travel-canvas');
  if (!canvas || !nbData) return;

  var page = nbData.pages[pageIndex];
  var ctx = canvas.getContext('2d');

  // Scale from page units (0..page.width/height) to canvas buffer pixels.
  // canvas.width was created as page.width * notebookCanvasScale(), so the
  // ratio is simply notebookCanvasScale().  Recalculate each render to stay
  // correct if the canvas was recreated after an image upload.
  var sx = canvas.width  / page.width;
  var sy = canvas.height / page.height;

  ctx.clearRect(0, 0, canvas.width, canvas.height);

  if (mapTravelPath.length < 1) return;

  ctx.save();
  ctx.lineWidth = 3;
  ctx.lineCap   = 'round';
  ctx.setLineDash([Math.round(8 * sx), Math.round(4 * sx)]);

  // Draw each segment in its terrain colour
  for (var i = 1; i < mapTravelPath.length; i++)
  {
    var color = terrainColor(mapTravelPath[i-1].terrain || 'road');
    ctx.strokeStyle = color;
    ctx.beginPath();
    ctx.moveTo(mapTravelPath[i-1].x * sx, mapTravelPath[i-1].y * sy);
    ctx.lineTo(mapTravelPath[i].x   * sx, mapTravelPath[i].y   * sy);
    ctx.stroke();
  }

  // Draw waypoint dots in their terrain colour
  ctx.setLineDash([]);
  mapTravelPath.forEach(function(pt)
  {
    ctx.fillStyle = terrainColor(pt.terrain || 'road');
    ctx.beginPath();
    ctx.arc(pt.x * sx, pt.y * sy, 5, 0, Math.PI * 2);
    ctx.fill();
  });

  ctx.restore();
}

// ── Map image upload ──

function triggerMapImageUpload()
{
  var inp = document.getElementById('map-image-input');
  if (inp) inp.click();
}

function handleMapImageUpload(event)
{
  var file = event.target.files[0];
  if (!file) return;

  event.target.value = '';

  var reader = new FileReader();

  reader.onload = function(e)
  {
    var img = new Image();

    img.onload = function()
    {
      var maxW = 1920, maxH = 1440;
      var scale = Math.min(1, maxW / img.naturalWidth, maxH / img.naturalHeight);
      var dw = Math.round(img.naturalWidth * scale);
      var dh = Math.round(img.naturalHeight * scale);

      var tmp = document.createElement('canvas');
      tmp.width = dw;
      tmp.height = dh;
      tmp.getContext('2d').drawImage(img, 0, 0, dw, dh);

      var page = nbData.pages[nbCurrentPage];
      page.imageData = tmp.toDataURL('image/jpeg', 0.85);
      page.width = dw;
      page.height = dh;

      commitNotebookChange();
      renderNotebookPages();
      notebookGoToPage(nbCurrentPage);
    };

    img.src = e.target.result;
  };

  reader.readAsDataURL(file);
}

// ── Map pointer dispatching ──

// Converts a pointer event to page-space coordinates using the PAGE ELEMENT's
// bounding rect directly — the same reference as pin % positions — rather than
// the ink-canvas buffer dimensions, which can diverge due to CSS inset vs
// width/height differences or max-width constraints.
function mapPagePoint(e, pageIndex)
{
  var pageEl = document.querySelector('.notebook-page[data-page-index="' + pageIndex + '"]');
  if (!pageEl || !nbData) return { x: 0, y: 0 };
  var rect = pageEl.getBoundingClientRect();
  var page = nbData.pages[pageIndex];
  return {
    x: Math.max(0, Math.min(page.width,  (e.clientX - rect.left) / rect.width  * page.width)),
    y: Math.max(0, Math.min(page.height, (e.clientY - rect.top)  / rect.height * page.height))
  };
}

function findNearestPin(page, pt, hitRadius)
{
  hitRadius = hitRadius || 20;
  var best = null, bestD = Infinity;

  (page.pins || []).forEach(function(pin)
  {
    var d = Math.hypot(pin.x - pt.x, pin.y - pt.y);
    if (d < hitRadius && d < bestD) { best = pin; bestD = d; }
  });

  return best;
}

function handleMapPointerDown(e, _pt, canvas, pageIndex)
{
  // Use page-element coordinates for all map operations so they align with
  // pin % positions regardless of canvas buffer size or CSS constraints.
  var pt = mapPagePoint(e, pageIndex);
  var page = nbData.pages[pageIndex];

  if (nbMapTool === 'pin')
  {
    var hit = findNearestPin(page, pt, 16);
    if (hit)
    {
      var fakeEvent = { stopPropagation: function(){}, target: canvas };
      openPinPopover(pageIndex, hit.id, fakeEvent);
    }
    else
    {
      addMapPin(pageIndex, pt.x, pt.y);
    }
    return;
  }

  if (nbMapTool === 'scale')
  {
    handleScaleClick(pageIndex, pt.x, pt.y, e.clientX, e.clientY);
    return;
  }

  if (nbMapTool === 'travel')
  {
    if (e.button === 0)
      addTravelWaypoint(pageIndex, pt.x, pt.y);
    return;
  }

  // No map tool active — check if clicking a pin
  var hit = findNearestPin(page, pt, 20);
  if (hit)
  {
    var pinEl = document.querySelector(
      '.notebook-page[data-page-index="' + pageIndex + '"] .map-pin[title="' + (hit.label || '') + '"]'
    );
    var fakeEvt = {
      stopPropagation: function(){},
      target: pinEl || canvas
    };
    openPinPopover(pageIndex, hit.id, fakeEvt);
    return;
  }

  // Fall through to normal notebook tools
  return false;
}

function handleMapContextMenu(e, canvas, pageIndex)
{
  // Suppress browser context menu while any map tool is active
  if (nbMapTool) e.preventDefault();
}

// ── Pointer handling per page ──

function wireNotebookPointerEvents(canvas, pageIndex)
{
  canvas.addEventListener('pointerdown', function(e){ notebookPointerDown(e, canvas, pageIndex); });
  canvas.addEventListener('pointermove', function(e){ notebookPointerMove(e, canvas, pageIndex); });
  canvas.addEventListener('pointerup', function(e){ notebookPointerUp(e, canvas, pageIndex); });
  canvas.addEventListener('pointercancel', function(e){ notebookPointerUp(e, canvas, pageIndex); });
  canvas.addEventListener('contextmenu', function(e){ handleMapContextMenu(e, canvas, pageIndex); });
}

function notebookCanvasPoint(e, canvas)
{
  // canvas.width/height are the raw (oversampled) buffer size, but
  // drawing/storage all happens in logical page units - dividing the canvas
  // scale back out here first recovers the logical size, so what's left is
  // purely the CSS-responsive-shrink factor (the page can be narrower than
  // its logical width on a smaller window).
  const rect = canvas.getBoundingClientRect(),
        canvasScale = notebookCanvasScale(),
        scaleX = (canvas.width / canvasScale) / rect.width,
        scaleY = (canvas.height / canvasScale) / rect.height;

  return {
    x: (e.clientX - rect.left) * scaleX,
    y: (e.clientY - rect.top) * scaleY,
    pressure: e.pressure,
    pointerType: e.pointerType
  };
}

function notebookEffectiveWidth(pt)
{
  if (pt.pointerType !== 'pen')
    return nbWidth;

  // A bare `pressure > 0` gate means a single stray zero-pressure sample (not
  // uncommon between driver samples) snaps the line straight to its thinnest
  // width for that point instead of holding steady - falling back to the
  // spec's neutral 0.5 avoids that visible pinch. The range here is wide
  // (0.3x-1.7x) since the previous 0.5x-1.5x spread was barely perceptible
  // once strokes were actually crisp.
  const pressure = (typeof pt.pressure === 'number' && pt.pressure > 0) ? pt.pressure : 0.5;

  return nbWidth * (0.3 + pressure * 1.4);
}

function notebookEraseAt(pt, active)
{
  const page = nbData.pages[active.pageIndex],
        r = active.radius,
        newStrokes = [];

  page.strokes.forEach(function(stroke)
  {
    // Filled strokes (SVG import fills) erase as whole objects: splitting a
    // fill outline into open fragments has no meaningful rendering.
    if (stroke.fill)
    {
      const hit = stroke.points.some(function(p){ return Math.hypot(p.x - pt.x, p.y - pt.y) <= r; });

      if (!hit)
        newStrokes.push(stroke);

      return;
    }

    const segments = [[]];

    stroke.points.forEach(function(p)
    {
      if (Math.hypot(p.x - pt.x, p.y - pt.y) <= r)
      {
        if (segments[segments.length - 1].length)
          segments.push([]);
      }
      else
      {
        segments[segments.length - 1].push(p);
      }
    });

    segments.forEach(function(seg)
    {
      if (seg.length > 1)
        newStrokes.push({ tool: stroke.tool, color: stroke.color, baseWidth: stroke.baseWidth, points: seg });
    });
  });

  page.strokes = newStrokes;
}

function notebookPointerDown(e, canvas, pageIndex)
{
  if (!nbData)
    return;

  canvas.setPointerCapture(e.pointerId);

  const pt = notebookCanvasPoint(e, canvas);
  nbCurrentPage = pageIndex;

  // A pending vector import captures all pointer input on its page: dragging
  // repositions it, the bottom-right handle resizes it, and normal tools stay
  // disabled until the import is accepted or cancelled.
  if (nbImport)
  {
    if (pageIndex !== nbImport.pageIndex)
      return;

    const hx = nbImport.x + nbImport.w * nbImport.scale + 6,
          hy = nbImport.y + nbImport.h * nbImport.scale + 6;

    if (Math.abs(pt.x - hx) <= 12 && Math.abs(pt.y - hy) <= 12)
      nbActive = { kind: 'import-scale', pageIndex: pageIndex };
    else
      nbActive = { kind: 'import-move', pageIndex: pageIndex, start: pt };

    return;
  }

  // Map pages: route to map tool handler first; fall through to drawing tools if not consumed.
  const _mapPage = nbData.pages[pageIndex];
  if (_mapPage && _mapPage.type === 'map')
  {
    const _consumed = handleMapPointerDown(e, pt, canvas, pageIndex);
    if (_consumed !== false)
      return;
  }

  if (nbTool === 'pen' || nbTool === 'highlighter')
  {
    notebookSnapshotForUndo(pageIndex);

    const stroke = { tool: nbTool, color: nbColor, baseWidth: nbWidth, points: [{ x: pt.x, y: pt.y, w: notebookEffectiveWidth(pt) }] };

    nbData.pages[pageIndex].strokes.push(stroke);
    // smoothX/Y track a filtered position (see notebookPointerMove) - seeded
    // from the real first point since there's nothing yet to smooth against.
    nbActive = { kind: 'stroke', pageIndex: pageIndex, stroke: stroke, smoothX: pt.x, smoothY: pt.y };
    redrawNotebookPage(pageIndex);
  }
  else if (nbTool === 'eraser')
  {
    notebookSnapshotForUndo(pageIndex);
    nbActive = { kind: 'erase', pageIndex: pageIndex, radius: Math.max(8, nbWidth * 3) };
    notebookEraseAt(pt, nbActive);
    redrawNotebookPage(pageIndex);
  }
  else if (nbTool === 'line' || nbTool === 'rectangle' || nbTool === 'ellipse')
  {
    notebookSnapshotForUndo(pageIndex);

    const shape = { tool: nbTool, color: nbColor, width: nbWidth, x1: pt.x, y1: pt.y, x2: pt.x, y2: pt.y };

    nbData.pages[pageIndex].shapes.push(shape);
    nbActive = { kind: 'shape', pageIndex: pageIndex, shape: shape };
    redrawNotebookPage(pageIndex);
  }
  else if (nbTool === 'select')
  {
    const page = nbData.pages[pageIndex],
          hit = notebookHitTest(page, pt),
          alreadySelected = hit && nbSelection.some(function(s){ return s.kind === hit.kind && s.index === hit.index; });

    if (hit && !alreadySelected)
      nbSelection = [hit];

    if (hit)
    {
      nbActive = { kind: 'move', pageIndex: pageIndex, start: pt, snapshotted: false };
    }
    else
    {
      nbSelection = [];
      nbActive = { kind: 'marquee', pageIndex: pageIndex, x1: pt.x, y1: pt.y, x2: pt.x, y2: pt.y };
    }

    redrawNotebookPage(pageIndex);
  }
  else if (nbTool === 'text')
  {
    notebookCreateTextBoxAt(pageIndex, pt.x, pt.y);
    nbActive = null;
  }
}

function notebookPointerMove(e, canvas, pageIndex)
{
  if (!nbActive || nbActive.pageIndex !== pageIndex)
    return;

  const pt = notebookCanvasPoint(e, canvas);

  if (nbActive.kind === 'stroke')
  {
    // Stylus samples carry real jitter that the quadratic-curve rendering
    // alone doesn't hide - it smooths the path *between* recorded points, but
    // can't fix noisy points. An exponential moving average on the position
    // (a simple, cheap "stroke stabilizer") filters that jitter out before the
    // point is even stored, which is what was making strokes look shaky/faceted
    // rather than like handwriting. 0.4 trades a little responsiveness for a
    // visibly steadier line; raw pressure is kept as-is for width.
    const SMOOTHING = 0.4;

    nbActive.smoothX += (pt.x - nbActive.smoothX) * SMOOTHING;
    nbActive.smoothY += (pt.y - nbActive.smoothY) * SMOOTHING;

    nbActive.stroke.points.push({ x: nbActive.smoothX, y: nbActive.smoothY, w: notebookEffectiveWidth(pt) });
    redrawNotebookPage(pageIndex);
  }
  else if (nbActive.kind === 'erase')
  {
    notebookEraseAt(pt, nbActive);
    redrawNotebookPage(pageIndex);
  }
  else if (nbActive.kind === 'shape')
  {
    nbActive.shape.x2 = pt.x;
    nbActive.shape.y2 = pt.y;
    redrawNotebookPage(pageIndex);
  }
  else if (nbActive.kind === 'marquee')
  {
    nbActive.x2 = pt.x;
    nbActive.y2 = pt.y;
    nbSelection = notebookItemsInRect(nbData.pages[pageIndex], nbActive);
    redrawNotebookPage(pageIndex);
  }
  else if (nbActive.kind === 'import-move')
  {
    nbImport.x += pt.x - nbActive.start.x;
    nbImport.y += pt.y - nbActive.start.y;
    nbActive.start = pt;
    redrawNotebookPage(pageIndex);
  }
  else if (nbActive.kind === 'import-scale')
  {
    // Scale so the corner handle tracks the pointer, anchored at the top-left.
    const s = Math.max((pt.x - nbImport.x) / nbImport.w, (pt.y - nbImport.y) / nbImport.h);

    nbImport.scale = Math.max(0.05, Math.min(20, s));
    redrawNotebookPage(pageIndex);
  }
  else if (nbActive.kind === 'move')
  {
    if (!nbActive.snapshotted)
    {
      notebookSnapshotForUndo(pageIndex);
      nbActive.snapshotted = true;
    }

    const dx = pt.x - nbActive.start.x,
          dy = pt.y - nbActive.start.y;

    nbActive.start = pt;
    notebookMoveSelection(nbData.pages[pageIndex], dx, dy);
    renderNotebookTextLayerForPage(pageIndex);
    redrawNotebookPage(pageIndex);
  }
}

function notebookPointerUp(e, canvas, pageIndex)
{
  if (!nbActive)
    return;

  const kind = nbActive.kind;

  // The smoothing filter in notebookPointerMove always trails slightly behind
  // the real pointer position: snap to the exact lift-off point so the stroke
  // visibly reaches all the way to where the pen actually left the surface.
  if (kind === 'stroke')
  {
    const pt = notebookCanvasPoint(e, canvas);
    nbActive.stroke.points.push({ x: pt.x, y: pt.y, w: notebookEffectiveWidth(pt) });
  }

  nbActive = null;

  if (kind === 'stroke' || kind === 'erase' || kind === 'shape' || kind === 'move')
  {
    rebuildNotebookCommittedCache(pageIndex);
    commitNotebookChange();
  }

  redrawNotebookPage(pageIndex);
}

// ── PDF import ──

function triggerNotebookPdfImport()
{
  document.getElementById('notebook-pdf-input').click();
}

async function handleNotebookPdfImport(e)
{
  const file = e.target.files[0];
  e.target.value = '';

  if (!file || !nbData)
    return;

  if (typeof pdfjsLib === 'undefined')
  {
    alert('PDF support is still loading - please try again in a moment.');
    return;
  }

  const arrayBuffer = await file.arrayBuffer(),
        bytes = new Uint8Array(arrayBuffer);

  let base64 = '';
  const chunkSize = 0x8000;

  for (let i = 0; i < bytes.length; i += chunkSize)
    base64 += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));

  base64 = btoa(base64);

  try
  {
    const pdfDoc = await pdfjsLib.getDocument({ data: bytes.slice() }).promise,
          newPages = [];

    for (let p = 1; p <= pdfDoc.numPages; p++)
    {
      const pdfPage = await pdfDoc.getPage(p),
            viewport = pdfPage.getViewport({ scale: 1.5 });

      newPages.push
      (
        {
          width: Math.round(viewport.width),
          height: Math.round(viewport.height),
          background: { type: 'pdf', pdfPageIndex: p - 1 },
          strokes: [],
          shapes: [],
          texts: []
        }
      );
    }

    nbData.pdfData = 'data:application/pdf;base64,' + base64;
    nbData.pages = nbData.pages.concat(newPages);
    const firstNewPageIndex = nbData.pages.length - newPages.length;

    nbPdfDocCache = pdfDoc;
    nbPdfDocCacheKey = nbData.pdfData;

    renderNotebookPages();
    commitNotebookChange();
    notebookGoToPage(firstNewPageIndex);
  }
  catch(err)
  {
    console.warn('PDF import error', err);
    alert('Could not import that PDF.');
  }
}

// ── SVG vector import ──

function triggerNotebookVectorImport()
{
  document.getElementById('notebook-svg-input').click();
}

async function handleNotebookVectorImport(e)
{
  const file = e.target.files[0];
  e.target.value = '';

  if (!file || !nbData)
    return;

  const text = await file.text(),
        converted = convertSvgToNotebookStrokes(text);

  if (!converted || !converted.strokes.length)
  {
    alert('No drawable vector shapes were found in that SVG.');
    return;
  }

  const page = nbData.pages[nbCurrentPage],
        // Start at a comfortable size: natural size when it fits, otherwise
        // shrunk to at most ~70% of the page.
        scale = Math.min(1, (page.width * 0.7) / converted.w, (page.height * 0.7) / converted.h);

  nbImport = {
    pageIndex: nbCurrentPage,
    strokes: converted.strokes,
    w: converted.w,
    h: converted.h,
    scale: scale,
    x: (page.width - converted.w * scale) / 2,
    y: (page.height - converted.h * scale) / 2
  };

  document.getElementById('nb-import-bar').style.display = 'flex';
  redrawNotebookPage(nbCurrentPage);
}

// Parses an SVG path's `d` into flattened polyline runs (one per subpath),
// evaluating lines and quadratic/cubic curves directly. The DOM alternative,
// getPointAtLength, re-walks the whole segment list on every call, which
// turns dense sampling of big paths (Excalidraw freedraw exports easily
// reach hundreds of curve segments each) quadratic - tens of seconds for a
// large drawing - while direct evaluation keeps the import linear.
// Returns null for anything it doesn't cover (arcs, malformed data), so the
// caller can fall back to the slow-but-complete DOM sampler.
function flattenSvgPathData(d)
{
  const tokens = d.match(/[A-Za-z]|-?(?:\d*\.\d+|\d+\.?)(?:[eE][+-]?\d+)?/g);

  if (!tokens || !tokens.length)
    return null;

  const runs = [];

  let run = null,
      i = 0,
      cx = 0, cy = 0,       // current point
      sx = 0, sy = 0,       // subpath start (Z closes back to it)
      pcx = null, pcy = null, // previous curve control point (S/T reflection)
      prevCmd = '';

  function num()
  {
    const v = parseFloat(tokens[i++]);
    return isFinite(v) ? v : NaN;
  }

  function quadTo(qx, qy, x, y)
  {
    const x0 = cx, y0 = cy;

    for (let t = 1; t <= 8; t++)
    {
      const u = t / 8, v = 1 - u;
      run.push({ x: v * v * x0 + 2 * v * u * qx + u * u * x, y: v * v * y0 + 2 * v * u * qy + u * u * y });
    }

    cx = x; cy = y; pcx = qx; pcy = qy;
  }

  function cubicTo(x1, y1, x2, y2, x, y)
  {
    const x0 = cx, y0 = cy;

    for (let t = 1; t <= 10; t++)
    {
      const u = t / 10, v = 1 - u;
      run.push
      (
        {
          x: v * v * v * x0 + 3 * v * v * u * x1 + 3 * v * u * u * x2 + u * u * u * x,
          y: v * v * v * y0 + 3 * v * v * u * y1 + 3 * v * u * u * y2 + u * u * u * y
        }
      );
    }

    cx = x; cy = y; pcx = x2; pcy = y2;
  }

  while (i < tokens.length)
  {
    let cmd = tokens[i];

    if (/^[A-Za-z]$/.test(cmd))
      i++;
    else
      cmd = (prevCmd === 'M') ? 'L' : (prevCmd === 'm') ? 'l' : prevCmd; // implicit command repetition

    const rel = (cmd >= 'a');
    let x, y, x1, y1, x2, y2;

    if (!run && cmd !== 'M' && cmd !== 'm')
      return null; // drawing command before any moveto

    switch (cmd)
    {
      case 'M': case 'm':
        x = num(); y = num();
        if (rel) { x += cx; y += cy; }
        if (run && run.length > 1) runs.push(run);
        run = [{ x: x, y: y }];
        cx = sx = x; cy = sy = y; pcx = pcy = null;
        break;

      case 'L': case 'l':
        x = num(); y = num();
        if (rel) { x += cx; y += cy; }
        run.push({ x: x, y: y }); cx = x; cy = y; pcx = pcy = null;
        break;

      case 'H': case 'h':
        x = num(); if (rel) x += cx;
        run.push({ x: x, y: cy }); cx = x; pcx = pcy = null;
        break;

      case 'V': case 'v':
        y = num(); if (rel) y += cy;
        run.push({ x: cx, y: y }); cy = y; pcx = pcy = null;
        break;

      case 'Q': case 'q':
        x1 = num(); y1 = num(); x = num(); y = num();
        if (rel) { x1 += cx; y1 += cy; x += cx; y += cy; }
        quadTo(x1, y1, x, y);
        break;

      case 'T': case 't':
        x = num(); y = num();
        if (rel) { x += cx; y += cy; }
        x1 = (pcx === null || !/^[QqTt]$/.test(prevCmd)) ? cx : 2 * cx - pcx;
        y1 = (pcy === null || !/^[QqTt]$/.test(prevCmd)) ? cy : 2 * cy - pcy;
        quadTo(x1, y1, x, y);
        break;

      case 'C': case 'c':
        x1 = num(); y1 = num(); x2 = num(); y2 = num(); x = num(); y = num();
        if (rel) { x1 += cx; y1 += cy; x2 += cx; y2 += cy; x += cx; y += cy; }
        cubicTo(x1, y1, x2, y2, x, y);
        break;

      case 'S': case 's':
        x2 = num(); y2 = num(); x = num(); y = num();
        if (rel) { x2 += cx; y2 += cy; x += cx; y += cy; }
        x1 = (pcx === null || !/^[CcSs]$/.test(prevCmd)) ? cx : 2 * cx - pcx;
        y1 = (pcy === null || !/^[CcSs]$/.test(prevCmd)) ? cy : 2 * cy - pcy;
        cubicTo(x1, y1, x2, y2, x, y);
        break;

      case 'Z': case 'z':
        if (!run || (i < tokens.length && !/^[A-Za-z]$/.test(tokens[i])))
          return null; // numbers after Z aren't valid - bail rather than loop
        run.push({ x: sx, y: sy }); cx = sx; cy = sy; pcx = pcy = null;
        break;

      default:
        return null; // arcs and anything unrecognized use the DOM sampler
    }

    if (!run || isNaN(cx) || isNaN(cy))
      return null;

    prevCmd = cmd;
  }

  if (run && run.length > 1)
    runs.push(run);

  return runs.length ? runs : null;
}

// Thins a polyline to roughly evenly spaced points at least 1.5 units apart
// (and at most maxPoints total), always preserving the endpoints.
function thinRunByArcLength(run, maxPoints)
{
  if (run.length < 3)
    return run;

  let len = 0;

  for (let i = 1; i < run.length; i++)
    len += Math.hypot(run[i].x - run[i - 1].x, run[i].y - run[i - 1].y);

  const step = Math.max(1.5, len / maxPoints),
        out = [run[0]];

  let acc = 0;

  for (let i = 1; i < run.length - 1; i++)
  {
    acc += Math.hypot(run[i].x - run[i - 1].x, run[i].y - run[i - 1].y);

    if (acc >= step)
    {
      out.push(run[i]);
      acc = 0;
    }
  }

  out.push(run[run.length - 1]);
  return out;
}

// Converts SVG markup into notebook pen strokes. Every geometry element
// (path, line, rect, circle, ellipse, polyline, polygon) is sampled along its
// outline with getPointAtLength and mapped through its cumulative transform,
// so the result is plain point lists in a common coordinate space - exactly
// the {points:[{x,y,w}]} shape the notebook's own pen produces, which is what
// lets the import render/erase/select like hand-drawn ink afterwards.
// Filled geometry additionally becomes fill:true strokes whose points trace
// the shape's outline(s), with `breaks` marking subpath starts - the renderer
// paints those as even-odd filled polygons, which is what keeps solid-color
// art (e.g. Excalidraw exports, where even pen strokes and text are filled
// paths) looking like the browser's rendering instead of hollow outlines.
// Returns { strokes, w, h } with points normalized so their bounds start at
// (0,0), or null if the markup isn't parseable SVG.
function convertSvgToNotebookStrokes(svgText)
{
  const doc = new DOMParser().parseFromString(svgText, 'image/svg+xml');

  if (doc.querySelector('parsererror') || !doc.documentElement || doc.documentElement.nodeName.toLowerCase() !== 'svg')
    return null;

  // The SVG has to actually be laid out in the live document for the geometry
  // APIs (getTotalLength / getPointAtLength / getScreenCTM) to work, so it's
  // parked far off-screen for the duration of the conversion.
  const host = document.createElement('div');
  host.style.cssText = 'position:fixed;left:-100000px;top:0;pointer-events:none;';

  const svgEl = document.importNode(doc.documentElement, true);
  host.appendChild(svgEl);
  document.body.appendChild(host);

  try
  {
    const rootCTM = svgEl.getScreenCTM();

    if (!rootCTM)
      return null;

    const rootInv = rootCTM.inverse(),
          strokes = [],
          MAX_ELEMENTS = 2000,
          MAX_POINTS_PER_ELEMENT = 600,
          round = function(v){ return Math.round(v * 100) / 100; };

    let els = Array.prototype.slice.call(svgEl.querySelectorAll('path,line,rect,circle,ellipse,polyline,polygon'));

    if (els.length > MAX_ELEMENTS)
      els = els.slice(0, MAX_ELEMENTS);

    // Phase 1: resolve each element's effective style and transform up front,
    // keeping the later geometry work free of layout queries.
    const jobs = [];

    els.forEach(function(el)
    {
      if (typeof el.getTotalLength !== 'function')
        return;

      const cs = getComputedStyle(el);

      if (cs.display === 'none' || cs.visibility === 'hidden')
        return;

      const ctm = el.getScreenCTM();

      if (!ctm)
        return;

      const m = rootInv.multiply(ctm),
            // Determinant-based estimate of the transform's uniform scale, so
            // stroke widths survive scaled/nested groups reasonably.
            mScale = Math.sqrt(Math.abs(m.a * m.d - m.b * m.c)) || 1;

      const tag = el.nodeName.toLowerCase(),
            fillColor = (cs.fill && cs.fill !== 'none' && tag !== 'line') ? cs.fill : null,
            strokeColor = (cs.stroke && cs.stroke !== 'none') ? cs.stroke : null;

      if (!fillColor && !strokeColor)
        return;

      jobs.push({
        el: el,
        m: m,
        width: Math.max(0.5, (parseFloat(cs.strokeWidth) || 1) * mScale),
        fillColor: fillColor,
        fillRule: cs.fillRule,
        strokeColor: strokeColor
      });
    });

    // Phase 2: turn each element into local-space polyline runs, one per
    // subpath - separate runs are what keep glyph holes and disjoint pieces
    // from getting connected. Paths flatten analytically (see
    // flattenSvgPathData for why the DOM sampling APIs are avoided);
    // everything else - rect/circle/ellipse/line/poly*, plus the rare path
    // with arcs - samples via getPointAtLength, which is fine for those
    // cheap single-geometry elements.
    jobs.forEach(function(job)
    {
      let rawRuns = null;

      if (job.el.nodeName.toLowerCase() === 'path')
        rawRuns = flattenSvgPathData(job.el.getAttribute('d') || '');

      if (!rawRuns)
      {
        let len;
        try { len = job.el.getTotalLength(); }
        catch(err) { return; }

        if (!isFinite(len) || len <= 0)
          return;

        const step = Math.max(1.5, len / MAX_POINTS_PER_ELEMENT),
              run = [];

        for (let d = 0; d < len; d += step)
        {
          const p = job.el.getPointAtLength(d);
          run.push({ x: p.x, y: p.y });
        }

        const pEnd = job.el.getPointAtLength(len);
        run.push({ x: pEnd.x, y: pEnd.y });
        rawRuns = [run];
      }

      // Map into the root coordinate space, then thin by arc length so the
      // dense curve flattening doesn't bloat the stored stroke data.
      const m = job.m,
            runs = [];

      rawRuns.forEach(function(rawRun)
      {
        const t = rawRun.map(function(p)
        {
          return { x: round(m.a * p.x + m.c * p.y + m.e), y: round(m.b * p.x + m.d * p.y + m.f) };
        });

        const thinned = thinRunByArcLength(t, MAX_POINTS_PER_ELEMENT);

        if (thinned.length >= 2)
          runs.push(thinned);
      });

      if (!runs.length)
        return;

      // Fill first, outline second, so the outline paints on top just like
      // the browser does.
      if (job.fillColor)
      {
        const points = [], breaks = [];

        runs.forEach(function(run)
        {
          if (points.length)
            breaks.push(points.length);

          run.forEach(function(p){ points.push({ x: p.x, y: p.y, w: 1 }); });
        });

        if (points.length >= 3)
        {
          const stroke = { tool: 'pen', fill: true, color: job.fillColor, baseWidth: 1, points: points, breaks: breaks };

          if (job.fillRule === 'evenodd')
            stroke.fillRule = 'evenodd';

          strokes.push(stroke);
        }
      }

      if (job.strokeColor)
      {
        runs.forEach(function(run)
        {
          strokes.push({
            tool: 'pen',
            color: job.strokeColor,
            baseWidth: job.width,
            points: run.map(function(p){ return { x: p.x, y: p.y, w: job.width }; })
          });
        });
      }
    });

    if (!strokes.length)
      return { strokes: [], w: 1, h: 1 };

    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;

    strokes.forEach(function(s)
    {
      s.points.forEach(function(p)
      {
        if (p.x < minX) minX = p.x;
        if (p.y < minY) minY = p.y;
        if (p.x > maxX) maxX = p.x;
        if (p.y > maxY) maxY = p.y;
      });
    });

    strokes.forEach(function(s)
    {
      s.points.forEach(function(p) { p.x -= minX; p.y -= minY; });
    });

    return { strokes: strokes, w: Math.max(1, maxX - minX), h: Math.max(1, maxY - minY) };
  }
  finally
  {
    document.body.removeChild(host);
  }
}

// Bakes the pending import into the page: placement offset/scale are applied
// to every point, then the strokes join page.strokes as ordinary pen ink
// (undoable, erasable, selectable, saved with the file).
function acceptNotebookVectorImport()
{
  if (!nbImport || !nbData)
    return;

  const imp = nbImport,
        page = nbData.pages[imp.pageIndex];

  if (!page)
  {
    cancelNotebookVectorImport();
    return;
  }

  notebookSnapshotForUndo(imp.pageIndex);

  imp.strokes.forEach(function(s)
  {
    const stroke = {
      tool: s.tool,
      color: s.color,
      baseWidth: Math.max(0.5, s.baseWidth * imp.scale),
      points: s.points.map(function(p)
      {
        return {
          x: imp.x + p.x * imp.scale,
          y: imp.y + p.y * imp.scale,
          w: Math.max(0.5, (p.w || s.baseWidth) * imp.scale)
        };
      })
    };

    if (s.fill)
    {
      stroke.fill = true;

      if (s.fillRule)
        stroke.fillRule = s.fillRule;

      if (s.breaks && s.breaks.length)
        stroke.breaks = s.breaks.slice();
    }

    page.strokes.push(stroke);
  });

  nbImport = null;
  nbActive = null;
  document.getElementById('nb-import-bar').style.display = 'none';

  rebuildNotebookCommittedCache(imp.pageIndex);
  commitNotebookChange();
  redrawNotebookPage(imp.pageIndex);
}

function cancelNotebookVectorImport()
{
  if (!nbImport)
    return;

  const pageIndex = nbImport.pageIndex;

  nbImport = null;
  nbActive = null;
  document.getElementById('nb-import-bar').style.display = 'none';
  redrawNotebookPage(pageIndex);
}

function notebookVectorScaleBy(factor)
{
  if (!nbImport)
    return;

  nbImport.scale = Math.max(0.05, Math.min(20, nbImport.scale * factor));
  redrawNotebookPage(nbImport.pageIndex);
}

// ── Export (composites each page to a PNG, embedded as <img> in the wrapped HTML) ──

async function buildNotebookExportHtml(file)
{
  const data = parseNotebookContent(file.content || '');

  let imagesHtml = '';

  for (let i = 0; i < data.pages.length; i++)
  {
    const page = data.pages[i],
          canvas = document.createElement('canvas');

    canvas.width = page.width;
    canvas.height = page.height;

    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    if (page.type === 'map' && page.imageData)
    {
      try
      {
        await new Promise(function(resolve, reject)
        {
          const mapImg = new Image();
          mapImg.onload = function()
          {
            ctx.drawImage(mapImg, 0, 0, canvas.width, canvas.height);
            resolve();
          };
          mapImg.onerror = reject;
          mapImg.src = page.imageData;
        });
		
		ctx.drawImage(mapImg, 0, 0, canvas.width, canvas.height);
      }
      catch(e)
      {
        console.warn('Notebook export map render error', e);
      }
    }
    else if (page.background.type === 'pdf' && data.pdfData && typeof pdfjsLib !== 'undefined')
    {
      try
      {
        const base64 = data.pdfData.split(',').pop(),
              binary = atob(base64),
              bytes = new Uint8Array(binary.length);

        for (let k = 0; k < binary.length; k++)
          bytes[k] = binary.charCodeAt(k);

        const pdfDoc = await pdfjsLib.getDocument({ data: bytes }).promise,
              pdfPage = await pdfDoc.getPage(page.background.pdfPageIndex + 1),
              scale = canvas.width / pdfPage.getViewport({ scale: 1 }).width,
              viewport = pdfPage.getViewport({ scale: scale });

        await pdfPage.render({ canvasContext: ctx, viewport: viewport }).promise;
      }
      catch(e)
      {
        console.warn('Notebook export PDF render error', e);
      }
    }
    else if (page.background.type === 'ruled' || page.background.type === 'grid')
    {
      renderNotebookPageBackground(page, canvas);
    }

    page.strokes.forEach(function(s){ drawNotebookStroke(ctx, s); });
    page.shapes.forEach(function(s){ drawNotebookShape(ctx, s); });

    page.texts.forEach(function(t)
    {
      ctx.fillStyle = t.color;
      ctx.font = t.fontSize + 'px sans-serif';
      ctx.textBaseline = 'top';

      const words = (t.text || '').split(/\s+/);
      let line = '', y = t.y;
      const lineHeight = t.fontSize * 1.3;

      words.forEach(function(word)
      {
        const test = line ? (line + ' ' + word) : word;

        if (line && (ctx.measureText(test).width > t.width))
        {
          ctx.fillText(line, t.x, y);
          line = word;
          y += lineHeight;
        }
        else
        {
          line = test;
        }
      });

      if (line)
        ctx.fillText(line, t.x, y);
    });

    if (page.type === 'map')
    {
      (page.pins || []).forEach(function(pin)
      {
        const color = pin.color || PIN_COLORS[0],
              px = pin.x, py = pin.y, r = 9;

        ctx.save();
        ctx.translate(px, py - r * 2);
        ctx.beginPath();
        ctx.moveTo(0, r * 2);
        ctx.bezierCurveTo(-r * 1.1, r * 0.6, -r, -r * 0.3, 0, -r * 1.1);
        ctx.bezierCurveTo(r, -r * 0.3, r * 1.1, r * 0.6, 0, r * 2);
        ctx.closePath();
        ctx.fillStyle = color;
        ctx.strokeStyle = 'rgba(0,0,0,0.3)';
        ctx.lineWidth = 0.5;
        ctx.fill();
        ctx.stroke();
        ctx.restore();

        if (pin.label)
        {
          ctx.fillStyle = '#1f2937';
          ctx.font = '13px sans-serif';
          ctx.textBaseline = 'middle';
          ctx.fillText(pin.label, px + r + 4, py - r * 2);
        }
      });
    }

    imagesHtml += '<img class="notebook-page" src="' + canvas.toDataURL('image/png') + '" style="display:block;max-width:100%;margin:0 auto 24px;border:1px solid #ddd;">\n';
  }

  return wrapExportHtml(file.name, '<h1>' + escHtml(file.name) + '</h1>\n' + imagesHtml);
}

// ── SHEET ──
document.addEventListener
(
  'mouseup',
  function()
  {
    isDragging = false;

    if (isRefDragging)
    {
      isRefDragging = false;
      formulaInsertStart = null;
      formulaInsertLength = 0;

      var inp = refDragViaBar ? document.getElementById('formula-bar')
                              : document.getElementById('inp-' + editingCell);
      refDragViaBar = false;

      if (inp)
        inp.focus();
    }

    if (isFillDragging)
    {
      isFillDragging = false;

      if (fillRange)
        performFill();

      clearFillPreview();
      fillRange = null;
      fillSourceBox = null;
    }
  }
);

document.addEventListener
(
  'copy',
  function(copyEvent)
  {
    // Let native copy behavior run while editing (e.g. copying part of a
    // formula's text) or when focus isn't on the grid at all.
    if (editingCell)
      return;

    var active = document.activeElement;

    if (!active || !active.closest('#sheet-body'))
      return;

    if (!selectionAnchor || !selectionEnd)
      return;

    copyEvent.preventDefault();
    copyEvent.clipboardData.setData('text/plain', buildRangeTSV(selectionAnchor, selectionEnd));
  }
);

document.addEventListener
(
  'paste',
  function(pasteEvent)
  {
    // Same guard as the copy handler above: let native paste run while
    // editing a single cell/formula, or when focus isn't on the grid at all
    // (e.g. pasting into the formula bar should just paste text normally).
    if (editingCell)
      return;

    var active = document.activeElement;

    if (!active || !active.closest('#sheet-body'))
      return;

    if (!selectionAnchor || !pasteEvent.clipboardData)
      return;

    var text = pasteEvent.clipboardData.getData('text/plain');

    if (!text)
      return;

    pasteEvent.preventDefault();
    pasteTSVAt(selectionAnchor, text);
  }
);

document.addEventListener
(
  'keydown',
  function(e)
  {
    if (currentAppType !== 'sheet')
      return;

    // Mid-edit (a cell or the formula bar), let native text-undo run instead
    // of jumping back through the sheet-wide stack - same exclusion the
    // copy/paste handlers above already use.
    if (editingCell || document.activeElement === document.getElementById('formula-bar'))
      return;

    if (!(e.ctrlKey || e.metaKey))
      return;

    if (!e.shiftKey && e.key.toLowerCase() === 'z')
    {
      e.preventDefault();
      sheetUndo();
    }
    else if (e.key.toLowerCase() === 'y' || (e.shiftKey && e.key.toLowerCase() === 'z'))
    {
      e.preventDefault();
      sheetRedo();
    }
  }
);

// Writes tab/newline-delimited clipboard text (the format every major
// spreadsheet - Excel, Google Sheets, OnlyOffice - puts on the clipboard for
// a copied range) into the grid starting at anchorRef, growing right/down.
function pasteTSVAt(anchorRef, text)
{
  var anchor = parseName(anchorRef);

  if (!anchor)
    return;

  sheetSnapshotForUndo();

  var rows = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');

  // Spreadsheets terminate clipboard TSV with a trailing newline - without
  // dropping it, the paste would write one extra (empty) row past the data.
  if (rows.length > 1 && rows[rows.length - 1] === '')
    rows.pop();

  var startCol = colIndex(anchor.col),
      startRow = anchor.row,
      lastRef = anchorRef;

  rows.forEach
  (
    function(rowText, r)
    {
      rowText.split('\t').forEach
      (
        function(cellText, c)
        {
          var col = startCol + c,
              row = startRow + r;

          // Pasted blocks push the grid outward instead of being clipped.
          ensureGridFits(row, col);

          if (col >= COLS || row > ROWS)
            return;

          var ref = colName(col) + row;

          sheetData[ref] = cellText;
          lastRef = ref;

          var inp = document.getElementById('inp-' + ref);

          if (inp)
          {
            inp.value = cellText.startsWith('=') ? evalCell(ref, cellText) : cellText;
            inp.classList.toggle('formula-result', cellText.startsWith('='));
          }
        }
      );
    }
  );

  // evaluateFormulas recalculates every formula cell on the sheet (not just
  // the pasted ones) - needed since other formulas may reference cells that
  // just changed - and also refreshes any charts built from this range.
  evaluateFormulas();
  saveSheetToFile();

  selectCell(anchorRef);
  selectionEnd = lastRef;
  renderRangeSelection();
}

function getDisplayValue(ref)
{
  var val = sheetData[ref] || '';
  return val.startsWith('=') ? evalCell(ref, val) : val;
}

function buildRangeTSV(anchorRef, endRef)
{
  var anchor = parseName(anchorRef),
      end = parseName(endRef);

  if (!anchor || !end)
    return '';

  var anchorCol = colIndex(anchor.col),
      endCol = colIndex(end.col),
      rowStart = Math.min(anchor.row, end.row),
      rowEnd = Math.max(anchor.row, end.row),
      colStart = Math.min(anchorCol, endCol),
      colEnd = Math.max(anchorCol, endCol),
      rows = [];

  for (var r = rowStart; r <= rowEnd; r++)
  {
    var cells = [];

    for (var c = colStart; c <= colEnd; c++)
      cells.push(getDisplayValue(colName(c) + r));

    rows.push(cells.join('\t'));
  }

  return rows.join('\n');
}

function buildSheet()
{
  var head = document.getElementById('sheet-head');

  var tableHeadNames ='<tr>'+
                        '<th class="row-header" style="top: 0; left: 0; z-index: 4;">'+
                        '</th>';

  for (var i = 0; i < COLS; i++)
    tableHeadNames +=   '<th class="col-header" id="colh-' + i + '" onmousedown="colHeaderMouseDown(event,' + i + ')">' +
                          colName(i) +
                        '</th>';

  head.innerHTML = tableHeadNames +
                      '</tr>';

  var body = document.getElementById('sheet-body'),
      html = '';

  for (var i = 1; i <= ROWS; i++)
  {
    html += '<tr>' +
              '<th class="row-header" id="rowh-' + i + '" onmousedown="rowHeaderMouseDown(event,' + i + ')" style="position: sticky; left: 0; z-index: 1;">' +
                i +
              '</th>';

    for (var j = 0; j < COLS; j++)
      html += sheetCellHtml(colName(j) + i);

    html += '</tr>';
  }

  document.getElementById('sheet-body').innerHTML = html;

  paintBucketSwatch();
  paintTextSwatch();
}

function sheetCellHtml(ref)
{
  return '<td id="cell-' + ref + '" onmousedown="cellMouseDown(event,\'' + ref + '\')" onmouseenter="cellMouseEnter(\'' + ref + '\')" ondblclick="cellDblClick(\'' + ref + '\')">'+
           '<input class="cell-input" id="inp-' + ref + '" value="" readonly onkeydown="cellKey(event, \'' + ref + '\')" onfocus="onCellFocus(\'' + ref + '\')" oninput="onCellInput(\'' + ref + '\')" onblur="onCellBlur(\'' + ref + '\')">'+
         '</td>';
}

// ── GRID GROWTH ──
// The grid isn't fixed-size: whenever content or the selection gets close to
// an edge, more rows/columns are appended (chunked, to keep DOM churn low).
// Columns run past Z into two-letter names, up to ZZ.

var MAX_SHEET_COLS = 702; // "ZZ"
var GROW_ROW_CHUNK = 12;
var GROW_COL_CHUNK = 4;

function growSheetGrid(targetRows, targetCols)
{
  targetCols = Math.min(targetCols, MAX_SHEET_COLS);
  if (targetRows <= ROWS && targetCols <= COLS)
    return;

  var head = document.getElementById('sheet-head'),
      body = document.getElementById('sheet-body');

  if (!head || !body || !head.rows.length)
  {
    ROWS = Math.max(ROWS, targetRows);
    COLS = Math.max(COLS, targetCols);
    return;
  }

  // New columns first: header cell plus a td on every existing row, so the
  // new rows below are built at the full width.
  if (targetCols > COLS)
  {
    var headRow = head.rows[0];
    for (var j = COLS; j < targetCols; j++)
    {
      var th = document.createElement('th');
      th.className = 'col-header';
      th.id = 'colh-' + j;
      th.setAttribute('onmousedown', 'colHeaderMouseDown(event,' + j + ')');
      th.textContent = colName(j);
      headRow.appendChild(th);
    }
    for (var r = 0; r < body.rows.length; r++)
    {
      var rowNum = r + 1, cells = '';
      for (var jj = COLS; jj < targetCols; jj++)
        cells += sheetCellHtml(colName(jj) + rowNum);
      body.rows[r].insertAdjacentHTML('beforeend', cells);
    }
    COLS = targetCols;
  }

  if (targetRows > ROWS)
  {
    var html = '';
    for (var i = ROWS + 1; i <= targetRows; i++)
    {
      html += '<tr><th class="row-header" id="rowh-' + i + '" onmousedown="rowHeaderMouseDown(event,' + i + ')" style="position: sticky; left: 0; z-index: 1;">' + i + '</th>';
      for (var j2 = 0; j2 < COLS; j2++)
        html += sheetCellHtml(colName(j2) + i);
      html += '</tr>';
    }
    body.insertAdjacentHTML('beforeend', html);
    ROWS = targetRows;
  }
}

/// Grows the grid so (row, col) sits comfortably inside it — being within a
/// couple of cells of an edge already triggers the next chunk.
function ensureGridFits(row, col)
{
  var targetRows = ROWS, targetCols = COLS;
  while (row > targetRows - 3) targetRows += GROW_ROW_CHUNK;
  while (col > targetCols - 3 && targetCols < MAX_SHEET_COLS) targetCols += GROW_COL_CHUNK;
  growSheetGrid(targetRows, targetCols);
}

function ensureGridFitsData(data)
{
  var maxRow = 0, maxCol = 0;
  Object.keys(data || {}).forEach(function(ref)
  {
    var p = parseName(ref);
    if (!p) return;
    if (p.row > maxRow) maxRow = p.row;
    var ci = colIndex(p.col);
    if (ci > maxCol) maxCol = ci;
  });
  if (maxRow || maxCol) ensureGridFits(maxRow, maxCol);
}

// Bijective base-26 column names: 0..25 = A..Z, 26 = AA, 27 = AB, … 701 = ZZ.
function colName(i)
{
    if (i < 26)
      return String.fromCharCode(65 + i);
    return String.fromCharCode(65 + Math.floor(i / 26) - 1) + String.fromCharCode(65 + (i % 26));
}

function colIndex(name)
{
    if (name.length === 1)
      return name.charCodeAt(0) - 65;
    return (name.charCodeAt(0) - 65 + 1) * 26 + (name.charCodeAt(1) - 65);
}

function parseName(cellReference)
{
    // '$' (absolute-reference markers, e.g. $A$1) carry no meaning for lookups,
    // only for how shiftFormulaRefs treats them during a fill-drag.
    var matched = cellReference.match(/^\$?([A-Z]+)\$?(\d+)$/);

    if(!matched)
      return null;

    return {
      col: matched[1],
      row: parseInt(matched[2])
    };
}

function selectCell(name)
{
  document.querySelectorAll('#sheet-body td.active').forEach
  (
    function(activeSheetCell)
    {
      activeSheetCell.classList.remove('active');
    }
  );

  document.querySelectorAll('#sheet-body td.selected').forEach
  (
    function(selectedSheetCell)
    {
      selectedSheetCell.classList.remove('selected');
    }
  );

  var parsedName = parseName(name);

  if(!parsedName)
    return;

  activeCell = {
    row: parsedName.row,
    col: colIndex(parsedName.col)
  };

  // Selection moving near an edge pulls more rows/columns into existence.
  ensureGridFits(activeCell.row, activeCell.col);

  selectionAnchor = name;
  selectionEnd = name;

  var selectedCell = document.getElementById('cell-' + name);

  if(selectedCell)
    selectedCell.classList.add('active');

  document.getElementById('cell-ref').value = name;
  document.getElementById('formula-bar').value = sheetData[name] || '';

  refreshFillHandle();
  updateHeaderHighlights();
  refreshRefHighlights();
}

// Marks the row/column header cells spanning the current selection, so the
// margins always show where you are.
function updateHeaderHighlights()
{
  document.querySelectorAll('#sheet-grid th.hl').forEach
  (
    function(th)
    {
      th.classList.remove('hl');
    }
  );

  if (!selectionAnchor)
    return;

  var box = getSelectionBoundingBox();
  if (!box)
    return;

  for (var r = box.rowStart; r <= box.rowEnd; r++)
  {
    var rh = document.getElementById('rowh-' + r);
    if (rh) rh.classList.add('hl');
  }

  for (var c = box.colStart; c <= box.colEnd; c++)
  {
    var ch = document.getElementById('colh-' + c);
    if (ch) ch.classList.add('hl');
  }
}

function cellMouseDown(mouseEvent, name)
{
  if (mouseEvent.button !== 0)
    return;

  var inp = document.getElementById('inp-' + name);

  // A cell mid-edit (double-clicked, readOnly removed) should keep native
  // text-caret/selection behavior instead of starting a range drag.
  if (inp && !inp.readOnly)
    return;

  // Cross-tab pick mode: a formula is pending on another tab; insert this
  // cell as [Tab name](REF) into it and jump back to where it was typed.
  if (formulaPickCtx && !editingCell)
  {
    mouseEvent.preventDefault();

    var ctx = formulaPickCtx;
    formulaPickCtx = null;

    var srcPage = sheetPages[ctx.pageIdx];
    if (!srcPage) return;

    var refText = '[' + sheetPages[sheetPageIdx].name + '](' + name + ')',
        srcVal = srcPage.data[ctx.ref] || '=',
        insertAt = Math.min(ctx.insertStart, srcVal.length);

    srcPage.data[ctx.ref] = srcVal.slice(0, insertAt) + refText + srcVal.slice(insertAt + ctx.insertLen);

    var caretPos = insertAt + refText.length,
        backRef = ctx.ref;

    switchSheetPage(ctx.pageIdx);

    if (ctx.viaBar)
    {
      // Resume in the formula bar: selectCell repopulates it from
      // sheetData (which now holds the formula with the inserted ref).
      selectCell(backRef);
      var backBar = document.getElementById('formula-bar');
      backBar.focus();
      backBar.setSelectionRange(caretPos, caretPos);
    }
    else
    {
      enterEditMode(backRef);
      var backInp = document.getElementById('inp-' + backRef);
      if (backInp) backInp.setSelectionRange(caretPos, caretPos);
    }
    return;
  }

  // While editing a formula, clicking/dragging other cells is a range
  // picker: it inserts the reference at the caret instead of navigating.
  if (editingCell && (sheetData[editingCell] || '').startsWith('='))
  {
    mouseEvent.preventDefault();

    var editingInput = document.getElementById('inp-' + editingCell);

    isRefDragging = true;
    refDragViaBar = false;
    selectionAnchor = name;
    selectionEnd = name;
    formulaInsertStart = editingInput ? editingInput.selectionStart : 0;
    formulaInsertLength = 0;

    insertFormulaRef(name, name);
    return;
  }

  // The same picker for formula-BAR editing: with the bar focused and
  // holding a formula, clicking cells inserts references instead of moving
  // the selection away (which used to clobber the half-typed formula).
  // Focus is still on the bar at mousedown time; preventDefault keeps it.
  var formulaBar = document.getElementById('formula-bar');
  if (document.activeElement === formulaBar && formulaBar.value.startsWith('='))
  {
    mouseEvent.preventDefault();

    isRefDragging = true;
    refDragViaBar = true;
    selectionAnchor = name;
    selectionEnd = name;
    formulaInsertStart = formulaBar.selectionStart;
    formulaInsertLength = 0;

    insertFormulaRef(name, name);
    return;
  }

  isDragging = true;
  selectCell(name);
}

// Clicking a column letter or row number in the margins: while a formula is
// being typed it inserts the whole-column/row range (A:A / 4:4) at the caret,
// mirroring cellMouseDown's cell picker; otherwise it selects the whole
// column or row.
function colHeaderMouseDown(mouseEvent, col)
{
  var letter = colName(col);
  headerMouseDown(mouseEvent, letter + ':' + letter, letter + '1', letter + ROWS);
}

function rowHeaderMouseDown(mouseEvent, row)
{
  headerMouseDown(mouseEvent, row + ':' + row, 'A' + row, colName(COLS - 1) + row);
}

function headerMouseDown(mouseEvent, rangeText, selAnchor, selEnd)
{
  if (mouseEvent.button !== 0)
    return;

  // Mid-formula in a cell: insert the range at the caret, keep typing.
  if (editingCell && (sheetData[editingCell] || '').startsWith('='))
  {
    mouseEvent.preventDefault();

    var editingInput = document.getElementById('inp-' + editingCell);
    refDragViaBar = false;
    formulaInsertStart = editingInput ? editingInput.selectionStart : 0;
    formulaInsertLength = 0;
    insertFormulaText(rangeText);
    formulaInsertStart = null;
    formulaInsertLength = 0;
    if (editingInput) editingInput.focus();
    return;
  }

  // Mid-formula in the formula bar: same insertion there.
  var formulaBar = document.getElementById('formula-bar');
  if (document.activeElement === formulaBar && formulaBar.value.startsWith('='))
  {
    mouseEvent.preventDefault();

    refDragViaBar = true;
    formulaInsertStart = formulaBar.selectionStart;
    formulaInsertLength = 0;
    insertFormulaText(rangeText);
    refDragViaBar = false;
    formulaInsertStart = null;
    formulaInsertLength = 0;
    formulaBar.focus();
    return;
  }

  // Otherwise: select the whole row/column.
  selectCell(selAnchor);
  selectionEnd = selEnd;
  renderRangeSelection();
}

function cellMouseEnter(name)
{
  if (isFillDragging)
  {
    updateFillPreview(name);
    return;
  }

  if (isRefDragging)
  {
    selectionEnd = name;
    insertFormulaRef(selectionAnchor, name);
    return;
  }

  if (!isDragging)
    return;

  selectionEnd = name;
  renderRangeSelection();
}

function insertFormulaRef(anchorRef, endRef)
{
  var refText = (anchorRef === endRef) ? anchorRef : (anchorRef + ':' + endRef);
  insertFormulaText(refText);
}

function insertFormulaText(refText)
{
  // The picker serves two editing surfaces: the in-cell input, or the
  // formula bar (refDragViaBar) — same insertion logic either way.
  var targetRef = refDragViaBar ? document.getElementById('cell-ref').value : editingCell,
      inp = refDragViaBar ? document.getElementById('formula-bar')
                          : document.getElementById('inp-' + editingCell);

  if (!inp || !targetRef)
    return;

  var before = inp.value.slice(0, formulaInsertStart),
      after = inp.value.slice(formulaInsertStart + formulaInsertLength);

  inp.value = before + refText + after;
  formulaInsertLength = refText.length;

  var caretPos = formulaInsertStart + formulaInsertLength;
  inp.setSelectionRange(caretPos, caretPos);

  sheetData[targetRef] = inp.value;

  // Bar editing: mirror the raw formula into the cell's own input too.
  if (refDragViaBar)
  {
    var cellInp = document.getElementById('inp-' + targetRef);
    if (cellInp) cellInp.value = inp.value;
  }

  evaluateFormulas(targetRef);
  saveSheetToFile();

  renderRangeSelection();
  refreshRefHighlights();
}

function renderRangeSelection()
{
  document.querySelectorAll('#sheet-body td.selected').forEach
  (
    function(selectedSheetCell)
    {
      selectedSheetCell.classList.remove('selected');
    }
  );

  if (!selectionAnchor || !selectionEnd || selectionAnchor === selectionEnd)
  {
    updateHeaderHighlights();
    return;
  }

  var anchor = parseName(selectionAnchor),
      end = parseName(selectionEnd);

  if (!anchor || !end)
    return;

  var anchorCol = colIndex(anchor.col),
      endCol = colIndex(end.col),
      rowStart = Math.min(anchor.row, end.row),
      rowEnd = Math.max(anchor.row, end.row),
      colStart = Math.min(anchorCol, endCol),
      colEnd = Math.max(anchorCol, endCol);

  for (var r = rowStart; r <= rowEnd; r++)
    for (var c = colStart; c <= colEnd; c++)
    {
      var cellRef = colName(c) + r,
          td = document.getElementById('cell-' + cellRef);

      if (td && cellRef !== selectionAnchor)
        td.classList.add('selected');
    }

  document.getElementById('cell-ref').value = selectionAnchor + ':' + selectionEnd;

  refreshFillHandle();
  updateHeaderHighlights();
}

// ── FILL HANDLE (drag-fill, like Excel/Sheets) ──
function refreshFillHandle()
{
  if (editingCell)
    removeFillHandle();
  else
    updateFillHandlePosition();
}

function updateFillHandlePosition()
{
  var endRef = selectionEnd || selectionAnchor;

  if (!endRef)
  {
    removeFillHandle();
    return;
  }

  var td = document.getElementById('cell-' + endRef);

  if (!td)
  {
    removeFillHandle();
    return;
  }

  var handle = document.getElementById('fill-handle');

  if (!handle)
  {
    handle = document.createElement('div');
    handle.id = 'fill-handle';
    handle.className = 'fill-handle';
    handle.addEventListener('mousedown', fillHandleMouseDown);
  }

  td.appendChild(handle);
}

function removeFillHandle()
{
  var handle = document.getElementById('fill-handle');

  if (handle)
    handle.remove();
}

// ── FORMULA REFERENCE HIGHLIGHTS ──
// While a formula is being edited (in-cell or in the bar) — or a cell holding
// one is selected — every cell/range reference in it gets a colored box on
// the grid. The boxes are live: dragging an edge moves the reference,
// dragging a range's corner handle resizes it, and the formula text updates
// as you go. Only box edges/handles take pointer events, so the cells inside
// stay clickable.

var REF_BOX_COLORS = ['#4a7fa8', '#c0574a', '#5a9a6e', '#9d6fd1', '#d18b3f', '#3fa7b0', '#d4a843'];

var refHighlights = [];  // tokens of the highlighted formula, in text order
var refBoxDrag = null;   // active box drag: {index, mode, src, lastRow, lastCol, anchorRow, anchorCol}

// Where the highlighted formula lives right now: the in-cell edit, the
// formula bar, or (fallback) the selected cell's stored value.
function refHighlightSource()
{
  if (editingCell)
  {
    var inp = document.getElementById('inp-' + editingCell);
    if (inp && !inp.readOnly)
      return { ref: editingCell, inp: inp, viaBar: false, editing: true };
  }

  var bar = document.getElementById('formula-bar');
  if (document.activeElement === bar && bar.value.startsWith('='))
    return { ref: document.getElementById('cell-ref').value, inp: bar, viaBar: true, editing: true };

  if (selectionAnchor && (!selectionEnd || selectionEnd === selectionAnchor))
    return { ref: selectionAnchor, inp: null, viaBar: false, editing: false };

  return null;
}

function refSourceText(src)
{
  return src.inp ? src.inp.value : (sheetData[src.ref] || '');
}

// Classifies one matched reference and normalizes it into a grid box
// (r1<=r2, c1<=c2). d holds the '$' markers so moving/resizing keeps them.
function refTokenInfo(text)
{
  var m;

  if ((m = text.match(/^(\$?)([A-Z]{1,2})(\$?)(\d+):(\$?)([A-Z]{1,2})(\$?)(\d+)$/)))
  {
    var ca = colIndex(m[2]), cb = colIndex(m[6]), ra = +m[4], rb = +m[8];
    return { kind: 'range', c1: Math.min(ca, cb), c2: Math.max(ca, cb),
             r1: Math.min(ra, rb), r2: Math.max(ra, rb), d: [m[1], m[3], m[5], m[7]] };
  }

  if ((m = text.match(/^(\$?)([A-Z]{1,2}):(\$?)([A-Z]{1,2})$/)))
  {
    var cc = colIndex(m[2]), cd = colIndex(m[4]);
    return { kind: 'col', c1: Math.min(cc, cd), c2: Math.max(cc, cd), d: [m[1], m[3]] };
  }

  if ((m = text.match(/^(\$?)(\d+):(\$?)(\d+)$/)))
  {
    var re = +m[2], rf = +m[4];
    if (re < 1 || rf < 1) return null;
    return { kind: 'row', r1: Math.min(re, rf), r2: Math.max(re, rf), d: [m[1], m[3]] };
  }

  if ((m = text.match(/^(\$?)([A-Z]{1,2})(\$?)(\d+)$/)))
    return { kind: 'cell', c1: colIndex(m[2]), c2: colIndex(m[2]), r1: +m[4], r2: +m[4], d: [m[1], m[3]] };

  return null;
}

function refTokenText(tok)
{
  var d = tok.d;

  if (tok.kind === 'cell') return d[0] + colName(tok.c1) + d[1] + tok.r1;
  if (tok.kind === 'col')  return d[0] + colName(tok.c1) + ':' + d[1] + colName(tok.c2);
  if (tok.kind === 'row')  return d[0] + tok.r1 + ':' + d[1] + tok.r2;

  return d[0] + colName(tok.c1) + d[1] + tok.r1 + ':' + d[2] + colName(tok.c2) + d[3] + tok.r2;
}

// Parses cell refs (A3), ranges (A1:B3) and whole-column/row ranges (A:A,
// 4:4) out of a formula's text, with their offsets, so each can later be
// replaced in place. Quoted strings and cross-tab [Tab](A1) links are masked
// (same length, so offsets hold) to keep their contents from reading as refs.
function parseFormulaRefTokens(text)
{
  if (!text || !text.startsWith('='))
    return [];

  var masked = text
    .replace(/"[^"]*"|'[^']*'/g, function(m){ return m.replace(/./g, ' '); })
    .replace(/\[[^\]]*\]\([^)]*\)/g, function(m){ return m.replace(/./g, ' '); });

  var re = /(?<![A-Za-z0-9_$:])(?:\$?[A-Z]{1,2}\$?\d+:\$?[A-Z]{1,2}\$?\d+|\$?[A-Z]{1,2}:\$?[A-Z]{1,2}|\$?\d+:\$?\d+|\$?[A-Z]{1,2}\$?\d+)(?![A-Za-z0-9_$:])/g,
      tokens = [],
      m;

  while ((m = re.exec(masked)))
  {
    var info = refTokenInfo(m[0]);
    if (!info) continue;

    info.start = m.index;
    info.end = m.index + m[0].length;
    info.text = m[0];
    tokens.push(info);
  }

  return tokens;
}

// Pixel geometry of a token's box, relative to the ref layer (which sits at
// the grid's content origin). Whole-column/row boxes span the current grid.
function refBoxGeometry(tok)
{
  var r1 = tok.kind === 'col' ? 1 : tok.r1,
      r2 = tok.kind === 'col' ? ROWS : Math.min(tok.r2, ROWS),
      c1 = tok.kind === 'row' ? 0 : tok.c1,
      c2 = tok.kind === 'row' ? COLS - 1 : Math.min(tok.c2, COLS - 1);

  if (r1 > ROWS || c1 > COLS - 1)
    return null;

  var layer = document.getElementById('sheet-ref-layer'),
      tdA = document.getElementById('cell-' + colName(c1) + r1),
      tdB = document.getElementById('cell-' + colName(c2) + r2);

  // Cells hidden by a merge report zero size; skip rather than draw a
  // zero-anchored box.
  if (!layer || !tdA || !tdB || !tdA.offsetWidth || !tdB.offsetWidth)
    return null;

  var la = layer.getBoundingClientRect(),
      ra = tdA.getBoundingClientRect(),
      rb = tdB.getBoundingClientRect();

  return { left: ra.left - la.left, top: ra.top - la.top,
           width: rb.right - ra.left, height: rb.bottom - ra.top };
}

function buildRefBoxElement(tok, index)
{
  var geo = refBoxGeometry(tok);
  if (!geo)
    return null;

  var color = REF_BOX_COLORS[index % REF_BOX_COLORS.length],
      box = document.createElement('div');

  box.className = 'ref-box';
  box.style.left = geo.left + 'px';
  box.style.top = geo.top + 'px';
  box.style.width = geo.width + 'px';
  box.style.height = geo.height + 'px';
  box.style.borderColor = color;
  box.style.background = color + '14';

  ['top', 'right', 'bottom', 'left'].forEach(function(side)
  {
    var edge = document.createElement('div');
    edge.className = 'ref-box-edge ' + side;
    edge.addEventListener('mousedown', function(e){ startRefBoxDrag(e, index, 'move'); });
    box.appendChild(edge);
  });

  if (tok.kind !== 'cell')
  {
    var handle = document.createElement('div');
    handle.className = 'ref-box-size';
    handle.style.background = color;
    handle.addEventListener('mousedown', function(e){ startRefBoxDrag(e, index, 'resize'); });
    box.appendChild(handle);
  }

  return box;
}

function refreshRefHighlights()
{
  var layer = document.getElementById('sheet-ref-layer');
  if (!layer)
    return;

  layer.innerHTML = '';
  refHighlights = [];

  var src = refHighlightSource();
  if (!src)
    return;

  refHighlights = parseFormulaRefTokens(refSourceText(src));

  refHighlights.forEach(function(tok, i)
  {
    var el = buildRefBoxElement(tok, i);
    if (el) layer.appendChild(el);
  });
}

// The grid cell under a viewport point, looked up through whatever element
// is there (usually the cell's own input).
function sheetCellFromPoint(x, y)
{
  var el = document.elementFromPoint(x, y);

  while (el && el !== document.body)
  {
    if (el.id && el.id.indexOf('cell-') === 0)
    {
      var p = parseName(el.id.slice(5));
      if (p) return { row: p.row, col: colIndex(p.col) };
    }
    el = el.parentElement;
  }

  return null;
}

function startRefBoxDrag(mouseEvent, index, mode)
{
  if (mouseEvent.button !== 0)
    return;

  var src = refHighlightSource(),
      tok = refHighlights[index];

  if (!src || !tok)
    return;

  // Keeps focus (and the mid-edit state) in whichever input holds the formula.
  mouseEvent.preventDefault();
  mouseEvent.stopPropagation();

  // Dragging on a committed (not-being-edited) formula rewrites it in place,
  // so it gets its own undo step; an active edit already snapshotted.
  if (!src.editing)
    sheetSnapshotForUndo();

  var cell = sheetCellFromPoint(mouseEvent.clientX, mouseEvent.clientY);

  refBoxDrag = {
    index: index,
    mode: mode,
    src: src,
    lastRow: cell ? cell.row : (tok.kind === 'col' ? 1 : tok.r1),
    lastCol: cell ? cell.col : (tok.kind === 'row' ? 0 : tok.c1),
    anchorRow: tok.kind === 'col' ? 1 : tok.r1,
    anchorCol: tok.kind === 'row' ? 0 : tok.c1
  };

  var layer = document.getElementById('sheet-ref-layer');
  if (layer) layer.classList.add('dragging');
}

// Writes the (moved/resized) token back into the formula text and refreshes
// everything that mirrors it.
function applyRefTokenEdit(tok)
{
  var src = refBoxDrag.src,
      text = refSourceText(src),
      newText = refTokenText(tok),
      full = text.slice(0, tok.start) + newText + text.slice(tok.end);

  if (src.inp)
  {
    src.inp.value = full;
    var caret = tok.start + newText.length;
    src.inp.setSelectionRange(caret, caret);
  }

  sheetData[src.ref] = full;

  // Bar editing mirrors the raw text into the cell input; the committed-cell
  // case mirrors it into the bar instead.
  if (src.viaBar)
  {
    var cellInp = document.getElementById('inp-' + src.ref);
    if (cellInp) cellInp.value = full;
  }
  if (!src.editing)
    document.getElementById('formula-bar').value = full;

  evaluateFormulas(src.editing ? src.ref : undefined);
  saveSheetToFile();

  // Re-parse: offsets after the replaced token have shifted.
  refreshRefHighlights();
}

document.addEventListener
(
  'mousemove',
  function(mouseEvent)
  {
    if (!refBoxDrag)
      return;

    var cell = sheetCellFromPoint(mouseEvent.clientX, mouseEvent.clientY);
    if (!cell)
      return;

    var tok = refHighlights[refBoxDrag.index];
    if (!tok)
    {
      refBoxDrag = null;
      return;
    }

    var changed = false;

    if (refBoxDrag.mode === 'move')
    {
      var dr = cell.row - refBoxDrag.lastRow,
          dc = cell.col - refBoxDrag.lastCol;

      if (tok.kind === 'col') dr = 0;
      if (tok.kind === 'row') dc = 0;

      // Clamp so the box never slides off the top/left of the sheet.
      if (tok.r1 !== undefined) dr = Math.max(dr, 1 - tok.r1);
      if (tok.c1 !== undefined) dc = Math.max(dc, -tok.c1);

      if (dr || dc)
      {
        if (tok.r1 !== undefined) { tok.r1 += dr; tok.r2 += dr; }
        if (tok.c1 !== undefined) { tok.c1 += dc; tok.c2 += dc; }
        changed = true;
      }
    }
    else // resize: the corner opposite the handle stays anchored
    {
      var nr1 = tok.r1, nr2 = tok.r2, nc1 = tok.c1, nc2 = tok.c2;

      if (tok.kind !== 'col')
      {
        nr1 = Math.min(refBoxDrag.anchorRow, cell.row);
        nr2 = Math.max(refBoxDrag.anchorRow, cell.row);
      }
      if (tok.kind !== 'row')
      {
        nc1 = Math.min(refBoxDrag.anchorCol, cell.col);
        nc2 = Math.max(refBoxDrag.anchorCol, cell.col);
      }

      if (nr1 !== tok.r1 || nr2 !== tok.r2 || nc1 !== tok.c1 || nc2 !== tok.c2)
      {
        tok.r1 = nr1; tok.r2 = nr2; tok.c1 = nc1; tok.c2 = nc2;
        changed = true;
      }
    }

    refBoxDrag.lastRow = cell.row;
    refBoxDrag.lastCol = cell.col;

    if (changed)
      applyRefTokenEdit(tok);
  }
);

document.addEventListener
(
  'mouseup',
  function()
  {
    if (!refBoxDrag)
      return;

    var src = refBoxDrag.src;
    refBoxDrag = null;

    var layer = document.getElementById('sheet-ref-layer');
    if (layer) layer.classList.remove('dragging');

    if (src.editing && src.inp)
      src.inp.focus();
  }
);

// ── DRAG AUTO-SCROLL ──
// While any grid drag is active (selection, formula ref picking, fill,
// reference-box move/resize, chart move/resize), the sheet scrolls by itself
// when the pointer nears the container's edges — the closer to the edge, the
// faster. Each scroll step retargets the drag, since the grid moved under a
// pointer that didn't.

var AUTOSCROLL_ZONE = 48; // px from an edge where scrolling kicks in
var AUTOSCROLL_MAX = 28;  // px per frame at (or past) the very edge

var autoScrollMouse = null;
var autoScrollRaf = null;

function sheetDragActive()
{
  return isDragging || isRefDragging || isFillDragging || !!refBoxDrag || !!sheetChartDrag;
}

document.addEventListener
(
  'mousemove',
  function(mouseEvent)
  {
    if (!sheetDragActive())
    {
      autoScrollMouse = null;
      return;
    }

    autoScrollMouse = { x: mouseEvent.clientX, y: mouseEvent.clientY };

    if (!autoScrollRaf)
      autoScrollRaf = requestAnimationFrame(autoScrollStep);
  }
);

function autoScrollSpeed(pos, lo, hi)
{
  var depth;

  if (pos < lo + AUTOSCROLL_ZONE)
  {
    depth = Math.min(1, (lo + AUTOSCROLL_ZONE - pos) / AUTOSCROLL_ZONE);
    return -Math.ceil(depth * AUTOSCROLL_MAX);
  }

  if (pos > hi - AUTOSCROLL_ZONE)
  {
    depth = Math.min(1, (pos - (hi - AUTOSCROLL_ZONE)) / AUTOSCROLL_ZONE);
    return Math.ceil(depth * AUTOSCROLL_MAX);
  }

  return 0;
}

function autoScrollStep()
{
  autoScrollRaf = null;

  if (!sheetDragActive() || !autoScrollMouse)
    return;

  var container = document.getElementById('sheet-container');
  if (!container)
    return;

  var rect = container.getBoundingClientRect(),
      dx = autoScrollSpeed(autoScrollMouse.x, rect.left, rect.right),
      dy = autoScrollSpeed(autoScrollMouse.y, rect.top, rect.bottom);

  if (dx || dy)
  {
    var beforeLeft = container.scrollLeft,
        beforeTop = container.scrollTop;

    container.scrollLeft += dx;
    container.scrollTop += dy;

    var movedX = container.scrollLeft - beforeLeft,
        movedY = container.scrollTop - beforeTop;

    if (movedX || movedY)
      retargetDragAtPointer(movedX, movedY);
  }

  // The retarget above can dispatch a synthetic mousemove, whose listener
  // already scheduled the next step — guard so each frame runs this once,
  // not an ever-doubling number of times.
  if (!autoScrollRaf)
    autoScrollRaf = requestAnimationFrame(autoScrollStep);
}

function retargetDragAtPointer(dx, dy)
{
  // Chart moves/resizes are anchored to viewport coordinates; the content
  // shifted under the stationary pointer, so shift the anchor with it.
  if (sheetChartDrag)
  {
    if (sheetChartDrag.kind === 'move')
    {
      sheetChartDrag.origX += dx;
      sheetChartDrag.origY += dy;
    }
    else
    {
      sheetChartDrag.origWidth += dx;
      sheetChartDrag.origHeight += dy;
    }
  }

  // Ref-box and chart drags do their work in document mousemove handlers;
  // feed them a synthetic move at the pointer's (unchanged) position.
  if (refBoxDrag || sheetChartDrag)
  {
    document.dispatchEvent(new MouseEvent('mousemove', { bubbles: true, clientX: autoScrollMouse.x, clientY: autoScrollMouse.y, button: 0 }));
    return;
  }

  // Selection, formula-ref picking, and fill drags track the hovered cell.
  var cell = sheetCellFromPoint(autoScrollMouse.x, autoScrollMouse.y);

  if (cell && (isDragging || isRefDragging || isFillDragging))
    cellMouseEnter(colName(cell.col) + cell.row);
}

function getSelectionBoundingBox()
{
  var anchor = parseName(selectionAnchor),
      end = parseName(selectionEnd || selectionAnchor);

  if (!anchor)
    return null;

  if (!end)
    end = anchor;

  var anchorCol = colIndex(anchor.col),
      endCol = colIndex(end.col);

  return  {
            rowStart: Math.min(anchor.row, end.row),
            rowEnd: Math.max(anchor.row, end.row),
            colStart: Math.min(anchorCol, endCol),
            colEnd: Math.max(anchorCol, endCol)
          };
}

function fillHandleMouseDown(mouseEvent)
{
  mouseEvent.preventDefault();
  mouseEvent.stopPropagation();

  fillSourceBox = getSelectionBoundingBox();
  isFillDragging = true;
  fillRange = null;
}

function updateFillPreview(targetRef)
{
  var target = parseName(targetRef);

  if (!target || !fillSourceBox)
    return;

  var targetCol = colIndex(target.col),
      box = fillSourceBox,
      vDist =   target.row < box.rowStart
                ?
                  box.rowStart - target.row
                :
                  (target.row > box.rowEnd ? target.row - box.rowEnd : 0),
      hDist =   targetCol < box.colStart
                ?
                  box.colStart - targetCol
                :
                  (targetCol > box.colEnd ? targetCol - box.colEnd : 0);

  if (!vDist && !hDist)
    fillRange = null;
  else if (vDist >= hDist)
    fillRange =   (target.row > box.rowEnd)
                  ?
                    { rowStart: box.rowEnd + 1, rowEnd: target.row, colStart: box.colStart, colEnd: box.colEnd }
                  :
                    { rowStart: target.row, rowEnd: box.rowStart - 1, colStart: box.colStart, colEnd: box.colEnd };
  else
    fillRange =   (targetCol > box.colEnd)
                  ?
                    { rowStart: box.rowStart, rowEnd: box.rowEnd, colStart: box.colEnd + 1, colEnd: targetCol }
                  :
                    { rowStart: box.rowStart, rowEnd: box.rowEnd, colStart: targetCol, colEnd: box.colStart - 1 };

  renderFillPreview();
}

function renderFillPreview()
{
  clearFillPreview();

  if (!fillRange)
    return;

  for (var r = fillRange.rowStart; r <= fillRange.rowEnd; r++)
    for (var c = fillRange.colStart; c <= fillRange.colEnd; c++)
    {
      var td = document.getElementById('cell-' + colName(c) + r);

      if (td)
        td.classList.add('fill-preview');
    }
}

function clearFillPreview()
{
  document.querySelectorAll('#sheet-body td.fill-preview').forEach
  (
    function(td)
    {
      td.classList.remove('fill-preview');
    }
  );
}

// Shifts every cell reference in a formula by a row/column delta - the same
// relative-reference adjustment Excel/Sheets apply when you drag-fill. A '$'
// before the column or row (e.g. $A$1, A$1, $A1) locks that axis in place.
function shiftFormulaRefs(formulaText, rowDelta, colDelta)
{
  return formulaText.replace
  (
    /(?<![A-Za-z])(\$?)([A-Z]{1,2})(\$?)(\d+)/g,
    function(match, colAbs, col, rowAbs, row)
    {
      var newCol = colAbs ? colIndex(col) : colIndex(col) + colDelta,
          newRow = rowAbs ? parseInt(row, 10) : parseInt(row, 10) + rowDelta;

      if (newCol < 0 || newCol >= COLS || newRow < 1 || newRow > ROWS)
        return match;

      return colAbs + colName(newCol) + rowAbs + newRow;
    }
  );
}

// Optional companion to shiftFormulaRefs: when the "increment numbers on
// fill" toggle is on, plain top-level numeric literals in the formula shift
// by the same delta, so =A1+1 filled down becomes =A2+2, =A3+3, etc. Numbers
// inside a function call's parens (ROUND(A1,2)'s 2, LOG10(100)'s 100) are
// left alone, since those are arguments, not series values - and so are
// cell-reference/function-name digits (A1, LOG10), which are never "bare".
function shiftFormulaLiterals(formulaText, delta)
{
  if (!delta)
    return formulaText;

  var result = '',
      depth = 0,
      i = 0;

  while (i < formulaText.length)
  {
    var ch = formulaText[i];

    if (ch === '(')
    {
      depth++;
      result += ch;
      i++;
      continue;
    }

    if (ch === ')')
    {
      depth--;
      result += ch;
      i++;
      continue;
    }

    var prevChar = i > 0 ? formulaText[i - 1] : '';

    if (depth === 0 && /\d/.test(ch) && !/[A-Za-z0-9.$]/.test(prevChar))
    {
      var match = /^\d+(?:\.\d+)?/.exec(formulaText.slice(i))[0];

      result += (parseFloat(match) + delta).toString();
      i += match.length;
      continue;
    }

    result += ch;
    i++;
  }

  return result;
}

function toggleIncrementLiterals(checked)
{
  incrementLiteralsOnFill = checked;
}

function refreshCellDisplay(ref)
{
  var inp = document.getElementById('inp-' + ref),
      v = sheetData[ref] || '';

  if (!inp)
    return;

  inp.value = v.startsWith('=') ? evalCell(ref, v) : v;
  inp.classList.toggle('formula-result', v.startsWith('='));
}

function performFill()
{
  sheetSnapshotForUndo();

  var box = fillSourceBox,
      range = fillRange,
      srcRows = box.rowEnd - box.rowStart + 1,
      srcCols = box.colEnd - box.colStart + 1;

  for (var r = range.rowStart; r <= range.rowEnd; r++)
  {
    for (var c = range.colStart; c <= range.colEnd; c++)
    {
      // Cycle back through the source range so a multi-cell source pattern
      // repeats across the filled region, same as Excel/Sheets.
      var srcRow = box.rowStart + (((r - box.rowStart) % srcRows) + srcRows) % srcRows,
          srcCol = box.colStart + (((c - box.colStart) % srcCols) + srcCols) % srcCols,
          srcRef = colName(srcCol) + srcRow,
          destRef = colName(c) + r,
          srcVal = sheetData[srcRef] || '',
          destVal = srcVal;

      if (srcVal.startsWith('='))
      {
        destVal = shiftFormulaRefs(srcVal, r - srcRow, c - srcCol);

        if (incrementLiteralsOnFill)
          destVal = shiftFormulaLiterals(destVal, (r - srcRow) + (c - srcCol));
      }

      sheetData[destRef] = destVal;

      refreshCellDisplay(destRef);
    }
  }

  selectionAnchor = colName(Math.min(box.colStart, range.colStart)) + Math.min(box.rowStart, range.rowStart);
  selectionEnd = colName(Math.max(box.colEnd, range.colEnd)) + Math.max(box.rowEnd, range.rowEnd);

  renderRangeSelection();
  evaluateFormulas();
  saveSheetToFile();
}

function enterEditMode(name, freshValue)
{
  var inp = document.getElementById('inp-' + name);

  if (!inp)
    return;

  sheetSnapshotForUndo();

  editingCell = name;
  inp.readOnly = false;
  inp.classList.remove('formula-result');

  if (freshValue === undefined)
    inp.value = sheetData[name] || '';
  else
  {
    inp.value = freshValue;
    sheetData[name] = freshValue;
  }

  selectCell(name);
  inp.focus();
}

function cellDblClick(name)
{
  enterEditMode(name);

  var inp = document.getElementById('inp-' + name),
      caretPos = inp.value.length;

  inp.setSelectionRange(caretPos, caretPos);
}

function isTypingKey(keyEvent)
{
  return  keyEvent.key.length === 1 &&
          !keyEvent.ctrlKey &&
          !keyEvent.metaKey &&
          !keyEvent.altKey;
}

// ── FUNCTION SUGGESTIONS (autocomplete dropdown while editing a formula) ──

// Scans backward from the caret for a run of letters - that's the partial
// function name currently being typed, ignoring everything before it
// (cell refs, numbers, operators).
function getTypedWordBounds(text, caretPos)
{
  var start = caretPos;

  while (start > 0 && /[A-Za-z]/.test(text[start - 1]))
    start--;

  return { start: start, end: caretPos, word: text.slice(start, caretPos) };
}

function updateFunctionSuggestions(name)
{
  var inp = document.getElementById('inp-' + name);

  if (!inp || inp.readOnly)
  {
    hideFunctionSuggestions();
    return;
  }

  var text = inp.value,
      caretPos = inp.selectionStart;

  if (!text.startsWith('='))
  {
    hideFunctionSuggestions();
    return;
  }

  var bounds = getTypedWordBounds(text, caretPos);

  // Show the full list right as '=' is typed (caret at 1, nothing after it
  // yet); otherwise only while there's a letter run to filter against.
  if (!bounds.word && caretPos !== 1)
  {
    hideFunctionSuggestions();
    return;
  }

  var filter = bounds.word.toUpperCase(),
      matches = ALL_SHEET_FUNCTIONS.filter
      (
        function(fn)
        {
          return fn.bareName.indexOf(filter) === 0;
        }
      )
      .slice(0, 8);

  if (!matches.length)
  {
    hideFunctionSuggestions();
    return;
  }

  renderFunctionSuggestions(name, matches, bounds);
}

function renderFunctionSuggestions(name, matches, bounds)
{
  var td = document.getElementById('cell-' + name);

  if (!td)
    return;

  var dropdown = document.getElementById('function-suggest-dropdown');

  if (!dropdown)
  {
    dropdown = document.createElement('div');
    dropdown.id = 'function-suggest-dropdown';
    dropdown.className = 'function-suggest-dropdown';
  }

  dropdown.innerHTML = matches.map
  (
    function(fn, i)
    {
      return  '<div class="function-suggest-item' + (i === 0 ? ' active' : '') + '" onmousedown="pickFunctionSuggestion(event,' + i + ')">'+
                '<span class="function-suggest-name">' + escHtml(fn.signature) + '</span>'+
                '<span class="function-suggest-desc">' + escHtml(fn.desc) + '</span>'+
              '</div>';
    }
  ).join('');

  td.appendChild(dropdown);

  functionSuggestState = { name: name, matches: matches, bounds: bounds, activeIndex: 0 };
}

function hideFunctionSuggestions()
{
  var dropdown = document.getElementById('function-suggest-dropdown');

  if (dropdown)
    dropdown.remove();

  functionSuggestState = null;
}

function moveFunctionSuggestion(delta)
{
  if (!functionSuggestState)
    return;

  var count = functionSuggestState.matches.length;

  functionSuggestState.activeIndex = (functionSuggestState.activeIndex + delta + count) % count;

  document.querySelectorAll('.function-suggest-item').forEach
  (
    function(item, i)
    {
      item.classList.toggle('active', i === functionSuggestState.activeIndex);
    }
  );
}

function pickFunctionSuggestion(mouseEvent, index)
{
  mouseEvent.preventDefault();

  if (!functionSuggestState)
    return;

  insertFunctionSuggestion(functionSuggestState.matches[index]);
}

function insertFunctionSuggestion(fn)
{
  var state = functionSuggestState;

  if (!state || !fn)
    return;

  var inp = document.getElementById('inp-' + state.name);

  if (!inp)
    return;

  var insertText = fn.bareName + '(',
      before = inp.value.slice(0, state.bounds.start),
      after = inp.value.slice(state.bounds.end);

  inp.value = before + insertText + after;

  var caretPos = before.length + insertText.length;

  inp.focus();
  inp.setSelectionRange(caretPos, caretPos);

  sheetData[state.name] = inp.value;
  evaluateFormulas(state.name);
  saveSheetToFile();

  hideFunctionSuggestions();
}

function onCellFocus(name)
{
  selectCell(name);
}

function onCellInput(name)
{
  sheetData[name] = document.getElementById('inp-' + name).value;
  evaluateFormulas(name);
  saveSheetToFile();
  updateFunctionSuggestions(name);
  refreshRefHighlights();
}

function onCellBlur(name)
{
  var inp = document.getElementById('inp-' + name);

  if (!inp)
    return;

  var val = sheetData[name] || '';

  inp.value = val.startsWith('=') ? evalCell(name, val) : val;
  inp.classList.toggle('formula-result', val.startsWith('='));
  inp.readOnly = true;

  if (editingCell === name)
    editingCell = null;

  hideFunctionSuggestions();
  refreshFillHandle();
  refreshRefHighlights();
}

function getSelectionRefs(anchorRef, endRef)
{
  var anchor = parseName(anchorRef),
      end = parseName(endRef || anchorRef);

  if (!anchor)
    return [];

  if (!end)
    end = anchor;

  var anchorCol = colIndex(anchor.col),
      endCol = colIndex(end.col),
      rowStart = Math.min(anchor.row, end.row),
      rowEnd = Math.max(anchor.row, end.row),
      colStart = Math.min(anchorCol, endCol),
      colEnd = Math.max(anchorCol, endCol),
      refs = [];

  for (var r = rowStart; r <= rowEnd; r++)
    for (var c = colStart; c <= colEnd; c++)
      refs.push(colName(c) + r);

  return refs;
}

function deleteSelectedCells()
{
  var refs = getSelectionRefs(selectionAnchor, selectionEnd);

  if (!refs.length)
    return;

  sheetSnapshotForUndo();

  refs.forEach
  (
    function(ref)
    {
      delete sheetData[ref];

      var inp = document.getElementById('inp-' + ref);

      if (inp)
      {
        inp.value = '';
        inp.classList.remove('formula-result');
      }
    }
  );

  document.getElementById('formula-bar').value = '';
  evaluateFormulas();
  saveSheetToFile();
  refreshRefHighlights();
}

function cellKey(keyEvent, name)
{
  var inp = document.getElementById('inp-' + name);

  if (functionSuggestState && functionSuggestState.name === name)
  {
    if (keyEvent.key === 'ArrowDown')
    {
      keyEvent.preventDefault();
      moveFunctionSuggestion(1);
      return;
    }

    if (keyEvent.key === 'ArrowUp')
    {
      keyEvent.preventDefault();
      moveFunctionSuggestion(-1);
      return;
    }

    if (keyEvent.key === 'Enter' || keyEvent.key === 'Tab')
    {
      keyEvent.preventDefault();
      insertFunctionSuggestion(functionSuggestState.matches[functionSuggestState.activeIndex]);
      return;
    }

    if (keyEvent.key === 'Escape')
    {
      keyEvent.preventDefault();
      hideFunctionSuggestions();
      return;
    }
  }

  if (inp && inp.readOnly)
  {
    if (isTypingKey(keyEvent))
    {
      enterEditMode(name, '');
      return;
    }

    if (keyEvent.key === 'Delete')
    {
      keyEvent.preventDefault();
      deleteSelectedCells();
      return;
    }

    if (keyEvent.key !== 'Enter' && keyEvent.key !== 'Tab' &&
        keyEvent.key !== 'ArrowUp' && keyEvent.key !== 'ArrowDown' &&
        keyEvent.key !== 'ArrowLeft' && keyEvent.key !== 'ArrowRight')
      return;
  }
  else if (keyEvent.key !== 'Enter' && keyEvent.key !== 'Tab')
  {
    // Editing this cell: let arrows/backspace/etc. move the text caret natively.
    return;
  }

  var parsedName = parseName(name),
      currentRow = parsedName.row,
      currentColumn = colIndex(parsedName.col);

  if (keyEvent.key==='Enter' || keyEvent.key==='ArrowDown')
  {
    keyEvent.preventDefault();
    currentRow++;
  }
  else if (keyEvent.key==='ArrowUp')
  {
    keyEvent.preventDefault();
    currentRow--;
  }
  else if (keyEvent.key==='Tab'||keyEvent.key==='ArrowRight')
  {
    keyEvent.preventDefault();
    currentColumn++;
  }
  else if (keyEvent.key==='ArrowLeft')
  {
    keyEvent.preventDefault();
    currentColumn--;
  }
  else
    return;

  currentRow = Math.max(1, Math.min(ROWS, currentRow));
  currentColumn = Math.max(0, Math.min(COLS - 1, currentColumn));

  var next = colName(currentColumn) + currentRow,
      nextInput = document.getElementById('inp-' + next);

  if (nextInput)
  {
    nextInput.focus();
    selectCell(next);
  }
}

function onFormulaBarFocus()
{
  sheetSnapshotForUndo();
}

function formulaBarKey(keyEvent)
{
  if (keyEvent.key !== 'Enter')
    return;

  var ref = document.getElementById('cell-ref').value,
      val = document.getElementById('formula-bar').value;

  sheetData[ref] = val;

  var inp = document.getElementById('inp-' + ref);

  if (inp)
  {
    inp.value = evalCell(ref, val);
    inp.classList.toggle('formula-result', val.startsWith('='));
  }

  evaluateFormulas();
  saveSheetToFile();
}

function formulaBarInput()
{
  var ref = document.getElementById('cell-ref').value,
      val = document.getElementById('formula-bar').value;

  sheetData[ref] = val;

  var inp = document.getElementById('inp-' + ref);
  if (inp)
    inp.value = val;

  refreshRefHighlights();
}

// Splits a function's argument list on top-level commas, skipping commas
// that fall inside quoted strings or (defensively) nested parens.
function splitTopLevelArgs(text)
{
  var args = [],
      current = '',
      depth = 0,
      inQuote = false,
      quoteChar = '';

  for (var i = 0; i < text.length; i++)
  {
    var ch = text[i];

    if (inQuote)
    {
      current += ch;

      if (ch === quoteChar)
        inQuote = false;

      continue;
    }

    if (ch === '"' || ch === "'")
    {
      inQuote = true;
      quoteChar = ch;
      current += ch;
      continue;
    }

    if (ch === '(')
      depth++;
    else if (ch === ')')
      depth--;

    if (ch === ',' && depth === 0)
    {
      args.push(current);
      current = '';
      continue;
    }

    current += ch;
  }

  args.push(current);

  return args.map(function(a) { return a.trim(); });
}

// Converts a resolved JS value back into text that can be spliced into an
// expression string and safely re-evaluated (strings need quoting).
function formatForExpr(value)
{
  if (typeof value === 'string')
    return JSON.stringify(value);

  // A negative number spliced right after a '-' would read as the JS "--"
  // operator ("A1-(-2)" collapsing to "A1--2" is a SyntaxError); keep it
  // parenthesized so it always reads as a value.
  if (typeof value === 'number' && value < 0)
    return '(' + value + ')';

  return String(value);
}

// Final safety net before handing assembled text to the JS evaluator: the
// formula is built by splicing computed values into a string, so JS-only
// token traps need defusing. Runs of '-' are spaced apart (every '-' past
// the first is a unary negation — "5--3" means 5-(-3), not a decrement),
// and Excel's '^' power operator becomes JS's '**'. Quoted strings pass
// through untouched.
function sanitizeExprForEval(expr)
{
  return expr.replace
  (
    /"(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*'|-{2,}|\^/g,
    function(m)
    {
      if (m === '^') return '**';
      if (m[0] === '-') return m.split('').join(' ');
      return m;
    }
  );
}

// Replaces cell references with literals usable in a JS expression: numeric
// cells become bare numbers, "TRUE"/"FALSE" become real booleans, text cells
// become quoted string literals, and formula cells resolve recursively.
function substituteCellRefs(exprText)
{
  return exprText.replace
  (
    /"(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*'|(?<![A-Za-z])\$?([A-Z]{1,2})\$?(\d+)/g,
    function(match, c, r)
    {
      if (c === undefined)
        return match; // quoted string — ref-like text inside stays literal

      var ref = c + r,
          v = sheetData[ref];

      if (v === undefined || v === '')
        return '0';

      if (v.startsWith('='))
        v = evalCell(ref, v);

      var upper = v.toUpperCase();

      if (upper === 'TRUE' || upper === 'FALSE')
        return upper.toLowerCase();

      var num = parseFloat(v);

      if (isNaN(num))
        return JSON.stringify(v);

      // Parenthesize negatives so "A1-B1" with a negative B1 doesn't
      // collapse into the JS "--" operator.
      return num < 0 ? '(' + v + ')' : v;
    }
  );
}

// Evaluates a flat (no nested parens) sub-expression like "A3-A4" or "5*2",
// resolving any cell references first - used as the argument to scalar
// functions such as ABS/ROUND/SQRT/IF/etc.
function resolveValue(exprText)
{
  return Function('"use strict"; return (' + sanitizeExprForEval(substituteCellRefs(exprText)) + ')')();
}

// Resolves innermost parenthesis groups that are NOT function calls (the
// preceding character isn't part of a name like ROUND or LOG10) into their
// computed value. Ranges (A1:B3) are skipped — those belong to whichever
// aggregation function wraps them. Groups that fail to evaluate are left
// untouched, so an unresolvable formula still degrades to #ERR normally.
function resolveBareGroups(expr)
{
  return expr.replace
  (
    /(?<![A-Za-z0-9_])\(([^()]+)\)/g,
    function(match, inner)
    {
      if (inner.indexOf(':') !== -1)
        return match;

      try { return formatForExpr(resolveValue(inner)); }
      catch(e) { return match; }
    }
  );
}

function cellNumericValue(ref)
{
  var v = sheetData[ref];

  if (v === undefined || v === '')
    return 0;

  if (v.startsWith('='))
    v = evalCell(ref, v);

  var n = parseFloat(v);

  return isNaN(n) ? 0 : n;
}

// The used extent of the active page — whole-column (A:A) and whole-row
// (4:4) ranges cover this instead of the unbounded grid, so they only walk
// cells that actually hold something.
function sheetUsedExtent()
{
  var maxRow = 1, maxCol = 0;

  Object.keys(sheetData).forEach(function(ref)
  {
    if (!sheetData[ref]) return;
    var p = parseName(ref);
    if (!p) return;
    if (p.row > maxRow) maxRow = p.row;
    var c = colIndex(p.col);
    if (c > maxCol) maxCol = c;
  });

  return { maxRow: maxRow, maxCol: maxCol };
}

function getRangeRefs(range)
{
  var text = range.trim(), refs, i, j;

  // Whole-column range: A:A grabs one column, A:C a block of columns.
  var colRange = text.match(/^\$?([A-Z]{1,2})\s*:\s*\$?([A-Z]{1,2})$/);
  if (colRange)
  {
    var cLo = colIndex(colRange[1]), cHi = colIndex(colRange[2]);
    if (cLo > cHi) { var cT = cLo; cLo = cHi; cHi = cT; }

    refs = [];
    var lastRow = sheetUsedExtent().maxRow;
    for (i = 1; i <= lastRow; i++)
      for (j = cLo; j <= cHi; j++)
        refs.push(colName(j) + i);
    return refs;
  }

  // Whole-row range: 4:4 grabs one row, 2:5 a block of rows.
  var rowRange = text.match(/^\$?(\d+)\s*:\s*\$?(\d+)$/);
  if (rowRange)
  {
    var rLo = parseInt(rowRange[1]), rHi = parseInt(rowRange[2]);
    if (rLo > rHi) { var rT = rLo; rLo = rHi; rHi = rT; }

    refs = [];
    var lastCol = sheetUsedExtent().maxCol;
    for (i = rLo; i <= rHi; i++)
      for (j = 0; j <= lastCol; j++)
        refs.push(colName(j) + i);
    return refs;
  }

  var parts = range.split(':');

  if (parts.length === 1)
  {
    var single = parseName(parts[0].trim());
    return single ? [single.col + single.row] : [];
  }

  var pa = parseName(parts[0].trim()),
      pb = parseName(parts[1].trim());

  if (!pa || !pb)
    return [];

  var refs = [];

  for (var i = pa.row; i <= pb.row; i++)
    for (var j = colIndex(pa.col); j <= colIndex(pb.col); j++)
      refs.push(colName(j) + i);

  return refs;
}

function getRangeVals(range)
{
  return getRangeRefs(range).map(cellNumericValue);
}

// Cell refs covered by one aggregation argument: a range (A1:B3) or a single
// ref (A1). Returns null for anything else (a literal or an expression), so
// callers can fall back to value-based handling.
function argCellRefs(arg)
{
  if (arg.indexOf(':') !== -1)
  {
    var refs = getRangeRefs(arg);
    return refs.length ? refs : null;
  }

  var single = parseName(arg.trim());
  return single ? [single.col + single.row] : null;
}

// Collects the numeric values behind a mixed aggregation argument list:
// each argument may be a range, a single ref, a literal, or any expression
// the resolver can evaluate — so MAX(A1,B2), MAX(3,7), SUM(A1:A4, 10) and
// AVG(A1, (B1+1)*2) all behave like they do in a real spreadsheet.
function aggregateArgValues(argText)
{
  var vals = [];

  splitTopLevelArgs(argText).forEach(function(arg)
  {
    if (!arg)
      return;

    var refs = argCellRefs(arg);
    if (refs)
    {
      refs.forEach(function(ref) { vals.push(cellNumericValue(ref)); });
      return;
    }

    var v = resolveValue(arg),
        n = (typeof v === 'number') ? v : parseFloat(v);
    vals.push(isNaN(n) ? 0 : n);
  });

  return vals;
}

function sumRange(r)
{
  return aggregateArgValues(r).reduce
  (
    function(a, b)
    {
      return a + b;
    },
    0
  );
}

function avgRange(r)
{
  var vals = aggregateArgValues(r);
  return vals.length ? vals.reduce(function(a, b){ return a + b; }, 0) / vals.length : 0;
}

function countRange(r)
{
  var total = 0;
  splitTopLevelArgs(r).forEach(function(arg)
  {
    if (!arg) return;
    var refs = argCellRefs(arg);
    total += refs ? refs.length : 1;
  });
  return total;
}

function countaRange(r)
{
  var total = 0;
  splitTopLevelArgs(r).forEach(function(arg)
  {
    if (!arg) return;
    var refs = argCellRefs(arg);
    if (refs)
      total += refs.filter(function(ref) { return (sheetData[ref] || '') !== ''; }).length;
    else
      total += 1; // a literal/expression argument is inherently non-blank
  });
  return total;
}

function countBlankRange(r)
{
  var total = 0;
  splitTopLevelArgs(r).forEach(function(arg)
  {
    if (!arg) return;
    var refs = argCellRefs(arg);
    if (refs)
      total += refs.filter(function(ref) { return (sheetData[ref] || '') === ''; }).length;
  });
  return total;
}

function maxRange(r)
{
  var vals = aggregateArgValues(r);
  return vals.length ? Math.max.apply(null, vals) : 0;
}

function minRange(r)
{
  var vals = aggregateArgValues(r);
  return vals.length ? Math.min.apply(null, vals) : 0;
}

function medianRange(r)
{
  var vals = aggregateArgValues(r).slice().sort(function(a, b) { return a - b; }),
      len = vals.length;

  if (!len)
    return 0;

  var mid = Math.floor(len / 2);

  return (len % 2) ? vals[mid] : (vals[mid - 1] + vals[mid]) / 2;
}

function productRange(r)
{
  return aggregateArgValues(r).reduce
  (
    function(a, b)
    {
      return a * b;
    },
    1
  );
}

function varianceRange(r)
{
  var vals = aggregateArgValues(r),
      n = vals.length;

  if (n < 2)
    return 0;

  var mean = vals.reduce(function(a, b) { return a + b; }, 0) / n;

  return vals.reduce(function(a, b) { return a + Math.pow(b - mean, 2); }, 0) / (n - 1);
}

function stdevRange(r)
{
  return Math.sqrt(varianceRange(r));
}

// A cell's value as an expression literal (numbers bare, TRUE/FALSE bare,
// text quoted, blanks 0) — the same shape substituteCellRefs produces, so
// lookup functions can drop their result straight into the formula text.
function cellExprLiteral(ref)
{
  var v = sheetData[ref];

  if (v === undefined || v === '')
    return '0';

  if (v.startsWith('='))
    v = evalCell(ref, v);

  var upper = String(v).toUpperCase();

  if (upper === 'TRUE' || upper === 'FALSE')
    return upper.toLowerCase();

  var num = parseFloat(v);

  if (isNaN(num))
    return JSON.stringify(v);

  // Parenthesize negatives so "A1-B1" with a negative B1 doesn't collapse
  // into the JS "--" operator.
  return num < 0 ? '(' + v + ')' : v;
}

function refsBoundingBox(refs)
{
  var minRow = Infinity, maxRow = 0, minCol = Infinity, maxCol = 0;

  refs.forEach(function(ref)
  {
    var p = parseName(ref);
    if (!p) return;
    minRow = Math.min(minRow, p.row); maxRow = Math.max(maxRow, p.row);
    minCol = Math.min(minCol, colIndex(p.col)); maxCol = Math.max(maxCol, colIndex(p.col));
  });

  return { minRow: minRow, maxRow: maxRow, minCol: minCol, maxCol: maxCol };
}

// True when a displayed cell value matches a lookup target — numerically
// when both sides are numbers, case-insensitive text comparison otherwise.
function lookupValueMatches(value, target)
{
  var targetNum = (typeof target === 'number') ? target : parseFloat(target),
      targetIsNum = (typeof target === 'number') || (String(target).trim() !== '' && !isNaN(targetNum));

  if (targetIsNum)
    return value !== '' && parseFloat(value) === targetNum;

  return String(value).toLowerCase() === String(target).toLowerCase();
}

// INDEX(range, row, [col]) — the value at a position inside a range. For a
// single-row range the second argument indexes along the row (like Excel's
// one-dimensional form), otherwise it picks the row and the optional third
// argument picks the column (both 1-based, defaulting to 1). The result is
// formatted like a substituted cell ref, so text comes back quoted and blank
// cells come back as 0.
function indexRange(argText)
{
  var args = splitTopLevelArgs(argText),
      refs = argCellRefs(args[0]);

  if (!refs || !refs.length)
    throw new Error('INDEX needs a range');

  var box = refsBoundingBox(refs),
      height = box.maxRow - box.minRow + 1,
      width = box.maxCol - box.minCol + 1,
      rowNum = args[1] !== undefined ? resolveValue(args[1]) : 1,
      colNum;

  if (args[2] !== undefined)
    colNum = resolveValue(args[2]);
  else if (height === 1 && width > 1)
  {
    colNum = rowNum;
    rowNum = 1;
  }
  else
    colNum = 1;

  if (rowNum < 1 || rowNum > height || colNum < 1 || colNum > width)
    throw new Error('INDEX out of range');

  return cellExprLiteral(colName(box.minCol + colNum - 1) + (box.minRow + rowNum - 1));
}

// VLOOKUP(value, range, colIndex) — finds the first row whose FIRST-column
// value matches, then returns that row's value from the 1-based colIndex-th
// column of the range. Exact matching only. HLOOKUP is the same across rows.
function vlookupRange(argText)
{
  var args = splitTopLevelArgs(argText),
      target = resolveValue(args[0]),
      refs = argCellRefs(args[1]);

  if (!refs || !refs.length || args[2] === undefined)
    throw new Error('VLOOKUP needs a value, range, and column index');

  var box = refsBoundingBox(refs),
      colOffset = resolveValue(args[2]) - 1;

  if (colOffset < 0 || box.minCol + colOffset > box.maxCol)
    throw new Error('VLOOKUP column index out of range');

  for (var r = box.minRow; r <= box.maxRow; r++)
    if (lookupValueMatches(getDisplayValue(colName(box.minCol) + r), target))
      return cellExprLiteral(colName(box.minCol + colOffset) + r);

  return '"#N/A"';
}

function hlookupRange(argText)
{
  var args = splitTopLevelArgs(argText),
      target = resolveValue(args[0]),
      refs = argCellRefs(args[1]);

  if (!refs || !refs.length || args[2] === undefined)
    throw new Error('HLOOKUP needs a value, range, and row index');

  var box = refsBoundingBox(refs),
      rowOffset = resolveValue(args[2]) - 1;

  if (rowOffset < 0 || box.minRow + rowOffset > box.maxRow)
    throw new Error('HLOOKUP row index out of range');

  for (var c = box.minCol; c <= box.maxCol; c++)
    if (lookupValueMatches(getDisplayValue(colName(c) + box.minRow), target))
      return cellExprLiteral(colName(c) + (box.minRow + rowOffset));

  return '"#N/A"';
}

// MATCH(value, range) — 1-based position of the first exact match walking
// the range row-major (use a single row or column range, like in Excel).
function matchRange(argText)
{
  var args = splitTopLevelArgs(argText),
      target = resolveValue(args[0]),
      refs = argCellRefs(args[1]);

  if (!refs || !refs.length)
    throw new Error('MATCH needs a range');

  for (var i = 0; i < refs.length; i++)
    if (lookupValueMatches(getDisplayValue(refs[i]), target))
      return i + 1;

  return '"#N/A"';
}

// Excel-style criterion (">5", "<>0", "apples", 12) → a predicate over a
// cell's displayed value, shared by COUNTIF / SUMIF / AVERAGEIF.
function makeCriterionTest(rawArg)
{
  var crit = resolveValue(rawArg),
      s = String(crit),
      m = s.match(/^(>=|<=|<>|=|>|<)([\s\S]*)$/),
      op = m ? m[1] : '=',
      rhs = m ? m[2] : s,
      rhsNum = parseFloat(rhs),
      rhsIsNum = rhs.trim() !== '' && !isNaN(rhsNum);

  return function(value)
  {
    var num = parseFloat(value),
        bothNum = rhsIsNum && value !== '' && !isNaN(num);

    switch (op)
    {
      case '>':  return bothNum && num > rhsNum;
      case '<':  return bothNum && num < rhsNum;
      case '>=': return bothNum && num >= rhsNum;
      case '<=': return bothNum && num <= rhsNum;
      case '<>': return bothNum ? num !== rhsNum : String(value).toLowerCase() !== rhs.toLowerCase();
      default:   return bothNum ? num === rhsNum : String(value).toLowerCase() === rhs.toLowerCase();
    }
  };
}

function countifRange(argText)
{
  var args = splitTopLevelArgs(argText),
      refs = argCellRefs(args[0]);

  if (!refs)
    throw new Error('COUNTIF needs a range');

  var test = makeCriterionTest(args[1]);

  return refs.filter(function(ref){ return test(getDisplayValue(ref)); }).length;
}

// SUMIF(range, criterion, [sumRange]): sums sumRange (default: the range
// itself) wherever the criterion matches the corresponding range cell.
function sumifRange(argText)
{
  var args = splitTopLevelArgs(argText),
      refs = argCellRefs(args[0]);

  if (!refs)
    throw new Error('SUMIF needs a range');

  var test = makeCriterionTest(args[1]),
      sumRefs = args[2] !== undefined ? argCellRefs(args[2]) : refs,
      total = 0;

  refs.forEach(function(ref, i)
  {
    if (test(getDisplayValue(ref)) && sumRefs && sumRefs[i] !== undefined)
      total += cellNumericValue(sumRefs[i]);
  });

  return total;
}

function averageifRange(argText)
{
  var args = splitTopLevelArgs(argText),
      refs = argCellRefs(args[0]);

  if (!refs)
    throw new Error('AVERAGEIF needs a range');

  var test = makeCriterionTest(args[1]),
      avgRefs = args[2] !== undefined ? argCellRefs(args[2]) : refs,
      total = 0,
      count = 0;

  refs.forEach(function(ref, i)
  {
    if (test(getDisplayValue(ref)) && avgRefs && avgRefs[i] !== undefined)
    {
      total += cellNumericValue(avgRefs[i]);
      count++;
    }
  });

  if (!count)
    throw new Error('AVERAGEIF matched nothing');

  return total / count;
}

// LARGE/SMALL(values..., k): the k-th largest/smallest value (1-based). The
// LAST argument is k and everything before it supplies values — so both
// LARGE(A1:A9, 2) and an @-expanded LARGE(@A1:A5*@B1:B5, 2) work.
function largeSmallRange(argText, wantLargest)
{
  var args = splitTopLevelArgs(argText),
      k = 1,
      valueArgs = args;

  if (args.length > 1)
  {
    k = resolveValue(args[args.length - 1]);
    valueArgs = args.slice(0, -1);
  }

  var vals = aggregateArgValues(valueArgs.join(',')).slice().sort
      (
        function(a, b) { return wantLargest ? b - a : a - b; }
      );

  if (k < 1 || k > vals.length)
    throw new Error('k out of range');

  return vals[k - 1];
}

// Runs every named-function pattern once. Each pattern requires a function
// name not preceded by a letter (so e.g. OR doesn't match inside FLOOR, or
// AND inside RAND) and an argument with no unresolved nested parens. evalCell
// re-runs this pass until the text stops changing, which is what lets calls
// like ROUND(SUM(A1:A4),2) resolve correctly regardless of nesting order.
function applyFunctionPass(expr, selfRef)
{
  return expr
    .replace(/(?<![A-Za-z])SUM\(([^()]+)\)/gi, function(_, r) { return sumRange(r); })
    .replace(/(?<![A-Za-z])AVG\(([^()]+)\)/gi, function(_, r) { return avgRange(r); })
    .replace(/(?<![A-Za-z])MEAN\(([^()]+)\)/gi, function(_, r) { return avgRange(r); })
    .replace(/(?<![A-Za-z])COUNTA\(([^()]+)\)/gi, function(_, r) { return countaRange(r); })
    .replace(/(?<![A-Za-z])COUNTBLANK\(([^()]+)\)/gi, function(_, r) { return countBlankRange(r); })
    .replace(/(?<![A-Za-z])COUNT\(([^()]+)\)/gi, function(_, r) { return countRange(r); })
    .replace(/(?<![A-Za-z])MAX\(([^()]+)\)/gi, function(_, r) { return maxRange(r); })
    .replace(/(?<![A-Za-z])MIN\(([^()]+)\)/gi, function(_, r) { return minRange(r); })
    .replace(/(?<![A-Za-z])MEDIAN\(([^()]+)\)/gi, function(_, r) { return medianRange(r); })
    .replace(/(?<![A-Za-z])PRODUCT\(([^()]+)\)/gi, function(_, r) { return productRange(r); })
    .replace(/(?<![A-Za-z])STDEV\(([^()]+)\)/gi, function(_, r) { return stdevRange(r); })
    .replace(/(?<![A-Za-z])VAR\(([^()]+)\)/gi, function(_, r) { return varianceRange(r); })
    .replace(/(?<![A-Za-z])INDEX\(([^()]+)\)/gi, function(_, a) { return indexRange(a); })
    .replace(/(?<![A-Za-z])VLOOKUP\(([^()]+)\)/gi, function(_, a) { return vlookupRange(a); })
    .replace(/(?<![A-Za-z])HLOOKUP\(([^()]+)\)/gi, function(_, a) { return hlookupRange(a); })
    .replace(/(?<![A-Za-z])MATCH\(([^()]+)\)/gi, function(_, a) { return matchRange(a); })
    .replace(/(?<![A-Za-z])COUNTIF\(([^()]+)\)/gi, function(_, a) { return countifRange(a); })
    .replace(/(?<![A-Za-z])SUMIF\(([^()]+)\)/gi, function(_, a) { return sumifRange(a); })
    .replace(/(?<![A-Za-z])AVERAGEIF\(([^()]+)\)/gi, function(_, a) { return averageifRange(a); })
    .replace(/(?<![A-Za-z])AVERAGE\(([^()]+)\)/gi, function(_, r) { return avgRange(r); })
    .replace(/(?<![A-Za-z])LARGE\(([^()]+)\)/gi, function(_, a) { return largeSmallRange(a, true); })
    .replace(/(?<![A-Za-z])SMALL\(([^()]+)\)/gi, function(_, a) { return largeSmallRange(a, false); })
    .replace
    (
      /(?<![A-Za-z])ROW\(([^()]*)\)/gi,
      function(match, a)
      {
        var p = parseName((a.trim() || selfRef || '').replace(/\$/g, '').toUpperCase());
        return p ? p.row : match;
      }
    )
    .replace
    (
      /(?<![A-Za-z])COLUMN\(([^()]*)\)/gi,
      function(match, a)
      {
        var p = parseName((a.trim() || selfRef || '').replace(/\$/g, '').toUpperCase());
        return p ? colIndex(p.col) + 1 : match;
      }
    )
    .replace(/(?<![A-Za-z])SIN\(([^()]+)\)/gi, function(_, a) { return Math.sin(resolveValue(a)); })
    .replace(/(?<![A-Za-z])COS\(([^()]+)\)/gi, function(_, a) { return Math.cos(resolveValue(a)); })
    .replace(/(?<![A-Za-z])TAN\(([^()]+)\)/gi, function(_, a) { return Math.tan(resolveValue(a)); })
    .replace(/(?<![A-Za-z])ASIN\(([^()]+)\)/gi, function(_, a) { return Math.asin(resolveValue(a)); })
    .replace(/(?<![A-Za-z])ACOS\(([^()]+)\)/gi, function(_, a) { return Math.acos(resolveValue(a)); })
    .replace(/(?<![A-Za-z])ATAN2\(([^()]+)\)/gi, function(_, a)
    {
      // Excel argument order is (x, y); JS atan2 takes (y, x).
      var args = splitTopLevelArgs(a);
      return Math.atan2(resolveValue(args[1]), resolveValue(args[0]));
    })
    .replace(/(?<![A-Za-z])ATAN\(([^()]+)\)/gi, function(_, a) { return Math.atan(resolveValue(a)); })
    .replace(/(?<![A-Za-z])SINH\(([^()]+)\)/gi, function(_, a) { return Math.sinh(resolveValue(a)); })
    .replace(/(?<![A-Za-z])COSH\(([^()]+)\)/gi, function(_, a) { return Math.cosh(resolveValue(a)); })
    .replace(/(?<![A-Za-z])TANH\(([^()]+)\)/gi, function(_, a) { return Math.tanh(resolveValue(a)); })
    .replace(/(?<![A-Za-z])DEGREES\(([^()]+)\)/gi, function(_, a) { return resolveValue(a) * 180 / Math.PI; })
    .replace(/(?<![A-Za-z])RADIANS\(([^()]+)\)/gi, function(_, a) { return resolveValue(a) * Math.PI / 180; })
    .replace(/(?<![A-Za-z])ABS\(([^()]+)\)/gi, function(_, a) { return Math.abs(resolveValue(a)); })
    .replace(/(?<![A-Za-z])SQRT\(([^()]+)\)/gi, function(_, a) { return Math.sqrt(resolveValue(a)); })
    .replace(/(?<![A-Za-z])SIGN\(([^()]+)\)/gi, function(_, a) { return Math.sign(resolveValue(a)); })
    .replace(/(?<![A-Za-z])EXP\(([^()]+)\)/gi, function(_, a) { return Math.exp(resolveValue(a)); })
    .replace(/(?<![A-Za-z])LN\(([^()]+)\)/gi, function(_, a) { return Math.log(resolveValue(a)); })
    .replace(/(?<![A-Za-z])LOG10\(([^()]+)\)/gi, function(_, a) { return Math.log10(resolveValue(a)); })
    .replace
    (
      /(?<![A-Za-z])LOG\(([^()]+)\)/gi,
      function(_, a)
      {
        var args = splitTopLevelArgs(a),
            num = resolveValue(args[0]),
            base = args[1] !== undefined ? resolveValue(args[1]) : 10;

        return Math.log(num) / Math.log(base);
      }
    )
    .replace(/(?<![A-Za-z])PI\(\)/gi, function() { return Math.PI; })
    .replace(/(?<![A-Za-z])RAND\(\)/gi, function() { return Math.random(); })
    .replace
    (
      /(?<![A-Za-z])RANDBETWEEN\(([^()]+)\)/gi,
      function(_, a)
      {
        var args = splitTopLevelArgs(a),
            lo = resolveValue(args[0]),
            hi = resolveValue(args[1]);

        return Math.floor(Math.random() * (hi - lo + 1)) + lo;
      }
    )
    .replace
    (
      /(?<![A-Za-z])POWER\(([^()]+)\)/gi,
      function(_, a)
      {
        var args = splitTopLevelArgs(a);
        return Math.pow(resolveValue(args[0]), resolveValue(args[1]));
      }
    )
    .replace
    (
      /(?<![A-Za-z])MOD\(([^()]+)\)/gi,
      function(_, a)
      {
        var args = splitTopLevelArgs(a);
        return resolveValue(args[0]) % resolveValue(args[1]);
      }
    )
    .replace
    (
      /(?<![A-Za-z])TRUNC\(([^()]+)\)/gi,
      function(_, a)
      {
        var args = splitTopLevelArgs(a),
            num = resolveValue(args[0]),
            digits = args[1] !== undefined ? resolveValue(args[1]) : 0,
            factor = Math.pow(10, digits);

        return Math.trunc(num * factor) / factor;
      }
    )
    .replace(/(?<![A-Za-z])INT\(([^()]+)\)/gi, function(_, a) { return Math.floor(resolveValue(a)); })
    .replace
    (
      /(?<![A-Za-z])CEILING\(([^()]+)\)/gi,
      function(_, a)
      {
        var args = splitTopLevelArgs(a),
            num = resolveValue(args[0]),
            sig = args[1] !== undefined ? resolveValue(args[1]) : 1;

        return sig ? Math.ceil(num / sig) * sig : 0;
      }
    )
    .replace
    (
      /(?<![A-Za-z])FLOOR\(([^()]+)\)/gi,
      function(_, a)
      {
        var args = splitTopLevelArgs(a),
            num = resolveValue(args[0]),
            sig = args[1] !== undefined ? resolveValue(args[1]) : 1;

        return sig ? Math.floor(num / sig) * sig : 0;
      }
    )
    .replace
    (
      /(?<![A-Za-z])CONCATENATE\(([^()]+)\)/gi,
      function(_, a)
      {
        return formatForExpr(splitTopLevelArgs(a).map(resolveValue).join(''));
      }
    )
    .replace
    (
      /(?<![A-Za-z])CONCAT\(([^()]+)\)/gi,
      function(_, a)
      {
        return formatForExpr(splitTopLevelArgs(a).map(resolveValue).join(''));
      }
    )
    .replace(/(?<![A-Za-z])LEN\(([^()]+)\)/gi, function(_, a) { return String(resolveValue(a)).length; })
    .replace(/(?<![A-Za-z])UPPER\(([^()]+)\)/gi, function(_, a) { return formatForExpr(String(resolveValue(a)).toUpperCase()); })
    .replace(/(?<![A-Za-z])LOWER\(([^()]+)\)/gi, function(_, a) { return formatForExpr(String(resolveValue(a)).toLowerCase()); })
    .replace(/(?<![A-Za-z])TRIM\(([^()]+)\)/gi, function(_, a) { return formatForExpr(String(resolveValue(a)).trim()); })
    .replace
    (
      /(?<![A-Za-z])LEFT\(([^()]+)\)/gi,
      function(_, a)
      {
        var args = splitTopLevelArgs(a),
            text = String(resolveValue(args[0])),
            n = args[1] !== undefined ? resolveValue(args[1]) : 1;

        return formatForExpr(text.slice(0, n));
      }
    )
    .replace
    (
      /(?<![A-Za-z])RIGHT\(([^()]+)\)/gi,
      function(_, a)
      {
        var args = splitTopLevelArgs(a),
            text = String(resolveValue(args[0])),
            n = args[1] !== undefined ? resolveValue(args[1]) : 1;

        return formatForExpr(n ? text.slice(-n) : '');
      }
    )
    .replace
    (
      /(?<![A-Za-z])MID\(([^()]+)\)/gi,
      function(_, a)
      {
        var args = splitTopLevelArgs(a),
            text = String(resolveValue(args[0])),
            start = resolveValue(args[1]),
            count = resolveValue(args[2]);

        return formatForExpr(text.substr(start - 1, count));
      }
    )
    .replace
    (
      /(?<![A-Za-z])SUBSTITUTE\(([^()]+)\)/gi,
      function(_, a)
      {
        var args = splitTopLevelArgs(a),
            text = String(resolveValue(args[0])),
            oldText = String(resolveValue(args[1])),
            newText = String(resolveValue(args[2]));

        if (!oldText)
          return formatForExpr(text);

        // Optional 4th argument: replace only the n-th occurrence.
        if (args[3] !== undefined)
        {
          var n = resolveValue(args[3]),
              at = -1;

          for (var i = 0; i < n; i++)
          {
            at = text.indexOf(oldText, at + 1);
            if (at === -1) break;
          }

          return formatForExpr(at === -1 ? text : text.slice(0, at) + newText + text.slice(at + oldText.length));
        }

        return formatForExpr(text.split(oldText).join(newText));
      }
    )
    .replace
    (
      /(?<![A-Za-z])REPT\(([^()]+)\)/gi,
      function(_, a)
      {
        var args = splitTopLevelArgs(a),
            text = String(resolveValue(args[0])),
            times = Math.max(0, Math.floor(resolveValue(args[1])));

        return formatForExpr(new Array(times + 1).join(text));
      }
    )
    .replace
    (
      /(?<![A-Za-z])FIND\(([^()]+)\)/gi,
      function(_, a)
      {
        var args = splitTopLevelArgs(a),
            needle = String(resolveValue(args[0])),
            hay = String(resolveValue(args[1])),
            start = args[2] !== undefined ? resolveValue(args[2]) : 1,
            at = hay.indexOf(needle, start - 1);

        if (at === -1)
          throw new Error('FIND: not found');

        return at + 1;
      }
    )
    .replace
    (
      /(?<![A-Za-z])SEARCH\(([^()]+)\)/gi,
      function(_, a)
      {
        var args = splitTopLevelArgs(a),
            needle = String(resolveValue(args[0])).toLowerCase(),
            hay = String(resolveValue(args[1])).toLowerCase(),
            start = args[2] !== undefined ? resolveValue(args[2]) : 1,
            at = hay.indexOf(needle, start - 1);

        if (at === -1)
          throw new Error('SEARCH: not found');

        return at + 1;
      }
    )
    .replace
    (
      /(?<![A-Za-z])PROPER\(([^()]+)\)/gi,
      function(_, a)
      {
        var text = String(resolveValue(a)).replace
        (
          /[A-Za-z]+/g,
          function(word) { return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase(); }
        );

        return formatForExpr(text);
      }
    )
    .replace
    (
      /(?<![A-Za-z])EXACT\(([^()]+)\)/gi,
      function(_, a)
      {
        var args = splitTopLevelArgs(a);
        return String(resolveValue(args[0])) === String(resolveValue(args[1]));
      }
    )
    .replace
    (
      /(?<![A-Za-z])IFERROR\(([^()]+)\)/gi,
      function(_, a)
      {
        var args = splitTopLevelArgs(a);

        try
        {
          var primary = resolveValue(args[0]);

          if (typeof primary === 'number' && (isNaN(primary) || !isFinite(primary)))
            throw new Error('bad number');

          return formatForExpr(primary);
        }
        catch (e)
        {
          return formatForExpr(resolveValue(args[1]));
        }
      }
    )
    .replace
    (
      /(?<![A-Za-z])IF\(([^()]+)\)/gi,
      function(_, a)
      {
        var args = splitTopLevelArgs(a),
            cond = resolveValue(args[0]);

        return formatForExpr(cond ? resolveValue(args[1]) : resolveValue(args[2]));
      }
    )
    .replace
    (
      /(?<![A-Za-z])AND\(([^()]+)\)/gi,
      function(_, a)
      {
        return splitTopLevelArgs(a).every(function(x) { return !!resolveValue(x); });
      }
    )
    .replace
    (
      /(?<![A-Za-z])OR\(([^()]+)\)/gi,
      function(_, a)
      {
        return splitTopLevelArgs(a).some(function(x) { return !!resolveValue(x); });
      }
    )
    .replace(/(?<![A-Za-z])NOT\(([^()]+)\)/gi, function(_, a) { return !resolveValue(a); })
    .replace
    (
      /(?<![A-Za-z])XOR\(([^()]+)\)/gi,
      function(_, a)
      {
        return splitTopLevelArgs(a).filter(function(x) { return !!resolveValue(x); }).length % 2 === 1;
      }
    )
    .replace
    (
      /(?<![A-Za-z])ROUNDUP\(([^()]+)\)/gi,
      function(_, a)
      {
        var args = splitTopLevelArgs(a),
            num = resolveValue(args[0]),
            digits = args[1] !== undefined ? resolveValue(args[1]) : 0,
            factor = Math.pow(10, digits);

        // Away from zero, like Excel.
        return Math.sign(num) * Math.ceil(Math.abs(num) * factor) / factor;
      }
    )
    .replace
    (
      /(?<![A-Za-z])ROUNDDOWN\(([^()]+)\)/gi,
      function(_, a)
      {
        var args = splitTopLevelArgs(a),
            num = resolveValue(args[0]),
            digits = args[1] !== undefined ? resolveValue(args[1]) : 0,
            factor = Math.pow(10, digits);

        // Toward zero, like Excel.
        return Math.sign(num) * Math.floor(Math.abs(num) * factor) / factor;
      }
    )
    .replace
    (
      /(?<![A-Za-z])FACT\(([^()]+)\)/gi,
      function(_, a)
      {
        var n = Math.floor(resolveValue(a));

        if (n < 0)
          throw new Error('FACT of a negative number');

        var result = 1;
        for (var i = 2; i <= n; i++)
          result *= i;

        return result;
      }
    )
    .replace
    (
      /(?<![A-Za-z])EVEN\(([^()]+)\)/gi,
      function(_, a)
      {
        var num = resolveValue(a);
        return Math.sign(num) * Math.ceil(Math.abs(num) / 2) * 2;
      }
    )
    .replace
    (
      /(?<![A-Za-z])ODD\(([^()]+)\)/gi,
      function(_, a)
      {
        var num = resolveValue(a);
        if (!num) return 1;
        return Math.sign(num) * (2 * Math.ceil((Math.abs(num) + 1) / 2) - 1);
      }
    )
    .replace
    (
      /(?<![A-Za-z])ROUND\(([^()]+)\)/gi,
      function(_, a)
      {
        var args = splitTopLevelArgs(a),
            num = resolveValue(args[0]),
            decimals = args[1] !== undefined ? resolveValue(args[1]) : 0;

        return Number(num.toFixed(decimals));
      }
    );
}

// Evaluates a cell that lives on another tab: the whole formula engine reads
// the global sheetData, so swap the alias for the duration of the call. The
// eval cache is keyed by bare refs, so it must not survive the swap in
// either direction.
function evalCellOnPage(page, ref, val)
{
  if (page.data === sheetData)
    return evalCell(ref, val);

  var saved = sheetData;
  sheetData = page.data;
  clearEvalCache();
  try { return evalCell(ref, val); }
  finally { sheetData = saved; clearEvalCache(); }
}

// Cross-tab references use the link-like syntax [Tab name](A1). Resolved
// before anything else so they compose with functions: =SUM([Costs](B2), 5).
// Unknown tab names come back as "#REF".
function substitutePageRefs(exprText)
{
  return exprText.replace
  (
    /\[([^\]]+)\]\(\s*\$?([A-Z]{1,2})\$?(\d+)\s*\)/g,
    function(_, pageName, c, r)
    {
      var target = pageName.trim().toLowerCase();
      var page = sheetPages.find(function(p){ return p.name.trim().toLowerCase() === target; });
      if (!page)
        return '"#REF"';

      var ref = c + r,
          v = page.data[ref];

      if (v === undefined || v === '')
        return '0';

      if (v.startsWith('='))
        v = evalCellOnPage(page, ref, v);

      var upper = String(v).toUpperCase();
      if (upper === 'TRUE' || upper === 'FALSE')
        return upper.toLowerCase();

      var num = parseFloat(v);

      if (isNaN(num))
        return JSON.stringify(v);

      // Parenthesize negatives so "A1-B1" with a negative B1 doesn't
      // collapse into the JS "--" operator.
      return num < 0 ? '(' + v + ')' : v;
    }
  );
}

// Refs currently being evaluated, so a formula that (directly or through a
// whole-column range like SUM(A:A) sitting inside column A) references
// itself resolves the cycle as 0 instead of recursing forever.
var evalCellInProgress = {};

// Memoized formula results. Within one recalculation pass the same cell gets
// evaluated over and over (every range that covers it re-reads it, the save
// serializer evaluates it again, charts again) — without a cache, a sheet of
// chained or whole-column formulas recalculates quadratically on every
// keystroke. Cleared at the start of each pass and whenever the sheetData
// alias switches pages, so entries can never outlive the data they came from.
var evalCellCache = {};

function clearEvalCache()
{
  evalCellCache = {};
}

// ── @ RANGES (elementwise / line-per-line operations) ──
// Two complementary behaviors, matching how Excel treats arrays vs implicit
// intersection:
//
//  1. Inside an aggregate function, an argument mixing @ranges with
//     operators is expanded one element per line and handed to the function
//     as a plain value list: =SUM(@I18:I22*@J18:J22) becomes
//     SUM(I18*J18, I19*J19, …) — a single SUMPRODUCT-style result.
//  2. Anywhere else, @range picks the one value on the formula's own
//     row/column, so =@A1:A9*@B1:B9 filled down computes each line's pair.

var AT_EXPANDING_FUNCTIONS = {
  SUM: 1, AVG: 1, AVERAGE: 1, MEAN: 1, MAX: 1, MIN: 1, MEDIAN: 1,
  PRODUCT: 1, STDEV: 1, VAR: 1, COUNT: 1, COUNTA: 1, LARGE: 1, SMALL: 1
};

// Matches quoted strings (passed through) or an @range (captured).
var AT_RANGE_RE = /"(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*'|@\s*(\$?[A-Z]{1,2}\$?\d+\s*:\s*\$?[A-Z]{1,2}\$?\d+|\$?[A-Z]{1,2}\s*:\s*\$?[A-Z]{1,2}|\$?\d+\s*:\s*\$?\d+)/g;

// Rewrites one paren/comma-free argument segment ("@A1:A3*@B1:B3+2") into a
// comma-separated element list ("A1*B1+2,A2*B2+2,A3*B3+2"). Scalars repeat
// on every line; all @ranges in the segment must have the same size.
function expandAtSegment(segment)
{
  var ranges = [],
      m;

  AT_RANGE_RE.lastIndex = 0;
  while ((m = AT_RANGE_RE.exec(segment)))
    if (m[1] !== undefined) ranges.push(m[1]);

  if (!ranges.length)
    return segment; // only @A5-style single refs — implicit intersection's job

  var refLists = ranges.map(function(r){ return getRangeRefs(r.replace(/[\s$]/g, '')); }),
      len = refLists[0].length;

  refLists.forEach(function(list)
  {
    if (list.length !== len)
      throw new Error('@ ranges in one expression must be the same size');
  });

  if (!len)
    throw new Error('@ range is empty');

  var elements = [];

  for (var k = 0; k < len; k++)
  {
    var idx = 0;
    elements.push(segment.replace(AT_RANGE_RE, function(match, r)
    {
      return r === undefined ? match : refLists[idx++][k];
    }));
  }

  return elements.join(',');
}

// Walks the formula tracking which function call each piece sits inside;
// @-bearing segments whose innermost enclosing call is an aggregate get
// elementwise expansion. Everything else is left for implicit intersection.
function expandArrayOps(expr)
{
  if (expr.indexOf('@') === -1)
    return expr;

  var result = '',
      segment = '',
      stack = [],
      i = 0;

  function flush()
  {
    var context = stack.length ? stack[stack.length - 1] : null;

    if (context && context.agg && segment.indexOf('@') !== -1)
      result += expandAtSegment(segment);
    else
      result += segment;

    segment = '';
  }

  while (i < expr.length)
  {
    var ch = expr.charAt(i);

    // Quoted strings ride along inside the segment untouched (and their
    // commas/parens don't count as delimiters).
    if (ch === '"' || ch === "'")
    {
      var j = i + 1;
      while (j < expr.length && expr.charAt(j) !== ch)
        j += (expr.charAt(j) === '\\') ? 2 : 1;
      segment += expr.slice(i, j + 1);
      i = j + 1;
      continue;
    }

    if (ch === '(')
    {
      var nameMatch = /([A-Za-z]+)\s*$/.exec(segment),
          fnName = nameMatch ? nameMatch[1].toUpperCase() : null;
      flush();
      result += '(';
      stack.push({ agg: !!(fnName && AT_EXPANDING_FUNCTIONS[fnName]) });
      i++;
      continue;
    }

    if (ch === ')')
    {
      flush();
      result += ')';
      stack.pop();
      i++;
      continue;
    }

    if (ch === ',')
    {
      flush();
      result += ',';
      i++;
      continue;
    }

    segment += ch;
    i++;
  }

  flush();
  return result;
}

// Excel-style implicit intersection: @range picks the single value from the
// range that lines up with the formula's own row (or column, for horizontal
// ranges) — the "same line as me" operator. Fill =@A1:A9*@B1:B9 down beside
// the data and each row multiplies its own pair; whole columns work too
// (=@A:A*@B:B). A range that doesn't cross the formula's line is an error,
// like Excel's #VALUE!.
function substituteImplicitIntersection(exprText, selfRef)
{
  var self = parseName((selfRef || '').replace(/\$/g, ''));

  return exprText.replace
  (
    /"(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*'|@\s*(\$?[A-Z]{1,2}\$?\d+\s*:\s*\$?[A-Z]{1,2}\$?\d+|\$?[A-Z]{1,2}\s*:\s*\$?[A-Z]{1,2}|\$?\d+\s*:\s*\$?\d+|\$?[A-Z]{1,2}\$?\d+)/g,
    function(match, rangeText)
    {
      if (rangeText === undefined)
        return match; // quoted string — leave alone

      if (!self)
        throw new Error('@ needs a host cell');

      var clean = rangeText.replace(/[\s$]/g, ''),
          m;

      // @A5 on a plain ref is a no-op.
      if ((m = clean.match(/^([A-Z]{1,2})(\d+)$/)))
        return m[1] + m[2];

      var selfCol = colIndex(self.col),
          rowLo = null, rowHi = null,
          colLo = null, colHi = null;

      if ((m = clean.match(/^([A-Z]{1,2})(\d+):([A-Z]{1,2})(\d+)$/)))
      {
        colLo = Math.min(colIndex(m[1]), colIndex(m[3]));
        colHi = Math.max(colIndex(m[1]), colIndex(m[3]));
        rowLo = Math.min(+m[2], +m[4]);
        rowHi = Math.max(+m[2], +m[4]);
      }
      else if ((m = clean.match(/^([A-Z]{1,2}):([A-Z]{1,2})$/)))
      {
        colLo = Math.min(colIndex(m[1]), colIndex(m[2]));
        colHi = Math.max(colIndex(m[1]), colIndex(m[2]));
      }
      else if ((m = clean.match(/^(\d+):(\d+)$/)))
      {
        rowLo = Math.min(+m[1], +m[2]);
        rowHi = Math.max(+m[1], +m[2]);
      }
      else
        return match;

      // Pick the row: a single-row range contributes its one row; a span (or
      // an unbounded whole-column dimension) must contain the formula's row.
      var row, col;

      if (rowLo === null || (rowLo !== rowHi && self.row >= rowLo && self.row <= rowHi))
        row = self.row;
      else if (rowLo === rowHi)
        row = rowLo;
      else
        throw new Error('@' + rangeText + ' does not reach row ' + self.row);

      if (colLo === null || (colLo !== colHi && selfCol >= colLo && selfCol <= colHi))
        col = selfCol;
      else if (colLo === colHi)
        col = colLo;
      else
        throw new Error('@' + rangeText + ' does not reach column ' + self.col);

      return colName(col) + row;
    }
  );
}

function evalCell(ref, val)
{
  if (!val || !val.startsWith('='))
    return val;

  if (evalCellInProgress[ref])
    return '0';

  var cached = evalCellCache[ref];
  if (cached && cached.val === val)
    return cached.result;

  evalCellInProgress[ref] = true;

  var out;

  try
  {
    var expr = substituteImplicitIntersection(expandArrayOps(substitutePageRefs(val.slice(1))), ref),
        previous,
        iterations = 25;

    do
    {
      previous = expr;
      expr = applyFunctionPass(expr, ref);

      // When no function matched (their arg regexes require paren-free
      // arguments), collapse innermost BARE (…) groups to values so e.g.
      // MAX((A1+1)*2, 5) flattens to MAX(4,5) for the next pass — plain
      // parenthesised sub-expressions resolve before anything around them.
      if (expr === previous)
        expr = resolveBareGroups(expr);
    }
    while (expr !== previous && --iterations > 0);

    expr = sanitizeExprForEval(substituteCellRefs(expr));

    var result = Function('"use strict"; return (' + expr + ')')();

    if (typeof result === 'boolean')
      out = result ? 'TRUE' : 'FALSE';
    else if (typeof result === 'string')
      out = result;
    else
      out =  (isNaN(result) || !isFinite(result))
             ?
               '#ERR'
             :
               parseFloat(result.toFixed(10)).toString();
  }
  catch(e)
  {
    out = '#ERR';
  }
  finally
  {
    delete evalCellInProgress[ref];
  }

  evalCellCache[ref] = { val: val, result: out };

  return out;
}

function evaluateFormulas(skipRef)
{
  // Fresh pass: anything might have changed since the last one.
  clearEvalCache();

  Object.keys(sheetData).forEach
  (
    function(ref)
    {
      // Skip the cell currently being typed into - overwriting it mid-edit
      // would clobber the raw formula text the user hasn't finished typing.
      if (ref === skipRef)
        return;

      var val=sheetData[ref];

      if(val && val.startsWith('='))
      {
        var inp = document.getElementById('inp-' + ref);
        if (inp)
        {
          inp.value = evalCell(ref, val);
          inp.classList.add('formula-result');
        }
      }
    }
  );

  refreshSheetCharts();
}

// Strips the cell-decoration suffixes off a serialized cell line's payload
// (everything after "REF="): a trailing `;#hex` paints the background, a
// trailing `;txt#hex` paints the text, and a trailing `:B2,B3` lists cells
// merged into this one. They may be chained in any order
// (B1=5:B2,B3;#FAFAFA == B1=5;#FAFAFA:B2,B3) — all parse the same because
// stripping repeats from the end until nothing matches.
function stripCellSuffixes(rest)
{
  var color = null,
      textColor = null,
      merged = null,
      changed = true;

  while (changed)
  {
    changed = false;

    var mt = rest.match(/;txt(#[0-9A-Fa-f]{3}(?:[0-9A-Fa-f]{3})?(?:[0-9A-Fa-f]{2})?)\s*$/);
    if (mt && textColor === null)
    {
      textColor = mt[1];
      rest = rest.slice(0, rest.length - mt[0].length);
      changed = true;
      continue;
    }

    var mc = rest.match(/;(#[0-9A-Fa-f]{3}(?:[0-9A-Fa-f]{3})?(?:[0-9A-Fa-f]{2})?)\s*$/);
    if (mc && color === null)
    {
      color = mc[1];
      rest = rest.slice(0, rest.length - mc[0].length);
      changed = true;
      continue;
    }

    var mm = rest.match(/:(\$?[A-Z]{1,2}\$?\d+(?:\s*,\s*\$?[A-Z]{1,2}\$?\d+)*)\s*$/);
    if (mm && merged === null)
    {
      merged = mm[1].split(',').map(function(t){ return t.trim().replace(/\$/g, ''); });
      rest = rest.slice(0, rest.length - mm[0].length);
      changed = true;
      continue;
    }
  }

  return { rest: rest, color: color, textColor: textColor, merged: merged };
}

// ── SHEET TABS (pages) ──

function renderSheetPageTabs()
{
  var bar = document.getElementById('sheet-tabs');
  if (!bar) return;

  bar.innerHTML = sheetPages.map(function(p, i)
  {
    // mousedown-default is prevented so a mid-formula edit keeps focus while
    // the user switches tabs to pick a cell from another page.
    return '<div class="sheet-tab' + (i === sheetPageIdx ? ' active' : '') + '" onmousedown="event.preventDefault()" onclick="switchSheetPage(' + i + ')" ondblclick="renameSheetPage(' + i + ')" title="Double-click to rename">' +
      escHtml(p.name) +
      (sheetPages.length > 1 ? '<button class="sheet-tab-del" onclick="deleteSheetPage(event,' + i + ')" title="Delete tab">×</button>' : '') +
    '</div>';
  }).join('') +
  '<button class="sheet-tab-add" onclick="addSheetPage()" title="New tab">+</button>' +
  '<span class="sheet-tab-hint">' +
    (formulaPickCtx
      ? 'Click a cell to insert it into the formula on “' + escHtml(sheetPages[formulaPickCtx.pageIdx].name) + '”'
      : 'Reference other tabs in formulas: [Tab name](A1)') +
  '</span>';
}

function aliasActiveSheetPage()
{
  var page = sheetPages[sheetPageIdx];
  if (!page.colors) page.colors = {};
  if (!page.textColors) page.textColors = {};
  if (!page.merges) page.merges = {};
  sheetData = page.data;
  sheetColors = page.colors;
  sheetTextColors = page.textColors;
  sheetMerges = page.merges;
  clearEvalCache();
}

function switchSheetPage(i)
{
  if (i === sheetPageIdx || !sheetPages[i]) return;

  // Leaving a tab mid-edit: keep the typed value, and if it's a formula,
  // arm pick mode so a cell click on the destination tab inserts a
  // [Tab name](REF) reference and returns here (see cellMouseDown). The tab
  // buttons prevent mousedown-default, so the editing input hasn't blurred.
  var pickCtx = null;
  if (editingCell)
  {
    var editInp = document.getElementById('inp-' + editingCell);
    if (editInp)
    {
      sheetData[editingCell] = editInp.value;
      if (editInp.value.startsWith('='))
        pickCtx = { pageIdx: sheetPageIdx, ref: editingCell, insertStart: editInp.selectionStart, insertLen: 0 };
      editInp.readOnly = true;
    }
    editingCell = null;
  }
  else
  {
    // Formula-BAR editing gets the same treatment: leave mid-formula, click
    // a cell on the destination tab, and come back to the bar.
    var pickBar = document.getElementById('formula-bar');
    if (document.activeElement === pickBar && pickBar.value.startsWith('='))
    {
      var barRef = document.getElementById('cell-ref').value;
      sheetData[barRef] = pickBar.value;
      pickCtx = { pageIdx: sheetPageIdx, ref: barRef, insertStart: pickBar.selectionStart, insertLen: 0, viaBar: true };
    }
  }

  sheetPageIdx = i;
  aliasActiveSheetPage();
  formulaPickCtx = pickCtx;
  // Undo snapshots belong to the page they were taken on; restoring one onto
  // a different page would corrupt it, so the stacks reset per page.
  sheetUndoStack = [];
  sheetRedoStack = [];

  renderSheetPageTabs();
  ensureGridFitsData(sheetData);
  refreshAllCellDisplays();
  selectCell('A1');
  renderSheetCharts();
}

function uniqueSheetPageName(base, excludeIdx)
{
  var name = base, n = 2;
  function taken(candidate)
  {
    return sheetPages.some(function(p, i){ return i !== excludeIdx && p.name.trim().toLowerCase() === candidate.trim().toLowerCase(); });
  }
  while (taken(name)) name = base + ' ' + n++;
  return name;
}

function addSheetPage()
{
  var page = { name: uniqueSheetPageName('Sheet ' + (sheetPages.length + 1), -1), data: {}, colors: {}, textColors: {}, merges: {} };
  sheetPages.push(page);
  switchSheetPage(sheetPages.length - 1);
  saveSheetToFile();
}

function renameSheetPage(i)
{
  var page = sheetPages[i];
  if (!page) return;

  openDataModal
  (
    'Rename tab',
    '<label class="field-label">Tab name<input class="modal-input" id="sheet-tab-name" value="' + escAttr(page.name) + '"></label>' +
    '<div style="color:var(--text3);font-size:12px;margin-top:8px">Formulas on other tabs reference this one by name — e.g. [' + escHtml(page.name) + '](A1) — and don\'t update automatically when it changes.</div>',
    function()
    {
      var name = document.getElementById('sheet-tab-name').value.trim();
      if (!name) return;
      var oldName = page.name;
      page.name = uniqueSheetPageName(name, i);
      // Charts follow their tab through renames.
      sheetCharts.forEach(function(c){ if (c.page === oldName) c.page = page.name; });
      closeDataModal();
      renderSheetPageTabs();
      refreshAllCellDisplays(); // refs to the old name now show #REF
      saveSheetToFile();
    }
  );
}

function deleteSheetPage(e, i)
{
  e.stopPropagation();
  if (sheetPages.length <= 1) return;
  if (!confirm('Delete tab "' + sheetPages[i].name + '" and everything on it?')) return;

  sheetPages.splice(i, 1);
  if (sheetPageIdx >= sheetPages.length) sheetPageIdx = sheetPages.length - 1;
  else if (i < sheetPageIdx) sheetPageIdx--;

  aliasActiveSheetPage();
  sheetCharts = sheetCharts.filter(function(c){ return !c.page || sheetPages.some(function(p){ return p.name === c.page; }); });
  sheetUndoStack = [];
  sheetRedoStack = [];

  renderSheetPageTabs();
  refreshAllCellDisplays();
  selectCell('A1');
  renderSheetCharts();
  saveSheetToFile();
}

function loadSheetFile(f)
{
  sheetData = {};
  sheetPages = [];
  sheetPageIdx = 0;
  sheetCharts = [];
  sheetUndoStack = [];
  sheetRedoStack = [];

  var lines = (f.content || '').split('\n'),
      dataStart = 0;

  if(lines[0]==='---')
  {
    var end = lines.indexOf('---',1);
    dataStart = end + 2;
  }

  // A '```charts' marker ends the cell-data section - everything from there
  // is the chart block, not cell lines. Found explicitly (rather than just
  // skipping any line without '=') since a chart's JSON could itself contain
  // an '=' (e.g. in a title) and would otherwise get misread as a cell.
  var cellLines = lines.slice(dataStart),
      chartsMarker = cellLines.indexOf('```charts');

  if (chartsMarker !== -1)
  {
    var chartsEnd = cellLines.indexOf('```', chartsMarker + 1);

    if (chartsEnd !== -1)
    {
      try
      {
        sheetCharts = JSON.parse(cellLines.slice(chartsMarker + 1, chartsEnd).join('\n')) || [];
      }
      catch(e)
      {
        console.warn('Sheet charts parse error', e);
        sheetCharts = [];
      }
    }

    cellLines = cellLines.slice(0, chartsMarker);
  }

  var curPage = null;

  cellLines.forEach
  (
    function(line)
    {
      if(!line.trim())
        return;

      // `# name` starts a new tab/page. Files from before multi-tab support
      // have no heading and fall into an implicit first page below.
      if (line.charAt(0) === '#')
      {
        curPage = { name: line.replace(/^#+\s*/, '').trim() || ('Sheet ' + (sheetPages.length + 1)), data: {}, colors: {}, textColors: {}, merges: {} };
        sheetPages.push(curPage);
        return;
      }

      var refEnd = line.indexOf('=');

      if(refEnd === -1)
        return;

      if (!curPage)
      {
        curPage = { name: 'Sheet 1', data: {}, colors: {}, textColors: {}, merges: {} };
        sheetPages.push(curPage);
      }

      var ref = line.slice(0, refEnd),
          stripped = stripCellSuffixes(line.slice(refEnd + 1)),
          rest = stripped.rest,
          formulaEnd = rest.indexOf('=');

      if (stripped.color)     curPage.colors[ref] = stripped.color;
      if (stripped.textColor) curPage.textColors[ref] = stripped.textColor;
      if (stripped.merged)    curPage.merges[ref] = stripped.merged;

      // Formula cells are written as ref=formula=result; the result is informational
      // (recomputed on load), so only the formula half is kept, with '=' restored.
      if (rest !== '')
        curPage.data[ref] = (formulaEnd === -1) ? rest : ('=' + rest.slice(0, formulaEnd));
    }
  );

  if (!sheetPages.length)
    sheetPages = [{ name: 'Sheet 1', data: {}, colors: {}, textColors: {}, merges: {} }];

  sheetPageIdx = 0;
  aliasActiveSheetPage();

  // Charts made before per-tab charts existed carry no page name; pin them
  // to the first tab so they don't float over every page.
  sheetCharts.forEach(function(c){ if (!c.page) c.page = sheetPages[0].name; });

  renderSheetPageTabs();
  ensureGridFitsData(sheetData);
  refreshAllCellDisplays();
  selectCell('A1');
  renderSheetCharts();
}

function refreshAllCellDisplays()
{
  clearEvalCache();

  for (var i = 1; i <= ROWS; i++)
    for (var j = 0; j < COLS; j++)
    {
      var ref = colName(j) + i,
          v = sheetData[ref] || '',
          inp = document.getElementById('inp-' + ref);

      if (inp)
      {
        inp.value = v.startsWith('=')
                    ?
                      evalCell(ref, v)
                    :
                      v;

        inp.classList.toggle('formula-result', v.startsWith('='));
      }
    }

  applyCellDecorations();
}

// ── CELL DECORATIONS (background colors + merged cells) ──

// Refs whose <td> was touched by the previous decoration pass, so a page
// switch / undo / unmerge can cheaply reset exactly those cells.
var decoratedCellRefs = [];

function applyCellDecorations()
{
  decoratedCellRefs.forEach(function(ref)
  {
    var td = document.getElementById('cell-' + ref);
    if (!td) return;
    td.style.background = '';
    td.style.display = '';
    td.removeAttribute('colspan');
    td.removeAttribute('rowspan');
    // Text color lives on the input (inputs don't inherit the td's color).
    var inp = document.getElementById('inp-' + ref);
    if (inp) inp.style.color = '';
  });
  decoratedCellRefs = [];

  Object.keys(sheetTextColors).forEach(function(ref)
  {
    var inp = document.getElementById('inp-' + ref);
    if (!inp || !sheetTextColors[ref]) return;
    inp.style.color = sheetTextColors[ref];
    decoratedCellRefs.push(ref);
  });

  Object.keys(sheetColors).forEach(function(ref)
  {
    var td = document.getElementById('cell-' + ref);
    if (!td || !sheetColors[ref]) return;
    td.style.background = sheetColors[ref];
    decoratedCellRefs.push(ref);
  });

  Object.keys(sheetMerges).forEach(function(anchor)
  {
    var covered = sheetMerges[anchor];
    if (!covered || !covered.length) return;

    // The merge is rendered as the bounding rectangle of anchor + covered
    // cells (colspan/rowspan on the anchor, everything else hidden), so even
    // a hand-written non-rectangular list ends up displayable.
    var pa = parseName(anchor);
    if (!pa) return;
    var minRow = pa.row, maxRow = pa.row,
        minCol = colIndex(pa.col), maxCol = colIndex(pa.col);

    covered.forEach(function(ref)
    {
      var pc = parseName(ref);
      if (!pc) return;
      minRow = Math.min(minRow, pc.row); maxRow = Math.max(maxRow, pc.row);
      minCol = Math.min(minCol, colIndex(pc.col)); maxCol = Math.max(maxCol, colIndex(pc.col));
    });

    ensureGridFits(maxRow, maxCol);

    var anchorTd = document.getElementById('cell-' + anchor);
    if (!anchorTd) return;

    anchorTd.colSpan = maxCol - minCol + 1;
    anchorTd.rowSpan = maxRow - minRow + 1;
    decoratedCellRefs.push(anchor);

    for (var r = minRow; r <= maxRow; r++)
      for (var c = minCol; c <= maxCol; c++)
      {
        var ref = colName(c) + r;
        if (ref === anchor) continue;
        var td = document.getElementById('cell-' + ref);
        if (td)
        {
          td.style.display = 'none';
          decoratedCellRefs.push(ref);
        }
      }
  });
}

// Applies (or clears, with null) a background color to every cell in the
// current selection rectangle.
function applyCellColor(color)
{
  if (!selectionAnchor) return;
  sheetSnapshotForUndo();

  var box = getSelectionBoundingBox();
  for (var r = box.rowStart; r <= box.rowEnd; r++)
    for (var c = box.colStart; c <= box.colEnd; c++)
    {
      var ref = colName(c) + r;
      if (color) sheetColors[ref] = color;
      else delete sheetColors[ref];
    }

  applyCellDecorations();
  saveSheetToFile();
}

function clearCellColor() { applyCellColor(null); }

// Applies (or clears, with null) a text color to every cell in the current
// selection rectangle — same shape as applyCellColor, different map.
function applyCellTextColor(color)
{
  if (!selectionAnchor) return;
  sheetSnapshotForUndo();

  var box = getSelectionBoundingBox();
  for (var r = box.rowStart; r <= box.rowEnd; r++)
    for (var c = box.colStart; c <= box.colEnd; c++)
    {
      var ref = colName(c) + r;
      if (color) sheetTextColors[ref] = color;
      else delete sheetTextColors[ref];
    }

  applyCellDecorations();
  saveSheetToFile();
}

function clearCellTextColor() { applyCellTextColor(null); }

// ── Bucket split-button ──
// The bucket itself paints the current selection with the last-chosen color;
// the thin arrow beside it opens the color chooser. Choosing a color both
// remembers it (per browser) and applies it right away.

var currentCellColor = (function()
{
  try { return localStorage.getItem('lk_bucket_color') || '#3a5a40'; }
  catch(e) { return '#3a5a40'; }
})();

function paintBucketSwatch()
{
  var sw = document.getElementById('bucket-swatch');
  if (sw) sw.style.background = currentCellColor;
  var inp = document.getElementById('cell-color-input');
  if (inp) inp.value = currentCellColor;
}

function applyBucketColor()
{
  applyCellColor(currentCellColor);
}

function onBucketColorChosen(color)
{
  currentCellColor = color;
  try { localStorage.setItem('lk_bucket_color', color); } catch(e) {}
  paintBucketSwatch();
  applyCellColor(color);
}

// ── Text-color split-button ──
// Works exactly like the bucket above, but colors the cells' text instead of
// their background.

var currentTextColor = (function()
{
  try { return localStorage.getItem('lk_text_color') || '#d4a843'; }
  catch(e) { return '#d4a843'; }
})();

function paintTextSwatch()
{
  var sw = document.getElementById('text-swatch');
  if (sw) sw.style.background = currentTextColor;
  var inp = document.getElementById('text-color-input');
  if (inp) inp.value = currentTextColor;
}

function applyTextColor()
{
  applyCellTextColor(currentTextColor);
}

function onTextColorChosen(color)
{
  currentTextColor = color;
  try { localStorage.setItem('lk_text_color', color); } catch(e) {}
  paintTextSwatch();
  applyCellTextColor(color);
}

// Merges the selected rectangle into one big cell (top-left becomes the
// anchor), or splits it back apart when the selection sits on an existing
// merge.
function toggleMergeSelection()
{
  if (!selectionAnchor) return;

  var box = getSelectionBoundingBox();

  // Unmerge: any existing anchor inside the selection wins.
  var existing = Object.keys(sheetMerges).find(function(anchor)
  {
    var pa = parseName(anchor);
    return pa && pa.row >= box.rowStart && pa.row <= box.rowEnd &&
           colIndex(pa.col) >= box.colStart && colIndex(pa.col) <= box.colEnd;
  });

  sheetSnapshotForUndo();

  if (existing)
  {
    delete sheetMerges[existing];
  }
  else
  {
    if (box.rowStart === box.rowEnd && box.colStart === box.colEnd)
      return; // a single cell has nothing to merge with

    var anchor = colName(box.colStart) + box.rowStart,
        covered = [];

    for (var r = box.rowStart; r <= box.rowEnd; r++)
      for (var c = box.colStart; c <= box.colEnd; c++)
      {
        var ref = colName(c) + r;
        if (ref !== anchor) covered.push(ref);
      }

    sheetMerges[anchor] = covered;
  }

  applyCellDecorations();
  saveSheetToFile();
  selectCell(existing || (colName(box.colStart) + box.rowStart));
}

function sheetSnapshotForUndo()
{
  sheetUndoStack.push({ data: JSON.parse(JSON.stringify(sheetData)), charts: JSON.parse(JSON.stringify(sheetCharts)), colors: JSON.parse(JSON.stringify(sheetColors)), textColors: JSON.parse(JSON.stringify(sheetTextColors)), merges: JSON.parse(JSON.stringify(sheetMerges)) });
  sheetRedoStack = [];

  if (sheetUndoStack.length > 100)
    sheetUndoStack.shift();
}

function sheetRestoreSnapshot(fromStack, toStack)
{
  if (!fromStack.length)
    return;

  toStack.push({ data: JSON.parse(JSON.stringify(sheetData)), charts: JSON.parse(JSON.stringify(sheetCharts)), colors: JSON.parse(JSON.stringify(sheetColors)), textColors: JSON.parse(JSON.stringify(sheetTextColors)), merges: JSON.parse(JSON.stringify(sheetMerges)) });

  var snap = fromStack.pop();

  sheetData = snap.data;
  sheetColors = snap.colors || {};
  sheetTextColors = snap.textColors || {};
  sheetMerges = snap.merges || {};
  var snapPage = sheetPages[sheetPageIdx]; // keep the page aliases in sync
  snapPage.data = sheetData;
  snapPage.colors = sheetColors;
  snapPage.textColors = sheetTextColors;
  snapPage.merges = sheetMerges;
  sheetCharts = snap.charts;
  ensureGridFitsData(sheetData);

  // Mid-edit cell state belongs to whatever was being typed, which the
  // snapshot doesn't track - drop back to read-only rather than leaving a
  // stale editable input pointed at a ref that may no longer make sense.
  if (editingCell)
  {
    var inp = document.getElementById('inp-' + editingCell);

    if (inp)
      inp.readOnly = true;

    editingCell = null;
  }

  refreshAllCellDisplays();
  renderSheetCharts();
  evaluateFormulas();
  saveSheetToFile();
  refreshRefHighlights();
}

function sheetUndo()
{
  sheetRestoreSnapshot(sheetUndoStack, sheetRedoStack);
}

function sheetRedo()
{
  sheetRestoreSnapshot(sheetRedoStack, sheetUndoStack);
}

// Serializing + persisting on every step of a drag (ref-box moves, fill
// drags, chart drags all save as they go) is wasted work: mark the sheet
// dirty instead and let the post-drag flush below write once.
var sheetSaveDeferred = false;

// Persisting (disk/network write, backlink scan, sidebar re-render) is also
// too heavy to run per keystroke — files[..].content is always updated
// synchronously, only the persist itself trails behind.
var sheetPersistTimer = null;
var sheetPersistId = null;

function scheduleSheetPersist(id)
{
  sheetPersistId = id;
  clearTimeout(sheetPersistTimer);
  sheetPersistTimer = setTimeout(function()
  {
    sheetPersistTimer = null;
    persistFileEntry(id);
  }, 400);
}

window.addEventListener('beforeunload', function()
{
  if (sheetPersistTimer)
  {
    clearTimeout(sheetPersistTimer);
    sheetPersistTimer = null;
    persistFileEntry(sheetPersistId);
  }
});

function saveSheetToFile()
{
  // The grid is always live even when no spreadsheet is open, so without the
  // type check an empty sheet would serialize itself over whatever file
  // happens to be current (e.g. the document opened before switching tabs).
  if (!currentFileId || !files[currentFileId] || files[currentFileId].type !== 'sheet')
    return;

  if (sheetDragActive())
  {
    sheetSaveDeferred = true;
    return;
  }

  var out = '---\ntype: spreadsheet\n---\n\n';

  sheetPages.forEach(function(page, p)
  {
    if (p > 0) out += '\n';
    out += '# ' + page.name + '\n';

    // Iterate the page's own cells (row-major) rather than the grid bounds:
    // the grid is sized for the ACTIVE page, and clipping another page's
    // far-out cells to it would silently drop them from the file. Cells with
    // only a color/merge (no value) still get a line, so the decoration keeps.
    var colors = page.colors || {},
        textColors = page.textColors || {},
        merges = page.merges || {},
        cellKeys = {};
    Object.keys(page.data).forEach(function(ref){ if (page.data[ref]) cellKeys[ref] = true; });
    Object.keys(colors).forEach(function(ref){ if (colors[ref]) cellKeys[ref] = true; });
    Object.keys(textColors).forEach(function(ref){ if (textColors[ref]) cellKeys[ref] = true; });
    Object.keys(merges).forEach(function(ref){ if (merges[ref] && merges[ref].length) cellKeys[ref] = true; });

    // Alias the page in once for the whole loop (rather than per formula via
    // evalCellOnPage) so the eval cache holds across its cells.
    var savedData = sheetData;
    sheetData = page.data;
    if (savedData !== page.data) clearEvalCache();

    Object.keys(cellKeys)
      .map(function(ref){ return { ref: ref, at: parseName(ref) }; })
      .filter(function(c){ return !!c.at; })
      .sort(function(a, b){ return a.at.row - b.at.row || colIndex(a.at.col) - colIndex(b.at.col); })
      .forEach(function(c)
      {
        var val = page.data[c.ref] || '';
        var cellLine = val.startsWith('=')
                       ? c.ref + '=' + val.slice(1) + '=' + evalCell(c.ref, val)
                       : c.ref + '=' + val;
        if (merges[c.ref] && merges[c.ref].length) cellLine += ':' + merges[c.ref].join(',');
        if (colors[c.ref]) cellLine += ';' + colors[c.ref];
        if (textColors[c.ref]) cellLine += ';txt' + textColors[c.ref];
        out += cellLine + '\n';
      });

    sheetData = savedData;
    if (savedData !== page.data) clearEvalCache();
  });

  if (sheetCharts.length)
    out += '\n```charts\n' + JSON.stringify(sheetCharts) + '\n```\n';

  files[currentFileId].content = out;
  files[currentFileId].modified = Date.now();

  scheduleSheetPersist(currentFileId);
}

// ── SHEET CHARTS (Chart.js) ──

var CHART_COLORS = ['#d4a843', '#4a7fa8', '#5a9a6e', '#c0574a', '#9d6fd1', '#d18b3f', '#3fa7b0'];

function configureChartJs()
{
  if (typeof Chart === 'undefined')
    return;

  Chart.register
  (
    Chart.BarController, Chart.LineController, Chart.PieController, Chart.DoughnutController,
    Chart.RadarController, Chart.ScatterController,
    Chart.BarElement, Chart.LineElement, Chart.PointElement, Chart.ArcElement,
    Chart.CategoryScale, Chart.LinearScale, Chart.RadialLinearScale,
    Chart.Legend, Chart.Title, Chart.Tooltip, Chart.Filler
  );
}

configureChartJs();

function parseChartRange(range)
{
  var parts = (range || '').split(':');

  if (parts.length !== 2)
    return null;

  var pa = parseName(parts[0].trim()),
      pb = parseName(parts[1].trim());

  if (!pa || !pb)
    return null;

  return  {
            colStart: Math.min(colIndex(pa.col), colIndex(pb.col)),
            colEnd: Math.max(colIndex(pa.col), colIndex(pb.col)),
            rowStart: Math.min(pa.row, pb.row),
            rowEnd: Math.max(pa.row, pb.row)
          };
}

function getChartRangeGrid(range)
{
  var dims = parseChartRange(range);

  if (!dims)
    return null;

  var grid = [];

  for (var r = dims.rowStart; r <= dims.rowEnd; r++)
  {
    var row = [];

    for (var c = dims.colStart; c <= dims.colEnd; c++)
      row.push(getDisplayValue(colName(c) + r));

    grid.push(row);
  }

  return grid;
}

// Standard Excel/Sheets table orientation: row 1 -> category labels,
// column 1 -> series names, the rest -> the data matrix.
// Categories from row 1 (across columns), series from column 1 (down rows) -
// the right shape for a "Quarter columns x Product rows" comparison table.
function categoricalByRowSeries(grid, firstRow, firstCol)
{
  var rowStart = firstRow ? 1 : 0,
      colStart = firstCol ? 1 : 0,
      labels = [],
      datasets = [],
      c, r;

  for (c = colStart; c < grid[0].length; c++)
    labels.push(firstRow ? grid[0][c] : ('Col ' + (c + 1)));

  for (r = rowStart; r < grid.length; r++)
  {
    var values = [];

    for (c = colStart; c < grid[r].length; c++)
      values.push(parseFloat(grid[r][c]) || 0);

    datasets.push({ label: firstCol ? grid[r][0] : ('Series ' + (r + 1 - rowStart)), data: values });
  }

  return { labels: labels, datasets: datasets };
}

// Categories from column 1 (down rows), series from row 1 (across columns) -
// the right shape for a simple "Month rows x Revenue column(s)" list, which
// is the more common case for small everyday tables.
function categoricalByColumnSeries(grid, firstRow, firstCol)
{
  var rowStart = firstRow ? 1 : 0,
      colStart = firstCol ? 1 : 0,
      labels = [],
      datasets = [],
      c, r;

  for (r = rowStart; r < grid.length; r++)
    labels.push(firstCol ? grid[r][0] : ('Row ' + (r + 1)));

  for (c = colStart; c < grid[0].length; c++)
  {
    var values = [];

    for (r = rowStart; r < grid.length; r++)
      values.push(parseFloat(grid[r][c]) || 0);

    datasets.push({ label: firstRow ? grid[0][c] : ('Series ' + (c + 1 - colStart)), data: values });
  }

  return { labels: labels, datasets: datasets };
}

// The effective orientation of a chart: an explicit choice wins; 'auto'
// matches Excel/Sheets' default — a range with more data rows than data
// columns (e.g. a simple Month/Revenue list) charts each COLUMN as a series,
// otherwise (e.g. a wide Quarter-by-Product table) each ROW is a series.
function chartEffectiveOrientation(chartDef, grid)
{
  if (chartDef.orientation === 'rows' || chartDef.orientation === 'cols')
    return chartDef.orientation;

  var dataRowCount = grid.length - (chartDef.firstRowLabels ? 1 : 0),
      dataColCount = grid[0].length - (chartDef.firstColSeries ? 1 : 0);

  return (dataRowCount > dataColCount) ? 'cols' : 'rows';
}

function buildCategoricalChartData(chartDef, grid)
{
  if (!grid.length || !grid[0].length)
    return { labels: [], datasets: [] };

  var firstRow = chartDef.firstRowLabels,
      firstCol = chartDef.firstColSeries;

  // 'rows': each row is a series (legend from the first column, axis labels
  // from the first row). 'cols': each column is a series (legend from the
  // first row, axis labels from the first column).
  return  chartEffectiveOrientation(chartDef, grid) === 'cols'
          ?
            categoricalByColumnSeries(grid, firstRow, firstCol)
          :
            categoricalByRowSeries(grid, firstRow, firstCol);
}

// Scatter wants {x,y} pairs rather than labels+parallel arrays: column 1 is
// the shared X axis, every other column is its own series plotted against it.
function buildScatterChartData(chartDef, grid)
{
  var firstRow = chartDef.firstRowLabels,
      rowStart = firstRow ? 1 : 0,
      datasets = [];

  if (!grid.length || grid[0].length < 2)
    return { labels: [], datasets: [] };

  for (var c = 1; c < grid[0].length; c++)
  {
    var points = [];

    for (var r = rowStart; r < grid.length; r++)
    {
      var x = parseFloat(grid[r][0]),
          y = parseFloat(grid[r][c]);

      if (!isNaN(x) && !isNaN(y))
        points.push({ x: x, y: y });
    }

    datasets.push({ label: firstRow ? grid[0][c] : ('Series ' + c), data: points });
  }

  return { labels: [], datasets: datasets };
}

function buildChartData(chartDef)
{
  var grid = getChartRangeGrid(chartDef.range);

  if (!grid || !grid.length)
    return { labels: [], datasets: [] };

  if (chartDef.type === 'scatter')
    return buildScatterChartData(chartDef, grid);

  var data = buildCategoricalChartData(chartDef, grid);

  // Pie/doughnut are single-series chart types - if the range has more than
  // one data row, only the first is used rather than silently merging them.
  if (chartDef.type === 'pie' || chartDef.type === 'doughnut')
    data.datasets = data.datasets.slice(0, 1);

  if (chartDef.type === 'percent')
    data = percentStackDatasets(data, chartDef);

  return data;
}

// % Stacked: absolute values render as segments of 100%-stacked columns,
// colored per series — a pie/doughnut reading, but as bars. A range with a
// SINGLE data series pivots into one bar whose segments are the categories
// (each category becomes its own colored, legended dataset), the closest
// bar-shaped equivalent of a pie chart; multi-series ranges normalize each
// category's stack. chartDef.percentTotal, when given, replaces the per-bar
// sum as the 100% denominator, so bars may then top out below 100%.
function percentStackDatasets(data, chartDef)
{
  if (!data.datasets.length)
    return data;

  function pct(v, total)
  {
    return total ? Math.round(((parseFloat(v) || 0) / total) * 1000) / 10 : 0;
  }

  if (data.datasets.length === 1)
  {
    var ds = data.datasets[0],
        total = chartDef.percentTotal ||
                ds.data.reduce(function(a, b){ return a + (parseFloat(b) || 0); }, 0);

    return {
      labels: [ds.label || 'Total'],
      datasets: data.labels.map(function(lab, i)
      {
        return { label: String(lab), data: [pct(ds.data[i], total)] };
      })
    };
  }

  var categoryTotals = data.labels.map(function(_, j)
  {
    return chartDef.percentTotal ||
           data.datasets.reduce(function(a, ds){ return a + (parseFloat(ds.data[j]) || 0); }, 0);
  });

  return {
    labels: data.labels,
    datasets: data.datasets.map(function(ds)
    {
      return {
        label: ds.label,
        data: ds.data.map(function(v, j){ return pct(v, categoryTotals[j]); })
      };
    })
  };
}

function buildChartScales(chartDef)
{
  var gridColor = 'rgba(255,255,255,0.08)',
      tickColor = '#9e9b94';

  if (chartDef.type === 'radar')
  {
    return { r: { angleLines: { color: gridColor }, grid: { color: gridColor }, pointLabels: { color: tickColor }, ticks: { color: tickColor, backdropColor: 'transparent' } } };
  }

  if (chartDef.type === 'scatter')
  {
    return  {
              x: { type: 'linear', position: 'bottom', grid: { color: gridColor }, ticks: { color: tickColor } },
              y: { grid: { color: gridColor }, ticks: { color: tickColor } }
            };
  }

  if (chartDef.type === 'percent')
  {
    return  {
              x: { grid: { color: gridColor }, ticks: { color: tickColor }, stacked: true },
              y: { beginAtZero: true, grid: { color: gridColor },
                   ticks: { color: tickColor, callback: function(v){ return v + '%'; } }, stacked: true }
            };
  }

  return  {
            x: { grid: { color: gridColor }, ticks: { color: tickColor } },
            y: { grid: { color: gridColor }, ticks: { color: tickColor } }
          };
}

function buildChartJsConfig(chartDef)
{
  var data = buildChartData(chartDef),
      isPie = (chartDef.type === 'pie' || chartDef.type === 'doughnut'),
      baseType = (chartDef.type === 'combo' || chartDef.type === 'percent') ? 'bar' : chartDef.type;

  var datasets = data.datasets.map(function(ds, i)
  {
    var color = CHART_COLORS[i % CHART_COLORS.length],
        styled =
        {
          label: ds.label,
          data: ds.data,
          backgroundColor: isPie ? ds.data.map(function(_, j){ return CHART_COLORS[j % CHART_COLORS.length]; }) : color,
          borderColor: color,
          borderWidth: (chartDef.type === 'line' || chartDef.type === 'radar') ? 2 : 1
        };

    if (chartDef.type === 'combo')
      styled.type = (i === 0) ? 'bar' : 'line';

    if (chartDef.type === 'line' || chartDef.type === 'radar' || (chartDef.type === 'combo' && i > 0))
    {
      styled.fill = false;
      styled.tension = 0.25;
    }

    return styled;
  });

  return  {
            type: baseType,
            data: { labels: data.labels, datasets: datasets },
            options:
            {
              responsive: true,
              maintainAspectRatio: false,
              animation: false,
              plugins:
              {
                title: { display: !!chartDef.title, text: chartDef.title, color: '#f0ede6' },
                legend: { display: datasets.length > 1 || isPie, labels: { color: '#9e9b94' } },
                tooltip: (chartDef.type === 'percent')
                  ? { callbacks: { label: function(ctx){ return (ctx.dataset.label ? ctx.dataset.label + ': ' : '') + ctx.parsed.y + '%'; } } }
                  : {}
              },
              scales: isPie ? {} : buildChartScales(chartDef)
            }
          };
}

function renderSheetCharts()
{
  var layer = document.getElementById('sheet-charts-layer');

  if (!layer)
    return;

  Object.values(sheetChartInstances).forEach(function(chart){ chart.destroy(); });
  sheetChartInstances = {};
  layer.innerHTML = '';

  // Only the active tab's charts render; switching tabs re-runs this.
  var activePage = sheetPages[sheetPageIdx].name;

  sheetCharts.forEach(function(chartDef)
  {
    if (chartDef.page && chartDef.page !== activePage)
      return;
    var box = document.createElement('div');
    box.className = 'sheet-chart-box';
    box.dataset.chartId = chartDef.id;
    box.style.left = chartDef.x + 'px';
    box.style.top = chartDef.y + 'px';
    box.style.width = chartDef.width + 'px';
    box.style.height = chartDef.height + 'px';

    var swapBtn = chartDef.type === 'scatter' ? '' :
        '<span class="sheet-chart-del" title="Swap legend and axis (rows ↔ columns)" onmousedown="event.stopPropagation()" onclick="toggleChartOrientation(event,\'' + chartDef.id + '\')">' +
          '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="17,3 21,7 17,11"/><line x1="21" y1="7" x2="9" y2="7"/><polyline points="7,13 3,17 7,21"/><line x1="3" y1="17" x2="15" y2="17"/></svg>' +
        '</span>';

    box.innerHTML =
      '<div class="sheet-chart-header" onmousedown="chartBoxMouseDown(event,\'' + chartDef.id + '\')">' +
        '<span class="sheet-chart-title">' + escHtml(chartDef.title || 'Chart') + '</span>' +
        swapBtn +
        '<span class="sheet-chart-del" onmousedown="event.stopPropagation()" onclick="deleteChart(event,\'' + chartDef.id + '\')">' +
          '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>' +
        '</span>' +
      '</div>' +
      '<div class="sheet-chart-canvas-wrap"><canvas></canvas></div>' +
      '<div class="sheet-chart-resize-handle" onmousedown="chartResizeMouseDown(event,\'' + chartDef.id + '\')"></div>';

    layer.appendChild(box);

    if (typeof Chart === 'undefined')
      return;

    var canvas = box.querySelector('canvas');
    sheetChartInstances[chartDef.id] = new Chart(canvas.getContext('2d'), buildChartJsConfig(chartDef));
  });
}

function refreshSheetCharts()
{
  sheetCharts.forEach(function(chartDef)
  {
    var chart = sheetChartInstances[chartDef.id];

    if (!chart)
      return;

    var data = buildChartData(chartDef);

    chart.data.labels = data.labels;

    data.datasets.forEach(function(ds, i)
    {
      if (chart.data.datasets[i])
      {
        chart.data.datasets[i].data = ds.data;
        chart.data.datasets[i].label = ds.label;
      }
    });

    chart.update();
  });
}

function openChartModal()
{
  newChartType = 'bar';

  document.querySelectorAll('#chart-modal .type-card').forEach(function(card)
  {
    card.classList.toggle('selected', card.dataset.type === 'bar');
  });

  var box = selectionAnchor ? getSelectionBoundingBox() : null,
      prefill = box ? (colName(box.colStart) + box.rowStart + ':' + colName(box.colEnd) + box.rowEnd) : '';

  document.getElementById('chart-range-input').value = prefill;
  document.getElementById('chart-title-input').value = '';
  document.getElementById('chart-percent-total').value = '';
  document.getElementById('chart-percent-total').style.display = 'none';
  document.getElementById('chart-first-row-labels').checked = true;
  document.getElementById('chart-first-col-series').checked = true;
  document.getElementById('chart-orientation').value = 'auto';

  document.getElementById('chart-modal').classList.add('open');
}

function closeChartModal()
{
  document.getElementById('chart-modal').classList.remove('open');
}

function selectChartType(type, el)
{
  newChartType = type;

  document.querySelectorAll('#chart-modal .type-card').forEach(function(c){ c.classList.remove('selected'); });
  el.classList.add('selected');

  // % Columns converts absolute values into percentages of a total; the
  // total is optional (defaults to each series' own sum).
  document.getElementById('chart-percent-total').style.display = (type === 'percent') ? '' : 'none';
}

document.getElementById('chart-modal').addEventListener
(
  'click',
  function(e)
  {
    if (e.target === document.getElementById('chart-modal'))
      closeChartModal();
  }
);

function createChart()
{
  var range = document.getElementById('chart-range-input').value.trim();

  if (!range || !parseChartRange(range))
  {
    alert('Enter a valid range, e.g. A1:C5');
    return;
  }

  sheetSnapshotForUndo();

  var offset = (sheetCharts.length % 4) * 30;

  var percentTotal = parseFloat(document.getElementById('chart-percent-total').value);

  sheetCharts.push
  (
    {
      id: 'chart_' + Date.now(),
      type: newChartType,
      title: document.getElementById('chart-title-input').value.trim(),
      range: range,
      // Charts belong to the tab they were created on and only render there.
      page: sheetPages[sheetPageIdx].name,
      percentTotal: (newChartType === 'percent' && !isNaN(percentTotal) && percentTotal > 0) ? percentTotal : null,
      firstRowLabels: document.getElementById('chart-first-row-labels').checked,
      firstColSeries: document.getElementById('chart-first-col-series').checked,
      orientation: document.getElementById('chart-orientation').value,
      x: 40 + offset,
      y: 40 + offset,
      width: 420,
      height: 280
    }
  );

  closeChartModal();
  renderSheetCharts();
  saveSheetToFile();
}

// Flips which table direction a chart uses as its legend (Excel's "Switch
// Row/Column"). An 'auto' chart resolves to what it's currently showing
// first, so the click always visibly flips it.
function toggleChartOrientation(e, chartId)
{
  e.stopPropagation();

  var chartDef = sheetCharts.find(function(c){ return c.id === chartId; });
  if (!chartDef)
    return;

  var grid = getChartRangeGrid(chartDef.range);
  if (!grid || !grid.length || !grid[0].length)
    return;

  sheetSnapshotForUndo();
  chartDef.orientation = chartEffectiveOrientation(chartDef, grid) === 'rows' ? 'cols' : 'rows';

  renderSheetCharts();
  saveSheetToFile();
}

function deleteChart(e, chartId)
{
  e.stopPropagation();

  if (!confirm('Delete this chart?'))
    return;

  var idx = sheetCharts.findIndex(function(c){ return c.id === chartId; });

  if (idx === -1)
    return;

  sheetSnapshotForUndo();
  sheetCharts.splice(idx, 1);

  if (sheetChartInstances[chartId])
  {
    sheetChartInstances[chartId].destroy();
    delete sheetChartInstances[chartId];
  }

  var box = document.querySelector('.sheet-chart-box[data-chart-id="' + chartId + '"]');

  if (box)
    box.remove();

  saveSheetToFile();
}

function chartBoxMouseDown(e, chartId)
{
  e.preventDefault();

  var chartDef = sheetCharts.find(function(c){ return c.id === chartId; });

  if (!chartDef)
    return;

  sheetSnapshotForUndo();
  sheetChartDrag = { kind: 'move', id: chartId, startX: e.clientX, startY: e.clientY, origX: chartDef.x, origY: chartDef.y };
}

function chartResizeMouseDown(e, chartId)
{
  e.preventDefault();
  e.stopPropagation();

  var chartDef = sheetCharts.find(function(c){ return c.id === chartId; });

  if (!chartDef)
    return;

  sheetSnapshotForUndo();
  sheetChartDrag = { kind: 'resize', id: chartId, startX: e.clientX, startY: e.clientY, origWidth: chartDef.width, origHeight: chartDef.height };
}

document.addEventListener
(
  'mousemove',
  function(e)
  {
    if (!sheetChartDrag)
      return;

    var chartDef = sheetCharts.find(function(c){ return c.id === sheetChartDrag.id; });

    if (!chartDef)
      return;

    var box = document.querySelector('.sheet-chart-box[data-chart-id="' + sheetChartDrag.id + '"]');

    if (sheetChartDrag.kind === 'move')
    {
      chartDef.x = Math.max(0, sheetChartDrag.origX + (e.clientX - sheetChartDrag.startX));
      chartDef.y = Math.max(0, sheetChartDrag.origY + (e.clientY - sheetChartDrag.startY));

      if (box)
      {
        box.style.left = chartDef.x + 'px';
        box.style.top = chartDef.y + 'px';
      }
    }
    else if (sheetChartDrag.kind === 'resize')
    {
      chartDef.width = Math.max(220, sheetChartDrag.origWidth + (e.clientX - sheetChartDrag.startX));
      chartDef.height = Math.max(160, sheetChartDrag.origHeight + (e.clientY - sheetChartDrag.startY));

      if (box)
      {
        box.style.width = chartDef.width + 'px';
        box.style.height = chartDef.height + 'px';
      }

      var chart = sheetChartInstances[sheetChartDrag.id];

      if (chart)
        chart.resize();
    }
  }
);

document.addEventListener
(
  'mouseup',
  function()
  {
    if (!sheetChartDrag)
      return;

    sheetChartDrag = null;
    saveSheetToFile();
  }
);

// Deferred-save flush. Registered after every drag's own mouseup handler, so
// by the time it runs the drag flags are cleared and the save goes through.
document.addEventListener
(
  'mouseup',
  function()
  {
    if (!sheetSaveDeferred || sheetDragActive())
      return;

    sheetSaveDeferred = false;
    saveSheetToFile();
  }
);

// ── FUNCTION REFERENCE ──
var SHEET_FUNCTION_GROUPS =
[
  {
    category: 'Aggregate',
    functions:
    [
      { name: 'SUM(range)', desc: 'Adds all numbers in a range.', example: '=SUM(A1:A4)' },
      { name: 'AVG(range)', desc: 'Averages the numbers in a range.', example: '=AVG(A1:A4)' },
      { name: 'AVERAGE(range)', desc: 'Same as AVG.', example: '=AVERAGE(A1:A4)' },
      { name: 'MEAN(range)', desc: 'Same as AVG.', example: '=MEAN(A1:A4)' },
      { name: 'COUNTIF(range, criterion)', desc: 'Counts cells matching a criterion like ">5" or "text".', example: '=COUNTIF(A1:A9, ">5")' },
      { name: 'SUMIF(range, criterion, [sum_range])', desc: 'Sums the cells (or sum_range) where the criterion matches.', example: '=SUMIF(A1:A9, ">5", B1:B9)' },
      { name: 'AVERAGEIF(range, criterion, [avg_range])', desc: 'Averages the cells where the criterion matches.', example: '=AVERAGEIF(A1:A9, "<>0")' },
      { name: 'LARGE(range, k)', desc: 'The k-th largest value in a range.', example: '=LARGE(A1:A9, 2)' },
      { name: 'SMALL(range, k)', desc: 'The k-th smallest value in a range.', example: '=SMALL(A1:A9, 2)' },
      { name: 'MEDIAN(range)', desc: 'Middle value of a range.', example: '=MEDIAN(A1:A4)' },
      { name: 'PRODUCT(range)', desc: 'Multiplies all numbers in a range.', example: '=PRODUCT(A1:A4)' },
      { name: 'MAX(range)', desc: 'Largest value in a range.', example: '=MAX(A1:A4)' },
      { name: 'MIN(range)', desc: 'Smallest value in a range.', example: '=MIN(A1:A4)' },
      { name: 'COUNT(range)', desc: 'Counts the cells in a range.', example: '=COUNT(A1:A4)' },
      { name: 'COUNTA(range)', desc: 'Counts the non-empty cells in a range.', example: '=COUNTA(A1:A4)' },
      { name: 'COUNTBLANK(range)', desc: 'Counts the empty cells in a range.', example: '=COUNTBLANK(A1:A4)' },
      { name: 'STDEV(range)', desc: 'Sample standard deviation of a range.', example: '=STDEV(A1:A4)' },
      { name: 'VAR(range)', desc: 'Sample variance of a range.', example: '=VAR(A1:A4)' }
    ]
  },
  {
    category: 'Math',
    functions:
    [
      { name: 'ABS(value)', desc: 'Absolute value of a number, cell, or expression.', example: '=ABS(A1-A2)' },
      { name: 'ROUND(value, decimals)', desc: 'Rounds to a number of decimal places.', example: '=ROUND(A1, 2)' },
      { name: 'ROUNDUP(value, [digits])', desc: 'Rounds away from zero.', example: '=ROUNDUP(A1, 1)' },
      { name: 'ROUNDDOWN(value, [digits])', desc: 'Rounds toward zero.', example: '=ROUNDDOWN(A1, 1)' },
      { name: 'TRUNC(value, [decimals])', desc: 'Truncates toward zero without rounding.', example: '=TRUNC(A1, 1)' },
      { name: 'INT(value)', desc: 'Rounds down to the nearest integer.', example: '=INT(A1)' },
      { name: 'CEILING(value, [significance])', desc: 'Rounds up to the nearest multiple.', example: '=CEILING(A1, 5)' },
      { name: 'FLOOR(value, [significance])', desc: 'Rounds down to the nearest multiple.', example: '=FLOOR(A1, 5)' },
      { name: 'SIGN(value)', desc: 'Returns -1, 0, or 1 depending on the sign.', example: '=SIGN(A1)' },
      { name: 'SQRT(value)', desc: 'Square root of a number, cell, or expression.', example: '=SQRT(A1)' },
      { name: 'POWER(base, exponent)', desc: 'Raises a number to a power (the ^ operator works too).', example: '=POWER(A1, 2)' },
      { name: 'MOD(number, divisor)', desc: 'Remainder after division.', example: '=MOD(A1, 3)' },
      { name: 'EXP(value)', desc: 'e raised to the given power.', example: '=EXP(A1)' },
      { name: 'LN(value)', desc: 'Natural logarithm.', example: '=LN(A1)' },
      { name: 'LOG(value, [base])', desc: 'Logarithm with an optional base (default 10).', example: '=LOG(A1, 2)' },
      { name: 'LOG10(value)', desc: 'Base-10 logarithm.', example: '=LOG10(A1)' },
      { name: 'PI()', desc: 'The value of pi.', example: '=PI()' },
      { name: 'RAND()', desc: 'Random number between 0 and 1.', example: '=RAND()' },
      { name: 'RANDBETWEEN(min, max)', desc: 'Random integer between two values.', example: '=RANDBETWEEN(1, 6)' },
      { name: 'FACT(n)', desc: 'Factorial of a number.', example: '=FACT(5)' },
      { name: 'EVEN(value)', desc: 'Rounds up (away from zero) to the nearest even integer.', example: '=EVEN(A1)' },
      { name: 'ODD(value)', desc: 'Rounds up (away from zero) to the nearest odd integer.', example: '=ODD(A1)' }
    ]
  },
  {
    category: 'Trigonometry',
    functions:
    [
      { name: 'SIN(angle)', desc: 'Sine of an angle in radians.', example: '=SIN(PI()/2)' },
      { name: 'COS(angle)', desc: 'Cosine of an angle in radians.', example: '=COS(A1)' },
      { name: 'TAN(angle)', desc: 'Tangent of an angle in radians.', example: '=TAN(A1)' },
      { name: 'ASIN(value)', desc: 'Inverse sine, in radians.', example: '=ASIN(A1)' },
      { name: 'ACOS(value)', desc: 'Inverse cosine, in radians.', example: '=ACOS(A1)' },
      { name: 'ATAN(value)', desc: 'Inverse tangent, in radians.', example: '=ATAN(A1)' },
      { name: 'ATAN2(x, y)', desc: 'Angle of the point (x, y), in radians.', example: '=ATAN2(1, 1)' },
      { name: 'SINH(value)', desc: 'Hyperbolic sine.', example: '=SINH(A1)' },
      { name: 'COSH(value)', desc: 'Hyperbolic cosine.', example: '=COSH(A1)' },
      { name: 'TANH(value)', desc: 'Hyperbolic tangent.', example: '=TANH(A1)' },
      { name: 'DEGREES(radians)', desc: 'Converts radians to degrees.', example: '=DEGREES(PI())' },
      { name: 'RADIANS(degrees)', desc: 'Converts degrees to radians.', example: '=RADIANS(180)' }
    ]
  },
  {
    category: 'Lookup',
    functions:
    [
      { name: 'INDEX(range, row, [col])', desc: 'Value at a position inside a range (1-based).', example: '=INDEX(A1:B10, 3, 2)' },
      { name: 'VLOOKUP(value, range, col)', desc: 'Finds value in the range\'s first column, returns the col-th column of that row.', example: '=VLOOKUP("Bob", A1:C9, 3)' },
      { name: 'HLOOKUP(value, range, row)', desc: 'Finds value in the range\'s first row, returns the row-th row of that column.', example: '=HLOOKUP("Q2", A1:F3, 2)' },
      { name: 'MATCH(value, range)', desc: 'Position of the first exact match in a row or column range (1-based).', example: '=MATCH(42, A1:A9)' },
      { name: 'ROW([ref])', desc: 'Row number of a reference, or of the formula\'s own cell.', example: '=ROW(B7)' },
      { name: 'COLUMN([ref])', desc: 'Column number of a reference, or of the formula\'s own cell.', example: '=COLUMN(B7)' }
    ]
  },
  {
    category: 'Text',
    functions:
    [
      { name: 'CONCAT(value1, value2, ...)', desc: 'Joins values into one piece of text.', example: '=CONCAT(A1, " ", A2)' },
      { name: 'CONCATENATE(value1, value2, ...)', desc: 'Same as CONCAT.', example: '=CONCATENATE(A1, A2)' },
      { name: 'LEN(text)', desc: 'Number of characters in the text.', example: '=LEN(A1)' },
      { name: 'UPPER(text)', desc: 'Converts text to uppercase.', example: '=UPPER(A1)' },
      { name: 'LOWER(text)', desc: 'Converts text to lowercase.', example: '=LOWER(A1)' },
      { name: 'TRIM(text)', desc: 'Removes leading/trailing whitespace.', example: '=TRIM(A1)' },
      { name: 'LEFT(text, [count])', desc: 'First N characters of the text.', example: '=LEFT(A1, 3)' },
      { name: 'RIGHT(text, [count])', desc: 'Last N characters of the text.', example: '=RIGHT(A1, 3)' },
      { name: 'MID(text, start, count)', desc: 'A substring starting at a given position.', example: '=MID(A1, 2, 3)' },
      { name: 'SUBSTITUTE(text, old, new, [n])', desc: 'Replaces occurrences of old with new (all, or only the n-th).', example: '=SUBSTITUTE(A1, "cat", "dog")' },
      { name: 'REPT(text, times)', desc: 'Repeats text a number of times.', example: '=REPT("-", 10)' },
      { name: 'FIND(needle, text, [start])', desc: 'Position of needle in text (case-sensitive, 1-based).', example: '=FIND("a", A1)' },
      { name: 'SEARCH(needle, text, [start])', desc: 'Like FIND, but case-insensitive.', example: '=SEARCH("a", A1)' },
      { name: 'PROPER(text)', desc: 'Capitalizes the first letter of each word.', example: '=PROPER(A1)' },
      { name: 'EXACT(text1, text2)', desc: 'TRUE if two texts are identical (case-sensitive).', example: '=EXACT(A1, B1)' }
    ]
  },
  {
    category: 'Logical',
    functions:
    [
      { name: 'IF(condition, true_value, false_value)', desc: 'Branches on a condition.', example: '=IF(A1>5,"big","small")' },
      { name: 'AND(cond1, cond2, ...)', desc: 'True if every condition is true.', example: '=AND(A1>0, A2>0)' },
      { name: 'OR(cond1, cond2, ...)', desc: 'True if any condition is true.', example: '=OR(A1>0, A2>0)' },
      { name: 'NOT(condition)', desc: 'Reverses a true/false value.', example: '=NOT(A1>5)' },
      { name: 'XOR(cond1, cond2, ...)', desc: 'True if an odd number of conditions are true.', example: '=XOR(A1>0, A2>0)' },
      { name: 'IFERROR(value, fallback)', desc: 'Returns a fallback if value errors out.', example: '=IFERROR(A1/A2, 0)' }
    ]
  }
];

// Flattened, autocomplete-friendly view of SHEET_FUNCTION_GROUPS - one entry
// per function with its bare name (the part before "(") split out, used by
// the in-cell function suggestion dropdown.
var ALL_SHEET_FUNCTIONS = SHEET_FUNCTION_GROUPS.reduce
(
  function(acc, group)
  {
    return acc.concat(group.functions);
  },
  []
).map
(
  function(fn)
  {
    return { bareName: fn.name.split('(')[0], signature: fn.name, desc: fn.desc };
  }
);

function renderFunctionHelp()
{
  var panel = document.getElementById('function-help-panel');

  if (!panel)
    return;

  var intro = '<div class="fx-help-intro">'+
                'Start a cell with <code>=</code> to write a formula. Reference cells like '+
                '<code>A1</code> or ranges like <code>A1:B3</code>. Whole columns and rows '+
                'work too: <code>A:A</code> is everything in column A, <code>4:4</code> '+
                'everything in row 4. Prefix ranges with <code>@</code> for line-by-line math: '+
                'inside an aggregate it combines the ranges element by element and returns one '+
                'result — <code>=SUM(@A1:A9*@B1:B9)</code> sums each row\'s product, like '+
                'SUMPRODUCT. Outside a function, <code>@</code> takes the value on the '+
                'formula\'s own line, so <code>=@A1:A9*@B1:B9</code> filled down computes '+
                'row by row. While editing a formula, '+
                'click or drag cells on the grid to insert their reference at the cursor. '+
                'Functions can be nested, e.g. <code>=ROUND(SUM(A1:A4),1)</code>. Prefix a '+
                'column or row with <code>$</code> (e.g. <code>$A$1</code>, <code>A$1</code>, '+
                '<code>$A1</code>) to lock it in place when filling.'+
              '</div>';

  var groups = SHEET_FUNCTION_GROUPS.map
  (
    function(group)
    {
      var rows = group.functions.map
      (
        function(fn)
        {
          return  '<div class="fx-help-row">'+
                    '<div class="fx-help-name">' + escHtml(fn.name) + '</div>'+
                    '<div class="fx-help-desc">' + escHtml(fn.desc) + ' <span class="fx-help-example">' + escHtml(fn.example) + '</span></div>'+
                  '</div>';
        }
      ).join('');

      return '<div class="fx-help-category">' + escHtml(group.category) + '</div>' + rows;
    }
  ).join('');

  panel.innerHTML = intro + groups;
}

function toggleFunctionHelp(clickEvent)
{
  if (clickEvent)
    clickEvent.stopPropagation();

  var panel = document.getElementById('function-help-panel');

  if (!panel)
    return;

  if (!panel.classList.contains('open'))
    renderFunctionHelp();

  panel.classList.toggle('open');
}

document.addEventListener
(
  'click',
  function(clickEvent)
  {
    var panel = document.getElementById('function-help-panel');

    if (!panel || !panel.classList.contains('open'))
      return;

    if (panel.contains(clickEvent.target) || clickEvent.target.id === 'fx-help-btn')
      return;

    panel.classList.remove('open');
  }
);

// ── MODAL ──
function openNewModal()
{
  newFileType = currentAppType;
  newDocTemplate = 'blank';

  document.getElementById('new-modal').classList.add('open');
  document.getElementById('new-name').value = '';

  var tplRow = document.getElementById('doc-template-row');
  if (tplRow) tplRow.style.display = (newFileType === 'doc') ? '' : 'none';

  document.querySelectorAll('.doc-tpl-card').forEach(function(c)
  {
    c.classList.toggle('selected', c.dataset.tpl === 'blank');
  });

  // Re-read the template store every time: another device may have synced
  // new templates since the last look.
  customTemplates = null;
  loadCustomTemplates().then(renderDocTemplateRow);

  setTimeout
  (
    function()
    {
      document.getElementById('new-name').focus();
    },
    50
  );

  document.querySelectorAll('.type-card').forEach
  (
    function(card)
    {
      card.classList.toggle('selected', card.dataset.type === newFileType);
    }
  );
}

function closeNewModal()
{
  document.getElementById('new-modal').classList.remove('open');
}

function selectNewType(type, el)
{
  newFileType = type;

  document.querySelectorAll('.type-card').forEach
  (
    function(c)
    {
      c.classList.remove('selected');
    }
  );

  el.classList.add('selected');

  var tplRow = document.getElementById('doc-template-row');
  if (tplRow) tplRow.style.display = (type === 'doc') ? '' : 'none';
}

function selectDocTemplate(tpl, el)
{
  newDocTemplate = tpl;
  document.querySelectorAll('.doc-tpl-card').forEach(function(c) { c.classList.remove('selected'); });
  el.classList.add('selected');
}

// ── CUSTOM TEMPLATES ──
// User-made document templates, alongside the built-in DOC_TEMPLATES. Stored
// as a `_lktpl.json` sidecar in the work folder (same pattern as the
// `_lkbl.json` backlinks index, so it works on web and desktop alike and
// never shows up in the file tree), or localStorage when no folder is open.

var customTemplates = null; // { name: content }, lazily loaded
var customTplKeys = [];     // sorted names; template ids are 'custom:<index>'

async function loadCustomTemplates()
{
  if (customTemplates !== null)
    return customTemplates;

  if (workFolderRoot)
  {
    try { customTemplates = JSON.parse(await Platform.readWorkFile(workFolderRoot, '_lktpl.json')) || {}; }
    catch(e) { customTemplates = {}; }
  }
  else
  {
    try { customTemplates = JSON.parse(localStorage.getItem('lk_custom_templates')) || {}; }
    catch(e) { customTemplates = {}; }
  }
  return customTemplates;
}

async function persistCustomTemplates()
{
  if (workFolderRoot)
  {
    try { await Platform.writeWorkFile(workFolderRoot, '_lktpl.json', JSON.stringify(customTemplates, null, 2)); }
    catch(e) { console.warn('Template store write error', e); }
  }
  else
  {
    try { localStorage.setItem('lk_custom_templates', JSON.stringify(customTemplates)); }
    catch(e) { console.warn('Template store write error', e); }
  }
}

function renderDocTemplateRow()
{
  var grid = document.querySelector('#doc-template-row .doc-tpl-grid');
  if (!grid) return;

  grid.querySelectorAll('.doc-tpl-custom').forEach(function(c){ c.remove(); });

  customTplKeys = Object.keys(customTemplates || {}).sort(function(a, b){ return a.localeCompare(b); });
  customTplKeys.forEach(function(name, i)
  {
    var card = document.createElement('div');
    card.className = 'doc-tpl-card doc-tpl-custom';
    card.dataset.tpl = 'custom:' + i;
    card.title = name;
    card.appendChild(document.createTextNode(name));
    card.onclick = function(){ selectDocTemplate('custom:' + i, card); };

    var del = document.createElement('button');
    del.className = 'doc-tpl-del';
    del.title = 'Delete this template';
    del.textContent = '×';
    del.onclick = function(ev){ deleteCustomTemplate(ev, name); };
    card.appendChild(del);

    grid.appendChild(card);
  });
}

async function deleteCustomTemplate(ev, name)
{
  ev.stopPropagation();
  delete customTemplates[name];
  await persistCustomTemplates();

  // If the deleted template was selected, fall back to Blank.
  if (newDocTemplate.indexOf('custom:') === 0 && customTplKeys[parseInt(newDocTemplate.slice(7), 10)] === name)
  {
    newDocTemplate = 'blank';
    document.querySelectorAll('.doc-tpl-card').forEach(function(c)
    {
      c.classList.toggle('selected', c.dataset.tpl === 'blank');
    });
  }
  renderDocTemplateRow();
}

function customTemplateContent(tplId)
{
  var name = customTplKeys[parseInt(tplId.slice(7), 10)];
  return name !== undefined ? (customTemplates || {})[name] : undefined;
}

function handleFileMenuSaveTemplate()
{
  closeFileMenu();

  var f = currentFileId && files[currentFileId];
  if (!f || f.type !== 'doc')
  {
    alert('Open a Document first — templates are created from the currently open Document.');
    return;
  }

  openDataModal
  (
    'Save as template',
    '<label class="field-label">Template name<input class="modal-input" id="tpl-name" value="' + escAttr(f.name || 'My template') + '"></label>' +
    '<div style="color:var(--text3);font-size:12px;margin-top:8px">The current contents of &ldquo;' + escHtml(f.name) + '&rdquo; will appear as a template choice in the New File dialog. Saving under an existing name replaces that template.</div>',
    async function()
    {
      var name = document.getElementById('tpl-name').value.trim();
      if (!name) return;
      await loadCustomTemplates();
      customTemplates[name] = f.content || '';
      await persistCustomTemplates();
      closeDataModal();
    }
  );
}

async function createNewFile()
{
  var name = document.getElementById('new-name').value.trim() || 'Untitled';
  var tplContent;
  if (newFileType === 'doc' && newDocTemplate !== 'blank')
  {
    tplContent = newDocTemplate.indexOf('custom:') === 0
                 ? customTemplateContent(newDocTemplate)
                 : (DOC_TEMPLATES[newDocTemplate] || undefined);
  }

  closeNewModal();

  if (workFolderRoot)
  {
    var workId = await createWorkFile(name, newFileType, tplContent);

    if (!workId)
      return;

    switchAppType(newFileType);
    await openFile(workId);
    return;
  }

  var id = createFile(name, newFileType, tplContent);

  switchAppType(newFileType);
  openFile(id);
}

document.getElementById('new-modal').addEventListener
(
  'click',
  function(e)
  {
    if (e.target === document.getElementById('new-modal'))
      closeNewModal();
  }
);

// ── EXPORT / IMPORT ──
async function exportCurrent()
{
  if (!currentFileId)
  {
    alert('No file open.');
    return;
  }

  var f = files[currentFileId],
      ext = fileExtensionFor(f.type);

  var fileName = f.name.replace(/\s+/g,'_') + '.' + ext;

  try
  {
    var saved = await Platform.saveFile(fileName, f.content);

    if (!saved)
      return;
  }
  catch(e)
  {
    console.warn('Save error', e);
    alert('Could not save file.');
  }
}

function toggleFileMenu(clickEvent)
{
  clickEvent.stopPropagation();

  var menu = document.getElementById('file-menu');
  if (!menu) return;

  var opening = !menu.classList.contains('open');
  if (!opening)
    document.getElementById('export-as-submenu').classList.remove('open');

  menu.classList.toggle('open');

  if (opening)
  {
    var btn  = document.getElementById('file-menu-btn');
    var rect = btn.getBoundingClientRect();
    menu.style.left   = Math.max(4, rect.left) + 'px';
    menu.style.bottom = (window.innerHeight - rect.top + 6) + 'px';
    menu.style.top    = 'auto';
  }
}

function toggleExportAsSubmenu(clickEvent)
{
  clickEvent.stopPropagation();

  var submenu = document.getElementById('export-as-submenu');

  if (submenu)
    submenu.classList.toggle('open');
}

function closeFileMenu()
{
  var menu = document.getElementById('file-menu'),
      submenu = document.getElementById('export-as-submenu');

  if (menu)
    menu.classList.remove('open');

  if (submenu)
    submenu.classList.remove('open');
}

document.addEventListener
(
  'click',
  function(clickEvent)
  {
    var menu = document.getElementById('file-menu');

    if (!menu || !menu.classList.contains('open'))
      return;

    if (menu.contains(clickEvent.target) || clickEvent.target.closest('#file-menu-btn'))
      return;

    closeFileMenu();
  }
);

function handleFileMenuExport()
{
  closeFileMenu();
  exportCurrent();
}

function handleFileMenuExportAs(format)
{
  closeFileMenu();
  exportCurrentAs(format);
}

function handleFileMenuImport()
{
  closeFileMenu();
  triggerImport();
}

function handleFileMenuSettings()
{
  closeFileMenu();
  openSettingsModal();
}

// ── ENABLED TYPES ──

var enabledTypes = null;
var ALL_APP_TYPES = ['doc','sheet','graph','notebook','glossary','calendar','economy','bestiary'];

function getEnabledTypes()
{
  if (!enabledTypes)
  {
    try
    {
      var saved = localStorage.getItem('lk_enabled_types');
      enabledTypes = saved ? JSON.parse(saved) : null;
    }
    catch(e) {}
    if (!enabledTypes) enabledTypes = ALL_APP_TYPES.slice();
  }
  return enabledTypes;
}

function saveEnabledTypes(types)
{
  enabledTypes = types.length ? types : [currentAppType || 'doc'];
  try { localStorage.setItem('lk_enabled_types', JSON.stringify(enabledTypes)); } catch(e) {}
  applyEnabledTypes();
}

function applyEnabledTypes()
{
  var enabled = getEnabledTypes();
  ALL_APP_TYPES.forEach(function(t)
  {
    var tab = document.getElementById('tab-' + t);
    if (tab) tab.style.display = enabled.includes(t) ? '' : 'none';
    var card = document.querySelector('.type-card[data-type="' + t + '"]');
    if (card) card.style.display = enabled.includes(t) ? '' : 'none';
  });
  if (enabled.length && !enabled.includes(currentAppType))
    switchAppType(enabled[0]);
}

function onTypeToggle()
{
  var enabled = [];
  document.querySelectorAll('#settings-type-toggles input[type=checkbox]:checked').forEach(function(cb){ enabled.push(cb.value); });
  if (!enabled.length)
  {
    var cur = document.querySelector('#settings-type-toggles input[value="' + (currentAppType||'doc') + '"]');
    if (cur) { cur.checked = true; enabled.push(cur.value); }
    else enabled = [currentAppType || 'doc'];
  }
  saveEnabledTypes(enabled);
}

// ── THEME ──

var THEME_VAR_NAMES = ['--bg','--bg2','--bg3','--bg4','--border','--border2','--text','--text2','--text3','--accent','--accent2','--accent-dim','--accent-dim2'];

var THEME_PRESETS =
[
  { name: 'Amber Dusk', vars: { '--bg':'#0f0e0c', '--bg2':'#1a1916', '--bg3':'#232220', '--bg4':'#2d2c29', '--border':'rgba(255,255,255,0.07)', '--border2':'rgba(255,255,255,0.13)', '--text':'#f0ede6', '--text2':'#9e9b94', '--text3':'#5c5a55', '--accent':'#d4a843', '--accent2':'#b8912e', '--accent-dim':'rgba(212,168,67,0.12)', '--accent-dim2':'rgba(212,168,67,0.06)' } },
  { name: 'Ocean Blue', vars: { '--bg':'#0b1218', '--bg2':'#151f27', '--bg3':'#1e2a33', '--bg4':'#28353f', '--border':'rgba(255,255,255,0.07)', '--border2':'rgba(255,255,255,0.13)', '--text':'#e8edf0', '--text2':'#94a3ac', '--text3':'#54626b', '--accent':'#4a9fd8', '--accent2':'#3a84b8', '--accent-dim':'rgba(74,159,216,0.12)', '--accent-dim2':'rgba(74,159,216,0.06)' } },
  { name: 'Forest Moss', vars: { '--bg':'#0d1410', '--bg2':'#161f19', '--bg3':'#1f2b23', '--bg4':'#29362d', '--border':'rgba(255,255,255,0.07)', '--border2':'rgba(255,255,255,0.13)', '--text':'#e8ede9', '--text2':'#93a398', '--text3':'#546155', '--accent':'#6fae5c', '--accent2':'#5a9048', '--accent-dim':'rgba(111,174,92,0.12)', '--accent-dim2':'rgba(111,174,92,0.06)' } },
  { name: 'Crimson Ember', vars: { '--bg':'#150e0d', '--bg2':'#201715', '--bg3':'#2b201d', '--bg4':'#362a26', '--border':'rgba(255,255,255,0.07)', '--border2':'rgba(255,255,255,0.13)', '--text':'#f0e8e6', '--text2':'#a89490', '--text3':'#63524e', '--accent':'#d9614a', '--accent2':'#bf4c37', '--accent-dim':'rgba(217,97,74,0.12)', '--accent-dim2':'rgba(217,97,74,0.06)' } },
  { name: 'Slate Violet', vars: { '--bg':'#100e17', '--bg2':'#1a1722', '--bg3':'#23202d', '--bg4':'#2d2a38', '--border':'rgba(255,255,255,0.07)', '--border2':'rgba(255,255,255,0.13)', '--text':'#eae7f0', '--text2':'#a19cac', '--text3':'#5f5a6b', '--accent':'#9a7fd4', '--accent2':'#8267bf', '--accent-dim':'rgba(154,127,212,0.12)', '--accent-dim2':'rgba(154,127,212,0.06)' } }
];

function hexToRgb(hex)
{
  hex = hex.replace('#', '');
  if (hex.length === 3) hex = hex.split('').map(function(c){ return c + c; }).join('');
  var n = parseInt(hex, 16);
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 };
}

function rgbToHex(r, g, b)
{
  return '#' + [r, g, b].map(function(v)
  {
    return Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0');
  }).join('');
}

function mixHex(hex, towardHex, amount)
{
  var a = hexToRgb(hex), b = hexToRgb(towardHex);
  return rgbToHex(a.r + (b.r - a.r) * amount, a.g + (b.g - a.g) * amount, a.b + (b.b - a.b) * amount);
}

function relativeLuminance(hex)
{
  var c = hexToRgb(hex);
  return (0.299 * c.r + 0.587 * c.g + 0.114 * c.b) / 255;
}

// Derives a full palette from just a background and an accent color, so a
// custom pick stays legible regardless of how light or dark it is.
function generateThemeVars(bgHex, accentHex)
{
  var isDark = relativeLuminance(bgHex) < 0.5,
      lightenTarget = isDark ? '#ffffff' : '#000000',
      text = isDark ? '#f0ede6' : '#1c1a17',
      accentRgb = hexToRgb(accentHex);

  return {
    '--bg': bgHex,
    '--bg2': mixHex(bgHex, lightenTarget, 0.045),
    '--bg3': mixHex(bgHex, lightenTarget, 0.085),
    '--bg4': mixHex(bgHex, lightenTarget, 0.125),
    '--border': isDark ? 'rgba(255,255,255,0.07)' : 'rgba(0,0,0,0.08)',
    '--border2': isDark ? 'rgba(255,255,255,0.13)' : 'rgba(0,0,0,0.15)',
    '--text': text,
    '--text2': mixHex(bgHex, text, 0.55),
    '--text3': mixHex(bgHex, text, 0.32),
    '--accent': accentHex,
    '--accent2': mixHex(accentHex, '#000000', 0.15),
    '--accent-dim': 'rgba(' + accentRgb.r + ',' + accentRgb.g + ',' + accentRgb.b + ',0.12)',
    '--accent-dim2': 'rgba(' + accentRgb.r + ',' + accentRgb.g + ',' + accentRgb.b + ',0.06)'
  };
}

function applyThemeVars(vars)
{
  var root = document.documentElement.style;
  THEME_VAR_NAMES.forEach(function(name)
  {
    if (vars[name]) root.setProperty(name, vars[name]);
  });
}

function resolveThemeChoice(choice)
{
  if (choice.type === 'custom') return generateThemeVars(choice.bg, choice.accent);
  var preset = THEME_PRESETS.find(function(p){ return p.name === choice.name; });
  return preset ? preset.vars : null;
}

function loadThemeChoice()
{
  try
  {
    var raw = localStorage.getItem('lk_theme');
    return raw ? JSON.parse(raw) : null;
  }
  catch(e) { return null; }
}

function saveThemeChoice(choice)
{
  try { localStorage.setItem('lk_theme', JSON.stringify(choice)); } catch(e) {}
  persistRoamingTheme(choice);
}

function initTheme()
{
  var choice = loadThemeChoice();
  if (choice)
  {
    var vars = resolveThemeChoice(choice);
    if (vars) applyThemeVars(vars);
  }
  // localStorage is only per-browser/per-install, so on a fresh device (or
  // after an app reinstall wiped webview storage) pull the roamed choice from
  // the workspace; repaints shortly after first paint when they differ.
  loadRoamingTheme();
}

// ── ROAMING APPEARANCE ──
// The theme choice also lives in a `_lkprefs.json` workspace sidecar (synced
// across devices like the `_lktpl.json` template store), with localStorage
// kept as the instant-paint cache.

async function loadRoamingTheme()
{
  if (!workFolderRoot) return;
  var prefs;
  try { prefs = JSON.parse(await Platform.readWorkFile(workFolderRoot, '_lkprefs.json')); }
  catch(e) { return; } // no prefs stored yet
  if (!prefs || !('theme' in prefs)) return;

  var localRaw  = localStorage.getItem('lk_theme'),
      roamedRaw = prefs.theme ? JSON.stringify(prefs.theme) : null;
  if (roamedRaw === localRaw) return;

  try
  {
    if (roamedRaw) localStorage.setItem('lk_theme', roamedRaw);
    else localStorage.removeItem('lk_theme');
  }
  catch(e) {}

  if (prefs.theme)
  {
    var vars = resolveThemeChoice(prefs.theme);
    if (vars) applyThemeVars(vars);
  }
  else
    THEME_VAR_NAMES.forEach(function(name){ document.documentElement.style.removeProperty(name); });
}

/// choice = null records an explicit reset, so it roams too.
async function persistRoamingTheme(choice)
{
  if (!workFolderRoot) return;
  var prefs = {};
  try { prefs = JSON.parse(await Platform.readWorkFile(workFolderRoot, '_lkprefs.json')) || {}; }
  catch(e) {}
  prefs.theme = choice;
  try { await Platform.writeWorkFile(workFolderRoot, '_lkprefs.json', JSON.stringify(prefs, null, 2)); }
  catch(e) { console.warn('Prefs write error', e); }
}

function selectThemePreset(name)
{
  var choice = { type: 'preset', name: name };
  saveThemeChoice(choice);
  applyThemeVars(resolveThemeChoice(choice));
  renderThemeSection();
}

function onCustomColorChange()
{
  var bg = document.getElementById('settings-bg-color').value,
      accent = document.getElementById('settings-accent-color').value,
      choice = { type: 'custom', bg: bg, accent: accent };
  saveThemeChoice(choice);
  applyThemeVars(generateThemeVars(bg, accent));
  renderThemeSection();
}

function resetTheme()
{
  localStorage.removeItem('lk_theme');
  THEME_VAR_NAMES.forEach(function(name){ document.documentElement.style.removeProperty(name); });
  persistRoamingTheme(null); // roam the reset too
  renderThemeSection();
}

function renderThemeSection()
{
  var choice = loadThemeChoice() || { type: 'preset', name: 'Amber Dusk' };

  document.getElementById('settings-theme-presets').innerHTML = THEME_PRESETS.map(function(p)
  {
    var active = (choice.type === 'preset' && choice.name === p.name) ? ' active' : '';
    return '<button type="button" class="settings-theme-swatch' + active + '" title="' + p.name + '" onclick="selectThemePreset(\'' + p.name + '\')">' +
      '<span class="settings-theme-swatch-dot" style="background:' + p.vars['--bg4'] + '"></span>' +
      '<span class="settings-theme-swatch-dot" style="background:' + p.vars['--accent'] + '"></span>' +
      '<span class="settings-theme-swatch-label">' + p.name + '</span>' +
    '</button>';
  }).join('');

  var defaultPreset = THEME_PRESETS[0];
  document.getElementById('settings-bg-color').value = choice.type === 'custom' ? choice.bg : defaultPreset.vars['--bg'];
  document.getElementById('settings-accent-color').value = choice.type === 'custom' ? choice.accent : defaultPreset.vars['--accent'];
}

// ── SETTINGS MODAL ──

function openSettingsModal()
{
  renderSettingsModal();
  document.getElementById('settings-modal').classList.add('open');
  refreshCloudStatus();

  if (!Platform.isNative)
    refreshCustomFonts().then(renderCustomFontsSection);
}

function closeSettingsModal()
{
  document.getElementById('settings-modal').classList.remove('open');
}

var TYPE_LABELS = { doc:'Documents', sheet:'Sheets', graph:'Diagrams', notebook:'Notebook / Maps', glossary:'Glossary', calendar:'Calendar', economy:'Economy', bestiary:'Bestiary' };

function renderSettingsModal()
{
  document.getElementById('settings-folder-path').textContent =
    workFolderRoot || 'This browser (no folder set)';

  document.getElementById('settings-stop-btn').style.display = workFolderRoot ? '' : 'none';

  var enabled = getEnabledTypes();
  document.getElementById('settings-type-toggles').innerHTML = ALL_APP_TYPES.map(function(t){
    var checked = enabled.includes(t) ? ' checked' : '';
    return '<label class="settings-type-toggle">' +
      '<input type="checkbox" value="' + t + '"' + checked + ' onchange="onTypeToggle()"> ' +
      '<span>' + (TYPE_LABELS[t] || t) + '</span>' +
    '</label>';
  }).join('');

  renderThemeSection();
  renderCloudSyncSection();
  renderCustomFontsSection();
}

// ── CUSTOM FONTS ──
// Account-level fonts (server: src/fonts.rs). Unlike the picker's
// queryLocalFonts() path (Chromium + secure-context only, and dependent on
// whatever happens to be installed on THIS device), an uploaded font is
// served as a real asset and always renders the same way for every viewer,
// in the live editor and in exported PDFs (see embed_account_fonts_style
// server-side and buildDocExportHtml's use of injectedFontStyleTag below).

var customFontsCache = null; // [{id, familyName, contentType, byteSize}]

async function refreshCustomFonts()
{
  customFontsCache = await Platform.listCustomFonts();
  return customFontsCache;
}

async function injectCustomFontsStyle()
{
  if (!customFontsCache) await refreshCustomFonts();

  var tag = document.getElementById('custom-fonts-style');
  if (!tag)
  {
    tag = document.createElement('style');
    tag.id = 'custom-fonts-style';
    document.head.appendChild(tag);
  }

  var rules = await Promise.all((customFontsCache || []).map(async function(f)
  {
    var url = await Platform.customFontDataUrl(f.id);
    var format = f.contentType.indexOf('woff2') !== -1 ? 'woff2'
               : f.contentType.indexOf('woff') !== -1  ? 'woff'
               : f.contentType.indexOf('otf') !== -1   ? 'opentype'
               : 'truetype';
    return "@font-face{font-family:'" + f.familyName.replace(/'/g, '') + "';src:url('" + url + "') format('" + format + "');}";
  }));

  tag.textContent = rules.join('\n');
}

function renderCustomFontsSection()
{
  var section = document.getElementById('custom-fonts-section');
  if (!section) return;

  if (Platform.isNative)
  {
    section.style.display = 'none';
    return;
  }

  section.style.display = '';
  document.getElementById('custom-font-error').style.display = 'none';

  var list = document.getElementById('custom-fonts-list');
  var fonts = customFontsCache || [];

  list.innerHTML = fonts.length
    ? fonts.map(function(f)
      {
        return '<div class="custom-font-row">' +
          '<span class="custom-font-name" style="font-family:' + escAttr(cssFontFamilyValue(f.familyName)) + '">' + escHtml(f.familyName) + '</span>' +
          '<span class="custom-font-size">' + Math.max(1, Math.round(f.byteSize / 1024)) + ' KB</span>' +
          '<button class="gls-card-del" style="position:static" onclick="removeCustomFont(\'' + escAttr(f.id) + '\')" title="Delete">×</button>' +
        '</div>';
      }).join('')
    : '<div class="settings-help-text" style="margin:0">No custom fonts yet.</div>';
}

async function onCustomFontFileChosen()
{
  var fileInput = document.getElementById('custom-font-file'),
      nameInput = document.getElementById('custom-font-name'),
      errEl = document.getElementById('custom-font-error'),
      file = fileInput.files && fileInput.files[0];

  if (!file) return;

  var familyName = nameInput.value.trim() || file.name.replace(/\.[^.]+$/, '');
  errEl.style.display = 'none';

  try
  {
    var bytes = await file.arrayBuffer();
    await Platform.uploadCustomFont(familyName, file.name, bytes);
    nameInput.value = '';
    fileInput.value = '';
    await refreshCustomFonts();
    renderCustomFontsSection();
    await injectCustomFontsStyle();
    cachedFontList = null; // let the "Font…" picker pick up the new family
  }
  catch(e)
  {
    errEl.textContent = e.message || 'Upload failed';
    errEl.style.display = '';
  }
}

async function removeCustomFont(id)
{
  try { await Platform.deleteCustomFont(id); } catch(e) {}
  await refreshCustomFonts();
  renderCustomFontsSection();
  await injectCustomFontsStyle();
  cachedFontList = null;
}

document.getElementById('settings-modal').addEventListener
(
  'click',
  function(e)
  {
    if (e.target === document.getElementById('settings-modal'))
      closeSettingsModal();
  }
);

// Writes every file currently in localStorage out as a real file at the root
// of `root`, de-duplicating names locally (the in-memory `files` are still
// keyed by the old synthetic ids at this point, so collisions can't be
// detected by checking `files[candidate]` the way uniqueRelPath does).
async function migrateLocalStorageFilesToFolder(root)
{
  const usedPaths = {},
        entries = Object.values(files);

  for (let i = 0; i < entries.length; i++)
  {
    const f = entries[i],
          ext = fileExtensionFor(f.type),
          base = sanitizeFileName(f.name) || 'Untitled';

    let candidate = base + '.' + ext,
        n = 2;

    while (usedPaths[candidate])
    {
      candidate = base + ' (' + n + ').' + ext;
      n++;
    }

    usedPaths[candidate] = true;

    try
    {
      await Platform.writeWorkFile(root, candidate, f.content || '');
    }
    catch(e)
    {
      console.warn('Migration write error for "' + f.name + '"', e);
    }
  }
}

async function chooseWorkFolder()
{
  let picked;

  try
  {
    picked = await Platform.pickWorkFolder();
  }
  catch(e)
  {
    console.warn('Folder picker error', e);
    return;
  }

  if (!picked)
    return;

  if (Object.keys(files).length > 0)
    await migrateLocalStorageFilesToFolder(picked);

  workFolderRoot = picked;
  saveSettings();

  currentFileId = null;
  expandedFolders = new Set();
  clearActiveEditors();

  await loadWorkFolderTree();
  renderSettingsModal();
}

function stopUsingWorkFolder()
{
  workFolderRoot = null;
  saveSettings();

  folders = {};
  currentFileId = null;
  expandedFolders = new Set();
  clearActiveEditors();

  loadFromStorage();
  renderFileList();
  renderSettingsModal();
}

// ── CLOUD SYNC (desktop only) ──
// Connects the active work folder to an officesuite-web server via the
// Tauri-only cloud* commands (platform.js -> src-tauri/src/cloud.rs). The web
// build already talks to the cloud natively through Platform's work-folder
// calls and has no separate "connect" concept, so this whole section stays
// hidden there (Platform.isNative is false).

var cloudStatusCache = { connected: false };
var syncConflicts = [];
var activeSyncConflict = null;

async function refreshCloudStatus()
{
  if (!Platform.isNative)
    return;

  try { cloudStatusCache = await Platform.cloudStatus(); }
  catch(e) { cloudStatusCache = { connected: false }; }

  renderCloudSyncSection();
}

function renderCloudSyncSection()
{
  var section = document.getElementById('cloud-sync-section');

  if (!section)
    return;

  if (!Platform.isNative)
  {
    section.style.display = 'none';
    return;
  }

  section.style.display = '';

  var needsFolderEl = document.getElementById('cloud-sync-needs-folder'),
      disconnectedEl = document.getElementById('cloud-sync-disconnected'),
      connectedEl = document.getElementById('cloud-sync-connected');

  if (!workFolderRoot)
  {
    needsFolderEl.style.display = '';
    disconnectedEl.style.display = 'none';
    connectedEl.style.display = 'none';
    return;
  }

  needsFolderEl.style.display = 'none';

  if (!cloudStatusCache.connected)
  {
    disconnectedEl.style.display = '';
    connectedEl.style.display = 'none';
    return;
  }

  disconnectedEl.style.display = 'none';
  connectedEl.style.display = '';

  var statusEl = document.getElementById('cloud-sync-status'),
      conflictsBtn = document.getElementById('cloud-conflicts-btn');

  var when = cloudStatusCache.lastSyncedAt ? new Date(cloudStatusCache.lastSyncedAt).toLocaleString() : 'not yet';
  statusEl.textContent = 'Syncing with ' + cloudStatusCache.serverUrl + ' as ' + cloudStatusCache.email + '. Last synced: ' + when + '.';

  if (cloudStatusCache.conflictCount > 0)
  {
    conflictsBtn.style.display = '';
    conflictsBtn.textContent = 'View conflicts (' + cloudStatusCache.conflictCount + ')';
  }
  else
  {
    conflictsBtn.style.display = 'none';
  }
}

async function connectCloud()
{
  var serverUrl = document.getElementById('cloud-server-url').value.trim(),
      email = document.getElementById('cloud-email').value.trim(),
      password = document.getElementById('cloud-password').value;

  if (!serverUrl || !email || !password)
  {
    alert('Enter a server URL, email, and password.');
    return;
  }

  if (!workFolderRoot)
  {
    alert('Choose a work folder first — that folder becomes the sync target.');
    return;
  }

  try
  {
    cloudStatusCache = await Platform.cloudConnect(serverUrl, email, password, workFolderRoot);
    document.getElementById('cloud-password').value = '';
    renderSettingsModal();
    await loadWorkFolderTree();
  }
  catch(e)
  {
    alert('Could not connect: ' + e.message);
  }
}

async function disconnectCloud()
{
  try { await Platform.cloudDisconnect(); }
  catch(e) { console.warn('Cloud disconnect error', e); }

  cloudStatusCache = { connected: false };
  renderCloudSyncSection();
}

async function syncNowCloud()
{
  try
  {
    cloudStatusCache = await Platform.cloudSyncNow();
    renderCloudSyncSection();
    await loadWorkFolderTree();
  }
  catch(e)
  {
    alert('Sync failed: ' + e.message);
  }
}

// ── Sync Conflicts modal ──
// Each conflict is a file both sides edited since the last sync in ways that
// couldn't be auto-merged (see three_way_merge in cloud.rs). The user picks
// which version wins, or keeps both as separate files.

async function openSyncConflictsModal()
{
  try { syncConflicts = await Platform.cloudListConflicts(); }
  catch(e) { syncConflicts = []; }

  renderSyncConflictsModal();
  document.getElementById('sync-conflicts-modal').classList.add('open');
}

function closeSyncConflictsModal()
{
  document.getElementById('sync-conflicts-modal').classList.remove('open');
}

function renderSyncConflictsModal()
{
  var list = document.getElementById('sync-conflicts-list');

  if (!syncConflicts.length)
  {
    list.innerHTML = '<div class="history-empty">No conflicts — everything is synced.</div>';
    return;
  }

  list.innerHTML = syncConflicts.map(function(c, i)
  {
    return '<div class="history-item">' +
             '<div class="history-item-info">' +
               '<div class="history-item-time">' + escHtml(c.relPath) + '</div>' +
               '<div class="history-item-preview">Both your device and the cloud changed this file since the last sync.</div>' +
             '</div>' +
             '<div class="history-item-actions">' +
               '<button type="button" class="btn-cancel history-compare-btn" onclick="openSyncConflictDiff(' + i + ')">Review</button>' +
             '</div>' +
           '</div>';
  }).join('');
}

function openSyncConflictDiff(index)
{
  activeSyncConflict = syncConflicts[index];

  if (!activeSyncConflict)
    return;

  document.getElementById('sync-conflict-diff-title').textContent = activeSyncConflict.relPath;
  renderWordDiff(activeSyncConflict.localContent, activeSyncConflict.remoteContent, 'sync-conflict-local-pane', 'sync-conflict-remote-pane');

  closeSyncConflictsModal();
  document.getElementById('sync-conflict-diff-modal').classList.add('open');
}

function closeSyncConflictDiffModal()
{
  document.getElementById('sync-conflict-diff-modal').classList.remove('open');
  document.getElementById('sync-conflicts-modal').classList.add('open');
}

async function resolveSyncConflict(choice)
{
  if (!activeSyncConflict)
    return;

  try
  {
    await Platform.cloudResolveConflict(activeSyncConflict.relPath, choice);
  }
  catch(e)
  {
    alert('Could not resolve conflict: ' + e.message);
    return;
  }

  document.getElementById('sync-conflict-diff-modal').classList.remove('open');
  activeSyncConflict = null;

  await refreshCloudStatus();
  await openSyncConflictsModal();
  await loadWorkFolderTree();
}

// Wraps rendered body markup (doc HTML or a sheet table) into a standalone,
// printable HTML document - intentionally light-themed regardless of the
// app's dark UI, since this is meant to be portable/readable on its own.
function wrapExportHtml(title, bodyHtml)
{
  return  '<!DOCTYPE html>\n'+
          '<html lang="en">\n'+
          '<head>\n'+
          '<meta charset="UTF-8">\n'+
          '<title>' + escHtml(title) + '</title>\n'+
          '<style>\n'+
          '@page{size:A4;margin:2cm;}\n'+
          'body{font-family:Georgia,\'Times New Roman\',serif;max-width:800px;margin:40px auto;padding:0 20px;color:#1a1a1a;line-height:1.6;}\n'+
          'h1,h2,h3,h4,h5,h6{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;}\n'+
          'table{border-collapse:collapse;width:100%;font-family:Consolas,monospace;font-size:13px;}\n'+
          'td{border:1px solid #ccc;padding:4px 8px;}\n'+
          'code{background:#f0f0f0;padding:2px 4px;border-radius:3px;}\n'+
          'pre{background:#f0f0f0;padding:12px;border-radius:6px;overflow-x:auto;}\n'+
          'blockquote{border-left:3px solid #ccc;margin-left:0;padding-left:16px;color:#555;}\n'+
          '.mermaid-diagram{text-align:center;margin:1.5em 0;}\n'+
          '.mermaid-diagram svg{max-width:100%;height:auto;}\n'+
          'table,tr,blockquote,pre,.mermaid-diagram,.notebook-page{break-inside:avoid;}\n'+
          'h1,h2,h3,h4,h5,h6{break-after:avoid;}\n'+
          '.notebook-page{break-after:page;}\n'+
          '.notebook-page:last-child{break-after:auto;}\n'+
          '@media print{\n'+
          'body{max-width:none;margin:0;padding:0;}\n'+
          '}\n'+
          '</style>\n'+
          '</head>\n'+
          '<body>\n' + bodyHtml + '\n</body>\n'+
          '</html>';
}

function buildDocExportHtml(file)
{
  var bodyHtml = escHtml(file.content || '');

  if (typeof marked !== 'undefined')
  {
    documentHeadings = [];
    slugCounts = {};
    headingRenderCursor = 0;

    var tokens = marked.lexer(file.content || '');
    collectHeadings(tokens);

    bodyHtml = marked.parser(tokens);
  }

  return wrapExportHtml(file.name, '<h1>' + escHtml(file.name) + '</h1>\n' + bodyHtml);
}

function buildSheetExportHtml(file)
{
  var sections = '',
      savedData = sheetData;

  sheetPages.forEach(function(page)
  {
    // The display/eval helpers read the global sheetData; point it at each
    // page in turn while its table is built.
    sheetData = page.data;

    var maxRow = 0,
        maxCol = 0;

    Object.keys(page.data).forEach
    (
      function(ref)
      {
        var parsed = parseName(ref);

        if (!parsed)
          return;

        maxRow = Math.max(maxRow, parsed.row);
        maxCol = Math.max(maxCol, colIndex(parsed.col));
      }
    );

    var rowsHtml = '';

    for (var r = 1; r <= maxRow; r++)
    {
      rowsHtml += '<tr>';

      for (var c = 0; c <= maxCol; c++)
        rowsHtml += '<td>' + escHtml(getDisplayValue(colName(c) + r)) + '</td>';

      rowsHtml += '</tr>\n';
    }

    if (sheetPages.length > 1)
      sections += '<h2>' + escHtml(page.name) + '</h2>\n';

    sections += maxRow
                ?
                  '<table>\n<tbody>\n' + rowsHtml + '</tbody>\n</table>\n'
                :
                  '<p>This tab is empty.</p>\n';
  });

  sheetData = savedData;

  return wrapExportHtml(file.name, '<h1>' + escHtml(file.name) + '</h1>\n' + sections);
}

// Mermaid's library isn't shipped inside exported files, so any `.mermaid`
// placeholder is rendered to a static <svg> here (via mermaid.render, which
// works off-DOM) and the markup is inlined directly - keeping the exported
// HTML/PDF fully self-contained and immune to script execution quirks in
// whatever renders it later (the server-side PDF renderer, or the Tauri
// print-preview iframe).
async function inlineMermaidSvgs(html)
{
  if (typeof mermaid === 'undefined')
    return html;

  var container = document.createElement('div');
  container.innerHTML = html;

  var nodes = Array.prototype.slice.call(container.querySelectorAll('.mermaid'));

  if (!nodes.length)
    return html;

  for (var i = 0; i < nodes.length; i++)
  {
    var node = nodes[i],
        source = node.textContent;

    try
    {
      var result = await mermaid.render('mermaid-export-' + Date.now() + '-' + i, source);
      node.outerHTML = '<div class="mermaid-diagram">' + result.svg + '</div>';
    }
    catch(e)
    {
      console.warn('Mermaid export render error', e);
      node.outerHTML = '<pre>' + escHtml(source) + '</pre>';
    }
  }

  return container.innerHTML;
}

async function buildGraphExportHtml(file)
{
  var bodyHtml = escHtml(file.content || '');

  if (typeof marked !== 'undefined')
    bodyHtml = marked.parser(marked.lexer(file.content || ''));

  bodyHtml = await inlineMermaidSvgs(bodyHtml);

  return wrapExportHtml(file.name, '<h1>' + escHtml(file.name) + '</h1>\n' + bodyHtml);
}

// Shows the HTML in a genuinely on-screen preview (rather than a hidden/
// off-screen iframe) before printing. An invisible or off-canvas iframe is
// the usual trick for "print this content", but some embedded webviews
// (e.g. Tauri's WebView2 on Windows) appear to rasterize whatever's at the
// iframe's on-page position for "Print to PDF" rather than the iframe's own
// internal layout - which is empty when the iframe sits off-screen, so the
// resulting PDF comes out blank. Keeping the content visibly on-screen
// sidesteps that entirely, and lets the user see what they're printing.
function printHtmlContent(html)
{
  var overlay = document.getElementById('print-preview-overlay'),
      iframe = document.getElementById('print-preview-frame');

  if (!overlay || !iframe)
    return;

  iframe.srcdoc = html;
  overlay.classList.add('open');
}

function printPreviewNow()
{
  var iframe = document.getElementById('print-preview-frame');

  if (!iframe || !iframe.contentWindow)
    return;

  iframe.contentWindow.focus();
  iframe.contentWindow.print();
}

function closePrintPreview()
{
  var overlay = document.getElementById('print-preview-overlay'),
      iframe = document.getElementById('print-preview-frame');

  if (overlay)
    overlay.classList.remove('open');

  if (iframe)
    iframe.srcdoc = '';
}

document.getElementById('print-preview-overlay').addEventListener
(
  'click',
  function(e)
  {
    if (e.target === document.getElementById('print-preview-overlay'))
      closePrintPreview();
  }
);

async function exportCurrentAs(format)
{
  if (!currentFileId)
  {
    alert('No file open.');
    return;
  }

  var f = files[currentFileId],
      html;

  if (f.type === 'sheet')
    html = buildSheetExportHtml(f);
  else if (f.type === 'graph')
    html = await buildGraphExportHtml(f);
  else if (f.type === 'notebook')
    html = await buildNotebookExportHtml(f);
  else
    html = buildDocExportHtml(f);

  if (format === 'pdf')
  {
    if (Platform.isNative)
    {
      printHtmlContent(html);
      return;
    }

    var pdfName = f.name.replace(/\s+/g,'_') + '.pdf';

    try
    {
      await Platform.exportPdf(pdfName, html);
    }
    catch(e)
    {
      console.warn('PDF export error', e);
      alert('Could not export PDF.');
    }

    return;
  }

  var fileName = f.name.replace(/\s+/g,'_') + '.html';

  try
  {
    var saved = await Platform.saveFile(fileName, html);

    if (!saved)
      return;
  }
  catch(e)
  {
    console.warn('Export error', e);
    alert('Could not export file.');
  }
}

function triggerImport()
{
  document.getElementById('import-input').click();
}

function handleImport(e)
{
  var file = e.target.files[0];
  if (!file)
    return;

  var reader = new FileReader();

  reader.onload = async function(ev)
  {
    var content = ev.target.result,
        type =  (file.name.endsWith('.mds') || file.name.endsWith('.csv')) ? 'sheet'
              : file.name.endsWith('.mdg') ? 'graph'
              : file.name.endsWith('.mdn') ? 'notebook'
              : file.name.endsWith('.mdl') ? 'glossary'
              : file.name.endsWith('.mdc') ? 'calendar'
              : file.name.endsWith('.mde') ? 'economy'
              : file.name.endsWith('.mdb') ? 'bestiary'
              : 'doc';

    var name = file.name.replace(/\.[^.]+$/,'');

    if (workFolderRoot)
    {
      var relPath = uniqueRelPath('', sanitizeFileName(name) || 'Untitled', fileExtensionFor(type));

      try
      {
        await Platform.writeWorkFile(workFolderRoot, relPath, content);
      }
      catch(err)
      {
        console.warn('Work folder import error', err);
        alert('Could not import that file.');
        return;
      }

      await loadWorkFolderTree();
      switchAppType(type);
      await openFile(relPath);
      return;
    }

    var id = createFile(name, type, content);
    switchAppType(type);
    openFile(id);
  };

  reader.readAsText(file);
  e.target.value = '';
}

function escHtml(s)
{
  return String(s)
          .replace(/&/g,'&amp;')
          .replace(/</g,'&lt;')
          .replace(/>/g,'&gt;')
          .replace(/"/g,'&quot;');
}

// Escapes a value for safe use inside a single-quoted JS string literal that
// itself sits inside a double-quoted HTML attribute (the onclick="fn('ID')"
// pattern used throughout the sidebar) - relevant since rel_paths/names can
// contain arbitrary text once a user renames something.
function escAttr(s)
{
  return String(s).replace(/\\/g,'\\\\').replace(/'/g,"\\'").replace(/"/g,'&quot;');
}


// ── SHARING & COMMENTS (web only) ──────────────────────────
// Files/folders can be shared with other accounts at view/comment/edit level
// (server: officesuite-web src/share.rs). Shared entries appear in their own
// sidebar section; opening one loads it into the normal editors under the
// reserved id SHARED_TMP_ID, so it never mixes with the user's own workspace
// files. Saves are routed through the share (edit permission) or dropped.

var SHARED_TMP_ID = '@shared';
var sharedCtx = null;              // { shareId, subPath, permission, ownerEmail } while a shared file is open
var sharedEntries = [];            // GET /api/shared result
var sharedChildren = {};           // shareId -> FsEntry[] (folder shares, lazily fetched)
var sharedExpanded = {};           // shareId -> bool
var sharedRegistry = {};           // rowKey -> openable file descriptor
var sharedCollapsed = false;

function typeForFileName(name)
{
  var n = (name || '').toLowerCase();
  return n.endsWith('.mds') ? 'sheet'
       : n.endsWith('.mdg') ? 'graph'
       : n.endsWith('.mdn') ? 'notebook'
       : n.endsWith('.mdl') ? 'glossary'
       : n.endsWith('.mdc') ? 'calendar'
       : n.endsWith('.mde') ? 'economy'
       : n.endsWith('.mdb') ? 'bestiary'
       : 'doc';
}

var SHARE_PERM_COLORS = { view: '#9ca3af', comment: '#f1c40f', edit: '#2ecc71' };
function sharePermBadge(perm)
{
  var c = SHARE_PERM_COLORS[perm] || '#9ca3af';
  return '<span class="share-perm-badge" style="background:' + c + '22;color:' + c + ';border-color:' + c + '55">' + escHtml(perm) + '</span>';
}

async function refreshSharedSection()
{
  if (Platform.isNative) return;
  var fetched;
  try { fetched = await Platform.listSharedWithMe(); }
  catch(e) { fetched = []; }
  // Link entries (from a #link= URL) aren't in the account's share list; keep
  // them across refreshes. Captured after the await so an entry added while
  // the request was in flight isn't lost.
  var linkEntries = sharedEntries.filter(function(s){ return s.link; });
  sharedEntries = fetched.concat(linkEntries);
  renderSharedSection();
}

function sharedEntryId(s) { return s.shareId || ('link:' + s.link); }

function toggleSharedSection()
{
  sharedCollapsed = !sharedCollapsed;
  renderSharedSection();
}

function sharedFileRowHtml(key, name, ownerLabel, permission, depth)
{
  var type = typeForFileName(name);
  return '<div class="file-item shared-row" style="padding-left:' + (10 + depth * 14) + 'px" onclick="openSharedRegistry(\'' + escAttr(key) + '\')">' +
    '<span class="file-icon">' + fileTypeIcon(type) + '</span>' +
    '<span class="file-name" title="' + escAttr(ownerLabel) + '">' + escHtml(name.replace(/\.[^.]+$/, '')) + '</span>' +
    sharePermBadge(permission) +
  '</div>';
}

function renderSharedSection()
{
  var header = document.getElementById('shared-section-header'),
      list   = document.getElementById('shared-list');
  if (!header || !list) return;

  if (!sharedEntries.length)
  {
    header.style.display = 'none';
    list.style.display = 'none';
    return;
  }

  header.style.display = '';
  document.getElementById('shared-section-chevron').textContent = sharedCollapsed ? '▸' : '▾';
  list.style.display = sharedCollapsed ? 'none' : '';
  if (sharedCollapsed) return;

  sharedRegistry = {};
  var html = '';

  sharedEntries.forEach(function(s)
  {
    var id = sharedEntryId(s);
    var ownerLabel = s.link ? 'Shared via link by ' + s.ownerEmail : 'Shared by ' + s.ownerEmail;

    if (!s.exists)
    {
      html += '<div class="file-item shared-row shared-gone" title="The owner removed or moved this">' +
        '<span class="file-name">' + escHtml(s.name) + '</span>' +
        '<span class="shared-owner">gone</span>' +
        (s.shareId ? '<button class="gls-card-del" style="position:static" onclick="leaveShare(event,\'' + escAttr(s.shareId) + '\')" title="Remove">×</button>' : '') +
      '</div>';
      return;
    }

    if (s.isDir)
    {
      var open = !!sharedExpanded[id];
      html += '<div class="file-item shared-row" onclick="toggleSharedFolder(\'' + escAttr(id) + '\')">' +
        '<span class="file-icon">' + FOLDER_ICON + '</span>' +
        '<span class="file-name" title="' + escAttr(ownerLabel) + '">' + (open ? '▾ ' : '▸ ') + escHtml(s.name) + '</span>' +
        sharePermBadge(s.permission) +
      '</div>';

      if (open && sharedChildren[id])
        html += sharedFolderChildrenHtml(s, sharedChildren[id], 1);
    }
    else
    {
      sharedRegistry[id] = { shareId: s.shareId, link: s.link, subPath: '', name: s.name, permission: s.permission, ownerEmail: s.ownerEmail };
      html += sharedFileRowHtml(id, s.name, ownerLabel, s.permission, 0);
    }
  });

  list.innerHTML = html;
}

function sharedFolderChildrenHtml(share, entries, depth)
{
  var html = '';
  var parentId = sharedEntryId(share);
  var ownerLabel = share.link ? 'Shared via link by ' + share.ownerEmail : 'Shared by ' + share.ownerEmail;
  entries.forEach(function(e)
  {
    if (e.isDir)
    {
      html += '<div class="file-item shared-row" style="padding-left:' + (10 + depth * 14) + 'px">' +
        '<span class="file-icon">' + FOLDER_ICON + '</span>' +
        '<span class="file-name">' + escHtml(e.name) + '</span>' +
      '</div>' + sharedFolderChildrenHtml(share, e.children || [], depth + 1);
    }
    else
    {
      var key = parentId + '|' + e.relPath;
      sharedRegistry[key] = { shareId: share.shareId, link: share.link, subPath: e.relPath, name: e.name, permission: share.permission, ownerEmail: share.ownerEmail };
      html += sharedFileRowHtml(key, e.name, ownerLabel, share.permission, depth);
    }
  });
  return html;
}

async function toggleSharedFolder(id)
{
  sharedExpanded[id] = !sharedExpanded[id];
  if (sharedExpanded[id] && !sharedChildren[id])
  {
    var s = sharedEntries.find(function(x){ return sharedEntryId(x) === id; });
    try
    {
      sharedChildren[id] = s && s.link
        ? await Platform.listLinkFolder(s.link, '')
        : await Platform.listSharedFolder(s.shareId, '');
    }
    catch(e) { sharedChildren[id] = []; }
  }
  renderSharedSection();
}

async function leaveShare(e, shareId)
{
  e.stopPropagation();
  try { await Platform.revokeShare(shareId); } catch(err) {}
  refreshSharedSection();
}

/// Entered via a share-link URL (…/#link=<token>): resolves the token, adds
/// the entry to the "Shared with me" section, and opens it if it's a file.
async function initLinkAccess(token)
{
  var meta;
  try { meta = await Platform.linkMeta(token); }
  catch(e) { console.warn('Share link error', e); return; }

  sharedEntries = sharedEntries.filter(function(s){ return s.link !== token; });
  sharedEntries.push({
    link: token,
    ownerEmail: meta.ownerEmail,
    name: meta.name,
    isDir: meta.isDir,
    permission: meta.permission,
    exists: meta.exists
  });
  renderSharedSection();

  if (!meta.exists) return;

  if (meta.isDir)
    await toggleSharedFolder('link:' + token);
  else
    await openSharedRegistry('link:' + token);
}

async function openSharedRegistry(key)
{
  var entry = sharedRegistry[key];
  if (!entry) return;

  var content;
  try
  {
    content = entry.link
      ? await Platform.readLinkFile(entry.link, entry.subPath)
      : await Platform.readSharedFile(entry.shareId, entry.subPath);
  }
  catch(err)
  {
    console.warn('Shared read error', err);
    return;
  }

  if (currentFileId && currentFileId !== SHARED_TMP_ID)
    recordFileHistory(currentFileId);

  var type = typeForFileName(entry.name);

  files[SHARED_TMP_ID] = {
    name: entry.name.replace(/\.[^.]+$/, ''),
    type: type,
    folder: '',
    modified: Date.now(),
    content: content,
    contentLoaded: true
  };

  sharedCtx = {
    shareId: entry.shareId,
    link: entry.link,
    subPath: entry.subPath,
    permission: entry.permission,
    ownerEmail: entry.ownerEmail
  };

  currentFileId = SHARED_TMP_ID;
  switchAppType(type, false);
  renderFileList();

  var file = files[SHARED_TMP_ID];
  if (type === 'doc')           loadDocFile(file);
  else if (type === 'graph')    loadGraphFile(file);
  else if (type === 'notebook') loadNotebookFile(file);
  else if (type === 'glossary') loadGlossaryFile(file);
  else if (type === 'calendar') loadCalendarFile(file);
  else if (type === 'economy')  loadEconomyFile(file);
  else if (type === 'bestiary') loadBestiaryFile(file);
  else                          loadSheetFile(file);

  applySharedReadOnly();
  updateCommentsUI();
  updateLiveSync();
}

function exitSharedMode()
{
  if (!sharedCtx && !files[SHARED_TMP_ID]) return;
  sharedCtx = null;
  delete files[SHARED_TMP_ID];
  applySharedReadOnly();
}

function applySharedReadOnly()
{
  var ro = !!(sharedCtx && sharedCtx.permission !== 'edit');
  var ed = document.getElementById('editor');
  if (ed) ed.readOnly = ro;
  var ti = document.getElementById('doc-title-input');
  if (ti) ti.readOnly = !!sharedCtx; // renaming someone else's file isn't a thing
  document.body.classList.toggle('shared-readonly', ro);

  var banner = document.getElementById('view-only-banner');
  if (!banner) return;
  if (sharedCtx)
  {
    var via = sharedCtx.link ? 'Shared via link by ' : 'Shared by ';
    banner.style.display = '';
    banner.textContent = sharedCtx.permission === 'edit'
      ? via + sharedCtx.ownerEmail + ' — you can edit'
      : sharedCtx.permission === 'comment'
        ? via + sharedCtx.ownerEmail + ' — view & comment only'
        : via + sharedCtx.ownerEmail + ' — view only';
  }
  else
    banner.style.display = 'none';
}

// ── Share modal (owner side) ──

var shareModalTarget = null; // { relPath, isDir, label }

function openShareModal()
{
  var t = contextMenuTarget;
  if (!t || t.kind === 'root') return;

  var isDir = t.kind === 'folder';
  var label = isDir ? t.id : (files[t.id] ? files[t.id].name : t.id);
  shareModalTarget = { relPath: t.id, isDir: isDir, label: label };

  closeContextMenu();

  document.getElementById('share-modal-title').textContent = 'Share "' + label + '"';
  document.getElementById('share-email').value = '';
  document.getElementById('share-error').style.display = 'none';
  document.getElementById('share-modal').style.display = 'flex';
  renderShareModalList();
  renderShareLinksList();
}

// ── Share links ("anyone with the link") ──

function shareLinkUrl(token)
{
  return window.location.origin + '/#link=' + token;
}

async function renderShareLinksList()
{
  if (!shareModalTarget) return;
  var box = document.getElementById('share-links-list');
  var links = [];
  try { links = await Platform.listShareLinks(shareModalTarget.relPath); }
  catch(e) {}

  box.innerHTML = links.length
    ? links.map(function(l){
        return '<div class="share-modal-row">' +
          '<span class="share-modal-email share-link-url" title="' + escAttr(shareLinkUrl(l.id)) + '">' + escHtml(shareLinkUrl(l.id)) + '</span>' +
          sharePermBadge(l.permission) +
          '<button class="share-copy-btn" onclick="copyShareLink(\'' + escAttr(l.id) + '\', this)">Copy</button>' +
          '<button class="gls-card-del" style="position:static" onclick="revokeLinkRow(\'' + escAttr(l.id) + '\')" title="Revoke link">×</button>' +
        '</div>';
      }).join('')
    : '<div class="share-modal-empty">No links yet — anyone who gets one can open this without an account.</div>';
}

async function createLinkRow()
{
  if (!shareModalTarget) return;
  var perm = document.getElementById('share-link-perm').value,
      errEl = document.getElementById('share-error');
  errEl.style.display = 'none';
  try
  {
    await Platform.createShareLink(shareModalTarget.relPath, shareModalTarget.isDir, perm);
    renderShareLinksList();
  }
  catch(e)
  {
    errEl.textContent = e.message || 'Creating the link failed';
    errEl.style.display = '';
  }
}

async function copyShareLink(token, btn)
{
  var url = shareLinkUrl(token);
  try
  {
    await navigator.clipboard.writeText(url);
    btn.textContent = 'Copied!';
    setTimeout(function(){ btn.textContent = 'Copy'; }, 1400);
  }
  catch(e)
  {
    // Clipboard API needs a secure context; fall back to a copyable prompt.
    window.prompt('Copy this link:', url);
  }
}

async function revokeLinkRow(linkId)
{
  try { await Platform.revokeShareLink(linkId); } catch(e) {}
  renderShareLinksList();
}

function closeShareModal()
{
  document.getElementById('share-modal').style.display = 'none';
  shareModalTarget = null;
}

async function renderShareModalList()
{
  if (!shareModalTarget) return;
  var box = document.getElementById('share-modal-list');
  var shares = [];
  try { shares = await Platform.listSharesFor(shareModalTarget.relPath); }
  catch(e) {}

  box.innerHTML = shares.length
    ? shares.map(function(s){
        return '<div class="share-modal-row">' +
          '<span class="share-modal-email">' + escHtml(s.email) + '</span>' +
          sharePermBadge(s.permission) +
          '<button class="gls-card-del" style="position:static" onclick="revokeShareRow(\'' + escAttr(s.id) + '\')" title="Revoke">×</button>' +
        '</div>';
      }).join('')
    : '<div class="share-modal-empty">Not shared with anyone yet.</div>';
}

async function revokeShareRow(shareId)
{
  try { await Platform.revokeShare(shareId); } catch(e) {}
  renderShareModalList();
}

async function submitShare()
{
  if (!shareModalTarget) return;
  var email = document.getElementById('share-email').value.trim(),
      perm  = document.getElementById('share-perm').value,
      errEl = document.getElementById('share-error');
  if (!email) return;

  errEl.style.display = 'none';
  try
  {
    await Platform.shareEntry(shareModalTarget.relPath, shareModalTarget.isDir, email, perm);
    document.getElementById('share-email').value = '';
    renderShareModalList();
  }
  catch(e)
  {
    errEl.textContent = e.message || 'Sharing failed';
    errEl.style.display = '';
  }
}

// ── Comments (any work file; anchors point at a spot inside it) ──

var commentsVisible = false;
var commentsPollTimer = null;
var lastCommentCount = -1;
var pendingAnchor = null;          // captured when the composer gains focus
var lastItemAnchor = null;         // { fileId, kind, id, label } — last opened item modal
var commentAnchorRegistry = {};    // commentId -> parsed anchor (for click-to-navigate)

function commentTarget()
{
  if (Platform.isNative) return null;
  if (!currentFileId || !files[currentFileId]) return null;
  if (sharedCtx && sharedCtx.link) return { link: sharedCtx.link, subPath: sharedCtx.subPath };
  if (sharedCtx) return { share: sharedCtx.shareId, subPath: sharedCtx.subPath };
  if (!workFolderRoot) return null; // not logged in / no workspace
  return { path: currentFileId };
}

function updateCommentsUI()
{
  var fab = document.getElementById('comments-fab');
  if (!fab) return;
  var target = commentTarget();
  fab.style.display = (target && !commentsVisible) ? '' : 'none';
  if (!target)
  {
    commentsVisible = false;
    document.getElementById('comments-panel').style.display = 'none';
    stopCommentsPolling();
    return;
  }
  startCommentsPolling();
  lastCommentCount = -1;
  if (commentsVisible) refreshComments();
  else refreshCommentsCount();
}

function startCommentsPolling()
{
  if (commentsPollTimer) return;
  commentsPollTimer = setInterval(function()
  {
    if (!commentTarget()) { stopCommentsPolling(); return; }
    if (commentsVisible) refreshComments();
    else refreshCommentsCount();
  }, 7000);
}

function stopCommentsPolling()
{
  if (!commentsPollTimer) return;
  clearInterval(commentsPollTimer);
  commentsPollTimer = null;
}

function toggleCommentsPanel()
{
  commentsVisible = !commentsVisible;
  document.getElementById('comments-panel').style.display = commentsVisible ? 'flex' : 'none';
  var fab = document.getElementById('comments-fab');
  if (fab) fab.style.display = (!commentsVisible && commentTarget()) ? '' : 'none';
  if (commentsVisible) { lastCommentCount = -1; refreshComments(); }
}

function setCommentsBadge(n)
{
  var badge = document.getElementById('comments-count');
  if (!badge) return;
  badge.textContent = n;
  badge.style.display = n ? '' : 'none';
}

async function refreshCommentsCount()
{
  var target = commentTarget();
  if (!target) return;
  try { setCommentsBadge((await Platform.listComments(target)).length); }
  catch(e) {}
}

// ── Comment anchors ──

function noteItemAnchor(kind, id, label)
{
  lastItemAnchor = { fileId: currentFileId, kind: kind, id: id, label: label };
}

function captureCommentAnchor()
{
  var f = currentFileId && files[currentFileId];
  if (!f) return;

  var a = null;

  if (f.type === 'doc')
  {
    var ed = document.getElementById('editor');
    if (ed && ed.selectionStart !== ed.selectionEnd)
    {
      var s = ed.selectionStart, e = ed.selectionEnd;
      a = { t: 'text', s: s, e: e, x: ed.value.slice(s, e).slice(0, 120) };
    }
  }
  else if (f.type === 'sheet')
  {
    if (selectionAnchor)
      a = {
        t: 'cells',
        p: sheetPages[sheetPageIdx].name,
        r: (selectionAnchor === selectionEnd || !selectionEnd) ? selectionAnchor : (selectionAnchor + ':' + selectionEnd)
      };
  }
  else if (f.type !== 'graph' && f.type !== 'notebook' && lastItemAnchor && lastItemAnchor.fileId === currentFileId)
    a = { t: 'item', k: lastItemAnchor.kind, id: lastItemAnchor.id, x: lastItemAnchor.label };

  pendingAnchor = a;
  renderPendingAnchor();
}

function clearPendingAnchor()
{
  pendingAnchor = null;
  renderPendingAnchor();
}

function anchorLabel(a)
{
  if (!a) return '';
  if (a.t === 'text')  return '“' + (a.x || '').slice(0, 40) + ((a.x || '').length > 40 ? '…' : '') + '”';
  if (a.t === 'cells') return a.p + ' · ' + a.r;
  if (a.t === 'item')  return a.x || 'item';
  return '';
}

function renderPendingAnchor()
{
  var chip = document.getElementById('comment-anchor-chip');
  if (!chip) return;
  if (!pendingAnchor) { chip.style.display = 'none'; chip.innerHTML = ''; return; }
  chip.style.display = '';
  chip.innerHTML = '<span class="comment-anchor">📍 ' + escHtml(anchorLabel(pendingAnchor)) + '</span>' +
    '<button class="gls-card-del" style="position:static" onclick="clearPendingAnchor()" title="Comment on the whole file instead">×</button>';
}

function gotoCommentAnchor(commentId)
{
  var a = commentAnchorRegistry[commentId];
  var f = currentFileId && files[currentFileId];
  if (!a || !f) return;

  if (a.t === 'text' && f.type === 'doc')
  {
    var ed = document.getElementById('editor');
    if (!ed) return;
    // The stored offsets drift as the doc is edited; re-find the excerpt and
    // fall back to the offsets when it's gone.
    var idx = a.x ? ed.value.indexOf(a.x) : -1;
    var s = idx !== -1 ? idx : Math.min(a.s || 0, ed.value.length),
        e = idx !== -1 ? idx + a.x.length : Math.min(a.e || 0, ed.value.length);
    ed.focus();
    ed.setSelectionRange(s, e);
    var totalLines = ed.value.split('\n').length - 1;
    var frac = totalLines ? (ed.value.slice(0, s).split('\n').length - 1) / totalLines : 0;
    ed.scrollTop = frac * Math.max(0, ed.scrollHeight - ed.clientHeight);
  }
  else if (a.t === 'cells' && f.type === 'sheet')
  {
    var pi = sheetPages.findIndex(function(p){ return p.name === a.p; });
    if (pi !== -1 && pi !== sheetPageIdx) switchSheetPage(pi);
    var parts = (a.r || 'A1').split(':');
    selectCell(parts[0]);
    if (parts[1]) { selectionEnd = parts[1]; renderRangeSelection(); }
  }
  else if (a.t === 'item')
  {
    if (a.k === 'glossary-entry')      { switchGlossaryTab('words'); openGlossaryEntryModal(a.id); }
    else if (a.k === 'glossary-root')  { switchGlossaryTab('roots'); openGlossaryRootModal(a.id); }
    else if (a.k === 'beast')          openBeastModal(a.id);
    else if (a.k === 'eco-currency')   openEcoCurrencyModal(a.id);
    else if (a.k === 'eco-good')       openEcoGoodModal(a.id);
    else if (a.k === 'eco-region')     openEcoRegionModal(a.id);
  }
}

async function refreshComments()
{
  var target = commentTarget();
  if (!target) return;

  var listEl = document.getElementById('comments-list');
  var canPost = !sharedCtx || sharedCtx.permission === 'comment' || sharedCtx.permission === 'edit';
  document.getElementById('comments-compose').style.display = canPost ? '' : 'none';
  document.getElementById('comments-no-perm').style.display = canPost ? 'none' : '';

  var comments = [];
  try { comments = await Platform.listComments(target); }
  catch(e)
  {
    listEl.innerHTML = '<div class="comments-empty">' + escHtml(e.message || 'Could not load comments') + '</div>';
    return;
  }

  setCommentsBadge(comments.length);

  commentAnchorRegistry = {};
  listEl.innerHTML = comments.length
    ? comments.map(function(c){
        var when = new Date(c.createdAt).toLocaleString();
        var anchorHtml = '';
        if (c.anchor)
        {
          try
          {
            var a = JSON.parse(c.anchor);
            commentAnchorRegistry[c.id] = a;
            anchorHtml = '<button class="comment-anchor comment-anchor-link" onclick="gotoCommentAnchor(&quot;' + escAttr(c.id) + '&quot;)" title="Jump to this spot">📍 ' + escHtml(anchorLabel(a)) + '</button>';
          }
          catch(e) {}
        }
        return '<div class="comment-item">' +
          '<div class="comment-meta">' +
            '<span class="comment-author">' + escHtml(c.authorEmail || 'Anonymous') + '</span>' +
            '<span class="comment-when">' + escHtml(when) + '</span>' +
            (c.mine ? '<button class="gls-card-del" style="position:static" onclick="removeComment(&quot;' + escAttr(c.id) + '&quot;)" title="Delete">×</button>' : '') +
          '</div>' +
          anchorHtml +
          '<div class="comment-body">' + escHtml(c.body) + '</div>' +
        '</div>';
      }).join('')
    : '<div class="comments-empty">No comments yet.</div>';

  // Only jump to the newest comment when something actually arrived, so the
  // live refresh never yanks the scroll position mid-read.
  if (comments.length !== lastCommentCount)
    listEl.scrollTop = listEl.scrollHeight;
  lastCommentCount = comments.length;
}

async function submitComment()
{
  var target = commentTarget();
  var input = document.getElementById('comment-input');
  if (!target || !input.value.trim()) return;
  try
  {
    await Platform.addComment(target, input.value.trim(), pendingAnchor ? JSON.stringify(pendingAnchor) : null);
    input.value = '';
    clearPendingAnchor();
    lastCommentCount = -1; // force the scroll-to-newest on the next render
    refreshComments();
  }
  catch(e)
  {
    document.getElementById('comments-list').innerHTML += '<div class="comments-empty">' + escHtml(e.message || 'Posting failed') + '</div>';
  }
}

async function removeComment(id)
{
  try { await Platform.deleteComment(id); } catch(e) {}
  refreshComments();
}

// ── LIVE SYNC (concurrent editing for the JSON data apps) ──
//
// While a Glossary / Bestiary / Calendar / Economy file is open, the server
// is polled for changes to that file (a cheap {modified, hash} stat) so the
// same file can be edited from several devices — or by several people
// through a share — at once, with everyone seeing everyone else's changes
// within a few seconds. Remote and local edits are reconciled with a
// three-way merge at item granularity (see mergeDataContent), which these
// JSON files support cleanly because every item carries a stable id.
// Web only: on Tauri the stat call resolves null and the poller never
// starts — the desktop app keeps its explicit cloud-sync flow.

var LIVE_SYNC_INTERVAL = 3000;

var liveSyncTimer      = null;
var liveSyncFileId     = null;   // file the poller is tracking
var liveSyncBase       = null;   // content as of the last agreement with the server
var liveSyncServerHash = null;   // server hash seen on the previous tick
var liveSyncBusy       = false;

function liveSyncTarget()
{
  if (Platform.isNative) return null;
  var f = currentFileId && files[currentFileId];
  if (!f || !DATA_FILE_TYPES[f.type] || !f.contentLoaded) return null;
  if (currentFileId === SHARED_TMP_ID)
  {
    if (!sharedCtx) return null;
    // Any permission level may watch; only edit may write back.
    return sharedCtx.link
      ? { link: sharedCtx.link, subPath: sharedCtx.subPath }
      : { share: sharedCtx.shareId, subPath: sharedCtx.subPath };
  }
  if (!workFolderRoot) return null; // local-storage mode: nothing to sync with
  return { path: currentFileId };
}

// (Re)aims the poller at whatever is currently open. Called after every
// file open; the tick also calls it when it notices the file changed.
function updateLiveSync()
{
  var target = liveSyncTarget();
  liveSyncFileId     = target ? currentFileId : null;
  liveSyncBase       = target ? (files[currentFileId].content || '') : null;
  liveSyncServerHash = null;

  var pill = document.getElementById('live-sync-pill');
  if (pill) pill.style.display = target ? '' : 'none';

  if (target && !liveSyncTimer)
    liveSyncTimer = setInterval(liveSyncTick, LIVE_SYNC_INTERVAL);
  if (!target && liveSyncTimer)
  {
    clearInterval(liveSyncTimer);
    liveSyncTimer = null;
  }
}

function flashLiveSyncPill()
{
  var pill = document.getElementById('live-sync-pill');
  if (!pill) return;
  pill.textContent = '↻ Updated';
  pill.classList.add('flash');
  setTimeout(function()
  {
    pill.textContent = '● Live';
    pill.classList.remove('flash');
  }, 1600);
}

async function liveSyncTick()
{
  if (liveSyncBusy) return;
  if (currentFileId !== liveSyncFileId) { updateLiveSync(); return; }
  var target = liveSyncTarget();
  if (!target) { updateLiveSync(); return; }

  liveSyncBusy = true;
  try
  {
    var stat;
    try
    {
      stat = target.link  ? await Platform.statLinkFile(target.link, target.subPath)
           : target.share ? await Platform.statSharedFile(target.share, target.subPath)
           :                await Platform.statWorkFile(workFolderRoot, target.path);
    }
    catch(e) { return; } // transient network error, or the file moved remotely — try again next tick

    if (!stat || !stat.hash) return;
    if (stat.hash === liveSyncServerHash) return; // nothing new on the server

    var remote;
    try
    {
      remote = target.link  ? await Platform.readLinkFile(target.link, target.subPath)
             : target.share ? await Platform.readSharedFile(target.share, target.subPath)
             :                await Platform.readWorkFile(workFolderRoot, target.path);
    }
    catch(e) { return; }

    // Things may have changed while awaiting the fetch.
    if (currentFileId !== liveSyncFileId || !files[currentFileId]) return;

    var f     = files[currentFileId];
    var local = f.content || '';
    var merged = (local === liveSyncBase || liveSyncBase === null)
      ? remote
      : mergeDataContent(f.type, liveSyncBase, local, remote);

    if (merged !== local)
    {
      f.content  = merged;
      f.fileTags = undefined; // recomputed lazily from the new content
      f.modified = stat.modified || Date.now();
      applyRemoteDataContent(currentFileId);
      flashLiveSyncPill();
    }

    if (merged !== remote)
    {
      // Our local edits survived the merge — push the merged result back so
      // the other devices converge on it. (Read-only shares never diverge:
      // their UI can't edit, so local always equals base.)
      var canWrite = !sharedCtx || sharedCtx.permission === 'edit';
      if (canWrite)
      {
        try
        {
          if (target.link)       await Platform.writeLinkFile(target.link, target.subPath, merged);
          else if (target.share) await Platform.writeSharedFile(target.share, target.subPath, merged);
          else                   await Platform.writeWorkFile(workFolderRoot, target.path, merged);
          liveSyncBase = merged;
          liveSyncServerHash = null; // our write changed the server hash; re-read it next tick
        }
        catch(e) {}
      }
      else
        liveSyncBase = merged;
    }
    else
    {
      liveSyncBase = remote;
      liveSyncServerHash = stat.hash;
    }
  }
  finally
  {
    liveSyncBusy = false;
  }
}

// Which top-level keys of each data file are item arrays (merged per item);
// everything else is treated as a scalar (local wins only when it changed).
var DATA_ARRAY_KEYS = {
  glossary: ['entries', 'roots'],
  bestiary: ['beasts'],
  calendar: ['seasons', 'months', 'holidays'],
  economy:  ['currencies', 'exchangeRates', 'tradeGoods', 'regions']
};

function mergeDataContent(type, baseStr, localStr, remoteStr)
{
  var base, local, remote;
  try { base = JSON.parse(baseStr || '{}'); } catch(e) { base = {}; }
  try { local = JSON.parse(localStr || '{}'); } catch(e) { return remoteStr; }
  try { remote = JSON.parse(remoteStr || '{}'); } catch(e) { return localStr; }

  var arrayKeys = DATA_ARRAY_KEYS[type] || [];
  var merged = {};
  var seenKeys = {};
  [remote, local, base].forEach(function(o){ Object.keys(o).forEach(function(k){ seenKeys[k] = 1; }); });

  Object.keys(seenKeys).forEach(function(k)
  {
    if (arrayKeys.indexOf(k) !== -1) return;
    var localChanged = JSON.stringify(local[k]) !== JSON.stringify(base[k]);
    var v = localChanged ? local[k] : remote[k];
    if (v !== undefined) merged[k] = v;
  });

  arrayKeys.forEach(function(k)
  {
    merged[k] = mergeItemArrays(base[k] || [], local[k] || [], remote[k] || []);
  });

  return JSON.stringify(merged, null, 2);
}

// Three-way merge of two divergent copies of an item array, by item id.
// Remote order wins (they may have re-ordered); local-only survivors append.
function mergeItemArrays(base, local, remote)
{
  function mapById(arr)
  {
    var m = {};
    arr.forEach(function(x){ if (x && x.id) m[x.id] = x; });
    return m;
  }

  var bm = mapById(base), lm = mapById(local);
  var out = [], seen = {};

  remote.forEach(function(r)
  {
    if (!r || !r.id || seen[r.id]) return;
    seen[r.id] = 1;
    var l = lm[r.id], b = bm[r.id];
    if (!l)
    {
      // Missing locally: new on the remote (keep) or deleted locally (drop).
      if (!b) out.push(r);
      return;
    }
    var localChanged = JSON.stringify(l) !== JSON.stringify(b);
    out.push(localChanged ? l : r);
  });

  local.forEach(function(l)
  {
    if (!l || !l.id) { out.push(l); return; } // id-less item: never drop silently
    if (seen[l.id]) return;
    seen[l.id] = 1;
    var b = bm[l.id];
    // Missing remotely: new local item (keep), or deleted remotely — where a
    // concurrent local edit beats the delete, and an untouched item follows it.
    if (!b || JSON.stringify(l) !== JSON.stringify(b)) out.push(l);
  });

  return out;
}

// Re-points the open editor at freshly merged content and re-renders,
// preserving the user's tab, search, facet filters, and whatever input
// they're typing in (inputs holding focus are never overwritten).
function applyRemoteDataContent(id)
{
  var f = files[id];
  if (!f || currentFileId !== id) return;

  var data;
  try { data = JSON.parse(f.content || '{}'); } catch(e) { return; }

  var focused = document.activeElement;
  function setInput(elId, val)
  {
    var el = document.getElementById(elId);
    if (el && el !== focused) el.value = val;
  }

  if (f.type === 'glossary')
  {
    glsData = data;
    if (!glsData.entries) glsData.entries = [];
    if (!glsData.roots)   glsData.roots   = [];
    setInput('gls-file-tags', (glsData.tags || []).join(', '));
    renderGlossary();
  }
  else if (f.type === 'bestiary')
  {
    bstData = data;
    if (!bstData.beasts) bstData.beasts = [];
    setInput('bst-file-tags', (bstData.tags || []).join(', '));
    renderBestiary();
  }
  else if (f.type === 'calendar')
  {
    calData = data;
    if (!calData.months)   calData.months   = [];
    if (!calData.seasons)  calData.seasons  = [];
    if (!calData.holidays) calData.holidays = [];
    if (!calData.daysPerYear) calData.daysPerYear = 365;
    setInput('cal-file-tags', (calData.tags || []).join(', '));
    setInput('cal-days-per-year', calData.daysPerYear);
    if (calData.epochRealDate) setInput('cal-epoch-real', calData.epochRealDate);
    renderCalendar();
  }
  else if (f.type === 'economy')
  {
    ecoData = data;
    if (!ecoData.currencies)    ecoData.currencies    = [];
    if (!ecoData.exchangeRates) ecoData.exchangeRates = [];
    if (!ecoData.tradeGoods)    ecoData.tradeGoods    = [];
    if (!ecoData.regions)       ecoData.regions       = [];
    setInput('eco-file-tags', (ecoData.tags || []).join(', '));
    renderEconomy();
  }

  renderFileList();
}


init();
