// =========Example code from Tauri docs - call to backend=========
/*
const { invoke } = window.__TAURI__.core;

let greetInputEl;
let greetMsgEl;

async function greet() {
  // Learn more about Tauri commands at https://tauri.app/develop/calling-rust/
  greetMsgEl.textContent = await invoke("greet", { name: greetInputEl.value });
}

window.addEventListener("DOMContentLoaded", () => {
  greetInputEl = document.querySelector("#greet-input");
  greetMsgEl = document.querySelector("#greet-msg");
  document.querySelector("#greet-form").addEventListener("submit", (e) => {
    e.preventDefault();
    greet();
  });
});
*/
//========================================

// TAURI

const { invoke } = window.__TAURI__.core;

// ── STATE ──
let files = {};
let currentFileId = null;
let currentAppType = 'doc';
let newFileType = 'doc';
let editorView = 'write';
let saveTimer = null;
let sheetData = {};
let activeCell =
{
  row: 1,
  col: 1
};

let ROWS = 40,
    COLS = 26;

let documentHeadings = [];
let headingRenderCursor = 0;
let slugCounts = {};

async function createDefaultFile()
{
  let fileContext = await invoke
  (
    "defaultFile",
    {
      name: "Decumano"
    }
  );

  const file = createFile
  (
    fileContext.name,
    fileContext.docType,
    fileContext.content
  );

  openFile(file);
}

function init()
{
  loadFromStorage();
  renderFileList();
  buildSheet();

  if (Object.keys(files).length === 0)
  {
    createDefaultFile();
  }
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

function createFile(name, type, content)
{
  const fileId = 'f_' + Date.now();
  
  if (content === undefined)
  {
    if (type === 'sheet')
      content = '---\ntype: spreadsheet\n---\n\n';

    else
      content = '';
  }

  files[fileId] =
  {
    name, type, content, modified: Date.now()
  };

  saveToStorage();
  
  return fileId;
}

function renderFileList()
{
  const fileList = document.getElementById('file-list');

  const filtered = Object.entries(files).filter(([, f]) => f.type === currentAppType);

  filtered.sort((a, b) => b[1].modified - a[1].modified);

  if (filtered.length === 0)
  {
    fileList.innerHTML =  '<div style="padding: 20px 12px; text-align: center; color: var(--text3); font-size: 12px;">' +
                            'No files yet.<br>Click <strong style=\"color:var(--accent)\">+ New</strong> to start.' +
                          '</div>';
    return;
  }
  
  fileList.innerHTML = filtered.map(([id,f]) => {
    const active = (id === currentFileId);

    const icon = 
      (f.type === 'sheet')
        ?
          '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">'+
            '<rect x="3" y="3" width="18" height="18" rx="2"/>'+
            '<line x1="3" y1="9" x2="21" y2="9"/>'+
            '<line x1="3" y1="15" x2="21" y2="15"/>'+
            '<line x1="9" y1="3" x2="9" y2="21"/>'+
            '<line x1="15" y1="3" x2="15" y2="21"/>'+
          '</svg>'
        :
          '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">'+
            '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>'+
            '<polyline points="14,2 14,8 20,8"/>'+
          '</svg>';
    
    const date = new Date(f.modified).toLocaleDateString(undefined,{month:'short',day:'numeric'});
    
    return  '<div class="file-item ' + (active?'active':'') + '" onclick="openFile(\'' + id + '\')">' +
              '<span class="file-icon">' +
                icon +
              '</span>' +
              '<div class="file-info">'+
                '<div class="file-name">' +
                  escHtml(f.name) +
                '</div>'+
                '<div class="file-meta">' +
                  date +
                '</div>'+
              '</div>'+
              '<span class="file-del" onclick="deleteFile(event,\'' + id + '\')">'+
                '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">'+
                  '<line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>'+
                '</svg>'+
              '</span>'+
            '</div>';
  }).join('');
}

function openFile(id)
{
  if (!files[id])
    return;

  currentFileId = id;
  const file = files[id];
  
  switchAppType(file.type, false);
  renderFileList();

  if (file.type === 'doc')
    loadDocFile(file);

  else
    loadSheetFile(file);
}

function deleteFile(e, id)
{
  e.stopPropagation();
  if (!confirm('Delete "' + files[id].name + '"?'))
    return;

  delete files[id];

  if (currentFileId === id)
  {
    currentFileId = null;
    document.getElementById('doc-title-input').value = '';
    document.getElementById('editor').value = '';
    document.getElementById('toolbar-title').innerHTML = '<span>No document open</span>';
    updatePreview();
    updateStatus();
  }

  saveToStorage();
  renderFileList();
}

function switchAppType(type, rerender)
{
  if (rerender === undefined)
    rerender = true;

  currentAppType = type;
  
  document.getElementById('doc-app').style.display = (type === 'doc') ? 'flex' : 'none';
  document.getElementById('sheet-app').style.display = (type === 'sheet') ? 'flex' : 'none';

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

// ── DOC ──
function loadDocFile(file)
{
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
  if (!currentFileId)
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
          files[currentFileId].modified = Date.now();

        saveToStorage();
        renderFileList();
    },
    800
  );
}

// ── MARKDOWN EXTENSIONS (styled blocks, font spans, table of contents) ──
// Adds Pandoc-style commands on top of standard markdown:
//   :::center / :::right / :::justify / :::left ... :::   wraps a block in a text-align div
//   :::font="Georgia" ... :::                              wraps a block in a font-family div
//   :::center font="Georgia" ... :::                       directives can be combined on one line
//   [text]{font="Georgia"}                                 changes the font of an inline passage
//   [TOC]                                                  expands into a generated table of contents
// Implemented as marked.js extensions (not a post-processing pass) so headings nested inside
// styled blocks are still visited in true document order, keeping TOC anchors in sync.
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
      style += "font-family:'" + value.replace(/'/g, "\\'") + "';";
  }

  return style;
}

function configureMarkedExtensions()
{
  if (typeof marked === 'undefined')
    return;

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
            return '<span style="font-family:\'' + token.fontName.replace(/'/g, "\\'") + '\';">' +
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
        }
      ]
    }
  );
}

configureMarkedExtensions();

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

    var tokens = marked.lexer(source);
    collectHeadings(tokens);

    html = marked.parser(tokens);
  }

  document.getElementById('preview-content').innerHTML = html;
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
  if (keyEvent.ctrlKey || keyEvent.metaKey)
  {
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
  }
);

function wrapSelectedText(before, after)
{
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

function wrapBlock(align)
{
  var editor = document.getElementById('editor'),
      selectionStart = editor.selectionStart,
      selectionEnd = editor.selectionEnd,
      selection = editor.value.slice(selectionStart, selectionEnd) || 'Text';

  var block = '\n:::' + align + '\n' + selection + '\n:::\n';

  editor.value = editor.value.slice(0, selectionStart) + block + editor.value.slice(selectionEnd);
  editor.selectionStart = editor.selectionEnd = selectionStart + block.length;

  editor.focus();
  onEditorChange();
}

function insertTOC()
{
  var editor = document.getElementById('editor'),
      selectionStart = editor.selectionStart,
      lineStart = editor.value.lastIndexOf('\n', selectionStart - 1) + 1,
      marker = '[TOC]\n\n';

  editor.value = editor.value.slice(0, lineStart) + marker + editor.value.slice(lineStart);
  editor.selectionStart = editor.selectionEnd = lineStart + marker.length;

  editor.focus();
  onEditorChange();
}

// Changes the font of the current selection, akin to LaTeX's \fontspec/\setfont commands.
// A selection spanning multiple paragraphs becomes a ":::" block (own font for the whole passage);
// a single-line/inline selection becomes a "[text]{font="..."}" span (own font for that phrase only).
function applyFont()
{
  var editor = document.getElementById('editor'),
      selectionStart = editor.selectionStart,
      selectionEnd = editor.selectionEnd,
      selection = editor.value.slice(selectionStart, selectionEnd),
      fontName = prompt('Font name (e.g. Georgia, "Comic Sans MS", monospace):', '');

  if (!fontName)
    return;

  fontName = fontName.replace(/"/g, '');

  if (selection && !/\n[ \t]*\n/.test(selection))
    wrapSelectedText('[', ']{font="' + fontName + '"}');

  else
    wrapBlock('font="' + fontName + '"');
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

// ── SHEET ──
function buildSheet()
{
  var head = document.getElementById('sheet-head');

  var tableHeadNames ='<tr>'+
                        '<th class="row-header" style="top: 0; left: 0; z-index: 4;">'+
                        '</th>';

  for (var i = 0; i < COLS; i++)
    tableHeadNames +=   '<th class="col-header">' +
                          colName(i) +
                        '</th>';

  head.innerHTML = tableHeadNames +
                      '</tr>';

  var body = document.getElementById('sheet-body'),
      html = '';

  for (var i = 1; i <= ROWS; i++)
  {
    html += '<tr>' +
              '<th class="row-header" style="position: sticky; left: 0; z-index: 1;">' +
                i +
              '</th>';

    for (var j = 0; j < COLS; j++)
    {
      var ref = colName(j) + i;
      html += '<td id="cell-' + ref + '" onclick="selectCell(\'' + ref + '\')">'+
                '<input class="cell-input" id="inp-' + ref + '" value="" onkeydown="cellKey(event, \'' + ref + '\')" onfocus="onCellFocus(\'' + ref + '\')" oninput="onCellInput(\'' + ref + '\')">'+
              '</td>';
    }
    
    html += '</tr>';
  }

  document.getElementById('sheet-body').innerHTML = html;
}

function colName(i)
{
    return String.fromCharCode(65 + i);
}

function colIndex(name)
{
    return name.charCodeAt(0) - 65;
}

function parseName(cellReference)
{
    var matched = cellReference.match(/^([A-Z]+)(\d+)$/);

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

  var parsedName = parseName(name);

  if(!parsedName)
    return;

  activeCell = {
    row: parsedName.row,
    col: colIndex(parsedName.col)
  };

  var selectedCell = document.getElementById('cell-' + name);

  if(selectedCell)
    selectedCell.classList.add('active');

  document.getElementById('cell-ref').value = name;
  document.getElementById('formula-bar').value = sheetData[name] || '';
}

function onCellFocus(name)
{
  selectCell(name);
}

function onCellInput(name)
{
  sheetData[name] = document.getElementById('inp-' + name).value;
  evaluateFormulas();
  saveSheetToFile();
}

function cellKey(keyEvent, name)
{
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
}

function evalCell(ref, val)
{
  if (!val || !val.startsWith('='))
    return val;

  try
  {
    var expr=val.slice(1)
      .replace
      (
        /SUM\(([^)]+)\)/gi,
        function(_, r)
        {
          return sumRange(r);
        }
      )
      .replace
      (
        /AVG\(([^)]+)\)/gi,
        function(_, r)
        {
          return avgRange(r);
        }
      )
      .replace
      (
        /COUNT\(([^)]+)\)/gi,
        function(_, r)
        {
          return countRange(r);
        }
      )
      .replace
      (
        /MAX\(([^)]+)\)/gi,
        function(_, r)
        {
          return maxRange(r);
        }
      )
      .replace
      (
        /MIN\(([^)]+)\)/gi,
        function(_, r)
        {
          return minRange(r);
        }
      )
      .replace
      (
        /([A-Z])(\d+)/g,
        function(_, c, r)
        {
          var v=sheetData[c+r]||'0';
          return
            isNaN(parseFloat(v))?'0':v;
        }
      );

    var result = Function('"use strict"; return ('+expr+')')();

    return  isNaN(result)
            ?
              '#ERR'
            :
              parseFloat(result.toFixed(10)).toString();
  }
  catch(e)
  {
    return '#ERR';
  }
}

function getRangeVals(range)
{
  var parts = range.split(':');

  if (parts.length === 1)
    return  [parseFloat(sheetData[parts[0].trim()]) || 0];

  var pa = parseName(parts[0].trim()),
      pb = parseName(parts[1].trim());
      
  if(!pa || !pb)
    return [];

  var vals=[];

  for (var i = pa.row; i <= pb.row; i++)
    for(var j = colIndex(pa.col); j <= colIndex(pb.col); j++)
      vals.push(parseFloat(sheetData[colName(j) + i]) || 0);
  
  return vals;
}

function sumRange(r)
{
  return getRangeVals(r).reduce
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
  return sumRange(r) / getRangeVals(r).length;
}

function countRange(r)
{
  return getRangeVals(r).length;
}

function maxRange(r)
{
  return Math.max.apply(null, getRangeVals(r));
}

function minRange(r)
{
  return Math.min.apply(null, getRangeVals(r));
}

function evaluateFormulas()
{
  Object.keys(sheetData).forEach
  (
    function(ref)
    {
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
}

function loadSheetFile(f)
{
  sheetData = {};
  var lines = (f.content || '').split('\n'),
      dataStart = 0;
  
  if(lines[0]==='---')
  {
    var end = lines.indexOf('---',1);
    dataStart = end + 2;
  }

  lines.slice(dataStart).forEach
  (
    function(line, ri)
    {
      if(!line.trim())
        return;

      line.split(',').forEach
      (
        function(cell, ci)
        {
          if(ci < COLS && ri < ROWS)
          {
            sheetData[colName(ci) + (ri + 1)] = cell.trim();
          }
        }
      );
    }
  );

  for (var i = 1; i <= ROWS; i++)
    for(var j = 0; j < COLS; j++)
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

  selectCell('A1');
}

function saveSheetToFile()
{
  if(!currentFileId)
    return;

  var csv = '---\ntype: spreadsheet\n---\n\n';

  for (var i = 1; i <= ROWS; i++)
  {
    var row = [];

    for (var j = 0; j < COLS; j++)
      row.push(sheetData[colName(j) + i] || '');

    while (row.length && !row[row.length-1])
      row.pop();

    if (row.some(function(v){ return v; }))
      csv += row.join(',') + '\n';
  }

  files[currentFileId].content = csv;
  files[currentFileId].modified = Date.now();

  saveToStorage();
  renderFileList();
}

// ── MODAL ──
function openNewModal()
{
  newFileType = currentAppType;

  document.getElementById('new-modal').classList.add('open');
  document.getElementById('new-name').value = '';

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
    function(card, i)
    {
      card.classList.toggle('selected', ['doc','sheet'][i] === newFileType);
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
}

function createNewFile()
{
  var name = document.getElementById('new-name').value.trim() || 'Untitled';

  var id = createFile(name, newFileType);

  closeNewModal();
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
      ext = f.type === 'sheet'
            ?
              'csv'
            :
              'md';

  var fileName = f.name.replace(/\s+/g,'_') + '.' + ext;

  try
  {
    var saved = await invoke
    (
      "save_file",
      {
        name: fileName,
        content: f.content
      }
    );

    if (!saved)
      return;
  }
  catch(e)
  {
    console.warn('Save error', e);
    alert('Could not save file.');
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

  reader.onload = function(ev)
  {
    var content = ev.target.result,
        type =  file.name.endsWith('.csv')
                ?
                  'sheet'
                :
                  'doc';

    var name = file.name.replace(/\.[^.]+$/,'');
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

init();
