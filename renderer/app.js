// ---- SVG Support ----
async function renderSvgBlocks() {
  const codeBlocks = previewEl.querySelectorAll('code.language-svg, code[class*="language-svg"]');
  for (const code of codeBlocks) {
    const pre = code.parentElement;
    const source = code.textContent.trim();
    try {
      // Parse SVG string and validate it's valid XML
      const parser = new DOMParser();
      const svgDoc = parser.parseFromString(source, 'image/svg+xml');

      // Check for parsing errors
      if (svgDoc.getElementsByTagName('parsererror').length > 0) {
        throw new Error('Invalid SVG syntax');
      }

      const container = document.createElement('div');
      container.className = 'svg-diagram';
      container.innerHTML = source;
      pre.replaceWith(container);
    } catch (e) {
      console.warn('SVG render error:', e);
      // Keep the code block visible if SVG rendering fails
    }
  }
}

// ---- Mermaid ----
let mermaidReady = false;
let mermaidModule = null;

async function initMermaid() {
  try {
    // Dynamic import since mermaid is ESM
    mermaidModule = await import('../node_modules/mermaid/dist/mermaid.esm.min.mjs');
    mermaidModule.default.initialize({
      startOnLoad: false,
      theme: 'default',
      securityLevel: 'loose',
    });
    mermaidReady = true;
  } catch (e) {
    console.warn('Mermaid failed to load:', e);
  }
}

async function renderMermaidBlocks() {
  if (!mermaidReady) return;

  const codeBlocks = previewEl.querySelectorAll('code.language-mermaid, code[class*="language-mermaid"], code[class*="language- mermaid"]');
  for (const code of codeBlocks) {
    const pre = code.parentElement;
    const source = code.textContent;
    try {
      const id = 'mermaid-' + Math.random().toString(36).slice(2, 10);
      const { svg } = await mermaidModule.default.render(id, source);
      const container = document.createElement('div');
      container.className = 'mermaid-diagram';
      container.innerHTML = svg;
      pre.replaceWith(container);
    } catch (e) {
      console.warn('Mermaid render error:', e);
      showDiagramError(pre, source, e);
    }
  }
}

// Replace a failed diagram's <pre> with a visible error box (message +
// collapsible source) instead of silently leaving raw code, so a bad
// diagram explains itself (e.g. a stray ";" in a sequence message).
function showDiagramError(pre, source, err) {
  const box = document.createElement('div');
  box.className = 'diagram-error';
  box.innerHTML =
    '<div class="diagram-error-title">⚠ Diagram failed to render</div>' +
    '<pre class="diagram-error-msg"></pre>' +
    '<details><summary>Show diagram source</summary><pre class="diagram-error-src"></pre></details>';
  box.querySelector('.diagram-error-msg').textContent =
    (err && err.message) ? err.message : String(err);
  box.querySelector('.diagram-error-src').textContent = source;
  pre.replaceWith(box);
}

/**
 * Capture all SVG blocks as PNG data URLs for Word export.
 * Returns an object mapping source text → { dataUrl, width, height }
 */
async function captureSvgImages(markdown) {
  const images = {};

  // Extract SVG blocks from markdown
  const regex = /```\s*svg\s*\n([\s\S]*?)```/g;
  let match;
  const sources = [];
  while ((match = regex.exec(markdown)) !== null) {
    sources.push(match[1].trim());
  }
  if (sources.length === 0) return images;

  for (const source of sources) {
    try {
      // Parse and validate SVG
      const parser = new DOMParser();
      const svgDoc = parser.parseFromString(source, 'image/svg+xml');

      if (svgDoc.getElementsByTagName('parsererror').length > 0) {
        console.warn('Invalid SVG in export:', source.slice(0, 50));
        continue;
      }

      const svgEl = svgDoc.querySelector('svg');
      if (!svgEl) continue;

      // Get dimensions
      let svgWidth, svgHeight;
      const viewBox = svgEl.getAttribute('viewBox');
      if (viewBox) {
        const parts = viewBox.split(/[\s,]+/).map(Number);
        if (parts.length === 4) {
          svgWidth = parts[2];
          svgHeight = parts[3];
        }
      }
      if (!svgWidth) svgWidth = parseFloat(svgEl.getAttribute('width')) || 600;
      if (!svgHeight) svgHeight = parseFloat(svgEl.getAttribute('height')) || 400;

      // Set explicit dimensions
      svgEl.setAttribute('width', String(svgWidth));
      svgEl.setAttribute('height', String(svgHeight));
      const fixedSvg = new XMLSerializer().serializeToString(svgEl);

      // Convert SVG to PNG
      const svgBase64 = btoa(unescape(encodeURIComponent(fixedSvg)));
      const svgDataUrl = 'data:image/svg+xml;base64,' + svgBase64;

      const img = new Image();
      await new Promise((resolve, reject) => {
        img.onload = resolve;
        img.onerror = reject;
        img.src = svgDataUrl;
      });

      const scale = 3;
      const canvas = document.createElement('canvas');
      canvas.width = svgWidth * scale;
      canvas.height = svgHeight * scale;
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.scale(scale, scale);
      ctx.drawImage(img, 0, 0, svgWidth, svgHeight);

      const dataUrl = canvas.toDataURL('image/png');
      images[source] = {
        dataUrl,
        width: svgWidth,
        height: svgHeight,
      };
    } catch (e) {
      console.warn('SVG capture error:', e);
    }
  }

  return images;
}

/**
 * Pre-render all mermaid blocks to PNG data URLs for Word export.
 * Returns an object mapping source text → { dataUrl, width, height }
 */
async function captureMermaidImages(markdown) {
  if (!mermaidReady) return {};
  const images = {};

  // Extract mermaid blocks from markdown — handles ```mermaid and ``` mermaid (with space)
  const regex = /```\s*mermaid\s*\n([\s\S]*?)```/g;
  let match;
  const sources = [];
  while ((match = regex.exec(markdown)) !== null) {
    sources.push(match[1].trim());
  }
  if (sources.length === 0) return images;

  for (const source of sources) {
    try {
      const id = 'cap-' + Math.random().toString(36).slice(2, 10);
      const { svg } = await mermaidModule.default.render(id, source);

      // Parse SVG to get actual dimensions from viewBox or attributes
      const parser = new DOMParser();
      const svgDoc = parser.parseFromString(svg, 'image/svg+xml');
      const svgEl = svgDoc.querySelector('svg');

      let svgWidth, svgHeight;
      const viewBox = svgEl?.getAttribute('viewBox');
      if (viewBox) {
        const parts = viewBox.split(/[\s,]+/).map(Number);
        if (parts.length === 4) {
          svgWidth = parts[2];
          svgHeight = parts[3];
        }
      }
      // Fallback to explicit width/height attributes
      if (!svgWidth) svgWidth = parseFloat(svgEl?.getAttribute('width')) || 800;
      if (!svgHeight) svgHeight = parseFloat(svgEl?.getAttribute('height')) || 600;

      // Set explicit dimensions on SVG so Image renders at full size
      svgEl.setAttribute('width', String(svgWidth));
      svgEl.setAttribute('height', String(svgHeight));
      const fixedSvg = new XMLSerializer().serializeToString(svgEl);

      // Convert SVG to PNG via base64 data URL (blob URLs get blocked by canvas security)
      const svgBase64 = btoa(unescape(encodeURIComponent(fixedSvg)));
      const svgDataUrl = 'data:image/svg+xml;base64,' + svgBase64;

      const img = new Image();
      await new Promise((resolve, reject) => {
        img.onload = resolve;
        img.onerror = reject;
        img.src = svgDataUrl;
      });

      const scale = 3; // high-res for crisp text in Word
      const canvas = document.createElement('canvas');
      canvas.width = svgWidth * scale;
      canvas.height = svgHeight * scale;
      const ctx = canvas.getContext('2d');
      // White background for clean Word rendering
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.scale(scale, scale);
      ctx.drawImage(img, 0, 0, svgWidth, svgHeight);

      const dataUrl = canvas.toDataURL('image/png');

      // Store under multiple keys to handle whitespace differences
      const entry = {
        dataUrl,
        width: svgWidth,
        height: svgHeight,
      };
      // Store under trimmed source (matches marked lexer's token.text)
      images[source] = entry;
    } catch (e) {
      console.warn('Mermaid capture error for diagram:', e);
    }
  }

  return images;
}

initMermaid();

// ---- State ----
let currentTree = [];
let currentFilePath = null;
let currentContent = '';
let currentTitle = '';

// ---- DOM ----
const $ = (sel) => document.querySelector(sel);
const btnOpenFolder = $('#btn-open-folder');
const btnOpenEmpty = $('#btn-open-empty');
const btnOpenFile = $('#btn-open-file');
const btnOpenFileEmpty = $('#btn-open-file-empty');
const btnExportWord = $('#btn-export-word');
const fileTreeEl = $('#file-tree');
const previewEl = $('#preview');
const emptyState = $('#empty-state');
const folderName = $('#folder-name');
const statusText = $('#status-text');
const statusFile = $('#status-file');
const sidebar = $('#sidebar');
const resizeHandle = $('#resize-handle');

// ---- Open Folder ----
async function openFolder() {
  const result = await window.lumadesk.openFolder();
  if (!result) return;

  currentTree = result.tree;
  folderName.textContent = result.path.split(/[/\\]/).pop();
  renderTree(result.tree);
  // Folder mode: enable the sidebar file/content search.
  clearFolderSearch();
  $('#sidebar-search-wrap').style.display = 'block';
  statusText.textContent = `Opened: ${result.path}`;

  // Auto-select first file
  const firstFile = findFirstFile(result.tree);
  if (firstFile) {
    await selectFile(firstFile.path);
  }
}

btnOpenFolder.addEventListener('click', openFolder);
btnOpenEmpty.addEventListener('click', openFolder);

// ---- Open Single File ----
async function openSingleFile() {
  const result = await window.lumadesk.openFile();
  if (!result) return;
  if (result.error) {
    statusText.textContent = `Error: ${result.error}`;
    return;
  }

  // Single-file mode: show just this file in the sidebar so it stays
  // visible and selectable, then render it via the shared select path.
  currentTree = [{ type: 'file', name: result.fileName, path: result.path }];
  folderName.textContent = result.path.split(/[/\\]/).pop();
  renderTree(currentTree);
  // Single-file mode: no folder to search across.
  clearFolderSearch();
  $('#sidebar-search-wrap').style.display = 'none';
  await selectFile(result.path);
  statusText.textContent = `Opened: ${result.path}`;
}

btnOpenFile.addEventListener('click', openSingleFile);
btnOpenFileEmpty.addEventListener('click', openSingleFile);

// ---- File Tree ----
function renderTree(tree, depth = 0) {
  if (depth === 0) fileTreeEl.innerHTML = '';
  const container = depth === 0 ? fileTreeEl : document.createDocumentFragment();

  for (const item of tree) {
    if (item.type === 'folder') {
      const folder = document.createElement('div');
      folder.className = `tree-folder depth-${depth}`;

      const header = document.createElement('div');
      header.className = 'tree-folder-header';
      header.innerHTML = `
        <svg class="chevron" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"/></svg>
        <span>${item.name}</span>
      `;
      header.addEventListener('click', () => {
        header.classList.toggle('collapsed');
        children.classList.toggle('collapsed');
      });

      const children = document.createElement('div');
      children.className = 'tree-folder-children';

      folder.appendChild(header);
      folder.appendChild(children);

      // Render children into the children container
      for (const child of item.children) {
        if (child.type === 'folder') {
          const subContainer = document.createElement('div');
          subContainer.className = `tree-folder depth-${depth + 1}`;
          const subHeader = document.createElement('div');
          subHeader.className = 'tree-folder-header';
          subHeader.innerHTML = `
            <svg class="chevron" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"/></svg>
            <span>${child.name}</span>
          `;
          const subChildren = document.createElement('div');
          subChildren.className = 'tree-folder-children';
          subHeader.addEventListener('click', () => {
            subHeader.classList.toggle('collapsed');
            subChildren.classList.toggle('collapsed');
          });
          renderFilesInto(subChildren, child.children, depth + 2);
          subContainer.appendChild(subHeader);
          subContainer.appendChild(subChildren);
          children.appendChild(subContainer);
        } else {
          renderFileItem(children, child, depth + 1);
        }
      }

      container.appendChild(folder);
    } else {
      renderFileItem(container, item, depth);
    }
  }

  if (depth === 0 && container !== fileTreeEl) {
    fileTreeEl.appendChild(container);
  }
}

function renderFilesInto(container, items, depth) {
  for (const item of items) {
    if (item.type === 'folder') {
      const folder = document.createElement('div');
      folder.className = `tree-folder depth-${depth}`;
      const header = document.createElement('div');
      header.className = 'tree-folder-header';
      header.innerHTML = `
        <svg class="chevron" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"/></svg>
        <span>${item.name}</span>
      `;
      const children = document.createElement('div');
      children.className = 'tree-folder-children';
      header.addEventListener('click', () => {
        header.classList.toggle('collapsed');
        children.classList.toggle('collapsed');
      });
      renderFilesInto(children, item.children, depth + 1);
      folder.appendChild(header);
      folder.appendChild(children);
      container.appendChild(folder);
    } else {
      renderFileItem(container, item, depth);
    }
  }
}

function renderFileItem(container, item, depth) {
  const fileEl = document.createElement('div');
  fileEl.className = `tree-file depth-${depth}`;
  fileEl.dataset.path = item.path;
  fileEl.innerHTML = `
    <svg class="tree-file-icon" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
    <span>${item.name}</span>
  `;
  fileEl.addEventListener('click', () => selectFile(item.path));
  fileEl.addEventListener('contextmenu', (e) => {
    e.preventDefault();
    openFileContextMenu(e.clientX, e.clientY, item.path);
  });
  container.appendChild(fileEl);
}

function findFirstFile(tree) {
  for (const item of tree) {
    if (item.type === 'file') return item;
    if (item.type === 'folder' && item.children) {
      const found = findFirstFile(item.children);
      if (found) return found;
    }
  }
  return null;
}

// ---- Select & Preview File ----
async function selectFile(filePath) {
  // Guard unsaved edits when switching files.
  if (editorState.active && editorState.dirty && filePath !== currentFilePath) {
    if (!confirm('You have unsaved changes. Discard them and switch files?')) return;
  }

  // Update active state in tree
  document.querySelectorAll('.tree-file.active').forEach(el => el.classList.remove('active'));
  const fileEl = document.querySelector(`.tree-file[data-path="${CSS.escape(filePath)}"]`);
  if (fileEl) fileEl.classList.add('active');

  const result = await window.lumadesk.readFile(filePath);
  if (result.error) {
    statusText.textContent = `Error: ${result.error}`;
    return;
  }

  currentFilePath = filePath;
  currentContent = result.content;
  currentTitle = result.fileName;

  const html = await window.lumadesk.renderMarkdown(result.content);
  previewEl.innerHTML = html;
  previewEl.style.display = 'block';
  emptyState.style.display = 'none';

  // Close any open search when switching files
  if (searchBar.style.display !== 'none') closeSearch();

  // Render SVG and mermaid diagrams
  await renderSvgBlocks();
  await renderMermaidBlocks();

  btnExportWord.disabled = false;
  btnExportPdf.disabled = false;
  btnExportPpt.disabled = false;
  btnPresent.disabled = false;
  btnEdit.disabled = false;
  statusFile.textContent = filePath.split(/[/\\]/).pop();
  statusText.textContent = 'Ready';

  // Scroll to top
  $('#preview-container').scrollTop = 0;
  previewPane.scrollTop = 0;

  // If editing, load the newly selected file into the editor too.
  if (editorState.active) {
    editorTextarea.value = currentContent;
    editorState.fm = '';
    if (editorState.mode === 'wysiwyg') {
      const { fm, body } = splitFrontmatter(currentContent);
      editorState.fm = fm;
      await renderWysiwygFromMarkdown(body);
    }
    setEditorDirty(false);
  }

  // If this file was opened from a folder-search content result, re-apply the
  // query as an in-document highlight and jump to the first hit.
  if (pendingPreviewHighlight) {
    const q = pendingPreviewHighlight;
    pendingPreviewHighlight = '';
    searchBar.style.display = 'flex';
    searchInput.value = q;
    highlightSearch(q);
  }
}

// ---- Word Export Settings (font/size for body text) ----
const btnWordSettings = $('#btn-word-settings');
const wordSettingsPopover = $('#word-settings-popover');
const wordFontInput = $('#word-font-input');
const wordSizeInput = $('#word-size-input');
const wordMermaidSidecar = $('#word-mermaid-sidecar');

const WORD_SETTINGS_KEY = 'lumadesk.wordExportSettings';

function loadWordSettings() {
  try {
    const raw = localStorage.getItem(WORD_SETTINGS_KEY);
    if (!raw) return { font: '', size: '', mermaidSidecar: true };
    const parsed = JSON.parse(raw);
    return {
      font: parsed.font || '',
      size: parsed.size || '',
      mermaidSidecar: parsed.mermaidSidecar !== false,
    };
  } catch {
    return { font: '', size: '', mermaidSidecar: true };
  }
}

function saveWordSettings() {
  const data = {
    font: wordFontInput.value.trim(),
    size: wordSizeInput.value.trim(),
    mermaidSidecar: wordMermaidSidecar.checked,
  };
  localStorage.setItem(WORD_SETTINGS_KEY, JSON.stringify(data));
}

(function initWordSettings() {
  const s = loadWordSettings();
  wordFontInput.value = s.font;
  wordSizeInput.value = s.size;
  wordMermaidSidecar.checked = s.mermaidSidecar;
})();

btnWordSettings.addEventListener('click', (e) => {
  e.stopPropagation();
  wordSettingsPopover.style.display = wordSettingsPopover.style.display === 'none' ? 'block' : 'none';
});

document.addEventListener('click', (e) => {
  if (!wordSettingsPopover.contains(e.target) && e.target !== btnWordSettings && !btnWordSettings.contains(e.target)) {
    wordSettingsPopover.style.display = 'none';
  }
});

wordFontInput.addEventListener('change', saveWordSettings);
wordSizeInput.addEventListener('change', saveWordSettings);
wordMermaidSidecar.addEventListener('change', saveWordSettings);

wordSettingsPopover.querySelectorAll('.preset-btn').forEach((btn) => {
  btn.addEventListener('click', () => {
    wordFontInput.value = btn.dataset.font || '';
    wordSizeInput.value = btn.dataset.size || '';
    saveWordSettings();
    statusText.textContent = btn.dataset.font
      ? `Word body: ${btn.dataset.font} ${btn.dataset.size}pt`
      : 'Word body: default';
  });
});

// ---- Word Export ----
btnExportWord.addEventListener('click', async () => {
  if (!currentContent) return;

  statusText.textContent = 'Rendering diagrams...';
  btnExportWord.disabled = true;

  // Capture SVG and mermaid diagrams as PNG for embedding in Word
  const svgImages = await captureSvgImages(currentContent);
  const mermaidImages = await captureMermaidImages(currentContent);

  statusText.textContent = 'Exporting to Word...';

  const settings = loadWordSettings();
  const result = await window.lumadesk.exportWordWithImages({
    title: currentTitle,
    content: currentContent,
    author: '',
    mermaidImages,
    svgImages,
    bodyFont: settings.font || undefined,
    bodySize: settings.size ? Number(settings.size) : undefined,
    saveMermaidSidecar: settings.mermaidSidecar,
  });

  if (result.success) {
    statusText.textContent = `Exported: ${result.path}`;
  } else if (result.error) {
    statusText.textContent = `Export failed: ${result.error}`;
  } else {
    statusText.textContent = 'Export cancelled';
  }

  btnExportWord.disabled = false;
});

// ---- File Watcher ----
window.lumadesk.onFileChanged(async (filePath) => {
  // Don't hot-reload over an open edit session (would discard the user's work).
  if (editorState.active && filePath === currentFilePath) {
    statusText.textContent = 'File changed on disk — not reloaded while editing';
    return;
  }
  if (filePath === currentFilePath) {
    const result = await window.lumadesk.readFile(filePath);
    if (!result.error) {
      currentContent = result.content;
      const html = await window.lumadesk.renderMarkdown(result.content);
      previewEl.innerHTML = html;
      await renderSvgBlocks();
      await renderMermaidBlocks();
      // Re-apply active search after hot-reload
      if (searchBar.style.display !== 'none' && searchInput.value) {
        highlightSearch(searchInput.value);
      }
      statusText.textContent = 'File updated';
    }
  }
});

window.lumadesk.onTreeChanged((tree) => {
  currentTree = tree;
  renderTree(tree);
  // Re-highlight active file
  if (currentFilePath) {
    const fileEl = document.querySelector(`.tree-file[data-path="${CSS.escape(currentFilePath)}"]`);
    if (fileEl) fileEl.classList.add('active');
  }
});

// ---- Resize Handle ----
let isResizing = false;

resizeHandle.addEventListener('mousedown', (e) => {
  isResizing = true;
  e.preventDefault();
});

document.addEventListener('mousemove', (e) => {
  if (!isResizing) return;
  const width = Math.min(500, Math.max(180, e.clientX));
  sidebar.style.width = width + 'px';
});

document.addEventListener('mouseup', () => {
  isResizing = false;
});

// ---- PDF Export ----
const btnExportPdf = $('#btn-export-pdf');

// A doc is a slide deck only on an explicit signal: Marp "marp: true"
// frontmatter or "# SLIDE N --" markers. Bare "---" lines are NOT used —
// they are ordinary horizontal rules in normal documents.
function isSlideDeck(md) {
  if (!md) return false;
  const fmMatch = md.match(/^﻿?---[ \t]*\r?\n([\s\S]*?)\r?\n---[ \t]*(\r?\n|$)/);
  if (fmMatch && /^\s*marp\s*:\s*true\s*$/m.test(fmMatch[1])) return true;
  if (/^#\s+SLIDE\s+\d+\s*--/m.test(md)) return true;
  return false;
}

btnExportPdf.addEventListener('click', async () => {
  if (!currentContent) return;

  statusText.textContent = 'Exporting to PDF...';
  btnExportPdf.disabled = true;

  // Send the rendered HTML (including mermaid SVGs) for a clean document PDF.
  // Strip search highlights from a clone so they don't leak into the exported PDF.
  const clone = previewEl.cloneNode(true);
  clone.querySelectorAll('mark.search-hit').forEach((m) => {
    const parent = m.parentNode;
    while (m.firstChild) parent.insertBefore(m.firstChild, m);
    parent.removeChild(m);
  });
  const result = await window.lumadesk.exportPdf({
    title: currentTitle,
    html: clone.innerHTML,
    slides: isSlideDeck(currentContent),
  });

  if (result.success) {
    statusText.textContent = `Exported: ${result.path}`;
  } else if (result.error) {
    statusText.textContent = `Export failed: ${result.error}`;
  } else {
    statusText.textContent = 'Export cancelled';
  }

  btnExportPdf.disabled = false;
});

// ---- Mermaid Fullscreen Modal ----
const modal = $('#mermaid-modal');
const modalCanvas = $('#modal-canvas');
const modalDiagram = $('#modal-diagram');
const modalZoomLevel = $('#modal-zoom-level');

let modalScale = 1;
let modalTranslate = { x: 0, y: 0 };
let modalPanning = false;
let modalLastMouse = { x: 0, y: 0 };

function openMermaidModal(svgHtml) {
  modalDiagram.innerHTML = svgHtml;
  modalScale = 1;
  modalTranslate = { x: 0, y: 0 };
  updateModalTransform();
  modal.style.display = 'flex';
  document.body.style.overflow = 'hidden';

  // Auto-fit
  requestAnimationFrame(() => {
    const svg = modalDiagram.querySelector('svg');
    if (!svg) return;
    const svgW = svg.getBoundingClientRect().width;
    const svgH = svg.getBoundingClientRect().height;
    if (!svgW || !svgH) return;
    const viewW = window.innerWidth - 120;
    const viewH = window.innerHeight - 120;
    const fit = Math.min(viewW / svgW, viewH / svgH, 3);
    if (fit > 1.1) {
      modalScale = fit;
      updateModalTransform();
    }
  });
}

function closeMermaidModal() {
  modal.style.display = 'none';
  document.body.style.overflow = '';
}

function updateModalTransform() {
  modalDiagram.style.transform = `translate(${modalTranslate.x}px, ${modalTranslate.y}px) scale(${modalScale})`;
  modalZoomLevel.textContent = Math.round(modalScale * 100) + '%';
}

$('#modal-close').addEventListener('click', closeMermaidModal);
$('#modal-zoom-in').addEventListener('click', () => { modalScale = Math.min(modalScale * 1.3, 8); updateModalTransform(); });
$('#modal-zoom-out').addEventListener('click', () => { modalScale = Math.max(modalScale * 0.7, 0.2); updateModalTransform(); });
$('#modal-reset').addEventListener('click', () => { modalScale = 1; modalTranslate = { x: 0, y: 0 }; updateModalTransform(); });

modalCanvas.addEventListener('wheel', (e) => {
  e.preventDefault();
  const delta = e.deltaY > 0 ? 0.9 : 1.1;
  modalScale = Math.min(Math.max(modalScale * delta, 0.2), 8);
  updateModalTransform();
});

modalCanvas.addEventListener('mousedown', (e) => {
  if (e.button !== 0) return;
  modalPanning = true;
  modalLastMouse = { x: e.clientX, y: e.clientY };
  modalDiagram.style.transition = 'none';
});

document.addEventListener('mousemove', (e) => {
  if (!modalPanning) return;
  modalTranslate.x += e.clientX - modalLastMouse.x;
  modalTranslate.y += e.clientY - modalLastMouse.y;
  modalLastMouse = { x: e.clientX, y: e.clientY };
  updateModalTransform();
});

document.addEventListener('mouseup', () => {
  if (modalPanning) {
    modalPanning = false;
    modalDiagram.style.transition = 'transform 0.1s ease-out';
  }
});

// Click on mermaid/svg diagram in preview → open modal
previewEl.addEventListener('click', (e) => {
  const diagram = e.target.closest('.mermaid-diagram, .svg-diagram');
  if (diagram) {
    openMermaidModal(diagram.innerHTML);
  }
});

// In-page anchor links (Table of Contents) → smooth-scroll to the heading.
// Handled explicitly because the preview scrolls inside #preview-container,
// and we want smooth scrolling without pushing a hash onto the URL.
previewEl.addEventListener('click', (e) => {
  const link = e.target.closest('a[href^="#"]');
  if (!link) return;
  const id = decodeURIComponent(link.getAttribute('href').slice(1));
  if (!id) return;
  // Match by id, or by name attribute as a fallback.
  const target =
    previewEl.querySelector(`#${CSS.escape(id)}`) ||
    previewEl.querySelector(`[name="${CSS.escape(id)}"]`);
  if (target) {
    e.preventDefault();
    target.scrollIntoView({ block: 'start', behavior: 'smooth' });
  }
});

// ---- Mermaid Playground ----
const playgroundModal = $('#playground-modal');
const playgroundInput = $('#playground-input');
const playgroundCanvas = $('#playground-canvas');
const playgroundDiagram = $('#playground-diagram');
const playgroundStatus = $('#playground-status');
const playgroundZoomLevel = $('#playground-zoom-level');
const playgroundDivider = document.querySelector('.playground-divider');

const PLAYGROUND_DEFAULT = `flowchart LR
    A[Start] --> B{Decision}
    B -->|Yes| C[Do this]
    B -->|No| D[Do that]
    C --> E[End]
    D --> E`;

let pgScale = 1;
let pgTranslate = { x: 0, y: 0 };
let pgPanning = false;
let pgLastMouse = { x: 0, y: 0 };
let pgRenderTimer = null;

function updatePlaygroundTransform() {
  playgroundDiagram.style.transform = `translate(${pgTranslate.x}px, ${pgTranslate.y}px) scale(${pgScale})`;
  playgroundZoomLevel.textContent = Math.round(pgScale * 100) + '%';
}

function resetPlaygroundView() {
  pgScale = 1;
  pgTranslate = { x: 0, y: 0 };
  updatePlaygroundTransform();
}

async function renderPlayground() {
  const source = playgroundInput.value.trim();

  if (!source) {
    playgroundDiagram.className = 'empty';
    playgroundDiagram.textContent = 'Type or paste mermaid code to render';
    playgroundStatus.textContent = 'Empty';
    playgroundStatus.classList.remove('error');
    return;
  }

  if (!mermaidReady) {
    playgroundStatus.textContent = 'Loading...';
    return;
  }

  playgroundStatus.textContent = 'Rendering...';
  playgroundStatus.classList.remove('error');

  try {
    const id = 'pg-' + Math.random().toString(36).slice(2, 10);
    const { svg } = await mermaidModule.default.render(id, source);
    playgroundDiagram.className = '';
    playgroundDiagram.innerHTML = svg;
    playgroundStatus.textContent = 'OK';
  } catch (err) {
    playgroundDiagram.className = 'error';
    playgroundDiagram.textContent = String(err?.message || err);
    playgroundStatus.textContent = 'Error';
    playgroundStatus.classList.add('error');
  }
}

function schedulePlaygroundRender() {
  clearTimeout(pgRenderTimer);
  pgRenderTimer = setTimeout(renderPlayground, 250);
}

function openPlayground() {
  playgroundModal.style.display = 'flex';
  if (!playgroundInput.value) playgroundInput.value = PLAYGROUND_DEFAULT;
  resetPlaygroundView();
  renderPlayground();
  setTimeout(() => playgroundInput.focus(), 0);
}

function closePlayground() {
  playgroundModal.style.display = 'none';
}

playgroundInput.addEventListener('input', schedulePlaygroundRender);
$('#btn-mermaid-playground').addEventListener('click', openPlayground);
$('#playground-close').addEventListener('click', closePlayground);
$('#playground-zoom-in').addEventListener('click', () => {
  pgScale = Math.min(pgScale * 1.3, 8);
  updatePlaygroundTransform();
});
$('#playground-zoom-out').addEventListener('click', () => {
  pgScale = Math.max(pgScale * 0.7, 0.2);
  updatePlaygroundTransform();
});
$('#playground-reset').addEventListener('click', resetPlaygroundView);

playgroundCanvas.addEventListener('wheel', (e) => {
  e.preventDefault();
  const delta = e.deltaY > 0 ? 0.9 : 1.1;
  pgScale = Math.min(Math.max(pgScale * delta, 0.2), 8);
  updatePlaygroundTransform();
}, { passive: false });

playgroundCanvas.addEventListener('mousedown', (e) => {
  if (e.button !== 0) return;
  pgPanning = true;
  pgLastMouse = { x: e.clientX, y: e.clientY };
  playgroundDiagram.style.transition = 'none';
});

document.addEventListener('mousemove', (e) => {
  if (!pgPanning) return;
  pgTranslate.x += e.clientX - pgLastMouse.x;
  pgTranslate.y += e.clientY - pgLastMouse.y;
  pgLastMouse = { x: e.clientX, y: e.clientY };
  updatePlaygroundTransform();
});

document.addEventListener('mouseup', () => {
  if (pgPanning) {
    pgPanning = false;
    playgroundDiagram.style.transition = 'transform 0.1s ease-out';
  }
});

// Resizable divider between editor and canvas
const playgroundEditor = document.querySelector('.playground-editor');
let pgResizing = false;

playgroundDivider.addEventListener('mousedown', (e) => {
  pgResizing = true;
  playgroundDivider.classList.add('active');
  document.body.style.userSelect = 'none';
  e.preventDefault();
});

document.addEventListener('mousemove', (e) => {
  if (!pgResizing) return;
  const total = playgroundModal.clientWidth;
  // Allow editor to shrink down to 0 (snap-collapse near the left edge)
  // and canvas to keep at least 200px on the right
  let w = Math.max(0, Math.min(e.clientX, total - 200));
  if (w < 60) w = 0; // snap-collapse
  playgroundEditor.style.width = w + 'px';
  playgroundEditor.classList.toggle('collapsed', w === 0);
});

document.addEventListener('mouseup', () => {
  if (pgResizing) {
    pgResizing = false;
    playgroundDivider.classList.remove('active');
    document.body.style.userSelect = '';
  }
});

// Double-click divider to toggle editor visibility
playgroundDivider.addEventListener('dblclick', () => {
  const collapsed = playgroundEditor.classList.toggle('collapsed');
  playgroundEditor.style.width = collapsed ? '0px' : '40%';
});

// ---- File Filter ----
const btnFilter = $('#btn-filter');
const filterPopover = $('#filter-popover');
const filterInput = $('#filter-input');
const filterLabel = $('#filter-label');
const btnFilterApply = $('#btn-filter-apply');

btnFilter.addEventListener('click', (e) => {
  e.stopPropagation();
  filterPopover.style.display = filterPopover.style.display === 'none' ? 'block' : 'none';
});

document.addEventListener('click', (e) => {
  if (!filterPopover.contains(e.target) && e.target !== btnFilter) {
    filterPopover.style.display = 'none';
  }
});

btnFilterApply.addEventListener('click', async () => {
  const exts = filterInput.value.trim().split(/\s+/).filter(Boolean);
  if (exts.length === 0) return;
  filterLabel.textContent = exts.join(' ');
  filterPopover.style.display = 'none';
  await window.lumadesk.setExtensions(exts);
  statusText.textContent = `Filter: ${exts.join(' ')}`;
});

filterInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') btnFilterApply.click();
});

// Load saved extensions on start
(async () => {
  const exts = await window.lumadesk.getExtensions();
  if (exts) {
    filterInput.value = exts.join(' ');
    filterLabel.textContent = exts.join(' ');
  }
})();

// ---- File Context Menu ----
const fileContextMenu = $('#file-context-menu');
let contextMenuTargetPath = null;

function openFileContextMenu(x, y, filePath) {
  contextMenuTargetPath = filePath;
  // Show first to measure, then clamp into viewport
  fileContextMenu.style.display = 'block';
  fileContextMenu.style.left = '0px';
  fileContextMenu.style.top = '0px';
  const rect = fileContextMenu.getBoundingClientRect();
  const maxX = window.innerWidth - rect.width - 4;
  const maxY = window.innerHeight - rect.height - 4;
  fileContextMenu.style.left = Math.min(x, maxX) + 'px';
  fileContextMenu.style.top = Math.min(y, maxY) + 'px';
}

function closeFileContextMenu() {
  fileContextMenu.style.display = 'none';
  contextMenuTargetPath = null;
}

fileContextMenu.addEventListener('click', async (e) => {
  const item = e.target.closest('.context-menu-item');
  if (!item) return;
  const action = item.dataset.action;
  const target = contextMenuTargetPath;
  closeFileContextMenu();
  if (action === 'show-in-folder' && target) {
    const result = await window.lumadesk.showInFolder(target);
    if (result?.error) statusText.textContent = `Error: ${result.error}`;
  }
});

document.addEventListener('click', (e) => {
  if (fileContextMenu.style.display !== 'none' && !fileContextMenu.contains(e.target)) {
    closeFileContextMenu();
  }
});

document.addEventListener('contextmenu', (e) => {
  if (!e.target.closest('.tree-file') && fileContextMenu.style.display !== 'none') {
    closeFileContextMenu();
  }
});

window.addEventListener('blur', closeFileContextMenu);
window.addEventListener('resize', closeFileContextMenu);

// ---- In-document Search ----
const searchBar = $('#search-bar');
const searchInput = $('#search-input');
const searchCount = $('#search-count');
const searchPrev = $('#search-prev');
const searchNext = $('#search-next');
const searchClose = $('#search-close');

let searchHits = [];
let searchIndex = -1;

function clearSearchHighlights() {
  const marks = previewEl.querySelectorAll('mark.search-hit');
  marks.forEach((m) => {
    const parent = m.parentNode;
    while (m.firstChild) parent.insertBefore(m.firstChild, m);
    parent.removeChild(m);
    parent.normalize();
  });
  searchHits = [];
  searchIndex = -1;
}

function highlightSearch(query) {
  clearSearchHighlights();
  if (!query) {
    updateSearchUI();
    return;
  }

  const lower = query.toLowerCase();
  const walker = document.createTreeWalker(previewEl, NodeFilter.SHOW_TEXT, {
    acceptNode(node) {
      if (!node.nodeValue || !node.nodeValue.toLowerCase().includes(lower)) {
        return NodeFilter.FILTER_REJECT;
      }
      // Skip text inside mermaid SVGs, script, style
      let p = node.parentNode;
      while (p && p !== previewEl) {
        if (p.classList && p.classList.contains('mermaid-diagram')) return NodeFilter.FILTER_REJECT;
        const tag = p.nodeName;
        if (tag === 'SCRIPT' || tag === 'STYLE' || tag === 'MARK') return NodeFilter.FILTER_REJECT;
        p = p.parentNode;
      }
      return NodeFilter.FILTER_ACCEPT;
    },
  });

  const textNodes = [];
  let n;
  while ((n = walker.nextNode())) textNodes.push(n);

  for (const textNode of textNodes) {
    const text = textNode.nodeValue;
    const lowerText = text.toLowerCase();
    const frag = document.createDocumentFragment();
    let lastIdx = 0;
    let idx = lowerText.indexOf(lower);
    while (idx !== -1) {
      if (idx > lastIdx) {
        frag.appendChild(document.createTextNode(text.slice(lastIdx, idx)));
      }
      const mark = document.createElement('mark');
      mark.className = 'search-hit';
      mark.textContent = text.slice(idx, idx + query.length);
      frag.appendChild(mark);
      searchHits.push(mark);
      lastIdx = idx + query.length;
      idx = lowerText.indexOf(lower, lastIdx);
    }
    if (lastIdx < text.length) {
      frag.appendChild(document.createTextNode(text.slice(lastIdx)));
    }
    textNode.parentNode.replaceChild(frag, textNode);
  }

  if (searchHits.length > 0) {
    searchIndex = 0;
    focusCurrentHit();
  }
  updateSearchUI();
}

function focusCurrentHit() {
  searchHits.forEach((m) => m.classList.remove('current'));
  if (searchIndex >= 0 && searchIndex < searchHits.length) {
    const current = searchHits[searchIndex];
    current.classList.add('current');
    current.scrollIntoView({ block: 'center', behavior: 'smooth' });
  }
}

function updateSearchUI() {
  const total = searchHits.length;
  const pos = total === 0 ? 0 : searchIndex + 1;
  searchCount.textContent = `${pos} / ${total}`;
  searchPrev.disabled = total === 0;
  searchNext.disabled = total === 0;
}

function nextHit() {
  if (searchHits.length === 0) return;
  searchIndex = (searchIndex + 1) % searchHits.length;
  focusCurrentHit();
  updateSearchUI();
}

function prevHit() {
  if (searchHits.length === 0) return;
  searchIndex = (searchIndex - 1 + searchHits.length) % searchHits.length;
  focusCurrentHit();
  updateSearchUI();
}

function openSearch() {
  if (previewEl.style.display === 'none') return;
  searchBar.style.display = 'flex';
  searchInput.focus();
  searchInput.select();
  if (searchInput.value) highlightSearch(searchInput.value);
}

function closeSearch() {
  searchBar.style.display = 'none';
  clearSearchHighlights();
  updateSearchUI();
}

searchInput.addEventListener('input', () => highlightSearch(searchInput.value));
searchInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') {
    e.preventDefault();
    if (e.shiftKey) prevHit();
    else nextHit();
  } else if (e.key === 'Escape') {
    e.preventDefault();
    closeSearch();
  }
});
searchNext.addEventListener('click', () => { nextHit(); searchInput.focus(); });
searchPrev.addEventListener('click', () => { prevHit(); searchInput.focus(); });
searchClose.addEventListener('click', closeSearch);

// ---- Folder Search (file names + file contents) ----
const folderSearchInput = $('#folder-search-input');
const folderSearchClear = $('#folder-search-clear');
const folderSearchResults = $('#folder-search-results');

let folderSearchTimer = null;
// Query to re-apply as an in-document highlight after a content result opens.
let pendingPreviewHighlight = '';
// Search mode: match file names OR file contents (one at a time).
let searchMode = 'content';
// Keyboard navigation over the flat list of clickable result rows.
let searchNavEls = [];
let searchNavIndex = -1;

function escapeHtml(s) {
  return String(s).replace(/[&<>"]/g, (c) => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]
  ));
}

// Wrap each case-insensitive match of `query` in <mark> for the results list.
function highlightSnippet(text, query) {
  const lower = text.toLowerCase();
  const q = query.toLowerCase();
  let out = '';
  let last = 0;
  let idx = lower.indexOf(q);
  while (idx !== -1) {
    out += escapeHtml(text.slice(last, idx));
    out += '<mark>' + escapeHtml(text.slice(idx, idx + query.length)) + '</mark>';
    last = idx + query.length;
    idx = lower.indexOf(q, last);
  }
  out += escapeHtml(text.slice(last));
  return out;
}

function showFolderSearchResults(show) {
  folderSearchResults.style.display = show ? 'block' : 'none';
  fileTreeEl.style.display = show ? 'none' : '';
}

function clearFolderSearch() {
  folderSearchInput.value = '';
  folderSearchClear.style.display = 'none';
  folderSearchResults.innerHTML = '';
  searchNavEls = [];
  searchNavIndex = -1;
  showFolderSearchResults(false);
}

async function runFolderSearch(query) {
  const q = query.trim();
  folderSearchClear.style.display = q ? 'flex' : 'none';

  if (!q) {
    showFolderSearchResults(false);
    folderSearchResults.innerHTML = '';
    return;
  }

  const { results, truncated } = await window.lumadesk.searchFolder({
    query: q,
    matchName: searchMode === 'name',
    matchContent: searchMode === 'content',
  });
  renderFolderSearchResults(results || [], q, truncated);
}

function renderFolderSearchResults(results, query, truncated) {
  showFolderSearchResults(true);
  searchNavIndex = -1;

  if (results.length === 0) {
    folderSearchResults.innerHTML =
      `<div class="search-empty">No ${searchMode === 'name' ? 'file names' : 'content'} match &ldquo;${escapeHtml(query)}&rdquo;</div>`;
    searchNavEls = [];
    return;
  }

  const fileCount = results.length;
  let summary;
  if (searchMode === 'name') {
    summary = `${fileCount} file${fileCount === 1 ? '' : 's'}`;
  } else {
    const matchCount = results.reduce((n, r) => n + r.totalContentMatches, 0);
    summary = `${matchCount} match${matchCount === 1 ? '' : 'es'} in ${fileCount} file${fileCount === 1 ? '' : 's'}`;
  }
  if (truncated) summary += ' (first 500)';

  let html = `<div class="search-summary">${summary}</div>`;

  for (const r of results) {
    const active = r.path === currentFilePath ? ' active' : '';
    // In name mode the whole file row is the clickable nav target; in content
    // mode each snippet is its own target (jumps to that line's term).
    const fileNav = searchMode === 'name' ? ' data-nav="1"' : '';
    html += `<div class="search-result${active}">`;
    html += `<div class="search-result-file" data-path="${escapeHtml(r.path)}"${fileNav}>`;
    html += `<svg class="tree-file-icon" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>`;
    html += `<span class="search-result-name">${(r.nameMatch && searchMode === 'name') ? highlightSnippet(r.name, query) : escapeHtml(r.name)}</span>`;
    if (searchMode === 'content' && r.totalContentMatches > 0) {
      html += `<span class="search-result-count">${r.totalContentMatches}</span>`;
    }
    html += `</div>`;
    if (searchMode === 'content') {
      for (const sn of r.snippets) {
        html += `<div class="search-result-snippet" data-path="${escapeHtml(r.path)}" data-query="${escapeHtml(query)}" data-nav="1"><span class="snippet-line">${sn.line}</span><span class="snippet-text">${highlightSnippet(sn.text, query)}</span></div>`;
      }
      if (r.totalContentMatches > r.snippets.length) {
        html += `<div class="search-result-more">+${r.totalContentMatches - r.snippets.length} more in this file</div>`;
      }
    }
    html += `</div>`;
  }

  folderSearchResults.innerHTML = html;
  searchNavEls = Array.from(folderSearchResults.querySelectorAll('[data-nav="1"]'));
}

function setSearchNav(index) {
  if (!searchNavEls.length) return;
  searchNavIndex = (index + searchNavEls.length) % searchNavEls.length;
  searchNavEls.forEach((el, i) => el.classList.toggle('nav-active', i === searchNavIndex));
  searchNavEls[searchNavIndex].scrollIntoView({ block: 'nearest' });
}

function openNavEl(el) {
  if (!el) return;
  const filePath = el.dataset.path;
  const query = el.dataset.query || folderSearchInput.value.trim();
  if (filePath) openSearchResult(filePath, query);
}

// Open a result: load the file, then (for content matches) re-highlight the
// query in the preview via the in-document search bar and jump to the first hit.
async function openSearchResult(filePath, query) {
  pendingPreviewHighlight = (searchMode === 'content') ? (query || '') : '';
  await selectFile(filePath);
}

folderSearchResults.addEventListener('click', (e) => {
  const el = e.target.closest('[data-path]');
  if (el) openNavEl(el);
});

folderSearchInput.addEventListener('input', () => {
  clearTimeout(folderSearchTimer);
  folderSearchTimer = setTimeout(() => runFolderSearch(folderSearchInput.value), 180);
});

folderSearchInput.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    e.preventDefault();
    clearFolderSearch();
    folderSearchInput.blur();
  } else if (e.key === 'ArrowDown') {
    e.preventDefault();
    setSearchNav(searchNavIndex + 1);
  } else if (e.key === 'ArrowUp') {
    e.preventDefault();
    setSearchNav(searchNavIndex - 1);
  } else if (e.key === 'Enter') {
    e.preventDefault();
    openNavEl(searchNavIndex >= 0 ? searchNavEls[searchNavIndex] : searchNavEls[0]);
  } else if ((e.ctrlKey || e.metaKey) && (e.key === 'ArrowLeft' || e.key === 'ArrowRight')) {
    // Ctrl+←/→ flips between Name and Content modes.
    e.preventDefault();
    setSearchMode(searchMode === 'name' ? 'content' : 'name');
  }
});

folderSearchClear.addEventListener('click', () => {
  clearFolderSearch();
  folderSearchInput.focus();
});

// Single-select scope: Name OR Content.
function setSearchMode(mode) {
  if (mode !== 'name' && mode !== 'content') return;
  searchMode = mode;
  document.querySelectorAll('#search-scope .scope-btn').forEach((b) =>
    b.classList.toggle('active', b.dataset.scope === mode));
  folderSearchInput.placeholder = mode === 'name' ? 'Search file names…' : 'Search in file contents…';
  runFolderSearch(folderSearchInput.value);
  folderSearchInput.focus();
}

$('#search-scope').addEventListener('click', (e) => {
  const btn = e.target.closest('.scope-btn');
  if (btn) setSearchMode(btn.dataset.scope);
});

// Apply the default mode's placeholder/active state on load.
setSearchMode(searchMode);

// ---- PPT Export ----
const btnExportPpt = $('#btn-export-ppt');
const btnPptSettings = $('#btn-ppt-settings');
const pptSettingsPopover = $('#ppt-settings-popover');

const PPT_SETTINGS_KEY = 'lumadesk.pptExportSettings';

function loadPptSettings() {
  try {
    const raw = localStorage.getItem(PPT_SETTINGS_KEY);
    if (!raw) return { theme: 'light' };
    const parsed = JSON.parse(raw);
    return { theme: ['light', 'dark', 'branded'].includes(parsed.theme) ? parsed.theme : 'light' };
  } catch {
    return { theme: 'light' };
  }
}

function savePptSettings(theme) {
  localStorage.setItem(PPT_SETTINGS_KEY, JSON.stringify({ theme }));
}

function refreshPptThemeButtons() {
  const current = loadPptSettings().theme;
  pptSettingsPopover.querySelectorAll('.ppt-theme-btn').forEach((b) => {
    b.classList.toggle('active', b.dataset.theme === current);
  });
}

btnPptSettings.addEventListener('click', (e) => {
  e.stopPropagation();
  pptSettingsPopover.style.display = pptSettingsPopover.style.display === 'none' ? 'block' : 'none';
  if (pptSettingsPopover.style.display === 'block') refreshPptThemeButtons();
});

document.addEventListener('click', (e) => {
  if (!pptSettingsPopover.contains(e.target) && e.target !== btnPptSettings && !btnPptSettings.contains(e.target)) {
    pptSettingsPopover.style.display = 'none';
  }
});

pptSettingsPopover.querySelectorAll('.ppt-theme-btn').forEach((btn) => {
  btn.addEventListener('click', () => {
    savePptSettings(btn.dataset.theme);
    refreshPptThemeButtons();
    statusText.textContent = `PPT theme: ${btn.dataset.theme}`;
  });
});

btnExportPpt.addEventListener('click', async () => {
  if (!currentContent) return;

  statusText.textContent = 'Rendering diagrams...';
  btnExportPpt.disabled = true;

  const mermaidImages = await captureMermaidImages(currentContent);
  const { theme } = loadPptSettings();

  statusText.textContent = `Building .pptx (${theme})...`;

  try {
    const buffer = await window.LumaPpt.exportToPptx({
      title: currentTitle,
      content: currentContent,
      theme,
      mermaidImages,
    });
    const result = await window.lumadesk.savePptx({
      title: currentTitle,
      buffer,
    });
    if (result.success) statusText.textContent = `Exported: ${result.path}`;
    else if (result.error) statusText.textContent = `Export failed: ${result.error}`;
    else statusText.textContent = 'Export cancelled';
  } catch (err) {
    statusText.textContent = `PPT failed: ${err.message || err}`;
  }

  btnExportPpt.disabled = false;
});

// ---- Presentation Mode ----
const btnPresent = $('#btn-present');
const presentOverlay = $('#present-overlay');
const presentSlideEl = $('#present-slide');
const presentCounter = $('#present-counter');
const presentProgress = $('#present-progress');

let presentSlides = [];
let presentIndex = 0;

function buildSlideMarkdown(slide) {
  // Reconstruct a slide as markdown for the existing renderMarkdown IPC.
  const parts = [];
  if (slide.title) parts.push(`# ${slide.title}`);
  if (slide.subtitle) parts.push(`## ${slide.subtitle}`);
  if (slide.content) parts.push(slide.content);
  return parts.join('\n\n');
}

async function renderCurrentSlide() {
  if (!presentSlides.length) return;
  const slide = presentSlides[presentIndex];
  const md = buildSlideMarkdown(slide);
  const html = await window.lumadesk.renderMarkdown(md);
  // Wrap in a <section> so the deck's own CSS (section, section.lead,
  // section.divider, ...) applies just like Marp.
  const cls = slide.sectionClass ? ` class="${slide.sectionClass}"` : '';
  presentSlideEl.innerHTML = `<section${cls}>${html}</section>`;
  presentCounter.textContent = `${presentIndex + 1} / ${presentSlides.length}`;
  const pct = ((presentIndex + 1) / presentSlides.length) * 100;
  presentProgress.style.width = pct + '%';

  // Render mermaid blocks in the slide
  if (mermaidReady) {
    const codeBlocks = presentSlideEl.querySelectorAll('code.language-mermaid, code[class*="language-mermaid"]');
    for (const code of codeBlocks) {
      const pre = code.parentElement;
      const source = code.textContent;
      try {
        const id = 'pres-' + Math.random().toString(36).slice(2, 10);
        const { svg } = await mermaidModule.default.render(id, source);
        const container = document.createElement('div');
        container.className = 'mermaid-diagram';
        container.innerHTML = svg;
        pre.replaceWith(container);
        hardenDiagramColors(container);
      } catch (e) {
        console.warn('Mermaid render error (present):', e);
        showDiagramError(pre, source, e);
      }
    }
  }

  presentSlideEl.scrollTop = 0;
}

// A slide's injected deck <style> may include broad rules like
// `section span { color: ... !important }`. Mermaid renders flowchart labels
// as HTML spans inside the SVG, so those rules can override the diagram's own
// (non-important) label colors and make text unreadable (e.g. white-on-dark
// nodes turning dark-on-dark). Re-assert Mermaid's own colors with !important
// so they win the cascade again.
function hardenDiagramColors(container) {
  // 1. Bump declarations in the SVG's internal <style> to !important.
  container.querySelectorAll('svg style').forEach((styleEl) => {
    styleEl.textContent = styleEl.textContent.replace(
      /([\w-]+)\s*:\s*([^;{}!]+?)\s*(!important)?\s*;/g,
      (m, prop, val, imp) => (imp ? m : `${prop}:${val} !important;`)
    );
  });
  // 2. Re-apply any inline label colors as !important so they beat the deck.
  container.querySelectorAll('svg [style]').forEach((el) => {
    const color = el.style.color;
    const bg = el.style.backgroundColor;
    if (color) el.style.setProperty('color', color, 'important');
    if (bg) el.style.setProperty('background-color', bg, 'important');
  });
}

function presentNext() {
  if (presentIndex < presentSlides.length - 1) {
    presentIndex++;
    renderCurrentSlide();
  }
}

function presentPrev() {
  if (presentIndex > 0) {
    presentIndex--;
    renderCurrentSlide();
  }
}

function presentFirst() {
  presentIndex = 0;
  renderCurrentSlide();
}

function presentLast() {
  presentIndex = presentSlides.length - 1;
  renderCurrentSlide();
}

function injectDeckStyle(css) {
  let el = document.getElementById('present-deck-style');
  if (!css) {
    if (el) el.remove();
    return;
  }
  if (!el) {
    el = document.createElement('style');
    el.id = 'present-deck-style';
    document.head.appendChild(el);
  }
  el.textContent = css;
}

async function openPresent() {
  if (!currentContent) return;
  const deck = window.LumaSlides.parseDeck(currentContent);
  presentSlides = deck.slides;
  if (!presentSlides.length) {
    statusText.textContent = 'No slides parsed';
    return;
  }
  // Apply the deck's own <style> so the presentation matches its design.
  injectDeckStyle(deck.style);
  presentIndex = 0;
  presentOverlay.style.display = 'flex';
  await renderCurrentSlide();
}

function closePresent() {
  presentOverlay.style.display = 'none';
  injectDeckStyle(''); // remove deck CSS so it doesn't bleed into the app
  if (document.fullscreenElement) document.exitFullscreen();
}

function toggleFullscreen() {
  if (document.fullscreenElement) document.exitFullscreen();
  else presentOverlay.requestFullscreen?.();
}

btnPresent.addEventListener('click', openPresent);
$('#present-prev').addEventListener('click', presentPrev);
$('#present-next').addEventListener('click', presentNext);
$('#present-fullscreen').addEventListener('click', toggleFullscreen);
$('#present-close').addEventListener('click', closePresent);

// Scroll-wheel navigation: scroll within the slide first; only switch
// slides when the slide is at the edge in the wheel direction. Debounced
// so trackpad inertia doesn't blast through multiple slides at the edge.
let wheelLockUntil = 0;
presentOverlay.addEventListener('wheel', (e) => {
  if (presentOverlay.style.display === 'none') return;
  if (Math.abs(e.deltaY) < 10) return;

  const s = presentSlideEl;
  const atTop = s.scrollTop <= 0;
  const atBottom = s.scrollTop + s.clientHeight >= s.scrollHeight - 1;
  const goingDown = e.deltaY > 0;

  // If the slide can still scroll in this direction, let the browser do it.
  if (goingDown && !atBottom) return;
  if (!goingDown && !atTop) return;

  // At the edge — navigate, with debounce.
  e.preventDefault();
  const now = performance.now();
  if (now < wheelLockUntil) return;
  wheelLockUntil = now + 350;
  if (goingDown) presentNext();
  else presentPrev();
}, { passive: false });

// ---- Document Editor (split edit mode) ----
const btnEdit = $('#btn-edit');
const previewContainer = $('#preview-container');
const previewPane = $('#preview-pane');
const editorDivider = $('#editor-divider');
const editorPane = $('#editor-pane');
const editorToolbar = $('#editor-toolbar');
const editorTextarea = $('#editor-textarea');
const editorWysiwyg = $('#editor-wysiwyg');
const editorSaveBtn = $('#editor-save');
const editorDirtyDot = $('#editor-dirty');
const editorModeSwitch = $('#editor-mode-switch');

const editorState = { active: false, mode: 'text', dirty: false, fm: '', wysiwygRange: null };
let editorSyncTimer = null;

// Clicking the mode switch blurs the contenteditable and clears its selection,
// so keep the last caret/selection that lived inside it for the mode switch.
document.addEventListener('selectionchange', () => {
  if (!editorState.active || editorState.mode !== 'wysiwyg') return;
  const sel = window.getSelection();
  if (sel && sel.rangeCount && editorWysiwyg.contains(sel.anchorNode)) {
    editorState.wysiwygRange = sel.getRangeAt(0).cloneRange();
  }
});

function setEditorDirty(d) {
  editorState.dirty = d;
  editorDirtyDot.style.display = d ? 'inline' : 'none';
}

// Split a leading YAML frontmatter block off so WYSIWYG mode can preserve it
// verbatim (marked strips it from the rendered body).
function splitFrontmatter(md) {
  const m = md.match(/^﻿?---[ \t]*\r?\n[\s\S]*?\r?\n---[ \t]*(\r?\n|$)/);
  if (m) return { fm: m[0].replace(/\s+$/, ''), body: md.slice(m[0].length) };
  return { fm: '', body: md };
}

// Current editor contents as Markdown, from whichever mode is active.
function getEditorMarkdown() {
  if (editorState.mode === 'wysiwyg') {
    let md = htmlToMarkdown(editorWysiwyg);
    if (editorState.fm) md = editorState.fm + '\n' + md;
    return md;
  }
  return editorTextarea.value;
}

// Re-render the top (read-only) preview from the live editor content.
async function renderTopPreview(md) {
  const html = await window.lumadesk.renderMarkdown(md);
  previewEl.innerHTML = html;
  await renderSvgBlocks();
  await renderMermaidBlocks();
  if (searchBar.style.display !== 'none' && searchInput.value) {
    highlightSearch(searchInput.value);
  }
}

function scheduleEditorSync() {
  clearTimeout(editorSyncTimer);
  editorSyncTimer = setTimeout(() => renderTopPreview(getEditorMarkdown()), 300);
}

function markEditorDirty() {
  if (!editorState.dirty) setEditorDirty(true);
}

// Render Markdown into the contenteditable WYSIWYG surface, turning mermaid/svg
// code blocks into rendered (non-editable) diagrams so they're easy to read.
async function renderWysiwygFromMarkdown(body) {
  const html = await window.lumadesk.renderMarkdown(body);
  editorWysiwyg.innerHTML = html;
  await renderEditorDiagrams(editorWysiwyg);
}

async function renderEditorDiagrams(scope) {
  if (mermaidReady) {
    const blocks = scope.querySelectorAll('code.language-mermaid, code[class*="language-mermaid"], code[class*="language- mermaid"]');
    for (const code of blocks) {
      const pre = code.parentElement;
      const source = code.textContent;
      try {
        const id = 'edm-' + Math.random().toString(36).slice(2, 10);
        const { svg } = await mermaidModule.default.render(id, source);
        const c = document.createElement('div');
        c.className = 'mermaid-diagram';
        c.setAttribute('contenteditable', 'false');
        c.dataset.src = source.replace(/\n+$/, '');
        c.innerHTML = svg;
        pre.replaceWith(c);
      } catch (e) {
        console.warn('Editor mermaid render error:', e);
      }
    }
  }
  const svgBlocks = scope.querySelectorAll('code.language-svg, code[class*="language-svg"]');
  for (const code of svgBlocks) {
    const pre = code.parentElement;
    const source = code.textContent.trim();
    const c = document.createElement('div');
    c.className = 'svg-diagram';
    c.setAttribute('contenteditable', 'false');
    c.dataset.src = source;
    c.innerHTML = source;
    pre.replaceWith(c);
  }
}

// ---- HTML → Markdown (for WYSIWYG mode) ----
function htmlToMarkdown(root) {
  const parts = [];
  root.childNodes.forEach((n) => {
    const block = serializeBlock(n);
    if (block.trim() !== '') parts.push(block);
  });
  return parts.join('\n\n').replace(/\n{3,}/g, '\n\n').trim() + '\n';
}

function serializeInline(node) {
  let out = '';
  node.childNodes.forEach((n) => {
    if (n.nodeType === 3) { out += n.nodeValue; return; }
    if (n.nodeType !== 1) return;
    const tag = n.nodeName.toLowerCase();
    const inner = serializeInline(n);
    switch (tag) {
      case 'strong': case 'b': out += inner.trim() ? '**' + inner + '**' : ''; break;
      case 'em': case 'i': out += inner.trim() ? '*' + inner + '*' : ''; break;
      case 'del': case 's': case 'strike': out += '~~' + inner + '~~'; break;
      case 'code': out += '`' + n.textContent + '`'; break;
      case 'a': out += '[' + inner + '](' + (n.getAttribute('href') || '') + ')'; break;
      case 'br': out += '  \n'; break;
      case 'img': out += '![' + (n.getAttribute('alt') || '') + '](' + (n.getAttribute('src') || '') + ')'; break;
      default: out += inner;
    }
  });
  return out;
}

function serializeBlock(node) {
  if (node.nodeType === 3) return node.nodeValue.trim();
  if (node.nodeType !== 1) return '';
  const el = node;
  const tag = el.nodeName.toLowerCase();

  // Rendered diagrams carry their original source.
  if (el.classList && el.classList.contains('mermaid-diagram') && el.dataset.src != null) {
    return '```mermaid\n' + el.dataset.src + '\n```';
  }
  if (el.classList && el.classList.contains('svg-diagram') && el.dataset.src != null) {
    return '```svg\n' + el.dataset.src + '\n```';
  }

  switch (tag) {
    case 'h1': return '# ' + serializeInline(el);
    case 'h2': return '## ' + serializeInline(el);
    case 'h3': return '### ' + serializeInline(el);
    case 'h4': return '#### ' + serializeInline(el);
    case 'h5': return '##### ' + serializeInline(el);
    case 'h6': return '###### ' + serializeInline(el);
    case 'p': return serializeInline(el);
    case 'hr': return '---';
    case 'blockquote': {
      const inner = Array.from(el.childNodes).map((c) => serializeBlock(c)).filter(Boolean).join('\n\n');
      return inner.split('\n').map((l) => (l ? '> ' + l : '>')).join('\n');
    }
    case 'pre': {
      const code = el.querySelector('code');
      const text = (code ? code.textContent : el.textContent).replace(/\n$/, '');
      const m = code && code.className.match(/language-([\w-]+)/);
      return '```' + (m ? m[1] : '') + '\n' + text + '\n```';
    }
    case 'ul': return serializeList(el, false, '');
    case 'ol': return serializeList(el, true, '');
    case 'table': return serializeTable(el);
    case 'div': case 'section': case 'article':
      return Array.from(el.childNodes).map((c) => serializeBlock(c)).filter(Boolean).join('\n\n');
    default: return serializeInline(el);
  }
}

function serializeList(el, ordered, indent) {
  const lines = [];
  let i = 1;
  Array.from(el.children).forEach((li) => {
    if (li.nodeName.toLowerCase() !== 'li') return;
    const marker = ordered ? (i++) + '. ' : '- ';
    let inline = '';
    li.childNodes.forEach((n) => {
      if (n.nodeType === 1 && (n.nodeName === 'UL' || n.nodeName === 'OL')) return;
      if (n.nodeType === 3) inline += n.nodeValue;
      else if (n.nodeType === 1) inline += serializeInline(n);
    });
    lines.push(indent + marker + inline.trim());
    Array.from(li.children).forEach((ch) => {
      const t = ch.nodeName.toLowerCase();
      if (t === 'ul') lines.push(serializeList(ch, false, indent + '  '));
      if (t === 'ol') lines.push(serializeList(ch, true, indent + '  '));
    });
  });
  return lines.join('\n');
}

function serializeTable(table) {
  let head = Array.from(table.querySelectorAll('thead th, thead td')).map((c) => serializeInline(c).trim());
  let bodyRows = Array.from(table.querySelectorAll('tbody tr'));
  if (head.length === 0) {
    const first = table.querySelector('tr');
    if (first) {
      head = Array.from(first.children).map((c) => serializeInline(c).trim());
      bodyRows = Array.from(table.querySelectorAll('tr')).slice(1);
    }
  }
  if (head.length === 0) return '';
  const rows = ['| ' + head.join(' | ') + ' |', '| ' + head.map(() => '---').join(' | ') + ' |'];
  bodyRows.forEach((tr) => {
    const cells = Array.from(tr.children).map((c) => serializeInline(c).trim());
    rows.push('| ' + cells.join(' | ') + ' |');
  });
  return rows.join('\n');
}

// ---- Edit mode enter/exit ----
function enterEditMode() {
  if (!currentFilePath) { statusText.textContent = 'Open a file to edit'; return; }
  editorState.active = true;
  editorState.mode = 'text';
  editorState.fm = '';
  editorTextarea.value = currentContent;
  editorTextarea.style.display = 'block';
  editorWysiwyg.style.display = 'none';
  editorModeSwitch.querySelectorAll('.mode-opt').forEach((b) => b.classList.toggle('active', b.dataset.mode === 'text'));

  previewContainer.classList.add('editing');
  editorDivider.style.display = 'block';
  editorPane.style.display = 'flex';
  if (!editorPane.style.height) {
    editorPane.style.height = Math.round(previewContainer.clientHeight * 0.42) + 'px';
  }
  btnEdit.classList.add('active');
  setEditorDirty(false);
  editorTextarea.focus();
  statusText.textContent = 'Editing';
}

function exitEditMode() {
  if (editorState.dirty && !confirm('Discard unsaved changes?')) return;
  editorState.active = false;
  previewContainer.classList.remove('editing');
  editorDivider.style.display = 'none';
  editorPane.style.display = 'none';
  btnEdit.classList.remove('active');
  setEditorDirty(false);
  // Restore the preview to the last saved content.
  renderTopPreview(currentContent);
  statusText.textContent = 'Ready';
}

async function setEditorMode(mode) {
  if (mode === editorState.mode) return;
  // Capture the caret/selection in the pane we're leaving as word-character
  // counts, plus a scroll fraction as a fallback if we can't locate it.
  const anchor = captureEditorAnchor();
  const fromEl = editorState.mode === 'wysiwyg' ? editorWysiwyg : editorTextarea;
  const fromMax = fromEl.scrollHeight - fromEl.clientHeight;
  const scrollFrac = fromMax > 0 ? fromEl.scrollTop / fromMax : 0;
  if (mode === 'wysiwyg') {
    const { fm, body } = splitFrontmatter(editorTextarea.value);
    editorState.fm = fm;
    await renderWysiwygFromMarkdown(body);
    editorTextarea.style.display = 'none';
    editorWysiwyg.style.display = 'block';
    editorState.mode = 'wysiwyg';
    editorWysiwyg.focus({ preventScroll: true });
    if (!restoreWysiwygAnchor(anchor)) setScrollFrac(editorWysiwyg, scrollFrac);
  } else {
    editorTextarea.value = getEditorMarkdown();
    editorState.fm = '';
    editorWysiwyg.style.display = 'none';
    editorTextarea.style.display = 'block';
    editorState.mode = 'text';
    editorTextarea.focus({ preventScroll: true });
    if (!restoreTextareaAnchor(anchor)) setScrollFrac(editorTextarea, scrollFrac);
  }
  editorModeSwitch.querySelectorAll('.mode-opt').forEach((b) => b.classList.toggle('active', b.dataset.mode === mode));
}

function setScrollFrac(el, frac) {
  const max = el.scrollHeight - el.clientHeight;
  el.scrollTop = max > 0 ? Math.round(frac * max) : 0;
}

// --- Caret mapping across the two representations -----------------------------
// The caret is anchored by the TEXT around it, not by an absolute position: we
// take a small window of word characters (letters/digits) just before and after
// the caret and search for that same window in the target view. Only word chars
// are compared, so markdown syntax/whitespace, frontmatter and diagrams between
// the two views don't throw the match off. Matching is local, so it doesn't
// drift the way a running character count would.
const ANCHOR_WIN = 10;     // word chars of context captured on each side of the caret
const MAX_AMBIGUOUS = 8;   // more identical windows than this → don't trust the match
const AMBIGUOUS_TOL = 40;  // word chars; best match must be this close to the estimate
const isWordChar = (ch) => /[\p{L}\p{N}]/u.test(ch);

// Reduce a string to its word chars, keeping each one's original index.
function wordStream(str) {
  let chars = '';
  const idx = [];
  for (let i = 0; i < str.length; i++) {
    if (isWordChar(str[i])) { chars += str[i]; idx.push(i); }
  }
  return { chars, idx };
}

function countWordChars(str) {
  let n = 0;
  for (let i = 0; i < str.length; i++) if (isWordChar(str[i])) n++;
  return n;
}

// Diagrams render as SVG in the visual view but as fenced source in markdown, so
// their characters never match across views. We exclude them from both word
// streams and treat each as an opaque boundary the caret can't fall inside.
const DIAGRAM_SEL = '.mermaid-diagram, .svg-diagram';
const inDiagram = (node) => !!(node.parentElement && node.parentElement.closest(DIAGRAM_SEL));
const diagramFilter = {
  acceptNode: (n) => (inDiagram(n) ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT),
};

// Markdown word stream with fenced mermaid/svg blocks masked out. `idx` maps each
// kept word char back to its offset in the original string.
function markdownWordStream(str) {
  const excluded = [];
  // Frontmatter lives in the textarea but not the visual view, so mask it to
  // keep the two streams aligned (and the position estimate honest).
  const fmLen = str.length - splitFrontmatter(str).body.length;
  if (fmLen > 0) excluded.push([0, fmLen]);
  const re = /```[ \t]*(?:mermaid|svg)\b[\s\S]*?```/gi;
  let m;
  while ((m = re.exec(str))) excluded.push([m.index, m.index + m[0].length]);
  let chars = '';
  const idx = [];
  let ei = 0;
  for (let i = 0; i < str.length; i++) {
    while (ei < excluded.length && i >= excluded[ei][1]) ei++;
    if (ei < excluded.length && i >= excluded[ei][0]) continue; // inside a diagram
    if (isWordChar(str[i])) { chars += str[i]; idx.push(i); }
  }
  return { chars, idx };
}

// Visual word stream (diagrams removed).
function visualWordChars() {
  const clone = editorWysiwyg.cloneNode(true);
  clone.querySelectorAll(DIAGRAM_SEL).forEach((e) => e.remove());
  return wordStream(clone.textContent).chars;
}

// Word chars before a WYSIWYG selection boundary, diagrams excluded.
function visualCountBefore(node, offset) {
  const r = document.createRange();
  r.selectNodeContents(editorWysiwyg);
  r.setEnd(node, offset);
  const frag = r.cloneContents();
  frag.querySelectorAll(DIAGRAM_SEL).forEach((e) => e.remove());
  return countWordChars(frag.textContent);
}

// The context window (before/after word chars, plus a position fraction for
// tie-breaking) around a caret sitting after `k` word chars of `chars`.
function windowAround(chars, k) {
  return {
    before: chars.slice(Math.max(0, k - ANCHOR_WIN), k),
    after: chars.slice(k, k + ANCHOR_WIN),
    frac: chars.length ? k / chars.length : 0,
  };
}

// Number of entries in a sorted `idx` array that are less than `off`.
function countBefore(idx, off) {
  let lo = 0, hi = idx.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (idx[mid] < off) lo = mid + 1; else hi = mid;
  }
  return lo;
}

// Capture { start, end } context windows for the current selection, or null.
function captureEditorAnchor() {
  if (editorState.mode === 'text') {
    const { chars, idx } = markdownWordStream(editorTextarea.value);
    return {
      start: windowAround(chars, countBefore(idx, editorTextarea.selectionStart)),
      end: windowAround(chars, countBefore(idx, editorTextarea.selectionEnd)),
    };
  }
  const range = editorState.wysiwygRange;
  if (!range || !editorWysiwyg.contains(range.startContainer)) return null;
  const chars = visualWordChars();
  return {
    start: windowAround(chars, visualCountBefore(range.startContainer, range.startOffset)),
    end: windowAround(chars, visualCountBefore(range.endContainer, range.endOffset)),
  };
}

// All positions (word-char count before the seam) where `key` occurs in `chars`.
function allSplits(chars, key, beforeLen) {
  const splits = [];
  for (let at = chars.indexOf(key); at !== -1; at = chars.indexOf(key, at + 1)) {
    splits.push(at + beforeLen);
  }
  return splits;
}

// Find where a captured window lands in `chars`: the word-char count before the
// caret in the target. Returns null only if there's nothing to go on.
//
// A single match is trusted outright. When the window repeats, we fall back to
// the caret's proportional position (`expected`): we accept the nearest copy
// only if it's close to that estimate and there aren't too many copies —
// otherwise the pattern is too confusing to trust and we use the estimate.
function locateWindow(chars, win) {
  if (!chars.length) return null;
  const expected = win.frac * chars.length;
  const nearest = (arr) => arr.reduce((b, s) => (Math.abs(s - expected) < Math.abs(b - expected) ? s : b));

  const key = win.before + win.after;
  if (!key) return Math.round(expected);

  const splits = allSplits(chars, key, win.before.length);
  if (splits.length === 1) return splits[0];
  if (splits.length > 1) {
    const best = nearest(splits);
    const trustworthy = splits.length <= MAX_AMBIGUOUS && Math.abs(best - expected) <= AMBIGUOUS_TOL;
    return trustworthy ? best : Math.round(expected);
  }

  // Full window not found (e.g. a link URL sits between the halves) — try each
  // half, again preferring a copy near the estimate; else use the estimate.
  const half = [];
  if (win.before) {
    for (let at = chars.indexOf(win.before); at !== -1; at = chars.indexOf(win.before, at + 1)) half.push(at + win.before.length);
  }
  if (win.after) {
    for (let at = chars.indexOf(win.after); at !== -1; at = chars.indexOf(win.after, at + 1)) half.push(at);
  }
  if (!half.length) return Math.round(expected);
  const best = nearest(half);
  return Math.abs(best - expected) <= AMBIGUOUS_TOL ? best : Math.round(expected);
}

// Offsets are snapped to word boundaries so punctuation/whitespace in the gap
// isn't swept in: a selection START sits just before its next word char, an END
// just after its previous one. `bias` is 'before' (start) or 'after' (end).
// Offset just after the `p`-th word char.
function offsetAfterWord(str, idx, p) {
  if (p <= 0) return 0;
  if (p >= idx.length) return str.length;
  return idx[p - 1] + 1;
}
// Offset just before the (p+1)-th word char (i.e. skip the gap forward).
function offsetBeforeWord(str, idx, p) {
  if (p >= idx.length) return str.length;
  return idx[p];
}

function restoreTextareaAnchor(anchor) {
  if (!anchor) return false;
  const v = editorTextarea.value;
  const { chars, idx } = markdownWordStream(v);
  let ps = locateWindow(chars, anchor.start);
  let pe = locateWindow(chars, anchor.end);
  if (ps === null && pe === null) return false;
  ps = ps ?? pe;
  pe = pe ?? ps;
  let s, e;
  if (ps === pe) {
    s = e = offsetAfterWord(v, idx, ps);
  } else {
    s = offsetBeforeWord(v, idx, Math.min(ps, pe));
    e = offsetAfterWord(v, idx, Math.max(ps, pe));
  }
  editorTextarea.setSelectionRange(s, e);
  const line = v.slice(0, s).split('\n').length - 1;
  const lh = parseFloat(getComputedStyle(editorTextarea).lineHeight) || 20;
  editorTextarea.scrollTop = Math.max(0, line * lh - editorTextarea.clientHeight / 2);
  return true;
}

// Resolve a word-char count to a [textNode, offset] position in the WYSIWYG DOM,
// skipping diagram subtrees so it stays in step with the captured stream.
// bias 'after' → just after the p-th word char; 'before' → just before the
// (p+1)-th word char.
function wysiwygPos(p, bias) {
  if (p <= 0 && bias === 'after') return [editorWysiwyg, 0];
  const walker = document.createTreeWalker(editorWysiwyg, NodeFilter.SHOW_TEXT, diagramFilter);
  let acc = 0, node, last = null;
  while ((node = walker.nextNode())) {
    const val = node.nodeValue;
    for (let i = 0; i < val.length; i++) {
      if (isWordChar(val[i])) {
        if (bias === 'before' && acc === p) return [node, i];
        if (++acc === p && bias === 'after') return [node, i + 1];
      }
    }
    last = node;
  }
  return last ? [last, last.nodeValue.length] : [editorWysiwyg, 0];
}

function restoreWysiwygAnchor(anchor) {
  if (!anchor) return false;
  const chars = visualWordChars();
  let ps = locateWindow(chars, anchor.start);
  let pe = locateWindow(chars, anchor.end);
  if (ps === null && pe === null) return false;
  ps = ps ?? pe;
  pe = pe ?? ps;
  if (ps > pe) [ps, pe] = [pe, ps];
  const collapsed = ps === pe;
  const [sn, so] = collapsed ? wysiwygPos(ps, 'after') : wysiwygPos(ps, 'before');
  const [en, eo] = wysiwygPos(pe, 'after');
  const range = document.createRange();
  range.setStart(sn, so);
  range.setEnd(en, eo);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
  // Scroll the caret roughly to the middle of the editable pane.
  let rect = range.getBoundingClientRect();
  if (!rect.height && sn.parentElement) rect = sn.parentElement.getBoundingClientRect();
  const box = editorWysiwyg.getBoundingClientRect();
  if (rect.height || rect.width) {
    editorWysiwyg.scrollTop += (rect.top - box.top) - editorWysiwyg.clientHeight / 2;
  }
  return true;
}

async function saveEditorFile() {
  if (!editorState.active || !currentFilePath) return;
  const md = getEditorMarkdown();
  const res = await window.lumadesk.writeFile(currentFilePath, md);
  if (res && res.success) {
    currentContent = md;
    if (editorState.mode === 'wysiwyg') editorTextarea.value = md;
    setEditorDirty(false);
    statusText.textContent = 'Saved';
  } else {
    statusText.textContent = 'Save failed: ' + ((res && res.error) || 'unknown');
  }
}

// ---- Formatting toolbar ----
function surroundSelection(before, after, placeholder) {
  const ta = editorTextarea;
  const s = ta.selectionStart, e = ta.selectionEnd;
  const sel = ta.value.slice(s, e) || placeholder;
  ta.setRangeText(before + sel + after, s, e, 'end');
  ta.focus();
  markEditorDirty();
  scheduleEditorSync();
}

function prefixLines(prefix) {
  const ta = editorTextarea;
  const s = ta.selectionStart, e = ta.selectionEnd;
  const val = ta.value;
  const lineStart = val.lastIndexOf('\n', s - 1) + 1;
  let lineEnd = val.indexOf('\n', e);
  if (lineEnd === -1) lineEnd = val.length;
  const block = val.slice(lineStart, lineEnd).split('\n').map((l) => prefix + l).join('\n');
  ta.setRangeText(block, lineStart, lineEnd, 'end');
  ta.focus();
  markEditorDirty();
  scheduleEditorSync();
}

function insertBlock(text) {
  const ta = editorTextarea;
  const s = ta.selectionStart;
  const before = s > 0 && ta.value[s - 1] !== '\n' ? '\n\n' : '';
  ta.setRangeText(before + text + '\n', s, ta.selectionEnd, 'end');
  ta.focus();
  markEditorDirty();
  scheduleEditorSync();
}

function applyEditorCommand(cmd) {
  if (editorState.mode === 'wysiwyg') return applyWysiwygCommand(cmd);
  switch (cmd) {
    case 'bold': surroundSelection('**', '**', 'bold text'); break;
    case 'italic': surroundSelection('*', '*', 'italic'); break;
    case 'code': surroundSelection('`', '`', 'code'); break;
    case 'h1': prefixLines('# '); break;
    case 'h2': prefixLines('## '); break;
    case 'h3': prefixLines('### '); break;
    case 'ul': prefixLines('- '); break;
    case 'ol': prefixLines('1. '); break;
    case 'quote': prefixLines('> '); break;
    case 'codeblock': insertBlock('```\n\n```'); break;
    case 'table': insertBlock('| Column A | Column B |\n| --- | --- |\n| Cell 1 | Cell 2 |'); break;
    case 'mermaid': insertBlock('```mermaid\nflowchart LR\n  A[Start] --> B[End]\n```'); break;
    case 'link': {
      const url = prompt('Link URL', 'https://');
      if (url !== null) surroundSelection('[', '](' + url + ')', 'text');
      break;
    }
  }
}

function applyWysiwygCommand(cmd) {
  const ex = (c, v) => document.execCommand(c, false, v);
  switch (cmd) {
    case 'bold': ex('bold'); break;
    case 'italic': ex('italic'); break;
    case 'h1': ex('formatBlock', 'H1'); break;
    case 'h2': ex('formatBlock', 'H2'); break;
    case 'h3': ex('formatBlock', 'H3'); break;
    case 'ul': ex('insertUnorderedList'); break;
    case 'ol': ex('insertOrderedList'); break;
    case 'quote': ex('formatBlock', 'BLOCKQUOTE'); break;
    case 'link': { const url = prompt('Link URL', 'https://'); if (url) ex('createLink', url); break; }
    default:
      statusText.textContent = 'Switch to Text mode to insert ' + cmd;
      return;
  }
  editorWysiwyg.focus();
  markEditorDirty();
  scheduleEditorSync();
}

// ---- Wiring ----
btnEdit.addEventListener('click', () => (editorState.active ? exitEditMode() : enterEditMode()));
editorSaveBtn.addEventListener('click', saveEditorFile);

editorToolbar.addEventListener('click', (e) => {
  const btn = e.target.closest('.ed-btn');
  if (!btn) return;
  if (btn.id === 'ed-playground') { openPlayground(); return; }
  if (btn.dataset.cmd) applyEditorCommand(btn.dataset.cmd);
});

editorModeSwitch.addEventListener('click', (e) => {
  const b = e.target.closest('.mode-opt');
  if (b) setEditorMode(b.dataset.mode);
});

editorTextarea.addEventListener('input', () => { markEditorDirty(); scheduleEditorSync(); });
editorWysiwyg.addEventListener('input', () => { markEditorDirty(); scheduleEditorSync(); });

// Tab inserts two spaces in the textarea instead of leaving the field.
editorTextarea.addEventListener('keydown', (e) => {
  if (e.key === 'Tab') {
    e.preventDefault();
    editorTextarea.setRangeText('  ', editorTextarea.selectionStart, editorTextarea.selectionEnd, 'end');
    markEditorDirty();
  }
});

// Draggable horizontal divider between preview (top) and editor (bottom).
let edResizing = false;
editorDivider.addEventListener('mousedown', (e) => {
  edResizing = true;
  editorDivider.classList.add('active');
  document.body.style.userSelect = 'none';
  e.preventDefault();
});
document.addEventListener('mousemove', (e) => {
  if (!edResizing) return;
  const rect = previewContainer.getBoundingClientRect();
  let h = rect.bottom - e.clientY;
  h = Math.max(80, Math.min(h, rect.height - 120));
  editorPane.style.height = h + 'px';
});
document.addEventListener('mouseup', () => {
  if (edResizing) {
    edResizing = false;
    editorDivider.classList.remove('active');
    document.body.style.userSelect = '';
  }
});

// ---- Keyboard Shortcuts ----
document.addEventListener('keydown', (e) => {
  // Presentation mode owns most keys when open
  if (presentOverlay.style.display !== 'none') {
    if (e.key === 'Escape') { e.preventDefault(); closePresent(); return; }
    if (e.key === 'ArrowRight' || e.key === 'PageDown' || e.key === ' ') {
      if (e.key === ' ' && e.shiftKey) { e.preventDefault(); presentPrev(); return; }
      e.preventDefault(); presentNext(); return;
    }
    if (e.key === 'ArrowLeft' || e.key === 'PageUp') { e.preventDefault(); presentPrev(); return; }
    if (e.key === 'Home') { e.preventDefault(); presentFirst(); return; }
    if (e.key === 'End') { e.preventDefault(); presentLast(); return; }
    if (e.key === 'f' || e.key === 'F') { e.preventDefault(); toggleFullscreen(); return; }
  }

  if (e.key === 'Escape' && modal.style.display !== 'none') {
    closeMermaidModal();
    return;
  }
  if (e.key === 'Escape' && fileContextMenu.style.display !== 'none') {
    closeFileContextMenu();
    return;
  }
  if (e.key === 'Escape' && playgroundModal.style.display !== 'none') {
    closePlayground();
    return;
  }
  if ((e.ctrlKey || e.metaKey) && e.key === 'm') {
    e.preventDefault();
    openPlayground();
    return;
  }
  if ((e.ctrlKey || e.metaKey) && e.key === 'f') {
    if (playgroundModal.style.display !== 'none') return;
    e.preventDefault();
    openSearch();
    return;
  }
  if ((e.ctrlKey || e.metaKey) && e.key === 'o') {
    e.preventDefault();
    if (e.shiftKey) openSingleFile();
    else openFolder();
  }
  if ((e.ctrlKey || e.metaKey) && e.key === 's') {
    e.preventDefault();
    if (editorState.active) saveEditorFile();
    return;
  }
  if ((e.ctrlKey || e.metaKey) && e.shiftKey && (e.key === 'e' || e.key === 'E')) {
    e.preventDefault();
    if (!btnEdit.disabled) btnEdit.click();
    return;
  }
  if ((e.ctrlKey || e.metaKey) && e.key === 'e') {
    e.preventDefault();
    if (!btnExportWord.disabled) btnExportWord.click();
  }
  if ((e.ctrlKey || e.metaKey) && e.key === 'p') {
    e.preventDefault();
    if (!btnExportPdf.disabled) btnExportPdf.click();
  }
  if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key === 'Enter') {
    e.preventDefault();
    if (!btnPresent.disabled) openPresent();
  }
});
