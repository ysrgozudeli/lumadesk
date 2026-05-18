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
    }
  }
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
  statusText.textContent = `Opened: ${result.path}`;

  // Auto-select first file
  const firstFile = findFirstFile(result.tree);
  if (firstFile) {
    await selectFile(firstFile.path);
  }
}

btnOpenFolder.addEventListener('click', openFolder);
btnOpenEmpty.addEventListener('click', openFolder);

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

  // Render mermaid diagrams
  await renderMermaidBlocks();

  btnExportWord.disabled = false;
  btnExportPdf.disabled = false;
  btnExportPpt.disabled = false;
  btnPresent.disabled = false;
  statusFile.textContent = filePath.split(/[/\\]/).pop();
  statusText.textContent = 'Ready';

  // Scroll to top
  $('#preview-container').scrollTop = 0;
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

  // Capture mermaid diagrams as PNG for embedding in Word
  const mermaidImages = await captureMermaidImages(currentContent);

  statusText.textContent = 'Exporting to Word...';

  const settings = loadWordSettings();
  const result = await window.lumadesk.exportWordWithImages({
    title: currentTitle,
    content: currentContent,
    author: '',
    mermaidImages,
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
  if (filePath === currentFilePath) {
    const result = await window.lumadesk.readFile(filePath);
    if (!result.error) {
      currentContent = result.content;
      const html = await window.lumadesk.renderMarkdown(result.content);
      previewEl.innerHTML = html;
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

// Click on mermaid diagram in preview → open modal
previewEl.addEventListener('click', (e) => {
  const diagram = e.target.closest('.mermaid-diagram');
  if (diagram) {
    openMermaidModal(diagram.innerHTML);
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
  presentSlideEl.innerHTML = html;
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
      } catch {
        // Leave the code block visible if mermaid fails — user still sees source
      }
    }
  }

  presentSlideEl.scrollTop = 0;
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

async function openPresent() {
  if (!currentContent) return;
  presentSlides = window.LumaSlides.parseSlides(currentContent);
  if (!presentSlides.length) {
    statusText.textContent = 'No slides parsed';
    return;
  }
  presentIndex = 0;
  presentOverlay.style.display = 'flex';
  await renderCurrentSlide();
}

function closePresent() {
  presentOverlay.style.display = 'none';
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
    openFolder();
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
