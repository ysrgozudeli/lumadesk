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

  // Render mermaid diagrams
  await renderMermaidBlocks();

  btnExportWord.disabled = false;
  btnExportPdf.disabled = false;
  statusFile.textContent = filePath.split(/[/\\]/).pop();
  statusText.textContent = 'Ready';

  // Scroll to top
  $('#preview-container').scrollTop = 0;
}

// ---- Word Export ----
btnExportWord.addEventListener('click', async () => {
  if (!currentContent) return;

  statusText.textContent = 'Rendering diagrams...';
  btnExportWord.disabled = true;

  // Capture mermaid diagrams as PNG for embedding in Word
  const mermaidImages = await captureMermaidImages(currentContent);

  statusText.textContent = 'Exporting to Word...';

  const result = await window.lumadesk.exportWordWithImages({
    title: currentTitle,
    content: currentContent,
    author: '',
    mermaidImages,
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

  // Send the rendered HTML (including mermaid SVGs) for a clean document PDF
  const result = await window.lumadesk.exportPdf({
    title: currentTitle,
    html: previewEl.innerHTML,
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

// ---- Keyboard Shortcuts ----
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && modal.style.display !== 'none') {
    closeMermaidModal();
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
});
