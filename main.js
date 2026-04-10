const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const chokidar = require('chokidar');
const { exportToWord } = require('./lib/wordExport');

let mainWindow;
let watcher;
let currentRootDir = null;
let fileExtensions = ['.md', '.txt'];

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    title: 'LumaDesk',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  mainWindow.loadFile('renderer/index.html');

  if (process.argv.includes('--dev')) {
    mainWindow.webContents.openDevTools();
  }
}

app.whenReady().then(createWindow);
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });

// ---- IPC Handlers ----

ipcMain.handle('open-folder', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openDirectory'],
    title: 'Select Documentation Folder',
  });
  if (result.canceled || result.filePaths.length === 0) return null;

  const dirPath = result.filePaths[0];
  currentRootDir = dirPath;
  setupWatcher(dirPath);
  return { path: dirPath, tree: scanDirectory(dirPath) };
});

ipcMain.handle('read-file', async (_event, filePath) => {
  try {
    const content = fs.readFileSync(filePath, 'utf-8');
    return { content, fileName: path.basename(filePath, '.md') };
  } catch (err) {
    return { error: err.message };
  }
});

ipcMain.handle('render-markdown', async (_event, markdown) => {
  const { marked } = await import('marked');
  return marked.parse(markdown, { async: false });
});

ipcMain.handle('export-word', async (_event, { title, content, author }) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    title: 'Export as Word Document',
    defaultPath: `${sanitizeFilename(title)}.docx`,
    filters: [{ name: 'Word Document', extensions: ['docx'] }],
  });
  if (result.canceled) return { canceled: true };

  try {
    await exportToWord({ title, content, author, savePath: result.filePath });
    return { success: true, path: result.filePath };
  } catch (err) {
    return { error: err.message };
  }
});

ipcMain.handle('export-word-with-images', async (_event, { title, content, author, mermaidImages }) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    title: 'Export as Word Document',
    defaultPath: `${sanitizeFilename(title)}.docx`,
    filters: [{ name: 'Word Document', extensions: ['docx'] }],
  });
  if (result.canceled) return { canceled: true };

  try {
    await exportToWord({ title, content, author, savePath: result.filePath, mermaidImages });
    return { success: true, path: result.filePath };
  } catch (err) {
    return { error: err.message };
  }
});

ipcMain.handle('export-pdf', async (_event, { title, html }) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    title: 'Export as PDF',
    defaultPath: `${sanitizeFilename(title)}.pdf`,
    filters: [{ name: 'PDF Document', extensions: ['pdf'] }],
  });
  if (result.canceled) return { canceled: true };

  try {
    // Create a hidden window with only the document content
    const printWin = new BrowserWindow({
      width: 800,
      height: 600,
      show: false,
      webPreferences: {
        nodeIntegration: false,
        contextIsolation: true,
      },
    });

    const styledHtml = `<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>${title}</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    color: #1f2937; line-height: 1.7; font-size: 11pt;
    padding: 0; background: white;
  }
  .header { margin-bottom: 1.5em; padding-bottom: 0.8em; border-bottom: 2px solid #e5e7eb; }
  .header h1 { font-size: 22pt; margin-bottom: 0.2em; }
  .header .meta { color: #6b7280; font-size: 9pt; }
  h1 { font-size: 18pt; border-bottom: 1px solid #e5e7eb; padding-bottom: 0.25em; margin: 1.2em 0 0.4em; }
  h2 { font-size: 14pt; border-bottom: 1px solid #e5e7eb; padding-bottom: 0.2em; margin: 1em 0 0.4em; }
  h3 { font-size: 12pt; margin: 0.8em 0 0.3em; }
  h4 { font-size: 11pt; margin: 0.8em 0 0.3em; }
  p { margin: 0.5em 0; }
  a { color: #2563eb; text-decoration: none; }
  code { background: #f3f4f6; padding: 1px 4px; border-radius: 3px; font-family: Consolas, monospace; font-size: 0.9em; }
  pre { background: #f3f4f6; padding: 12px 16px; border-radius: 6px; overflow-x: auto; margin: 0.8em 0; border-left: 4px solid #d1d5db; }
  pre code { background: none; padding: 0; font-size: 9pt; }
  blockquote { margin: 0.8em 0; padding: 0.4em 1em; border-left: 4px solid #d1d5db; background: #f9fafb; color: #4b5563; }
  table { width: 100%; border-collapse: collapse; margin: 0.8em 0; font-size: 10pt; }
  th, td { border: 1px solid #d1d5db; padding: 6px 10px; text-align: left; }
  th { background: #e2e8f0; font-weight: 600; }
  tr:nth-child(even) { background: #f8fafc; }
  ul, ol { margin: 0.4em 0; padding-left: 1.8em; }
  li { margin: 0.15em 0; }
  hr { margin: 1.5em 0; border: none; border-top: 1px solid #e5e7eb; }
  img, svg { max-width: 100%; height: auto; }
  .mermaid-diagram { text-align: center; margin: 1em 0; }
  .mermaid-diagram svg { max-width: 100%; height: auto; }
  .footer { margin-top: 2em; padding-top: 0.8em; border-top: 1px solid #e5e7eb; color: #9ca3af; font-size: 8pt; text-align: center; }
  @page { margin: 1.5cm; }
</style>
</head><body>
  <div class="header">
    <h1>${title.replace(/</g, '&lt;')}</h1>
    <div class="meta">Exported ${new Date().toLocaleDateString()}</div>
  </div>
  ${html}
  <div class="footer">Exported from LumaDesk &mdash; <a href="https://peerluma.com">peerluma.com</a></div>
</body></html>`;

    await printWin.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(styledHtml));

    // Wait for images/content to load
    await new Promise(resolve => setTimeout(resolve, 500));

    const pdfData = await printWin.webContents.printToPDF({
      printBackground: true,
      pageSize: 'A4',
    });

    printWin.close();
    fs.writeFileSync(result.filePath, pdfData);
    return { success: true, path: result.filePath };
  } catch (err) {
    return { error: err.message };
  }
});

ipcMain.handle('get-recent-folder', () => {
  return currentRootDir;
});

ipcMain.handle('rescan-folder', async () => {
  if (!currentRootDir) return null;
  return { path: currentRootDir, tree: scanDirectory(currentRootDir) };
});

ipcMain.handle('get-extensions', () => fileExtensions);

ipcMain.handle('set-extensions', async (_event, extensions) => {
  fileExtensions = extensions.map(e => e.startsWith('.') ? e : '.' + e).filter(Boolean);
  if (currentRootDir) {
    const tree = scanDirectory(currentRootDir);
    mainWindow.webContents.send('tree-changed', tree);
    return { path: currentRootDir, tree };
  }
  return null;
});

// ---- File Scanner ----

function scanDirectory(dirPath, depth = 0) {
  const entries = [];
  let items;

  try {
    items = fs.readdirSync(dirPath, { withFileTypes: true });
  } catch {
    return entries;
  }

  // Sort: directories first, then files; both alphabetically (numbered prefixes sort naturally)
  const dirs = items.filter(i => i.isDirectory() && !i.name.startsWith('.')).sort((a, b) => a.name.localeCompare(b.name));
  const files = items.filter(i => i.isFile() && fileExtensions.some(ext => i.name.endsWith(ext))).sort((a, b) => a.name.localeCompare(b.name));

  for (const dir of dirs) {
    const fullPath = path.join(dirPath, dir.name);
    const children = scanDirectory(fullPath, depth + 1);
    if (children.length > 0) { // Only show folders that contain .md files
      entries.push({
        type: 'folder',
        name: cleanDisplayName(dir.name),
        rawName: dir.name,
        path: fullPath,
        children,
      });
    }
  }

  for (const file of files) {
    entries.push({
      type: 'file',
      name: cleanDisplayName(file.name.replace(/\.[^.]+$/, '')),
      rawName: file.name,
      path: path.join(dirPath, file.name),
    });
  }

  return entries;
}

/**
 * Remove leading number prefixes like "01-", "02_", "1. "
 */
function cleanDisplayName(name) {
  return name.replace(/^\d+[-_.\s]+/, '').trim();
}

function sanitizeFilename(name) {
  return name.replace(/[^a-zA-Z0-9-_ ]/g, '').substring(0, 100).trim();
}

// ---- File Watcher ----

function setupWatcher(dirPath) {
  if (watcher) watcher.close();

  watcher = chokidar.watch(dirPath, {
    ignored: /(^|[\/\\])\./,
    persistent: true,
    ignoreInitial: true,
    depth: 10,
  });

  watcher.on('change', (filePath) => {
    if (fileExtensions.some(ext => filePath.endsWith(ext)) && mainWindow) {
      mainWindow.webContents.send('file-changed', filePath);
    }
  });

  watcher.on('add', () => notifyTreeChange());
  watcher.on('unlink', () => notifyTreeChange());
  watcher.on('addDir', () => notifyTreeChange());
  watcher.on('unlinkDir', () => notifyTreeChange());
}

let treeChangeTimeout;
function notifyTreeChange() {
  clearTimeout(treeChangeTimeout);
  treeChangeTimeout = setTimeout(() => {
    if (currentRootDir && mainWindow) {
      mainWindow.webContents.send('tree-changed', scanDirectory(currentRootDir));
    }
  }, 500);
}
