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
