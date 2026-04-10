const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('lumadesk', {
  openFolder: () => ipcRenderer.invoke('open-folder'),
  readFile: (filePath) => ipcRenderer.invoke('read-file', filePath),
  renderMarkdown: (markdown) => ipcRenderer.invoke('render-markdown', markdown),
  exportWord: (options) => ipcRenderer.invoke('export-word', options),
  rescanFolder: () => ipcRenderer.invoke('rescan-folder'),
  getExtensions: () => ipcRenderer.invoke('get-extensions'),
  setExtensions: (exts) => ipcRenderer.invoke('set-extensions', exts),

  exportWordWithImages: (options) => ipcRenderer.invoke('export-word-with-images', options),
  exportPdf: (options) => ipcRenderer.invoke('export-pdf', options),

  // Events from main process
  onFileChanged: (callback) => ipcRenderer.on('file-changed', (_e, filePath) => callback(filePath)),
  onTreeChanged: (callback) => ipcRenderer.on('tree-changed', (_e, tree) => callback(tree)),
});
