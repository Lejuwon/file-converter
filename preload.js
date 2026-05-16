const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  selectFolder: () => ipcRenderer.invoke('select-folder'),
  saveFile: (folderPath, filename, buffer) =>
    ipcRenderer.invoke('save-file', folderPath, filename, buffer),
});
