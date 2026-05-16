const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 800,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  mainWindow.loadFile(path.join(__dirname, 'index.html'));
  mainWindow.setMenuBarVisibility(false);
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

// 폴더 선택 다이얼로그
ipcMain.handle('select-folder', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openDirectory'],
  });
  if (result.canceled || !result.filePaths.length) return null;
  return result.filePaths[0];
});

// 파일 저장 (중복 시 (1), (2) 붙이기)
ipcMain.handle('save-file', async (event, folderPath, filename, buffer) => {
  const safeName = filename.replace(/[/\\/]/g, '_').replace(/^\.+/, '') || 'download';
  const ext = path.extname(safeName);
  const base = path.basename(safeName, ext);

  let finalPath = path.join(folderPath, safeName);
  let counter = 1;
  while (fs.existsSync(finalPath)) {
    finalPath = path.join(folderPath, `${base}(${counter})${ext}`);
    counter++;
  }

  fs.writeFileSync(finalPath, Buffer.from(buffer));
  return finalPath;
});