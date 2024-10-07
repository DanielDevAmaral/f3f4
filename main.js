const { app, BrowserWindow } = require('electron');
const path = require('path');


let mainWindow;
const isDev = process.env.NODE_ENV === "production";
const isMac = process.platform !== 'darwin';

function createMainWindow() {
  mainWindow = new BrowserWindow({
    width: isDev ? 1200 : 800,
    height: 600,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: true, // Para permitir o uso de require no frontend
      preload: path.join(__dirname, 'preload.js'),
    },
  });

  // open devtools
  if (isDev){
    mainWindow.webContents.openDevTools();
  }

  mainWindow.loadFile(path.join(__dirname, './page/index.html'))
}

app.whenReady().then(() => {
  createMainWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createMainWindow()
    }
  })
})

app.on('window-all-closed', function () {
  if (!isMac) {
    app.quit();
  }
});
