const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const isDev = require('electron-is-dev');
const db = require('../backend/db');
const projectService = require('../backend/projectService');

function createWindow() {
  const win = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true
    }
  });

  win.loadURL(
    isDev
      ? 'http://localhost:3000'
      : `file://${path.join(__dirname, '../build/index.html')}`
  );
}

app.on('ready', () => {
  db.init();
  createWindow();
});

ipcMain.handle('project:create', async (event, data) => {
  return projectService.createProject(data);
});
console.log("Loading URL:", isDev
  ? 'http://localhost:3000'
  : `file://${path.join(__dirname, '../build/index.html')}`);

ipcMain.handle('project:list', async () => {
  return projectService.getProjects();
});
