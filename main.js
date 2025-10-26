const { app, BrowserWindow, ipcMain, dialog } = require("electron");
const path = require("path");
const { spawn } = require("child_process");
const fs = require("fs");

let mainWindow;
let backendProcess;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
    },
  });

  const prodIndex = path.join(__dirname, "frontend-build", "index.html");
  const isDev = !!process.env.ELECTRON_DEV || !fs.existsSync(prodIndex);
  if (isDev) {
    // 🔹 Development (fallback to dev server if no build present)
    mainWindow.loadURL("http://localhost:5173");
    mainWindow.webContents.openDevTools();
  } else {
    // 🔹 Production build exists
    mainWindow.loadFile(prodIndex);
  }

  mainWindow.on("closed", () => {
    mainWindow = null;
  });
}

function startBackend() {
  const backendPath = path.join(__dirname, "backend", "src", "server.js");

  backendProcess = spawn(process.execPath, [backendPath], {
    env: { ...process.env, PORT: 5000 },
    stdio: "inherit",
  });

  backendProcess.on("error", (err) => {
    console.error("❌ Failed to start backend:", err);
  });
}

app.on("ready", () => {
  startBackend();
  createWindow();
});

// IPC handlers for dialogs
ipcMain.handle("dialog:selectDirectory", async () => {
  const result = await dialog.showOpenDialog(mainWindow || undefined, {
    properties: ["openDirectory", "createDirectory"],
  });
  if (result.canceled || !result.filePaths || result.filePaths.length === 0) {
    return null;
  }
  return result.filePaths[0];
});

ipcMain.handle("dialog:saveFile", async (_event, opts) => {
  const result = await dialog.showSaveDialog(mainWindow || undefined, {
    title: "Select Excel file",
    defaultPath: (opts && opts.defaultPath) || "workbook.xlsx",
    filters: [
      { name: "Excel Workbook", extensions: ["xlsx", "xls"] },
      { name: "All Files", extensions: ["*"] },
    ],
  });
  if (result.canceled || !result.filePath) {
    return null;
  }
  return result.filePath;
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});

app.on("before-quit", () => {
  if (backendProcess) backendProcess.kill();
});

app.on("activate", () => {
  if (mainWindow === null) createWindow();
});
