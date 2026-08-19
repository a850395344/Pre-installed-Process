"use strict";

const { app, BrowserWindow, shell, ipcMain } = require("electron");
const path = require("path");
const fs = require("fs");
const http = require("http");
const { spawn } = require("child_process");

app.disableHardwareAcceleration();

const PORT_FILE = path.join(__dirname, "data", "port.json");
function getPort() {
  try {
    const p = Number(fs.readFileSync(PORT_FILE, "utf8").trim());
    if (Number.isInteger(p) && p > 0 && p < 65536) return p;
  } catch {}
  return 1314;
}

const port = getPort();
const url = `http://127.0.0.1:${port}`;
let serverProcess = null;
let mainWindow = null;

// ==================== 单实例（窗口版）：只允许一个主窗口 ====================
const gotLock = app.requestSingleInstanceLock();
if (!gotLock) {
  app.quit();
}
app.on("second-instance", () => {
  if (mainWindow) {
    if (mainWindow.isMinimized()) mainWindow.restore();
    mainWindow.focus();
  }
});

function isServerUp() {
  return new Promise((resolve) => {
    const req = http.get(url, (res) => {
      res.resume();
      resolve(true);
    });
    req.setTimeout(1200, () => {
      req.destroy();
      resolve(false);
    });
    req.on("error", () => resolve(false));
  });
}

function startServer() {
  return new Promise((resolve) => {
    serverProcess = spawn(process.execPath, [path.join(__dirname, "server.js")], {
      cwd: __dirname,
      env: { ...process.env, PORT: String(port), ELECTRON_RUN_AS_NODE: "1" },
      stdio: "ignore",
      windowsHide: true
    });
    serverProcess.on("exit", () => { serverProcess = null; });
    const tryConnect = async () => {
      if (await isServerUp()) return resolve();
      setTimeout(tryConnect, 300);
    };
    tryConnect();
  });
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 1024,
    minHeight: 700,
    title: "四轮汽车线束后段预装工艺文件生成器",
    autoHideMenuBar: true,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      preload: path.join(__dirname, "preload.js")
    }
  });

  mainWindow.loadURL(url);

  mainWindow.webContents.setWindowOpenHandler(({ url: targetUrl }) => {
    shell.openExternal(targetUrl);
    return { action: "deny" };
  });

  mainWindow.on("closed", () => {
    mainWindow = null;
  });
}

app.whenReady().then(async () => {
  if (!gotLock) return;
  if (!(await isServerUp())) {
    await startServer();
  }
  createWindow();

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on("window-all-closed", () => {
  if (serverProcess) {
    serverProcess.kill();
    serverProcess = null;
  }
  app.quit();
});

// 更新完成后的“立即重启”：先结束内嵌 server（停掉旧代码进程），再 relaunch 应用，
// 新实例启动时会自动应用 data/pending_update 里的更新包。
ipcMain.on("app-relaunch", () => {
  if (serverProcess) {
    try { serverProcess.kill(); } catch {}
    serverProcess = null;
  }
  app.relaunch();
  app.exit(0);
});
