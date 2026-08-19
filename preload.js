"use strict";
// Electron 预加载脚本：向主窗口暴露"是否 Electron"标志与"立即重启以完成更新"能力。
const { contextBridge, ipcRenderer } = require("electron");
contextBridge.exposeInMainWorld("dsh", {
  isElectron: true,
  relaunch: () => ipcRenderer.send("app-relaunch")
});