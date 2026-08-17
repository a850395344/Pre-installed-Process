"use strict";

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const DATA_DIR = path.join(__dirname, "..", "data");
const HISTORY_FILE = path.join(DATA_DIR, "version_history.json");
const AUTH_FILE = path.join(DATA_DIR, "backend_auth.json");

function readJson(file, fallback) {
  try {
    return JSON.parse(fs.readFileSync(file, "utf8"));
  } catch {
    return fallback;
  }
}

function writeJson(file, data) {
  fs.writeFileSync(file, JSON.stringify(data, null, 2), "utf8");
}

function getVersionHistory() {
  const data = readJson(HISTORY_FILE, { version: "1.0.0", records: [] });
  return data;
}

function bumpVersion(version) {
  const parts = String(version || "1.0.0").split(".").map(Number);
  parts[2] = (parts[2] || 0) + 1;
  return parts.join(".");
}

function addRecord({ action = "修改", detail = "", modules = [], bump = true } = {}) {
  const data = getVersionHistory();
  const newVersion = bump ? bumpVersion(data.version) : data.version;
  const record = {
    version: newVersion,
    time: new Date().toISOString(),
    action,
    detail,
    modules: Array.isArray(modules) ? modules : []
  };
  data.version = newVersion;
  data.records.unshift(record);
  writeJson(HISTORY_FILE, data);
  return { version: newVersion, record };
}

function verifyBackendPassword(password) {
  const auth = readJson(AUTH_FILE, null);
  if (!auth || !auth.salt || !auth.hash || !auth.iterations) return false;
  const hash = crypto.pbkdf2Sync(String(password || ""), auth.salt, auth.iterations, 32, "sha256").toString("hex");
  return crypto.timingSafeEqual(Buffer.from(hash, "hex"), Buffer.from(auth.hash, "hex"));
}

function buildBackendData(extra = {}) {
  const history = getVersionHistory();
  const files = [];
  const root = path.join(__dirname, "..");
  const skipDirs = new Set(["node_modules", ".git", "data", "dist"]);
  function walk(dir, prefix) {
    let entries = [];
    try { entries = fs.readdirSync(dir, { withFileTypes: true }); } catch { return; }
    for (const entry of entries) {
      if (skipDirs.has(entry.name)) continue;
      const full = path.join(dir, entry.name);
      if (entry.isDirectory()) walk(full, prefix + entry.name + "/");
      else files.push({ path: prefix + entry.name, size: fs.statSync(full).size });
    }
  }
  walk(root, "");
  const continuationGuide = [
    "1. 将本后台数据（含“功能一览”“核心文件说明”“续作约定”）完整复制给新AI。",
    "2. 把 D:\\分工艺软件\\harness-process-app 整个目录加入新聊天窗口的工作目录。",
    "3. 软件定位 = 根据客户上传的项目资料（标准工时/EBOM/MBOM/工艺PDF）生成四轮汽车线束后段预装工艺。",
    "4. 【最重要约定】所有数据（配置号、护套、导线、工时、岗位数）只来自客户在页面“放入项目资料”上传的文件；绝不读取/参考工作目录里任何项目资料文件（那些可能被移走或属于其他项目）。",
    "5. 正式工时只来自标准工时文件；用户明确提供的固定工时（如KIT打圈50秒/件）除外；禁止经验式虚构工时。",
    "6. 继续修改请按 data/version_history.json 的版本记录演进，不得回退已有功能；改动后必须 node --check 校验 JS，并确保 server.js、src/*、public/* 三者一致。",
    "7. dist\\预装工艺生成器-win32-x64\\resources\\app 与 dist\\预装工艺生成器-Dream.zip 是分发/打包副本；改完源码如需更新，先关闭正在运行的软件再打包（运行时会占用文件导致压缩失败）。",
    "8. 优先阅读与维护：README.md、server.js、src/generator.js（核心）、src/parse.js、src/validation.js、src/templates.js、src/version.js、public/index.html、public/app.js、public/ai.html、public/backend.html、public/version.html。",
    "9. 【维护纪律】每次功能更新后，务必重新检查并同步本后台数据的“功能一览/续作说明/核心文件说明”与 README.md，确保接续信息反映最新状态，便于新AI无损续作（勿让接续信息过期）。"
  ];
  const featureSummary = [
    "总体：本地 Node+浏览器（可 Electron 窗口）的线束后段预装工艺生成器；读取客户上传的标准工时/EBOM/MBOM/工艺PDF，按 V2.6.1 预装规则输出：结构化工艺数据包、预装工艺岗位表、人工核查表。",
    "岗位数口径：岗位数以“单配置最高工时 ÷ 目标节拍TT”为准（TT优先，最多岗位数填少时自动扩岗）；最多岗位数可不填，分析文件后按最高配置工时÷TT 在输入框给出“预计岗位数”占位提示。",
    "岗位级同色线编码：八色（PK/RD/OG/YE/GN/WH/VT/BK）无序多色组合，按岗位建立编码池并查重。",
    "同岗位护套分组：一组护套=一个岗位、组间强制分岗；组超节拍时可“强制同岗”或“按最佳插接率拆分”（输出到“同岗位分组节拍处理”工作表）。",
    "其他：在线超声波限制（需热缩管SP/SC点、单配置/总组数上限）、强制线下插接、胶套专属岗位。",
    "前端：必填项红星 + 缺失弹窗 + 红框闪烁；上传文件服务端暂存/刷新恢复/删除/清空；三个导出Excel。",
    "AI多模态：PDF 用 pdf-to-img 转图片后走视觉模型分批识别（每批≤4张/约8MB，模型按 needMore 决定是否继续）。",
    "页面与版本：AI多模态、版本记录、后台数据（本页，供接续）。"
  ];
  const coreFiles = [
    "server.js: Express 服务端路由（文件分析、生成、导出、项目资料暂存/删除、AI多模态、后台鉴权、版本）。",
    "src/generator.js: 核心生成逻辑（解析→工时→工作包→护套关联→同色编码→岗位分配→TT→导出工作簿）。",
    "src/parse.js: 解析 MBOM/EBOM/标准工时/PDF。",
    "src/validation.js: 上传文件模板校验。",
    "src/templates.js: 三套 Excel 模板下载。",
    "src/version.js: 版本历史、后台数据、密码校验。",
    "public/index.html + app.js + styles.css: 主页面与交互。",
    "public/ai.html: AI多模态页面；public/version.html: 版本记录；public/backend.html: 后台数据。"
  ];
  return {
    generatedAt: new Date().toISOString(),
    appVersion: require("../package.json").version,
    versionHistory: history,
    files,
    continuationGuide,
    featureSummary,
    coreFiles,
    ...extra
  };
}

module.exports = {
  getVersionHistory,
  addRecord,
  verifyBackendPassword,
  buildBackendData
};
