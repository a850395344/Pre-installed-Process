"use strict";

const path = require("path");
const fs = require("fs");
const express = require("express");
const multer = require("multer");
const pdfParse = require("pdf-parse");
const crypto = require("crypto");
const { execSync } = require("child_process");
const { analyzeProject, analyzeFilesForPreview, buildWorkbook, buildProcessWorkbook, buildReviewWorkbook } = require("./src/generator");
const { buildTemplate } = require("./src/templates");
const { getVersionHistory, addRecord, verifyBackendPassword, buildBackendData } = require("./src/version");

// ==================== 单实例（窗口版 + 浏览器版共用） ====================
// 用 data/instance.lock 记录PID；已有存活实例时拒绝启动第二个服务。
const INSTANCE_LOCK = path.join(__dirname, "data", "instance.lock");
function acquireSingleInstanceLock() {
  try {
    const old = String(fs.readFileSync(INSTANCE_LOCK, "utf8") || "").trim();
    if (old) {
      const pid = Number(old);
      if (pid && pid !== process.pid) {
        try {
          process.kill(pid, 0);
          return false; // 已有存活实例
        } catch {
          // 旧锁进程已不存在（陈旧锁），继续占用
        }
      }
    }
  } catch {}
  try {
    fs.mkdirSync(path.dirname(INSTANCE_LOCK), { recursive: true });
    fs.writeFileSync(INSTANCE_LOCK, String(process.pid), "utf8");
  } catch {}
  process.on("exit", () => { try { fs.unlinkSync(INSTANCE_LOCK); } catch {} });
  return true;
}
if (!acquireSingleInstanceLock()) {
  console.error("[单实例] 预装工艺生成器已在运行，请关闭已有窗口/服务或直接使用其页面。本实例退出。");
  process.exit(1);
}

// ==================== 本地生成历史（data/gen_history，与版本记录分开） ====================
const GEN_HISTORY_DIR = path.join(__dirname, "data", "gen_history");
const GEN_HISTORY_MAX_ENTRIES = 30;
const GEN_HISTORY_MAX_TOTAL_MB = 300;
function genHistoryIndexFile() {
  return path.join(GEN_HISTORY_DIR, "index.json");
}
function readGenHistoryIndex() {
  try { return JSON.parse(fs.readFileSync(genHistoryIndexFile(), "utf8")); } catch { return []; }
}
function writeGenHistoryIndex(list) {
  fs.mkdirSync(GEN_HISTORY_DIR, { recursive: true });
  fs.writeFileSync(genHistoryIndexFile(), JSON.stringify(list, null, 2), "utf8");
}
// 计算目录总大小（MB）
function dirSizeMB(dir) {
  let total = 0;
  try {
    const walk = (d) => {
      for (const ent of fs.readdirSync(d, { withFileTypes: true })) {
        const full = path.join(d, ent.name);
        if (ent.isDirectory()) walk(full);
        else if (ent.isFile()) total += fs.statSync(full).size;
      }
    };
    walk(dir);
  } catch {}
  return total / (1024 * 1024);
}
// 生成完成后：备份上传文件 + 结果JSON，写入索引（时间倒序，最新在前）
function saveGenHistory(files, result, regions, modeLabel) {
  try {
    const id = new Date().toISOString().replace(/[:.]/g, "-");
    const dir = path.join(GEN_HISTORY_DIR, id);
    fs.mkdirSync(dir, { recursive: true });
    const fileNames = {};
    for (const key of ["standard", "ebom", "mbom", "pdf"]) {
      const f = files[key];
      if (!f || !f.buffer) continue;
      fs.writeFileSync(path.join(dir, key), f.buffer);
      fileNames[key] = f.originalname || key;
    }
    fs.writeFileSync(path.join(dir, "result.json"), JSON.stringify(result), "utf8");
    const entry = {
      id,
      time: new Date().toISOString(),
      regions: (regions || []).join("、") || "未选择",
      modeLabel: modeLabel || "未选择",
      stations: (result.plan && result.plan.stationDetails && result.plan.stationDetails.length) || 0,
      wires: (result.summary && result.summary.wires) || 0,
      packages: (result.summary && result.summary.packages) || 0,
      issues: (result.summary && result.summary.issues) || 0,
      tt: result.options && result.options.tt != null ? result.options.tt : "",
      missingTimeSource: !!(result.plan && result.plan.missingTimeSource),
      fileNames
    };
    const list = [entry, ...readGenHistoryIndex()].slice(0, GEN_HISTORY_MAX_ENTRIES);
    // 超容量时删除最旧条目
    while (dirSizeMB(GEN_HISTORY_DIR) > GEN_HISTORY_MAX_TOTAL_MB && list.length > 1) {
      const last = list.pop();
      try { fs.rmSync(path.join(GEN_HISTORY_DIR, last.id), { recursive: true, force: true }); } catch {}
    }
    writeGenHistoryIndex(list);
    return entry;
  } catch (e) {
    console.error("保存生成历史失败:", e.message);
    return null;
  }
}
function deleteGenHistory(id) {
  const list = readGenHistoryIndex();
  const next = list.filter((e) => e.id !== id);
  if (next.length === list.length) return false;
  try { fs.rmSync(path.join(GEN_HISTORY_DIR, id), { recursive: true, force: true }); } catch {}
  writeGenHistoryIndex(next);
  return true;
}

function parseList(value) {
  if (!value) return [];
  if (Array.isArray(value)) return value.map((v) => String(v).trim()).filter(Boolean);
  return String(value).split(/[\n,，;；]+/).map((s) => s.trim()).filter(Boolean);
}

const PORT_FILE = path.join(__dirname, "data", "port.json");
const TEMP_PDF = path.join(__dirname, "data", "temp_upload.pdf");
function saveTempPdf(buffer) {
  if (buffer) fs.writeFileSync(TEMP_PDF, buffer);
}
function getConfiguredPort() {
  try {
    const p = Number(fs.readFileSync(PORT_FILE, "utf8").trim());
    if (Number.isInteger(p) && p > 0 && p < 65536) return p;
  } catch {}
  return 1314;
}

// ==================== AI多模态：PDF转图片 + 分批识别 ====================
const MAX_AI_BATCH_PAGES = 4;            // 每批最多几张图纸图片
const MAX_AI_BATCH_BASE64 = 8 * 1024 * 1024; // 每批base64估算上限（约8MB）

// PDF → 每页PNG图片（pdf-to-img 为 ESM，需动态导入）
async function renderPdfPages(buffer) {
  const { pdf } = await import("pdf-to-img");
  const pages = [];
  const document = await pdf(buffer, { scale: 1.5 });
  for await (const img of document) pages.push(img);
  return pages;
}

function base64Estimate(bytes) {
  return Math.ceil(bytes * 4 / 3) + 60;
}

// 把图片按"张数上限 + base64总量上限"分批；单张超大图单独一批
function splitImagesIntoBatches(images) {
  const batches = [];
  let cur = [];
  let curBytes = 0;
  for (const img of images) {
    const est = base64Estimate(img.length);
    if (est > MAX_AI_BATCH_BASE64) {
      if (cur.length) { batches.push(cur); cur = []; curBytes = 0; }
      batches.push([img]);
      continue;
    }
    if (cur.length >= MAX_AI_BATCH_PAGES || curBytes + est > MAX_AI_BATCH_BASE64) {
      batches.push(cur);
      cur = [];
      curBytes = 0;
    }
    cur.push(img);
    curBytes += est;
  }
  if (cur.length) batches.push(cur);
  return batches;
}

function imageContentBlock(img, mime) {
  return { type: "image_url", image_url: { url: `data:${mime || "image/png"};base64,${img.toString("base64")}` } };
}

// 合并多批识别结果（数组拼接；对象按字段合并，数组字段追加）
function mergeAiBatches(batches) {
  const arr = [];
  const obj = {};
  let isArrayMode = null;
  for (const b of batches) {
    let v = b;
    if (v && typeof v === "object" && Object.prototype.hasOwnProperty.call(v, "needMore")) {
      const copy = { ...v };
      delete copy.needMore;
      v = Object.keys(copy).length ? copy : {};
    }
    if (Array.isArray(v)) {
      isArrayMode = true;
      arr.push(...v);
    } else if (v && typeof v === "object") {
      for (const [k, val] of Object.entries(v)) {
        if (Array.isArray(obj[k]) && Array.isArray(val)) obj[k] = [...obj[k], ...val];
        else obj[k] = val;
      }
    }
  }
  return isArrayMode ? arr : obj;
}

async function callAiModel(endpoint, apiKey, model, content) {
  const payload = {
    model,
    temperature: 0.2,
    messages: [
      { role: "system", content: "你是汽车线束工艺图纸识别AI。请只输出JSON，不要输出额外解释。JSON字段建议包含：housings、splices、branches、protections、dimensions、notes。若需要继续查看后续图纸，请在JSON中加入 needMore: true。" },
      { role: "user", content }
    ]
  };
  const resp = await fetch(endpoint, {
    method: "POST",
    headers: { "Content-Type": "application/json", "Authorization": `Bearer ${apiKey}` },
    body: JSON.stringify(payload)
  });
  if (!resp.ok) {
    const errText = await resp.text();
    throw new Error(`AI接口返回错误 ${resp.status}: ${errText.slice(0, 500)}`);
  }
  const data = await resp.json();
  const raw = data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content;
  if (!raw) throw new Error("AI接口未返回内容");
  let parsed = raw;
  try { parsed = JSON.parse(raw); } catch {}
  return parsed;
}

const app = express();
// 导出接口通过JSON把完整result回传生成Excel；限制放宽到500mb作为保险（本机单用户无安全风险，
// 真实体积取决于项目规模：当前2000根导线量级仅数MB；pdf.text已在结果中截断）
app.use(express.json({ limit: "500mb" }));

const aiJobs = new Map();
let aiJobSeq = 1;

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 1024 * 1024 * 1024, files: 6 }
});

// ==================== 项目资料暂存（页面跳转/刷新后恢复，避免重复上传） ====================
const STASH_DIR = path.join(__dirname, "data", "stash");
const STASH_META = path.join(STASH_DIR, "meta.json");
const STASH_FIELDS = ["standard", "ebom", "mbom", "pdf"];
function readStashMeta() {
  try { return JSON.parse(fs.readFileSync(STASH_META, "utf8")); } catch { return {}; }
}
function writeStashMeta(meta) {
  fs.mkdirSync(STASH_DIR, { recursive: true });
  fs.writeFileSync(STASH_META, JSON.stringify(meta, null, 2), "utf8");
}

// 上传文件时立即暂存到服务端（覆盖式，每个字段保留最新一份）
app.post("/api/stash", upload.fields([
  { name: "standard", maxCount: 1 },
  { name: "ebom", maxCount: 1 },
  { name: "mbom", maxCount: 1 },
  { name: "pdf", maxCount: 1 }
]), (req, res) => {
  try {
    const meta = readStashMeta();
    for (const key of STASH_FIELDS) {
      if (req.files && req.files[key] && req.files[key][0]) {
        const f = req.files[key][0];
        fs.mkdirSync(STASH_DIR, { recursive: true });
        fs.writeFileSync(path.join(STASH_DIR, key), f.buffer);
        meta[key] = {
          name: Buffer.from(f.originalname, "latin1").toString("utf8") || f.originalname,
          size: f.buffer.length
        };
      }
    }
    writeStashMeta(meta);
    res.json({ ok: true, files: meta });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: error.message || "暂存失败" });
  }
});

// 列出已暂存的项目资料
app.get("/api/stash", (req, res) => {
  const meta = readStashMeta();
  const out = {};
  for (const key of STASH_FIELDS) {
    const p = path.join(STASH_DIR, key);
    if (meta[key] && fs.existsSync(p)) out[key] = { name: meta[key].name, size: fs.statSync(p).size };
  }
  res.json({ files: out });
});

// 下载暂存文件（用于返回主页面时恢复）
app.get("/api/stash/:field", (req, res) => {
  const field = req.params.field;
  if (!STASH_FIELDS.includes(field)) return res.status(400).json({ error: "未知字段" });
  const p = path.join(STASH_DIR, field);
  if (!fs.existsSync(p)) return res.status(404).json({ error: "无暂存文件" });
  const meta = readStashMeta();
  const name = (meta[field] && meta[field].name) || field;
  res.setHeader("Content-Type", "application/octet-stream");
  res.setHeader("Content-Disposition", `attachment; filename="stash"; filename*=UTF-8''${encodeURIComponent(name)}`);
  res.send(fs.readFileSync(p));
});

// 删除单个暂存文件
app.delete("/api/stash/:field", (req, res) => {
  const field = req.params.field;
  if (!STASH_FIELDS.includes(field)) return res.status(400).json({ error: "未知字段" });
  try {
    const p = path.join(STASH_DIR, field);
    if (fs.existsSync(p)) fs.unlinkSync(p);
    const meta = readStashMeta();
    delete meta[field];
    writeStashMeta(meta);
    res.json({ ok: true });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: error.message || "删除失败" });
  }
});

// 清空全部暂存文件
app.delete("/api/stash", (req, res) => {
  try {
    for (const key of STASH_FIELDS) {
      const p = path.join(STASH_DIR, key);
      if (fs.existsSync(p)) fs.unlinkSync(p);
    }
    writeStashMeta({});
    res.json({ ok: true });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: error.message || "清空失败" });
  }
});

app.post("/api/analyze-files", upload.fields([
  { name: "standard", maxCount: 1 },
  { name: "ebom", maxCount: 1 },
  { name: "mbom", maxCount: 1 },
  { name: "pdf", maxCount: 1 }
]), async (req, res) => {
  try {
    const files = {};
    for (const key of ["standard", "ebom", "mbom", "pdf"]) {
      if (req.files && req.files[key] && req.files[key][0]) {
        const f = req.files[key][0];
        files[key] = { buffer: f.buffer, originalname: Buffer.from(f.originalname, "latin1").toString("utf8") || f.originalname };
      }
    }
    if (!files.mbom) {
      return res.status(400).json({ error: "请先上传MBOM/Cutting文件。" });
    }
    if (files.pdf) saveTempPdf(files.pdf.buffer);
    const tt = req.body.tt ? Number(req.body.tt) : null;
    const analysis = await analyzeFilesForPreview(files, {
      tt,
      preassemblyMode: req.body.preassemblyMode || "",
      loopTimes: {
        singleKit: req.body.loopSingleKit ? Number(req.body.loopSingleKit) : 0,
        kitTransferMiddle: req.body.loopKitTransferMiddle ? Number(req.body.loopKitTransferMiddle) : 0,
        kitTransferLast: req.body.loopKitTransferLast ? Number(req.body.loopKitTransferLast) : 0,
        subLast: req.body.loopSubLast ? Number(req.body.loopSubLast) : 0
      },
      overTtPolicy: req.body.overTtPolicy === "allow" ? "allow" : "forbid",
      maxPeoplePerStation: req.body.maxPeoplePerStation ? Number(req.body.maxPeoplePerStation) : 2,
      standardIncludesLoop: req.body.standardIncludesLoop === "true" || req.body.standardIncludesLoop === "yes"
    });
    res.json(analysis);
  } catch (error) {
    console.error(error);
    res.status(400).json({ error: error.message || "文件分析失败" });
  }
});

app.post("/api/analyze", upload.fields([
  { name: "standard", maxCount: 1 },
  { name: "ebom", maxCount: 1 },
  { name: "mbom", maxCount: 1 },
  { name: "pdf", maxCount: 1 }
]), async (req, res) => {
  try {
    const files = {};
    for (const key of ["standard", "ebom", "mbom", "pdf"]) {
      if (req.files && req.files[key] && req.files[key][0]) {
        const f = req.files[key][0];
        files[key] = { buffer: f.buffer, originalname: Buffer.from(f.originalname, "latin1").toString("utf8") || f.originalname };
      }
    }
    if (!files.standard || !files.ebom || !files.mbom) {
      return res.status(400).json({ error: "必须上传标准工时文件、EBOM文件、MBOM文件；PDF可选。" });
    }
    if (files.pdf) saveTempPdf(files.pdf.buffer);

    const tt = req.body.tt ? Number(req.body.tt) : null;
    const maxStations = req.body.maxStations ? Number(req.body.maxStations) : null;
    const autoStations = req.body.autoStations === "true" || req.body.autoStations === "yes" || req.body.autoStations === "1";
    const maxSubFrames = req.body.maxSubFrames ? Number(req.body.maxSubFrames) : null;
    // V2.6.1 参考打圈工时：50秒/件/岗位（按岗位计，与岗内人数无关）。非强制——未填写一律按0秒，
    // 只有现场确实对半成品打圈时才填写；若勾选“标准工时已含打圈”则在生成侧统一置0查重（见 analyzeProject）。
    const standardIncludesLoop = req.body.standardIncludesLoop === "true" || req.body.standardIncludesLoop === "yes";
    const loopDef = (v) => {
      const n = req.body[v];
      if (n === "" || n == null) return 0;
      return Number(n);
    };
    const loopTimes = {
      singleKit: loopDef("loopSingleKit"),
      kitTransferMiddle: loopDef("loopKitTransferMiddle"),
      kitTransferLast: loopDef("loopKitTransferLast"),
      subLast: loopDef("loopSubLast")
    };
    // 硬阻断：TT 与 最多岗位数 至少提供一项，否则不给静默0岗位结果
    if ((!tt || tt <= 0) && (!maxStations || maxStations <= 0)) {
      return res.status(400).json({ error: "必须填写目标节拍 TT（秒/件）或最多预装岗位数（至少一项），否则无法形成岗位。" });
    }
    const preassemblyMode = req.body.preassemblyMode || "";
    const onlineUltrasonic = req.body.onlineUltrasonic === "yes" || req.body.onlineUltrasonic === "true" || req.body.onlineUltrasonic === "是";
    const onlineUltrasonicMaxGroupsPerConfig = req.body.onlineUltrasonicMaxGroupsPerConfig ? Number(req.body.onlineUltrasonicMaxGroupsPerConfig) : null;
    const onlineUltrasonicMaxTotalGroups = req.body.onlineUltrasonicMaxTotalGroups ? Number(req.body.onlineUltrasonicMaxTotalGroups) : null;
    const noOnlineUltrasonicSplices = parseList(req.body.noOnlineUltrasonicSplices);
    const forcedOfflineHousings = parseList(req.body.forcedOfflineHousings);
    const sameStationHousings = parseList(req.body.sameStationHousings);
    const sameStationOverTtMode = req.body.sameStationOverTtMode === "best-rate" ? "best-rate" : "force-same";
    const overTtPolicy = req.body.overTtPolicy === "allow" ? "allow" : "forbid";
    const maxPeoplePerStation = req.body.maxPeoplePerStation ? Number(req.body.maxPeoplePerStation) : 2;
    let sameStationGroups = [];
    try { sameStationGroups = JSON.parse(req.body.sameStationGroups || "[]"); } catch {}
    let grommetStations = [];
    try { grommetStations = JSON.parse(req.body.grommetStations || "[]"); } catch {}
    const rawRegions = req.body.regions;
    const regions = rawRegions
      ? (Array.isArray(rawRegions) ? rawRegions.slice(0, 1) : [rawRegions])
      : [];
    const result = await analyzeProject(files, {
      tt,
      maxStations,
      autoStations,
      maxSubFrames,
      loopTimes,
      standardIncludesLoop,
      preassemblyMode,
      regions,
      onlineUltrasonic,
      onlineUltrasonicMaxGroupsPerConfig,
      onlineUltrasonicMaxTotalGroups,
      noOnlineUltrasonicSplices,
      forcedOfflineHousings,
      sameStationHousings,
      sameStationGroups,
      sameStationOverTtMode,
      overTtPolicy,
      maxPeoplePerStation,
      grommetStations
    });
    // 本地生成不写入版本历史（版本记录只记录作者发布）；仅保存到本地生成历史。
    saveGenHistory(files, result, regions, (result.plan && result.plan.modeLabel) || "");
    res.json(result);
  } catch (error) {
    console.error(error);
    res.status(400).json({ error: error.message || "解析失败" });
  }
});

app.get("/api/version", (req, res) => {
  res.json(getVersionHistory());
});

app.post("/api/version/record", (req, res) => {
  try {
    const { action, detail, modules } = req.body || {};
    const result = addRecord({ action, detail, modules });
    res.json(result);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/backend/auth", (req, res) => {
  try {
    const { password } = req.body || {};
    if (!verifyBackendPassword(password)) {
      return res.status(401).json({ ok: false, error: "密码错误" });
    }
    res.json({ ok: true, data: buildBackendData() });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/export", (req, res) => {
  try {
    const result = req.body;
    if (!result || !result.generatedAt) {
      return res.status(400).json({ error: "缺少生成结果，请先执行分析。" });
    }
    const buffer = buildWorkbook(result);
    const stamp = new Date().toISOString().slice(0, 10);
    const filename = `preassembly_package_${stamp}.xlsx`;
    const displayName = encodeURIComponent(`预装工艺结构化数据包_${stamp}.xlsx`);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"; filename*=UTF-8''${displayName}`);
    res.send(buffer);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: error.message || "导出失败" });
  }
});

app.post("/api/export/process", (req, res) => {
  try {
    const result = req.body;
    if (!result || !result.generatedAt) return res.status(400).json({ error: "缺少生成结果，请先执行分析。" });
    const buffer = buildProcessWorkbook(result);
    const stamp = new Date().toISOString().slice(0, 10);
    const displayName = encodeURIComponent(`预装工艺岗位表_${stamp}.xlsx`);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="preassembly_stations_${stamp}.xlsx"; filename*=UTF-8''${displayName}`);
    res.send(buffer);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: error.message || "导出失败" });
  }
});

app.post("/api/export/review", (req, res) => {
  try {
    const result = req.body;
    if (!result || !result.generatedAt) return res.status(400).json({ error: "缺少生成结果，请先执行分析。" });
    const buffer = buildReviewWorkbook(result);
    const stamp = new Date().toISOString().slice(0, 10);
    const displayName = encodeURIComponent(`预装工艺人工核查表_${stamp}.xlsx`);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="preassembly_review_${stamp}.xlsx"; filename*=UTF-8''${displayName}`);
    res.send(buffer);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: error.message || "导出失败" });
  }
});

app.get("/api/ai/multimodal", (req, res) => {
  res.json({
    available: true,
    required: false,
    reason: "多模态AI用于读取PDF图纸/图片，并返回结构化数据（护套、分支、保护件、SP/SC点、尺寸、备注等）。该功能为可选项，PDF图纸上传为非必选项。"
  });
});

app.post("/api/ai/analyze", upload.single("file"), (req, res) => {
  try {
    const apiUrl = String(req.body.apiUrl || "").trim();
    const apiKey = String(req.body.apiKey || "").trim();
    const model = String(req.body.model || "").trim();
    if (!apiUrl || !apiKey || !model) {
      return res.status(400).json({ error: "请填写API地址、API Key和模型名称" });
    }
    if (!/^https?:\/\//i.test(apiUrl)) {
      return res.status(400).json({ error: "API地址必须以 http:// 或 https:// 开头" });
    }

    const jobId = "ai-" + (aiJobSeq++);
    const job = {
      id: jobId,
      status: "running",
      createdAt: new Date().toISOString(),
      result: null,
      error: null,
      progress: null
    };
    aiJobs.set(jobId, job);

    const fileBuffer = req.file ? req.file.buffer : null;
    const fileMime = req.file ? req.file.mimetype : "";
    const fileName = req.file ? req.file.originalname : "";

    (async () => {
      try {
        const endpoint = /\/chat\/completions$/.test(apiUrl) ? apiUrl : apiUrl.replace(/\/+$/, "") + "/chat/completions";

        // 1. 确定输入：上传的PDF/图片，或主页上传的PDF（无文件时兜底）
        let source = "无文件";
        let batches = []; // [{ buffer, mime }]
        let pages = 0;
        if (fileBuffer) {
          const isPdf = /\.pdf$/i.test(fileName) || (fileMime || "").includes("pdf");
          if (isPdf) {
            source = "PDF（已转换为图片）";
            const imgs = await renderPdfPages(fileBuffer);
            pages = imgs.length;
            batches = splitImagesIntoBatches(imgs).map((b) => b.map((buf) => ({ buffer: buf, mime: "image/png" })));
          } else {
            source = "图片";
            pages = 1;
            const mime = fileMime || "image/png";
            batches = [[{ buffer: fileBuffer, mime }]];
          }
        } else if (fs.existsSync(TEMP_PDF)) {
          source = "主页上传的PDF（已转换为图片）";
          const imgs = await renderPdfPages(fs.readFileSync(TEMP_PDF));
          pages = imgs.length;
          batches = splitImagesIntoBatches(imgs).map((b) => b.map((buf) => ({ buffer: buf, mime: "image/png" })));
        }

        // 2. 分批调用视觉模型：模型通过 needMore 决定是否继续识别后续图纸
        const parsedBatches = [];
        if (!batches.length) {
          const parsed = await callAiModel(endpoint, apiKey, model, [
            { type: "text", text: "未提供图纸文件，请根据已有信息输出线束工艺识别JSON。" }
          ]);
          job.status = "done";
          job.result = { source, pages: 0, batches: [], merged: parsed };
          return;
        }
        for (let i = 0; i < batches.length; i++) {
          const batch = batches[i];
          job.progress = { done: i + 1, total: batches.length, pages };
          const text = `这是汽车线束工艺图纸识别任务，当前为第${i + 1}/${batches.length}批，共${batch.length}张图纸图片。请识别其中的护套、SP/SC压接点、分支、保护件、尺寸和备注，只输出JSON；如果识别信息不完整或需要继续查看后续图纸，请返回 needMore: true。`;
          const content = [
            { type: "text", text },
            ...batch.map((b) => imageContentBlock(b.buffer, b.mime))
          ];
          const parsed = await callAiModel(endpoint, apiKey, model, content);
          parsedBatches.push(parsed);
          const needMore = parsed && typeof parsed === "object" && parsed.needMore === true;
          if (needMore === false && i < batches.length - 1) {
            // 模型认为无需继续，提前结束（剩余批次不再发送）
            break;
          }
        }
        job.status = "done";
        job.result = { source, pages, batches: parsedBatches, merged: mergeAiBatches(parsedBatches) };
      } catch (error) {
        console.error(error);
        job.status = "error";
        job.error = error.message || "AI识别失败";
      }
    })();

    res.json({ ok: true, jobId });
  } catch (error) {
    res.status(400).json({ error: error.message || "提交失败" });
  }
});

app.get("/api/ai/jobs", (req, res) => {
  const list = [...aiJobs.values()]
    .sort((a, b) => b.createdAt.localeCompare(a.createdAt))
    .slice(0, 20)
    .map((j) => ({ id: j.id, status: j.status, createdAt: j.createdAt, error: j.error || null, progress: j.progress || null }));
  res.json({ jobs: list });
});

app.get("/api/ai/jobs/:id", (req, res) => {
  const job = aiJobs.get(req.params.id);
  if (!job) return res.status(404).json({ error: "任务不存在" });
  res.json(job);
});

// 检查更新：读取更新源(data/update_config.json 的 updateUrl 指向 Gitee 上的 version.json)并与当前版本比对
const UPDATE_CONFIG_FILE = path.join(__dirname, "data", "update_config.json");
function readUpdateConfig() {
  try { return JSON.parse(fs.readFileSync(UPDATE_CONFIG_FILE, "utf8").replace(/^\uFEFF/, "")); } catch { return {}; }
}
function parseVersionNum(v) {
  const p = String(v || "").replace(/^v/i, "").split(".").map(Number);
  return ((p[0] || 0) * 10000) + ((p[1] || 0) * 100) + (p[2] || 0);
}
app.get("/api/update/check", async (req, res) => {
  try {
    const cfg = readUpdateConfig();
    const url = String(cfg.updateUrl || "").trim();
    let current = "0.0.0";
    try { current = require("./package.json").version; } catch {}
    if (!url) {
      return res.json({ ok: false, configured: false, current, message: "未配置更新源：请在 data/update_config.json 的 updateUrl 填入 Gitee 上的 version.json 地址。" });
    }
    const r = await fetch(url);
    if (!r.ok) return res.json({ ok: false, configured: true, current, error: `更新源请求失败（HTTP ${r.status}）` });
    const info = await r.json();
    const latest = String(info.version || "").trim();
    const hasUpdate = !!latest && parseVersionNum(latest) > parseVersionNum(current);
    res.json({
      ok: true,
      configured: true,
      current,
      latest,
      hasUpdate,
      notes: info.notes || "",
      downloadUrl: info.downloadUrl || "",
      sha256: info.sha256 || ""
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: error.message || "检查更新失败" });
  }
});

// ============ 一键更新：下载待更新包 + 重启后自动应用 ============
const PENDING_DIR = path.join(__dirname, "data", "pending_update");
const APP_ROOT = __dirname; // resources/app（server.js 所在目录即应用根）

// 把 src 目录内容整体覆盖到 dest（不删除目标中多余的既有文件）
function copyOverlay(src, dest) {
  for (const ent of fs.readdirSync(src, { withFileTypes: true })) {
    const sp = path.join(src, ent.name);
    const dp = path.join(dest, ent.name);
    if (ent.isDirectory()) {
      fs.mkdirSync(dp, { recursive: true });
      copyOverlay(sp, dp);
    } else if (ent.isFile()) {
      fs.mkdirSync(path.dirname(dp), { recursive: true });
      fs.copyFileSync(sp, dp);
    }
  }
}

// 启动时检查是否存在“待应用更新”，校验后覆盖应用源码（重启后生效）
function robustRemove(dir) {
  for (let i = 0; i < 6; i++) {
    try { fs.rmSync(dir, { recursive: true, force: true }); return; } catch (e) { /* Windows 偶发句柄未释放，重试 */ }
  }
}
function applyPendingUpdate() {
  try {
    const zipPath = path.join(PENDING_DIR, "update.zip");
    const metaPath = path.join(PENDING_DIR, "pending.json");
    if (!fs.existsSync(zipPath) || !fs.existsSync(metaPath)) return;
    const meta = JSON.parse(fs.readFileSync(metaPath, "utf8").replace(/^\uFEFF/, ""));
    const buf = fs.readFileSync(zipPath);
    if (meta.sha256) {
      const hash = crypto.createHash("sha256").update(buf).digest("hex");
      if (hash !== meta.sha256) {
        console.error(`待更新包 sha256 校验失败（期望 ${meta.sha256}，实际 ${hash}），已丢弃`);
        robustRemove(PENDING_DIR);
        return;
      }
    }
    const tmp = path.join(PENDING_DIR, "extract");
    robustRemove(tmp);
    fs.mkdirSync(tmp, { recursive: true });
    execSync(`tar -xf "${zipPath}" -C "${tmp}"`, { stdio: "ignore" });
    copyOverlay(tmp, APP_ROOT);
    robustRemove(PENDING_DIR);
    console.log(`已应用待更新（${meta.targetVersion || "未知版本"}），重启后将完整生效`);
  } catch (e) {
    console.error("应用待更新失败:", e.message);
  }
}

// 下载更新包到待更新目录（本进程继续跑旧代码；下次重启时 applyPendingUpdate 覆盖）
app.post("/api/update/download", async (req, res) => {
  try {
    const { downloadUrl, sha256, targetVersion } = req.body || {};
    if (!downloadUrl) return res.status(400).json({ error: "缺少 downloadUrl" });
    const r = await fetch(downloadUrl);
    if (!r.ok) return res.status(400).json({ error: `下载更新包失败（HTTP ${r.status}）` });
    const buf = Buffer.from(await r.arrayBuffer());
    if (sha256) {
      const hash = crypto.createHash("sha256").update(buf).digest("hex");
      if (hash !== sha256) return res.status(400).json({ error: "更新包 sha256 校验失败" });
    }
    fs.mkdirSync(PENDING_DIR, { recursive: true });
    fs.writeFileSync(path.join(PENDING_DIR, "update.zip"), buf);
    fs.writeFileSync(path.join(PENDING_DIR, "pending.json"), JSON.stringify({ targetVersion: targetVersion || "", sha256: sha256 || "" }, null, 2), "utf8");
    res.json({ ok: true, message: "更新包已下载，重启软件后将自动完成更新" });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: error.message || "下载更新失败" });
  }
});

app.get("/api/templates/:type", (req, res) => {
  try {
    const type = req.params.type;
    const buffer = buildTemplate(type);
    const names = {
      standard: "标准工时文件模板.xlsx",
      ebom: "EBOM文件模板.xlsx",
      mbom: "MBOM-Cutting模板.xlsx"
    };
    const displayName = encodeURIComponent(names[type] || "模板.xlsx");
    const asciiName = "template_" + type + ".xlsx";
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${asciiName}"; filename*=UTF-8''${displayName}`);
    res.send(buffer);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: error.message || "模板生成失败" });
  }
});

// ==================== 本地生成历史 ====================
app.get("/api/gen-history", (req, res) => {
  try {
    res.json({ entries: readGenHistoryIndex() });
  } catch (error) {
    res.status(500).json({ error: error.message || "读取生成历史失败" });
  }
});

app.get("/api/gen-history/:id", (req, res) => {
  try {
    const id = req.params.id;
    const entry = readGenHistoryIndex().find((e) => e.id === id);
    if (!entry) return res.status(404).json({ error: "该记录不存在" });
    const result = JSON.parse(fs.readFileSync(path.join(GEN_HISTORY_DIR, id, "result.json"), "utf8"));
    res.json({ meta: entry, result });
  } catch (error) {
    res.status(500).json({ error: error.message || "打开生成历史失败" });
  }
});

app.get("/api/gen-history/:id/files/:field", (req, res) => {
  try {
    const id = req.params.id;
    const field = req.params.field;
    if (!["standard", "ebom", "mbom", "pdf"].includes(field)) return res.status(400).json({ error: "未知字段" });
    const entry = readGenHistoryIndex().find((e) => e.id === id);
    const p = path.join(GEN_HISTORY_DIR, id, field);
    if (!fs.existsSync(p)) return res.status(404).json({ error: "无该文件（或未上传）" });
    const name = (entry && entry.fileNames && entry.fileNames[field]) || field;
    res.setHeader("Content-Type", "application/octet-stream");
    res.setHeader("Content-Disposition", `attachment; filename="history_${field}"; filename*=UTF-8''${encodeURIComponent(name)}`);
    res.send(fs.readFileSync(p));
  } catch (error) {
    res.status(500).json({ error: error.message || "下载失败" });
  }
});

app.delete("/api/gen-history/:id", (req, res) => {
  try {
    res.json({ ok: deleteGenHistory(req.params.id) });
  } catch (error) {
    res.status(500).json({ error: error.message || "删除失败" });
  }
});

app.use(express.static(path.join(__dirname, "public")));

const PORT = process.env.PORT ? Number(process.env.PORT) : getConfiguredPort();
app.listen(PORT, () => {
  console.log(`预装工艺文件生成器已启动： http://localhost:${PORT}`);
  // 若存在“待应用更新”，覆盖应用（重启后完整生效）
  applyPendingUpdate();
});
