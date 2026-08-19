"use strict";
// 发布更新脚本：生成只含源码的 update.zip + version.json（含sha256、下载直链）
// 用法: node publish-update.js <新版本号x.y.z> [更新说明]
// 例如: node publish-update.js 1.0.24 "新增一键更新"
// 产物: 发布输出/update.zip、发布输出/version.json （传到 Gitee 仓库根目录即可）
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { execSync } = require("child_process");

const ROOT = __dirname;
const OUT_DIR = path.join(ROOT, "发布输出");

const args = process.argv.slice(2);
const newVersion = String(args[0] || "").trim();
// 更新说明优先读“发布说明.txt”(UTF-8)，避免命令行中文在 Windows 下乱码；命令行第二参可作备选
let notes = (args[1] || "").trim();
const notesFile = path.join(ROOT, "发布说明.txt");
if (fs.existsSync(notesFile)) notes = fs.readFileSync(notesFile, "utf8").trim();
if (!/^\d+\.\d+\.\d+$/.test(newVersion)) {
  console.error("用法: node publish-update.js <新版本号x.y.z> [更新说明]");
  console.error("例如: node publish-update.js 1.0.24 \"新增一键更新\"");
  process.exit(1);
}

const cfg = JSON.parse(fs.readFileSync(path.join(ROOT, "data", "update_config.json"), "utf8").replace(/^\uFEFF/, ""));
const owner = String(cfg.owner || "").trim();
const repo = String(cfg.repo || "").trim();
const branch = String(cfg.branch || "master").trim();
if (!owner || !repo) {
  console.error("请在 data/update_config.json 里的 owner/repo 填入 Gitee 仓库信息");
  process.exit(1);
}

// ---- 先把发布版本号写入 package.json —— 必须在“打包”之前完成：
//      否则打进 update.zip 的 package.json 仍是旧版本号，用户更新后用上新代码，
//      但 /api/update/check 读到旧版本，会一直提示“有新版本、请更新”，形成死循环。 ----
{
  const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, "package.json"), "utf8"));
  pkg.version = newVersion;
  fs.writeFileSync(path.join(ROOT, "package.json"), JSON.stringify(pkg, null, 2) + "\n", "utf8");
}

// ---- 收集要发布的源码文件（排除 node_modules/dist/敏感数据/大文件） ----
const tmpRoot = path.join(ROOT, "_publish_tmp");
const tmpUpdate = path.join(tmpRoot, "update");
fs.rmSync(tmpRoot, { recursive: true, force: true });
fs.mkdirSync(tmpUpdate, { recursive: true });

function copyDir(src, rel) {
  for (const ent of fs.readdirSync(src, { withFileTypes: true })) {
    if (ent.name === "node_modules" || ent.name === "dist" || ent.name === "data" || ent.name.startsWith("_") || ent.name === "发布输出") continue;
    const sp = path.join(src, ent.name);
    const rp = path.join(rel, ent.name);
    if (ent.isDirectory()) copyDir(sp, rp);
    else if (ent.isFile()) {
      fs.mkdirSync(path.dirname(path.join(tmpUpdate, rp)), { recursive: true });
      fs.copyFileSync(sp, path.join(tmpUpdate, rp));
    }
  }
}
function copyFile(sp, rel) {
  fs.mkdirSync(path.dirname(path.join(tmpUpdate, rel)), { recursive: true });
  fs.copyFileSync(sp, path.join(tmpUpdate, rel));
}

// 根文件 + src + public
for (const item of ["server.js", "electron-main.js", "preload.js", "package.json", "package-lock.json", "LICENSE", "README.md", "src", "public"]) {
  const sp = path.join(ROOT, item);
  if (!fs.existsSync(sp)) continue;
  if (fs.statSync(sp).isDirectory()) copyDir(sp, item);
  else copyFile(sp, item);
}
// data 只带版本记录与更新配置（不含后台密码/端口/暂存/临时PDF）
for (const df of ["data/version_history.json", "data/update_config.json"]) {
  const sp = path.join(ROOT, df);
  if (fs.existsSync(sp)) copyFile(sp, df);
}

// ---- 打 zip（Windows 自带 tar，-a 按 .zip 输出） ----
fs.mkdirSync(OUT_DIR, { recursive: true });
const zipPath = path.join(OUT_DIR, "update.zip");
if (fs.existsSync(zipPath)) fs.rmSync(zipPath);
execSync(`tar -a -c -f "${zipPath}" -C "${tmpUpdate}" .`, { stdio: "inherit" });

const zipBuf = fs.readFileSync(zipPath);
const sha256 = crypto.createHash("sha256").update(zipBuf).digest("hex");
const pkgUrl = `https://gitee.com/${owner}/${repo}/raw/${branch}/update.zip`;
fs.writeFileSync(path.join(OUT_DIR, "version.json"), JSON.stringify({ version: newVersion, notes, downloadUrl: pkgUrl, sha256 }, null, 2) + "\n", "utf8");

fs.rmSync(tmpRoot, { recursive: true, force: true });

console.log("\n发布包已生成：");
console.log("  update.zip   (" + Math.round(zipBuf.length / 1024) + " KB)  sha256=" + sha256);
console.log("  version.json");
console.log("\n请把 发布输出/ 下两个文件传到 Gitee 仓库 " + branch + " 分支根目录：");
console.log("  - update.zip   → 根目录 update.zip");
console.log("  - version.json → 根目录 version.json");
console.log("软件“检查更新”读取：https://gitee.com/" + owner + "/" + repo + "/raw/" + branch + "/version.json");
console.log("package.json version 已更新为 " + newVersion);
