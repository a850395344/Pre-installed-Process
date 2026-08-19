"use strict";

const files = {};
let result = null;
let fileAnalysis = null;
let activeTab = "";
const renderLimits = new Map();

const $ = (sel) => document.querySelector(sel);

function toast(text, type = "success") {
  const el = $("#toast");
  el.textContent = text;
  el.className = "toast " + type;
  el.hidden = false;
  clearTimeout(el._timer);
  el._timer = setTimeout(() => { el.hidden = true; }, 2600);
}

function humanNumber(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n.toLocaleString("zh-CN") : String(value || "");
}

// 岗位数显示：若按TT自动扩岗，展示"最终值（按TT由原设置扩岗）"
function stationCountText(plan) {
  if (!plan || plan.maxStations == null) return "未设置";
  const req = plan.maxStationsRequested;
  if (req != null && req !== plan.maxStations) return `${plan.maxStations}（按TT由${req}扩岗）`;
  return String(plan.maxStations);
}

function addGrommetRow(name = "", housings = "", time = "") {
  const tbody = $("#grommetRows");
  const tr = document.createElement("tr");
  tr.innerHTML = `
    <td><input class="grommet-name" value="${name}" placeholder="例如 G01"></td>
    <td><input class="grommet-housings" value="${housings}" placeholder="例如 W25LH1/B601E1"></td>
    <td><input class="grommet-time" type="number" min="0" step="1" value="${time}" placeholder="秒"></td>
    <td><button class="btn" type="button">删除</button></td>
  `;
  tr.querySelector("button").addEventListener("click", () => tr.remove());
  tbody.appendChild(tr);
}

function collectGrommetRows() {
  return Array.from(document.querySelectorAll("#grommetRows tr")).map((tr) => ({
    name: tr.querySelector(".grommet-name").value.trim(),
    housings: tr.querySelector(".grommet-housings").value.trim(),
    time: Number(tr.querySelector(".grommet-time").value) || 0
  })).filter((g) => g.name && g.housings);
}

function addSameStationGroupRow(name = "", housings = "", mode = "pure-kit") {
  const tbody = $("#sameStationGroupRows");
  const tr = document.createElement("tr");
  const modeOptions = [
    ["pure-kit", "纯KIT岗位"],
    ["kit-transfer", "KIT传递岗"],
    ["sub", "SUB岗位"]
  ].map(([v, l]) => `<option value="${v}"${mode === v ? " selected" : ""}>${l}</option>`).join("");
  tr.innerHTML = `
    <td><input class="ssg-name" value="${name}" placeholder="例如 左前门接口"></td>
    <td><input class="ssg-housings" value="${housings}" placeholder="例如 W25LH1/B601E1"></td>
    <td><select class="ssg-mode">${modeOptions}</select></td>
    <td><button class="btn" type="button">删除</button></td>
  `;
  tr.querySelector("button").addEventListener("click", () => tr.remove());
  tbody.appendChild(tr);
}

function collectSameStationGroups() {
  return Array.from(document.querySelectorAll("#sameStationGroupRows tr")).map((tr) => ({
    name: tr.querySelector(".ssg-name").value.trim(),
    housings: tr.querySelector(".ssg-housings").value.trim(),
    mode: tr.querySelector(".ssg-mode").value
  })).filter((g) => g.housings);
}

function resetAnalysisState() {
  if (!fileAnalysis) return;
  fileAnalysis = null;
  result = null;
  $("#analysisPanel").hidden = true;
  $("#analyzeBtn").disabled = true;
  $("#result").hidden = true;
  $("#exportBtn").disabled = true;
  $("#exportProcessBtn").disabled = true;
  $("#exportReviewBtn").disabled = true;
  $("#noOnlineUltrasonicSplices").value = "";
  $("#forcedOfflineHousings").value = "";
}

// ==================== 必填项校验 + 红框闪烁 + 弹窗 ====================
const REQUIRED_MAP = [
  { key: "standard", label: "标准工时文件", check: () => !!files.standard, sel: '.file-drop[data-field="standard"]' },
  { key: "ebom", label: "EBOM文件", check: () => !!files.ebom, sel: '.file-drop[data-field="ebom"]' },
  { key: "mbom", label: "MBOM/Cutting文件", check: () => !!files.mbom, sel: '.file-drop[data-field="mbom"]' },
  { key: "tt", label: "目标节拍 TT（秒/件）", check: () => $("#tt").value.trim() !== "", sel: "#tt" },
  { key: "preassemblyMode", label: "预装模式", check: () => $("#preassemblyMode").value !== "", sel: "#preassemblyMode" }
];
function validateRequiredFields() {
  return REQUIRED_MAP.filter((r) => !r.check());
}
function clearRequiredFlash(sel) {
  const el = document.querySelector(sel);
  if (el) el.classList.remove("required-flash");
}
function flashMissing(missing) {
  missing.forEach((m) => {
    const el = document.querySelector(m.sel);
    if (el) {
      el.classList.remove("required-flash");
      void el.offsetWidth; // 重触发动画
      el.classList.add("required-flash");
    }
  });
}
// 页面内模态框：列出缺失项，点击可定位到对应输入框
function showRequiredModal(missing) {
  const overlay = document.createElement("div");
  overlay.className = "req-modal-overlay";
  const box = document.createElement("div");
  box.className = "req-modal";
  const h = document.createElement("h3");
  h.textContent = "请先完善必填项";
  const p = document.createElement("p");
  p.textContent = "点击下方缺失项可直接跳到对应输入位置：";
  const ul = document.createElement("ul");
  ul.className = "req-missing-list";
  for (const m of missing) {
    const li = document.createElement("li");
    li.innerHTML = '<span class="req-star">*</span>' + m.label;
    li.dataset.sel = m.sel;
    li.addEventListener("click", () => {
      const el = document.querySelector(m.sel);
      if (el) {
        el.scrollIntoView({ behavior: "smooth", block: "center" });
        el.classList.remove("required-flash");
        void el.offsetWidth;
        el.classList.add("required-flash");
      }
      overlay.remove();
    });
    ul.appendChild(li);
  }
  const btn = document.createElement("button");
  btn.className = "btn primary";
  btn.type = "button";
  btn.textContent = "知道了";
  btn.addEventListener("click", () => overlay.remove());
  overlay.addEventListener("click", (e) => { if (e.target === overlay) overlay.remove(); });
  box.appendChild(h); box.appendChild(p); box.appendChild(ul); box.appendChild(btn);
  overlay.appendChild(box);
  document.body.appendChild(overlay);
}
// 用守卫包裹主流程按钮：缺失必填时弹窗+闪烁，不执行
function guardRequired(fn) {
  return (...args) => {
    const missing = validateRequiredFields();
    if (missing.length) {
      flashMissing(missing);
      showRequiredModal(missing);
      return;
    }
    return fn(...args);
  };
}

// ---------- 检查更新 ----------
function showMessageModal(title, bodyHtml) {
  const overlay = document.createElement("div");
  overlay.className = "req-modal-overlay";
  const box = document.createElement("div");
  box.className = "req-modal";
  const h = document.createElement("h3");
  h.textContent = title;
  const b = document.createElement("div");
  b.style = "line-height:1.7; color:var(--text-2); font-size:13px;";
  b.innerHTML = bodyHtml;
  const btn = document.createElement("button");
  btn.className = "btn primary";
  btn.type = "button";
  btn.textContent = "关闭";
  btn.addEventListener("click", () => overlay.remove());
  overlay.addEventListener("click", (e) => { if (e.target === overlay) overlay.remove(); });
  box.appendChild(h); box.appendChild(b); box.appendChild(btn);
  overlay.appendChild(box);
  document.body.appendChild(overlay);
  return overlay;
}
// ---------- 检查更新 / 下载进度 / 重启 ----------
let updateChecking = false;
let updateBtnEl = null;
const IS_ELECTRON = !!(window.dsh && window.dsh.isElectron);
const resetUpdateBusy = () => {
  updateChecking = false;
  if (updateBtnEl) { updateBtnEl.disabled = false; updateBtnEl = null; }
};

// 统一更新弹窗骨架：渐变头部徽标 + 卡片主体
function createUpdateModal(title, icon) {
  const overlay = document.createElement("div");
  overlay.className = "upd-overlay";
  const card = document.createElement("div");
  card.className = "upd-modal";
  const head = document.createElement("div");
  head.className = "upd-head";
  const ic = document.createElement("span"); ic.className = "upd-icon"; ic.textContent = icon || "🚀";
  const tt = document.createElement("span"); tt.className = "upd-title"; tt.textContent = title;
  head.appendChild(ic); head.appendChild(tt);
  const body = document.createElement("div");
  body.className = "upd-body";
  card.appendChild(head); card.appendChild(body);
  overlay.appendChild(card);
  const close = () => overlay.remove();
  overlay.addEventListener("click", (e) => { if (e.target === overlay) close(); });
  document.body.appendChild(overlay);
  return { overlay, card, head, ic, tt, body, close };
}
// 版本对比块
function updVersionsHtml(cur, next) {
  return `<div class="upd-vers">
    <div class="upd-ver"><b>当前版本</b><span>${escapeHtml(cur)}</span></div>
    <div class="upd-sep">→</div>
    <div class="upd-ver next"><b>最新版本</b><span>${escapeHtml(next)}</span></div>
  </div>`;
}

// 统一“下载→进度→成功→重启提示”流程
async function beginUpdateDownload(d, closeOverlay) {
  let resp;
  try {
    resp = await fetch("/api/update/download", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ downloadUrl: d.downloadUrl, sha256: d.sha256 || "", targetVersion: d.latest })
    });
  } catch (e2) {
    if (closeOverlay) closeOverlay();
    resetUpdateBusy();
    showMessageModal("更新", "下载启动失败：" + escapeHtml(e2 && e2.message ? e2.message : e2));
    return;
  }
  const data = await resp.json();
  if (!data.ok || !data.jobId) {
    if (closeOverlay) closeOverlay();
    resetUpdateBusy();
    showMessageModal("更新", (data.error || "下载更新失败"));
    return;
  }
  if (closeOverlay) closeOverlay();

  const m = createUpdateModal("正在下载更新包", "⬇️");
  const body = m.body;
  const track = document.createElement("div"); track.className = "upd-track";
  const bar = document.createElement("div"); bar.className = "upd-bar";
  track.appendChild(bar);
  const meta = document.createElement("div"); meta.className = "upd-meta";
  const info = document.createElement("span"); info.className = "upd-info-text"; info.textContent = "正在连接更新源…";
  const pct = document.createElement("span"); pct.className = "upd-pct"; pct.textContent = "0%";
  meta.appendChild(info); meta.appendChild(pct);
  const prog = document.createElement("div"); prog.className = "upd-progress";
  prog.appendChild(track); prog.appendChild(meta);
  body.appendChild(prog);

  const finish = () => { resetUpdateBusy(); };
  const setDone = () => { m.tt.textContent = "下载完成"; m.ic.textContent = "✅"; };

  const poll = async () => {
    try {
      const jresp = await fetch("/api/update/jobs/" + data.jobId);
      const j = await jresp.json();
      if (!j || j.status === "error") {
        bar.style.width = "0%";
        pct.textContent = "失败"; pct.style.color = "#dc2626";
        m.ic.textContent = "⚠️"; m.tt.textContent = "下载失败";
        info.textContent = (j && j.error) || "下载失败";
        const ok = document.createElement("button"); ok.className = "btn primary"; ok.type = "button"; ok.textContent = "关闭";
        ok.onclick = () => { m.close(); finish(); };
        const row = document.createElement("div"); row.className = "upd-actions"; row.appendChild(ok);
        body.appendChild(row);
        return;
      }
      if (j.progress != null) { bar.style.width = j.progress + "%"; pct.textContent = Math.round(j.progress) + "%"; }
      if (j.downloaded != null && j.total) info.textContent = `已下载 ${(j.downloaded / 1048576).toFixed(1)}MB / ${(j.total / 1048576).toFixed(1)}MB`;
      else if (j.downloaded != null) info.textContent = `已下载 ${(j.downloaded / 1048576).toFixed(1)}MB（等待校验…）`;
      if (j.status === "done") {
        bar.style.width = "100%"; pct.textContent = "100%";
        setDone();
        const ok = document.createElement("div"); ok.className = "upd-success-msg"; ok.textContent = "更新包已下载并校验完成（sha256 通过）✅";
        const note = document.createElement("div"); note.className = "upd-info-row";
        note.textContent = IS_ELECTRON
          ? "点击“现在重启”将自动重启软件并应用更新，重启后版本显示为新版、不再提示更新。"
          : "独立服务版：请关闭并重新打开软件以应用更新（仅刷新页面不会生效），重开后显示新版、不再提示更新。";
        const later = document.createElement("button"); later.className = "btn"; later.type = "button"; later.textContent = "稍后";
        const now = document.createElement("button"); now.className = "btn primary"; now.type = "button"; now.textContent = IS_ELECTRON ? "现在重启" : "我已重新打开";
        const row = document.createElement("div"); row.className = "upd-actions";
        row.appendChild(later); row.appendChild(now);
        later.onclick = () => { m.close(); finish(); };
        now.onclick = () => {
          if (IS_ELECTRON && window.dsh.relaunch) {
            try { window.dsh.relaunch(); } catch (e3) { showMessageModal("更新", "重启失败：" + escapeHtml(e3 && e3.message ? e3.message : e3)); }
          } else {
            try { fetch("/api/update/restart", { method: "POST", headers: { "Content-Type": "application/json" }, body: "{}" }); } catch {}
            info.textContent = "请关闭本窗口并重新打开软件完成更新；重开后页面将显示新版。";
            later.remove(); now.remove();
          }
        };
        body.appendChild(ok); body.appendChild(note); body.appendChild(row);
        finish();
        return;
      }
      setTimeout(poll, 700);
    } catch (err) {
      info.textContent = "进度查询失败：" + escapeHtml(err.message || err);
      finish();
    }
  };
  poll();
}

async function checkUpdate() {
  if (updateChecking) return;
  updateChecking = true;
  const btn = $("#checkUpdateBtn");
  updateBtnEl = btn;
  if (btn) btn.disabled = true;
  try {
    const resp = await fetch("/api/update/check");
    const d = await resp.json();
    if (!d.configured) {
      resetUpdateBusy();
      const m = createUpdateModal("检查更新", "🔍");
      const info = document.createElement("div"); info.className = "upd-info-row";
      info.innerHTML = escapeHtml(d.message) + `<div style="margin-top:8px">更新源地址形如：<code>https://gitee.com/你的用户名/仓库/raw/master/version.json</code></div>`;
      const b = document.createElement("button"); b.className = "btn primary"; b.type = "button"; b.textContent = "关闭"; b.onclick = () => m.close();
      const row = document.createElement("div"); row.className = "upd-actions"; row.appendChild(b);
      m.body.appendChild(info); m.body.appendChild(row);
      return;
    }
    if (!d.hasUpdate) {
      resetUpdateBusy();
      const m = createUpdateModal("检查更新", "✅");
      const v = document.createElement("div"); v.className = "upd-vers";
      v.innerHTML = `<div class="upd-ver"><b>当前版本</b><span>${escapeHtml(d.current)}</span></div>`;
      const info = document.createElement("div"); info.className = "upd-info-row"; info.textContent = "当前已是最新版本，无需更新。";
      const b = document.createElement("button"); b.className = "btn primary"; b.type = "button"; b.textContent = "关闭"; b.onclick = () => m.close();
      const row = document.createElement("div"); row.className = "upd-actions"; row.appendChild(b);
      m.body.appendChild(v); m.body.appendChild(info); m.body.appendChild(row);
      return;
    }
    const m = createUpdateModal("发现新版本", "🚀");
    const v = document.createElement("div"); v.innerHTML = updVersionsHtml(d.current, d.latest);
    m.body.appendChild(v);
    if (d.notes) { const n = document.createElement("div"); n.className = "upd-notes"; n.textContent = d.notes; m.body.appendChild(n); }
    const row = document.createElement("div"); row.className = "upd-actions";
    if (d.downloadUrl) { const a = document.createElement("a"); a.className = "btn"; a.href = d.downloadUrl; a.target = "_blank"; a.rel = "noopener"; a.textContent = "去下载页"; row.appendChild(a); }
    const dl = document.createElement("button"); dl.className = "btn primary"; dl.type = "button"; dl.textContent = "立即下载并更新"; dl.onclick = () => { m.close(); beginUpdateDownload(d); };
    row.appendChild(dl);
    m.body.appendChild(row);
  } catch (e) {
    resetUpdateBusy();
    showMessageModal("检查更新", "检查更新失败：" + escapeHtml(e && e.message ? e.message : e));
  }
}

// 启动自动检查更新弹窗（稍后 / 现在更新）
async function autoUpdatePrompt() {
  try {
    if (sessionStorage.getItem("dsh_updatePromptShown")) return;
    sessionStorage.setItem("dsh_updatePromptShown", "1");
    const resp = await fetch("/api/update/check");
    const d = await resp.json();
    if (!d || !d.hasUpdate || !d.configured) return;
    const m = createUpdateModal("发现新版本 " + escapeHtml(d.latest), "🚀");
    const v = document.createElement("div"); v.innerHTML = updVersionsHtml(d.current, d.latest);
    m.body.appendChild(v);
    if (d.notes) { const n = document.createElement("div"); n.className = "upd-notes"; n.textContent = d.notes; m.body.appendChild(n); }
    const row = document.createElement("div"); row.className = "upd-actions";
    const later = document.createElement("button"); later.className = "btn"; later.type = "button"; later.textContent = "稍后"; later.onclick = () => m.close();
    const now = document.createElement("button"); now.className = "btn primary"; now.type = "button"; now.textContent = "现在更新"; now.onclick = () => beginUpdateDownload(d, m.close);
    row.appendChild(later); row.appendChild(now);
    m.body.appendChild(row);
  } catch {}
}

// 统一的确认弹窗（替代原生 confirm）：返回 Promise<boolean>
function showConfirmModal(message, opts = {}) {
  return new Promise((resolve) => {
    const overlay = document.createElement("div");
    overlay.className = "upd-overlay";
    const box = document.createElement("div");
    box.className = "req-modal";
    if (opts.title) {
      const h = document.createElement("h3");
      h.textContent = opts.title;
      box.appendChild(h);
    }
    const p = document.createElement("p");
    p.textContent = message;
    p.style = "line-height:1.7; white-space:pre-wrap; color:var(--text-2); font-size:13px;";
    box.appendChild(p);
    const row = document.createElement("div");
    row.style = "display:flex; gap:10px; justify-content:flex-end; margin-top:4px;";
    const btnCancel = document.createElement("button");
    btnCancel.className = "btn"; btnCancel.type = "button"; btnCancel.textContent = opts.cancelText || "取消";
    btnCancel.onclick = () => { overlay.remove(); resolve(false); };
    const btnOk = document.createElement("button");
    btnOk.className = opts.danger ? "btn danger" : "btn primary"; btnOk.type = "button"; btnOk.textContent = opts.confirmText || "确定";
    btnOk.onclick = () => { overlay.remove(); resolve(true); };
    row.appendChild(btnCancel); row.appendChild(btnOk);
    box.appendChild(row);
    overlay.appendChild(box);
    overlay.addEventListener("click", (e) => { if (e.target === overlay) { overlay.remove(); resolve(false); } });
    document.body.appendChild(overlay);
  });
}

// ==================== 本地生成历史 ====================
function timeText(iso) {
  try {
    const d = new Date(iso);
    const p = (n) => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
  } catch { return iso || ""; }
}
async function fetchGenHistory() {
  const resp = await fetch("/api/gen-history");
  const data = await resp.json();
  return (data && data.entries) || [];
}
function renderHistoryModal() {
  const overlay = document.createElement("div");
  overlay.className = "req-modal-overlay";
  const box = document.createElement("div");
  box.className = "req-modal";
  box.style = "width:min(860px,92vw); max-height:86vh; overflow:auto;";
  const h = document.createElement("h3");
  h.textContent = "本地生成历史（本机data\\gen_history）";
  const tip = document.createElement("p");
  tip.style = "color:var(--text-3); font-size:12px; line-height:1.6;";
  tip.textContent = "每次“开始解析生成”自动保存上传文件与生成结果，可按时间倒序打开重看/重新导出。此记录在软件目录下，与“版本记录”分开，本地生成不写入版本历史。";
  const listWrap = document.createElement("div");
  listWrap.style = "margin-top:12px;";
  const btns = document.createElement("div");
  btns.style = "margin-top:16px;";
  const closeBtn = document.createElement("button");
  closeBtn.className = "btn primary"; closeBtn.type = "button"; closeBtn.textContent = "关闭";
  closeBtn.addEventListener("click", () => overlay.remove());
  btns.appendChild(closeBtn);
  box.appendChild(h); box.appendChild(tip); box.appendChild(listWrap); box.appendChild(btns);
  overlay.appendChild(box);
  overlay.addEventListener("click", (e) => { if (e.target === overlay) overlay.remove(); });
  document.body.appendChild(overlay);
  listWrap.innerHTML = '<div style="color:var(--text-3);padding:8px;">加载中…</div>';
  fetchGenHistory().then((entries) => {
    if (!entries.length) {
      listWrap.innerHTML = '<div style="color:var(--text-3);padding:8px;">暂无生成历史。</div>';
      return;
    }
    const rows = entries.map((e) => {
      const fileText = Object.entries(e.fileNames || {}).map(([k, n]) => `${k}:${n}`).join("；");
      return `<div style="border-bottom:1px solid var(--border); padding:9px 4px; display:flex; align-items:flex-start; gap:10px;">
        <div style="flex:1; min-width:0;">
          <div><strong>${escapeHtml(timeText(e.time))}</strong> ｜ 部位：${escapeHtml(e.regions || "未选择")} ｜ 模式：${escapeHtml(e.modeLabel || "")}</div>
          <div style="font-size:12px; color:var(--text-2); margin-top:2px;">岗位 ${e.stations || 0} ｜ 导线 ${e.wires || 0} ｜ 工作包 ${e.packages || 0} ｜ 待确认/冲突 ${e.issues || 0} ｜ TT ${e.tt || "—"}${e.missingTimeSource ? " ｜【未找到工时源】" : ""}</div>
          ${fileText ? `<div style="font-size:12px; color:var(--text-3); margin-top:2px;">文件：${escapeHtml(fileText)}</div>` : ""}
        </div>
        <button class="btn" type="button" data-open="${e.id}">打开</button>
        <button class="btn" type="button" data-del="${e.id}">删除</button>
      </div>`;
    }).join("");
    listWrap.innerHTML = rows;
    listWrap.querySelectorAll("[data-open]").forEach((btn) => btn.addEventListener("click", () => { openGenHistory(btn.dataset.open); overlay.remove(); }));
    listWrap.querySelectorAll("[data-del]").forEach((btn) => btn.addEventListener("click", async () => {
      const ok = await showConfirmModal("确定删除该条生成历史吗？（本机原始项目文件不受影响）", { title: "删除生成历史", confirmText: "删除", danger: true });
      if (!ok) return;
      try { await fetch("/api/gen-history/" + btn.dataset.del, { method: "DELETE" }); renderHistoryModal(); } catch {}
    }));
  }).catch((err) => { listWrap.innerHTML = '<div class="validation-error">读取失败：' + escapeHtml(err.message || err) + '</div>'; });
}
async function openGenHistory(id) {
  try {
    const resp = await fetch("/api/gen-history/" + id);
    if (!resp.ok) throw new Error("打开失败");
    const data = await resp.json();
    result = data.result;
    activeTab = "";
    renderResult();
    window.scrollTo({ top: document.body.scrollHeight, behavior: "smooth" });
    toast("已从生成历史打开该次生成结果，可查看或导出。");
  } catch (err) {
    toast(err.message, "error");
  }
}

function bindFileCards() {
  document.querySelectorAll(".file-drop").forEach((card) => {
    const field = card.dataset.field;
    const input = card.querySelector("input");

    card.addEventListener("click", () => input.click());

    input.addEventListener("change", () => {
      const file = input.files && input.files[0];
      if (file) {
        files[field] = file;
        stashFile(field, file);
        card.classList.remove("required-flash");
        setFileFilled(field, file.name);
        resetAnalysisState();
      }
    });

    card.addEventListener("dragover", (e) => {
      e.preventDefault();
      card.classList.add("dragover");
    });
    card.addEventListener("dragleave", () => card.classList.remove("dragover"));
    card.addEventListener("drop", (e) => {
      e.preventDefault();
      card.classList.remove("dragover");
      const file = e.dataTransfer.files && e.dataTransfer.files[0];
      if (file) {
        files[field] = file;
        input.files = e.dataTransfer.files;
        stashFile(field, file);
        card.classList.remove("required-flash");
        setFileFilled(field, file.name);
        resetAnalysisState();
      }
    });
  });
}

// 上传文件时立即暂存到服务端，页面跳转/刷新后可通过 loadStashedFiles 恢复
async function stashFile(key, file) {
  try {
    const fd = new FormData();
    fd.append(key, file);
    await fetch("/api/stash", { method: "POST", body: fd });
  } catch {}
}

// 文件卡片状态：已填（显示文件名+删除按钮） / 空白（未选择，隐藏删除按钮）
function fileCard(field) {
  return document.querySelector(`.file-drop[data-field="${field}"]`);
}
function setFileFilled(field, text) {
  const card = fileCard(field);
  if (!card) return;
  const name = card.querySelector(".file-name");
  name.textContent = text;
  name.classList.add("ready");
  const btn = card.querySelector(".file-clear");
  if (btn) btn.hidden = false;
}
function setFileBlank(field) {
  const card = fileCard(field);
  if (!card) return;
  const name = card.querySelector(".file-name");
  name.textContent = "未选择";
  name.classList.remove("ready", "required-flash");
  const btn = card.querySelector(".file-clear");
  if (btn) btn.hidden = true;
}
// 删除单个已上传文件（清浏览器内存 + 服务端暂存）
async function deleteUploadedFile(field) {
  const card = fileCard(field);
  const label = card ? (card.dataset.label || field) : field;
  if (!files[field]) { toast(`「${label}」尚未上传，无需删除`, "error"); return; }
  const ok = await showConfirmModal(`确定删除已上传的「${label}」吗？（不影响你本机原始文件，删除后需重新上传）`, { title: "删除文件", confirmText: "删除", danger: true });
  if (!ok) return;
  delete files[field];
  setFileBlank(field);
  try { await fetch("/api/stash/" + field, { method: "DELETE" }); } catch {}
  resetAnalysisState();
  toast(`已删除「${label}」`);
}
// 清空全部已上传文件
async function clearAllUploadedFiles() {
  const fields = ["standard", "ebom", "mbom", "pdf"];
  const hasAny = fields.some((k) => !!files[k]);
  if (!hasAny) { toast("当前没有已上传的文件", "error"); return; }
  const ok = await showConfirmModal("确定清空全部已上传的文件吗？（不影响你本机原始文件，清空后需重新上传）", { title: "清空已上传文件", confirmText: "清空", danger: true });
  if (!ok) return;
  for (const k of fields) { delete files[k]; setFileBlank(k); }
  try { await fetch("/api/stash", { method: "DELETE" }); } catch {}
  resetAnalysisState();
  toast("已清空全部上传文件");
}

// 页面加载时从服务端恢复上次暂存的项目资料
async function loadStashedFiles() {
  try {
    const resp = await fetch("/api/stash");
    const data = await resp.json();
    let restored = 0;
    for (const [key, meta] of Object.entries(data.files || {})) {
      if (files[key]) continue; // 用户本次会话已上传，不覆盖
      const r = await fetch("/api/stash/" + key);
      if (!r.ok) continue;
      const buf = await r.arrayBuffer();
      files[key] = { buffer: buf, originalname: meta.name };
      setFileFilled(key, meta.name + "（已恢复）");
      restored += 1;
    }
    if (restored) {
      toast(`已恢复上次暂存的${restored}个文件，请重新点击“分析文件”。`);
    }
  } catch {}
}

function appendFileToFd(fd, key, f) {
  if (!f) return;
  if (f.buffer) fd.append(key, new Blob([f.buffer]), f.originalname);
  else fd.append(key, f);
}

function buildAnalysisFd() {
  const fd = new FormData();
  appendFileToFd(fd, "standard", files.standard);
  appendFileToFd(fd, "ebom", files.ebom);
  appendFileToFd(fd, "mbom", files.mbom);
  appendFileToFd(fd, "pdf", files.pdf);
  return fd;
}

// 统一的工艺策划参数表单字段（供“分析文件”与“开始解析生成”共用，保证预览与正式生成口径一致）
function appendAnalysisOptions(fd) {
  fd.append("tt", $("#tt").value || "");
  fd.append("preassemblyMode", $("#preassemblyMode").value || "");
  fd.append("overTtPolicy", (document.querySelector('input[name="overTtPolicy"]:checked') || {}).value || "forbid");
  fd.append("loopSingleKit", $("#loopSingleKit").value || "");
  fd.append("loopKitTransferMiddle", $("#loopKitTransferMiddle").value || "");
  fd.append("loopKitTransferLast", $("#loopKitTransferLast").value || "");
  fd.append("loopSubLast", $("#loopSubLast").value || "");
  fd.append("standardIncludesLoop", ($("#loopIncludeCheck") && $("#loopIncludeCheck").checked) ? "true" : "false");
}

async function analyzeFiles() {
  if (!files.mbom) {
    toast("请先上传MBOM/Cutting文件。", "error");
    return;
  }
  const status = $("#status");
  status.hidden = false;
  status.className = "status";
  status.textContent = "正在分析MBOM，识别护套、SP/SC压接点和配置……";
  $("#analyzeFilesBtn").disabled = true;
  try {
    const resp = await fetch("/api/analyze-files", { method: "POST", body: (() => {
      const fd = buildAnalysisFd();
      appendAnalysisOptions(fd);
      return fd;
    })() });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.error || "文件分析失败");
    fileAnalysis = data;
    renderFileAnalysis();
    const limitWarn = validateLimitInputs();
    $("#analyzeBtn").disabled = !data.canGenerate;
    if (data.canGenerate) {
      status.textContent = `分析完成：护套 ${data.housingCount} 个，SP/SC压接点 ${data.spliceCount} 个，模板检查通过，请填写限制后点击“开始解析生成”。`;
      if (limitWarn) status.textContent += " 注意：" + limitWarn;
      status.classList.remove("error");
      toast("文件分析完成，模板检查通过。");
    } else {
      status.textContent = "模板存在问题，请根据下方检查结果修正文件后重新点击“分析文件”。";
      status.classList.add("error");
      toast("模板存在问题，请修正后重新分析。", "error");
    }
  } catch (err) {
    status.textContent = "分析失败：" + err.message;
    status.classList.add("error");
    toast(err.message, "error");
  } finally {
    $("#analyzeFilesBtn").disabled = false;
  }
}

function splitInput(value) {
  return String(value || "").split(/[/\n,，;；]+/).map((s) => s.trim()).filter(Boolean);
}

function validateLimitInputs() {
  if (!fileAnalysis) return "";
  const a = fileAnalysis;
  const knownSplices = new Set((a.spliceList || []).map((s) => s.code));
  const knownHousings = new Set((a.housingList || []).map((h) => h.code));
  const unknownSplices = splitInput($("#noOnlineUltrasonicSplices").value).filter((code) => !knownSplices.has(code));
  const unknownHousings = splitInput($("#forcedOfflineHousings").value).filter((code) => !knownHousings.has(code));
  const unknownGroupHousings = collectSameStationGroups()
    .flatMap((g) => splitInput(g.housings))
    .filter((code) => !knownHousings.has(code));
  const msgs = [];
  if (unknownSplices.length) msgs.push(`以下SP/SC点不在分析清单中：${unknownSplices.join("、")}`);
  if (unknownHousings.length) msgs.push(`以下护套不在分析清单中：${unknownHousings.join("、")}`);
  if (unknownGroupHousings.length) msgs.push(`以下同岗位分组护套不在分析清单中：${unknownGroupHousings.join("、")}`);
  return msgs.join("；");
}

function renderFileAnalysis() {
  const panel = $("#analysisPanel");
  panel.hidden = false;
  const a = fileAnalysis;
  const groupText = Object.entries(a.onlineUltrasonicGroupsPerConfig || {})
    .map(([cfg, n]) => `${cfg}: ${n}个`)
    .join("<br>") || "无";
  // 按最高配置工时 ÷ TT 测算预计岗位数，回填到“最多预装岗位数”占位符
  if (a.estimatedStations) {
    $("#maxStations").placeholder = `按节拍测算约需 ${a.estimatedStations} 个岗位（可留空自动 / 可作上限修改）`;
  }
  $("#analysisSummary").innerHTML = [
    ["MBOM导线", a.totalWires + " 行"],
    ["唯一W1", a.uniqueW1 + " 个"],
    ["护套数量", a.housingCount + " 个"],
    ["SP/SC压接点", a.spliceCount + " 个"],
    ["配置数量", (a.configs || []).length + " 个"],
    ["预计岗位数(按最高配置工时÷每岗承载容量)", a.estimatedStations ? `${a.estimatedStations} 个` : ((a.tt ? "（TT未算）" : "需填写TT") + (a.maxConfigSeconds ? ` / 最高配置${a.maxConfigSeconds}s` : ""))],
    ["单配置压接点数量", groupText]
  ].map(([k, v]) => `<div class="analysis-card"><div class="label">${k}</div><div class="value">${v}</div></div>`).join("");

  const validation = a.validation || {};
  const fileLabels = { standard: "标准工时文件", ebom: "EBOM文件", mbom: "MBOM/Cutting文件" };
  const validationHtml = Object.entries(fileLabels).map(([key, label]) => {
    const v = validation[key];
    if (!v) return "";
    const icon = v.ok ? "✅" : "❌";
    const issueHtml = (v.issues || []).map((it) => {
      const cls = it.level === "error" ? "validation-error" : "validation-warning";
      return `<div class="${cls}">${it.level === "error" ? "【错误】" : "【警告】"}${it.message}</div>`;
    }).join("");
    return `<div class="validation-file">
      <div class="validation-title">${icon} ${label}：${v.ok ? "模板检查通过" : "模板存在问题"}（${v.rowCount || 0}行）</div>
      ${issueHtml || ""}
    </div>`;
  }).join("");
  $("#analysisValidation").innerHTML = validationHtml || "";

  const housingHtml = (a.housingList || []).slice(0, 60).map((h) => `<code title="端数：${h.endCount}">${h.code}</code>`).join("");
  const spliceHtml = (a.spliceList || []).map((s) => `<code title="端数：${s.endCount}；配置：${s.configs}">${s.code}</code>`).join("");
  $("#analysisLists").innerHTML = `
    <div class="analysis-list">
      <h4>MBOM识别到的SP/SC压接点（${a.spliceCount}）</h4>
      ${spliceHtml || "无"}
    </div>
    <div class="analysis-list">
      <h4>MBOM识别到的护套（${a.housingCount}）</h4>
      ${housingHtml || "无"}
    </div>
  `;
}

async function analyze() {
  if (!files.standard || !files.ebom || !files.mbom) {
    toast("请先选择标准工时、EBOM、MBOM三个必填文件。", "error");
    return;
  }
  if (!fileAnalysis) {
    toast("请先点击“分析文件”，完成MBOM分析后再生成。", "error");
    return;
  }

  const status = $("#status");
  status.hidden = false;
  status.className = "status";
  status.textContent = "正在解析文件并生成结构化数据包……";
  $("#analyzeBtn").disabled = true;

  const maxStationsVal = $("#maxStations").value.trim();
  let autoStations = false;
  if (!maxStationsVal) {
    // 最多预装岗位数可不填：有TT时直接按节拍自动计算，不再二次确认
    const ttVal = $("#tt").value.trim();
    if (!ttVal) {
      status.textContent = "未填写最多岗位数，也未填写TT，无法自动计算岗位数。请填写TT（或填写最多岗位数）。";
      status.classList.add("error");
      $("#analyzeBtn").disabled = false;
      return;
    }
    autoStations = true;
  }

  const fd = buildAnalysisFd();
  appendAnalysisOptions(fd);
  const regionChecked = document.querySelector('input[name="region"]:checked');
  if (regionChecked) fd.append("regions", regionChecked.value);
  fd.append("maxStations", maxStationsVal);
  fd.append("autoStations", autoStations ? "true" : "false");
  fd.append("onlineUltrasonic", $("#onlineUltrasonic").value || "no");
  fd.append("onlineUltrasonicMaxGroupsPerConfig", $("#onlineUltrasonicMaxGroupsPerConfig").value || "");
  fd.append("onlineUltrasonicMaxTotalGroups", $("#onlineUltrasonicMaxTotalGroups").value || "");
  fd.append("noOnlineUltrasonicSplices", $("#noOnlineUltrasonicSplices").value || "");
  fd.append("forcedOfflineHousings", $("#forcedOfflineHousings").value || "");
  fd.append("sameStationGroups", JSON.stringify(collectSameStationGroups()));
  fd.append("sameStationOverTtMode", (document.querySelector('input[name="sameStationOverTtMode"]:checked') || {}).value || "force-same");
  fd.append("maxSubFrames", $("#maxSubFrames").value || "");
  fd.append("grommetStations", JSON.stringify(collectGrommetRows()));

  const limitWarn = validateLimitInputs();
  if (limitWarn) toast(limitWarn, "error");

  try {
    const resp = await fetch("/api/analyze", { method: "POST", body: fd });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.error || "解析失败");
    result = data;
    activeTab = "";
    renderResult();
    status.textContent = (limitWarn ? "注意：" + limitWarn + "；" : "") + `生成完成：${result.summary.wires}根导线，${result.summary.packages}个候选工作包，${result.summary.stations || 0}个候选岗位，${result.summary.issues}项待确认/冲突。`;
    status.classList.remove("error");
    toast("解析生成完成，可查看下方结果并导出Excel。" + (limitWarn ? "（有未识别限制输入，请查看“冲突与待确认”）" : ""));
  } catch (err) {
    status.textContent = "解析失败：" + err.message;
    status.classList.add("error");
    toast(err.message, "error");
  } finally {
    $("#analyzeBtn").disabled = !fileAnalysis;
  }
}

function renderSummary() {
  const s = result.summary;
  const plan = result.plan || {};
  const cards = [
    ["制作部位", (plan.regions || []).join("、") || "未选择"],
    ["预装模式", plan.modeLabel || "未选择"],
    ["最大岗位数", stationCountText(plan)],
    ["在线超声波", plan.onlineUltrasonic ? "是" : "否"],
    ["单配置超声限组", plan.onlineUltrasonicMaxGroupsPerConfig || "未设置"],
    ["超声总组数限制", plan.onlineUltrasonicMaxTotalGroups || "未设置"],
    ["MBOM导线", s.wires + " 行 / " + (s.uniqueW1 || s.wires) + " 唯一W1"],
    ["EBOM物料", s.materials + " 条"],
    ["标准工时", s.standardEntries + " 条"],
    ["候选工作包", s.packages + " 个"],
    ["候选岗位", s.stations + " 个"],
    ["护套关联对", s.housingPairs + " 对"],
    ["配置", (result.configs || []).join("、") || "无"],
    ["待确认/冲突", s.issues + " 项"],
    ["同色编码行", s.dotRows + " 条"]
  ];
  $("#summaryCards").innerHTML = cards.map(([k, v]) => `
    <div class="summary-card">
      <div class="label">${k}</div>
      <div class="value">${v}</div>
    </div>
  `).join("");
}

function defineTabs() {
  const plan = result.plan || {};
  return [
    { id: "plan", label: "工艺策划参数", headers: ["参数", "值"], rows: [
      ["制作部位", (plan.regions || []).join("、") || "未选择"],
      ["最多预装岗位数", stationCountText(plan)],
      ["预装模式代码", plan.preassemblyMode || ""],
      ["预装模式说明", plan.modeLabel || "未选择"],
      ["是否有在线超声波", plan.onlineUltrasonic ? "是" : "否"],
      ["单配置在线超声波最高组数", plan.onlineUltrasonicMaxGroupsPerConfig || "未设置"],
      ["在线超声波最高总组数限制", plan.onlineUltrasonicMaxTotalGroups || "未设置"],
      ["不能在线SP/SC压接点", (plan.noOnlineUltrasonicSplices || []).join("、") || "无"],
      ["强制线下插接护套", (plan.forcedOfflineHousings || []).join("、") || "无"],
      ["组超节拍处理方式", plan.sameStationOverTtMode === "best-rate" ? "按最佳插接率拆分" : "强制同岗（默认）"],
      ["超节拍处理", plan.overTtPolicy === "allow" ? "可超节拍（多人合干，自动建议N人）" : "禁止超节拍（默认，单岗单人）"],
      ["打圈工时(秒/岗位)", `单KIT:${(plan.loopTimes && plan.loopTimes.singleKit) || 0}；KIT传递中间:${(plan.loopTimes && plan.loopTimes.kitTransferMiddle) || 0}；KIT传递末:${(plan.loopTimes && plan.loopTimes.kitTransferLast) || 0}；SUB末:${(plan.loopTimes && plan.loopTimes.subLast) || 0}`],
      ["标准工时已含打圈", plan.standardIncludesLoop ? "是（不重复计50秒）" : "否（按固定值计入）"],
      ["工时源完整", plan.missingTimeSource ? "否（存在【未找到工时源】动作，TT/岗位数为候选估算）" : "是"],
      ["同岗位护套分组", (plan.sameStationGroups || []).map(g => {
        const ml = ({ "pure-kit": "纯KIT岗位", "kit-transfer": "KIT传递岗", "sub": "SUB岗位" })[g.mode] || g.mode || "纯KIT岗位";
        return `${g.name || "同岗组"}[${ml}](${g.housings})`;
      }).join("；") || "无"],
      ["TT(秒)", result.options && result.options.tt != null ? result.options.tt : "未设置"],
      ["说明", plan.note || ""]
    ]},
    ...((plan.groupTtRows && plan.groupTtRows.length) ? [{
      id: "groupTt",
      label: "同岗位分组节拍处理",
      headers: ["序号", "组名", "护套", "岗位类型", "合并工时(秒)", "TT(秒)", "负荷率%", "处理方式", "结果", "拆分情况/说明", "状态"],
      rows: plan.groupTtRows.map(r => [r.idx, r.groupName, r.housings, r.modeLabel, r.mergedSeconds, r.tt, r.loadPercent, r.handleMode, r.outcome, r.splitDetail, r.status])
    }] : []),
    { id: "stations", label: "候选岗位分配", headers: ["岗位号", "预装模式", "包含工作包", "工作包数", "导线数", "估算工时(秒)", "各配置工时(秒)", "TT(秒)", "负荷率%", "配置", "状态"], rows: (plan.stationAllocation || []).map(r => [r.stationNo, r.modeLabel, r.packageIds, r.packageCount, r.wireCount, r.totalSeconds, Object.entries(r.configSeconds || {}).map(([c, s]) => `${c}:${s}s`).join("；"), r.tt, r.loadPercent, r.configs, r.status]) },
    { id: "stationDetails", label: "岗位明细", headers: ["岗位号", "岗位名称", "制作部位", "预装模式", "导线数", "估算工时(秒)", "负荷率%", "打圈(秒)", "建议人数", "各配置工时(秒)", "胶带包胶备注", "状态"], rows: (plan.stationDetails || []).map(r => [r.stationNo, r.stationName, r.region, r.modeLabel, (r.wireRows || []).length, r.totalSeconds, r.loadPercent != null ? r.loadPercent : "", r.loopTimeSeconds || 0, r.workerCount || 1, Object.entries(r.configTime || {}).map(([c, s]) => `${c}:${s}s`).join("；"), r.tapeRemark, r.status]) },
    { id: "processFlow", label: "过程流程图", headers: (result.processFlow || [])[0] || [], rows: (result.processFlow || []).slice(1) },
    { id: "pfmea", label: "PFMEA", headers: (result.pfmeaRows || [])[0] || [], rows: (result.pfmeaRows || []).slice(1) },
    { id: "controlPlan", label: "控制计划", headers: (result.controlPlanRows || [])[0] || [], rows: (result.controlPlanRows || []).slice(1) },
    { id: "ledger", label: "输入台账", headers: ["文件类型", "文件名称", "大小KB", "读取状态", "简要信息"], rows: result.ledger.map(r => [r.fileType, r.fileName, r.sizeKB, r.status, r.info]) },
    { id: "issues", label: "冲突与待确认", headers: ["序号", "类别", "说明"], rows: result.issues.map((r, i) => [i + 1, r.category, r.detail]) },
    { id: "wires", label: "MBOM导线表", headers: ["序号", "W3看板号", "W2看板号", "W1看板号", "材料名称", "图纸号", "客户号", "规格", "颜色", "下料长度", "单位", "端子1", "雨塞1", "端子2", "雨塞2", "护套1", "孔位1", "护套2", "孔位2", "配置", "状态"], rows: result.wires.map(r => [r.idx, r.w3, r.w2, r.w1, r.material, r.drawingId, r.customerNo, r.spec, r.color, r.length, r.unit, r.terminal1, r.seal1, r.terminal2, r.seal2, r.housing1, r.position1, r.housing2, r.position2, r.configs, r.status]) },
    { id: "ebom", label: "EBOM物料表", headers: ["序号", "模块号", "模块名称", "材料名称", "Description", "图纸号", "捷翼号", "厂家号", "厂家", "图纸用量", "工艺余量", "总用量", "单位", "单价", "价格汇总", "理论铜重", "备注"], rows: result.ebomMaterials.map(r => [r.idx, r.moduleNo, r.moduleName, r.materialName, r.description, r.drawingId, r.jettyNo, r.spn, r.supplier, r.designQty, r.processAllowance, r.totalQty, r.unit, r.unitPrice, r.totalPrice, r.copperWeight, r.notes]) },
    { id: "standard", label: "标准工时表", headers: ["序号", "工序", "工作要素", "描述", "动作开始", "动作结束", "标准工时", "单位"], rows: result.standardHours.map(r => [r.idx, r.process, r.activity, r.comments, r.clockStart, r.clockStop, r.time, r.unit]) },
    { id: "path", label: "导线完整路径", headers: ["W3", "W2", "W1", "图纸号", "材料", "颜色", "长度", "护套1", "孔位1", "端子1", "雨塞1", "护套2", "孔位2", "端子2", "雨塞2", "焊点关系", "主干", "分支", "分支点", "滑板/工装板槽位", "保护区域", "附件区域", "配置", "候选包", "状态"], rows: result.pathRows.map(r => [r.w3, r.w2, r.w1, r.drawingId, r.material, r.color, r.length, r.housing1, r.position1, r.terminal1, r.seal1, r.housing2, r.position2, r.terminal2, r.seal2, r.spliceRelation, r.trunk, r.branch, r.branchPoint, r.boardSlot, r.protectArea, r.accessoryArea, r.configs, r.pkgId, r.status]) },
    { id: "matrix", label: "护套关联矩阵", headers: ["起始护套", "目标护套", "关联导线数(并集)", "各配置关联数(Nij)", "最强/次强关联", "导线编号", "候选工作包", "配置", "状态"], rows: result.housingMatrix.map(r => [r.housingA, r.housingB, r.count, r.perConfig, r.strength, r.wires, r.pkgIds, r.configs, r.status]) },
    { id: "packages", label: "候选预装工作包", headers: ["工作包编号", "层级", "名称/看板号", "导线数", "总长度", "最长导线", "包含护套", "锚点", "配置", "路线候选", "估算工时(秒)", "状态"], rows: result.packages.map(r => [r.id, r.kind, r.name, r.wireCount, r.totalLength, r.maxLength, r.housings, r.anchor, r.configs, r.routeType, r.estimatedSeconds, r.status]) },
    { id: "positions", label: "孔位责任矩阵", headers: ["工作包", "护套", "孔位", "导线", "端子", "雨塞", "配置", "状态"], rows: result.positionRows.map(r => [r.pkgId, r.housing, r.position, r.wire, r.terminal, r.seal, r.configs, r.status]) },
    { id: "dots", label: "同色线编码", headers: ["工作包", "导线颜色", "导线", "目标护套/孔位", "配置", "标准化编码", "查重", "状态"], rows: result.dotRows.map(r => [r.pkgId, r.wireColor, r.wireId, r.target, r.config, r.standardCode, r.checkResult, r.status]) },
    { id: "times", label: "岗位×配置工时", headers: ["工作包", "名称", "配置", "估算工时(秒)", "TT(秒)", "负荷率%", "建议人数", "备注"], rows: result.timeRows.map(r => [r.pkgId, r.pkgName, r.config, r.estimatedSeconds, r.tt, r.loadPercent, r.workerSuggestion, r.note]) },
    { id: "pdf", label: "PDF图纸关键词", headers: ["关键词", "出现次数"], rows: result.pdfKeywords.map(r => [r.keyword, r.count]) }
  ];
}

function escapeHtml(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function renderTable(tab, headers, allRows) {
  const total = allRows.length;
  const limit = renderLimits.get(tab.id) || 100;
  const panel = $("#panel");

  let html = `<h3>${tab.label}</h3>`;
  html += `<div class="search-bar"><input type="search" placeholder="搜索当前表" data-table-search="${tab.id}"><span>共 ${humanNumber(total)} 行</span></div>`;
  if (total > limit) {
    html += `<button class="btn" type="button" data-more="${tab.id}" style="margin-bottom:10px">显示更多（每次+200）</button>`;
  }
  html += `<div class="table-wrap"><table><thead><tr>${headers.map(h => `<th>${escapeHtml(h)}</th>`).join("")}</tr></thead><tbody>`;

  const sliceRows = allRows.slice(0, limit);
  html += sliceRows.map(row => `<tr>${row.map(cell => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`).join("");
  if (!sliceRows.length) html += `<tr><td colspan="${headers.length}">无数据</td></tr>`;
  html += `</tbody></table></div>`;
  if (total > limit) {
    html += `<p class="muted">当前显示 ${Math.min(limit, total)} / ${humanNumber(total)} 行</p>`;
  }
  panel.innerHTML = html;

  const input = panel.querySelector(`[data-table-search="${tab.id}"]`);
  if (input) {
    input.addEventListener("input", () => {
      const q = input.value.trim().toLowerCase();
      const filtered = q
        ? allRows.filter(row => row.some(cell => String(cell == null ? "" : cell).toLowerCase().includes(q)))
        : allRows;
      renderTableBody(tab, headers, filtered, q ? filtered.length : total);
    });
  }

  const moreBtn = panel.querySelector(`[data-more="${tab.id}"]`);
  if (moreBtn) {
    moreBtn.addEventListener("click", () => {
      const q = String((panel.querySelector(`[data-table-search="${tab.id}"]`) || {}).value || "").trim().toLowerCase();
      renderLimits.set(tab.id, (renderLimits.get(tab.id) || 100) + 200);
      if (q) {
        const filtered = allRows.filter(row => row.some(cell => String(cell == null ? "" : cell).toLowerCase().includes(q)));
        renderTableBody(tab, headers, filtered, filtered.length);
      } else {
        renderTable(tab, headers, allRows);
      }
    });
  }
}

function renderTableBody(tab, headers, rows, displayTotal) {
  const panel = $("#panel");
  const tbody = panel.querySelector("table tbody");
  const limit = renderLimits.get(tab.id) || 100;
  const sliceRows = rows.slice(0, limit);
  tbody.innerHTML = sliceRows.map(row => `<tr>${row.map(cell => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`).join("");
  if (!sliceRows.length) tbody.innerHTML = `<tr><td colspan="${headers.length}">无匹配数据</td></tr>`;
  const info = panel.querySelector(".muted");
  if (info) info.textContent = `当前显示 ${Math.min(limit, rows.length)} / ${humanNumber(displayTotal)} 行`;
}

function renderTabs(tabs) {
  const wrap = $("#tabs");
  wrap.innerHTML = tabs.map(t => `<button class="tab-btn" data-tab="${t.id}">${t.label}</button>`).join("");
  wrap.querySelectorAll(".tab-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      activeTab = btn.dataset.tab;
      wrap.querySelectorAll(".tab-btn").forEach(b => b.classList.toggle("active", b.dataset.tab === activeTab));
      const tab = tabs.find(t => t.id === activeTab);
      renderTab(tab);
    });
  });
}

function renderTab(tab) {
  renderTable(tab, tab.headers, tab.rows);
}

function renderResult() {
  $("#result").hidden = false;
  renderSummary();
  const tabs = defineTabs();
  renderTabs(tabs);
  activeTab = tabs[0].id;
  document.querySelectorAll(".tab-btn").forEach(b => b.classList.toggle("active", b.dataset.tab === activeTab));
  renderTab(tabs[0]);
  $("#exportBtn").disabled = false;
  $("#exportProcessBtn").disabled = false;
  $("#exportReviewBtn").disabled = false;
}

async function exportWorkbook(endpoint, filename, message) {
  if (!result) return;
  try {
    const resp = await fetch(endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(result)
    });
    if (!resp.ok) {
      const data = await resp.json().catch(() => ({}));
      throw new Error(data.error || "导出失败");
    }
    const blob = await resp.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    toast(message || "文件已生成。");
  } catch (err) {
    toast(err.message, "error");
  }
}

function exportExcel() {
  const stamp = new Date().toISOString().slice(0, 10);
  return exportWorkbook("/api/export", `预装工艺结构化数据包_${stamp}.xlsx`, "结构化数据包已生成。");
}

function exportProcess() {
  const stamp = new Date().toISOString().slice(0, 10);
  return exportWorkbook("/api/export/process", `预装工艺岗位表_${stamp}.xlsx`, "预装工艺岗位表已生成。");
}

function exportReview() {
  const stamp = new Date().toISOString().slice(0, 10);
  return exportWorkbook("/api/export/review", `预装工艺人工核查表_${stamp}.xlsx`, "人工核查表已生成。");
}

async function downloadTemplate(type) {
  const names = {
    standard: "标准工时文件模板.xlsx",
    ebom: "EBOM文件模板.xlsx",
    mbom: "MBOM-Cutting模板.xlsx"
  };
  try {
    const resp = await fetch("/api/templates/" + type);
    if (!resp.ok) throw new Error("模板下载失败");
    const blob = await resp.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = names[type] || ("template_" + type + ".xlsx");
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    toast("模板已下载：" + names[type]);
  } catch (err) {
    toast(err.message, "error");
  }
}

function closeAllCustomSelects() {
  document.querySelectorAll(".custom-select-wrap.open").forEach((wrap) => {
    wrap.classList.remove("open");
    const menu = wrap.querySelector(".custom-select-menu");
    if (menu) menu.hidden = true;
  });
}

function updateUltrasonicInputState() {
  const enabled = $("#onlineUltrasonic").value === "yes";
  $("#fieldOnlineUltrasonicMaxGroupsPerConfig").hidden = !enabled;
  $("#fieldOnlineUltrasonicMaxTotalGroups").hidden = !enabled;
  $("#onlineUltrasonicMaxGroupsPerConfig").disabled = !enabled;
  $("#onlineUltrasonicMaxTotalGroups").disabled = !enabled;
  if (!enabled) {
    $("#onlineUltrasonicMaxGroupsPerConfig").value = "";
    $("#onlineUltrasonicMaxTotalGroups").value = "";
  }
}

function updateModeFields() {
  const mode = $("#preassemblyMode").value;
  const showSub = ["pure-sub", "sub-kit", "sub-kit-transfer"].includes(mode);
  const showKitTransfer = ["kit-transfer-kit", "sub-kit-transfer"].includes(mode);
  const showSingleKit = mode !== "" && mode !== "pure-sub";
  // 隐藏时暂存值并清空、恢复时回填：避免“残留值”被误读为当前模式生效的打圈/SUB参数
  manageModeField("fieldMaxSubFrames", "maxSubFrames", showSub);
  manageModeField("fieldLoopSubLast", "loopSubLast", showSub);
  manageModeField("fieldLoopKitTransferMiddle", "loopKitTransferMiddle", showKitTransfer);
  manageModeField("fieldLoopKitTransferLast", "loopKitTransferLast", showKitTransfer);
  manageModeField("fieldLoopSingleKit", "loopSingleKit", showSingleKit);
}

// 模式相关字段的显隐 + 值暂存/回填 + disabled 同步
function manageModeField(labelId, inputId, shown) {
  const lbl = document.getElementById(labelId);
  const inp = document.getElementById(inputId);
  if (!lbl || !inp) return;
  if (shown) {
    if (inp.dataset.saved !== undefined) {
      inp.value = inp.dataset.saved;
      delete inp.dataset.saved;
    }
    lbl.hidden = false;
    inp.disabled = false;
  } else {
    if (!inp.disabled && inp.dataset.saved === undefined) {
      inp.dataset.saved = inp.value;
      inp.value = "";
    }
    lbl.hidden = true;
    inp.disabled = true;
  }
}

function updateConditionalFields() {
  updateUltrasonicInputState();
  updateModeFields();
}

function initCustomSelects() {
  document.querySelectorAll("select.custom-select").forEach((select) => {
    if (select.dataset.customReady) return;
    select.dataset.customReady = "1";

    const wrap = document.createElement("div");
    wrap.className = "custom-select-wrap";

    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "custom-select-btn";
    const btnText = document.createElement("span");
    btnText.className = "custom-select-btn-text";
    btnText.textContent = select.options[select.selectedIndex]
      ? select.options[select.selectedIndex].textContent
      : "";
    btn.appendChild(btnText);

    const menu = document.createElement("div");
    menu.className = "custom-select-menu";
    menu.hidden = true;

    Array.from(select.options).forEach((opt) => {
      const item = document.createElement("button");
      item.type = "button";
      item.className = "custom-select-option" + (opt.selected ? " selected" : "");
      item.textContent = opt.textContent;
      item.dataset.value = opt.value;
      item.addEventListener("click", () => {
        select.value = opt.value;
        btnText.textContent = opt.textContent;
        menu.querySelectorAll(".custom-select-option").forEach((o) => {
          o.classList.toggle("selected", o === item);
        });
        menu.hidden = true;
        wrap.classList.remove("open");
        select.dispatchEvent(new Event("change", { bubbles: true }));
      });
      menu.appendChild(item);
    });

    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const isOpen = !menu.hidden;
      closeAllCustomSelects();
      menu.hidden = isOpen;
      wrap.classList.toggle("open", !isOpen);
    });

    select.addEventListener("change", () => {
      const selected = select.options[select.selectedIndex];
      btnText.textContent = selected ? selected.textContent : "";
      menu.querySelectorAll(".custom-select-option").forEach((o) => {
        o.classList.toggle("selected", o.dataset.value === select.value);
      });
    });

    wrap.appendChild(btn);
    wrap.appendChild(menu);
    select.style.display = "none";
    select.parentNode.insertBefore(wrap, select.nextSibling);
  });

  document.addEventListener("click", closeAllCustomSelects);
}

function init() {
  bindFileCards();
  initCustomSelects();
  loadStashedFiles();
  $("#addGrommetRow").addEventListener("click", () => addGrommetRow());
  $("#addSameStationGroup").addEventListener("click", () => addSameStationGroupRow());
  updateConditionalFields();
  $("#onlineUltrasonic").addEventListener("change", updateConditionalFields);
  $("#preassemblyMode").addEventListener("change", updateConditionalFields);
  $("#analyzeFilesBtn").addEventListener("click", guardRequired(analyzeFiles));
  $("#analyzeBtn").addEventListener("click", guardRequired(analyze));
  $("#exportBtn").addEventListener("click", guardRequired(exportExcel));
  $("#exportProcessBtn").addEventListener("click", guardRequired(exportProcess));
  $("#exportReviewBtn").addEventListener("click", guardRequired(exportReview));
  // 补填后自动清除红框
  $("#tt").addEventListener("input", () => { if ($("#tt").value.trim()) clearRequiredFlash("#tt"); });
  $("#preassemblyMode").addEventListener("change", () => { if ($("#preassemblyMode").value) clearRequiredFlash("#preassemblyMode"); });
  // 已上传文件：逐项删除 / 一键清空
  document.querySelectorAll(".file-clear").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      deleteUploadedFile(btn.dataset.clr);
    });
  });
  $("#clearAllFiles").addEventListener("click", clearAllUploadedFiles);
  $("#checkUpdateBtn").addEventListener("click", checkUpdate);
  $("#genHistoryBtn").addEventListener("click", renderHistoryModal);
  document.querySelectorAll("[data-template]").forEach(btn => {
    btn.addEventListener("click", () => downloadTemplate(btn.dataset.template));
  });
  // 启动时自动检查更新（有新版则弹“稍后/现在更新”）
  autoUpdatePrompt();
}

init();

