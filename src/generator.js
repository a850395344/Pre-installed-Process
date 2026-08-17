"use strict";

const XLSX = require("xlsx-js-style");
const {
  clean,
  number,
  parseMBOM,
  parseEBOM,
  parseStandardHours,
  parsePDF
} = require("./parse");
const { validateUploadedFiles } = require("./validation");

const APPROVED_DOT_COLORS = ["PK", "RD", "OG", "YE", "GN", "WH", "VT", "BK"];

const REGION_OPTIONS = [
  "车身线束",
  "仪表线束",
  "前舱线束",
  "发动机线束",
  "左前门线束",
  "右前门线束",
  "左后门线束",
  "右后门线束",
  "顶棚线束",
  "前保险杠线束",
  "后保险杠线束"
];

const REGION_RULES = {
  "车身线束": "车身线束为整体域；预装以护套—护套关联、主干/分支点、受控半成品和TT为主线，区域名称只作数据隔离。",
  "仪表线束": "锚点：中心高密度护套组、主干工装板基准、最强关联护套组、短分支首次出口、可视插接面。候选包：中心护套插接包→外围短分支插接/定向包→保护件装配包。",
  "前舱线束": "锚点：车身固定主接口、穿隔板件、主干分支点、长线工装槽位、保护件装入窗口。候选包：主接口/主干定位包→多向分支护套插接包→后护套/保护件/固定件包。",
  "发动机线束": "锚点：主接口、工装固定点、分支出口基准、自然汇合点、保护件装配窗口。候选包：主干骨架定位包→分支插接/定向包→保护件与最终重组包。",
  "左前门线束": "锚点：左件车侧护套—连续过渡护套—门侧主干、板上定位点、实际分支汇合点。候选包：车侧插接/过渡段成形包→门侧主干/分支插接包→规定保护件与输出包。连续过渡段不可任意截断。",
  "右前门线束": "锚点：右件车侧护套—连续过渡护套—门侧主干、右件板基准、实际分支汇合点。方法可参考左前门，但图号、孔位、物料、工时、治具和漆点池必须独立建模。",
  "左后门线束": "锚点：左后件车侧/门侧护套、连续过渡段、主干分支点、可视插接面。候选包：主接口/过渡段连续成形包→底层分支插接包→保护件与输出包。是否单KIT由实际关联和工时决定。",
  "右后门线束": "锚点：右后件车侧/门侧护套、连续过渡段、主干分支点、可视插接面。与左后门仅共用方法模板，不共用未验证数据。",
  "顶棚线束": "锚点：连续主干工装槽位、前/中/后连接点、短引出出口、固定件装入窗口。候选包：主干展线/定位包→短引出插接/定向包→固定件/保护件包。连续长主干不得为均工时硬拆。",
  "前保险杠线束": "锚点：主接口、横向主干工装板基准、重复护套组、可拆装接口完成状态、实际分支点。候选包：主干骨架定位包→重复护套组插接/定向包→保护件与接口状态确认包。",
  "后保险杠线束": "锚点：主接口、横向主干工装板基准、重复护套组A/B/C、可拆装接口完成状态、实际分支点。候选包：主干骨架定位包→重复护套组插接/定向包→保护件与接口状态确认包。"
};

const MODE_OPTIONS = [
  { value: "pure-kit", label: "纯KIT岗位" },
  { value: "kit-transfer-kit", label: "KIT传递岗+KIT岗位" },
  { value: "pure-sub", label: "纯SUB岗位" },
  { value: "sub-kit", label: "SUB岗位+纯KIT岗位" },
  { value: "sub-kit-transfer", label: "SUB岗位+纯KIT岗位+KIT传递岗位" }
];

function modeShort(label) {
  const s = String(label || "");
  if (s.includes("SUB") && s.includes("KIT传递")) return "SUB+KIT+传";
  if (s.includes("SUB") && s.includes("KIT")) return "SUB+KIT";
  if (s.includes("KIT传递")) return "KIT传+KIT";
  if (s.includes("纯KIT")) return "纯KIT";
  if (s.includes("纯SUB")) return "纯SUB";
  if (s.includes("SUB")) return "SUB";
  return "岗位";
}

function uniq(values) {
  return [...new Set(values.map((v) => String(v == null ? "" : v)).filter(Boolean))];
}

function sum(values) {
  return values.reduce((a, b) => a + (Number(b) || 0), 0);
}

function round(value, digits = 2) {
  const n = Number(value || 0);
  return Number.isFinite(n) ? Number(n.toFixed(digits)) : null;
}

// 包在TT校验与岗位分配下的“真实工时” = 该包在其最重配置下的装配工时。
// 一个线束有多个配置，单件产品只生产其中一个配置的导线，因此不把各配置工时叠加。
function pkgWeight(p) {
  const ct = p && p.configTime;
  if (ct && typeof ct === "object") {
    const vals = Object.values(ct).filter((v) => typeof v === "number" && v > 0);
    if (vals.length) return Math.max(...vals);
  }
  return (p && p.estimatedSeconds) || 0;
}

function normalizeTerminal(value) {
  const s = clean(value);
  return s === "-" ? "" : s;
}

function mergeWiresByW1(wires) {
  const groups = new Map();
  for (const w of wires) {
    const key = w.w1 || ("NO_W1:" + (w.drawingId || Math.random()));
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(w);
  }

  const merged = [];
  const conflicts = [];
  const warnings = [];

  for (const [key, group] of groups.entries()) {
    if (key.startsWith("NO_W1:")) {
      merged.push(...group);
      continue;
    }
    const base = { ...group[0] };
    const configUnion = new Set(base.configCodes || []);
    const configRows = new Map();

    for (const w of group) {
      for (const cfg of w.configCodes || []) {
        configUnion.add(cfg);
        if (!configRows.has(cfg)) configRows.set(cfg, []);
        configRows.get(cfg).push(w);
      }
    }

    for (const [cfg, rows] of configRows.entries()) {
      const first = rows[0];
      for (const row of rows.slice(1)) {
        const termDiff =
          normalizeTerminal(first.terminal1) !== normalizeTerminal(row.terminal1) ||
          normalizeTerminal(first.terminal2) !== normalizeTerminal(row.terminal2);
        const lenDiff = Number(first.length) !== Number(row.length);
        if (termDiff || lenDiff) {
          conflicts.push({
            w1: key,
            config: cfg,
            detail: `同一配置下W1【${key}】出现两次，端子或长度不一致（端子1:${first.terminal1}/${row.terminal1}，端子2:${first.terminal2}/${row.terminal2}，长度:${first.length}/${row.length}），判定为真冲突，不合并看板。`
          });
        } else if (
          clean(first.housing1) !== clean(row.housing1) ||
          clean(first.position1) !== clean(row.position1) ||
          clean(first.housing2) !== clean(row.housing2) ||
          clean(first.position2) !== clean(row.position2)
        ) {
          warnings.push({
            w1: key,
            config: cfg,
            detail: `同一配置下W1【${key}】端子/长度一致，但护套或孔位不同（${first.housing1}/${first.position1}-${first.housing2}/${first.position2} vs ${row.housing1}/${row.position1}-${row.housing2}/${row.position2}），已按首次记录合并，请人工复核。`
          });
        }
      }
    }

    base.config = {};
    for (const cfg of configUnion) base.config[cfg] = true;
    base.configCodes = [...configUnion];
    merged.push(base);
  }

  return { wires: merged, conflicts, warnings };
}

function mergeSpecialWires(mbomWires, specialWires = []) {
  const wires = mbomWires.slice();
  const conflicts = [];
  const warnings = [];
  const byDrawing = new Map();
  for (const w of wires) {
    if (w.drawingId) byDrawing.set(clean(w.drawingId).toUpperCase(), w);
  }

  for (const sw of specialWires) {
    const sheetName = sw.sourceSheet || "";
    const mergeable = /铝线|气囊线/.test(sheetName);
    const key = clean(sw.drawingId).toUpperCase();
    const existing = key ? byDrawing.get(key) : null;

    if (mergeable && existing) {
      // 合并配置，并做端子/长度核对
      const termDiff =
        normalizeTerminal(existing.terminal1) !== normalizeTerminal(sw.terminal1) ||
        normalizeTerminal(existing.terminal2) !== normalizeTerminal(sw.terminal2);
      const lenDiff = Number(existing.length) !== Number(sw.length);
      if (termDiff || lenDiff) {
        conflicts.push({
          w1: existing.w1 || existing.drawingId,
          config: "跨MBOM/特殊表",
          detail: `铝线/气囊线【${sw.drawingId}】与MBOM中【${existing.drawingId}】端子或长度不一致，判定为真冲突，不合并。`
        });
      } else {
        for (const cfg of sw.configCodes || []) {
          existing.config[cfg] = true;
          if (!existing.configCodes.includes(cfg)) existing.configCodes.push(cfg);
        }
        warnings.push({
          w1: existing.w1 || existing.drawingId,
          config: "跨MBOM/特殊表",
          detail: `特殊线束表【${sheetName}】中的【${sw.drawingId}】已与MBOM合并，配置已取并集。`
        });
      }
    } else {
      // 未匹配或以太网/FAKRA等：保留为独立导线，参与工艺制作但不参与MBOM核对
      wires.push(sw);
    }
  }

  return { wires, conflicts, warnings };
}

function lengthBand(length) {
  if (length == null) return "";
  if (length <= 1000) return "0-1000 mm";
  if (length <= 2000) return "1001-2000 mm";
  if (length <= 3000) return "2001-3000 mm";
  if (length <= 4000) return "3001-4000 mm";
  if (length <= 5000) return "4001-5000 mm";
  if (length <= 6000) return "5001-6000 mm";
  if (length <= 7000) return "6001-7000 mm";
  if (length <= 8000) return "7001-8000 mm";
  if (length <= 9000) return "8001-9000 mm";
  if (length <= 10000) return "9001-10000 mm";
  if (length <= 11000) return "10001-11000 mm";
  return ">11000 mm";
}

function findStandardEntry(entries, predicates) {
  for (const e of entries) {
    let ok = true;
    for (const p of predicates) {
      if (typeof p === "string") {
        if (!String(e.process + e.activity + e.comments).includes(p)) {
          ok = false;
          break;
        }
      } else if (p && p.test && !p.test(String(e.process + e.activity + e.comments))) {
        ok = false;
        break;
      }
    }
    if (ok) return e;
  }
  return null;
}

function getRouteStandardTimes(entries, wire) {
  const len = wire.length;
  const band = lengthBand(len);
  const take = findStandardEntry(entries, [
    "取/布导线",
    "remove wires from rack",
    band
  ]);
  const route = findStandardEntry(entries, [
    "取/布导线",
    "Route wires on board",
    band
  ]);
  return {
    takeTime: take ? take.time : null,
    routeTime: route ? route.time : null,
    takeRef: take ? take.process + "｜" + take.activity + "｜" + take.comments : "",
    routeRef: route ? route.process + "｜" + route.activity + "｜" + route.comments : ""
  };
}

function isSpecialWire(wire) {
  const s = wire.material + " " + wire.drawingId + " " + wire.customerNo;
  return /双绞|屏蔽|气囊|铝线|以太网|FAKRA|外购/.test(s);
}

function isSpliceCode(value) {
  return /^(SP|SC)[0-9A-Za-z]/.test(clean(value));
}

function isUltrasonicCandidate(wire) {
  const s = wire.material + " " + wire.drawingId + " " + wire.customerNo + " " + (wire.notes || "");
  if (/双绞压接|并头|焊点|超声波|USW|在线超声波/.test(s)) return true;
  if (isSpliceCode(wire.housing1) || isSpliceCode(wire.housing2)) return true;
  return false;
}

function getPlugEndText(wire, forcedOfflineHousings = []) {
  const forced = new Set((forcedOfflineHousings || []).map(clean).filter(Boolean));
  const h1 = clean(wire.housing1) && clean(wire.housing1) !== "-";
  const h2 = clean(wire.housing2) && clean(wire.housing2) !== "-";
  const sp1 = h1 && isSpliceCode(wire.housing1);
  const sp2 = h2 && isSpliceCode(wire.housing2);
  const real1 = h1 && !sp1;
  const real2 = h2 && !sp2;
  const end1 = real1 && forced.has(clean(wire.housing1)) ? "A端【强制线下插接完毕】" : real1 ? "A端插接（护套1）" : "";
  const end2 = real2 && forced.has(clean(wire.housing2)) ? "B端【强制线下插接完毕】" : real2 ? "B端插接（护套2）" : "";
  const splicePart = [];
  if (sp1) splicePart.push(`SP/SC端1【${clean(wire.housing1)}】`);
  if (sp2) splicePart.push(`SP/SC端2【${clean(wire.housing2)}】`);
  const parts = [end1, end2, splicePart.join("＋")].filter(Boolean);
  if (parts.length) return parts.join("＋");
  return "不插接/待确认";
}

function getPlugTime(entries, wire, sealPresent, special) {
  const base = sealPresent
    ? (findStandardEntry(entries, ["Pluging", "sealed terminals to unsealed connector"]) || {}).time || 4.0
    : (findStandardEntry(entries, ["Pluging", "Unsealed terminals to unsealed connector"]) || {}).time || 2.5;
  const twistComp = special
    ? (findStandardEntry(entries, ["Pluging", "shielded / Twisted pluging compensation"]) || {}).time || 1.6
    : 0;
  return {
    baseTime: base,
    twistComp,
    total: round(base + twistComp),
    ref: sealPresent
      ? "标准工时：带密封端子插接/未密封塑件（默认，待确认）"
      : "标准工时：不带密封端子插接/不带密封塑件（默认，待确认）"
  };
}

function parseInputs(files) {
  const standard = parseStandardHours(files.standard.buffer, files.standard.originalname);
  const ebom = parseEBOM(files.ebom.buffer, files.ebom.originalname);
  const mbom = parseMBOM(files.mbom.buffer, files.mbom.originalname);
  return { standard, ebom, mbom };
}

async function parsePdfInput(files) {
  if (!files.pdf || !files.pdf.buffer) {
    return {
      filename: "未提供",
      numPages: 0,
      textLength: 0,
      text: "",
      keywordHits: {},
      status: "未读取到工艺PDF"
    };
  }
  const data = await parsePDF(files.pdf.buffer, files.pdf.originalname);
  // 关键词统计已用完整文本完成；返回/导出的 text 截断，避免超大PDF文本拖累结果JSON体积
  const MAX_PDF_TEXT = 100000;
  return {
    ...data,
    text: data.text.slice(0, MAX_PDF_TEXT),
    textLength: data.text.length
  };
}

function buildWireRows(wires) {
  return wires.map((w, idx) => ({
    idx: idx + 1,
    w3: w.w3,
    w2: w.w2,
    w1: w.w1,
    material: w.material,
    drawingId: w.drawingId,
    customerNo: w.customerNo,
    jettyNo: w.jettyNo,
    spec: w.spec,
    color: w.color,
    length: w.length,
    unit: w.lengthUnit,
    terminal1: w.terminal1,
    seal1: w.seal1,
    terminal2: w.terminal2,
    seal2: w.seal2,
    housing1: w.housing1,
    position1: w.position1,
    housing2: w.housing2,
    position2: w.position2,
    configs: w.configCodes.join(" / "),
    status: w.status || "已读取"
  }));
}

function buildHousingMatrix(wires) {
  const map = new Map();
  for (const w of wires) {
    const h1 = clean(w.housing1);
    const h2 = clean(w.housing2);
    if (!h1 || !h2 || h1 === "-" || h2 === "-" || h1 === h2) continue;
    if (isSpliceCode(h1) || isSpliceCode(h2)) continue;
    const a = h1 <= h2 ? h1 : h2;
    const b = h1 <= h2 ? h2 : h1;
    const key = a + "|" + b;
    if (!map.has(key)) {
      map.set(key, { housingA: a, housingB: b, count: 0, wires: [], configs: new Set() });
    }
    const row = map.get(key);
    row.count += 1;
    row.wires.push(w.w1 || w.drawingId);
    for (const c of w.configCodes) row.configs.add(c);
  }
  return [...map.values()]
    .map((row) => ({
      housingA: row.housingA,
      housingB: row.housingB,
      count: row.count,
      wires: row.wires.join("、"),
      configs: [...row.configs].join(" / ")
    }))
    .sort((a, b) => b.count - a.count || a.housingA.localeCompare(b.housingA));
}

function buildPackages(wires, configs, entries) {
  const groups = new Map();
  for (const w of wires) {
    const key = w.w3 || w.w2 || "W1:" + (w.w1 || w.drawingId);
    if (!groups.has(key)) {
      groups.set(key, {
        key,
        kind: w.w3 ? "W3" : w.w2 ? "W2" : "W1直挂",
        name: w.w3 || w.w2 || w.w1 || w.drawingId,
        wires: []
      });
    }
    groups.get(key).wires.push(w);
  }

  const packages = [];
  let seq = 0;
  for (const group of groups.values()) {
    seq += 1;
    const ws = group.wires;
    const housingMap = new Map();
    const spliceMap = new Map();
    for (const w of ws) {
      for (const h of [w.housing1, w.housing2]) {
        const hh = clean(h);
        if (!hh || hh === "-") continue;
        if (isSpliceCode(hh)) {
          spliceMap.set(hh, (spliceMap.get(hh) || 0) + 1);
        } else {
          housingMap.set(hh, (housingMap.get(hh) || 0) + 1);
        }
      }
    }
    const housingEntries = [...housingMap.entries()].sort((a, b) => b[1] - a[1]);
    const spliceEntries = [...spliceMap.entries()].sort((a, b) => b[1] - a[1]);
    const anchor = housingEntries.length ? housingEntries[0][0] : (spliceEntries.length ? spliceEntries[0][0] : "");
    const wireCount = ws.length;
    const totalLength = sum(ws.map((w) => w.length || 0));
    const maxLength = Math.max(0, ...ws.map((w) => w.length || 0));

    let routeType = "待定";
    if (wireCount <= 5 && housingEntries.length <= 3 && maxLength <= 3000) {
      routeType = "固定KIT候选";
    } else if (wireCount > 12 || maxLength > 6000 || housingEntries.length > 8) {
      routeType = "SUB滑板候选";
    }

    const pkg = {
      id: "PKG-" + String(seq).padStart(4, "0"),
      kind: group.kind,
      name: group.name,
      key: group.key,
      wireCount,
      totalLength: round(totalLength, 0),
      maxLength: round(maxLength, 0),
      housings: housingEntries.map(([h, c]) => `${h}(${c})`).join("、"),
      housingCount: housingEntries.length,
      spliceCodes: spliceEntries.map(([h, c]) => `${h}(${c})`).join("、"),
      spliceCount: spliceEntries.length,
      anchor,
      configs: uniq(ws.flatMap((w) => w.configCodes)),
      routeType,
      status: "【候选方案，待工艺验证】",
      wires: ws,
      estimatedSeconds: null,
      configTime: {}
    };

    let total = 0;
    for (const w of ws) {
      const rt = getRouteStandardTimes(entries, w);
      const special = isSpecialWire(w);
      const h1Real = clean(w.housing1) && clean(w.housing1) !== "-" && !isSpliceCode(w.housing1);
      const h2Real = clean(w.housing2) && clean(w.housing2) !== "-" && !isSpliceCode(w.housing2);
      const plug1 = h1Real ? getPlugTime(entries, w, !!(clean(w.seal1) && clean(w.seal1) !== "-"), special) : { total: 0, baseTime: 0, twistComp: 0, ref: "SP/SC端不计插接工时" };
      const plug2 = h2Real ? getPlugTime(entries, w, !!(clean(w.seal2) && clean(w.seal2) !== "-"), special) : { total: 0, baseTime: 0, twistComp: 0, ref: "SP/SC端不计插接工时" };
      const spliceEnds = [w.housing1, w.housing2].filter((h) => isSpliceCode(h)).length;
      const wireTime = (rt.takeTime || 0) + (rt.routeTime || 0) + plug1.total + plug2.total;
      w._time = {
        rt,
        plug1,
        plug2,
        spliceEnds,
        spliceTime: spliceEnds > 0 ? null : null,
        spliceTimeNote: spliceEnds > 0 ? "SP/SC压接点工时待确认（在线超声波/冷压接/热缩）" : "",
        wireTime: round(wireTime, 2)
      };
      total += wireTime;
    }

    let connTime = 0;
    for (const [h, c] of housingEntries) {
      const maxPos = Math.max(
        0,
        ...ws
          .filter((w) => clean(w.housing1) === h)
          .map((w) => number(w.position1) || 0),
        ...ws
          .filter((w) => clean(w.housing2) === h)
          .map((w) => number(w.position2) || 0)
      );
      let t = 1.7;
      if (maxPos > 40) t = 5.4;
      else if (maxPos > 6) t = 4.0;
      connTime += t;
    }
    pkg.connectorPlacementSeconds = round(connTime, 2);
    pkg.estimatedSeconds = round(total + connTime, 2);
    pkg.estimatedSecondsNote = "估算值，来自标准工时；实际以正式工时分摊与现场验证为准";
    for (const cfg of configs) {
      const cfgWires = ws.filter((w) => w.config[cfg]);
      const cfgTime =
        sum(cfgWires.map((w) => w._time?.wireTime || 0)) + connTime;
      pkg.configTime[cfg] = round(cfgTime, 2);
    }

    packages.push(pkg);
  }

  return packages.sort((a, b) => a.id.localeCompare(b.id, undefined, { numeric: true }));
}

function buildPositionMatrix(packages) {
  const rows = [];
  const seen = new Map();
  for (const pkg of packages) {
    for (const w of pkg.wires) {
      if (w.housing1 && w.housing1 !== "-" && !isSpliceCode(w.housing1) && w.position1 !== "") {
        const key = w.housing1 + "|" + w.position1;
        rows.push({
          pkgId: pkg.id,
          housing: w.housing1,
          position: w.position1,
          wire: w.w1 || w.drawingId,
          terminal: w.terminal1 || "",
          seal: clean(w.seal1) && clean(w.seal1) !== "-" ? w.seal1 : "",
          configs: w.configCodes.join(" / "),
          status: "待唯一责任确认"
        });
        if (seen.has(key)) seen.get(key).push(w.w1 || w.drawingId);
        else seen.set(key, [w.w1 || w.drawingId]);
      }
      if (w.housing2 && w.housing2 !== "-" && !isSpliceCode(w.housing2) && w.position2 !== "") {
        const key = w.housing2 + "|" + w.position2;
        rows.push({
          pkgId: pkg.id,
          housing: w.housing2,
          position: w.position2,
          wire: w.w1 || w.drawingId,
          terminal: w.terminal2 || "",
          seal: clean(w.seal2) && clean(w.seal2) !== "-" ? w.seal2 : "",
          configs: w.configCodes.join(" / "),
          status: "待唯一责任确认"
        });
        if (seen.has(key)) seen.get(key).push(w.w1 || w.drawingId);
        else seen.set(key, [w.w1 || w.drawingId]);
      }
    }
  }
  for (const row of rows) {
    const key = row.housing + "|" + row.position;
    const list = [...new Set(seen.get(key) || [])];
    if (list.length > 1) {
      row.status = `【存在冲突】同一孔位出现${list.length}根导线：${list.join("、")}`;
    }
  }
  return rows;
}

function generateDotCodes(n) {
  const codes = [];
  const palette = APPROVED_DOT_COLORS;

  function gen(totalDots, startIndex, currentCounts) {
    if (totalDots === 0) {
      const parts = [];
      for (let i = 0; i < palette.length; i++) {
        if ((currentCounts[i] || 0) > 0) {
          parts.push(`${palette[i]}×${currentCounts[i]}`);
        }
      }
      codes.push(parts.join("+"));
      return;
    }
    for (let i = startIndex; i < palette.length; i++) {
      currentCounts[i] = (currentCounts[i] || 0) + 1;
      gen(totalDots - 1, i, currentCounts);
      currentCounts[i] -= 1;
    }
  }

  for (let dots = 1; dots <= 4; dots++) {
    gen(dots, 0, []);
  }
  return codes;
}

const DOT_CODE_POOL = generateDotCodes(500);

function buildTimeMatrix(packages, configs, tt) {
  const rows = [];
  for (const pkg of packages) {
    for (const cfg of configs) {
      const seconds = pkg.configTime[cfg];
      if (seconds == null) continue; // 该包不包含此配置，不输出空白行
      const load = tt && tt > 0 ? round((seconds / tt) * 100, 1) : null;
      rows.push({
        pkgId: pkg.id,
        pkgName: pkg.kind + ":" + pkg.name,
        config: cfg,
        estimatedSeconds: seconds,
        tt: tt || "",
        loadPercent: load,
        workerSuggestion: tt && tt > 0 ? Math.max(1, Math.ceil(seconds / tt)) : "",
        note: "估算值，需正式工时确认"
      });
    }
  }
  return rows;
}

function buildIssues(mbom, ebom, standard, pdf, packages, dotIssues, positionRows) {
  const issues = [];
  const missingHousing = mbom.wires.filter(
    (w) => (!w.housing1 || w.housing1 === "-") && (!w.housing2 || w.housing2 === "-")
  ).length;
  const missingPosition = mbom.wires.filter((w) =>
    (w.housing1 && w.housing1 !== "-" && !isSpliceCode(w.housing1) && !w.position1) ||
    (w.housing2 && w.housing2 !== "-" && !isSpliceCode(w.housing2) && !w.position2)
  ).length;
  const missingTerminal = mbom.wires.filter((w) =>
    (w.housing1 && w.housing1 !== "-" && !isSpliceCode(w.housing1) && (!w.terminal1 || w.terminal1 === "-")) ||
    (w.housing2 && w.housing2 !== "-" && !isSpliceCode(w.housing2) && (!w.terminal2 || w.terminal2 === "-"))
  ).length;

  if (missingHousing) issues.push({ category: "MBOM数据完整性", detail: `${missingHousing}根导线缺少护套信息，不能建立完整护套关联。` });
  if (missingPosition) issues.push({ category: "MBOM数据完整性", detail: `${missingPosition}个护套端缺少孔位信息。` });
  if (missingTerminal) issues.push({ category: "MBOM数据完整性", detail: `${missingTerminal}个护套端缺少端子信息。` });

  const w1Dup = new Map();
  for (const w of mbom.wires) {
    if (!w.w1) continue;
    if (!w1Dup.has(w.w1)) w1Dup.set(w.w1, []);
    w1Dup.get(w.w1).push(`${w.w3 || "无W3"}/${w.w2 || "无W2"}`);
  }
  const duplicatedW1 = [...w1Dup.entries()].filter(([, v]) => v.length > 1);
  if (duplicatedW1.length) {
    const sample = duplicatedW1.slice(0, 5).map(([id, v]) => `${id}(${v.join("、")})`).join("；");
    issues.push({
      category: "MBOM跨W3/W2复用",
      detail: `共${duplicatedW1.length}个W1编号出现在多个组合中，已按各组合分别纳入候选包；请确认属于配置拆分复用而非重复。示例：${sample}。`
    });
  }

  const posConflicts = positionRows.filter((r) => r.status.includes("存在冲突"));
  for (const r of posConflicts.slice(0, 20)) {
    issues.push({ category: "孔位冲突", detail: `${r.housing}/${r.position}：${r.status}` });
  }

  for (const dot of dotIssues) {
    // 来自 buildStationDotMatrix 的 issue（{stationNo, stationName, color, count, detail}）
    issues.push({ category: "同色线编码容量", detail: dot.detail || `岗位${dot.stationNo || ""}颜色${dot.color || ""}同色线数量超出编码容量，需现场确认。` });
  }

  if (pdf.status === "未读取到工艺PDF" || !pdf.textLength) {
    issues.push({ category: "PDF图纸", detail: "未读取到工艺PDF文本，无法进行图纸关键词核对。" });
  } else {
    issues.push({ category: "PDF图纸", detail: "PDF已提取文本，但图形、尺寸线、分支走向等版面关系仍需人工或后续AI核对。" });
  }

  issues.push({ category: "版本", detail: "文件版本/生效日期未在文件内自动读取，需人工补充后关闭。" });
  issues.push({ category: "TT", detail: "TT/产量未在四个文件中识别，需用户输入或在输出中补充。" });
  issues.push({ category: "预装工作包", detail: "当前工作包为数据驱动候选包（W3/W2/W1层级），必须按V2.6.1规则用护套关联、待插端、保护件顺序、受控半成品和TT验证后转正式。" });
  issues.push({ category: "打圈工时", detail: "KIT/SUB输出岗位是否增加50秒打圈工时，需根据实际受控输出判定并查重。" });
  issues.push({ category: "打圈工时", detail: "KIT/SUB输出岗位是否增加50秒打圈工时，需根据实际受控输出判定并查重。" });
  issues.push({ category: "布局", detail: "未提供现场布局图，人员站位、物流、缓存仅能输出候选。" });

  // 工时源缺失统计：取/布线动作或长度分段未匹配、插接工时使用默认值兜底
  let missingRouteCount = 0;
  let defaultPlugCount = 0;
  for (const p of packages) {
    for (const w of p.wires || []) {
      const t = w._time || {};
      if (t.rt && (t.rt.takeTime == null || t.rt.routeTime == null)) missingRouteCount += 1;
      for (const side of [t.plug1, t.plug2]) {
        if (side && /默认/.test(String(side.ref || ""))) defaultPlugCount += 1;
      }
    }
  }
  if (missingRouteCount) {
    issues.push({
      category: "工时源",
      detail: `${missingRouteCount}根导线未匹配到标准工时中的取/布线动作或对应长度分段，其取线/布线工时为空（表格对应列空白），请补充标准工时表对应分段后再复核TT。`
    });
  }
  if (defaultPlugCount) {
    issues.push({
      category: "工时源",
      detail: `${defaultPlugCount}个插接端未匹配到标准工时插接动作，已使用默认估算值（带密封4.0s/不带密封2.5s/屏蔽补偿1.6s）参与估算，【未找到工时源】，需以正式工时分摊复核。`
    });
  }

  return issues;
}

function buildLedger(files, mbom, ebom, standard, pdf) {
  const items = [
    {
      fileType: "标准工时文件",
      fileName: standard.filename,
      sizeKB: files.standard.buffer ? Math.round(files.standard.buffer.length / 1024) : 0,
      status: "已读取",
      info: `${standard.rowCount}条标准工时要素`
    },
    {
      fileType: "EBOM文件",
      fileName: ebom.filename,
      sizeKB: files.ebom.buffer ? Math.round(files.ebom.buffer.length / 1024) : 0,
      status: "已读取",
      info: `${ebom.rowCount}条物料；特殊线束子表：${ebom.specialSheets.map((s) => `${s.sheetName}(${s.rowCount})`).join("、") || "无"}`
    },
    {
      fileType: "MBOM/Cutting文件",
      fileName: mbom.filename,
      sizeKB: files.mbom.buffer ? Math.round(files.mbom.buffer.length / 1024) : 0,
      status: "已读取",
      info: `${mbom.rowCount}根导线；配置：${mbom.configs.join("、")}`
    }
  ];
  if (files.pdf && files.pdf.buffer) {
    items.push({
      fileType: "工艺图纸PDF",
      fileName: pdf.filename,
      sizeKB: Math.round(files.pdf.buffer.length / 1024),
      status: pdf.status,
      info: `${pdf.numPages}页；文本${pdf.textLength}字符`
    });
  } else {
    items.push({
      fileType: "工艺图纸PDF",
      fileName: "未提供",
      sizeKB: 0,
      status: "未读取到",
      info: "缺少PDF将影响图纸识别与保护件/分支核对"
    });
  }
  return items;
}

function buildPDFKeywordTable(pdf) {
  if (!pdf.keywordHits || !Object.keys(pdf.keywordHits).length) {
    return [];
  }
  return Object.entries(pdf.keywordHits)
    .sort((a, b) => b[1] - a[1])
    .map(([keyword, count]) => ({ keyword, count }));
}

function buildFileAnalysis(mbom) {
  const housings = new Map();
  const splices = new Map();
  for (const w of mbom.wires) {
    const ends = [
      [w.housing1, w.position1, w.configCodes],
      [w.housing2, w.position2, w.configCodes]
    ];
    for (const [h, pos, cfgCodes] of ends) {
      if (!h || h === "-") continue;
      if (isSpliceCode(h)) {
        if (!splices.has(h)) splices.set(h, { code: h, count: 0, configs: new Set() });
        splices.get(h).count += 1;
        for (const c of cfgCodes) splices.get(h).configs.add(c);
      } else {
        if (!housings.has(h)) housings.set(h, { code: h, count: 0, positions: 0 });
        housings.get(h).count += 1;
        if (pos !== "") housings.get(h).positions += 1;
      }
    }
  }

  const configs = mbom.configs;
  const onlineUltrasonicGroupsPerConfig = {};
  for (const cfg of configs) {
    const set = new Set();
    for (const w of mbom.wires) {
      if (!w.config[cfg]) continue;
      for (const h of [w.housing1, w.housing2]) {
        if (isSpliceCode(h)) set.add(clean(h));
      }
    }
    onlineUltrasonicGroupsPerConfig[cfg] = set.size;
  }

  return {
    totalWires: mbom.wires.length,
    uniqueW1: mbom.uniqueW1 || new Set(mbom.wires.map((w) => w.w1).filter(Boolean)).size,
    configs,
    housingCount: housings.size,
    housingList: [...housings.values()].map((h) => ({ code: h.code, endCount: h.count, positionCount: h.positions })).sort((a, b) => b.endCount - a.endCount),
    spliceCount: splices.size,
    spliceList: [...splices.values()].map((s) => ({ code: s.code, endCount: s.count, configs: [...s.configs].join(" / ") })).sort((a, b) => b.endCount - a.endCount),
    onlineUltrasonicGroupsPerConfig
  };
}

function buildUltrasonicRules(wires, configs, noOnlineUltrasonicSplices = [], maxGroupsPerConfig = null, maxTotalGroups = null) {
  const userBlocked = new Set((noOnlineUltrasonicSplices || []).map(clean).filter(Boolean));
  const allSplices = new Map();
  for (const w of wires) {
    for (const h of [w.housing1, w.housing2]) {
      if (!isSpliceCode(h)) continue;
      const code = clean(h);
      if (!allSplices.has(code)) allSplices.set(code, { code, count: 0, configs: new Set() });
      allSplices.get(code).count += 1;
      for (const c of w.configCodes) allSplices.get(code).configs.add(c);
    }
  }

  const allowedByConfig = {};
  const groupsPerConfig = {};
  const blockedByLimit = new Set();
  for (const cfg of configs) {
    const allowed = [...allSplices.values()]
      .filter((s) => s.configs.has(cfg) && !userBlocked.has(s.code))
      .sort((a, b) => b.count - a.count || a.code.localeCompare(b.code));
    groupsPerConfig[cfg] = allowed.length;
    let finalAllowed = allowed;
    if (maxGroupsPerConfig != null && maxGroupsPerConfig > 0 && allowed.length > maxGroupsPerConfig) {
      finalAllowed = allowed.slice(0, maxGroupsPerConfig);
      for (const s of allowed.slice(maxGroupsPerConfig)) blockedByLimit.add(s.code);
    }
    allowedByConfig[cfg] = finalAllowed.map((s) => s.code);
  }

  const globallyAllowed = new Set(Object.values(allowedByConfig).flat());
  const blockedByLimitTotal = new Set();
  const remainingSorted = [...globallyAllowed]
    .map((code) => allSplices.get(code))
    .sort((a, b) => b.count - a.count || a.code.localeCompare(b.code));
  if (maxTotalGroups != null && maxTotalGroups > 0 && remainingSorted.length > maxTotalGroups) {
    for (const s of remainingSorted.slice(maxTotalGroups)) blockedByLimitTotal.add(s.code);
  }

  return {
    userBlocked: [...userBlocked],
    allowedByConfig,
    blockedByLimit: [...blockedByLimit],
    blockedByLimitTotal: [...blockedByLimitTotal],
    blocked: [...new Set([...blockedByLimit, ...blockedByLimitTotal])],
    groupsPerConfig,
    maxGroupsPerConfig: maxGroupsPerConfig ?? null,
    maxTotalGroups: maxTotalGroups ?? null,
    totalAllowedGroups: Math.min(remainingSorted.length, maxTotalGroups != null && maxTotalGroups > 0 ? maxTotalGroups : remainingSorted.length)
  };
}

function allocateStations(packages, maxStations, tt, modeLabel) {
  if (!maxStations || maxStations <= 0) return [];
  const count = Math.min(maxStations, Math.max(1, packages.length));
  const bins = Array.from({ length: count }, (_, i) => ({
    stationNo: i + 1,
    seconds: 0,
    wires: 0,
    packageIds: [],
    configs: new Set()
  }));

  const items = packages.slice().sort((a, b) => pkgWeight(b) - pkgWeight(a));
  for (const p of items) {
    const bin = bins.reduce((best, b) => (b.seconds <= best.seconds ? b : best), bins[0]);
    bin.seconds += pkgWeight(p);
    bin.wires += p.wireCount || 0;
    bin.packageIds.push(p.id);
    for (const c of p.configs) bin.configs.add(c);
  }

  return bins.map((b) => ({
    stationNo: b.stationNo,
    modeLabel: modeLabel || "待定",
    packageIds: b.packageIds.join("、"),
    packageCount: b.packageIds.length,
    wireCount: b.wires,
    totalSeconds: round(b.seconds, 2),
    tt: tt || "",
    loadPercent: tt && tt > 0 ? round((b.seconds / tt) * 100, 1) : null,
    status: tt && tt > 0 && b.seconds > tt ? "【超TT，需拆分或增岗】" : "【候选方案，待工艺验证】",
    configs: [...b.configs].join(" / ")
  }));
}

function packPackages(items, stationCount, tt, stationType) {
  if (!stationCount || stationCount <= 0) return [];
  const count = Math.min(stationCount, Math.max(1, items.length));
  const bins = Array.from({ length: count }, (_, i) => ({
    stationNo: 0,
    seconds: 0,
    wires: 0,
    packageIds: [],
    configs: new Set()
  }));
  const sorted = items.slice().sort((a, b) => pkgWeight(b) - pkgWeight(a));
  for (const p of sorted) {
    const bin = bins.reduce((best, b) => (b.seconds <= best.seconds ? b : best), bins[0]);
    bin.seconds += pkgWeight(p);
    bin.wires += p.wireCount || 0;
    bin.packageIds.push(p.id);
    for (const c of p.configs) bin.configs.add(c);
  }
  // 超限桶平衡：把超TT桶中最大的可移包移到最空桶（移后双方都不超TT才执行），减少轻微超限
  const ttLimit = Number(tt) || 0;
  const pkgSecMap = new Map(items.map((p) => [p.id, pkgWeight(p)]));
  const pkgWireMap = new Map(items.map((p) => [p.id, p.wireCount || 0]));
  const pkgCfgMap = new Map(items.map((p) => [p.id, p.configs || []]));
  if (ttLimit > 0) {
    let improved = true;
    let guard = 0;
    while (improved && guard < 200) {
      improved = false;
      guard += 1;
      for (let i = 0; i < bins.length; i++) {
        if (bins[i].seconds <= ttLimit) continue;
        let j = 0;
        for (let k = 1; k < bins.length; k++) {
          if (bins[k].seconds < bins[j].seconds) j = k;
        }
        if (j === i) continue;
        let bestMove = null;
        for (const pid of bins[i].packageIds) {
          const sec = pkgSecMap.get(pid) || 0;
          if (bins[i].seconds - sec <= ttLimit && bins[j].seconds + sec <= ttLimit) {
            if (!bestMove || sec > bestMove.sec) bestMove = { pid, sec };
          }
        }
        if (bestMove) {
          bins[i].packageIds = bins[i].packageIds.filter((p) => p !== bestMove.pid);
          bins[i].seconds -= bestMove.sec;
          bins[i].wires -= pkgWireMap.get(bestMove.pid) || 0;
          bins[j].packageIds.push(bestMove.pid);
          bins[j].seconds += bestMove.sec;
          bins[j].wires += pkgWireMap.get(bestMove.pid) || 0;
          for (const c of pkgCfgMap.get(bestMove.pid) || []) bins[j].configs.add(c);
          improved = true;
        }
      }
    }
  }
  return bins.map((b) => ({
    stationNo: 0,
    stationType,
    modeLabel: stationType,
    packageIds: b.packageIds.join("、"),
    packageCount: b.packageIds.length,
    wireCount: b.wires,
    totalSeconds: round(b.seconds, 2),
    tt: tt || "",
    loadPercent: tt && tt > 0 ? round((b.seconds / tt) * 100, 1) : null,
    status: tt && tt > 0 && b.seconds > tt ? "【超TT，需拆分或增岗】" : "【候选方案，待工艺验证】",
    configs: [...b.configs].join(" / ")
  }));
}

function labelTransferBins(bins) {
  if (bins.length <= 1) {
    return bins.map((b) => ({ ...b, stationType: "KIT岗位", modeLabel: "纯KIT岗位" }));
  }
  return bins.map((b, i) => {
    const isLast = i === bins.length - 1;
    return {
      ...b,
      stationType: isLast ? "KIT岗位" : "KIT传递岗",
      modeLabel: isLast ? "纯KIT岗位" : "KIT传递岗"
    };
  });
}

function allocateStationsByMode(packages, maxStations, tt, modeLabel, modeValue, maxSubFrames = null) {
  if (!maxStations || maxStations <= 0) return [];
  if (modeValue === "pure-kit") {
    return packPackages(packages, maxStations, tt, "纯KIT岗位").map((b, i) => ({ ...b, stationNo: i + 1 }));
  }
  if (modeValue === "kit-transfer-kit") {
    return labelTransferBins(packPackages(packages, maxStations, tt, "KIT传递岗+KIT岗位")).map((b, i) => ({ ...b, stationNo: i + 1 }));
  }
  if (modeValue === "pure-sub") {
    const count = maxSubFrames && maxSubFrames > 0
      ? Math.max(maxStations, Math.ceil(packages.length / maxSubFrames))
      : maxStations;
    return packPackages(packages, count, tt, "纯SUB岗位").map((b, i) => ({ ...b, stationNo: i + 1 }));
  }
  if (modeValue === "sub-kit" || modeValue === "sub-kit-transfer") {
    const totalSec = sum(packages.map((p) => pkgWeight(p))) || 1;
    const subItems = packages.filter((p) => p.routeType === "SUB滑板候选");
    const kitItems = packages.filter((p) => p.routeType !== "SUB滑板候选");
    const subSec = sum(subItems.map((p) => pkgWeight(p)));
    const kitSec = sum(kitItems.map((p) => pkgWeight(p)));
    let subCount = Math.max(1, Math.min(maxStations - 1, Math.round((maxStations * subSec) / totalSec)));
    if (maxSubFrames && maxSubFrames > 0) {
      subCount = Math.max(subCount, Math.ceil(subItems.length / maxSubFrames));
    }
    const kitCount = Math.max(1, maxStations - subCount);
    let subBins = packPackages(subItems, subCount, tt, "SUB岗位");
    let kitBins = packPackages(kitItems, kitCount, tt, modeValue === "sub-kit" ? "纯KIT岗位" : "KIT传递岗+KIT岗位");
    if (modeValue === "sub-kit-transfer") kitBins = labelTransferBins(kitBins);
    const all = [...subBins, ...kitBins];
    return all.map((b, i) => ({ ...b, stationNo: i + 1 }));
  }
  return packPackages(packages, maxStations, tt, modeLabel).map((b, i) => ({ ...b, stationNo: i + 1 }));
}

// ==================== 同岗位护套分组（一组=一个岗位，组间强制分岗） ====================
// 输入：sameStationGroups = [{ name, housings, mode }]
//   mode: pure-kit | kit-transfer | sub
// 语义：
//   1. 同一组内的护套 → 必须全部落在同一个岗位（强制合包，吸收包含该组护套的所有岗位）；
//   2. 不同组之间 → 不得在同一个岗位（若自动分配把多组护套放进同一岗位，按组拆分该岗位）；
//   3. 每组合并后的岗位采用该组所选岗位类型（纯KIT/KIT传递/SUB），打圈工时按该类型规则计算；
//   4. 同一个护套不允许出现在两个组；横跨两组的单个工作包无法自动拆分，输出冲突交人工确认。

const GROUP_MODE_LABELS = {
  "pure-kit": { stationType: "KIT岗位", modeLabel: "纯KIT岗位" },
  "kit-transfer": { stationType: "KIT传递岗", modeLabel: "KIT传递岗" },
  "sub": { stationType: "SUB岗位", modeLabel: "SUB岗位" }
};

// 各全局预装模式下允许的组岗位类型
const GROUP_MODE_ALLOWED_BY_GLOBAL = {
  "": ["pure-kit", "kit-transfer", "sub"],
  "pure-kit": ["pure-kit"],
  "kit-transfer-kit": ["pure-kit", "kit-transfer"],
  "pure-sub": ["sub"],
  "sub-kit": ["pure-kit", "sub"],
  "sub-kit-transfer": ["pure-kit", "kit-transfer", "sub"]
};

function parseSameStationGroups(groups, issues) {
  const parsed = [];
  (groups || []).forEach((g, gi) => {
    const housings = new Set(String((g && g.housings) || "").split(/[,，;；\n\r]+/).map(clean).filter(Boolean));
    if (!housings.size) return;
    const mode = g && g.mode && GROUP_MODE_LABELS[g.mode] ? g.mode : "pure-kit";
    parsed.push({
      idx: gi,
      name: clean(g && g.name) || `同岗组${String(gi + 1).padStart(2, "0")}`,
      housings,
      mode,
      typeLabel: GROUP_MODE_LABELS[mode]
    });
  });
  const owner = new Map();
  for (const g of parsed) {
    for (const h of g.housings) {
      if (owner.has(h)) {
        issues.push({
          category: "同岗位护套分组",
          detail: `护套【${h}】同时出现在“${owner.get(h)}”与“${g.name}”两组；按组输入不允许重复，请拆分或删除。`
        });
      } else {
        owner.set(h, g.name);
      }
    }
  }
  return parsed;
}

function applySameStationGroups(stationAllocation, parsedGroups, packages, issues, preassemblyMode, overTtMode = "force-same", tt = null, loopTimes = {}) {
  if (!parsedGroups.length) return { stationAllocation, ttRows: [] };
  const packageMap = new Map(packages.map((p) => [p.id, p]));
  const ttNum = Number(tt) || 0;
  const secOf = (pid) => (packageMap.get(pid) && pkgWeight(packageMap.get(pid))) || 0;
  const ttRows = [];
  // 组岗位类型对应的打圈工时估算（与 buildStationDetails 的规则对齐；
  // KIT传递/SUB 的"最后岗位"判断依赖整体序列，此处按保守口径估算并在明细中注明）
  const groupLoopEstimate = (g) => {
    const lt = loopTimes || {};
    if (g.mode === "pure-kit") return Number(lt.singleKit) || 0;
    if (g.mode === "kit-transfer") return Number(lt.kitTransferMiddle) || 0;
    if (g.mode === "sub") return Number(lt.subLast) || 0;
    return 0;
  };

  // 组岗位类型与全局预装模式兼容性提示（仍按组所选类型执行）
  const allowed = GROUP_MODE_ALLOWED_BY_GLOBAL[preassemblyMode] || GROUP_MODE_ALLOWED_BY_GLOBAL[""];
  for (const g of parsedGroups) {
    if (!allowed.includes(g.mode)) {
      issues.push({
        category: "同岗位护套分组",
        detail: `组【${g.name}】选择岗位类型【${g.typeLabel.modeLabel}】，与全局预装模式【${preassemblyMode || "未选择"}】不一致，仍按组所选类型执行，【待确认】。`
      });
    }
  }

  // 工作包 → 组归属
  const pkgGroups = new Map(); // pid -> Set<gi>
  for (const p of packages) {
    const gs = new Set();
    for (const w of p.wires || []) {
      for (const h of [w.housing1, w.housing2]) {
        const hh = clean(h);
        if (!hh) continue;
        parsedGroups.forEach((g, gi) => { if (g.housings.has(hh)) gs.add(gi); });
      }
    }
    if (gs.size) pkgGroups.set(p.id, gs);
  }

  // 横跨两组的工作包：无法自动拆分，保留在原岗位并输出冲突
  for (const [pid, gs] of pkgGroups) {
    if (gs.size > 1) {
      issues.push({
        category: "同岗位护套分组",
        detail: `工作包【${pid}】同时包含${[...gs].map((gi) => parsedGroups[gi].name).join("、")}组的护套，无法自动按组分岗，需人工确认。`
      });
    }
  }

  const stationPkgs = stationAllocation.map((a) => String(a.packageIds || "").split("、").filter((id) => packageMap.has(id)));
  // 各岗位包含的"独占组"工作包（只属于单一组的工作包才参与吸收/分岗判定）
  const stationExclusive = stationPkgs.map((ids) => {
    const m = new Map();
    for (const pid of ids) {
      const gs = pkgGroups.get(pid);
      if (gs && gs.size === 1) {
        const gi = [...gs][0];
        m.set(gi, (m.get(gi) || 0) + 1);
      }
    }
    return m;
  });
  const groupStations = parsedGroups.map(() => new Set());
  stationExclusive.forEach((m, si) => {
    for (const gi of m.keys()) groupStations[gi].add(si);
  });
  if (!groupStations.some((s) => s.size)) return { stationAllocation, ttRows };

  // 各组最终工作包集合：独占工作包 + 被吸收岗位携带的未归属/跨组工作包
  const groupPkgs = parsedGroups.map(() => []);
  const groupRideAlong = parsedGroups.map(() => []);
  stationPkgs.forEach((ids, si) => {
    const ex = stationExclusive[si];
    if (ex.size === 1) {
      const gi = [...ex.keys()][0];
      for (const pid of ids) {
        const gs = pkgGroups.get(pid);
        if (gs && gs.size === 1 && [...gs][0] === gi) groupPkgs[gi].push(pid);
        else groupRideAlong[gi].push(pid);
      }
    } else {
      for (const pid of ids) {
        const gs = pkgGroups.get(pid);
        if (gs && gs.size === 1) groupPkgs[[...gs][0]].push(pid);
      }
    }
  });
  for (const list of [...groupPkgs, ...groupRideAlong]) {
    const seen = new Set();
    for (let i = 0; i < list.length; i++) {
      if (seen.has(list[i])) { list.splice(i, 1); i--; }
      seen.add(list[i]);
    }
  }

  const makeAlloc = (pkgIds, no, extra = {}) => {
    const seconds = pkgIds.reduce((a, pid) => a + ((packageMap.get(pid) && pkgWeight(packageMap.get(pid))) || 0), 0);
    const wires = pkgIds.reduce((a, pid) => {
      const p = packageMap.get(pid);
      return a + (p ? (p.wireCount || (p.wires || []).length) : 0);
    }, 0);
    const cfgSet = new Set();
    for (const pid of pkgIds) for (const c of (packageMap.get(pid) && packageMap.get(pid).configs) || []) cfgSet.add(c);
    return {
      stationNo: no,
      packageIds: pkgIds.join("、"),
      packageCount: pkgIds.length,
      wireCount: wires,
      totalSeconds: round(seconds, 2),
      tt,
      loadPercent: tt && tt > 0 ? round((seconds / tt) * 100, 1) : null,
      status: tt && tt > 0 && seconds > tt ? "【超TT，需拆分或增岗】" : "【候选方案，待工艺验证】",
      configs: [...cfgSet].join(" / "),
      ...extra
    };
  };

  const result = [];
  const inserted = new Set();
  const firstStationOf = (gi) => (groupStations[gi].size ? Math.min(...groupStations[gi]) : -1);
  const pushGroup = (gi) => {
    if (inserted.has(gi)) return;
    const g = parsedGroups[gi];
    const core = groupPkgs[gi];
    const ride = groupRideAlong[gi];
    const all = [...core, ...ride];
    const seconds = round(all.reduce((a, pid) => a + secOf(pid), 0), 2);
    const loopEst = groupLoopEstimate(g);
    const secondsWithLoop = round(seconds + loopEst, 2);
    const over = ttNum > 0 && secondsWithLoop > ttNum;
    const mode = overTtMode === "best-rate" ? "best-rate" : "force-same";
    const handleModeLabel = mode === "best-rate" ? "按最佳插接率拆分" : "强制同岗";
    const groupEntry = (pkgIds, no, extra = {}) => makeAlloc(pkgIds, no, {
      stationType: g.typeLabel.stationType,
      modeLabel: g.typeLabel.modeLabel,
      ...extra
    });
    const pushTtRow = (outcome, splitDetail, status) => {
      ttRows.push({
        idx: ttRows.length + 1,
        groupName: g.name,
        housings: [...g.housings].join("、"),
        modeLabel: g.typeLabel.modeLabel,
        mergedSeconds: secondsWithLoop,
        loopSeconds: loopEst,
        tt: ttNum || "",
        loadPercent: ttNum > 0 ? Math.round((secondsWithLoop / ttNum) * 100) : null,
        handleMode: handleModeLabel,
        outcome,
        splitDetail,
        status
      });
    };
    const fmtBins = (used, allCount, maxShow = 12) => {
      const parts = used.map((b, i) => `岗位${i + 1}:${round(b.sec, 2)}s(${Math.round((b.sec / ttNum) * 100)}%)`);
      if (parts.length > maxShow) return `共${used.length}个岗位，示例：` + parts.slice(0, maxShow).join("；") + `；…其余${used.length - maxShow}个岗位见候选岗位分配表`;
      return parts.join("；");
    };

    if (!over) {
      // 未超节拍：保留同岗
      result.push(groupEntry(all, result.length + 1, { groupName: g.name }));
      inserted.add(gi);
      pushTtRow("未超节拍（保留同岗）", loopEst ? `含打圈估算${loopEst}s` : "", "【候选方案，待工艺验证】");
      return;
    }

    if (mode === "force-same") {
      // 强制同岗：即使超节拍也保留在一个岗位
      result.push(groupEntry(all, result.length + 1, { groupName: g.name }));
      inserted.add(gi);
      pushTtRow(
        "超节拍-强制同岗",
        `合并后含打圈约${Math.round((secondsWithLoop / ttNum) * 100)}%负荷（包工时${seconds}s${loopEst ? `+打圈${loopEst}s` : ""}），超出TT ${ttNum}s，岗位已在岗位表标记【超TT，需拆分或增岗】，需人工调整`,
        "【超TT，需拆分或增岗】"
      );
      return;
    }

    // 按最佳插接率拆分：优先把用户填入的护套（core）尽量放在同一岗位
    const coreSec = core.reduce((a, pid) => a + secOf(pid), 0);
    const k = Math.max(2, Math.ceil(seconds / ttNum));
    const bins = Array.from({ length: k }, () => ({ pkgs: [], sec: 0 }));
    if (coreSec <= ttNum) {
      // 用户护套自身可同岗：主岗位放core，溢出工作按负载拆分
      bins[0].pkgs.push(...core);
      bins[0].sec = coreSec;
      const rest = ride.slice().sort((a, b) => secOf(b) - secOf(a));
      for (const pid of rest) {
        const b = bins.reduce((best, x) => (x.sec <= best.sec ? x : best), bins[0]);
        b.pkgs.push(pid);
        b.sec += secOf(pid);
      }
      const used = bins.filter((b) => b.pkgs.length);
      used.forEach((b, i) => {
        result.push(groupEntry(b.pkgs, result.length + 1, { groupName: i === 0 ? g.name : `${g.name}-拆分${i + 1}` }));
      });
      inserted.add(gi);
      const corePct = Math.round((coreSec / ttNum) * 100);
      pushTtRow(
        "超节拍-已拆分（用户护套保留同岗，溢出工作拆分）",
        `用户护套岗位约${corePct}%负荷；` + fmtBins(used, all.length) + (loopEst ? `；打圈按${loopEst}s估（各岗位最终以岗位表为准）` : ""),
        "【候选方案，待工艺验证】"
      );
    } else {
      // 用户护套自身工作量即超节拍：整体按负载拆分（护套会分散，见冲突清单）
      const sorted = all.slice().sort((a, b) => secOf(b) - secOf(a));
      for (const pid of sorted) {
        const b = bins.reduce((best, x) => (x.sec <= best.sec ? x : best), bins[0]);
        b.pkgs.push(pid);
        b.sec += secOf(pid);
      }
      const used = bins.filter((b) => b.pkgs.length);
      used.forEach((b, i) => {
        result.push(groupEntry(b.pkgs, result.length + 1, { groupName: i === 0 ? g.name : `${g.name}-拆分${i + 1}` }));
      });
      inserted.add(gi);
      pushTtRow(
        "超节拍-已拆分（用户护套自身即超节拍，无法全部同岗）",
        `对全部${all.length}个工作包按负载拆分；` + fmtBins(used, all.length) + (loopEst ? `；打圈按${loopEst}s估（各岗位最终以岗位表为准）` : ""),
        "【候选方案，待工艺验证】"
      );
    }
  };
  const pushStation = (si, pkgIds) => {
    if (!pkgIds.length) return;
    const orig = stationAllocation[si];
    result.push(makeAlloc(pkgIds, result.length + 1, {
      stationType: orig.stationType,
      modeLabel: orig.modeLabel
    }));
  };

  for (let si = 0; si < stationAllocation.length; si++) {
    const ex = stationExclusive[si];
    if (ex.size === 1) {
      const gi = [...ex.keys()][0];
      if (firstStationOf(gi) === si) pushGroup(gi);
      // 其余被吸收岗位跳过（内容已并入组岗位）
    } else if (ex.size > 1) {
      // 分岗点：先插入最早归属该位置的组岗位，再保留未归属/跨组工作包
      const gis = [...ex.keys()].sort((a, b) => firstStationOf(a) - firstStationOf(b));
      for (const gi of gis) {
        if (firstStationOf(gi) === si) pushGroup(gi);
      }
      const rest = stationPkgs[si].filter((pid) => {
        const gs = pkgGroups.get(pid);
        return !gs || gs.size > 1;
      });
      pushStation(si, rest);
    } else {
      pushStation(si, stationPkgs[si]);
    }
  }
  // 兜底：仍未插入的组（吸收岗位在分岗点之后被消耗等情况）
  parsedGroups.forEach((g, gi) => {
    if (!inserted.has(gi) && (groupPkgs[gi].length || groupRideAlong[gi].length)) pushGroup(gi);
  });
  return { stationAllocation: result, ttRows };
}

// 生成后的分组结果一致性校验（基于已生成的岗位明细）
function checkSameStationGroupResult(stationDetails, parsedGroups, issues) {
  for (const g of parsedGroups) {
    const housingStations = new Map(); // h -> Set<stationNo>
    for (const st of stationDetails || []) {
      for (const w of st.wireRows || []) {
        for (const h of [w.housing1, w.housing2]) {
          const hh = clean(h);
          if (!hh || !g.housings.has(hh)) continue;
          if (!housingStations.has(hh)) housingStations.set(hh, new Set());
          housingStations.get(hh).add(st.stationNo);
        }
      }
    }
    if (!housingStations.size) {
      issues.push({
        category: "同岗位护套分组",
        detail: `组【${g.name}】的护套${[...g.housings].join("、")}未在MBOM导线中识别到，未生成组岗位。`
      });
      continue;
    }
    const scattered = [...housingStations.entries()].filter(([, v]) => v.size > 1);
    if (scattered.length) {
      issues.push({
        category: "同岗位护套分组",
        detail: `组【${g.name}】要求同岗位的护套未完全在同一岗位：${scattered.map(([h, v]) => `${h}(${[...v].join("、")})`).join("；")}（可能因跨组工作包无法自动拆分）。`
      });
    }
  }
  // 跨组同岗检查：同一岗位同时含两个不同组的护套
  const stationGroups = new Map(); // stationNo -> Set<groupName>
  for (const st of stationDetails || []) {
    for (const w of st.wireRows || []) {
      for (const h of [w.housing1, w.housing2]) {
        const hh = clean(h);
        if (!hh) continue;
        for (const g of parsedGroups) {
          if (g.housings.has(hh)) {
            if (!stationGroups.has(st.stationNo)) stationGroups.set(st.stationNo, new Set());
            stationGroups.get(st.stationNo).add(g.name);
          }
        }
      }
    }
  }
  for (const [no, gs] of stationGroups) {
    if (gs.size > 1) {
      issues.push({
        category: "同岗位护套分组",
        detail: `岗位${no}同时包含组【${[...gs].join("、")}】的护套，跨组同岗，需人工确认。`
      });
    }
  }
}

function buildStationDetails(packages, stationAllocation, options, ebomMaterials, pdfKeywords, ultrasonicRules, forcedOfflineHousings = [], loopTimes = {}, configs = []) {
  const region = (options.regions && options.regions[0]) || "未指定部位";
  const onlineUltrasonic = options.onlineUltrasonic === true || options.onlineUltrasonic === "yes" || options.onlineUltrasonic === "是";
  const forcedSet = new Set((forcedOfflineHousings || []).map(clean).filter(Boolean));
  const packageMap = new Map(packages.map((p) => [p.id, p]));

  const allHousingPositions = new Map();
  const allHousingPositionsByConfig = new Map();
  for (const p of packages) {
    for (const w of p.wires) {
      if (w.housing1 && w.housing1 !== "-" && !isSpliceCode(w.housing1)) {
        if (!allHousingPositions.has(w.housing1)) allHousingPositions.set(w.housing1, new Set());
        if (w.position1 !== "") allHousingPositions.get(w.housing1).add(String(w.position1));
        for (const cfg of w.configCodes || []) {
          if (!allHousingPositionsByConfig.has(cfg)) allHousingPositionsByConfig.set(cfg, new Map());
          const cfgMap = allHousingPositionsByConfig.get(cfg);
          if (!cfgMap.has(w.housing1)) cfgMap.set(w.housing1, new Set());
          if (w.position1 !== "") cfgMap.get(w.housing1).add(String(w.position1));
        }
      }
      if (w.housing2 && w.housing2 !== "-" && !isSpliceCode(w.housing2)) {
        if (!allHousingPositions.has(w.housing2)) allHousingPositions.set(w.housing2, new Set());
        if (w.position2 !== "") allHousingPositions.get(w.housing2).add(String(w.position2));
        for (const cfg of w.configCodes || []) {
          if (!allHousingPositionsByConfig.has(cfg)) allHousingPositionsByConfig.set(cfg, new Map());
          const cfgMap = allHousingPositionsByConfig.get(cfg);
          if (!cfgMap.has(w.housing2)) cfgMap.set(w.housing2, new Set());
          if (w.position2 !== "") cfgMap.get(w.housing2).add(String(w.position2));
        }
      }
    }
  }

  const stations = (stationAllocation || []).map((all) => {
    const binPackages = String(all.packageIds || "").split("、").map((id) => packageMap.get(id)).filter(Boolean);
    const wires = [];
    for (const p of binPackages) {
      for (const w of p.wires) {
        wires.push({ pkgId: p.id, ...w });
      }
    }

    const stationHousingPositions = new Map();
    const stationHousingPositionsByConfig = new Map();
    for (const w of wires) {
      if (w.housing1 && w.housing1 !== "-" && !isSpliceCode(w.housing1)) {
        if (!stationHousingPositions.has(w.housing1)) stationHousingPositions.set(w.housing1, new Set());
        if (w.position1 !== "") stationHousingPositions.get(w.housing1).add(String(w.position1));
        for (const cfg of w.configCodes || []) {
          if (!stationHousingPositionsByConfig.has(cfg)) stationHousingPositionsByConfig.set(cfg, new Map());
          const cfgMap = stationHousingPositionsByConfig.get(cfg);
          if (!cfgMap.has(w.housing1)) cfgMap.set(w.housing1, new Set());
          if (w.position1 !== "") cfgMap.get(w.housing1).add(String(w.position1));
        }
      }
      if (w.housing2 && w.housing2 !== "-" && !isSpliceCode(w.housing2)) {
        if (!stationHousingPositions.has(w.housing2)) stationHousingPositions.set(w.housing2, new Set());
        if (w.position2 !== "") stationHousingPositions.get(w.housing2).add(String(w.position2));
        for (const cfg of w.configCodes || []) {
          if (!stationHousingPositionsByConfig.has(cfg)) stationHousingPositionsByConfig.set(cfg, new Map());
          const cfgMap = stationHousingPositionsByConfig.get(cfg);
          if (!cfgMap.has(w.housing2)) cfgMap.set(w.housing2, new Set());
          if (w.position2 !== "") cfgMap.get(w.housing2).add(String(w.position2));
        }
      }
    }

    const wireRows = wires.map((w) => {
      const h1 = w.housing1 && w.housing1 !== "-";
      const h2 = w.housing2 && w.housing2 !== "-";
      const realH1 = h1 && !isSpliceCode(w.housing1);
      const realH2 = h2 && !isSpliceCode(w.housing2);
      const h1FullConfigs = [];
      const h2FullConfigs = [];
      if (realH1) {
        for (const cfg of w.configCodes || []) {
          const allSet = allHousingPositionsByConfig.get(cfg)?.get(w.housing1);
          const stSet = stationHousingPositionsByConfig.get(cfg)?.get(w.housing1);
          if (allSet && stSet && stSet.size >= allSet.size) h1FullConfigs.push(cfg);
        }
      }
      if (realH2) {
        for (const cfg of w.configCodes || []) {
          const allSet = allHousingPositionsByConfig.get(cfg)?.get(w.housing2);
          const stSet = stationHousingPositionsByConfig.get(cfg)?.get(w.housing2);
          if (allSet && stSet && stSet.size >= allSet.size) h2FullConfigs.push(cfg);
        }
      }
      const fullRemark = [];
      if (h1FullConfigs.length) fullRemark.push(`护套1【${w.housing1}】本岗位满插（配置：${h1FullConfigs.join("、")}）`);
      if (h2FullConfigs.length) fullRemark.push(`护套2【${w.housing2}】本岗位满插（配置：${h2FullConfigs.join("、")}）`);
      if (isSpliceCode(w.housing1)) fullRemark.push(`SP端1【${w.housing1}】为焊点/超声波点，不按满插判定`);
      if (isSpliceCode(w.housing2)) fullRemark.push(`SP端2【${w.housing2}】为焊点/超声波点，不按满插判定`);

      const time = w._time || {};
      const spliceCodes = [w.housing1, w.housing2].filter((h) => isSpliceCode(h)).map(clean);
      let ultrasonic = "否";
      let ultrasonicRemark = "";
      if (onlineUltrasonic && isUltrasonicCandidate(w)) {
        const blockedUser = spliceCodes.filter((code) => ultrasonicRules && ultrasonicRules.userBlocked.includes(code));
        const blockedLimitCodes = spliceCodes.filter((code) =>
          ultrasonicRules &&
          (w.configCodes || []).some((cfg) => !((ultrasonicRules.allowedByConfig || {})[cfg] || []).includes(code))
        );
        const blockedTotal = spliceCodes.filter((code) => ultrasonicRules && ultrasonicRules.blockedByLimitTotal.includes(code));
        if (blockedUser.length) {
          ultrasonic = "否（需热缩管，不能在线）";
          ultrasonicRemark = `客户指定SP/SC压接点【${blockedUser.join("、")}】不能上在线超声波，需热缩管。`;
        } else if (blockedLimitCodes.length) {
          ultrasonic = "否（超单配置限组）";
          ultrasonicRemark = `单配置在线超声波最高${ultrasonicRules.maxGroupsPerConfig}组，SP/SC点【${blockedLimitCodes.join("、")}】在当前配置中超出限制。`;
        } else if (blockedTotal.length) {
          ultrasonic = "否（超在线超声波总组数限制）";
          ultrasonicRemark = `在线超声波最高总组数限制为${ultrasonicRules.maxTotalGroups}组，SP/SC点【${blockedTotal.join("、")}】超出总组数限制。`;
        } else {
          ultrasonic = "是（候选）";
          ultrasonicRemark = "满足在线超声波候选条件，待设备/工艺验证。";
        }
      } else if (onlineUltrasonic && !isUltrasonicCandidate(w)) {
        ultrasonicRemark = "该导线未识别到SP/SC焊点或双绞压接。";
      }

      return {
        pkgId: w.pkgId,
        w3: w.w3,
        w2: w.w2,
        w1: w.w1,
        drawingId: w.drawingId,
        material: w.material,
        color: w.color,
        spec: w.spec,
        length: w.length,
        terminal1: w.terminal1,
        seal1: w.seal1,
        housing1: w.housing1,
        position1: w.position1,
        terminal2: w.terminal2,
        seal2: w.seal2,
        housing2: w.housing2,
        position2: w.position2,
        configs: w.configCodes ? w.configCodes.join(" / ") : "",
        ultrasonic,
        ultrasonicRemark,
        plugEnd: getPlugEndText(w, forcedOfflineHousings),
        fullRemark: fullRemark.join("；") || "部分插接/待确认",
        insert1: realH1 && !forcedSet.has(clean(w.housing1)),
        insert2: realH2 && !forcedSet.has(clean(w.housing2)),
        takeTime: time.rt ? time.rt.takeTime : null,
        routeTime: time.rt ? time.rt.routeTime : null,
        plug1Time: time.plug1 ? time.plug1.total : null,
        plug2Time: time.plug2 ? time.plug2.total : null,
        totalTime: time.wireTime != null ? time.wireTime : null
      };
    });

    const relevantIds = new Set();
    for (const w of wires) {
      for (const v of [w.housing1, w.housing2, w.drawingId, w.terminal1, w.terminal2, w.seal1, w.seal2, w.customerNo, w.jettyNo]) {
        if (v && !isSpliceCode(v)) relevantIds.add(clean(v));
      }
    }
    const materialsFiltered = ebomMaterials.filter((m) => {
      if (m.drawingId && relevantIds.has(clean(m.drawingId))) return true;
      if (m.jettyNo && relevantIds.has(clean(m.jettyNo))) return true;
      if (m.spn && relevantIds.has(clean(m.spn))) return true;
      return false;
    });
    const seenMaterial = new Set();
    const materials = materialsFiltered.filter((m) => {
      const key = clean(m.drawingId) || clean(m.jettyNo) || clean(m.spn) || (clean(m.materialName) + "|" + clean(m.description));
      if (seenMaterial.has(key)) return false;
      seenMaterial.add(key);
      return true;
    });

    const tapeRemark =
      ebomMaterials.some((m) => m.materialName === "胶带") ||
      (pdfKeywords || []).some((k) => k.keyword === "胶带")
        ? "需按图纸核查胶带包胶"
        : "无胶带包胶/待确认";

    const stationName = all.groupName
      ? `${String(all.stationNo).padStart(2, "0")}-${region}-${clean(all.groupName)}-${modeShort(all.modeLabel || "岗位")}`
      : `${String(all.stationNo).padStart(2, "0")}-${region}-${modeShort(all.modeLabel || "岗位")}`;
    const stType = all.stationType || all.modeLabel || "";
    const idx = all.stationNo - 1;
    const next = stationAllocation[idx + 1];
    const isLastTransfer = stType.includes("KIT传递") && next && !String(next.stationType || "").includes("KIT传递");
    const isLastSub = stType.includes("SUB") && next && !String(next.stationType || "").includes("SUB");
    let loopSeconds = 0;
    if (stType.includes("KIT岗位") || stType.includes("纯KIT")) {
      loopSeconds = Number(loopTimes.singleKit || 0);
    } else if (stType.includes("KIT传递")) {
      loopSeconds = isLastTransfer ? Number(loopTimes.kitTransferLast || 0) : Number(loopTimes.kitTransferMiddle || 0);
    } else if (stType.includes("SUB")) {
      loopSeconds = isLastSub ? Number(loopTimes.subLast || 0) : 0;
    }
    all.totalSeconds = round((all.totalSeconds || 0) + loopSeconds);
    const configTime = {};
    for (const cfg of configs) {
      const cfgSum = binPackages.reduce((sum, p) => sum + (p.configTime && p.configTime[cfg] ? p.configTime[cfg] : 0), 0);
      configTime[cfg] = round(cfgSum + loopSeconds, 2);
    }

    return {
      stationNo: all.stationNo,
      stationName,
      region,
      regionRule: REGION_RULES[region] || "按V2.6.1预装工作包规则：以护套关联、待插端、保护件顺序、受控半成品和TT形成工作包。",
      modeLabel: all.modeLabel,
      totalSeconds: all.totalSeconds,
      loopTimeSeconds: loopSeconds,
      configTime,
      wireRows,
      materials,
      packageIds: all.packageIds,
      tapeRemark,
      status: all.status
    };
  });

  return stations;
}

function buildStationDotMatrix(stationDetails) {
  const rows = [];
  const issues = [];
  for (const st of stationDetails || []) {
    const byColor = new Map();
    for (const w of st.wireRows || []) {
      const color = clean(w.color);
      if (!color) continue;
      if (!byColor.has(color)) byColor.set(color, []);
      byColor.get(color).push(w);
    }
    for (const [color, group] of byColor.entries()) {
      if (group.length < 2) continue;
      const sorted = group.slice().sort((a, b) => (a.w1 || "").localeCompare(b.w1 || ""));
      if (sorted.length > DOT_CODE_POOL.length) {
        issues.push({
          stationNo: st.stationNo,
          color,
          count: sorted.length,
          detail: `岗位${st.stationName}颜色${color}同色线数量${sorted.length}超过可生成编码容量${DOT_CODE_POOL.length}，超出部分留空，需现场确认实际编码容量。`
        });
      }
      sorted.forEach((w, i) => {
        const code = i < DOT_CODE_POOL.length ? DOT_CODE_POOL[i] : "";
        rows.push({
          stationNo: st.stationNo,
          stationName: st.stationName,
          pkgId: w.pkgId || "",
          wireColor: color,
          wireId: w.w1,
          target: w.target || [w.housing1 && w.housing1 !== "-" ? `${w.housing1}/${w.position1}` : "", w.housing2 && w.housing2 !== "-" ? `${w.housing2}/${w.position2}` : ""].filter(Boolean).join("、"),
          config: w.configs,
          standardCode: code,
          checkResult: code ? "唯一" : "编码容量不足",
          status: code ? "候选编码，待现场容量与打点规则确认" : "【编码容量不足，需人工确认】"
        });
      });
    }
  }
  return { rows, issues };
}

function buildGrommetStations(grommetStations, region, configs, startNo) {
  return (grommetStations || []).map((g, i) => {
    const no = startNo + i;
    const name = `${String(no).padStart(2, "0")}-${region || "未指定"}-胶套-${clean(g.name)}`;
    const housings = String(g.housings || "").split(/[,，;；]+/).map(clean).filter(Boolean);
    const time = Number(g.time) || 0;
    const configTime = {};
    for (const cfg of configs) configTime[cfg] = time;
    const wireRows = housings.map((h) => ({
      w1: "",
      drawingId: "",
      material: "胶套/防水泥",
      color: "",
      housing1: h,
      position1: "",
      housing2: "",
      position2: "",
      configs: configs.join(" / "),
      insert1: false,
      insert2: false,
      fullRemark: "胶套后护套，强制线下插接完毕",
      plugEnd: "胶套后护套（线下）",
      totalTime: null,
      ultrasonic: "否"
    }));
    return {
      stationNo: no,
      stationName: name,
      region: region || "未指定",
      regionRule: "胶套专属岗位：安装防水泥和胶套；胶套后面的护套强制线下插接完毕。",
      modeLabel: "胶套专属岗位",
      stationType: "胶套岗位",
      totalSeconds: time,
      loopTimeSeconds: 0,
      configTime,
      wireRows,
      materials: [],
      packageIds: "",
      tapeRemark: "胶套/防水泥作业",
      status: "【候选方案，待工艺验证】"
    };
  });
}

async function analyzeProject(files, options = {}) {
  const inputs = parseInputs(files);
  const { standard, ebom, mbom } = inputs;
  const validationResult = validateUploadedFiles(files);
  const pdf = await parsePdfInput(files);
  const tt = options.tt != null ? Number(options.tt) : null;
  const maxStations = options.maxStations != null ? Number(options.maxStations) : null;
  const autoStations = options.autoStations === true || options.autoStations === "true" || options.autoStations === "yes";
  const maxSubFrames = options.maxSubFrames != null ? Number(options.maxSubFrames) : null;
  const loopTimes = options.loopTimes || {};
  const preassemblyMode = options.preassemblyMode || "";
  const onlineUltrasonic = options.onlineUltrasonic === true || options.onlineUltrasonic === "yes" || options.onlineUltrasonic === "是";
  const onlineUltrasonicMaxGroupsPerConfig = options.onlineUltrasonicMaxGroupsPerConfig != null ? Number(options.onlineUltrasonicMaxGroupsPerConfig) : null;
  const onlineUltrasonicMaxTotalGroups = options.onlineUltrasonicMaxTotalGroups != null ? Number(options.onlineUltrasonicMaxTotalGroups) : null;
  const noOnlineUltrasonicSplices = Array.isArray(options.noOnlineUltrasonicSplices) ? options.noOnlineUltrasonicSplices.map(clean).filter(Boolean) : [];
  const forcedOfflineHousings = Array.isArray(options.forcedOfflineHousings) ? options.forcedOfflineHousings.map(clean).filter(Boolean) : [];
  const sameStationHousings = Array.isArray(options.sameStationHousings) ? options.sameStationHousings.map(clean).filter(Boolean) : [];
  let sameStationGroups = Array.isArray(options.sameStationGroups) ? options.sameStationGroups : [];
  // 兼容旧版单个文本框：视为一组，默认纯KIT岗位
  if (!sameStationGroups.length && sameStationHousings.length) {
    sameStationGroups = [{ name: "同岗组01", housings: sameStationHousings.join(","), mode: "pure-kit" }];
  }
  const groupIssues = [];
  const parsedSameStationGroups = parseSameStationGroups(sameStationGroups, groupIssues);
  const sameStationOverTtMode = options.sameStationOverTtMode === "best-rate" ? "best-rate" : "force-same";
  const forcedOfflineSet = new Set(forcedOfflineHousings.map(clean).filter(Boolean));
  for (const g of parsedSameStationGroups) {
    for (const h of g.housings) {
      if (forcedOfflineSet.has(h)) {
        groupIssues.push({
          category: "同岗位护套分组",
          detail: `护套【${h}】在组【${g.name}】中，同时属于强制线下插接护套（胶套后护套或手动指定），组岗位类型与强制线下作业可能冲突，【待确认】。`
        });
      }
    }
  }
  const grommetStations = Array.isArray(options.grommetStations) ? options.grommetStations.filter((g) => g && g.name && String(g.housings || "").trim()) : [];
  for (const g of grommetStations) {
    const hs = String(g.housings || "").split(/[,，;；]+/).map(clean).filter(Boolean);
    for (const h of hs) {
      if (!forcedOfflineHousings.includes(h)) forcedOfflineHousings.push(h);
    }
  }
  const regions = Array.isArray(options.regions) ? options.regions.filter(Boolean) : [];
  const modeObj = MODE_OPTIONS.find((m) => m.value === preassemblyMode) || null;
  const modeLabel = modeObj ? modeObj.label : "未选择预装模式";

  const configs = mbom.configs.length ? mbom.configs : ["未识别配置"];
  const mergedResult = mergeWiresByW1(mbom.wires);
  const specialResult = mergeSpecialWires(mergedResult.wires, ebom.specialWires || []);
  const wires = specialResult.wires;
  const mergedMbom = {
    ...mbom,
    wires,
    rowCount: wires.length,
    uniqueW1: new Set(wires.map((w) => w.w1).filter(Boolean)).size
  };
  const fileAnalysis = buildFileAnalysis(mergedMbom);
  const ultrasonicRules = buildUltrasonicRules(wires, configs, noOnlineUltrasonicSplices, onlineUltrasonicMaxGroupsPerConfig, onlineUltrasonicMaxTotalGroups);
  const wireRows = buildWireRows(wires);
  const housingMatrix = buildHousingMatrix(wires);
  const packages = buildPackages(wires, configs, standard.entries);
  const positionRows = buildPositionMatrix(packages);
  const timeRows = buildTimeMatrix(packages, configs, tt);
  const ledger = buildLedger(files, mergedMbom, ebom, standard, pdf);
  const pdfKeywords = buildPDFKeywordTable(pdf);
  let stationCount = maxStations;
  let stationCountNote = "";
  if (tt && tt > 0) {
    // 岗位数以“单配置最高工时”为基准：先算出每个配置的总工时，取最高者 ÷ TT。
    // 单件产品只生产一个配置的导线，故不叠加各配置工时。
    const cfgTotals = configs.map((cfg) => round(sum(packages.map((p) => (p.configTime && p.configTime[cfg]) || 0)), 2));
    const maxCfg = Math.max(...cfgTotals.filter((t) => t > 0), 0);
    const needByTt = maxCfg > 0 ? Math.ceil(maxCfg / tt) : 0;
    if (needByTt > 0 && needByTt > (stationCount || 0)) {
      // 用户填的最多岗位数不足，按TT自动扩岗（最高配置也不超节拍）
      stationCount = needByTt;
      stationCountNote = `按TT自动扩岗：最高配置工时${maxCfg}s÷TT${tt}s需${needByTt}个岗位（原设置${maxStations || "未设置"}；各配置工时：${configs.map((c, i) => `${c}=${cfgTotals[i]}s`).join("；")}）`;
    }
  }
  if (!stationCount && autoStations && tt > 0) {
    // 兜底：未设置岗位数且TT有效时，按最高配置工时计算
    const cfgTotals = configs.map((cfg) => round(sum(packages.map((p) => (p.configTime && p.configTime[cfg]) || 0)), 2));
    const maxCfg = Math.max(...cfgTotals.filter((t) => t > 0), 0);
    stationCount = Math.max(1, Math.ceil(maxCfg / tt));
  }
  let stationAllocation = allocateStationsByMode(packages, stationCount, tt, modeLabel, preassemblyMode, maxSubFrames);
  const groupApplyResult = applySameStationGroups(stationAllocation, parsedSameStationGroups, packages, groupIssues, preassemblyMode, sameStationOverTtMode, tt, loopTimes);
  stationAllocation = groupApplyResult.stationAllocation;
  for (const r of groupApplyResult.ttRows) {
    if (String(r.outcome).startsWith("超节拍")) {
      groupIssues.push({
        category: "同岗位分组节拍处理",
        detail: `组【${r.groupName}】${r.handleMode}：合并工时${r.mergedSeconds}s vs TT ${r.tt}s（约${r.loadPercent}%），${r.outcome}${r.splitDetail ? "；" + r.splitDetail : ""}；详见“同岗位分组节拍处理”工作表。`
      });
    }
  }
  const stationDetails = buildStationDetails(
    packages,
    stationAllocation,
    { regions, onlineUltrasonic, preassemblyMode },
    ebom.materials,
    pdfKeywords,
    ultrasonicRules,
    forcedOfflineHousings,
    loopTimes,
    configs
  );
  checkSameStationGroupResult(stationDetails, parsedSameStationGroups, groupIssues);
  const grommetDetails = buildGrommetStations(grommetStations, regions[0] || "未指定", configs, stationDetails.length + 1);
  stationDetails.push(...grommetDetails);
  stationAllocation.push(...grommetDetails.map((g) => ({
    stationNo: g.stationNo,
    stationType: g.stationType,
    modeLabel: g.modeLabel,
    packageIds: "",
    packageCount: 0,
    wireCount: g.wireRows ? g.wireRows.length : 0,
    totalSeconds: g.totalSeconds,
    tt: tt || "",
    loadPercent: tt && tt > 0 ? round((g.totalSeconds / tt) * 100, 1) : null,
    status: g.status,
    configs: configs.join(" / ")
  })));
  const dot = buildStationDotMatrix(stationDetails);
  const pkgByWire = new Map();
  for (const p of packages) for (const w of p.wires) if (!pkgByWire.has(w)) pkgByWire.set(w, p.id);
  const issues = buildIssues(mergedMbom, ebom, standard, pdf, packages, dot.issues, positionRows);
  for (const c of mergedResult.conflicts || []) {
    issues.push({ category: "MBOM真冲突", detail: c.detail });
  }
  for (const w of mergedResult.warnings || []) {
    issues.push({ category: "MBOM合并复核", detail: w.detail });
  }
  for (const c of specialResult.conflicts || []) {
    issues.push({ category: "特殊线束真冲突", detail: c.detail });
  }
  for (const w of specialResult.warnings || []) {
    issues.push({ category: "特殊线束合并", detail: w.detail });
  }
  issues.push(...groupIssues);

  // 限制输入命中校验：未匹配到MBOM识别清单的强制提示（不静默忽略）
  const knownSpliceCodes = new Set((fileAnalysis.spliceList || []).map((s) => s.code));
  const knownHousingCodes = new Set((fileAnalysis.housingList || []).map((h) => h.code));
  const unmatchedNo = noOnlineUltrasonicSplices.filter((c) => !knownSpliceCodes.has(c));
  const unmatchedForced = forcedOfflineHousings.filter((c) => !knownHousingCodes.has(c));
  if (unmatchedNo.length) {
    issues.push({
      category: "限制输入校验",
      detail: `以下“不能上在线超声波”的SP/SC压接点不在MBOM识别清单中，输入未生效：${unmatchedNo.join("、")}`
    });
  }
  if (unmatchedForced.length) {
    issues.push({
      category: "限制输入校验",
      detail: `以下“强制线下插接”护套不在MBOM识别清单中，请确认拼写；若为胶套后护套（自动加入）可忽略：${unmatchedForced.join("、")}`
    });
  }
  if (regions.length) {
    issues.push({
      category: "部位筛选",
      detail: `已记录制作部位：${regions.join("、")}。请确认本次上传的MBOM/工艺图纸与该部位一致（一份文件对应一个部位）；当前部位用于工艺标识与岗位命名，不执行按部位数据过滤，也不改变分工艺算法。`
    });
  }
  if (stationCountNote) {
    issues.push({
      category: "岗位数提示",
      detail: stationCountNote + "；岗位数已按节拍自动调整，最终以TT验证为准。"
    });
  }
  const overTtAlloc = stationAllocation.filter((a) => a.loadPercent != null && a.loadPercent > 100);
  if (overTtAlloc.length && tt && tt > 0) {
    issues.push({
      category: "TT验证",
      detail: `仍有${overTtAlloc.length}个岗位估算负荷超过TT ${tt}s（多因单个工作包工时即超TT、或包粒度无法精确装入，少量为组合略超）；需人工拆分工作包或进一步调整岗位，详见“候选岗位分配”表。`
    });
  }
  if (stationCount && stationAllocation.length > stationCount) {
    issues.push({
      category: "岗位数提示",
      detail: `候选岗位数${stationAllocation.length}超过设置的最多预装岗位数${stationCount}：可能因“每个SUB组最多${maxSubFrames || "未设"}个预装”约束或同岗位分组按节拍拆分导致，属候选结果，最终岗位数量须以正式工时、TT和布局验证为准。`
    });
  }
  const plan = {
    regions,
    maxStations: stationCount ?? null,
    maxStationsRequested: maxStations ?? null,
    stationCountNote,
    autoStations,
    maxSubFrames,
    loopTimes,
    preassemblyMode,
    modeLabel,
    onlineUltrasonic,
    onlineUltrasonicMaxGroupsPerConfig,
    onlineUltrasonicMaxTotalGroups,
    noOnlineUltrasonicSplices,
    forcedOfflineHousings,
    sameStationHousings,
    sameStationGroups,
    sameStationOverTtMode,
    groupTtRows: groupApplyResult.ttRows,
    ultrasonicRules,
    stationAllocation,
    stationDetails,
    note: "部位、最大岗位数、预装模式、在线超声波、单配置在线超声波限组、在线超声波总组数限制、不可在线SP/SC点、强制线下插接护套、同岗位护套分组用于生成工艺策划参数和候选岗位分配；实际岗位数量仍须由正式工时、TT、布局和工作包验证。"
  };

  return {
    generatedAt: new Date().toISOString(),
    appVersion: "1.0.0",
    knowledgeVersion: "四轮汽车线束后段预装工艺规则 V2.6.1 摘要",
    options: {
      tt: tt ?? null,
      maxStations: stationCount ?? null,
      autoStations,
      maxSubFrames,
      loopTimes,
      preassemblyMode,
      modeLabel,
      onlineUltrasonic,
      onlineUltrasonicMaxGroupsPerConfig,
      onlineUltrasonicMaxTotalGroups,
      noOnlineUltrasonicSplices,
      forcedOfflineHousings,
      sameStationHousings,
      sameStationGroups,
      sameStationOverTtMode,
      regions
    },
    plan,
    fileAnalysis,
    validation: validationResult.validation,
    canGenerate: validationResult.canGenerate,
    summary: {
      files: ledger.length,
      wires: wires.length,
      uniqueW1: mbom.uniqueW1 || wires.length,
      materials: ebom.materials.length,
      standardEntries: standard.entries.length,
      packages: packages.length,
      stations: stationAllocation.length,
      configs: configs,
      housingPairs: housingMatrix.length,
      dotRows: dot.rows.length,
      issues: issues.length
    },
    ledger,
    configs,
    wires: wireRows,
    ebomMaterials: ebom.materials.map((m, i) => ({ idx: i + 1, ...m })),
    standardHours: standard.entries.map((e, i) => ({ idx: i + 1, ...e })),
    housingMatrix,
    pathRows: wires.map((w) => ({
      w3: w.w3,
      w2: w.w2,
      w1: w.w1,
      drawingId: w.drawingId,
      material: w.material,
      color: w.color,
      length: w.length,
      housing1: w.housing1,
      position1: w.position1,
      terminal1: w.terminal1,
      seal1: w.seal1,
      housing2: w.housing2,
      position2: w.position2,
      terminal2: w.terminal2,
      seal2: w.seal2,
      configs: w.configCodes.join(" / "),
      pkgId: pkgByWire.get(w) || "",
      status: "待唯一责任确认"
    })),
    packages: packages.map((p) => ({
      id: p.id,
      kind: p.kind,
      name: p.name,
      key: p.key,
      wireCount: p.wireCount,
      totalLength: p.totalLength,
      maxLength: p.maxLength,
      housings: p.housings,
      housingCount: p.housingCount,
      anchor: p.anchor,
      configs: p.configs.join(" / "),
      routeType: p.routeType,
      status: p.status,
      estimatedSeconds: p.estimatedSeconds,
      connectorPlacementSeconds: p.connectorPlacementSeconds,
      configTime: p.configTime,
      wires: p.wires.map((w) => w.w1 || w.drawingId).join("、")
    })),
    positionRows,
    dotRows: dot.rows,
    dotIssues: dot.issues,
    timeRows,
    issues,
    pdf,
    pdfKeywords,
    specialSheets: ebom.specialSheets
  };
}

function aoaToSheet(aoa) {
  return XLSX.utils.aoa_to_sheet(aoa);
}

function addSheet(wb, name, aoa) {
  const sheet = aoaToSheet(aoa);
  XLSX.utils.book_append_sheet(wb, sheet, name);
}

function grayCell(sheet, r, c) {
  const addr = XLSX.utils.encode_cell({ r, c });
  if (!sheet[addr]) sheet[addr] = { v: "" };
  sheet[addr].s = { fill: { fgColor: { rgb: "BFBFBF" } } };
}

// 同岗位分组节拍处理工作表（页面Tab与导出工作表共用此结构）
function buildGroupTtSheetAoa(plan) {
  const rows = (plan && plan.groupTtRows) || [];
  return [
    ["序号", "组名", "护套", "岗位类型", "合并工时(秒)", "TT(秒)", "负荷率%", "处理方式", "结果", "拆分情况/说明", "状态"],
    ...rows.map((r) => [r.idx, r.groupName, r.housings, r.modeLabel, r.mergedSeconds, r.tt, r.loadPercent, r.handleMode, r.outcome, r.splitDetail, r.status])
  ];
}

function overTtModeLabel(plan) {
  return plan && plan.sameStationOverTtMode === "best-rate" ? "按最佳插接率拆分" : "强制同岗（默认）";
}

// ==================== 整体预装插接率（按配置） ====================
// 定义（用户确认口径）：
//   分母 = 当前配置所有护套需要插接的端数（全束，含强制线下/胶套后护套，排除SP/SC与空端）
//   分子 = 整体预装（所有预装岗位合计）需要插接的端数
//   插接率 = 分子 ÷ 分母（按配置分别统计，不按单个工位输出）
function wireConfigCodes(w) {
  return String(w.configs || "").split("/").map((s) => s.trim()).filter(Boolean);
}
function realPlugEndsOfWire(w) {
  let n = 0;
  if (w.housing1 && w.housing1 !== "-" && !isSpliceCode(w.housing1) && w.insert1) n += 1;
  if (w.housing2 && w.housing2 !== "-" && !isSpliceCode(w.housing2) && w.insert2) n += 1;
  return n;
}
function stationPlugEndsText(st) {
  const m = new Map();
  for (const w of st.wireRows || []) {
    const n = realPlugEndsOfWire(w);
    if (!n) continue;
    for (const cfg of wireConfigCodes(w)) m.set(cfg, (m.get(cfg) || 0) + n);
  }
  return [...m.entries()].map(([c, v]) => `${c}:${v}端`).join("；") || "0端";
}
function buildPreassemblyPlugRate(result, stations) {
  // 分母：全束需插接端数（按配置）
  const totalByCfg = new Map();
  for (const w of result.wires || []) {
    let n = 0;
    if (w.housing1 && w.housing1 !== "-" && !isSpliceCode(w.housing1)) n += 1;
    if (w.housing2 && w.housing2 !== "-" && !isSpliceCode(w.housing2)) n += 1;
    if (!n) continue;
    for (const cfg of wireConfigCodes(w)) totalByCfg.set(cfg, (totalByCfg.get(cfg) || 0) + n);
  }
  // 分子：整体预装插接端数（按配置）
  const plugByCfg = new Map();
  for (const st of stations || []) {
    for (const w of st.wireRows || []) {
      const n = realPlugEndsOfWire(w);
      if (!n) continue;
      for (const cfg of wireConfigCodes(w)) plugByCfg.set(cfg, (plugByCfg.get(cfg) || 0) + n);
    }
  }
  const cfgs = [...new Set([...totalByCfg.keys(), ...plugByCfg.keys()])];
  return cfgs.map((cfg) => {
    const t = totalByCfg.get(cfg) || 0;
    const p = plugByCfg.get(cfg) || 0;
    return [cfg, p, t, t ? round((p / t) * 100, 1) + "%" : "无数据"];
  });
}

function buildProcessWorkbook(result) {
  const wb = XLSX.utils.book_new();
  const plan = result.plan || {};
  const stations = plan.stationDetails || [];

  addSheet(wb, "汇总", [
    ["项目", "值"],
    ["生成时间", result.generatedAt],
    ["制作部位", (plan.regions || []).join("、")],
    ["最多预装岗位数", plan.maxStations || ""],
    ["预装模式", plan.modeLabel || ""],
    ["是否有在线超声波", plan.onlineUltrasonic ? "是" : "否"],
    ["单配置在线超声波最高组数", plan.onlineUltrasonicMaxGroupsPerConfig || ""],
    ["在线超声波最高总组数限制", plan.onlineUltrasonicMaxTotalGroups || ""],
    ["不能在线SP/SC压接点", (plan.noOnlineUltrasonicSplices || []).join("、")],
    ["强制线下插接护套", (plan.forcedOfflineHousings || []).join("、")],
    ["组超节拍处理方式", overTtModeLabel(plan)],
    ["TT(秒)", result.options && result.options.tt != null ? result.options.tt : ""],
    ["岗位数", stations.length],
    ["说明", plan.note || ""]
  ]);

  addSheet(wb, "模板校验结果", [
    ["文件", "状态", "问题级别", "问题说明"],
    ...Object.entries(result.validation || {}).flatMap(([key, v]) => {
      const labels = { standard: "标准工时文件", ebom: "EBOM文件", mbom: "MBOM/Cutting文件" };
      const rows = (v.issues || []).map((it) => [labels[key] || key, v.ok ? "通过" : "问题", it.level, it.message]);
      if (!rows.length) rows.push([labels[key] || key, "通过", "", ""]);
      return rows;
    })
  ]);

  addSheet(wb, "在线超声波限制汇总", [
    ["参数", "值"],
    ["是否有在线超声波", plan.onlineUltrasonic ? "是" : "否"],
    ["单配置在线超声波最高组数", plan.onlineUltrasonicMaxGroupsPerConfig || ""],
    ["在线超声波最高总组数限制", plan.onlineUltrasonicMaxTotalGroups || ""],
    ["不能在线SP/SC压接点", (plan.noOnlineUltrasonicSplices || []).join("、")],
    ["各配置允许在线超声波组数", Object.entries((plan.ultrasonicRules && plan.ultrasonicRules.groupsPerConfig) || {}).map(([c, n]) => `${c}:${n}`).join("、")],
    ["全局允许在线超声波组数", plan.ultrasonicRules ? plan.ultrasonicRules.totalAllowedGroups : ""],
    ["超出单配置限组SP/SC点", (plan.ultrasonicRules && plan.ultrasonicRules.blockedByLimit || []).join("、")],
    ["超出总组数限制SP/SC点", (plan.ultrasonicRules && plan.ultrasonicRules.blockedByLimitTotal || []).join("、")]
  ]);

  addSheet(wb, "线下插接率与工时", [
    ["整体预装插接率（按配置）：分子=整体预装插接端数，分母=全束需插接端数（含强制线下/胶套后护套，排除SP/SC与空端）"],
    ["配置", "整体预装插接端数", "全束需插接端数", "整体预装插接率"],
    ...buildPreassemblyPlugRate(result, stations),
    [],
    ["岗位号", "岗位名称", "本岗位插接端数(按配置)", "满插护套", "是否需要包扎", "各配置工时(秒)"],
    ...stations.map((st) => {
      const fullSet = new Set();
      for (const w of st.wireRows || []) {
        for (const part of String(w.fullRemark || "").split("；")) {
          if (part.includes("满插")) fullSet.add(part);
        }
      }
      const configTimeText = Object.entries(st.configTime || {}).map(([c, s]) => `${c}:${s}s`).join("；");
      return [
        st.stationNo,
        st.stationName,
        stationPlugEndsText(st),
        [...fullSet].join("；") || "无",
        st.tapeRemark && st.tapeRemark.includes("需按图纸核查") ? "是" : "否/待确认",
        configTimeText
      ];
    })
  ]);

  addSheet(wb, "岗位物料与看板汇总", [
    ["岗位号", "岗位名称", "物料", "看板", "护套满插", "看板是否超声波"],
    ...stations.map((st) => {
      const materials = uniq((st.materials || []).map((m) => m.materialName)).join("、") || "无";
      const kanbans = uniq((st.wireRows || []).map((w) => [w.w3, w.w2, w.w1].filter(Boolean).join("/")).filter(Boolean)).join("；") || "无";
      const full = uniq((st.wireRows || []).flatMap((w) => String(w.fullRemark || "").split("；").filter((p) => p.includes("满插")))).join("；") || "无";
      const ultrasonic = uniq((st.wireRows || []).filter((w) => String(w.ultrasonic || "").includes("是")).map((w) => w.w1).filter(Boolean)).join("、") || "无";
      return [st.stationNo, st.stationName, materials, kanbans, full, ultrasonic];
    })
  ]);

  addSheet(wb, "同岗位分组节拍处理", buildGroupTtSheetAoa(plan));

  for (const st of stations) {
    const aoa = [];
    aoa.push(["岗位号", st.stationNo]);
    aoa.push(["岗位名称", st.stationName]);
    aoa.push(["制作部位", st.region]);
    aoa.push(["V2.6.1部位工艺要点", st.regionRule || ""]);
    aoa.push(["预装模式", st.modeLabel]);
    aoa.push(["是否有在线超声波", plan.onlineUltrasonic ? "是" : "否"]);
    aoa.push(["单配置在线超声波最高组数", plan.onlineUltrasonicMaxGroupsPerConfig || ""]);
    aoa.push(["在线超声波最高总组数限制", plan.onlineUltrasonicMaxTotalGroups || ""]);
    aoa.push(["不能在线SP/SC压接点", (plan.noOnlineUltrasonicSplices || []).join("、")]);
    aoa.push(["强制线下插接护套", (plan.forcedOfflineHousings || []).join("、")]);
    aoa.push(["TT(秒)", result.options && result.options.tt != null ? result.options.tt : ""]);
    for (const [cfg, sec] of Object.entries(st.configTime || {})) {
      aoa.push([`岗位估算工时-配置 ${cfg}(秒)`, sec]);
    }
    aoa.push(["状态", st.status]);
    aoa.push([]);
    aoa.push(["本岗位相关EBOM物料"]);
    aoa.push(["材料名称", "Description", "图纸号", "捷翼号", "厂家号", "厂家", "图纸用量", "工艺余量", "总用量", "单位", "备注"]);
    if (st.materials && st.materials.length) {
      for (const m of st.materials) {
        aoa.push([m.materialName, m.description, m.drawingId, m.jettyNo, m.spn, m.supplier, m.designQty, m.processAllowance, m.totalQty, m.unit, m.notes]);
      }
    } else {
      aoa.push(["未匹配到明确物料", "", "", "", "", "", "", "", "", "", "需人工按图纸补充"]);
    }
    aoa.push([]);
    aoa.push(["本岗位导线/看板明细"]);
    aoa.push([
      "工作包", "W3看板", "W2看板", "W1看板", "导线图号", "材料名称", "颜色", "规格", "长度",
      "端子1", "雨塞1", "护套1", "孔位1", "端子2", "雨塞2", "护套2", "孔位2",
      "配置", "是否在线超声波", "在线超声波备注", "本岗位需插接端", "满插备注", "取线工时", "布线工时", "插接1工时", "插接2工时", "导线小计工时"
    ]);
    for (const w of st.wireRows || []) {
      aoa.push([
        w.pkgId, w.w3, w.w2, w.w1, w.drawingId, w.material, w.color, w.spec, w.length,
        w.terminal1, w.seal1, w.housing1, w.position1, w.terminal2, w.seal2, w.housing2, w.position2,
        w.configs, w.ultrasonic, w.ultrasonicRemark, w.plugEnd, w.fullRemark,
        w.takeTime, w.routeTime, w.plug1Time, w.plug2Time, w.totalTime
      ]);
    }
    aoa.push([]);
    aoa.push(["胶带包胶备注", st.tapeRemark]);
    aoa.push(["人工核查说明", "满插、胶带、超声波、工时均为候选/估算；正式工艺需按图纸、现场和TT复核。"]);
    const sheetName = st.stationName.slice(0, 31);
    addSheet(wb, sheetName, aoa);
    const sheet = wb.Sheets[sheetName];
    const detailHeaderIdx = aoa.findIndex((r) => r[0] === "工作包");
    if (detailHeaderIdx >= 0) {
      (st.wireRows || []).forEach((w, k) => {
        const r = detailHeaderIdx + 1 + k;
        if (!w.insert1) {
          grayCell(sheet, r, 11);
          grayCell(sheet, r, 12);
        }
        if (!w.insert2) {
          grayCell(sheet, r, 15);
          grayCell(sheet, r, 16);
        }
      });
    }
  }

  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
}

function buildReviewWorkbook(result) {
  const wb = XLSX.utils.book_new();
  const plan = result.plan || {};
  const stations = plan.stationDetails || [];

  addSheet(wb, "核查说明", [
    ["人工核查表说明"],
    [""],
    ["每个岗位一个工作表，请重点核对："],
    ["1. 本岗位物料是否齐全；"],
    ["2. 看板信息是否完整；"],
    ["3. 是否在线超声波线标识是否正确；"],
    ["4. 本岗位需要插接哪一头；"],
    ["5. 护套是否满插；"],
    ["6. 是否需要胶带包胶；"],
    ["7. 工时是否与正式标准工时一致。"]
  ]);

  addSheet(wb, "模板校验结果", [
    ["文件", "状态", "问题级别", "问题说明"],
    ...Object.entries(result.validation || {}).flatMap(([key, v]) => {
      const labels = { standard: "标准工时文件", ebom: "EBOM文件", mbom: "MBOM/Cutting文件" };
      const rows = (v.issues || []).map((it) => [labels[key] || key, v.ok ? "通过" : "问题", it.level, it.message]);
      if (!rows.length) rows.push([labels[key] || key, "通过", "", ""]);
      return rows;
    })
  ]);

  addSheet(wb, "在线超声波限制汇总", [
    ["参数", "值"],
    ["是否有在线超声波", plan.onlineUltrasonic ? "是" : "否"],
    ["单配置在线超声波最高组数", plan.onlineUltrasonicMaxGroupsPerConfig || ""],
    ["在线超声波最高总组数限制", plan.onlineUltrasonicMaxTotalGroups || ""],
    ["不能在线SP/SC压接点", (plan.noOnlineUltrasonicSplices || []).join("、")],
    ["各配置允许在线超声波组数", Object.entries((plan.ultrasonicRules && plan.ultrasonicRules.groupsPerConfig) || {}).map(([c, n]) => `${c}:${n}`).join("、")],
    ["全局允许在线超声波组数", plan.ultrasonicRules ? plan.ultrasonicRules.totalAllowedGroups : ""],
    ["超出单配置限组SP/SC点", (plan.ultrasonicRules && plan.ultrasonicRules.blockedByLimit || []).join("、")],
    ["超出总组数限制SP/SC点", (plan.ultrasonicRules && plan.ultrasonicRules.blockedByLimitTotal || []).join("、")]
  ]);

  addSheet(wb, "线下插接率与工时", [
    ["整体预装插接率（按配置）：分子=整体预装插接端数，分母=全束需插接端数（含强制线下/胶套后护套，排除SP/SC与空端）"],
    ["配置", "整体预装插接端数", "全束需插接端数", "整体预装插接率"],
    ...buildPreassemblyPlugRate(result, stations),
    [],
    ["岗位号", "岗位名称", "本岗位插接端数(按配置)", "满插护套", "是否需要包扎", "各配置工时(秒)"],
    ...stations.map((st) => {
      const fullSet = new Set();
      for (const w of st.wireRows || []) {
        for (const part of String(w.fullRemark || "").split("；")) {
          if (part.includes("满插")) fullSet.add(part);
        }
      }
      const configTimeText = Object.entries(st.configTime || {}).map(([c, s]) => `${c}:${s}s`).join("；");
      return [
        st.stationNo,
        st.stationName,
        stationPlugEndsText(st),
        [...fullSet].join("；") || "无",
        st.tapeRemark && st.tapeRemark.includes("需按图纸核查") ? "是" : "否/待确认",
        configTimeText
      ];
    })
  ]);

  addSheet(wb, "岗位物料与看板汇总", [
    ["岗位号", "岗位名称", "物料", "看板", "护套满插", "看板是否超声波"],
    ...stations.map((st) => {
      const materials = uniq((st.materials || []).map((m) => m.materialName)).join("、") || "无";
      const kanbans = uniq((st.wireRows || []).map((w) => [w.w3, w.w2, w.w1].filter(Boolean).join("/")).filter(Boolean)).join("；") || "无";
      const full = uniq((st.wireRows || []).flatMap((w) => String(w.fullRemark || "").split("；").filter((p) => p.includes("满插")))).join("；") || "无";
      const ultrasonic = uniq((st.wireRows || []).filter((w) => String(w.ultrasonic || "").includes("是")).map((w) => w.w1).filter(Boolean)).join("、") || "无";
      return [st.stationNo, st.stationName, materials, kanbans, full, ultrasonic];
    })
  ]);

  addSheet(wb, "同岗位分组节拍处理", buildGroupTtSheetAoa(plan));

  for (const st of stations) {
    const aoa = [];
    aoa.push(["岗位号", st.stationNo]);
    aoa.push(["岗位名称", st.stationName]);
    aoa.push(["制作部位", st.region]);
    aoa.push(["V2.6.1部位工艺要点", st.regionRule || ""]);
    aoa.push(["预装模式", st.modeLabel]);
    aoa.push(["是否有在线超声波", plan.onlineUltrasonic ? "是" : "否"]);
    aoa.push(["单配置在线超声波最高组数", plan.onlineUltrasonicMaxGroupsPerConfig || ""]);
    aoa.push(["在线超声波最高总组数限制", plan.onlineUltrasonicMaxTotalGroups || ""]);
    aoa.push(["不能在线SP/SC压接点", (plan.noOnlineUltrasonicSplices || []).join("、")]);
    aoa.push(["强制线下插接护套", (plan.forcedOfflineHousings || []).join("、")]);
    for (const [cfg, sec] of Object.entries(st.configTime || {})) {
      aoa.push([`岗位估算工时-配置 ${cfg}(秒)`, sec]);
    }
    aoa.push(["胶带包胶备注", st.tapeRemark]);
    aoa.push([]);
    aoa.push(["物料信息"]);
    aoa.push(["材料名称", "图纸号", "捷翼号", "厂家号", "厂家", "图纸用量", "单位", "备注"]);
    if (st.materials && st.materials.length) {
      for (const m of st.materials) {
        aoa.push([m.materialName, m.drawingId, m.jettyNo, m.spn, m.supplier, m.designQty, m.unit, m.notes]);
      }
    }
    aoa.push([]);
    aoa.push(["看板/导线信息与核查"]);
    aoa.push([
      "W3看板", "W2看板", "W1看板", "导线图号", "材料/颜色/规格", "长度",
      "是否在线超声波", "在线超声波备注", "本岗位需插接端", "满插备注", "工时小计", "配置"
    ]);
    for (const w of st.wireRows || []) {
      aoa.push([
        w.w3, w.w2, w.w1, w.drawingId,
        `${w.material} / ${w.color} / ${w.spec}`,
        w.length,
        w.ultrasonic,
        w.ultrasonicRemark,
        w.plugEnd,
        w.fullRemark,
        w.totalTime,
        w.configs
      ]);
    }
    addSheet(wb, st.stationName.slice(0, 31), aoa);
  }

  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
}

function buildWorkbook(result) {
  const wb = XLSX.utils.book_new();
  const plan = result.plan || {};

  addSheet(wb, "项目输入资料台账", [
    ["文件类型", "文件名称", "大小KB", "读取状态", "简要信息"],
    ...result.ledger.map((r) => [r.fileType, r.fileName, r.sizeKB, r.status, r.info])
  ]);

  addSheet(wb, "总体摘要", [
    ["指标", "值"],
    ["生成时间", result.generatedAt],
    ["软件版本", result.appVersion],
    ["知识规则", result.knowledgeVersion],
    ["制作部位", (result.plan && result.plan.regions && result.plan.regions.length) ? result.plan.regions.join("、") : "未选择"],
    ["预装模式", (result.plan && result.plan.modeLabel) || "未选择"],
    ["最大预装岗位数", (result.plan && result.plan.maxStations) || "未设置"],
    ["MBOM导线数", result.summary.wires + " 行（唯一W1 " + (result.summary.uniqueW1 || result.summary.wires) + "）"],
    ["EBOM物料数", result.summary.materials],
    ["标准工时要素数", result.summary.standardEntries],
    ["候选工作包数", result.summary.packages],
    ["候选岗位分配", result.summary.stations || 0],
    ["护套关联对", result.summary.housingPairs],
    ["配置", result.configs.join("、")],
    ["待确认/冲突项", result.summary.issues]
  ]);

  addSheet(wb, "工艺策划参数", [
    ["参数", "值"],
    ["制作部位", (result.plan && result.plan.regions) ? result.plan.regions.join("、") : ""],
    ["最多预装岗位数", (result.plan && result.plan.maxStations) ?? ""],
    ["预装模式代码", (result.plan && result.plan.preassemblyMode) || ""],
    ["预装模式说明", (result.plan && result.plan.modeLabel) || ""],
    ["是否有在线超声波", (result.plan && result.plan.onlineUltrasonic) ? "是" : "否"],
    ["单配置在线超声波最高组数", (result.plan && result.plan.onlineUltrasonicMaxGroupsPerConfig) ?? ""],
    ["在线超声波最高总组数限制", (result.plan && result.plan.onlineUltrasonicMaxTotalGroups) ?? ""],
    ["不能在线SP/SC压接点", (result.plan && result.plan.noOnlineUltrasonicSplices || []).join("、")],
    ["强制线下插接护套", (result.plan && result.plan.forcedOfflineHousings || []).join("、")],
    ["组超节拍处理方式", overTtModeLabel(result.plan)],
    ["TT(秒)", (result.options && result.options.tt) ?? ""],
    ["说明", (result.plan && result.plan.note) || ""]
  ]);

  addSheet(wb, "同岗位分组节拍处理", buildGroupTtSheetAoa(result.plan));

  addSheet(wb, "文件分析结果", [
    ["指标", "值"],
    ["MBOM导线数", result.fileAnalysis ? result.fileAnalysis.totalWires : result.summary.wires],
    ["唯一W1", result.fileAnalysis ? result.fileAnalysis.uniqueW1 : ""],
    ["护套数量", result.fileAnalysis ? result.fileAnalysis.housingCount : ""],
    ["SP/SC压接点数量", result.fileAnalysis ? result.fileAnalysis.spliceCount : ""],
    ["单配置在线超声波组数", result.fileAnalysis ? Object.entries(result.fileAnalysis.onlineUltrasonicGroupsPerConfig || {}).map(([c, n]) => `${c}:${n}`).join("、") : ""],
    [],
    ["护套清单", "端数", "孔位数"],
    ...((result.fileAnalysis && result.fileAnalysis.housingList) || []).map((h) => [h.code, h.endCount, h.positionCount]),
    [],
    ["SP/SC压接点清单", "端数", "配置"],
    ...((result.fileAnalysis && result.fileAnalysis.spliceList) || []).map((s) => [s.code, s.endCount, s.configs])
  ]);

  addSheet(wb, "候选岗位分配", [
    ["岗位号", "预装模式", "包含工作包", "工作包数", "导线数", "估算工时(秒)", "TT(秒)", "负荷率%", "配置", "状态"],
    ...((result.plan && result.plan.stationAllocation) || []).map((r) => [r.stationNo, r.modeLabel, r.packageIds, r.packageCount, r.wireCount, r.totalSeconds, r.tt, r.loadPercent, r.configs, r.status])
  ]);

  addSheet(wb, "MBOM导线表", [
    ["序号", "W3看板号", "W2看板号", "W1看板号", "材料名称", "图纸号", "客户号", "规格", "颜色", "下料长度", "单位", "端子1", "雨塞1", "端子2", "雨塞2", "护套1", "孔位1", "护套2", "孔位2", "配置", "状态"],
    ...result.wires.map((r) => [r.idx, r.w3, r.w2, r.w1, r.material, r.drawingId, r.customerNo, r.spec, r.color, r.length, r.unit, r.terminal1, r.seal1, r.terminal2, r.seal2, r.housing1, r.position1, r.housing2, r.position2, r.configs, r.status])
  ]);

  addSheet(wb, "EBOM物料表", [
    ["序号", "模块号", "模块名称", "材料名称", "Description", "图纸号", "捷翼号", "厂家号", "厂家", "图纸用量", "工艺余量", "总用量", "单位", "单价", "价格汇总", "理论铜重", "备注"],
    ...result.ebomMaterials.map((r) => [r.idx, r.moduleNo, r.moduleName, r.materialName, r.description, r.drawingId, r.jettyNo, r.spn, r.supplier, r.designQty, r.processAllowance, r.totalQty, r.unit, r.unitPrice, r.totalPrice, r.copperWeight, r.notes])
  ]);

  addSheet(wb, "标准工时表", [
    ["序号", "工序", "工作要素", "描述", "动作开始", "动作结束", "标准工时", "单位"],
    ...result.standardHours.map((r) => [r.idx, r.process, r.activity, r.comments, r.clockStart, r.clockStop, r.time, r.unit])
  ]);

  addSheet(wb, "导线完整路径表", [
    ["W3", "W2", "W1", "图纸号", "材料", "颜色", "长度", "护套1", "孔位1", "端子1", "雨塞1", "护套2", "孔位2", "端子2", "雨塞2", "配置", "候选包", "状态"],
    ...result.pathRows.map((r) => [r.w3, r.w2, r.w1, r.drawingId, r.material, r.color, r.length, r.housing1, r.position1, r.terminal1, r.seal1, r.housing2, r.position2, r.terminal2, r.seal2, r.configs, r.pkgId, r.status])
  ]);

  addSheet(wb, "护套关联矩阵", [
    ["起始护套", "目标护套", "关联导线数", "导线编号", "配置"],
    ...result.housingMatrix.map((r) => [r.housingA, r.housingB, r.count, r.wires, r.configs])
  ]);

  addSheet(wb, "候选预装工作包", [
    ["工作包编号", "层级", "名称/看板号", "导线数", "总长度", "最长导线", "包含护套", "锚点", "配置", "路线候选", "估算工时(秒)", "状态"],
    ...result.packages.map((r) => [r.id, r.kind, r.name, r.wireCount, r.totalLength, r.maxLength, r.housings, r.anchor, r.configs, r.routeType, r.estimatedSeconds, r.status])
  ]);

  addSheet(wb, "孔位责任矩阵", [
    ["工作包", "护套", "孔位", "导线", "端子", "雨塞", "配置", "状态"],
    ...result.positionRows.map((r) => [r.pkgId, r.housing, r.position, r.wire, r.terminal, r.seal, r.configs, r.status])
  ]);

  addSheet(wb, "同色线编码方案", [
    ["工作包", "导线颜色", "导线", "目标护套/孔位", "配置", "标准化编码", "查重", "状态"],
    ...result.dotRows.map((r) => [r.pkgId, r.wireColor, r.wireId, r.target, r.config, r.standardCode, r.checkResult, r.status])
  ]);

  addSheet(wb, "岗位×配置工时矩阵", [
    ["工作包", "名称", "配置", "估算工时(秒)", "TT(秒)", "负荷率%", "建议人数", "备注"],
    ...result.timeRows.map((r) => [r.pkgId, r.pkgName, r.config, r.estimatedSeconds, r.tt, r.loadPercent, r.workerSuggestion, r.note])
  ]);

  addSheet(wb, "冲突与待确认事项", [
    ["序号", "类别", "说明"],
    ...result.issues.map((r, i) => [i + 1, r.category, r.detail])
  ]);

  addSheet(wb, "PDF图纸关键词", [
    ["关键词", "出现次数"],
    ...result.pdfKeywords.map((r) => [r.keyword, r.count])
  ]);

  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
}

async function analyzeFilesForPreview(files, options = {}) {
  if (!files.mbom || !files.mbom.buffer) {
    throw new Error("请先上传MBOM/Cutting文件后再分析。");
  }
  const { validation, canGenerate, mbom } = validateUploadedFiles(files);
  if (!mbom) {
    throw new Error("MBOM/Cutting文件无法解析，请检查模板。");
  }
  const ebom = files.ebom && files.ebom.buffer ? parseEBOM(files.ebom.buffer, files.ebom.originalname) : null;
  const mergedResult = mergeWiresByW1(mbom.wires);
  const specialResult = ebom ? mergeSpecialWires(mergedResult.wires, ebom.specialWires || []) : { wires: mergedResult.wires, conflicts: [], warnings: [] };
  const mergedMbom = {
    ...mbom,
    wires: specialResult.wires,
    rowCount: specialResult.wires.length,
    uniqueW1: new Set(specialResult.wires.map((w) => w.w1).filter(Boolean)).size
  };
  const analysis = buildFileAnalysis(mergedMbom);
  // 按“最高配置工时 ÷ TT”测算预计岗位数（用于前端占位提示；与正式生成口径一致）
  const tt = options.tt != null ? Number(options.tt) : null;
  let estimatedStations = null;
  let maxConfigSeconds = null;
  let configTotals = {};
  if (tt && tt > 0) {
    const standard = files.standard && files.standard.buffer
      ? parseStandardHours(files.standard.buffer, files.standard.originalname)
      : { entries: [] };
    const packages = buildPackages(mergedMbom.wires, mergedMbom.configs, standard.entries);
    configTotals = {};
    for (const cfg of mergedMbom.configs) {
      configTotals[cfg] = round(sum(packages.map((p) => (p.configTime && p.configTime[cfg]) || 0)), 2);
    }
    const vals = Object.values(configTotals).filter((v) => v > 0);
    maxConfigSeconds = vals.length ? Math.max(...vals) : null;
    estimatedStations = maxConfigSeconds ? Math.ceil(maxConfigSeconds / tt) : null;
  }
  return {
    ...analysis,
    mergeConflicts: mergedResult.conflicts,
    mergeWarnings: mergedResult.warnings,
    specialConflicts: specialResult.conflicts,
    specialWarnings: specialResult.warnings,
    validation,
    canGenerate,
    tt: tt || null,
    maxConfigSeconds,
    configTotals,
    estimatedStations
  };
}

async function analyzeWithFiles(files, options = {}) {
  return await analyzeProject(files, options);
}

module.exports = {
  analyzeProject,
  analyzeWithFiles,
  analyzeFilesForPreview,
  buildWorkbook,
  buildProcessWorkbook,
  buildReviewWorkbook
};
