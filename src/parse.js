"use strict";

const XLSX = require("xlsx");
const pdfParse = require("pdf-parse");

function clean(value) {
  return String(value == null ? "" : value).replace(/\u0000/g, "").trim();
}

function number(value) {
  if (value == null || value === "") return null;
  const n = Number(String(value).replace(/[,，]/g, "").trim());
  return Number.isFinite(n) ? n : null;
}

function sheetToRows(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  if (!ws) return [];
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true });
}

function findHeaderRow(rows, predicate) {
  for (let i = 0; i < Math.min(rows.length, 30); i++) {
    if (predicate(rows[i])) return i;
  }
  return -1;
}

function detectConfigHeaders(header) {
  const configs = [];
  for (let i = 20; i < header.length; i++) {
    const h = clean(header[i]);
    if (!h) continue;
    // 配置列通常是项目/零件号，不应包含中文；跳过中文说明列
    if (/[\u4e00-\u9fff]/.test(h)) continue;
    if (/^[A-Za-z0-9][A-Za-z0-9_\-]{3,}$/.test(h)) {
      configs.push({ code: h, index: i });
    }
  }
  return configs;
}

function parseMBOM(buffer, filename) {
  const wb = XLSX.read(buffer, { type: "buffer" });
  const sheetName =
    wb.SheetNames.find((name) => /sheet/i.test(name) && !/wire|线/i.test(name)) ||
    wb.SheetNames[0];
  const rows = sheetToRows(wb, sheetName);
  const headerRowIndex = findHeaderRow(rows, (r) =>
    clean(r[0]).includes("看板号3") &&
    clean(r[4]).includes("图纸号") &&
    clean(r[2]).includes("看板号1")
  );
  if (headerRowIndex < 0) {
    throw new Error(`MBOM文件(${filename})未找到标准表头，请确认是MBOM/Cutting导线表。`);
  }

  const header = rows[headerRowIndex].map(clean);
  const configs = detectConfigHeaders(header);
  const wires = [];
  let currentW3 = "";
  let currentW2 = "";

  for (let i = headerRowIndex + 1; i < rows.length; i++) {
    const r = rows[i];
    const w1 = clean(r[2]);
    const drawing = clean(r[4]);
    if (!w1 && !drawing) continue;

    if (clean(r[0])) {
      // 新的W3块开始；若该行没有W2，则为W3直挂W1，必须清空上一W2
      currentW3 = clean(r[0]);
      currentW2 = clean(r[1]);
    } else if (clean(r[1])) {
      // W2块开始；在W3区域内继续保留当前W3，在W3区域外currentW3保持为空
      currentW2 = clean(r[1]);
    }

    const configMap = {};
    for (const cfg of configs) {
      configMap[cfg.code] = clean(r[cfg.index]).toUpperCase() === "X";
    }

    const wire = {
      w3: currentW3,
      w2: currentW2,
      w1,
      material: clean(r[3]),
      drawingId: drawing,
      customerNo: clean(r[5]),
      jettyNo: clean(r[6]),
      spec: clean(r[7]),
      color: clean(r[8]),
      length: number(r[9]),
      lengthUnit: clean(r[10]) || "mm",
      terminal1: clean(r[11]),
      terminal1Jetty: clean(r[12]),
      seal1: clean(r[13]),
      terminal2: clean(r[14]),
      seal2: clean(r[15]),
      housing1: clean(r[16]) || "",
      position1: clean(r[17]),
      housing2: clean(r[18]) || "",
      position2: clean(r[19]),
      config: configMap,
      configCodes: Object.keys(configMap).filter((k) => configMap[k]),
      status: "已读取"
    };
    wires.push(wire);
  }

  const nonEmptyConfigs = configs
    .map((c) => c.code)
    .filter((code) => wires.some((w) => w.config[code]));

  return {
    sheetName,
    filename,
    header,
    configs: nonEmptyConfigs,
    configHeaders: configs,
    wires,
    rowCount: wires.length,
    uniqueW1: new Set(wires.map((w) => w.w1).filter(Boolean)).size
  };
}

function detectConfigHeadersAnywhere(header) {
  const configs = [];
  for (let i = 0; i < header.length; i++) {
    const h = clean(header[i]);
    if (!h) continue;
    if (/[\u4e00-\u9fff]/.test(h)) continue;
    if (/^[A-Za-z0-9][A-Za-z0-9_\-]{3,}$/.test(h)) {
      configs.push({ code: h, index: i });
    }
  }
  return configs;
}

function parseSpecialWireSheet(rows, headerIndex, sheetName) {
  const header = rows[headerIndex].map(clean);
  const findCol = (name) => header.findIndex((h) => h.includes(name));
  const lineCol = findCol("线号");
  if (lineCol < 0) return [];
  const moduleNoCol = findCol("线束零件号");
  const moduleNameCol = findCol("线束零件名称");
  const drawingCol = lineCol;
  const customerCol = findCol("电线");
  const colorCol = findCol("颜色");
  const lengthCol = findCol("理论长度");
  const h1Col = findCol("A端Housing");
  const t1Col = findCol("A端-端子");
  const s1Col = findCol("A端-防水塞");
  const h2Col = findCol("B端Housing");
  const t2Col = findCol("B端-端子");
  const s2Col = findCol("B端-防水塞");
  const posIndices = header.map((h, i) => [h, i]).filter(([h]) => h.includes("孔位")).map(([, i]) => i);
  const p1Col = posIndices[0] >= 0 ? posIndices[0] : -1;
  const p2Col = posIndices[1] >= 0 ? posIndices[1] : -1;
  const configs = detectConfigHeadersAnywhere(header);

  const wires = [];
  for (let i = headerIndex + 1; i < rows.length; i++) {
    const r = rows[i];
    const lineNo = clean(r[lineCol]);
    if (!lineNo) continue;
    const configMap = {};
    for (const cfg of configs) {
      configMap[cfg.code] = clean(r[cfg.index]).toUpperCase() === "X";
    }
    wires.push({
      w3: "",
      w2: "",
      w1: lineNo,
      material: clean(r[moduleNameCol]) || "特殊线束",
      drawingId: lineNo,
      customerNo: customerCol >= 0 ? clean(r[customerCol]) : "",
      jettyNo: "",
      spec: "",
      color: colorCol >= 0 ? clean(r[colorCol]) : "",
      length: lengthCol >= 0 ? number(r[lengthCol]) : null,
      lengthUnit: "mm",
      terminal1: t1Col >= 0 ? clean(r[t1Col]) : "",
      seal1: s1Col >= 0 ? clean(r[s1Col]) : "",
      housing1: h1Col >= 0 ? clean(r[h1Col]) : "",
      position1: p1Col >= 0 ? clean(r[p1Col]) : "",
      terminal2: t2Col >= 0 ? clean(r[t2Col]) : "",
      seal2: s2Col >= 0 ? clean(r[s2Col]) : "",
      housing2: h2Col >= 0 ? clean(r[h2Col]) : "",
      position2: p2Col >= 0 ? clean(r[p2Col]) : "",
      config: configMap,
      configCodes: Object.keys(configMap).filter((k) => configMap[k]),
      status: "已读取",
      sourceSheet: sheetName
    });
  }
  return wires;
}

function parseEBOM(buffer, filename) {
  const wb = XLSX.read(buffer, { type: "buffer" });
  const allSheets = wb.SheetNames;
  let mainSheet = null;
  let mainHeaderIndex = -1;

  for (const name of allSheets) {
    const rows = sheetToRows(wb, name);
    const idx = findHeaderRow(rows, (r) =>
      clean(r[0]).includes("模块号") &&
      clean(r[2]).includes("材料名称")
    );
    if (idx >= 0) {
      mainSheet = name;
      mainHeaderIndex = idx;
      break;
    }
  }

  if (!mainSheet) {
    throw new Error(`EBOM文件(${filename})未找到主EBOM表头。`);
  }

  const rows = sheetToRows(wb, mainSheet);
  const header = rows[mainHeaderIndex].map(clean);
  const materials = [];

  for (let i = mainHeaderIndex + 1; i < rows.length; i++) {
    const r = rows[i];
    const materialName = clean(r[2]);
    if (!materialName) continue;
    materials.push({
      moduleNo: clean(r[0]),
      moduleName: clean(r[1]),
      materialName,
      description: clean(r[3]),
      drawingId: clean(r[4]),
      jettyNo: clean(r[5]),
      spn: clean(r[6]),
      supplier: clean(r[7]),
      designQty: number(r[8]),
      processAllowance: number(r[9]),
      totalQty: number(r[10]),
      unit: clean(r[11]),
      unitPrice: number(r[12]),
      totalPrice: number(r[13]),
      copperWeight: number(r[14]),
      notes: clean(r[15])
    });
  }

  // Extra sheets containing wire-level tables in EBOM workbook are collected separately.
  const specialSheets = [];
  const specialWires = [];
  for (const name of allSheets) {
    if (name === mainSheet) continue;
    const rows2 = sheetToRows(wb, name);
    const idx = findHeaderRow(rows2, (r) =>
      clean(r[0]).includes("线束零件号") || clean(r[4]).includes("线号")
    );
    if (idx >= 0) {
      const wires = parseSpecialWireSheet(rows2, idx, name);
      specialSheets.push({
        sheetName: name,
        header: rows2[idx].map(clean),
        rowCount: rows2.slice(idx + 1).filter((r) => r.some((c) => clean(c) !== "")).length,
        wireCount: wires.length
      });
      specialWires.push(...wires);
    }
  }

  return {
    filename,
    sheetName: mainSheet,
    header,
    materials,
    rowCount: materials.length,
    specialSheets,
    specialWires
  };
}

function parseStandardHours(buffer, filename) {
  const wb = XLSX.read(buffer, { type: "buffer" });
  const sheetName = wb.SheetNames[0];
  const rows = sheetToRows(wb, sheetName);
  const headerRowIndex = findHeaderRow(rows, (r) =>
    clean(r[5]).includes("labor time") || clean(r[5]).includes("标准工时")
  );
  if (headerRowIndex < 0) {
    throw new Error(`标准工时文件(${filename})未找到“标准工时”列。`);
  }

  const entries = [];
  let currentProcess = "";
  for (let i = headerRowIndex + 1; i < rows.length; i++) {
    const r = rows[i];
    const activity = clean(r[1]);
    const comments = clean(r[2]);
    const time = number(r[5]);
    if (!activity && !comments) continue;
    if (clean(r[0]) === "Porcess" || clean(r[0]).includes("工序")) continue;

    const proc = clean(r[0]);
    if (proc) currentProcess = proc;

    entries.push({
      process: currentProcess,
      activity: activity || currentProcess,
      comments,
      clockStart: clean(r[3]),
      clockStop: clean(r[4]),
      time,
      unit: clean(r[6]) || "pc"
    });
  }

  return {
    filename,
    sheetName,
    header: rows[headerRowIndex].map(clean),
    entries,
    rowCount: entries.length
  };
}

async function parsePDF(buffer, filename) {
  const data = await pdfParse(buffer);
  const text = data.text || "";
  const keywords = [
    "护套", "Housing", "波纹管", "胶带", "橡胶件", "Grommet", "屏蔽", "双绞",
    "节点", "分支", "支架", "卡钉", "扎带", "标签", "保险丝", "继电器",
    "后盖", "盲塞", "防水", "密封"
  ];
  const keywordHits = {};
  for (const kw of keywords) {
    const re = new RegExp(kw.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "g");
    const m = text.match(re);
    if (m) keywordHits[kw] = m.length;
  }

  return {
    filename,
    numPages: data.numpages || 1,
    textLength: text.length,
    text,
    keywordHits,
    status: "已提取文本；版面/图形关系仍需人工或后续AI核对"
  };
}

module.exports = {
  clean,
  number,
  parseMBOM,
  parseEBOM,
  parseStandardHours,
  parsePDF
};
