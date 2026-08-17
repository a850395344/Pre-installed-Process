"use strict";

const {
  clean,
  number,
  parseMBOM,
  parseEBOM,
  parseStandardHours
} = require("./parse");

function issue(level, message, field) {
  return { level, message: message || "", field: field || "" };
}

function validateStandard(parsed) {
  const issues = [];
  const required = [
    "工序",
    "工作要素说明",
    "描述",
    "动作开始",
    "动作结束",
    "标准工时",
    "单位"
  ];
  const headerText = (parsed.sheetName ? "" : "") + "";
  const header = parsed.header || [];
  const joined = header.join(" ");

  for (const key of required) {
    if (!joined.includes(key)) {
      issues.push(issue("error", `标准工时模板缺少必需列：${key}`, key));
    }
  }

  if (!parsed.entries || parsed.entries.length === 0) {
    issues.push(issue("error", "标准工时文件中没有可用的数据行", ""));
  } else {
    const noTime = parsed.entries.filter((e) => e.time == null);
    if (noTime.length > 0) {
      issues.push(issue("warning", `有 ${noTime.length} 行标准工时为空或不是数字，请检查。`, "labor time"));
    }
    const noUnit = parsed.entries.filter((e) => !e.unit);
    if (noUnit.length > 0) {
      issues.push(issue("warning", `有 ${noUnit.length} 行缺少单位，请补 pc/mm。`, "unit"));
    }
  }

  return {
    ok: !issues.some((i) => i.level === "error"),
    issues,
    rowCount: parsed.entries ? parsed.entries.length : 0
  };
}

function validateEBOM(parsed) {
  const issues = [];
  const required = [
    "模块号",
    "材料名称",
    "图纸号",
    "图纸用量",
    "单位"
  ];
  const joined = (parsed.header || []).join(" ");

  for (const key of required) {
    if (!joined.includes(key)) {
      issues.push(issue("error", `EBOM模板缺少必需列：${key}`, key));
    }
  }

  if (!parsed.materials || parsed.materials.length === 0) {
    issues.push(issue("error", "EBOM文件中没有可用的物料数据行", ""));
  } else {
    const noDrawing = parsed.materials.filter((m) => !m.drawingId);
    if (noDrawing.length > 0) {
      issues.push(issue("warning", `有 ${noDrawing.length} 行缺少图纸号，请检查。`, "图纸号"));
    }
    const noQty = parsed.materials.filter((m) => m.designQty == null);
    if (noQty.length > 0) {
      issues.push(issue("warning", `有 ${noQty.length} 行缺少图纸用量，请检查。`, "图纸用量"));
    }
    const noUnit = parsed.materials.filter((m) => !m.unit);
    if (noUnit.length > 0) {
      issues.push(issue("warning", `有 ${noUnit.length} 行缺少单位，请检查。`, "单位"));
    }
  }

  return {
    ok: !issues.some((i) => i.level === "error"),
    issues,
    rowCount: parsed.materials ? parsed.materials.length : 0
  };
}

function validateMBOM(parsed) {
  const issues = [];
  const required = [
    "看板号3",
    "看板号2",
    "看板号1",
    "材料名称",
    "图纸号",
    "规格",
    "颜色",
    "下料长度",
    "端子1",
    "端子2",
    "护套1",
    "孔位1",
    "护套2",
    "孔位2"
  ];
  const joined = (parsed.header || []).join(" ");

  for (const key of required) {
    if (!joined.includes(key)) {
      issues.push(issue("error", `MBOM/Cutting模板缺少必需列：${key}`, key));
    }
  }

  if (parsed.configHeaders && parsed.configHeaders.length === 0) {
    issues.push(issue("error", "MBOM/Cutting模板未识别到配置列（例如 8325003CDE8400）。", "配置"));
  }

  if (!parsed.wires || parsed.wires.length === 0) {
    issues.push(issue("error", "MBOM/Cutting文件中没有可用的导线数据行", ""));
  } else {
    const noW1 = parsed.wires.filter((w) => !w.w1);
    if (noW1.length > 0) {
      issues.push(issue("error", `有 ${noW1.length} 行缺少W1看板号。`, "看板号1"));
    }
    const noDrawing = parsed.wires.filter((w) => !w.drawingId);
    if (noDrawing.length > 0) {
      issues.push(issue("warning", `有 ${noDrawing.length} 行缺少图纸号。`, "图纸号"));
    }
    const noLength = parsed.wires.filter((w) => w.length == null);
    if (noLength.length > 0) {
      issues.push(issue("warning", `有 ${noLength.length} 行缺少下料长度。`, "下料长度"));
    }
    const noColor = parsed.wires.filter((w) => !w.color);
    if (noColor.length > 0) {
      issues.push(issue("warning", `有 ${noColor.length} 行缺少颜色。`, "颜色"));
    }
    const noHousing = parsed.wires.filter((w) => (!w.housing1 || w.housing1 === "-") && (!w.housing2 || w.housing2 === "-"));
    if (noHousing.length > 0) {
      issues.push(issue("warning", `有 ${noHousing.length} 行同时缺少护套1和护套2，可能是SP/SC点以外的悬空导线，请检查。`, "护套"));
    }
  }

  return {
    ok: !issues.some((i) => i.level === "error"),
    issues,
    rowCount: parsed.wires ? parsed.wires.length : 0
  };
}

function validateUploadedFiles(files) {
  const validation = {};
  const parsers = {
    standard: { parse: parseStandardHours, label: "标准工时文件" },
    ebom: { parse: parseEBOM, label: "EBOM文件" },
    mbom: { parse: parseMBOM, label: "MBOM/Cutting文件" }
  };

  let mbom = null;

  for (const key of Object.keys(parsers)) {
    const item = parsers[key];
    const file = files[key];
    if (!file || !file.buffer) {
      validation[key] = {
        ok: false,
        issues: [issue("error", `缺少${item.label}，请上传后再分析。`)],
        rowCount: 0
      };
      continue;
    }

    try {
      const parsed = item.parse(file.buffer, file.originalname);
      if (key === "mbom") mbom = parsed;
      if (key === "standard") validation[key] = validateStandard(parsed);
      if (key === "ebom") validation[key] = validateEBOM(parsed);
      if (key === "mbom") validation[key] = validateMBOM(parsed);
    } catch (e) {
      validation[key] = {
        ok: false,
        issues: [issue("error", `模板解析失败：${e.message}`)],
        rowCount: 0
      };
    }
  }

  const canGenerate = Object.values(validation).every((v) => v.ok);

  return { validation, canGenerate, mbom };
}

module.exports = {
  validateUploadedFiles,
  validateStandard,
  validateEBOM,
  validateMBOM
};
