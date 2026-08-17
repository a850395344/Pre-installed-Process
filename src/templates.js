"use strict";

const XLSX = require("xlsx");

function addSheet(wb, name, aoa) {
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  XLSX.utils.book_append_sheet(wb, ws, name);
}

function buildStandardTemplate() {
  const wb = XLSX.utils.book_new();
  addSheet(wb, "标准工时模板", [
    ["Porcess\n工序", "Activity Description\n工作要素说明", "Comments\n描述", "Clock Start\n动作开始", "Clock Stop\n动作结束", "labor time\n标准工时", "unit\n单位"],
    ["Build connectors to board\n护套放在案板上", "Assemble connectors\n护套组装", "1-6 way & Ring to V-pin\n将 1~6 孔的塑件/环端放在叉子上", "Reach for connector\n取护套", "Release connector after placing to V-pin\n将护套放在治具上", 1.7, "pc"],
    ["", "", "7~40 way to connector holder\n将 7~40 孔的塑件放在塑件托架上", "Reach for connector\n取护套", "Release connector after placing to connector holder\n将护套放在治具上", 4, "pc"],
    ["Get and Route wires ( get and remove wires from rack operation to be used only for uncoiled wires)\n取/布导线(仅适用于没有打圈的导线)", "remove wires from rack\n取线", "导线长度0-1000 mm", "Reach for wire\n伸手拿导线", "Wire in hand ready for first plug\n导线拿在手中", 1.625, "pc"],
    ["", "Route wires on board\n布线", "导线长度0-1000 mm", "Wire in hand after first plug\n手中导线布线开始", "Wire in hand ready for second plug\n手中导线布线完成", 1.46, "pc"],
    ["Pluging\n插接/插盲堵相关", "Pluging导线插接", "Unsealed terminals to unsealed connector\n不带密封圈的端子插入不带密封圈塑件", "将端子与护套孔位对齐", "端子插入护套并松开导线", 2.5, "pc"],
    ["Manual taping\n手动缠绕", "Apply spot tape\n点胶带", "Single spot tape单个点胶带", "Start tape\n开始缠绕", "Release taped branch\n固定到分支", 4.2, "pc"]
  ]);
  addSheet(wb, "编写规则", [
    ["标准工时文件编写规则"],
    ["列名", "是否必填", "单位", "说明"],
    ["工序 Porcess", "必填", "", "后段预装动作分组，例如：护套放在案板、取/布导线、插接、手动缠绕、波纹管、橡胶件、扎带/卡扣、电测、包装、外检等。"],
    ["工作要素说明 Activity Description", "必填", "", "具体工作要素名称。"],
    ["描述 Comments", "必填", "", "补充条件，例如导线长度区间、护套孔数范围、是否带密封圈等。"],
    ["动作开始 Clock Start", "必填", "", "动作开始描述。"],
    ["动作结束 Clock Stop", "必填", "", "动作结束描述。"],
    ["标准工时 labor time", "必填", "pc/mm", "数值；每件填 pc，每毫米填 mm。"],
    ["单位 unit", "必填", "", "pc 或 mm。"],
    [""],
    ["注意事项"],
    ["1. 该文件是正式工时唯一来源，软件不会自动生成经验工时。"],
    ["2. 前端裁线/剥皮/端子压接等动作请单独标注为“前端工时，不纳入本次后段工艺计算”。"],
    ["3. 内容行可以任意增加，但必须保留以上7列表头。"],
    ["4. 工序列在合并单元格时，只有首行填写工序名，后续行留空即可。"]
  ]);
  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
}

function buildEBOMTemplate() {
  const wb = XLSX.utils.book_new();
  addSheet(wb, "EBOM模板", [
    ["模块号/Harness Family Part No.", "模块名称/Harness Family Name", "材料名称/Part Name", "Description", "图纸号/Drawing ID", "捷翼号/Jetty No.(生产)", "厂家号/SPN(生产）", "厂家/Supplier(生产）", "图纸用量/Design Quantity", "工艺余量/Process Allowance", "总用量/Total Quantity", "单位/Unit（mm、p、g）", "单价/Unit Price", "价格汇总/Total Price", "理论铜重/Copper Weight", "备注/Notes"],
    ["8325003CDE8400", "车身线束", "护套", "Housing", "W25LH1", "1HS-07618", "AH7341F-0.64/2.2-21", "奥海", 1, "", "", "p", "", "", "", ""],
    ["8325003CDE8400", "车身线束", "胶带", "Tape", "SP797", "1TP-00352", "CH268-38*50m-R", "凯密科", 60, "", "", "mm", "", "", "", "焊点胶带"],
    ["8325003CDE8400", "车身线束", "波纹管", "Conduit", "CT-00146-140", "1CT-00146", "内径6.4", "示例供应商", 140, "", "", "mm", "", "", "", ""]
  ]);
  addSheet(wb, "编写规则", [
    ["EBOM文件编写规则"],
    ["列名", "是否必填", "单位", "说明"],
    ["模块号", "必填", "", "线束零件号/项目号。"],
    ["模块名称", "可空", "", "例如：车身线束。"],
    ["材料名称", "必填", "", "护套、胶带、波纹管、扎带、橡胶件、后盖、保险丝、标签等。"],
    ["Description", "可空", "", "英文/规格描述。"],
    ["图纸号", "必填", "", "物料图纸号/物料编号。"],
    ["捷翼号", "可空", "", "生产用捷翼号。"],
    ["厂家号", "可空", "", "SPN/厂家号。"],
    ["厂家", "可空", "", "供应商名称。"],
    ["图纸用量", "必填", "", "单个线束的设计用量。"],
    ["工艺余量", "可空", "", "如没有可留空。"],
    ["总用量", "可空", "", "优先填写；为空时软件会提示待确认。"],
    ["单位", "必填", "mm/p/g", "mm、p 或 g。"],
    ["单价/价格汇总/理论铜重", "可空", "", "成本相关，不参与工艺拆分。"],
    ["备注", "可空", "", "例如焊点胶带、客户要求等。"],
    [""],
    ["注意事项"],
    ["1. 每个物料一行。"],
    ["2. 特殊线束表（铝线/气囊线/以太网/FAKRA等）请放在额外Sheet中，并保留“线束零件号、线号、A端/B端、孔位、电线、颜色、长度、配置”等列。"]
  ]);
  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
}

function buildMBOMTemplate() {
  const configs = ["8325003CDE8400", "8325003CDE8600", "8325003CDE8500", "8325003CDE9100"];
  const header = [
    "看板号3", "看板号2", "看板号1", "材料名称", "图纸号", "客户号", "客户号Jetty码",
    "规格", "颜色", "下料长度_绞合后", "单位", "端子1", "端子1Jetty码", "雨塞1",
    "端子2", "雨塞2", "护套1", "孔位1", "护套2", "孔位2", ...configs, "", ""
  ];
  const wb = XLSX.utils.book_new();
  addSheet(wb, "MBOM-Cutting模板", [
    header,
    ["W3T09Y1100001", "W2T09Y1100084", "W1T09Y1100811", "压接线", "G404-21B", "FLRY-B0.35B", "1FL-05617", "0.35", "黑", 1555, "mm", "AH617-1.5X0.8A", "1TM-03702", "-", "*TM-1011000", "-", "B860A", "2", "SP4269", "L", "", "X", "X", "X", "", ""],
    ["", "", "W1T09Y1100001", "铝线", "30031A", "FLALR2X-B12R", "1ZC-04470", "12", "红", 4620, "mm", "1-2477962-1", "-", "-", "RSD2H0657", "-", "Z301C6", "1", "P601R4", "1", "X", "X", "X", "X", "", ""]
  ]);
  addSheet(wb, "编写规则", [
    ["MBOM/Cutting文件编写规则"],
    ["列名", "是否必填", "单位", "说明"],
    ["看板号3", "可空", "", "三级复杂组合；没有可留空。"],
    ["看板号2", "可空", "", "二级看板/半成品组合；没有可留空。"],
    ["看板号1", "必填", "", "具体导线编号 W1。"],
    ["材料名称", "必填", "", "铝线、单线、压接线、双绞线、屏蔽线、气囊线等。"],
    ["图纸号", "必填", "", "导线图号/回路号。"],
    ["客户号", "可空", "", "线缆规格代码，例如 FLRY-B0.35。"],
    ["客户号Jetty码", "可空", "", "捷翼物料号。"],
    ["规格", "必填", "", "线径，例如 0.35 / 0.5 / 2.5。"],
    ["颜色", "必填", "", "完整线色。"],
    ["下料长度_绞合后", "必填", "mm", "绞合后的下料长度。"],
    ["单位", "必填", "mm", "通常为 mm。"],
    ["端子1/端子2", "必填", "", "两端端子号；没有端子或为SP/SC焊点可填“-”或“超声波焊接”。"],
    ["雨塞1/雨塞2", "可空", "", "没有填“-”。"],
    ["护套1/护套2", "必填", "", "两端护套号；SP/SC压接点也填在此列，例如 SP030、SC050。"],
    ["孔位1/孔位2", "必填", "", "护套孔位或SP/SC点位。"],
    ["配置列", "必填", "", "每个配置一列，该配置下有此导线填 X。"],
    [""],
    ["注意事项"],
    ["1. 同一W1跨多个W3/W2且配置不同时，可以多行存在，软件会按W1+配置合并。"],
    ["2. 同一配置下同一个W1如果出现两次，并且端子或长度不一致，软件会判定为真冲突。"],
    ["3. 没有W3/W2时，W1也可以直接挂接。"],
    ["4. 配置列必须与EBOM、图纸中的项目配置号一致。"]
  ]);
  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
}

function buildTemplate(type) {
  if (type === "standard") return buildStandardTemplate();
  if (type === "ebom") return buildEBOMTemplate();
  if (type === "mbom") return buildMBOMTemplate();
  throw new Error(`未知模板类型：${type}`);
}

module.exports = { buildTemplate };
