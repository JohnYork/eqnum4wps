/*
 * WPS 公式自动编号 —— 第一阶段：只读扫描与编号预览
 *
 * 功能：
 *   1. 扫描全文 OMML 公式，只统计 Type=0 的显示公式
 *   2. 按可编号 / 正文混排 / 表格内 / 同段多公式 / 含其他对象 分类
 *   3. 识别已有手工编号，统计旧编号清除量
 *   4. 统计 MathType 公式与 MathType 编号域（保留不动）
 *   5. 预览连续编号、一级章编号、二级章编号
 *
 * 安全：只读，不修改当前文档；结果写入新建报告文档，不保存。
 *
 * 使用：在宏编辑器中运行入口函数 EQNUM_SCAN()。只读扫描约需 1-3 分钟。
 */

function EQNUM_SCAN() {
var MAX_SAMPLE = 12;
var PROGRESS_STEP = 500;
var FIELD_PROGRESS_STEP = 2000;

var reportLines = [];
var infoCount = 0;
var failCount = 0;
var warnCount = 0;
var eligibleCount = 0;

function consoleLog(s) {
  try {
    if (typeof console !== "undefined" && console && console.log) {
      console.log(s);
    }
  } catch (e) {}
}

function log(level, msg) {
  if (level === "FAIL") failCount++;
  else if (level === "WARN") warnCount++;
  else infoCount++;

  var line = "[" + level + "] " + msg;
  reportLines.push(line);
  consoleLog(line);
}

function section(title) {
  reportLines.push("");
  reportLines.push("=== " + title + " ===");
  consoleLog("");
  consoleLog("=== " + title + " ===");
}

function errText(e) {
  if (!e) return "未知错误";
  try { if (e.message) return String(e.message); } catch (e1) {}
  try { if (e.description) return String(e.description); } catch (e2) {}
  try { return String(e); } catch (e3) { return "无法读取的错误"; }
}

function run(fn) {
  try {
    return { ok: true, value: fn() };
  } catch (e) {
    return { ok: false, error: errText(e) };
  }
}

function clip(s, n) {
  s = String(s);
  if (s.length <= n) return s;
  return s.substring(0, n) + "...(共" + s.length + "字符)";
}

function esc(s) {
  s = String(s);
  s = s.replace(/\r/g, "\\r").replace(/\n/g, "\\n").replace(/\t/g, "\\t");
  s = s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, function (ch) {
    var code = ch.charCodeAt(0).toString(16);
    while (code.length < 4) code = "0" + code;
    return "\\u" + code;
  });
  return s;
}

function toText(v) {
  if (v === null) return "null";
  if (typeof v === "undefined") return "undefined";
  var t = "";
  try { t = String(v); } catch (e) { t = "<无法转字符串>"; }
  return esc(clip(t, 120));
}

function status(msg) {
  try { Application.StatusBar = msg; } catch (e) {}
}

function addDocument() {
  var beforeCount = -1;
  try { beforeCount = Application.Documents.Count; } catch (e) {}

  var r = run(function () { return Application.Documents.Add(); });
  if (r.ok && r.value) return r.value;

  var afterCount = -1;
  try { afterCount = Application.Documents.Count; } catch (e2) {}
  if (beforeCount >= 0 && afterCount === beforeCount + 1) {
    try { return Application.ActiveDocument; } catch (e3) {}
  }
  return null;
}

function styleName(p) {
  try { return String(p.Style.NameLocal); } catch (e1) {}
  try { return String(p.Style); } catch (e2) {}
  return "读取失败";
}

function isHeadingLevel(style, level) {
  if (level === 1) return /^(标题|Heading)\s*1$/i.test(style);
  if (level === 2) return /^(标题|Heading)\s*2$/i.test(style);
  return false;
}

function headingNumber(p) {
  var s = "";
  try { s = String(p.Range.ListFormat.ListString); } catch (e1) {}
  s = s.replace(/^\s+|\s+$/g, "");
  if (s) return s;

  // 回退：从标题文本开头的数字解析，如 "2 方法" → "2"
  var t = "";
  try { t = String(p.Range.Text); } catch (e2) {}
  var m = t.match(/^\s*(\d+(?:\.\d+)*)/);
  return m ? m[1] : "";
}

// 去掉允许出现的内容：空白、制表位、段落标记、单元格标记、旧编号
function cleanAllowed(s) {
  s = String(s);
  s = s.replace(/[\(（]\s*\d+(?:\s*[-.]\s*\d+)*\s*[\)）]/g, "");
  s = s.replace(/[\s\u0007]/g, "");
  return s;
}

function findNumberPattern(s) {
  var m = String(s).match(/[\(（]\s*\d+(?:\s*[-.]\s*\d+)*\s*[\)）]/);
  return m ? m[0] : "";
}

function pageOf(p) {
  var r = run(function () { return p.Range.Information(3); });
  return r.ok ? String(r.value) : "?";
}

// 选择 OMaths 最多的文档，避免误选报告文档
function findTargetDocument() {
  var best = null;
  var bestCount = -1;
  var bestName = "";
  var total = 0;
  try { total = Application.Documents.Count; } catch (e) {}

  for (var i = 1; i <= total; i++) {
    var d = null;
    try { d = Application.Documents.Item(i); } catch (e1) { continue; }
    var c = -1;
    try { c = Number(d.OMaths.Count); } catch (e2) { c = -1; }
    if (c > bestCount) {
      bestCount = c;
      best = d;
      try { bestName = String(d.Name); } catch (e3) {}
    }
  }

  var activeName = "";
  try { activeName = String(Application.ActiveDocument.Name); } catch (e4) {}
  if (best && activeName && bestName !== activeName) {
    log("WARN", "活动文档不是公式文档，已自动选择 OMaths 最多的文档：" + bestName + "（OMaths=" + bestCount + "）");
  }
  if (best) return best;

  var r = run(function () { return Application.ActiveDocument; });
  return r.ok ? r.value : null;
}

/* ============================================================
 * 一、文档信息
 * ============================================================ */
function probeDocument(doc) {
  section("一、文档信息");

  var items = [
    ["文档名称", function () { return doc.Name; }],
    ["完整路径", function () { return doc.FullName; }],
    ["段落总数", function () { return doc.Paragraphs.Count; }],
    ["OMaths 总数", function () { return doc.OMaths.Count; }],
    ["InlineShapes 总数", function () { return doc.InlineShapes.Count; }],
    ["Fields 总数", function () { return doc.Fields.Count; }],
    ["Bookmarks 总数", function () { return doc.Bookmarks.Count; }],
    ["Tables 总数", function () { return doc.Tables.Count; }]
  ];

  for (var i = 0; i < items.length; i++) {
    var r = run(items[i][1]);
    if (r.ok) log("INFO", items[i][0] + " = " + toText(r.value));
    else log("FAIL", items[i][0] + " 读取失败：" + r.error);
  }
}

/* ============================================================
 * 二、逐段扫描与分类
 * ============================================================ */
function scanDocument(doc) {
  section("二、公式扫描与分类");

  var total = 0;
  var tr = run(function () { return doc.Paragraphs.Count; });
  if (tr.ok) total = Number(tr.value);
  else {
    log("FAIL", "无法读取段落总数：" + tr.error);
    return null;
  }

  log("INFO", "段落总数=" + total + "，按文档顺序扫描，每 " + PROGRESS_STEP + " 段输出一次进度");

  var stats = {
    displayTotal: 0,
    inlineTotal: 0,
    eligible: [],
    skipTable: 0,
    skipMixed: 0,
    skipMulti: 0,
    skipOther: 0,
    invalid: 0,
    manualNumbers: 0,
    headingCount1: 0,
    headingCount2: 0,
    samplesTable: [],
    samplesMixed: [],
    samplesMulti: [],
    samplesOther: [],
    manualSamples: [],
    chapterCounts: {}
  };

  var h1 = "";
  var h2 = "";

  for (var i = 1; i <= total; i++) {
    if (i % PROGRESS_STEP === 0) {
      status("EQNUM 扫描中：" + i + " / " + total);
      log("INFO", "已扫描段落 " + i + " / " + total);
    }

    var p = null;
    try { p = doc.Paragraphs.Item(i); } catch (e1) { continue; }

    var sn = styleName(p);
    if (isHeadingLevel(sn, 1)) {
      stats.headingCount1++;
      var n1 = headingNumber(p);
      if (n1) {
        h1 = n1;
        h2 = "";
      }
    } else if (isHeadingLevel(sn, 2)) {
      stats.headingCount2++;
      var n2 = headingNumber(p);
      if (n2) h2 = n2;
    }

    var omCount = -1;
    var cr = run(function () { return p.Range.OMaths.Count; });
    if (cr.ok) omCount = Number(cr.value);
    else continue;
    if (omCount <= 0) continue;

    var pStart = -1;
    var pEnd = -1;
    var pText = "";
    try {
      pStart = Number(p.Range.Start);
      pEnd = Number(p.Range.End);
      pText = String(p.Range.Text);
    } catch (e2) {
      continue;
    }

    var inTable = false;
    var itr = run(function () { return p.Range.Information(12); });
    if (itr.ok) {
      inTable = (itr.value === true || itr.value === -1 ||
        String(itr.value) === "true" || String(itr.value) === "-1");
    }

    var shapeCount = -1;
    var sr = run(function () { return p.Range.InlineShapes.Count; });
    if (sr.ok) shapeCount = Number(sr.value);

    // 先收集本段的公式类型，避免把行内公式误算进“同段多公式”
    var displayItems = [];
    var displayInPara = 0;
    for (var j = 1; j <= omCount; j++) {
      var om = null;
      try { om = p.Range.OMaths.Item(j); } catch (e3) {
        stats.invalid++;
        continue;
      }

      var ty = null;
      try { ty = om.Type; } catch (e4) {}

      if (ty !== 0) {
        stats.inlineTotal++;
        continue;
      }

      stats.displayTotal++;
      displayInPara++;
      displayItems.push(om);
    }

    if (displayItems.length === 0) continue;

    for (var k = 0; k < displayItems.length; k++) {
      var dom = displayItems[k];

      var omStart = -1;
      var omEnd = -1;
      try {
        omStart = Number(dom.Range.Start);
        omEnd = Number(dom.Range.End);
      } catch (e5) {
        stats.invalid++;
        continue;
      }

      var before = "";
      var after = "";
      try { before = String(doc.Range(pStart, omStart).Text); } catch (e6) {}
      try { after = String(doc.Range(omEnd, pEnd).Text); } catch (e7) {}

      var existing = findNumberPattern(before) || findNumberPattern(after);
      if (existing) {
        stats.manualNumbers++;
        if (stats.manualSamples.length < MAX_SAMPLE) {
          stats.manualSamples.push({
            para: i,
            style: sn,
            existing: existing,
            text: esc(clip(pText, 80))
          });
        }
      }

      var rec = {
        para: i,
        omIndex: k + 1,
        style: sn,
        page: pageOf(p),
        chapter1: h1 || "0",
        chapter2: h2 || h1 || "0",
        existing: existing,
        text: esc(clip(pText, 80))
      };

      if (inTable) {
        stats.skipTable++;
        if (stats.samplesTable.length < MAX_SAMPLE) stats.samplesTable.push(rec);
      } else if (displayInPara > 1) {
        stats.skipMulti++;
        if (stats.samplesMulti.length < MAX_SAMPLE) stats.samplesMulti.push(rec);
      } else if (shapeCount > 0) {
        stats.skipOther++;
        if (stats.samplesOther.length < MAX_SAMPLE) stats.samplesOther.push(rec);
      } else if (cleanAllowed(before).length > 0 || cleanAllowed(after).length > 0) {
        stats.skipMixed++;
        if (stats.samplesMixed.length < MAX_SAMPLE) stats.samplesMixed.push(rec);
      } else {
        stats.eligible.push(rec);
        stats.chapterCounts[rec.chapter1] = (stats.chapterCounts[rec.chapter1] || 0) + 1;
      }
    }
  }

  status("");
  return stats;
}

/* ============================================================
 * 三、编号预览
 * ============================================================ */
function buildPreviews(stats) {
  section("三、编号预览");

  if (!stats || stats.eligible.length === 0) {
    log("WARN", "没有可编号的独立显示公式");
    return;
  }

  var c1 = {};
  var c2 = {};

  for (var i = 0; i < stats.eligible.length; i++) {
    var rec = stats.eligible[i];

    rec.numberContinuous = String(i + 1);

    var key1 = rec.chapter1;
    c1[key1] = (c1[key1] || 0) + 1;
    rec.numberChapter1 = key1 + "." + c1[key1];

    var key2 = rec.chapter2;
    c2[key2] = (c2[key2] || 0) + 1;
    rec.numberChapter2 = key2 + "." + c2[key2];
  }

  log("INFO", "可编号公式共 " + stats.eligible.length + " 个");

  log("INFO", "连续编号：");
  appendPreview(stats.eligible, "numberContinuous", "(");

  log("INFO", "一级章编号（标题 1）：");
  appendPreview(stats.eligible, "numberChapter1", "(");

  log("INFO", "二级章编号（标题 2）：");
  appendPreview(stats.eligible, "numberChapter2", "(");

  log("INFO", "一级章节分布：");
  var count = 0;
  for (var key in stats.chapterCounts) {
    if (stats.chapterCounts.hasOwnProperty(key)) {
      count++;
      if (count <= 30) {
        log("INFO", "  第 " + key + " 章：" + stats.chapterCounts[key] + " 个可编号公式");
      }
    }
  }
  if (count > 30) log("INFO", "  其余章节省略");
}

function appendPreview(list, field, open) {
  var n = Math.min(list.length, 15);
  var parts = [];
  for (var i = 0; i < n; i++) {
    parts.push(open + list[i][field] + ")");
  }
  log("INFO", "  前 " + n + " 个：" + parts.join("、"));
}

/* ============================================================
 * 四、跳过样例
 * ============================================================ */
function reportSamples(stats) {
  section("四、跳过原因与样例");

  log("INFO", "统计：可编号 " + stats.eligible.length +
    "，正文混排 " + stats.skipMixed +
    "，表格内 " + stats.skipTable +
    "，同段多公式 " + stats.skipMulti +
    "，含其他对象 " + stats.skipOther +
    "，读取失败 " + stats.invalid);

  logSampleGroup("正文混排", stats.samplesMixed);
  logSampleGroup("表格内", stats.samplesTable);
  logSampleGroup("同段多公式（子公式组候选）", stats.samplesMulti);
  logSampleGroup("含图片或其他对象", stats.samplesOther);
}

function logSampleGroup(title, list) {
  log("INFO", "--- " + title + " 样例（最多 " + MAX_SAMPLE + " 个）---");
  if (!list || list.length === 0) {
    log("INFO", "  无");
    return;
  }
  for (var i = 0; i < list.length; i++) {
    var r = list[i];
    log("INFO", "  段 #" + r.para +
      " 第 " + r.page + " 页" +
      " 样式=" + r.style +
      " 已有编号=" + (r.existing ? r.existing : "无"));
    log("INFO", "    文本=" + r.text);
  }
}

/* ============================================================
 * 五、现有手工编号统计
 * ============================================================ */
function reportManualNumbers(stats) {
  section("五、现有手工编号");

  log("INFO", "显示公式段落中检测到旧编号共 " + stats.manualNumbers + " 处");

  if (stats.manualSamples.length === 0) {
    log("INFO", "  未发现旧编号样例");
    return;
  }

  for (var i = 0; i < stats.manualSamples.length; i++) {
    var r = stats.manualSamples[i];
    log("INFO", "  段 #" + r.para + " 样式=" + r.style + " 旧编号=" + r.existing);
    log("INFO", "    文本=" + r.text);
  }

  log("INFO", "说明：实施全文重编号时，将清除这些旧编号文本；MathType 编号域不在清除范围内。");
}

/* ============================================================
 * 六、MathType 公式统计
 * ============================================================ */
function scanMathType(doc) {
  section("六、MathType 公式（保留旧编号，不参与本次编号）");

  var total = 0;
  var cr = run(function () { return doc.InlineShapes.Count; });
  if (cr.ok) total = Number(cr.value);
  else {
    log("FAIL", "无法读取 InlineShapes.Count：" + cr.error);
    return;
  }

  var mathType = 0;
  var display = 0;
  var inline = 0;
  var inTable = 0;
  var samples = 0;

  for (var i = 1; i <= total; i++) {
    var shape = null;
    try { shape = doc.InlineShapes.Item(i); } catch (e1) { continue; }

    var prog = "";
    try { prog = String(shape.OLEFormat.ProgID); } catch (e2) {}
    if (prog.indexOf("Equation") < 0) continue;

    mathType++;

    var p = null;
    try { p = shape.Range.Paragraphs.Item(1); } catch (e3) { continue; }

    var sn = styleName(p);
    var isDisplay = /MTDisplay|Equation Number/i.test(sn);

    var txt = "";
    try { txt = String(p.Range.Text); } catch (e4) {}
    if (!isDisplay && /^\s*\t+\u0001\t+[\s\u0015]*\r?\s*$/.test(txt)) {
      isDisplay = true;
    }

    if (isDisplay) display++;
    else inline++;

    var itr = run(function () { return p.Range.Information(12); });
    if (itr.ok && (itr.value === true || itr.value === -1 ||
      String(itr.value) === "true" || String(itr.value) === "-1")) {
      inTable++;
    }

    if (samples < MAX_SAMPLE) {
      samples++;
      log("INFO", "  MathType #" + i + " ProgID=" + prog +
        " 样式=" + sn +
        " 显示公式=" + (isDisplay ? "是" : "否") +
        " 在表格内=" + (itr.ok ? String(itr.value) : "未知"));
      log("INFO", "    文本=" + esc(clip(txt, 100)));
    }
  }

  log("INFO", "MathType 公式共 " + mathType + " 个：显示 " + display + "，行内 " + inline + "，其中表格内 " + inTable);
  log("INFO", "按当前方案，这些公式保留原有 MathType 编号，不纳入新编号体系。");
}

/* ============================================================
 * 七、MathType 编号域统计
 * ============================================================ */
function scanNumberingFields(doc) {
  section("七、现有 MathType 编号域（保留不动）");

  var total = 0;
  var cr = run(function () { return doc.Fields.Count; });
  if (cr.ok) total = Number(cr.value);
  else {
    log("FAIL", "无法读取 Fields.Count：" + cr.error);
    return;
  }

  log("INFO", "域总数=" + total + "，逐域扫描，每 " + FIELD_PROGRESS_STEP + " 个输出一次进度");

  var zEqn = 0;
  var mtPlace = 0;
  var mtSeq = 0;
  var seqTotal = 0;
  var broken = 0;
  var zSamples = 0;
  var brokenSamples = 0;

  for (var i = 1; i <= total; i++) {
    if (i % FIELD_PROGRESS_STEP === 0) {
      status("EQNUM 扫描域：" + i + " / " + total);
      log("INFO", "已扫描域 " + i + " / " + total);
    }

    var f = null;
    try { f = doc.Fields.Item(i); } catch (e1) { continue; }

    var code = "";
    try { code = String(f.Code.Text); } catch (e2) { continue; }

    if (/ZEqnNum/i.test(code)) {
      zEqn++;
      if (zSamples < 10) {
        zSamples++;
        var res = "";
        try { res = String(f.Result.Text); } catch (e3) {}
        log("INFO", "  ZEqnNum 域 #" + i + " 代码=" + esc(clip(code, 80)) + " 结果=" + esc(clip(res, 30)));
      }
    }
    if (/MTPlaceRef/i.test(code)) mtPlace++;
    if (/SEQ\s+MTEqn|SEQ\s+MTSec/i.test(code)) mtSeq++;
    if (/\bSEQ\b/i.test(code)) seqTotal++;

    var result = "";
    try { result = String(f.Result.Text); } catch (e4) {}
    if (result.indexOf("错误") >= 0 || /error!/i.test(result)) {
      broken++;
      if (brokenSamples < 5) {
        brokenSamples++;
        log("INFO", "  损坏域 #" + i + " 代码=" + esc(clip(code, 60)) + " 结果=" + esc(clip(result, 30)));
      }
    }
  }

  status("");

  log("INFO", "含 ZEqnNum 的域：" + zEqn);
  log("INFO", "含 MTPlaceRef 的域：" + mtPlace);
  log("INFO", "含 SEQ MTEqn / SEQ MTSec 的域：" + mtSeq);
  log("INFO", "全部 SEQ 域（包含图表编号）：" + seqTotal);
  log("INFO", "结果包含“错误”的域：" + broken);
  log("INFO", "结论：这些 MathType 编号域按方案保留不动，避免破坏现有引用。");
}

/* ============================================================
 * 八、汇总与报告
 * ============================================================ */
function summary(stats) {
  section("八、汇总");

  if (!stats) {
    log("FAIL", "扫描失败，无法给出汇总");
    return;
  }

  log("INFO", "OMaths 总数 = " + (stats.displayTotal + stats.inlineTotal));
  log("INFO", "显示公式 Type=0 = " + stats.displayTotal);
  log("INFO", "行内公式 Type=1 = " + stats.inlineTotal);
  log("INFO", "可编号（独立段落、非表格）= " + stats.eligible.length);
  log("INFO", "跳过：正文混排 = " + stats.skipMixed);
  log("INFO", "跳过：表格内 = " + stats.skipTable);
  log("INFO", "跳过：同段多公式 = " + stats.skipMulti);
  log("INFO", "跳过：含图片或其他对象 = " + stats.skipOther);
  log("INFO", "读取失败 = " + stats.invalid);
  log("INFO", "旧编号检测 = " + stats.manualNumbers);
  log("INFO", "标题 1 数量 = " + stats.headingCount1 + "，标题 2 数量 = " + stats.headingCount2);
  log("INFO", "下一步：确认这些数字后，再实现“编号当前公式”和“全文重编号”。");
}

function finish() {
  log("INFO", "扫描完成：INFO=" + infoCount + "，WARN=" + warnCount + "，FAIL=" + failCount);

  var text = "EQNUM 只读扫描报告\r\n时间：" + new Date().toString() + "\r\n" + reportLines.join("\r\n");
  var reportDoc = addDocument();

  if (reportDoc) {
    try {
      reportDoc.Content.Text = text;
      try { reportDoc.Content.Font.Name = "Consolas"; } catch (e1) {}
      try { reportDoc.Content.Font.Size = 10; } catch (e2) {}
      log("INFO", "报告已写入新建文档，可直接复制内容反馈");
    } catch (e) {
      log("WARN", "写入报告文档失败：" + errText(e));
    }
  } else {
    log("WARN", "无法创建报告文档，请直接复制宏编辑器日志");
  }

  try {
    if (typeof alert === "function") {
      alert("EQNUM 只读扫描完成：可编号 " + eligibleCount +
        " 个。\r\n完整统计见报告文档与宏编辑器日志。");
    }
  } catch (e3) {}
}

/* ============================================================
 * 主流程
 * ============================================================ */
function main() {
  section("零、脚本信息");
  log("INFO", "开始时间：" + new Date().toString());
  log("INFO", "只读扫描：不修改当前文档。");

  var doc = findTargetDocument();
  if (!doc) {
    log("FAIL", "没有可用的文档");
    return;
  }

  probeDocument(doc);

  var stats = scanDocument(doc);
  if (stats) {
    buildPreviews(stats);
    reportSamples(stats);
    reportManualNumbers(stats);
    scanMathType(doc);
    scanNumberingFields(doc);
    summary(stats);
    eligibleCount = stats.eligible.length;
  }

  status("");
  finish();
}

try {
  main();
} catch (e) {
  consoleLog("脚本执行异常：" + errText(e));
  try {
    if (typeof alert === "function") alert("脚本执行异常：" + errText(e));
  } catch (e2) {}
}
}
