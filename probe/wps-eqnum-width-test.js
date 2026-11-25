/*
 * WPS 公式自动编号 —— 编号列宽度试验（只在新文档中执行）
 *
 * 目标：
 *   1. 表格宽度固定为版心宽 100%
 *   2. 对比 5 种编号列宽度设置，找出 WPS 真正支持的方式
 *
 * 对比项：
 *   A. 右单元格 PreferredWidthType = 1（自动）
 *   B. 右列 Columns.PreferredWidthType = 1（自动）
 *   C. 右单元格固定点宽（按编号长度估算）
 *   D. 表格 AutoFitBehavior(2)（按窗口宽度）
 *   E. 表格 AutoFitBehavior(1)（按内容）
 *
 * 安全：只在新文档中执行，不修改当前文档；新文档保留打开供目视检查。
 *
 * 使用：在宏编辑器中运行入口函数 EQNUM_WIDTH_TEST()。
 */

function EQNUM_WIDTH_TEST() {
var reportLines = [];
var passCount = 0;
var failCount = 0;

function consoleLog(s) {
  try {
    if (typeof console !== "undefined" && console && console.log) {
      console.log(s);
    }
  } catch (e) {}
}

function log(level, msg) {
  if (level === "PASS") passCount++;
  else if (level === "FAIL") failCount++;
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

function endRange(d) {
  var end = 0;
  try { end = d.Content.End - 1; } catch (e) { end = 0; }
  if (end < 0) end = 0;
  return d.Range(end, end);
}

function ptToCm(pt) {
  var n = Number(pt);
  if (isNaN(n)) return "?";
  return (n / 28.35).toFixed(2) + "cm";
}

function readNum(obj, prop) {
  var r = run(function () { return obj[prop]; });
  return r.ok ? Number(r.value) : null;
}

function widths(t) {
  var out = {};
  out.tableType = readNum(t, "PreferredWidthType");
  out.tableWidth = readNum(t, "PreferredWidth");
  out.allowAutoFit = run(function () { return t.AllowAutoFit; });
  out.colCount = run(function () { return t.Columns.Count; });

  out.col1 = null;
  out.col2 = null;
  try { out.col1 = Number(t.Columns.Item(1).Width); } catch (e1) {}
  try { out.col2 = Number(t.Columns.Item(2).Width); } catch (e2) {}

  out.cell2Type = null;
  out.cell2Pref = null;
  out.cell2Width = null;
  try { out.cell2Type = Number(t.Cell(1, 2).PreferredWidthType); } catch (e3) {}
  try { out.cell2Pref = Number(t.Cell(1, 2).PreferredWidth); } catch (e4) {}
  try { out.cell2Width = Number(t.Cell(1, 2).Width); } catch (e5) {}
  return out;
}

function logWidths(label, w) {
  log("INFO", label + "：表宽类型=" + w.tableType +
    " 表宽值=" + w.tableWidth +
    " 列1=" + (w.col1 === null ? "?" : String(w.col1) + "pt/" + ptToCm(w.col1)) +
    " 列2=" + (w.col2 === null ? "?" : String(w.col2) + "pt/" + ptToCm(w.col2)) +
    " 右单元格类型=" + w.cell2Type +
    " 右单元格偏好值=" + w.cell2Pref +
    " 右单元格宽度=" + (w.cell2Width === null ? "?" : String(w.cell2Width) + "pt/" + ptToCm(w.cell2Width)) +
    " AllowAutoFit=" + (w.allowAutoFit.ok ? String(w.allowAutoFit.value) : "?"));
}

function makeTable(d, label) {
  d.Content.InsertAfter("\r" + label + "\r");
  var t = d.Tables.Add(endRange(d), 1, 2);

  try { t.Borders.Enable = 0; } catch (e1) {}
  t.Cell(1, 1).Range.Text = "公式内容占位文字，用来观察左列是否自动填满剩余宽度";
  t.Cell(1, 2).Range.Text = "(2.1)";

  try { t.PreferredWidthType = 2; } catch (e2) {}
  try { t.PreferredWidth = 100; } catch (e3) {}
  try { t.Cell(1, 1).Range.ParagraphFormat.Alignment = 1; } catch (e4) {}
  try { t.Cell(1, 2).Range.ParagraphFormat.Alignment = 2; } catch (e5) {}
  try { t.Cell(1, 1).VerticalAlignment = 1; } catch (e6) {}
  try { t.Cell(1, 2).VerticalAlignment = 1; } catch (e7) {}

  d.Content.InsertAfter("\r");
  return t;
}

/* ============================================================
 * 主流程
 * ============================================================ */
function main() {
  section("零、脚本信息");
  log("INFO", "开始时间：" + new Date().toString());
  log("INFO", "只在新文档中执行，不修改当前文档。");

  var d = addDocument();
  if (!d) {
    log("FAIL", "无法创建临时文档");
    return;
  }

  d.Content.InsertAfter("编号列宽度试验：比较下面 5 种设置的实际列宽\r");

  section("A. 右单元格 PreferredWidthType = 1（自动）");
  var tA = makeTable(d, "A. 右单元格自动宽度");
  var wA0 = widths(tA);
  logWidths("设置前", wA0);
  var a1 = run(function () {
    tA.Cell(1, 2).PreferredWidthType = 1;
    return tA.Cell(1, 2).PreferredWidthType;
  });
  if (a1.ok) {
    log("PASS", "右单元格 PreferredWidthType 可设为 1，读回=" + String(a1.value));
  } else {
    log("FAIL", "右单元格自动宽度设置失败：" + a1.error);
  }
  logWidths("设置后", widths(tA));

  section("B. 右列 Columns.PreferredWidthType = 1（自动）");
  var tB = makeTable(d, "B. 右列自动宽度");
  logWidths("设置前", widths(tB));
  var b1 = run(function () {
    tB.Columns.Item(2).PreferredWidthType = 1;
    return tB.Columns.Item(2).PreferredWidthType;
  });
  if (b1.ok) {
    log("PASS", "右列 PreferredWidthType 可设为 1，读回=" + String(b1.value));
  } else {
    log("FAIL", "右列自动宽度设置失败：" + b1.error);
  }
  logWidths("设置后", widths(tB));

  section("C. 右单元格固定点宽（按编号长度估算）");
  var tC = makeTable(d, "C. 右单元格固定宽度");
  // "(2.1)" 5 个字符，按 10.5pt 字号估算约 30pt，再加 12pt 内边距
  var estimated = 5 * 6 + 12;
  log("INFO", "估算宽度=" + estimated + "pt/" + ptToCm(estimated));
  var c1 = run(function () {
    tC.Cell(1, 2).PreferredWidthType = 3;
    tC.Cell(1, 2).PreferredWidth = estimated;
    return tC.Cell(1, 2).PreferredWidth;
  });
  if (c1.ok) log("PASS", "右单元格固定点宽设置成功，读回=" + String(c1.value));
  else log("FAIL", "右单元格固定点宽设置失败：" + c1.error);
  logWidths("设置后", widths(tC));

  section("D. 表格 AutoFitBehavior(2)（按窗口宽度）");
  var tD = makeTable(d, "D. 按窗口宽度");
  logWidths("设置前", widths(tD));
  var d1 = run(function () {
    tD.AutoFitBehavior(2);
    return true;
  });
  if (d1.ok) log("PASS", "AutoFitBehavior(2) 可调用");
  else log("FAIL", "AutoFitBehavior(2) 失败：" + d1.error);
  logWidths("设置后", widths(tD));

  section("E. 表格 AutoFitBehavior(1)（按内容）");
  var tE = makeTable(d, "E. 按内容自适应");
  logWidths("设置前", widths(tE));
  var e1 = run(function () {
    tE.AutoFitBehavior(1);
    return true;
  });
  if (e1.ok) log("PASS", "AutoFitBehavior(1) 可调用");
  else log("FAIL", "AutoFitBehavior(1) 失败：" + e1.error);
  logWidths("设置后", widths(tE));

  section("F. 补充：关闭 AllowAutoFit 后设置右单元格固定宽度");
  var tF = makeTable(d, "F. 关闭自动调整 + 固定右列");
  var f1 = run(function () {
    tF.AllowAutoFit = false;
    tF.Cell(1, 2).PreferredWidthType = 3;
    tF.Cell(1, 2).PreferredWidth = estimated;
    return true;
  });
  if (f1.ok) log("PASS", "关闭 AllowAutoFit 并设置固定宽度成功");
  else log("FAIL", "关闭 AllowAutoFit 组合设置失败：" + f1.error);
  logWidths("设置后", widths(tF));

  section("汇总");
  log("INFO", "PASS=" + passCount + "，FAIL=" + failCount);
  log("INFO", "判断标准：设置自动后，列2 宽度应明显小于列1，且接近 (2.1) 的实际宽度。");
  log("INFO", "本文档是临时试验文档，不需要保存。");

  var br = run(function () {
    endRange(d).InsertBreak(7);
    return true;
  });
  if (!br.ok) {
    try { d.Content.InsertAfter("\f"); } catch (e2) {}
  }

  var text = "===== 宽度试验结果 =====\r时间：" + new Date().toString() + "\r" + reportLines.join("\r");
  try { d.Content.InsertAfter("\r" + text); } catch (e3) {}

  try {
    var all = d.Content;
    all.Font.Name = "Consolas";
    all.Font.Size = 10;
  } catch (e4) {}

  try {
    if (typeof alert === "function") {
      alert("EQNUM 宽度试验完成：PASS " + passCount + "，FAIL " + failCount +
        "。\r\n请查看新建的试验文档，重点比较 A、B、C 三张表的右列宽度。");
    }
  } catch (e5) {}
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
