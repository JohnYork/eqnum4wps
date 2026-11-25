/*
 * WPS 公式自动编号 —— 编号列宽度试验 3（只在新文档中执行）
 *
 * 目标：版心宽 100% + 右列按编号宽度，通过“左列 = 版心宽 - 右列宽”实现。
 *
 * 对比：
 *   S1. Cell.Width：右列 36pt，左列 = 版心宽 - 36pt
 *   S2. Columns.Width：右列 36pt，左列 = 版心宽 - 36pt
 *   S3. (10.123)：右列 48pt，左列 = 版心宽 - 48pt
 *   T1. 在 S1 基础上再把表格宽度设回 100%
 *   T2. 在 S1 基础上把表格宽度设为版心宽点值
 *
 * 安全：只在新文档中执行，不修改当前文档；新文档保留打开供目视检查。
 *
 * 使用：在宏编辑器运行入口函数 EQNUM_WIDTH_TEST3()。
 */

function EQNUM_WIDTH_TEST3() {
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

function textWidthPt(d) {
  var r = run(function () {
    var ps = d.PageSetup;
    return Number(ps.PageWidth) - Number(ps.LeftMargin) - Number(ps.RightMargin);
  });
  if (r.ok && !isNaN(Number(r.value)) && Number(r.value) > 0) return Number(r.value);

  // 回退：用 100% 宽表格实测
  var t = run(function () { return d.Tables.Item(d.Tables.Count); });
  if (t.ok && t.value) {
    var w = run(function () { return t.value.Columns.Item(1).Width + t.value.Columns.Item(2).Width; });
    if (w.ok) return Number(w.value);
  }
  return 426.1;
}

function widths(t) {
  var out = {};
  out.tableType = readNum(t, "PreferredWidthType");
  out.tableWidth = readNum(t, "PreferredWidth");
  out.allowAutoFit = run(function () { return t.AllowAutoFit; });
  out.col1 = null;
  out.col2 = null;
  out.cell2Width = null;
  try { out.col1 = Number(t.Columns.Item(1).Width); } catch (e1) {}
  try { out.col2 = Number(t.Columns.Item(2).Width); } catch (e2) {}
  try { out.cell2Width = Number(t.Cell(1, 2).Width); } catch (e3) {}
  out.total = (out.col1 === null || out.col2 === null) ? null : out.col1 + out.col2;
  return out;
}

function logWidths(label, w) {
  log("INFO", label + "：表宽类型=" + w.tableType +
    " 表宽值=" + w.tableWidth +
    " 列1=" + (w.col1 === null ? "?" : String(w.col1) + "pt/" + ptToCm(w.col1)) +
    " 列2=" + (w.col2 === null ? "?" : String(w.col2) + "pt/" + ptToCm(w.col2)) +
    " 合计=" + (w.total === null ? "?" : String(w.total) + "pt/" + ptToCm(w.total)) +
    " 右单元格宽=" + (w.cell2Width === null ? "?" : String(w.cell2Width) + "pt") +
    " AllowAutoFit=" + (w.allowAutoFit.ok ? String(w.allowAutoFit.value) : "?"));
}

function makeTable(d, label, numText) {
  d.Content.InsertAfter("\r" + label + "\r");
  var t = d.Tables.Add(endRange(d), 1, 2);

  try { t.Borders.Enable = 0; } catch (e1) {}
  t.Cell(1, 1).Range.Text = "公式内容占位文字，用来观察左列是否填满剩余宽度";
  t.Cell(1, 2).Range.Text = numText;

  try { t.PreferredWidthType = 2; } catch (e2) {}
  try { t.PreferredWidth = 100; } catch (e3) {}
  try { t.AllowAutoFit = false; } catch (e4) {}

  try { t.TopPadding = 0; } catch (e5) {}
  try { t.BottomPadding = 0; } catch (e6) {}
  try { t.LeftPadding = 0; } catch (e7) {}
  try { t.RightPadding = 0; } catch (e8) {}

  try { t.Cell(1, 1).LeftPadding = 0; } catch (e9) {}
  try { t.Cell(1, 1).RightPadding = 0; } catch (e10) {}
  try { t.Cell(1, 2).LeftPadding = 0; } catch (e11) {}
  try { t.Cell(1, 2).RightPadding = 0; } catch (e12) {}

  try { t.Cell(1, 1).Range.ParagraphFormat.Alignment = 1; } catch (e13) {}
  try { t.Cell(1, 2).Range.ParagraphFormat.Alignment = 2; } catch (e14) {}
  try { t.Cell(1, 1).VerticalAlignment = 1; } catch (e15) {}
  try { t.Cell(1, 2).VerticalAlignment = 1; } catch (e16) {}

  d.Content.InsertAfter("\r");
  return t;
}

function setTwoWidths(d, t, numWidth) {
  var tw = textWidthPt(d);
  var leftWidth = Math.round((tw - numWidth) * 100) / 100;

  var r1 = run(function () {
    t.Cell(1, 2).Width = numWidth;
    return t.Cell(1, 2).Width;
  });
  if (!r1.ok) log("FAIL", "设置右列宽度失败：" + r1.error);
  else log("PASS", "右列 Cell.Width=" + String(numWidth) + "pt，读回=" + String(r1.value));

  var r2 = run(function () {
    t.Cell(1, 1).Width = leftWidth;
    return t.Cell(1, 1).Width;
  });
  if (!r2.ok) log("FAIL", "设置左列宽度失败：" + r2.error);
  else log("PASS", "左列 Cell.Width=" + String(leftWidth) + "pt，读回=" + String(r2.value));

  return { textWidth: tw, leftWidth: leftWidth };
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

  var tw = textWidthPt(d);
  log("INFO", "版心宽=" + String(Math.round(tw * 100) / 100) + "pt/" + ptToCm(tw));
  try { log("INFO", "页面宽=" + String(d.PageSetup.PageWidth) + "pt，左边距=" + String(d.PageSetup.LeftMargin) + "pt，右边距=" + String(d.PageSetup.RightMargin) + "pt"); } catch (e0) {}

  d.Content.InsertAfter("编号列宽度试验 3：版心宽 + 右列按编号宽度\r");

  section("S1. Cell.Width：右列 36pt + 左列 = 版心宽 - 36pt");
  var tS1 = makeTable(d, "S1. Cell.Width 36pt", "(2.1)");
  logWidths("设置前", widths(tS1));
  setTwoWidths(d, tS1, 36);
  logWidths("设置后", widths(tS1));

  section("S2. Columns.Width：右列 36pt + 左列 = 版心宽 - 36pt");
  var tS2 = makeTable(d, "S2. Columns.Width 36pt", "(2.1)");
  logWidths("设置前", widths(tS2));
  var s2a = run(function () {
    tS2.Columns.Item(2).Width = 36;
    tS2.Columns.Item(1).Width = Math.round((tw - 36) * 100) / 100;
    return true;
  });
  if (s2a.ok) log("PASS", "Columns.Width 两列设置成功");
  else log("FAIL", "Columns.Width 两列设置失败：" + s2a.error);
  logWidths("设置后", widths(tS2));

  section("S3. (10.123)：右列 48pt + 左列 = 版心宽 - 48pt");
  var tS3 = makeTable(d, "S3. 长编号 48pt", "(10.123)");
  logWidths("设置前", widths(tS3));
  setTwoWidths(d, tS3, 48);
  logWidths("设置后", widths(tS3));

  section("T1. 在 S1 基础上把表格宽度设回 100%");
  var tT1 = makeTable(d, "T1. 设回 100%", "(2.1)");
  setTwoWidths(d, tT1, 36);
  logWidths("设回前", widths(tT1));
  var t1a = run(function () {
    tT1.PreferredWidthType = 2;
    tT1.PreferredWidth = 100;
    return true;
  });
  if (t1a.ok) log("PASS", "表格宽度设回 100%");
  else log("FAIL", "设回 100% 失败：" + t1a.error);
  logWidths("设回后", widths(tT1));

  section("T2. 在 S1 基础上把表格宽度设为版心宽点值");
  var tT2 = makeTable(d, "T2. 点值版心宽", "(2.1)");
  setTwoWidths(d, tT2, 36);
  logWidths("设点值前", widths(tT2));
  var t2a = run(function () {
    tT2.PreferredWidthType = 3;
    tT2.PreferredWidth = tw;
    return true;
  });
  if (t2a.ok) log("PASS", "表格宽度设为点值 " + String(tw));
  else log("FAIL", "设置点值失败：" + t2a.error);
  logWidths("设点值后", widths(tT2));

  section("汇总");
  log("INFO", "PASS=" + passCount + "，FAIL=" + failCount);
  log("INFO", "判断标准：列1 + 列2 合计接近版心宽 " + String(Math.round(tw * 100) / 100) +
    "pt，且右列宽度等于设定的 36/48pt。");
  log("INFO", "本文档是临时试验文档，不需要保存。");

  var br = run(function () {
    endRange(d).InsertBreak(7);
    return true;
  });
  if (!br.ok) {
    try { d.Content.InsertAfter("\f"); } catch (e2) {}
  }

  var text = "===== 宽度试验 3 结果 =====\r时间：" + new Date().toString() + "\r" + reportLines.join("\r");
  try { d.Content.InsertAfter("\r" + text); } catch (e3) {}

  try {
    var all = d.Content;
    all.Font.Name = "Consolas";
    all.Font.Size = 10;
  } catch (e4) {}

  try {
    if (typeof alert === "function") {
      alert("EQNUM 宽度试验 3 完成：PASS " + passCount + "，FAIL " + failCount +
        "。\r\n请查看 S1/S2/S3 的表格总宽是否与版心一致。");
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
