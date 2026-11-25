/*
 * WPS 公式自动编号 —— 编号列宽度试验 2（只在新文档中执行）
 *
 * 目标：找出版心宽 100% + 编号列贴合内容 的可用组合。
 *
 * 对比：
 *   G. 内边距 0 + Cell.Width 直接设 36pt
 *   H. 内边距 0 + Columns.Width 直接设 36pt
 *   I. 内边距 0 + Cell.PreferredWidthType=3 + PreferredWidth=36
 *   J. AutoFitBehavior(1) 量出自然宽度 → 恢复版心宽 → 右列固定为自然宽度+6pt
 *
 * 安全：只在新文档中执行，不修改当前文档；新文档保留打开供目视检查。
 *
 * 使用：在宏编辑器运行入口函数 EQNUM_WIDTH_TEST2()。
 */

function EQNUM_WIDTH_TEST2() {
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
  out.col1 = null;
  out.col2 = null;
  out.cell2Width = null;
  out.cell2Type = null;
  out.cell2Pref = null;
  try { out.col1 = Number(t.Columns.Item(1).Width); } catch (e1) {}
  try { out.col2 = Number(t.Columns.Item(2).Width); } catch (e2) {}
  try { out.cell2Width = Number(t.Cell(1, 2).Width); } catch (e3) {}
  try { out.cell2Type = Number(t.Cell(1, 2).PreferredWidthType); } catch (e4) {}
  try { out.cell2Pref = Number(t.Cell(1, 2).PreferredWidth); } catch (e5) {}
  return out;
}

function logWidths(label, w) {
  log("INFO", label + "：表宽类型=" + w.tableType +
    " 表宽值=" + w.tableWidth +
    " 列1=" + (w.col1 === null ? "?" : String(w.col1) + "pt/" + ptToCm(w.col1)) +
    " 列2=" + (w.col2 === null ? "?" : String(w.col2) + "pt/" + ptToCm(w.col2)) +
    " 右单元格宽=" + (w.cell2Width === null ? "?" : String(w.cell2Width) + "pt/" + ptToCm(w.cell2Width)) +
    " 右单元格类型=" + w.cell2Type +
    " 右单元格偏好值=" + w.cell2Pref +
    " AllowAutoFit=" + (w.allowAutoFit.ok ? String(w.allowAutoFit.value) : "?"));
}

function makeTable(d, label, numText) {
  d.Content.InsertAfter("\r" + label + "\r");
  var t = d.Tables.Add(endRange(d), 1, 2);

  try { t.Borders.Enable = 0; } catch (e1) {}
  t.Cell(1, 1).Range.Text = "公式内容占位文字，用来观察左列是否自动填满剩余宽度";
  t.Cell(1, 2).Range.Text = numText || "(2.1)";

  try { t.PreferredWidthType = 2; } catch (e2) {}
  try { t.PreferredWidth = 100; } catch (e3) {}

  try { t.TopPadding = 0; } catch (e4) {}
  try { t.BottomPadding = 0; } catch (e5) {}
  try { t.LeftPadding = 0; } catch (e6) {}
  try { t.RightPadding = 0; } catch (e7) {}

  try { t.Cell(1, 1).Range.ParagraphFormat.Alignment = 1; } catch (e8) {}
  try { t.Cell(1, 2).Range.ParagraphFormat.Alignment = 2; } catch (e9) {}
  try { t.Cell(1, 1).VerticalAlignment = 1; } catch (e10) {}
  try { t.Cell(1, 2).VerticalAlignment = 1; } catch (e11) {}

  d.Content.InsertAfter("\r");
  return t;
}

function measureNatural(d, label, numText) {
  d.Content.InsertAfter("\r" + label + "\r");
  var t = d.Tables.Add(endRange(d), 1, 2);

  try { t.Borders.Enable = 0; } catch (e1) {}
  t.Cell(1, 1).Range.Text = "公式";
  t.Cell(1, 2).Range.Text = numText;
  try { t.TopPadding = 0; } catch (e2) {}
  try { t.BottomPadding = 0; } catch (e3) {}
  try { t.LeftPadding = 0; } catch (e4) {}
  try { t.RightPadding = 0; } catch (e5) {}

  var af = run(function () {
    t.AutoFitBehavior(1);
    return true;
  });
  if (!af.ok) {
    log("FAIL", "AutoFitBehavior(1) 失败：" + af.error);
    return null;
  }

  var w = null;
  try { w = Number(t.Columns.Item(2).Width); } catch (e6) {}
  log("INFO", numText + " 自然宽度=" + (w === null ? "?" : String(w) + "pt/" + ptToCm(w)));
  d.Content.InsertAfter("\r");
  return { table: t, natural: w };
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

  d.Content.InsertAfter("编号列宽度试验 2：版心宽 + 右列贴合编号\r");

  section("G. 内边距 0 + Cell.Width 直接设 36pt");
  var tG = makeTable(d, "G. Cell.Width = 36pt");
  logWidths("设置前", widths(tG));
  var g1 = run(function () {
    tG.Cell(1, 2).Width = 36;
    return tG.Cell(1, 2).Width;
  });
  if (g1.ok) log("PASS", "Cell.Width 可设置，读回=" + String(g1.value));
  else log("FAIL", "Cell.Width 设置失败：" + g1.error);
  logWidths("设置后", widths(tG));

  section("H. 内边距 0 + Columns.Width 直接设 36pt");
  var tH = makeTable(d, "H. Columns.Width = 36pt");
  logWidths("设置前", widths(tH));
  var h1 = run(function () {
    tH.Columns.Item(2).Width = 36;
    return tH.Columns.Item(2).Width;
  });
  if (h1.ok) log("PASS", "Columns.Width 可设置，读回=" + String(h1.value));
  else log("FAIL", "Columns.Width 设置失败：" + h1.error);
  logWidths("设置后", widths(tH));

  section("I. 内边距 0 + PreferredWidthType=3 + PreferredWidth=36");
  var tI = makeTable(d, "I. PreferredWidth = 36pt");
  logWidths("设置前", widths(tI));
  var i1 = run(function () {
    tI.Cell(1, 2).PreferredWidthType = 3;
    tI.Cell(1, 2).PreferredWidth = 36;
    return tI.Cell(1, 2).PreferredWidth;
  });
  if (i1.ok) log("PASS", "固定偏好宽度可设置，读回=" + String(i1.value));
  else log("FAIL", "固定偏好宽度设置失败：" + i1.error);
  logWidths("设置后", widths(tI));

  section("J. 自然宽度测量 + 恢复版心宽 + 右列固定");
  var s1 = measureNatural(d, "J1. 测量 (2.1) 的自然宽度", "(2.1)");
  if (s1) {
    var tJ1 = s1.table;
    var target1 = Math.ceil(s1.natural + 6);

    var j1a = run(function () {
      tJ1.AllowAutoFit = false;
      tJ1.PreferredWidthType = 2;
      tJ1.PreferredWidth = 100;
      tJ1.Cell(1, 2).Width = target1;
      return true;
    });
    if (j1a.ok) log("PASS", "恢复版心宽并设置右列宽度成功，目标=" + target1 + "pt/" + ptToCm(target1));
    else log("FAIL", "恢复设置失败：" + j1a.error);
    logWidths("J1 设置后", widths(tJ1));
  }

  var s2 = measureNatural(d, "J2. 测量 (10.123) 的自然宽度", "(10.123)");
  if (s2) {
    var tJ2 = s2.table;
    var target2 = Math.ceil(s2.natural + 6);

    var j2a = run(function () {
      tJ2.AllowAutoFit = false;
      tJ2.PreferredWidthType = 2;
      tJ2.PreferredWidth = 100;
      tJ2.Columns.Item(2).Width = target2;
      return true;
    });
    if (j2a.ok) log("PASS", "恢复版心宽并用 Columns.Width 设置右列成功，目标=" + target2 + "pt/" + ptToCm(target2));
    else log("FAIL", "恢复设置失败：" + j2a.error);
    logWidths("J2 设置后", widths(tJ2));
  }

  section("K. 自然宽度测量 + 恢复版心宽 + PreferredWidthType=3");
  var s3 = measureNatural(d, "K1. 测量 (2.1) 备用", "(2.1)");
  if (s3) {
    var tK = s3.table;
    var target3 = Math.ceil(s3.natural + 6);

    var k1 = run(function () {
      tK.AllowAutoFit = false;
      tK.PreferredWidthType = 2;
      tK.PreferredWidth = 100;
      tK.Cell(1, 2).PreferredWidthType = 3;
      tK.Cell(1, 2).PreferredWidth = target3;
      return true;
    });
    if (k1.ok) log("PASS", "恢复版心宽并设置右列偏好宽度成功，目标=" + target3 + "pt/" + ptToCm(target3));
    else log("FAIL", "恢复设置失败：" + k1.error);
    logWidths("K1 设置后", widths(tK));
  }

  section("汇总");
  log("INFO", "PASS=" + passCount + "，FAIL=" + failCount);
  log("INFO", "判断标准：表宽类型=2/100，列1+列2 合计约 426pt（15cm），列2 接近自然宽度+6pt。");
  log("INFO", "本文档是临时试验文档，不需要保存。");

  var br = run(function () {
    endRange(d).InsertBreak(7);
    return true;
  });
  if (!br.ok) {
    try { d.Content.InsertAfter("\f"); } catch (e2) {}
  }

  var text = "===== 宽度试验 2 结果 =====\r时间：" + new Date().toString() + "\r" + reportLines.join("\r");
  try { d.Content.InsertAfter("\r" + text); } catch (e3) {}

  try {
    var all = d.Content;
    all.Font.Name = "Consolas";
    all.Font.Size = 10;
  } catch (e4) {}

  try {
    if (typeof alert === "function") {
      alert("EQNUM 宽度试验 2 完成：PASS " + passCount + "，FAIL " + failCount +
        "。\r\n请查看 G/H/I 与 J/K 表格的右列宽度差异。");
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
