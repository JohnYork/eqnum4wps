/*
 * WPS 公式自动编号 —— 第二轮探测：OMaths 结构、现有编号机制、光标定位（只读）
 *
 * 安全约定：
 *   1. 不修改任何文档，只读取。
 *   2. 报告写入新建文档，不保存。
 *   3. 每个 API 单独 try/catch，失败不中断。
 *
 * 使用：在宏编辑器中运行入口函数 EQNUM_PROBE2()。
 */

function EQNUM_PROBE2() {
var MAX_SAMPLE = 15;
var PROGRESS_STEP = 2000;

var reportLines = [];
var passCount = 0;
var failCount = 0;
var warnCount = 0;
var infoCount = 0;

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

function readValue(label, fn) {
  var r = run(fn);
  if (r.ok) {
    log("INFO", label + " = " + toText(r.value));
    return r.value;
  }
  log("FAIL", label + " 不可用：" + r.error);
  return null;
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

function paragraphStyleName(p) {
  try { return String(p.Style.NameLocal); } catch (e1) {}
  try { return String(p.Style); } catch (e2) {}
  return "读取失败";
}

// 在多份文档中选出 OMaths 最多的那个，避免误选探测报告文档
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
    try { c = d.OMaths.Count; } catch (e2) { c = -1; }
    if (c > bestCount) {
      bestCount = c;
      best = d;
      try { bestName = String(d.Name); } catch (e3) {}
    }
  }

  var active = null;
  var activeName = "";
  try {
    active = Application.ActiveDocument;
    activeName = String(active.Name);
  } catch (e4) {}

  if (best && activeName && bestName !== activeName) {
    log("INFO", "活动文档不是公式文档，已自动选择 OMaths 最多的文档：" + bestName + "（OMaths=" + bestCount + "）");
  }
  if (best) return best;

  var r = run(function () { return Application.ActiveDocument; });
  return r.ok ? r.value : null;
}

/* ============================================================
 * 一、目标文档
 * ============================================================ */
function probeTarget(doc) {
  section("一、目标文档");

  readValue("文档名称", function () { return doc.Name; });
  readValue("完整路径", function () { return doc.FullName; });
  readValue("OMaths 总数", function () { return doc.OMaths.Count; });
  readValue("InlineShapes 总数", function () { return doc.InlineShapes.Count; });
  readValue("Fields 总数", function () { return doc.Fields.Count; });
  readValue("Bookmarks 总数", function () { return doc.Bookmarks.Count; });
  readValue("Paragraphs 总数", function () { return doc.Paragraphs.Count; });
}

/* ============================================================
 * 二、OMaths 类型统计
 * ============================================================ */
function probeOMathTypes(doc) {
  section("二、OMaths 类型统计（wdOMathDisplay=0，wdOMathInline=1）");

  var total = 0;
  var cr = run(function () { return doc.OMaths.Count; });
  if (!cr.ok) {
    log("FAIL", "无法读取 OMaths.Count：" + cr.error);
    return null;
  }
  total = Number(cr.value);

  var display = 0;
  var inline = 0;
  var unknown = 0;
  var typeValues = {};
  var displayIndexes = [];

  var t0 = new Date().getTime();

  for (var i = 1; i <= total; i++) {
    if (i % PROGRESS_STEP === 0) {
      log("INFO", "已扫描 " + i + " / " + total);
    }

    var om = null;
    try { om = doc.OMaths.Item(i); } catch (e1) { unknown++; continue; }

    var ty = null;
    try { ty = om.Type; } catch (e2) {}
    typeValues[String(ty)] = (typeValues[String(ty)] || 0) + 1;

    if (ty === 0) {
      display++;
      displayIndexes.push(i);
    } else if (ty === 1) {
      inline++;
    } else {
      unknown++;
    }
  }

  var elapsed = new Date().getTime() - t0;

  log("INFO", "OMaths 总数=" + total);
  log("INFO", "显示公式 Type=0：" + display);
  log("INFO", "行内公式 Type=1：" + inline);
  log("INFO", "无法读取或未知类型：" + unknown);
  log("INFO", "Type 取值分布：" + toText(typeValues));
  log("INFO", "类型统计耗时 " + elapsed + " ms");

  return { total: total, display: display, inline: inline, displayIndexes: displayIndexes };
}

/* ============================================================
 * 三、显示公式结构抽样
 * ============================================================ */
function probeDisplaySamples(doc, stats) {
  section("三、显示公式结构抽样");

  if (!stats || stats.displayIndexes.length === 0) {
    log("WARN", "没有识别到 Type=0 的显示公式，需要改用其他特征（段落样式/独占段落/制表位）判断");
    return;
  }

  var styleStat = {};
  var inTableCount = 0;
  var numberPatternCount = 0;
  var fieldCount = 0;
  var sampleCount = 0;

  var list = stats.displayIndexes;
  var limit = Math.min(list.length, 1500);

  for (var k = 0; k < limit; k++) {
    var idx = list[k];
    var om = null;
    try { om = doc.OMaths.Item(idx); } catch (e1) { continue; }

    var p = null;
    try { p = om.Range.Paragraphs.Item(1); } catch (e2) { continue; }

    var sn = paragraphStyleName(p);
    styleStat[sn] = (styleStat[sn] || 0) + 1;

    var inTable = null;
    try { inTable = p.Range.Information(12); } catch (e3) {}
    if (inTable === true || String(inTable) === "true" || String(inTable) === "-1") inTableCount++;

    var txt = "";
    try { txt = String(p.Range.Text); } catch (e4) {}
    var stripped = txt.replace(/\u0001/g, "");

    var hasNumber = /[\(（]\s*\d+\s*[-.]\s*\d+\s*[\)）]/.test(stripped);
    if (hasNumber) numberPatternCount++;

    var fcount = -1;
    try { fcount = om.Range.Fields.Count; } catch (e5) {}
    if (fcount > 0) fieldCount++;

    if (sampleCount < MAX_SAMPLE) {
      sampleCount++;
      var listStr = "";
      var start = "";
      var end = "";
      var bm = "";
      try { listStr = String(p.Range.ListFormat.ListString); } catch (e6) {}
      try { start = String(om.Range.Start); } catch (e7) {}
      try { end = String(om.Range.End); } catch (e8) {}
      try { bm = String(om.Range.Bookmarks.Count); } catch (e9) {}

      log("INFO", "样本 #OMath" + idx +
        " 样式=" + sn +
        " 在表格内=" + inTable +
        " 章节号=" + (listStr ? listStr : "(空)") +
        " Range=" + start + "-" + end +
        " 域数=" + fcount +
        " 书签数=" + bm +
        " 疑似已有编号=" + hasNumber);
      log("INFO", "    段落文本=" + esc(clip(txt, 120)));
      try {
        log("INFO", "    公式 Range 文本=" + esc(clip(String(om.Range.Text), 60)));
      } catch (e10) {}
    }
  }

  log("INFO", "已分析显示公式 " + limit + " 个" +
    (list.length > limit ? "（共 " + list.length + " 个，仅分析前 " + limit + " 个）" : ""));
  log("INFO", "所在样式分布：" + toText(styleStat));
  log("INFO", "位于表格内的显示公式：" + inTableCount);
  log("INFO", "段落文本中已含 (数字-数字) 编号特征：" + numberPatternCount);
  log("INFO", "公式 Range 内包含域的：" + fieldCount);
}

/* ============================================================
 * 四、现有编号机制：域与书签
 * ============================================================ */
function probeExistingNumbering(doc) {
  section("四、现有编号机制：域与书签");

  var total = 0;
  var cr = run(function () { return doc.Fields.Count; });
  if (!cr.ok) {
    log("FAIL", "无法读取 Fields.Count：" + cr.error);
    return;
  }
  total = Number(cr.value);

  var zEqn = 0;
  var seq = 0;
  var eqField = 0;
  var refField = 0;
  var zSamples = 0;

  var t0 = new Date().getTime();

  for (var i = 1; i <= total; i++) {
    if (i % PROGRESS_STEP === 0) log("INFO", "已扫描域 " + i + " / " + total);

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
        log("INFO", "ZEqnNum 域 #" + i + " 代码=" + esc(clip(code, 80)) + " 结果=" + esc(clip(res, 30)));
      }
    }
    if (/^\s*SEQ\b/i.test(code)) seq++;
    if (/^\s*EQ\b/i.test(code)) eqField++;
    if (/^\s*REF\b/i.test(code)) refField++;
  }

  var elapsed = new Date().getTime() - t0;

  log("INFO", "域总数=" + total);
  log("INFO", "含 ZEqnNum 的域：" + zEqn);
  log("INFO", "SEQ 域：" + seq);
  log("INFO", "EQ 域：" + eqField);
  log("INFO", "REF 域：" + refField);
  log("INFO", "域扫描耗时 " + elapsed + " ms");

  // ZEqnNum 书签
  var bmTotal = 0;
  var bcr = run(function () { return doc.Bookmarks.Count; });
  if (bcr.ok) bmTotal = Number(bcr.value);
  var zBookmarks = 0;

  for (var j = 1; j <= bmTotal; j++) {
    var bm = null;
    try { bm = doc.Bookmarks.Item(j); } catch (e4) { continue; }
    var name = "";
    try { name = String(bm.Name); } catch (e5) {}
    if (name.indexOf("ZEqnNum") === 0) {
      zBookmarks++;
      var btxt = "";
      var bpara = "";
      try { btxt = String(bm.Range.Text); } catch (e6) {}
      try { bpara = String(bm.Range.Paragraphs.Item(1).Range.Text); } catch (e7) {}
      log("INFO", "ZEqnNum 书签：" + name +
        " 内容=" + esc(clip(btxt, 40)) +
        " 所在段=" + esc(clip(bpara, 80)));
    }
  }
  log("INFO", "ZEqnNum 书签数量：" + zBookmarks);

  // MathType 自定义属性
  var props = null;
  var pr = run(function () { return doc.CustomDocumentProperties; });
  if (pr.ok) {
    props = pr.value;
    var pcount = 0;
    try { pcount = props.Count; } catch (e8) {}
    for (var m = 1; m <= pcount; m++) {
      try {
        var prop = props.Item(m);
        var pn = String(prop.Name);
        if (pn.indexOf("MT") === 0) {
          log("INFO", "MathType 属性：" + pn + " = " + esc(clip(String(prop.Value), 60)));
        }
      } catch (e9) {}
    }
  }
}

/* ============================================================
 * 五、MathType OLE 公式的编号位置
 * ============================================================ */
function probeMathTypeShapes(doc) {
  section("五、MathType OLE 公式的编号位置");

  var total = 0;
  var cr = run(function () { return doc.InlineShapes.Count; });
  if (!cr.ok) {
    log("FAIL", "无法读取 InlineShapes.Count：" + cr.error);
    return;
  }
  total = Number(cr.value);

  var found = 0;

  for (var i = 1; i <= total; i++) {
    var shape = null;
    try { shape = doc.InlineShapes.Item(i); } catch (e1) { continue; }

    var prog = "";
    try { prog = String(shape.OLEFormat.ProgID); } catch (e2) {}
    if (prog.indexOf("Equation") < 0) continue;

    found++;
    var p = null;
    try { p = shape.Range.Paragraphs.Item(1); } catch (e3) { continue; }

    var sn = paragraphStyleName(p);
    var inTable = "";
    var txt = "";
    var fc = -1;
    var codes = [];

    try { inTable = String(p.Range.Information(12)); } catch (e4) {}
    try { txt = String(p.Range.Text); } catch (e5) {}
    try { fc = p.Range.Fields.Count; } catch (e6) {}

    if (fc > 0) {
      for (var k = 1; k <= Math.min(fc, 5); k++) {
        try {
          codes.push(String(p.Range.Fields.Item(k).Code.Text));
        } catch (e7) {}
      }
    }

    log("INFO", "MathType 对象 #" + i + " ProgID=" + prog +
      " 样式=" + sn +
      " 在表格内=" + inTable +
      " 段内域数=" + fc);
    log("INFO", "    段文本=" + esc(clip(txt, 120)));
    if (codes.length > 0) {
      log("INFO", "    段内域代码=" + esc(clip(codes.join(" / "), 150)));
    }
  }

  log("INFO", "MathType 公式对象共 " + found + " 个");
}

/* ============================================================
 * 六、光标与选区中的公式
 * ============================================================ */
function probeSelectionFormulas(doc) {
  section("六、光标与选区中的公式");

  var sr = run(function () { return Application.Selection; });
  if (!sr.ok || !sr.value) {
    log("FAIL", "无法获取 Selection：" + (sr.error || "返回空"));
    return;
  }

  var sel = sr.value;
  var sStart = null;
  var sEnd = null;
  try { sStart = Number(sel.Range.Start); } catch (e1) {}
  try { sEnd = Number(sel.Range.End); } catch (e2) {}

  var selType = null;
  var stR = run(function () { return sel.Type; });
  if (stR.ok) selType = stR.value;
  log("INFO", "Selection.Type=" + toText(selType));
  log("INFO", "选区范围=" + sStart + "-" + sEnd + "，长度=" + (sEnd - sStart));

  var cnt = -1;
  var cr = run(function () { return sel.Range.OMaths.Count; });
  if (cr.ok) cnt = Number(cr.value);
  log("INFO", "选区 Range.OMaths.Count=" + cnt);

  if (cnt > 0) {
    for (var i = 1; i <= Math.min(cnt, 10); i++) {
      var om = null;
      try { om = sel.Range.OMaths.Item(i); } catch (e3) { continue; }
      var ty = "";
      var os = null;
      var oe = null;
      try { ty = String(om.Type); } catch (e4) {}
      try { os = Number(om.Range.Start); } catch (e5) {}
      try { oe = Number(om.Range.End); } catch (e6) {}

      var dist = "";
      if (sStart !== null && os !== null) {
        if (os <= sEnd && oe >= sStart) dist = "与选区重叠";
        else if (os > sEnd) dist = "在光标右侧 " + (os - sEnd) + " 字符";
        else dist = "在光标左侧 " + (sStart - oe) + " 字符";
      }

      log("INFO", "选区公式 #" + i + " Type=" + ty + " Range=" + os + "-" + oe + "，" + dist);

      try {
        var p = om.Range.Paragraphs.Item(1);
        log("INFO", "    段落样式=" + paragraphStyleName(p) +
          " 在表格内=" + String(p.Range.Information(12)) +
          " 段落文本=" + esc(clip(String(p.Range.Text), 100)));
      } catch (e7) {}
    }
  }

  // 段落级公式数量：判断“光标所在段落有几个公式”
  try {
    var para = sel.Paragraphs.Item(1);
    log("INFO", "光标所在段落样式=" + paragraphStyleName(para) +
      "，段落内 OMaths 数=" + para.Range.OMaths.Count +
      "，段落文本=" + esc(clip(String(para.Range.Text), 120)));
  } catch (e8) {
    log("FAIL", "读取光标所在段落失败：" + errText(e8));
  }

  // 选区内的嵌入对象
  var isc = -1;
  var ir = run(function () { return sel.Range.InlineShapes.Count; });
  if (ir.ok) isc = Number(ir.value);
  log("INFO", "选区 Range.InlineShapes.Count=" + isc);

  if (isc > 0) {
    for (var j = 1; j <= Math.min(isc, 10); j++) {
      try {
        var shape = sel.Range.InlineShapes.Item(j);
        var prog = "";
        try { prog = String(shape.OLEFormat.ProgID); } catch (e9) {}
        log("INFO", "选区对象 #" + j + " Type=" + String(shape.Type) + " ProgID=" + prog);
      } catch (e10) {}
    }
  }
}

/* ============================================================
 * 汇总与报告
 * ============================================================ */
function finish() {
  section("七、汇总");

  log("INFO", "PASS=" + passCount + "，FAIL=" + failCount + "，WARN=" + warnCount + "，INFO=" + infoCount);
  log("INFO", "说明：FAIL 仅表示该项 API 不可用或行为不符，不代表脚本执行失败。");

  var text = "EQNUM 第二轮探测报告\r\n时间：" + new Date().toString() + "\r\n" + reportLines.join("\r\n");
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
      alert("EQNUM 第二轮探测完成：PASS " + passCount + "，FAIL " + failCount + "，WARN " + warnCount +
        "。\r\n结果已输出到宏编辑器日志，并写入新建的探测报告文档。");
    }
  } catch (e3) {}
}

/* ============================================================
 * 主流程
 * ============================================================ */
function main() {
  section("零、脚本信息");
  log("INFO", "开始时间：" + new Date().toString());
  log("INFO", "只读探测：不修改任何文档。");

  var doc = findTargetDocument();
  if (!doc) {
    log("FAIL", "没有可用的文档，无法继续");
    finish();
    return;
  }

  probeTarget(doc);
  var stats = probeOMathTypes(doc);
  probeDisplaySamples(doc, stats);
  probeExistingNumbering(doc);
  probeMathTypeShapes(doc);
  probeSelectionFormulas(doc);
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
