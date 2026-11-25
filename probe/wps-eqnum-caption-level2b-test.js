/*
 * WPS 题注二级章与 SEQ 切换补充探测
 *
 * 1. 在正式文档副本的两个标题 2 后，分别测试：
 *      A: STYLEREF 2 \n
 *      B: STYLEREF 2 \w
 *      C: STYLEREF "标题 2" \n
 *      D: STYLEREF 1 \s + "." + STYLEREF 2 \s
 *      E: STYLEREF 2 \s（参考）
 *    比较结果与标题的 ListString（如 1.1、1.2）。
 *
 * 2. 用全局字段索引重测：SEQ EQNUM 两次、SEQ 公式 两次，
 *    确认切换标识后是否从 1 重新开始。
 *
 * 安全：正式文档只读；副本测完关闭不保存。
 */

function EQNUM_CAPTION_LEVEL2B_TEST() {
var reportLines = [];

function consoleLog(s) {
  try {
    if (typeof console !== "undefined" && console && console.log) console.log(s);
  } catch (e) {}
}

function log(level, msg) {
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
  try { return { ok: true, value: fn() }; }
  catch (e) { return { ok: false, error: errText(e) }; }
}

function clip(s, n) {
  s = String(s);
  if (s.length <= n) return s;
  return s.substring(0, n) + "...(共" + s.length + "字符)";
}

function esc(s) {
  s = String(s);
  return s.replace(/\r/g, "\\r").replace(/\n/g, "\\n").replace(/\t/g, "\\t");
}

function norm(s) {
  return String(s).replace(/\s/g, "");
}

function addDocument() {
  var before = -1;
  try { before = Application.Documents.Count; } catch (e) {}
  var r = run(function () { return Application.Documents.Add(); });
  if (r.ok && r.value) return r.value;
  var after = -1;
  try { after = Application.Documents.Count; } catch (e2) {}
  if (before >= 0 && after === before + 1) {
    try { return Application.ActiveDocument; } catch (e3) {}
  }
  return null;
}

function styleName(p) {
  try { return String(p.Style.NameLocal); } catch (e1) {}
  try { return String(p.Style); } catch (e2) {}
  return "读取失败";
}

function listString(p) {
  try { return String(p.Range.ListFormat.ListString); } catch (e) { return ""; }
}

function fieldInfo(r, limit) {
  var out = [];
  try {
    var fs = r.Fields;
    var n = Number(fs.Count);
    for (var i = 1; i <= n && i <= limit; i++) {
      var f = null;
      try { f = fs.Item(i); } catch (e1) { continue; }
      var code = "", res = "";
      try { code = String(f.Code.Text); } catch (e2) {}
      try { res = String(f.Result.Text); } catch (e3) {}
      out.push({ code: code, result: res });
    }
  } catch (e4) {}
  return out;
}

function findTargetDoc() {
  var best = null;
  var bestCount = -1;
  var total = 0;
  try { total = Application.Documents.Count; } catch (e) {}
  for (var i = 1; i <= total; i++) {
    var d = null;
    try { d = Application.Documents.Item(i); } catch (e1) { continue; }
    var c = -1;
    try { c = Number(d.OMaths.Count); } catch (e2) {}
    if (c > bestCount) { bestCount = c; best = d; }
  }
  return best;
}

function findHeading2s(doc, maxScan, wanted) {
  var total = 0;
  try { total = Number(doc.Paragraphs.Count); } catch (e) { return []; }
  var limit = Math.min(total, maxScan || 3000);
  var out = [];
  for (var i = 1; i <= limit && out.length < wanted; i++) {
    var p = null;
    try { p = doc.Paragraphs.Item(i); } catch (e1) { continue; }
    if (/^(标题|Heading)\s*2$/i.test(styleName(p))) {
      out.push({ p: p, text: listString(p), para: i });
    }
  }
  return out;
}

function fieldAtEnd(doc, p, code) {
  var end = -1;
  try { end = Number(p.Range.End) - 1; } catch (e0) { return { ok: false, error: errText(e0) }; }
  return run(function () {
    var f = doc.Fields.Add(doc.Range(end, end), -1, code, false);
    f.Update();
    return true;
  });
}

function textAtEnd(doc, p, text) {
  var end = -1;
  try { end = Number(p.Range.End) - 1; } catch (e0) { return { ok: false, error: errText(e0) }; }
  return run(function () {
    doc.Range(end, end).InsertAfter(text);
    return true;
  });
}

function insertVariantAfter(doc, prev, label, parts) {
  var idx = 0;
  try { idx = Number(doc.Range(0, prev.Range.Start).Paragraphs.Count) + 1; } catch (e0) {}

  var ins = run(function () { prev.Range.InsertParagraphAfter(); return true; });
  if (!ins.ok) return { ok: false, error: ins.error };

  var np = null;
  try { np = prev.Next(); } catch (e1) {}
  if (!np) {
    try { np = doc.Paragraphs.Item(idx + 1); } catch (e2) {}
  }
  if (!np) return { ok: false, error: "无法读取新段落" };

  var t0 = textAtEnd(doc, np, label);
  if (!t0.ok) return { ok: false, error: t0.error };

  for (var i = 0; i < parts.length; i++) {
    if (parts[i].type === "text") {
      var tr = textAtEnd(doc, np, parts[i].text);
      if (!tr.ok) return { ok: false, error: tr.error };
    } else {
      var fr = fieldAtEnd(doc, np, parts[i].code);
      if (!fr.ok) return { ok: false, error: fr.error };
    }
  }

  return { ok: true, para: np };
}

function logVariant(v, headingText) {
  if (!v.ok) {
    log("FAIL", "  变体插入失败：" + v.error);
    return;
  }

  var fs = fieldInfo(v.para.Range, 6);
  var text = "";
  try { text = String(v.para.Range.Text); } catch (e1) {}
  log("INFO", "  文本=" + esc(clip(text, 80)) + " 域数=" + fs.length);

  var values = [];
  for (var i = 0; i < fs.length; i++) {
    values.push(fs[i].result);
    log("INFO", "    域#" + (i + 1) + " 代码=" + esc(clip(fs[i].code, 80)) +
      " 结果=" + esc(clip(fs[i].result, 40)));
  }

  var combined = values.join(".");
  var ok = (norm(combined) === norm(headingText));
  log(ok ? "PASS" : "WARN",
    "  结果是否等于标题编号 " + esc(headingText) + "：" +
    (ok ? "是" : "否") + "，实际=" + esc(clip(combined, 40)));
}

/* ============================================================
 * 主流程
 * ============================================================ */
function main() {
  section("零、脚本信息");
  log("INFO", "开始时间：" + new Date().toString());
  log("INFO", "正式文档只读；副本测完关闭不保存。");

  var target = findTargetDoc();
  if (target) log("INFO", "目标文档=" + target.Name);

  section("一、SEQ 标识切换（全局字段索引）");
  var tmp = addDocument();
  if (!tmp) {
    log("FAIL", "无法创建临时文档");
  } else {
    run(function () { tmp.Activate(); return true; });

    var seqCases = ["EQNUM", "EQNUM", "公式", "公式"];
    var seqResults = [];

    for (var si = 0; si < seqCases.length; si++) {
      tmp.Content.InsertAfter("\r");
      var end = Number(tmp.Content.End) - 1;
      if (end < 0) end = 0;

      var before = -1;
      try { before = Number(tmp.Fields.Count); } catch (e1) {}

      var fr = run(function () {
        var f = tmp.Fields.Add(tmp.Range(end, end), -1, " SEQ " + seqCases[si] + " \\* ARABIC ", false);
        f.Update();
        return true;
      });

      if (!fr.ok) {
        log("FAIL", "插入 SEQ " + seqCases[si] + " 失败：" + fr.error);
        seqResults.push("");
        continue;
      }

      var after = -1;
      try { after = Number(tmp.Fields.Count); } catch (e2) {}

      var field = null;
      try { field = tmp.Fields.Item(after); } catch (e3) {}
      if (!field) {
        try { field = tmp.Fields.Item(before + 1); } catch (e4) {}
      }

      var code = "", res = "";
      try { code = String(field.Code.Text); } catch (e5) {}
      try { res = String(field.Result.Text); } catch (e6) {}

      seqResults.push(res);
      log("INFO", "SEQ " + seqCases[si] + "：域数 " + before + "→" + after +
        "，代码=" + esc(clip(code, 60)) + "，结果=" + esc(res));
    }

    var switchOk = (seqResults.length === 4 &&
      seqResults[0] === "1" && seqResults[1] === "2" &&
      seqResults[2] === "1" && seqResults[3] === "2");
    log(switchOk ? "PASS" : "FAIL",
      "SEQ 公式 是否从 1 重新开始：" + (switchOk ? "是" : "否") +
      "，实测=" + seqResults.join(" / "));

    var closeTmp = run(function () { tmp.Close(0); return true; });
    if (closeTmp.ok) log("INFO", "临时文档已关闭（未保存）");
    else log("WARN", "关闭临时文档失败：" + closeTmp.error);
  }

  section("二、标题 2 的 STYLEREF 变体");
  if (!target) {
    log("FAIL", "没有目标文档");
    writeReport();
    return;
  }

  var copy = null;
  var cr = run(function () { return Application.Documents.Add(target.FullName); });
  if (cr.ok && cr.value) {
    var same = false;
    try { same = (String(cr.value.FullName) === String(target.FullName)); } catch (e7) {}
    if (!same) {
      copy = cr.value;
      log("INFO", "已创建内存副本：" + copy.Name);
    } else {
      run(function () { cr.value.Close(0); return true; });
      log("FAIL", "Documents.Add 返回了正式文档本身");
    }
  } else {
    log("FAIL", "无法创建内存副本：" + (cr.error || "返回空"));
  }

  if (copy) {
    run(function () { copy.Activate(); return true; });

    var headings = findHeading2s(copy, 3000, 2);
    log("INFO", "找到标题 2 数量=" + headings.length +
      (headings.length ? "，编号=" + headings[0].text +
        (headings[1] ? " / " + headings[1].text : "") : ""));

    if (headings.length === 0) {
      log("FAIL", "副本中没有标题 2");
    } else {
      for (var hi = 0; hi < headings.length; hi++) {
        var h = headings[hi];
        log("INFO", "--- 标题 2 #" + (hi + 1) + "，编号=" + esc(h.text) + " ---");

        var prev = h.p;

        var va = insertVariantAfter(copy, prev, "A: ", [
          { type: "field", code: ' STYLEREF 2 \\n ' }
        ]);
        log(va.ok ? "PASS" : "FAIL", "A 插入：" + (va.ok ? "成功" : va.error));
        if (va.ok) { logVariant(va, h.text); prev = va.para; }

        var vb = insertVariantAfter(copy, prev, "B: ", [
          { type: "field", code: ' STYLEREF 2 \\w ' }
        ]);
        log(vb.ok ? "PASS" : "FAIL", "B 插入：" + (vb.ok ? "成功" : vb.error));
        if (vb.ok) { logVariant(vb, h.text); prev = vb.para; }

        var vc = insertVariantAfter(copy, prev, "C: ", [
          { type: "field", code: ' STYLEREF "标题 2" \\n ' }
        ]);
        log(vc.ok ? "PASS" : "FAIL", "C 插入：" + (vc.ok ? "成功" : vc.error));
        if (vc.ok) { logVariant(vc, h.text); prev = vc.para; }

        var vd = insertVariantAfter(copy, prev, "D: ", [
          { type: "field", code: ' STYLEREF 1 \\s ' },
          { type: "text", text: "." },
          { type: "field", code: ' STYLEREF 2 \\s ' }
        ]);
        log(vd.ok ? "PASS" : "FAIL", "D 插入：" + (vd.ok ? "成功" : vd.error));
        if (vd.ok) { logVariant(vd, h.text); prev = vd.para; }

        var ve = insertVariantAfter(copy, prev, "E: ", [
          { type: "field", code: ' STYLEREF 2 \\s ' }
        ]);
        log(ve.ok ? "PASS" : "FAIL", "E 插入：" + (ve.ok ? "成功" : ve.error));
        if (ve.ok) { logVariant(ve, h.text); prev = ve.para; }
      }
    }

    var cl = run(function () { copy.Close(0); return true; });
    if (cl.ok) log("INFO", "内存副本已关闭（未保存）");
    else log("WARN", "关闭内存副本失败：" + cl.error);
  }

  section("三、汇总");
  log("INFO", "A/B/C/D/E 中，结果等于标题编号的写法可用于二级章题注。");
  log("INFO", "如果 A 或 C 等于 1.1/1.2，优先使用 \\n；如果 B 可用，\\w 也可以。");

  writeReport();

  try {
    if (typeof alert === "function") alert("题注二级章补充测试完成，详见报告文档。");
  } catch (e9) {}
}

function writeReport() {
  var text = "EQNUM 题注二级章补充报告\r时间：" + new Date().toString() + "\r" + reportLines.join("\r");
  var rd = addDocument();
  if (rd) {
    try {
      rd.Content.Text = text;
      try { rd.Content.Font.Name = "Consolas"; } catch (e1) {}
      try { rd.Content.Font.Size = 10; } catch (e2) {}
    } catch (e3) {}
  }
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
