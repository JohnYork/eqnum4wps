/*
 * WPS 题注括号测试
 *
 * 目的：确认原生题注交叉引用（REF _Ref...）是否包含题注外侧的字面括号。
 *
 * 构造 4 种题注：
 *   1. 手工 ( SEQ 公式 )                    —— 纯序号
 *   2. 手工 ( STYLEREF 1 \s . SEQ 公式 )    —— 实际章号结构
 *   3. InsertCaption 后手工补括号
 *   4. 表格单元格内手工 ( SEQ 公式 )
 *
 * 然后分别插入交叉引用，检查 REF 结果是否以括号包裹。
 *
 * 安全：只写临时文档，正式文档不触碰。
 */

function EQNUM_CAPTION_PAREN_TEST() {
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

function countFields(doc) {
  try { return Number(doc.Fields.Count); } catch (e) { return -1; }
}

function addParagraphAtEnd(doc) {
  var r = run(function () { return doc.Paragraphs.Add(); });
  if (r.ok && r.value) return r.value;
  var r2 = run(function () {
    doc.Content.InsertAfter("\r");
    return doc.Paragraphs.Item(doc.Paragraphs.Count);
  });
  return r2.ok ? r2.value : null;
}

function lastFieldParagraph(doc) {
  var total = 0;
  try { total = Number(doc.Paragraphs.Count); } catch (e) { return null; }
  for (var i = total; i >= 1; i--) {
    var p = null;
    try { p = doc.Paragraphs.Item(i); } catch (e1) { continue; }
    var n = -1;
    try { n = Number(p.Range.Fields.Count); } catch (e2) {}
    if (n > 0) return p;
  }
  return null;
}

function makeManualCaptionInParagraph(doc, p, withStyleref) {
  if (!p) return { ok: false, error: "段落为空" };

  var start = -1;
  try { start = Number(p.Range.Start); } catch (e0) { return { ok: false, error: errText(e0) }; }

  var r0 = run(function () { doc.Range(start, start).InsertAfter("("); return true; });
  if (!r0.ok) return { ok: false, error: r0.error };

  if (withStyleref) {
    var r1 = run(function () {
      var end = Number(p.Range.End) - 1;
      var f = doc.Fields.Add(doc.Range(end, end), -1, ' STYLEREF 1 \\s \\* ARABIC ', false);
      f.Update();
      return true;
    });
    if (!r1.ok) return { ok: false, error: r1.error };

    var rd = run(function () {
      var end = Number(p.Range.End) - 1;
      doc.Range(end, end).InsertAfter(".");
      return true;
    });
    if (!rd.ok) return { ok: false, error: rd.error };
  }

  var r2 = run(function () {
    var end = Number(p.Range.End) - 1;
    var f = doc.Fields.Add(doc.Range(end, end), -1, " SEQ 公式 \\* ARABIC \\s 1 ", false);
    f.Update();
    return true;
  });
  if (!r2.ok) return { ok: false, error: r2.error };

  var r3 = run(function () {
    var end = Number(p.Range.End) - 1;
    doc.Range(end, end).InsertAfter(")");
    return true;
  });
  if (!r3.ok) return { ok: false, error: r3.error };

  run(function () {
    try { p.Style = "题注"; } catch (e1) {
      try { p.Style = "Caption"; } catch (e2) {}
    }
    return true;
  });

  return { ok: true };
}

function makeManualCaption(doc, withStyleref) {
  var p = addParagraphAtEnd(doc);
  if (!p) return { ok: false, error: "无法新建段落" };
  return makeManualCaptionInParagraph(doc, p, withStyleref);
}

function describeParagraph(p, label) {
  if (!p) {
    log("FAIL", label + " 段落为空");
    return;
  }
  var text = "";
  try { text = String(p.Range.Text); } catch (e1) {}
  var fs = fieldInfo(p.Range, 8);
  log("INFO", label + " 样式=" + styleName(p) +
    " 文本=" + esc(clip(text, 100)) + " 域数=" + fs.length);
  for (var i = 0; i < fs.length; i++) {
    log("INFO", "  " + label + " 域#" + (i + 1) +
      " 代码=" + esc(clip(fs[i].code, 100)) +
      " 结果=" + esc(clip(fs[i].result, 50)));
  }
}

function hasParens(s) {
  s = String(s);
  var open = (s.indexOf("(") >= 0 || s.indexOf("（") >= 0);
  var close = (s.indexOf(")") >= 0 || s.indexOf("）") >= 0);
  return open && close;
}

function insertRefAndFind(doc, kind, item) {
  var p = addParagraphAtEnd(doc);
  if (p) run(function () { p.Range.Select(); return true; });

  var before = countFields(doc);

  var r = run(function () {
    Application.Selection.InsertCrossReference("公式", kind, item, true, false, false, "");
    return true;
  });
  if (!r.ok) {
    r = run(function () {
      Application.Selection.InsertCrossReference("公式", kind, item, true);
      return true;
    });
  }
  if (!r.ok) {
    r = run(function () {
      Application.Selection.InsertCrossReference("公式", kind, item);
      return true;
    });
  }

  if (!r.ok) return { ok: false, error: r.error };

  var after = countFields(doc);
  var candidates = [];

  for (var i = before + 1; i <= after; i++) {
    var f = null;
    try { f = doc.Fields.Item(i); } catch (e1) { continue; }
    var code = "", res = "";
    try { code = String(f.Code.Text); } catch (e2) {}
    try { res = String(f.Result.Text); } catch (e3) {}
    candidates.push({ code: code, result: res });
  }

  if (candidates.length === 0 && p) {
    candidates = fieldInfo(p.Range, 6);
  }

  var ref = null;
  for (i = 0; i < candidates.length; i++) {
    if (candidates[i].code.indexOf("REF") >= 0) {
      ref = candidates[i];
      break;
    }
  }

  return { ok: true, before: before, after: after, fields: candidates, ref: ref };
}

/* ============================================================
 * 主流程
 * ============================================================ */
function main() {
  section("零、脚本信息");
  log("INFO", "开始时间：" + new Date().toString());
  log("INFO", "只写临时文档。");

  var doc = addDocument();
  if (!doc) {
    log("FAIL", "无法创建临时文档");
    writeReport();
    return;
  }
  log("INFO", "临时文档=" + doc.Name);
  run(function () { doc.Activate(); return true; });

  section("一、构造 4 种题注");

  var c1 = makeManualCaption(doc, false);
  log(c1.ok ? "PASS" : "FAIL", "题注1 手工 (SEQ)：" + (c1.ok ? "成功" : c1.error));
  describeParagraph(lastFieldParagraph(doc), "题注1");

  var c2 = makeManualCaption(doc, true);
  log(c2.ok ? "PASS" : "FAIL", "题注2 手工 (STYLEREF . SEQ)：" + (c2.ok ? "成功" : c2.error));
  describeParagraph(lastFieldParagraph(doc), "题注2");

  // 题注3：InsertCaption 后补括号
  var p3 = addParagraphAtEnd(doc);
  var c3 = { ok: false };
  if (p3) {
    run(function () { p3.Range.Select(); return true; });
    c3 = run(function () { Application.Selection.InsertCaption("公式", "", "", 0, true); return true; });
  }
  log(c3.ok ? "PASS" : "FAIL", "题注3 InsertCaption：" + (c3.ok ? "成功" : c3.error));

  var p3cap = lastFieldParagraph(doc);
  if (p3cap) {
    var s3 = -1;
    try { s3 = Number(p3cap.Range.Start); } catch (e0) {}
    run(function () { if (s3 >= 0) doc.Range(s3, s3).InsertAfter("("); return true; });
    run(function () {
      var end = Number(p3cap.Range.End) - 1;
      doc.Range(end, end).InsertAfter(")");
      return true;
    });
  }
  describeParagraph(p3cap, "题注3");

  // 题注4：表格单元格内手工 (SEQ)
  var tbl = null;
  var tr = run(function () {
    return doc.Tables.Add(doc.Range(doc.Content.End - 1, doc.Content.End - 1), 1, 2);
  });
  if (tr.ok && tr.value) {
    tbl = tr.value;
    run(function () { tbl.Borders.Enable = 0; return true; });
    run(function () { tbl.Cell(1, 1).Range.Text = "公式"; return true; });
    run(function () { tbl.Cell(1, 2).Range.Text = ""; return true; });

    var cp = run(function () { return tbl.Cell(1, 2).Range.Paragraphs.Item(1); });
    if (cp.ok && cp.value) {
      var c4 = makeManualCaptionInParagraph(doc, cp.value, false);
      log(c4.ok ? "PASS" : "FAIL", "题注4 表格单元格 (SEQ)：" + (c4.ok ? "成功" : c4.error));
      describeParagraph(cp.value, "题注4");
    } else {
      log("FAIL", "无法读取表格单元格段落");
    }
  } else {
    log("FAIL", "创建表格失败");
  }

  section("二、交叉引用结果是否带括号");

  for (var item = 1; item <= 4; item++) {
    var xr = insertRefAndFind(doc, 1, item);
    if (!xr.ok) {
      log("FAIL", "引用第 " + item + " 项失败：" + xr.error);
      continue;
    }
    if (!xr.ref) {
      log("FAIL", "引用第 " + item + " 项未找到 REF 域，新增域数=" +
        (xr.after - xr.before));
      for (var fi = 0; fi < xr.fields.length; fi++) {
        log("INFO", "  新增域#" + (fi + 1) +
          " 代码=" + esc(clip(xr.fields[fi].code, 80)) +
          " 结果=" + esc(clip(xr.fields[fi].result, 60)));
      }
      continue;
    }

    log(hasParens(xr.ref.result) ? "PASS" : "FAIL",
      "第 " + item + " 项 REF 结果=" + esc(clip(xr.ref.result, 80)) +
      "，是否带括号=" + (hasParens(xr.ref.result) ? "是" : "否") +
      "，代码=" + esc(clip(xr.ref.code, 60)));
  }

  // 仅题注文字
  var xrText = insertRefAndFind(doc, 3, 1);
  if (xrText.ok && xrText.ref) {
    log(hasParens(xrText.ref.result) ? "PASS" : "FAIL",
      "仅题注文字 结果=" + esc(clip(xrText.ref.result, 80)) +
      "，是否带括号=" + (hasParens(xrText.ref.result) ? "是" : "否"));
  } else {
    log("FAIL", "仅题注文字引用失败：" + (xrText.error || "未找到 REF"));
  }

  section("三、汇总");
  log("INFO", "以上 4 项分别对应：");
  log("INFO", "  1=手工 (SEQ)，2=手工 (STYLEREF . SEQ)，3=InsertCaption 后补括号，4=表格单元格 (SEQ)");
  log("INFO", "如果某项 PASS，说明该写法下原生题注引用会带外层括号。");

  writeReport();

  var close = run(function () { doc.Close(0); return true; });
  if (close.ok) log("INFO", "临时文档已关闭（未保存）");
  else log("WARN", "关闭临时文档失败：" + close.error);

  try {
    if (typeof alert === "function") alert("题注括号测试完成，详见报告文档。");
  } catch (e9) {}
}

function writeReport() {
  var text = "EQNUM 题注括号测试报告\r时间：" + new Date().toString() + "\r" + reportLines.join("\r");
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
