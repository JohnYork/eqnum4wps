/*
 * WPS 题注兼容性测试 3
 *
 * 目标：验证“两栏表格右单元格 + 题注字段”的可行实现
 *
 * 测试：
 *   1. 两个题注分别放在独立段落，交叉引用是否各指各的
 *   2. 在表格单元格内手工插入 STYLEREF 1 \s + "." + SEQ 公式 \s 1
 *   3. 该单元格题注能否被 InsertCrossReference 引用
 *   4. 在正式文档副本中验证 STYLEREF 1 \s 和 STYLEREF "标题 1" \n 的章号
 *
 * 安全：临时文档和内存副本，正式文档只读。
 */

function EQNUM_CAPTION_TEST3() {
var reportLines = [];
var passCount = 0;
var failCount = 0;

function consoleLog(s) {
  try {
    if (typeof console !== "undefined" && console && console.log) console.log(s);
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
  s = s.replace(/\r/g, "\\r").replace(/\n/g, "\\n").replace(/\t/g, "\\t");
  return s;
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

function moveToEnd(doc) {
  run(function () {
    var end = Number(doc.Content.End) - 1;
    if (end < 0) end = 0;
    doc.Range(end, end).Select();
    return true;
  });
}

function addParagraphAtEnd(doc) {
  var r = run(function () { return doc.Paragraphs.Add(); });
  if (r.ok && r.value) return r.value;

  var r2 = run(function () {
    doc.Content.InsertAfter("\r");
    var end = Number(doc.Content.End) - 1;
    if (end < 0) end = 0;
    doc.Range(end, end).Select();
    return doc.Paragraphs.Item(doc.Paragraphs.Count);
  });
  return r2.ok ? r2.value : null;
}

function logFields(prefix, fields) {
  for (var i = 0; i < fields.length; i++) {
    log("INFO", prefix + " 域#" + (i + 1) +
      " 代码=" + esc(clip(fields[i].code, 120)) +
      " 结果=" + esc(clip(fields[i].result, 60)));
  }
}

function lastParagraphInfo(doc, label) {
  var p = null;
  try { p = doc.Paragraphs.Item(doc.Paragraphs.Count); } catch (e1) { return null; }
  var text = "";
  try { text = String(p.Range.Text); } catch (e2) {}
  var fields = fieldInfo(p.Range, 10);
  log("INFO", label + " 样式=" + styleName(p) +
    " 文本=" + esc(clip(text, 100)) + " 域数=" + fields.length);
  logFields(label, fields);
  return { p: p, text: text, fields: fields };
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

function findFirstHeading(d, maxScan) {
  var total = 0;
  try { total = Number(d.Paragraphs.Count); } catch (e) { return null; }
  var limit = Math.min(total, maxScan || 1500);
  for (var i = 1; i <= limit; i++) {
    var p = null;
    try { p = d.Paragraphs.Item(i); } catch (e1) { continue; }
    if (/^(标题|Heading)\s*1$/i.test(styleName(p))) return p;
  }
  return null;
}

function insertCrossRefAtNewParagraph(doc, label, kind, item) {
  var p = addParagraphAtEnd(doc);
  if (!p) return { ok: false, error: "无法新建段落" };

  var target = p.Range;
  run(function () { target.Select(); return true; });

  var before = countFields(doc);
  var sigs = [
    function () { Application.Selection.InsertCrossReference(label, kind, item, true, false, false, ""); },
    function () { Application.Selection.InsertCrossReference(label, kind, item, true); },
    function () { Application.Selection.InsertCrossReference(label, kind, item); }
  ];

  var lastErr = "";
  for (var i = 0; i < sigs.length; i++) {
    var r = run(sigs[i]);
    if (r.ok) {
      var after = countFields(doc);
      var lp = doc.Paragraphs.Item(doc.Paragraphs.Count);
      var fields = fieldInfo(lp.Range, 10);
      return {
        ok: true,
        before: before,
        after: after,
        fields: fields,
        text: String(lp.Range.Text),
        signature: i + 1
      };
    }
    lastErr = r.error;
  }
  return { ok: false, error: lastErr || "未生成引用" };
}

function insertCaptionFieldsInParagraph(doc, p) {
  var start = -1;
  try { start = Number(p.Range.Start); } catch (e0) { return { ok: false, error: e0 }; }

  var f1 = run(function () {
    var f = doc.Fields.Add(doc.Range(start, start), -1, ' STYLEREF 1 \\s \\* ARABIC ', false);
    f.Update();
    return true;
  });
  if (!f1.ok) return { ok: false, error: f1.error };

  var dot = run(function () {
    var end = Number(p.Range.End) - 1;
    doc.Range(end, end).InsertAfter(".");
    return true;
  });
  if (!dot.ok) return { ok: false, error: dot.error };

  var f2 = run(function () {
    var end = Number(p.Range.End) - 1;
    var f = doc.Fields.Add(doc.Range(end, end), -1, " SEQ 公式 \\* ARABIC \\s 1 ", false);
    f.Update();
    return true;
  });
  if (!f2.ok) return { ok: false, error: f2.error };

  run(function () {
    try { p.Style = "题注"; } catch (e1) {
      try { p.Style = "Caption"; } catch (e2) {}
    }
    return true;
  });

  return { ok: true };
}

/* ============================================================
 * 主流程
 * ============================================================ */
function main() {
  section("零、脚本信息");
  log("INFO", "开始时间：" + new Date().toString());
  log("INFO", "正式文档只读；测试在临时文档或内存副本中进行。");

  var target = findTargetDoc();
  if (target) log("INFO", "目标文档=" + target.Name);

  var testDoc = addDocument();
  if (!testDoc) {
    log("FAIL", "无法创建临时文档");
    writeReport();
    return;
  }
  log("INFO", "临时文档=" + testDoc.Name);
  run(function () { testDoc.Activate(); return true; });

  section("一、两个独立段落的题注");
  for (var ci = 1; ci <= 2; ci++) {
    var p = addParagraphAtEnd(testDoc);
    if (!p) {
      log("FAIL", "无法为第 " + ci + " 个题注新建段落");
      break;
    }
    run(function () { p.Range.Select(); return true; });

    var ic = run(function () { Application.Selection.InsertCaption("公式", "", "", 0, true); return true; });
    log(ic.ok ? "PASS" : "FAIL", "第 " + ci + " 个题注插入：" + (ic.ok ? "成功" : ic.error));
    if (ic.ok) lastParagraphInfo(testDoc, "题注" + ci);
  }

  section("二、交叉引用是否各指各的");
  var cases = [
    { kind: 1, item: 1 },
    { kind: 1, item: 2 },
    { kind: 2, item: 1 },
    { kind: 3, item: 1 }
  ];
  for (var xi = 0; xi < cases.length; xi++) {
    var c = cases[xi];
    var xr = insertCrossRefAtNewParagraph(testDoc, "公式", c.kind, c.item);

    if (!xr.ok) {
      log("FAIL", "类型" + c.kind + " 第" + c.item + "项失败：" + xr.error);
      continue;
    }

    var hasRef = false, okRes = false;
    for (var fi = 0; fi < xr.fields.length; fi++) {
      if (xr.fields[fi].code.indexOf("REF") >= 0) hasRef = true;
      if (xr.fields[fi].result && xr.fields[fi].result.indexOf("错误") < 0) okRes = true;
    }

    log(hasRef && okRes ? "PASS" : "FAIL",
      "类型" + c.kind + " 第" + c.item + "项：域数 " + xr.before + "→" + xr.after +
      "，文本=" + esc(clip(xr.text, 80)));
    logFields("类型" + c.kind + "_第" + c.item + "项", xr.fields);
  }

  section("三、表格单元格内手工题注字段");
  var tbl = null;
  var tr = run(function () {
    return testDoc.Tables.Add(testDoc.Range(testDoc.Content.End - 1, testDoc.Content.End - 1), 1, 2);
  });
  if (tr.ok && tr.value) {
    tbl = tr.value;
    run(function () { tbl.Borders.Enable = 0; return true; });
    run(function () { tbl.Cell(1, 1).Range.Text = "公式"; return true; });
    run(function () { tbl.Cell(1, 2).Range.Text = ""; return true; });

    var cellP = null;
    var cp = run(function () { return tbl.Cell(1, 2).Range.Paragraphs.Item(1); });
    if (cp.ok && cp.value) cellP = cp.value;

    if (cellP) {
      var ins = insertCaptionFieldsInParagraph(testDoc, cellP);
      log(ins.ok ? "PASS" : "FAIL", "单元格内插入题注字段：" + (ins.ok ? "成功" : ins.error));

      var ct = "";
      try { ct = String(tbl.Cell(1, 2).Range.Text); } catch (e1) {}
      var cf = [];
      try { cf = fieldInfo(tbl.Cell(1, 2).Range, 10); } catch (e2) {}
      log("INFO", "单元格文本=" + esc(clip(ct, 100)) + " 域数=" + cf.length);
      logFields("单元格", cf);

      var xr3 = insertCrossRefAtNewParagraph(testDoc, "公式", 1, 3);
      if (xr3.ok) {
        log("INFO", "引用表格题注：域数 " + xr3.before + "→" + xr3.after +
          "，文本=" + esc(clip(xr3.text, 80)));
        logFields("表格引用", xr3.fields);
      } else {
        log("FAIL", "引用表格题注失败：" + xr3.error);
      }
    } else {
      log("FAIL", "无法读取表格单元格段落");
    }
  } else {
    log("FAIL", "创建表格失败：" + (tr.error || "返回空"));
  }

  section("四、正式文档副本中的章号");
  if (target) {
    var copy = null;
    var cr = run(function () { return Application.Documents.Add(target.FullName); });
    if (cr.ok && cr.value) {
      var same = false;
      try { same = (String(cr.value.FullName) === String(target.FullName)); } catch (e1) {}
      if (!same) {
        copy = cr.value;
        log("INFO", "已创建内存副本：" + copy.Name);
      } else {
        run(function () { cr.value.Close(0); return true; });
        log("WARN", "Documents.Add 返回了正式文档本身，已关闭");
      }
    } else {
      log("WARN", "无法创建内存副本：" + (cr.error || "返回空"));
    }

    if (copy) {
      run(function () { copy.Activate(); return true; });

      var heading = findFirstHeading(copy, 2000);
      if (!heading) {
        log("FAIL", "副本中没有找到标题 1");
      } else {
        var hIdx = 0;
        try { hIdx = Number(copy.Range(0, heading.Range.Start).Paragraphs.Count) + 1; } catch (e2) {}

        // 在标题后插入新段落
        run(function () { heading.Range.InsertParagraphAfter(); return true; });

        var np = null;
        try { np = copy.Paragraphs.Item(hIdx + 1); } catch (e3) {}
        if (!np) {
          log("FAIL", "标题后新段落读取失败");
        } else {
          var ins2 = insertCaptionFieldsInParagraph(copy, np);
          log(ins2.ok ? "PASS" : "FAIL", "副本题注字段插入：" + (ins2.ok ? "成功" : ins2.error));

          var pt = "";
          try { pt = String(np.Range.Text); } catch (e4) {}
          var pf = fieldInfo(np.Range, 10);
          log("INFO", "副本题注文本=" + esc(clip(pt, 100)) + " 域数=" + pf.length);
          logFields("副本 STYLEREF 1", pf);

          var chapterOk = false;
          for (var pi = 0; pi < pf.length; pi++) {
            if (pf[pi].code.indexOf("STYLEREF") >= 0 &&
                pf[pi].result &&
                pf[pi].result.indexOf("错误") < 0 &&
                pf[pi].result.indexOf("未找到") < 0) {
              chapterOk = true;
            }
          }
          log(chapterOk ? "PASS" : "FAIL", "STYLEREF 1 \\s 是否解析出章号：" + (chapterOk ? "是" : "否"));
        }

        // 再测 STYLEREF "标题 1" \n
        var np2 = addParagraphAtEnd(copy);
        if (np2) {
          var start2 = Number(np2.Range.Start);
          var f3 = run(function () {
            var f = copy.Fields.Add(copy.Range(start2, start2), -1, ' STYLEREF "标题 1" \\n ', false);
            f.Update();
            return true;
          });
          if (f3.ok) {
            var pf2 = fieldInfo(np2.Range, 4);
            log("INFO", "STYLEREF \"标题 1\" \\n 结果：" +
              esc(clip(pf2.length ? pf2[0].result : "(空)", 60)));
            log(pf2.length && pf2[0].result && pf2[0].result.indexOf("错误") < 0
              ? "PASS" : "FAIL", "STYLEREF \"标题 1\" \\n 是否可用");
          } else {
            log("FAIL", "STYLEREF \"标题 1\" 插入失败：" + f3.error);
          }
        }
      }

      var closeCopy = run(function () { copy.Close(0); return true; });
      if (closeCopy.ok) log("INFO", "内存副本已关闭（未保存）");
      else log("WARN", "关闭内存副本失败：" + closeCopy.error);
    }
  } else {
    log("WARN", "没有找到目标文档，跳过副本测试");
  }

  section("五、汇总");
  log("INFO", "PASS=" + passCount + "，FAIL=" + failCount);
  log("INFO", "重点：独立题注是否各指各的、单元格题注能否引用、副本章号是否解析。");

  writeReport();

  var close = run(function () { testDoc.Close(0); return true; });
  if (close.ok) log("INFO", "临时文档已关闭（未保存）");
  else log("WARN", "关闭临时文档失败：" + close.error);

  try {
    if (typeof alert === "function") {
      alert("题注测试3完成：PASS " + passCount + "，FAIL " + failCount +
        "。\r\n详细结果见新建的报告文档。");
    }
  } catch (e9) {}
}

function writeReport() {
  var text = "EQNUM 题注测试3报告\r时间：" + new Date().toString() + "\r" + reportLines.join("\r");
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
