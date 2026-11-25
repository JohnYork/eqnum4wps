/*
 * WPS 题注兼容性测试 2
 *
 * 修正第一轮的两个问题：
 *   1. 交叉引用插入到全新的空段落，避免读到题注自带的 STYLEREF 域
 *   2. 在真实文档副本中验证 STYLEREF 1 \s + SEQ 公式 \s 1 的章号
 *
 * 验证：
 *   - InsertCaption 生成的题注字段
 *   - InsertCrossReference 实际插入的域代码和结果
 *   - 引用后是否产生 _Ref 书签
 *   - 表格单元格内的题注能否被引用
 *   - 真实文档副本中题注章号是否正确
 *
 * 安全：临时文档和内存副本，正式文档只读。
 */

function EQNUM_CAPTION_TEST2() {
var reportLines = [];
var passCount = 0;
var failCount = 0;
var infoCount = 0;

function consoleLog(s) {
  try {
    if (typeof console !== "undefined" && console && console.log) console.log(s);
  } catch (e) {}
}

function log(level, msg) {
  if (level === "PASS") passCount++;
  else if (level === "FAIL") failCount++;
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
  s = s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, function (ch) {
    var code = ch.charCodeAt(0).toString(16);
    while (code.length < 4) code = "0" + code;
    return "\\u" + code;
  });
  return s;
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

function fieldInfo(r, limit) {
  var out = [];
  try {
    var fs = r.Fields;
    var n = Number(fs.Count);
    for (var i = 1; i <= n && i <= limit; i++) {
      var f = null;
      try { f = fs.Item(i); } catch (e1) { continue; }
      var code = "";
      var res = "";
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

function refBookmarks(doc) {
  var out = [];
  try {
    var bs = doc.Bookmarks;
    for (var i = 1; i <= bs.Count; i++) {
      var n = "";
      try { n = String(bs.Item(i).Name); } catch (e1) {}
      if (n.indexOf("_Ref") === 0) out.push(n);
    }
  } catch (e2) {}
  return out;
}

function moveToEnd(doc) {
  run(function () {
    var end = Number(doc.Content.End) - 1;
    if (end < 0) end = 0;
    doc.Range(end, end).Select();
    return true;
  });
}

function newEmptyParagraphAtEnd(doc) {
  var r = run(function () {
    doc.Content.InsertAfter("\r");
    var end = Number(doc.Content.End) - 1;
    if (end < 0) end = 0;
    doc.Range(end, end).Select();
    return end;
  });
  return r.ok ? r.value : -1;
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
    " 文本=" + esc(clip(text, 100)) +
    " 域数=" + fields.length);
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
    if (c > bestCount) {
      bestCount = c;
      best = d;
    }
  }
  return best;
}

function tryCrossRef(doc, label, kind, item) {
  var end = newEmptyParagraphAtEnd(doc);
  if (end < 0) return { ok: false, error: "无法创建空段落" };

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
      var p = doc.Paragraphs.Item(doc.Paragraphs.Count);
      var fields = fieldInfo(p.Range, 10);
      return {
        ok: true,
        before: before,
        after: after,
        fields: fields,
        text: String(p.Range.Text),
        signature: i + 1
      };
    }
    lastErr = r.error;
  }
  return { ok: false, error: lastErr || "未生成引用" };
}

/* ============================================================
 * 主流程
 * ============================================================ */
function main() {
  section("零、脚本信息");
  log("INFO", "开始时间：" + new Date().toString());
  log("INFO", "正式文档只读；测试在临时文档或内存副本中进行。");

  var target = findTargetDoc();
  if (target) log("INFO", "目标文档=" + target.Name + "（OMaths=" + target.OMaths.Count + "）");

  var testDoc = addDocument();
  if (!testDoc) {
    log("FAIL", "无法创建临时文档");
    writeReport();
    return;
  }
  log("INFO", "临时文档=" + testDoc.Name);
  run(function () { testDoc.Activate(); return true; });

  section("一、InsertCaption 生成题注");
  moveToEnd(testDoc);
  var c1 = run(function () { Application.Selection.InsertCaption("公式", "", "", 0, true); return true; });
  log(c1.ok ? "PASS" : "FAIL", "插入第一个题注：" + (c1.ok ? "成功" : c1.error));
  if (c1.ok) lastParagraphInfo(testDoc, "题注1");

  moveToEnd(testDoc);
  var c2 = run(function () { Application.Selection.InsertCaption("公式", "", "", 0, true); return true; });
  log(c2.ok ? "PASS" : "FAIL", "插入第二个题注：" + (c2.ok ? "成功" : c2.error));
  if (c2.ok) lastParagraphInfo(testDoc, "题注2");

  section("二、隔离交叉引用域");
  var kinds = [1, 2, 3];
  for (var ki = 0; ki < kinds.length; ki++) {
    var kind = kinds[ki];
    var xr = tryCrossRef(testDoc, "公式", kind, 1);

    if (!xr.ok) {
      log("FAIL", "类型" + kind + " 插入失败：" + xr.error);
      continue;
    }

    log("INFO", "类型" + kind + " 插入前域数=" + xr.before + "，插入后域数=" + xr.after +
      "，调用方式#" + xr.signature + "，段落文本=" + esc(clip(xr.text, 80)));
    logFields("类型" + kind, xr.fields);

    var hasRef = false;
    var hasStyleref = false;
    var hasSeq = false;
    var okResult = false;
    for (var fi = 0; fi < xr.fields.length; fi++) {
      var code = xr.fields[fi].code.toUpperCase();
      var res = xr.fields[fi].result;
      if (code.indexOf("REF") >= 0) hasRef = true;
      if (code.indexOf("STYLEREF") >= 0) hasStyleref = true;
      if (code.indexOf("SEQ") >= 0) hasSeq = true;
      if (res && res.indexOf("错误") < 0) okResult = true;
    }

    log(hasRef ? "PASS" : "WARN",
      "类型" + kind + " 是否插入 REF 域：" + (hasRef ? "是" : "否") +
      "，是否含 STYLEREF：" + (hasStyleref ? "是" : "否") +
      "，是否含 SEQ：" + (hasSeq ? "是" : "否"));

    log(okResult ? "PASS" : "FAIL", "类型" + kind + " 引用结果是否有效：" + (okResult ? "是" : "否"));
  }

  // 引用第 2 项
  var xr2 = tryCrossRef(testDoc, "公式", 1, 2);
  if (xr2.ok) {
    log("INFO", "引用第 2 项：插入前=" + xr2.before + "，插入后=" + xr2.after +
      "，文本=" + esc(clip(xr2.text, 80)));
    logFields("第2项", xr2.fields);
  } else {
    log("FAIL", "引用第 2 项失败：" + xr2.error);
  }

  var refs = refBookmarks(testDoc);
  log(refs.length > 0 ? "PASS" : "WARN",
    "引用后 _Ref 书签数量=" + refs.length +
    (refs.length ? "：" + refs.slice(0, 10).join(" / ") : ""));

  section("三、表格单元格内的题注");
  moveToEnd(testDoc);
  var tbl = null;
  var tr = run(function () { return testDoc.Tables.Add(testDoc.Range(testDoc.Content.End - 1, testDoc.Content.End - 1), 1, 2); });
  if (tr.ok && tr.value) {
    tbl = tr.value;
    try { tbl.Borders.Enable = 0; } catch (e1) {}
    tbl.Cell(1, 1).Range.Text = "公式";

    var clear = run(function () { tbl.Cell(1, 2).Range.Text = ""; return true; });
    if (clear.ok) {
      var selCell = run(function () { tbl.Cell(1, 2).Range.Select(); return true; });
      if (selCell.ok) {
        var tc = run(function () { Application.Selection.InsertCaption("公式", "", "", 0, true); return true; });
        log(tc.ok ? "PASS" : "FAIL", "在表格单元格内 InsertCaption：" + (tc.ok ? "成功" : tc.error));

        var cp = tbl.Cell(1, 2).Range.Paragraphs.Item(1);
        var ctext = String(cp.Range.Text);
        var cfields = fieldInfo(cp.Range, 10);
        log("INFO", "单元格段落 样式=" + styleName(cp) +
          " 文本=" + esc(clip(ctext, 100)) +
          " 域数=" + cfields.length);
        logFields("单元格题注", cfields);

        var xr3 = tryCrossRef(testDoc, "公式", 1, 3);
        if (xr3.ok) {
          log("INFO", "引用表格题注：插入前=" + xr3.before + "，插入后=" + xr3.after +
            "，文本=" + esc(clip(xr3.text, 80)));
          logFields("表格引用", xr3.fields);
        } else {
          log("FAIL", "引用表格题注失败：" + xr3.error);
        }
      } else {
        log("FAIL", "无法选中表格单元格：" + selCell.error);
      }
    } else {
      log("FAIL", "清空表格单元格失败：" + clear.error);
    }
  } else {
    log("FAIL", "创建表格失败：" + (tr.error || "返回空"));
  }

  section("四、真实文档副本中的章号");
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
      moveToEnd(copy);

      var cc = run(function () { Application.Selection.InsertCaption("公式", "", "", 0, true); return true; });
      if (cc.ok) {
        var info = lastParagraphInfo(copy, "副本题注");
        var gotChapter = false;
        var gotError = false;
        if (info) {
          for (var fi2 = 0; fi2 < info.fields.length; fi2++) {
            var code2 = info.fields[fi2].code;
            var res2 = info.fields[fi2].result;
            if (code2.indexOf("STYLEREF") >= 0 && res2 && res2.indexOf("错误") < 0 &&
                res2.indexOf("未找到") < 0) gotChapter = true;
            if (res2.indexOf("未找到") >= 0 || res2.indexOf("错误") >= 0) gotError = true;
          }
        }
        log(gotChapter ? "PASS" : "FAIL",
          "副本题注章号是否解析成功：" + (gotChapter ? "是" : "否") +
          "，是否有错误结果：" + (gotError ? "是" : "否"));
      } else {
        log("FAIL", "副本中 InsertCaption 失败：" + cc.error);
      }

      var closeCopy = run(function () { copy.Close(0); return true; });
      if (closeCopy.ok) log("INFO", "内存副本已关闭（未保存）");
      else log("WARN", "关闭内存副本失败：" + closeCopy.error);
    }
  } else {
    log("WARN", "没有找到目标文档，跳过副本测试");
  }

  section("五、汇总");
  log("INFO", "PASS=" + passCount + "，FAIL=" + failCount + "，INFO=" + infoCount);
  log("INFO", "判断重点：交叉引用是否插入 REF 域、是否产生 _Ref 书签、副本章号是否解析。");

  writeReport();

  var close = run(function () { testDoc.Close(0); return true; });
  if (close.ok) log("INFO", "临时文档已关闭（未保存）");
  else log("WARN", "关闭临时文档失败：" + close.error);

  try {
    if (typeof alert === "function") {
      alert("题注测试2完成：PASS " + passCount + "，FAIL " + failCount +
        "。\r\n详细结果见新建的报告文档。");
    }
  } catch (e9) {}
}

function writeReport() {
  var text = "EQNUM 题注测试2报告\r时间：" + new Date().toString() + "\r" + reportLines.join("\r");
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
