/*
 * WPS 题注（Caption）兼容性测试
 *
 * 验证：
 *   1. CaptionLabels 集合与 Add("公式") 是否可用
 *   2. InsertCaption 是否可用，生成的域、样式、_Ref 书签
 *   3. InsertCrossReference 是否可用，引用结果显示什么
 *   4. 表格单元格内的 SEQ 公式 是否能被题注引用识别
 *   5. STYLEREF "标题 1" \n + SEQ 公式 能否生成 2.1 这种章号
 *
 * 安全：所有写入都在临时文档中完成，正式文档只读。
 * 使用：宏编辑器运行入口函数 EQNUM_CAPTION_TEST()。
 */

function EQNUM_CAPTION_TEST() {
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

function endRange(d) {
  var end = 0;
  try { end = d.Content.End - 1; } catch (e) { end = 0; }
  if (end < 0) end = 0;
  return d.Range(end, end);
}

function moveToEnd(d) {
  run(function () {
    d.Range(d.Content.End - 1, d.Content.End - 1).Select();
    return true;
  });
}

function styleName(p) {
  try { return String(p.Style.NameLocal); } catch (e1) {}
  try { return String(p.Style); } catch (e2) {}
  return "读取失败";
}

function listCaptionLabels(doc, limit) {
  var names = [];
  var r = run(function () { return doc.CaptionLabels; });
  if (!r.ok) return { ok: false, error: r.error, names: names };

  var n = -1;
  try { n = Number(r.value.Count); } catch (e1) {}
  for (var i = 1; i <= n && i <= limit; i++) {
    var it = null;
    try { it = r.value.Item(i); } catch (e2) { continue; }
    var nm = "";
    try { nm = String(it.Name); } catch (e3) {}
    names.push(nm || ("#" + i));
  }
  return { ok: true, count: n, names: names };
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

function bookmarkNames(r, limit) {
  var out = [];
  try {
    var bs = r.Bookmarks;
    var n = Number(bs.Count);
    for (var i = 1; i <= n && i <= limit; i++) {
      try { out.push(String(bs.Item(i).Name)); } catch (e1) {}
    }
  } catch (e2) {}
  return out;
}

function describeParagraph(p) {
  var text = "";
  try { text = String(p.Range.Text); } catch (e1) {}
  var fields = fieldInfo(p.Range, 8);
  var bms = bookmarkNames(p.Range, 8);

  log("INFO", "  段落样式=" + styleName(p) +
    " 文本=" + esc(clip(text, 80)) +
    " 域数=" + fields.length +
    " 书签数=" + bms.length);

  for (var i = 0; i < fields.length; i++) {
    log("INFO", "    域#" + (i + 1) + " 代码=" + esc(clip(fields[i].code, 100)) +
      " 结果=" + esc(clip(fields[i].result, 60)));
  }
  for (i = 0; i < bms.length; i++) {
    log("INFO", "    书签#" + (i + 1) + "=" + bms[i]);
  }

  return { fields: fields, bookmarks: bms, text: text };
}

function countSeqFormula(doc) {
  var count = 0;
  try {
    var fs = doc.Fields;
    for (var i = 1; i <= fs.Count; i++) {
      var code = "";
      try { code = String(fs.Item(i).Code.Text); } catch (e1) {}
      if (code.indexOf("SEQ") >= 0 && code.indexOf("公式") >= 0) count++;
    }
  } catch (e2) {}
  return count;
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

function tryCrossRef(doc, label, kind, item) {
  var sigs = [
    function () { Application.Selection.InsertCrossReference(label, kind, item, true, false, false, ""); },
    function () { Application.Selection.InsertCrossReference(label, kind, item, true); },
    function () { Application.Selection.InsertCrossReference(label, kind, item); },
    function () { doc.Selection.InsertCrossReference(label, kind, item, true, false, false, ""); }
  ];

  var lastErr = "";
  for (var i = 0; i < sigs.length; i++) {
    moveToEnd(doc);
    var r = run(sigs[i]);
    if (r.ok) {
      var p = doc.Paragraphs.Item(doc.Paragraphs.Count);
      var fs = fieldInfo(p.Range, 6);
      if (fs.length > 0) {
        return { ok: true, fields: fs, signature: i + 1 };
      }
    } else {
      lastErr = r.error;
    }
  }
  return { ok: false, error: lastErr || "未生成域" };
}

/* ============================================================
 * 主流程
 * ============================================================ */
function main() {
  section("零、脚本信息");
  log("INFO", "开始时间：" + new Date().toString());
  log("INFO", "正式文档只读；所有写入都在临时文档中完成。");

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

  section("一、题注标签 CaptionLabels");
  var labels0 = listCaptionLabels(testDoc, 30);
  if (!labels0.ok) {
    log("FAIL", "CaptionLabels 不可用：" + labels0.error);
  } else {
    log("PASS", "CaptionLabels 可用，数量=" + labels0.count);
    log("INFO", "已有标签=" + (labels0.names.length ? labels0.names.join(" / ") : "(空)"));
  }

  var addLabel = run(function () { return testDoc.CaptionLabels.Add("公式"); });
  if (addLabel.ok) {
    log("PASS", "CaptionLabels.Add(\"公式\") 成功");
  } else {
    log("WARN", "CaptionLabels.Add(\"公式\") 失败或已存在：" + addLabel.error);
  }

  var labels1 = listCaptionLabels(testDoc, 30);
  if (labels1.ok) {
    log("INFO", "添加后标签=" + (labels1.names.length ? labels1.names.join(" / ") : "(空)"));
    var found = false;
    for (var li = 0; li < labels1.names.length; li++) {
      if (labels1.names[li] === "公式") found = true;
    }
    log(found ? "PASS" : "FAIL", "标签列表中是否出现“公式”：" + (found ? "是" : "否"));
  }

  section("二、InsertCaption 插入题注");
  var captionOk = false;
  var captionSignature = "";

  var attempts = [
    { name: "5参数 Position=0 ExcludeLabel=true", fn: function () { Application.Selection.InsertCaption("公式", "", "", 0, true); } },
    { name: "5参数 Position=1 ExcludeLabel=true", fn: function () { Application.Selection.InsertCaption("公式", "", "", 1, true); } },
    { name: "1参数", fn: function () { Application.Selection.InsertCaption("公式"); } },
    { name: "doc.Selection 5参数", fn: function () { testDoc.Selection.InsertCaption("公式", "", "", 0, true); } }
  ];

  for (var ai = 0; ai < attempts.length; ai++) {
    moveToEnd(testDoc);
    var ar = run(attempts[ai].fn);
    if (!ar.ok) {
      log("WARN", "InsertCaption 尝试失败（" + attempts[ai].name + "）：" + ar.error);
      continue;
    }

    var lastPara = null;
    try { lastPara = testDoc.Paragraphs.Item(testDoc.Paragraphs.Count); } catch (e1) {}
    if (!lastPara) continue;

    var desc = describeParagraph(lastPara);
    if (desc.fields.length > 0) {
      captionOk = true;
      captionSignature = attempts[ai].name;
      log("PASS", "InsertCaption 可用：" + captionSignature);
      break;
    } else {
      log("WARN", "InsertCaption 未产生域（" + attempts[ai].name + "）");
    }
  }

  if (!captionOk) {
    log("WARN", "InsertCaption 不可用，改用手工方式插入 SEQ 公式 域");
    moveToEnd(testDoc);
    var mf = run(function () {
      testDoc.Content.InsertAfter("\r公式 ");
      var f = testDoc.Fields.Add(endRange(testDoc), -1, " SEQ 公式 \\* ARABIC ", false);
      f.Update();
      return true;
    });
    if (mf.ok) {
      var lp = testDoc.Paragraphs.Item(testDoc.Paragraphs.Count);
      try { lp.Style = "题注"; } catch (e2) {
        try { lp.Style = "Caption"; } catch (e3) {}
      }
      describeParagraph(lp);
      log("PASS", "手工 SEQ 公式 题注已插入");
    } else {
      log("FAIL", "手工插入 SEQ 公式 失败：" + mf.error);
    }
  }

  // 再插入一个题注，保证至少两个可引用项
  moveToEnd(testDoc);
  var second = run(function () { Application.Selection.InsertCaption("公式", "", "", 0, true); });
  if (!second.ok) {
    run(function () {
      testDoc.Content.InsertAfter("\r公式 ");
      var f = testDoc.Fields.Add(endRange(testDoc), -1, " SEQ 公式 \\* ARABIC ", false);
      f.Update();
      return true;
    });
  }
  var secondPara = testDoc.Paragraphs.Item(testDoc.Paragraphs.Count);
  describeParagraph(secondPara);

  var seqCountBeforeTable = countSeqFormula(testDoc);
  log("INFO", "当前 SEQ 公式 域数量=" + seqCountBeforeTable);

  section("三、InsertCrossReference 交叉引用");
  var refResults = [];

  for (var kind = 1; kind <= 3; kind++) {
    var xr = tryCrossRef(testDoc, "公式", kind, 1);

    if (!xr.ok) {
      log("WARN", "InsertCrossReference 类型" + kind + " 失败：" + xr.error);
      refResults.push({ kind: kind, ok: false, result: "", code: "" });
      continue;
    }

    var rp = testDoc.Paragraphs.Item(testDoc.Paragraphs.Count);
    var last = xr.fields[xr.fields.length - 1];

    run(function () {
      try { rp.Range.Fields.Item(rp.Range.Fields.Count).Update(); } catch (e1) {}
      return true;
    });

    var rfs2 = fieldInfo(rp.Range, 6);
    if (rfs2.length > 0) last = rfs2[rfs2.length - 1];

    var okRef = last.code.indexOf("REF") >= 0 && last.result.indexOf("错误") < 0;
    log(okRef ? "PASS" : "FAIL",
      "InsertCrossReference 类型" + kind + " 结果=" + esc(clip(last.result, 60)) +
      " 代码=" + esc(clip(last.code, 80)));
    refResults.push({ kind: kind, ok: okRef, result: last.result, code: last.code });
  }

  section("四、表格单元格内的题注");
  moveToEnd(testDoc);
  var tbl = null;
  var tr = run(function () { return testDoc.Tables.Add(endRange(testDoc), 1, 2); });
  if (tr.ok && tr.value) {
    tbl = tr.value;
    try { tbl.Borders.Enable = 0; } catch (e1) {}
    tbl.Cell(1, 1).Range.Text = "公式";
    tbl.Cell(1, 2).Range.Text = "";

    var tf = run(function () {
      var f = testDoc.Fields.Add(tbl.Cell(1, 2).Range, -1, " SEQ 公式 \\* ARABIC ", false);
      f.Update();
      return true;
    });
    if (tf.ok) {
      log("PASS", "表格单元格内插入 SEQ 公式 成功");
      var cellPara = tbl.Cell(1, 2).Range.Paragraphs.Item(1);
      describeParagraph(cellPara);

      var seqCount = countSeqFormula(testDoc);
      log("INFO", "表格题注后的 SEQ 公式 域数量=" + seqCount);

      var tableRef = tryCrossRef(testDoc, "公式", 1, seqCount);
      if (tableRef.ok) {
        var tlast = tableRef.fields[tableRef.fields.length - 1];
        var trp = testDoc.Paragraphs.Item(testDoc.Paragraphs.Count);
        run(function () {
          try { trp.Range.Fields.Item(trp.Range.Fields.Count).Update(); } catch (e1) {}
          return true;
        });
        var trfs2 = fieldInfo(trp.Range, 6);
        if (trfs2.length > 0) tlast = trfs2[trfs2.length - 1];

        log(tlast.result.indexOf("错误") < 0 ? "PASS" : "FAIL",
          "引用表格内题注（第 " + seqCount + " 项）结果=" + esc(clip(tlast.result, 60)));
      } else {
        log("FAIL", "引用表格内题注失败：" + tableRef.error);
      }
    } else {
      log("FAIL", "表格单元格内插入 SEQ 公式 失败：" + tf.error);
    }
  } else {
    log("FAIL", "创建测试表格失败：" + (tr.error || "返回空"));
  }

  section("五、STYLEREF + SEQ 生成章号");
  if (target) {
    var srcHeading = findFirstHeading(target, 1500);
    if (srcHeading) {
      var cp = run(function () {
        srcHeading.Range.Copy();
        testDoc.Content.InsertAfter("\r");
        endRange(testDoc).Select();
        Application.Selection.Paste();
        return true;
      });
      if (cp.ok) {
        var hp = testDoc.Paragraphs.Item(testDoc.Paragraphs.Count);
        log("INFO", "已复制标题段落：样式=" + styleName(hp) +
          " 文本=" + esc(clip(String(hp.Range.Text), 60)));

        moveToEnd(testDoc);
        var sf = run(function () {
          testDoc.Content.InsertAfter("\r章号测试：");
          var f1 = testDoc.Fields.Add(endRange(testDoc), -1, ' STYLEREF "标题 1" \\n ', false);
          f1.Update();
          testDoc.Content.InsertAfter(".");
          var f2 = testDoc.Fields.Add(endRange(testDoc), -1, " SEQ 公式 \\* ARABIC ", false);
          f2.Update();
          return true;
        });

        if (sf.ok) {
          var sp = testDoc.Paragraphs.Item(testDoc.Paragraphs.Count);
          var sfs = fieldInfo(sp.Range, 4);
          var sres = "";
          for (var si = 0; si < sfs.length; si++) sres += sfs[si].result;
          var stext = String(sp.Range.Text);
          var looksOk = (stext.indexOf("错误") < 0 && /[0-9]+\.[0-9]+/.test(stext));
          log(looksOk ? "PASS" : "FAIL",
            "STYLEREF + SEQ 域结果=" + esc(clip(sres, 60)) +
            " 文本=" + esc(clip(stext, 80)));
        } else {
          log("FAIL", "STYLEREF + SEQ 插入失败：" + sf.error);
        }
      } else {
        log("WARN", "复制标题失败：" + cp.error);
      }
    } else {
      log("WARN", "目标文档没有找到标题 1，跳过 STYLEREF 测试");
    }
  } else {
    log("WARN", "没有找到目标文档，跳过 STYLEREF 测试");
  }

  section("六、汇总");
  log("INFO", "CaptionLabels 可用=" + (labels1.ok ? "是" : "否"));
  log("INFO", "InsertCaption 可用=" + (captionOk ? "是（" + captionSignature + "）" : "否"));
  log("INFO", "InsertCrossReference 结果：" + refResults.length + " 种类型");
  for (var ri = 0; ri < refResults.length; ri++) {
    log("INFO", "  类型" + refResults[ri].kind + " 通过=" + (refResults[ri].ok ? "是" : "否") +
      " 结果=" + esc(clip(refResults[ri].result, 40)));
  }
  log("INFO", "PASS=" + passCount + "，FAIL=" + failCount);

  writeReport();

  var close = run(function () { testDoc.Close(0); return true; });
  if (close.ok) log("INFO", "临时文档已关闭（未保存）");
  else log("WARN", "关闭临时文档失败：" + close.error);

  try {
    if (typeof alert === "function") {
      alert("题注测试完成：PASS " + passCount + "，FAIL " + failCount +
        "。\r\n详细结果见新建的报告文档。");
    }
  } catch (e9) {}
}

function writeReport() {
  var text = "EQNUM 题注测试报告\r时间：" + new Date().toString() + "\r" + reportLines.join("\r");
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
