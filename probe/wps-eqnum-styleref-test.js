/*
 * WPS STYLEREF 域兼容性测试（不修改正式文档）
 *
 * 测试内容：
 *   1. STYLEREF "标题 1" 是否能返回标题文本
 *   2. STYLEREF "标题 1" \n 是否能返回标题编号
 *   3. STYLEREF "标题 1" \w 是否能返回完整编号
 *   4. 在表格单元格内是否可用
 *   5. 修改标题文本后，更新域是否能反映变化
 *
 * 做法：
 *   - 优先用 Documents.Add(目标文件) 生成一份内存副本
 *   - 失败则新建空白文档，并从目标文档复制一个“标题 1”段落
 *   - 所有测试都在测试文档中完成，正式文档只读
 *   - 测试结束后关闭测试文档（不保存），保留报告文档
 *
 * 使用：在宏编辑器运行入口函数 EQNUM_STYLEREF_TEST()。
 */

function EQNUM_STYLEREF_TEST() {
var reportLines = [];
var infoCount = 0;
var failCount = 0;
var passCount = 0;

function consoleLog(s) {
  try {
    if (typeof console !== "undefined" && console && console.log) console.log(s);
  } catch (e) {}
}

function log(level, msg) {
  if (level === "FAIL") failCount++;
  else if (level === "PASS") passCount++;
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

function styleName(p) {
  try { return String(p.Style.NameLocal); } catch (e1) {}
  try { return String(p.Style); } catch (e2) {}
  return "读取失败";
}

function isHeading1(style) {
  return /^(标题|Heading)\s*1$/i.test(style);
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
    if (isHeading1(styleName(p))) return p;
  }
  return null;
}

function addField(d, label, code) {
  d.Content.InsertAfter("\r" + label + "：");

  var f = null;
  var fr = run(function () {
    return d.Fields.Add(endRange(d), -1, code, false);
  });

  if (!fr.ok || !fr.value) {
    log("FAIL", label + " 插入失败：" + (fr.error || "返回空"));
    return { ok: false, result: "" };
  }

  f = fr.value;
  var ur = run(function () { f.Update(); return true; });

  var res = "";
  try { res = String(f.Result.Text); } catch (e1) {}

  var bad = (!ur.ok) || res === "" || res.indexOf("错误") >= 0 || /error!/i.test(res);
  log(bad ? "FAIL" : "PASS",
    label + "：代码=" + code + "，更新=" + (ur.ok ? "成功" : "失败") +
    "，结果=" + esc(clip(res, 60)));

  return { ok: !bad, result: res, field: f };
}

/* ============================================================
 * 主流程
 * ============================================================ */
function main() {
  section("零、脚本信息");
  log("INFO", "开始时间：" + new Date().toString());
  log("INFO", "正式文档只读；所有写入都在测试文档中完成。");

  var target = findTargetDoc();
  if (!target) {
    alert("没有找到包含公式的文档");
    return;
  }
  log("INFO", "目标文档：" + target.Name + "（OMaths=" + target.OMaths.Count + "）");

  section("一、准备测试文档");

  var testDoc = null;
  var created = false;

  // 方式 A：基于目标文件新建内存副本
  var ar = run(function () { return Application.Documents.Add(target.FullName); });
  if (ar.ok && ar.value) {
    var same = false;
    try { same = (String(ar.value.FullName) === String(target.FullName)); } catch (e1) {}
    if (!same) {
      testDoc = ar.value;
      created = true;
      log("PASS", "已用 Documents.Add(目标文件) 创建内存副本：" + testDoc.Name);
    } else {
      run(function () { ar.value.Close(0); return true; });
      log("WARN", "Documents.Add 返回了正式文档本身，已关闭以避免误改");
    }
  } else {
    log("WARN", "Documents.Add(目标文件) 不可用：" + (ar.error || "返回空"));
  }

  if (!testDoc) {
    testDoc = addDocument();
    if (!testDoc) {
      log("FAIL", "无法创建测试文档");
      writeReport();
      return;
    }
    created = true;
    log("INFO", "已新建空白测试文档：" + testDoc.Name);
  }

  run(function () { testDoc.Activate(); return true; });

  // 确保测试文档中有一个“标题 1”
  var heading = findFirstHeading(testDoc, 400);
  if (!heading) {
    var srcHeading = findFirstHeading(target, 1500);
    if (srcHeading) {
      var cp = run(function () {
        srcHeading.Range.Copy();
        testDoc.Content.InsertAfter("\r");
        endRange(testDoc).Select();
        Application.Selection.Paste();
        return true;
      });
      if (cp.ok) log("INFO", "已从目标文档复制一个标题 1 段落");
      else log("WARN", "复制标题 1 失败：" + cp.error);
    }

    heading = findFirstHeading(testDoc, 400);
  }

  if (!heading) {
    testDoc.Content.InsertAfter("1 测试章\r");
    var p = testDoc.Paragraphs.Item(testDoc.Paragraphs.Count);
    var sr = run(function () { p.Style = "标题 1"; return true; });
    if (!sr.ok) {
      sr = run(function () { p.Style = "Heading 1"; return true; });
    }
    log(sr.ok ? "PASS" : "FAIL", "手动创建标题 1 段落：" + (sr.ok ? "成功" : sr.error));
    heading = p;
  }

  // 尝试把目标文档标题 1 的编号模板链接到测试文档
  if (heading) {
    var ls = "";
    try { ls = String(heading.Range.ListFormat.ListString); } catch (e2) {}
    if (!ls) {
      var lr = run(function () {
        var lt = target.Styles.Item("标题 1").ListTemplate;
        testDoc.Styles.Item("标题 1").LinkToListTemplate(lt);
        return true;
      });
      if (lr.ok) log("INFO", "已尝试把目标文档的标题编号模板链接到测试文档");
      else log("WARN", "链接标题编号模板失败：" + lr.error);
    }

    try { ls = String(heading.Range.ListFormat.ListString); } catch (e3) {}
    log("INFO", "测试标题：样式=" + styleName(heading) +
      " 自动编号=" + (ls ? ls : "(空)") +
      " 文本=" + esc(clip(String(heading.Range.Text), 60)));

    if (!ls) {
      log("WARN", "测试文档没有自动编号，\\n 与 \\w 的结果可能不具代表性");
    }
  }

  section("二、STYLEREF 基本字段");
  var textField = addField(testDoc, "标题文本", ' STYLEREF "标题 1" ');
  var numField = addField(testDoc, "标题编号(\\n)", ' STYLEREF "标题 1" \\n ');
  var fullField = addField(testDoc, "完整编号(\\w)", ' STYLEREF "标题 1" \\w ');
  var mergeField = addField(testDoc, "带格式(\\* MERGEFORMAT)", ' STYLEREF "标题 1" \\* MERGEFORMAT ');

  section("三、括号组合测试");
  testDoc.Content.InsertAfter("\r括号组合：(");
  var parenField = null;
  var pfr = run(function () {
    return testDoc.Fields.Add(endRange(testDoc), -1, ' STYLEREF "标题 1" \\n ', false);
  });
  if (pfr.ok && pfr.value) {
    parenField = pfr.value;
    var pur = run(function () { parenField.Update(); return true; });
    testDoc.Content.InsertAfter(")");
    var pres = "";
    try { pres = String(parenField.Result.Text); } catch (e4) {}
    log(pur.ok && pres !== "" ? "PASS" : "FAIL",
      "括号内的编号域结果=" + esc(clip(pres, 40)) +
      (pur.ok ? "" : "，更新失败：" + pur.error));
  } else {
    log("FAIL", "括号组合测试插入失败：" + (pfr.error || "返回空"));
  }

  section("四、表格单元格内测试");
  testDoc.Content.InsertAfter("\r");
  var tbl = null;
  var tr = run(function () { return testDoc.Tables.Add(endRange(testDoc), 1, 2); });
  if (tr.ok && tr.value) {
    tbl = tr.value;
    try { tbl.Borders.Enable = 0; } catch (e5) {}
    tbl.Cell(1, 1).Range.Text = "编号";
    tbl.Cell(1, 2).Range.Text = "";

    var cf = null;
    var cfr = run(function () {
      return testDoc.Fields.Add(tbl.Cell(1, 2).Range, -1, ' STYLEREF "标题 1" \\n ', false);
    });

    if (cfr.ok && cfr.value) {
      cf = cfr.value;
      var cur = run(function () { cf.Update(); return true; });
      var cres = "";
      try { cres = String(cf.Result.Text); } catch (e6) {}
      log(cur.ok && cres !== "" ? "PASS" : "FAIL",
        "单元格内 STYLEREF \\n 结果=" + esc(clip(cres, 40)) +
        (cur.ok ? "" : "，更新失败：" + cur.error));
    } else {
      log("FAIL", "单元格内插入 STYLEREF 失败：" + (cfr.error || "返回空"));
    }
  } else {
    log("FAIL", "创建测试表格失败：" + (tr.error || "返回空"));
  }

  section("五、修改标题后的更新测试");
  if (heading && textField.ok) {
    var before = textField.result;
    var mod = run(function () {
      heading.Range.InsertAfter("（修改）");
      return true;
    });
    if (mod.ok) {
      var f2 = textField.field;
      var upd = run(function () { if (f2) f2.Update(); return true; });
      var after = "";
      try { if (f2) after = String(f2.Result.Text); } catch (e8) {}
      log(upd.ok && after !== before ? "PASS" : "FAIL",
        "标题修改后文本域：前=" + esc(before) + "，后=" + esc(after));
    } else {
      log("FAIL", "修改标题失败：" + mod.error);
    }
  } else {
    log("WARN", "缺少标题或文本域，跳过修改测试");
  }

  section("六、结论");
  log("INFO", "编号域(\\n) 结果=" + esc(numField.result || "(空)"));
  log("INFO", "完整编号(\\w) 结果=" + esc(fullField.result || "(空)"));
  log("INFO", "带格式域结果=" + esc(mergeField.result || "(空)"));
  log("INFO", "如果 \\n 能返回类似 1 或 1.1 的编号，说明可以用 STYLEREF 生成章节前缀。");
  log("INFO", "如果 \\n 返回空或错误，章节前缀仍需用普通文本，并通过全文重编号刷新。");

  writeReport();

  if (testDoc && created) {
    var close = run(function () { testDoc.Close(0); return true; });
    if (close.ok) log("INFO", "测试文档已关闭（未保存）");
    else log("WARN", "关闭测试文档失败：" + close.error);
  }

  try {
    if (typeof alert === "function") {
      alert("STYLEREF 测试完成：PASS " + passCount + "，FAIL " + failCount +
        "。\r\n详细结果见新建的报告文档。");
    }
  } catch (e9) {}
}

function writeReport() {
  var text = "EQNUM STYLEREF 测试报告\r时间：" + new Date().toString() + "\r" + reportLines.join("\r");
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
