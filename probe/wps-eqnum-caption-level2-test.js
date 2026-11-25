/*
 * WPS 题注剩余问题探测
 *
 * 验证三件事：
 *   1. 二级章层级：STYLEREF 2 \s + SEQ 公式 \s 2 是否解析、是否按标题 2 重置
 *   2. 在表格单元格套用「题注」样式，是否破坏两栏表格结构或单元格垂直居中
 *   3. SEQ EQNUM 切换到 SEQ 公式 后，编号是否从 1 重新开始
 *
 * 安全：临时文档 + 正式文档内存副本，正式文档只读。
 */

function EQNUM_CAPTION_LEVEL2_TEST() {
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

function logParagraph(label, p) {
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

function addFieldAtEnd(doc, p, code) {
  var end = -1;
  try { end = Number(p.Range.End) - 1; } catch (e0) { return { ok: false, error: errText(e0) }; }
  return run(function () {
    var f = doc.Fields.Add(doc.Range(end, end), -1, code, false);
    f.Update();
    return true;
  });
}

function makeSimpleCaption(doc, seqId) {
  var p = addParagraphAtEnd(doc);
  if (!p) return { ok: false, error: "无法新建段落" };

  var start = -1;
  try { start = Number(p.Range.Start); } catch (e0) { return { ok: false, error: errText(e0) }; }

  run(function () { doc.Range(start, start).InsertAfter("("); return true; });
  var f = addFieldAtEnd(doc, p, " SEQ " + seqId + " \\* ARABIC ");
  if (!f.ok) return f;

  var end = -1;
  try { end = Number(p.Range.End) - 1; } catch (e1) {}
  run(function () { if (end >= 0) doc.Range(end, end).InsertAfter(")"); return true; });

  return { ok: true, para: lastFieldParagraph(doc) };
}

function makeLevel2Caption(doc, p, level) {
  var idx = 0;
  try { idx = Number(doc.Range(0, p.Range.Start).Paragraphs.Count) + 1; } catch (e0) {}

  run(function () { p.Range.InsertParagraphAfter(); return true; });

  var np = null;
  try { np = doc.Paragraphs.Item(idx + 1); } catch (e1) {}
  if (!np) return { ok: false, error: "无法读取标题后的新段落" };

  var start = -1;
  try { start = Number(np.Range.Start); } catch (e2) { return { ok: false, error: errText(e2) }; }

  run(function () { doc.Range(start, start).InsertAfter("("); return true; });

  var f1 = addFieldAtEnd(doc, np, ' STYLEREF ' + level + ' \\s \\* ARABIC ');
  if (!f1.ok) return f1;

  var end1 = -1;
  try { end1 = Number(np.Range.End) - 1; } catch (e3) {}
  run(function () { if (end1 >= 0) doc.Range(end1, end1).InsertAfter("."); return true; });

  var f2 = addFieldAtEnd(doc, np, " SEQ 公式 \\* ARABIC \\s " + level + " ");
  if (!f2.ok) return f2;

  var end2 = -1;
  try { end2 = Number(np.Range.End) - 1; } catch (e4) {}
  run(function () { if (end2 >= 0) doc.Range(end2, end2).InsertAfter(")"); return true; });

  run(function () {
    try { np.Style = "题注"; } catch (e5) {
      try { np.Style = "Caption"; } catch (e6) {}
    }
    return true;
  });

  return { ok: true, para: np };
}

function findHeading2Groups(doc, maxScan) {
  var total = 0;
  try { total = Number(doc.Paragraphs.Count); } catch (e) { return { all: [] }; }
  var limit = Math.min(total, maxScan || 3000);

  var h1 = "";
  var groups = [];
  for (var i = 1; i <= limit; i++) {
    var p = null;
    try { p = doc.Paragraphs.Item(i); } catch (e1) { continue; }
    var sn = styleName(p);

    if (/^(标题|Heading)\s*1$/i.test(sn)) {
      h1 = listString(p);
    } else if (/^(标题|Heading)\s*2$/i.test(sn)) {
      groups.push({ p: p, h1: h1, h2: listString(p), para: i });
      if (groups.length >= 12) break;
    }
  }

  var a = null, b = null;
  for (var g = 0; g < groups.length; g++) {
    if (!a) a = groups[g];
    else if (a.h1 !== groups[g].h1) { b = groups[g]; break; }
  }

  return { a: a, b: b, all: groups };
}

function num(v, def) {
  if (v === undefined || v === null) return def;
  var n = Number(v);
  return isNaN(n) ? def : n;
}

function snapshotCell(table) {
  var s = {};
  try { s.rows = num(table.Rows.Count, -1); } catch (e1) { s.rows = -1; }
  try { s.cols = num(table.Columns.Count, -1); } catch (e2) { s.cols = -1; }
  try { s.widthType = num(table.PreferredWidthType, -1); } catch (e3) { s.widthType = -1; }
  try { s.width = num(table.PreferredWidth, -1); } catch (e4) { s.width = -1; }
  try { s.borders = num(table.Borders.Enable, -2); } catch (e5) { s.borders = -2; }

  var c = null;
  try { c = table.Cell(1, 2); } catch (e6) {}
  if (c) {
    try { s.cellWidth = num(c.Width, -1); } catch (e7) { s.cellWidth = -1; }
    try { s.valign = num(c.VerticalAlignment, -1); } catch (e8) { s.valign = -1; }
    try { s.paraCount = num(c.Range.Paragraphs.Count, -1); } catch (e9) { s.paraCount = -1; }
    try {
      var pf = c.Range.ParagraphFormat;
      s.align = num(pf.Alignment, -1);
      s.spaceBefore = num(pf.SpaceBefore, -1);
      s.spaceAfter = num(pf.SpaceAfter, -1);
      s.lineSpacing = num(pf.LineSpacing, -1);
    } catch (e10) {}
    try {
      s.fontName = String(c.Range.Font.Name);
      s.fontSize = num(c.Range.Font.Size, -1);
    } catch (e11) {}
  }
  return s;
}

function snapshotText(s) {
  var keys = ["rows", "cols", "widthType", "width", "borders", "cellWidth",
    "valign", "paraCount", "align", "spaceBefore", "spaceAfter", "lineSpacing",
    "fontName", "fontSize"];
  var parts = [];
  for (var i = 0; i < keys.length; i++) {
    if (s[keys[i]] !== undefined) parts.push(keys[i] + "=" + s[keys[i]]);
  }
  return parts.join("，");
}

/* ============================================================
 * 主流程
 * ============================================================ */
function main() {
  section("零、脚本信息");
  log("INFO", "开始时间：" + new Date().toString());
  log("INFO", "正式文档只读；测试在临时文档和内存副本中进行。");

  var target = null;
  var totalDocs = 0;
  try { totalDocs = Application.Documents.Count; } catch (e0) {}
  var best = -1;
  for (var di = 1; di <= totalDocs; di++) {
    var d = null;
    try { d = Application.Documents.Item(di); } catch (e1) { continue; }
    var oc = -1;
    try { oc = Number(d.OMaths.Count); } catch (e2) {}
    if (oc > best) { best = oc; target = d; }
  }
  if (target) log("INFO", "目标文档=" + target.Name);

  var tmp = addDocument();
  if (!tmp) {
    log("FAIL", "无法创建临时文档");
    writeReport();
    return;
  }
  log("INFO", "临时文档=" + tmp.Name);
  run(function () { tmp.Activate(); return true; });

  section("一、SEQ EQNUM 切换到 SEQ 公式");
  var e1 = makeSimpleCaption(tmp, "EQNUM");
  var e2 = makeSimpleCaption(tmp, "EQNUM");
  var f1 = makeSimpleCaption(tmp, "公式");
  var f2 = makeSimpleCaption(tmp, "公式");

  log(e1.ok ? "PASS" : "FAIL", "第一个 SEQ EQNUM：" + (e1.ok ? "成功" : e1.error));
  if (e1.ok) logParagraph("EQNUM#1", e1.para);
  log(e2.ok ? "PASS" : "FAIL", "第二个 SEQ EQNUM：" + (e2.ok ? "成功" : e2.error));
  if (e2.ok) logParagraph("EQNUM#2", e2.para);
  log(f1.ok ? "PASS" : "FAIL", "第一个 SEQ 公式：" + (f1.ok ? "成功" : f1.error));
  if (f1.ok) logParagraph("公式#1", f1.para);
  log(f2.ok ? "PASS" : "FAIL", "第二个 SEQ 公式：" + (f2.ok ? "成功" : f2.error));
  if (f2.ok) logParagraph("公式#2", f2.para);

  var switchOk = false;
  if (f1.ok && f2.ok) {
    var t1 = String(f1.para.Range.Text);
    var t2 = String(f2.para.Range.Text);
    switchOk = (t1.indexOf("(1)") >= 0 && t2.indexOf("(2)") >= 0);
  }
  log(switchOk ? "PASS" : "FAIL",
    "切换 SEQ 标识后是否从 1 开始：" + (switchOk ? "是" : "否"));

  section("二、表格单元格套用题注样式");
  var tbl = null;
  var tr = run(function () {
    return tmp.Tables.Add(tmp.Range(tmp.Content.End - 1, tmp.Content.End - 1), 1, 2);
  });

  if (tr.ok && tr.value) {
    tbl = tr.value;
    run(function () { tbl.Borders.Enable = 0; return true; });
    run(function () { tbl.PreferredWidthType = 2; tbl.PreferredWidth = 100; return true; });
    run(function () { tbl.Cell(1, 1).Range.Text = "公式"; return true; });
    run(function () { tbl.Cell(1, 2).Range.Text = ""; return true; });
    run(function () { tbl.Cell(1, 2).VerticalAlignment = 1; return true; });
    run(function () { tbl.Cell(1, 2).Range.ParagraphFormat.Alignment = 2; return true; });

    var before = snapshotCell(tbl);
    log("INFO", "套用前：" + snapshotText(before));

    var cp = run(function () { return tbl.Cell(1, 2).Range.Paragraphs.Item(1); });
    if (cp.ok && cp.value) {
      var cStart = -1;
      try { cStart = Number(cp.value.Range.Start); } catch (e2) {}
      run(function () { if (cStart >= 0) tmp.Range(cStart, cStart).InsertAfter("("); return true; });

      var cf1 = addFieldAtEnd(tmp, cp.value, " SEQ 公式 \\* ARABIC ");
      if (cf1.ok) {
        var cEnd = -1;
        try { cEnd = Number(cp.value.Range.End) - 1; } catch (e3) {}
        run(function () { if (cEnd >= 0) tmp.Range(cEnd, cEnd).InsertAfter(")"); return true; });
      }

      run(function () {
        try { cp.value.Style = "题注"; } catch (e4) {
          try { cp.value.Style = "Caption"; } catch (e5) {}
        }
        return true;
      });

      var after = snapshotCell(tbl);
      log("INFO", "套用后：" + snapshotText(after));

      var structuralOk = (before.rows === after.rows &&
        before.cols === after.cols &&
        before.widthType === after.widthType &&
        before.width === after.width &&
        before.borders === after.borders &&
        before.valign === after.valign &&
        before.paraCount === after.paraCount);

      log(structuralOk ? "PASS" : "FAIL",
        "表格结构和单元格垂直居中是否保持不变：" + (structuralOk ? "是" : "否"));

      var formatChanged = (before.align !== after.align ||
        before.spaceBefore !== after.spaceBefore ||
        before.spaceAfter !== after.spaceAfter ||
        before.fontName !== after.fontName ||
        before.fontSize !== after.fontSize);

      log(formatChanged ? "WARN" : "INFO",
        "段落/字体格式是否变化：" + (formatChanged ? "是（题注样式带来的正常变化）" : "否"));

      logParagraph("单元格题注", cp.value);
    } else {
      log("FAIL", "无法读取单元格段落");
    }
  } else {
    log("FAIL", "创建表格失败");
  }

  section("三、二级章层级 STYLEREF 2 + SEQ \\s 2");
  if (target) {
    var copy = null;
    var cr = run(function () { return Application.Documents.Add(target.FullName); });
    if (cr.ok && cr.value) {
      var same = false;
      try { same = (String(cr.value.FullName) === String(target.FullName)); } catch (e6) {}
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

      var groups = findHeading2Groups(copy, 3000);
      log("INFO", "找到标题 2 数量=" + groups.all.length +
        "，可用于重置测试的两组=" + (groups.a && groups.b ? "有" : "无"));

      if (groups.all.length === 0) {
        log("FAIL", "副本中没有标题 2");
      } else {
        var ga = groups.a;
        var gb = groups.b || (groups.all.length > 1 ? groups.all[1] : null);

        var ra = makeLevel2Caption(copy, ga.p, 2);
        log(ra.ok ? "PASS" : "FAIL", "第一组题注插入：" + (ra.ok ? "成功" : ra.error));
        if (ra.ok) {
          logParagraph("标题2-A h1=" + ga.h1 + " h2=" + ga.h2, ra.para);
        }

        if (gb) {
          var rb = makeLevel2Caption(copy, gb.p, 2);
          log(rb.ok ? "PASS" : "FAIL", "第二组题注插入：" + (rb.ok ? "成功" : rb.error));
          if (rb.ok) {
            logParagraph("标题2-B h1=" + gb.h1 + " h2=" + gb.h2, rb.para);
          }

          if (ra.ok && rb.ok) {
            var fa = fieldInfo(ra.para.Range, 4);
            var fb = fieldInfo(rb.para.Range, 4);
            var sa = "";
            var sb = "";
            var qa = "";
            var qb = "";
            for (var i = 0; i < fa.length; i++) {
              if (fa[i].code.indexOf("STYLEREF") >= 0) sa = fa[i].result;
              if (fa[i].code.indexOf("SEQ") >= 0) qa = fa[i].result;
            }
            for (i = 0; i < fb.length; i++) {
              if (fb[i].code.indexOf("STYLEREF") >= 0) sb = fb[i].result;
              if (fb[i].code.indexOf("SEQ") >= 0) qb = fb[i].result;
            }

            log(sa && sa.indexOf("未找到") < 0 && sa.indexOf("错误") < 0 ? "PASS" : "FAIL",
              "STYLEREF 2 第一组结果=" + esc(clip(sa, 40)));
            log(sb && sb.indexOf("未找到") < 0 && sb.indexOf("错误") < 0 ? "PASS" : "FAIL",
              "STYLEREF 2 第二组结果=" + esc(clip(sb, 40)));

            log("INFO", "SEQ 结果：A=" + esc(qa) + "，B=" + esc(qb));

            if (ga.h1 !== gb.h1) {
              var ta = String(ra.para.Range.Text);
              var tb = String(rb.para.Range.Text);
              log("INFO", "两组题注文本：A=" + esc(clip(ta, 40)) + "，B=" + esc(clip(tb, 40)));
              log(qa === "1" && qb === "1" ? "PASS" : "WARN",
                "SEQ \\s 2 是否在两组间重置为 1：" +
                (qa === "1" && qb === "1" ? "是" : "否或无法判断"));
            } else {
              log("WARN", "两组标题 2 属于同一标题 1，无法判断是否按标题 2 重置");
            }
          }
        } else {
          log("WARN", "没有第二组标题 2，无法测试重置");
        }
      }

      var cl = run(function () { copy.Close(0); return true; });
      if (cl.ok) log("INFO", "内存副本已关闭（未保存）");
      else log("WARN", "关闭内存副本失败：" + cl.error);
    }
  } else {
    log("WARN", "没有目标文档，跳过二级章测试");
  }

  section("四、汇总");
  log("INFO", "重点：SEQ 切换是否重头、表格结构是否保持、STYLEREF 2 / SEQ \\s 2 是否可用。");

  writeReport();

  var close = run(function () { tmp.Close(0); return true; });
  if (close.ok) log("INFO", "临时文档已关闭（未保存）");
  else log("WARN", "关闭临时文档失败：" + close.error);

  try {
    if (typeof alert === "function") alert("题注剩余问题测试完成，详见报告文档。");
  } catch (e9) {}
}

function writeReport() {
  var text = "EQNUM 题注剩余问题报告\r时间：" + new Date().toString() + "\r" + reportLines.join("\r");
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
