/*
 * WPS 公式自动编号 —— 环境与 API 能力探测脚本（只读）
 *
 * 安全约定：
 *   1. 不修改当前活动文档，只读取。
 *   2. 所有写入类测试都在新建的临时文档中执行，结束后关闭且不保存。
 *   3. 任何 API 缺失只记录为 FAIL，不中断脚本。
 *
 * 使用方法：
 *   1. 用 WPS 文字打开一个包含公式的 .docx。
 *   2. 打开「工具 → 开发工具 → JS 宏 / 宏编辑器」，新建脚本并粘贴本文件全部内容。
 *   3. 在宏编辑器中运行入口函数 EQNUM_PROBE()。
 *   4. 查看宏编辑器日志，以及自动新建的「EQNUM 探测报告」文档。
 *   5. 把报告内容反馈回来，用于确定实现方案。
 *
 * 说明：WPS 宏编辑器不允许在函数外做变量初始化和函数调用，
 *       因此除本入口函数外，全部常量、状态变量和辅助函数都放在 EQNUM_PROBE() 内部。
 */

function EQNUM_PROBE() {
var MAX_LIST = 30;              // 每类对象最多列出的条数
var MAX_PARAGRAPH_SCAN = 1200;  // 最多扫描的段落数
var OBJ_CHAR = String.fromCharCode(1); // 嵌入式对象在段落文本中的占位字符

var reportLines = [];
var passCount = 0;
var failCount = 0;
var warnCount = 0;
var infoCount = 0;

var INLINE_TYPE_NAMES = {
  1: "嵌入式OLE对象",
  3: "图片",
  4: "链接图片",
  5: "OLE控件",
  6: "水平线",
  7: "图片水平线",
  8: "链接图片水平线",
  9: "项目符号图片",
  10: "脚本锚点",
  11: "OWS锚点",
  12: "图表",
  15: "SmartArt",
  16: "Web视频",
  17: "内容控件"
};

var SELECTION_TYPE_NAMES = {
  0: "无选区",
  1: "插入点",
  2: "普通选区",
  3: "框架",
  4: "列选区",
  5: "行选区",
  6: "块选区",
  7: "行内图形",
  8: "图形"
};

function consoleLog(s) {
  try {
    if (typeof console !== "undefined" && console && console.log) {
      console.log(s);
    }
  } catch (e) {
    // 日志输出失败不影响探测
  }
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
  try {
    if (e.message) return String(e.message);
  } catch (e1) {}
  try {
    if (e.description) return String(e.description);
  } catch (e2) {}
  try {
    return String(e);
  } catch (e3) {
    return "无法读取的错误";
  }
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
  try {
    t = String(v);
  } catch (e) {
    t = "<无法转字符串>";
  }
  return esc(clip(t, 120));
}

// 读取类探测：成功记 INFO，失败记 FAIL
function readValue(label, fn) {
  var r = run(fn);
  if (r.ok) {
    log("INFO", label + " = " + toText(r.value));
    return r.value;
  }
  log("FAIL", label + " 不可用：" + r.error);
  return null;
}

// 新建文档：兼容 Add() 不返回对象的情况，并确认确实新建成功
function addDocument() {
  var beforeCount = -1;
  try {
    beforeCount = Application.Documents.Count;
  } catch (e) {}

  var r = run(function () {
    return Application.Documents.Add();
  });
  if (r.ok && r.value) return r.value;

  var afterCount = -1;
  try {
    afterCount = Application.Documents.Count;
  } catch (e2) {}

  if (beforeCount >= 0 && afterCount === beforeCount + 1) {
    try {
      return Application.ActiveDocument;
    } catch (e3) {}
  }
  return null;
}

function docEndRange(d) {
  var end = 0;
  try {
    end = d.Content.End - 1;
  } catch (e) {
    end = 0;
  }
  if (end < 0) end = 0;
  return d.Range(end, end);
}

function paragraphStyleName(p) {
  try {
    return String(p.Style.NameLocal);
  } catch (e1) {}
  try {
    return String(p.Style);
  } catch (e2) {}
  return "读取失败";
}

/* ============================================================
 * 一、运行环境
 * ============================================================ */
function probeEnvironment() {
  section("一、运行环境");

  readValue("Application.Name", function () { return Application.Name; });
  readValue("Application.Version", function () { return Application.Version; });
  readValue("Application.Build", function () { return Application.Build; });
  readValue("Application.OperatingSystem", function () { return Application.OperatingSystem; });
  readValue("Application.Path", function () { return Application.Path; });
  readValue("Application.Documents.Count", function () { return Application.Documents.Count; });
}

/* ============================================================
 * 二、活动文档概览（只读）
 * ============================================================ */
function probeDocumentOverview(doc) {
  section("二、活动文档概览（只读）");

  readValue("文档名称", function () { return doc.Name; });
  readValue("完整路径", function () { return doc.FullName; });
  readValue("是否只读", function () { return doc.ReadOnly; });
  readValue("是否已保存", function () { return doc.Saved; });
  readValue("段落数", function () { return doc.Paragraphs.Count; });
  readValue("节数", function () { return doc.Sections.Count; });
  readValue("表格数", function () { return doc.Tables.Count; });
  readValue("嵌入式对象数 InlineShapes.Count", function () { return doc.InlineShapes.Count; });
  readValue("浮动图形数 Shapes.Count", function () { return doc.Shapes.Count; });
  readValue("域数 Fields.Count", function () { return doc.Fields.Count; });
  readValue("书签数 Bookmarks.Count", function () { return doc.Bookmarks.Count; });

  readValue("Document.OMaths 集合", function () {
    if (typeof doc.OMaths === "undefined" || doc.OMaths === null) {
      throw new Error("doc.OMaths 不存在");
    }
    return "存在，数量 " + doc.OMaths.Count;
  });

  readValue("Range(0,1).OMaths 集合", function () {
    return "数量 " + doc.Range(0, 1).OMaths.Count;
  });

  readValue("Range.WordOpenXML", function () {
    var xml = doc.Range(0, 1).WordOpenXML;
    return "类型=" + typeof xml + "，长度=" + (xml ? String(xml).length : 0);
  });
}

/* ============================================================
 * 三、嵌入式公式对象（只读）
 * ============================================================ */
function probeInlineShapes(doc) {
  section("三、嵌入式公式对象（只读）");

  var count = -1;
  var cr = run(function () { return doc.InlineShapes.Count; });
  if (cr.ok) {
    count = Number(cr.value);
    log("INFO", "InlineShapes 共 " + count + " 个" + (count > MAX_LIST ? "，以下仅列出前 " + MAX_LIST + " 个" : ""));
  } else {
    log("FAIL", "无法读取 InlineShapes.Count：" + cr.error);
    return;
  }

  var progStat = {};
  var limit = count > 0 ? Math.min(count, MAX_LIST) : 0;

  for (var i = 1; i <= limit; i++) {
    try {
      var shape = doc.InlineShapes.Item(i);
      var type = null;
      try { type = shape.Type; } catch (e1) {}

      var line = "#" + i + " Type=" + type + "(" + (INLINE_TYPE_NAMES[type] || "未知") + ")";

      var prog = null;
      try { prog = shape.OLEFormat.ProgID; } catch (e2) {}
      if (prog) {
        line += " ProgID=" + prog;
        progStat[prog] = (progStat[prog] || 0) + 1;
      }

      var cls = null;
      try { cls = shape.OLEFormat.ClassType; } catch (e3) {}
      if (cls) line += " ClassType=" + cls;

      try { line += " 尺寸=" + Math.round(shape.Width) + "x" + Math.round(shape.Height); } catch (e4) {}
      try { line += " 文本=" + esc(clip(String(shape.Range.Text), 40)); } catch (e5) {}

      log("INFO", line);

      try {
        var para = shape.Range.Paragraphs.Item(1);
        var style = paragraphStyleName(para);
        var listStr = "";
        var inTable = "";
        try { listStr = String(para.Range.ListFormat.ListString); } catch (e6) {}
        try { inTable = String(para.Range.Information(12)); } catch (e7) {}
        log("INFO", "    所在段落：样式=" + style +
          " 自动编号=" + (listStr ? listStr : "(空)") +
          " 在表格内=" + (inTable ? inTable : "未知") +
          " 文本=" + esc(clip(String(para.Range.Text), 60)));
      } catch (e8) {}

      // 仅对前 5 个对象尝试读取 OMML，避免耗时过长
      if (i <= 5) {
        try {
          var xml = shape.Range.WordOpenXML;
          if (typeof xml === "string") {
            log("INFO", "    WordOpenXML 含 m:oMath = " + (xml.indexOf("<m:oMath") >= 0 ? "是" : "否"));
          }
        } catch (e9) {}
      }
    } catch (e) {
      log("FAIL", "读取 InlineShapes #" + i + " 失败：" + errText(e));
    }
  }

  var hasProg = false;
  for (var k in progStat) {
    if (progStat.hasOwnProperty(k)) {
      hasProg = true;
      log("INFO", "ProgID 统计：" + k + " × " + progStat[k]);
    }
  }
  if (count > 0 && !hasProg) {
    log("WARN", "存在嵌入式对象，但读不到任何 OLEFormat.ProgID，公式识别需要改用其他特征（类型/尺寸/文本/XML）");
  }
}

/* ============================================================
 * 四、浮动图形对象（只读）
 * ============================================================ */
function probeShapes(doc) {
  section("四、浮动图形对象（只读）");

  var count = -1;
  var cr = run(function () { return doc.Shapes.Count; });
  if (cr.ok) {
    count = Number(cr.value);
    log("INFO", "Shapes 共 " + count + " 个");
  } else {
    log("FAIL", "无法读取 Shapes.Count：" + cr.error);
    return;
  }

  var limit = count > 0 ? Math.min(count, MAX_LIST) : 0;
  for (var i = 1; i <= limit; i++) {
    try {
      var shape = doc.Shapes.Item(i);
      var line = "#" + i;
      try { line += " 名称=" + shape.Name; } catch (e1) {}
      try { line += " Type=" + shape.Type; } catch (e2) {}
      try { line += " ProgID=" + shape.OLEFormat.ProgID; } catch (e3) {}
      log("INFO", line);
    } catch (e) {
      log("FAIL", "读取 Shapes #" + i + " 失败：" + errText(e));
    }
  }
}

/* ============================================================
 * 五、域与书签（只读）
 * ============================================================ */
function probeFields(doc) {
  section("五、域（只读）");

  var count = -1;
  var cr = run(function () { return doc.Fields.Count; });
  if (cr.ok) {
    count = Number(cr.value);
    log("INFO", "域共 " + count + " 个");
  } else {
    log("FAIL", "无法读取 Fields.Count：" + cr.error);
    return;
  }

  var eqCount = 0;
  var limit = count > 0 ? Math.min(count, MAX_LIST) : 0;
  for (var i = 1; i <= limit; i++) {
    try {
      var f = doc.Fields.Item(i);
      var type = "";
      var code = "";
      var result = "";
      try { type = String(f.Type); } catch (e1) {}
      try { code = String(f.Code.Text); } catch (e2) {}
      try { result = String(f.Result.Text); } catch (e3) {}
      if (/^\s*EQ\b/i.test(code)) eqCount++;
      log("INFO", "#" + i + " Type=" + type + " 代码=" + esc(clip(code, 80)) + " 结果=" + esc(clip(result, 40)));
    } catch (e) {
      log("FAIL", "读取域 #" + i + " 失败：" + errText(e));
    }
  }

  if (eqCount > 0) {
    log("WARN", "检测到 " + eqCount + " 个 EQ 域，可能包含 Word 公式域，需要在识别逻辑中单独处理");
  }
}

function probeBookmarks(doc) {
  section("六、书签（只读）");

  var count = -1;
  var cr = run(function () { return doc.Bookmarks.Count; });
  if (cr.ok) {
    count = Number(cr.value);
    log("INFO", "书签共 " + count + " 个");
  } else {
    log("FAIL", "无法读取 Bookmarks.Count：" + cr.error);
    return;
  }

  var limit = count > 0 ? Math.min(count, 20) : 0;
  for (var i = 1; i <= limit; i++) {
    try {
      log("INFO", "书签 #" + i + "：" + doc.Bookmarks.Item(i).Name);
    } catch (e) {}
  }

  readValue("Bookmarks.Exists 方法", function () {
    return doc.Bookmarks.Exists("EQNUM_NOT_EXIST_TEST");
  });
}

/* ============================================================
 * 七、段落扫描：标题样式与对象字符（只读）
 * ============================================================ */
function probeParagraphs(doc) {
  section("七、段落扫描：标题与对象字符（只读）");

  var total = -1;
  var cr = run(function () { return doc.Paragraphs.Count; });
  if (cr.ok) total = Number(cr.value);
  else {
    log("FAIL", "无法读取 Paragraphs.Count：" + cr.error);
    return;
  }

  var scan = Math.min(total, MAX_PARAGRAPH_SCAN);
  log("INFO", "段落总数=" + total + "，本次扫描前 " + scan + " 个");

  var headingCount = 0;
  var objParaCount = 0;

  for (var i = 1; i <= scan; i++) {
    var p = null;
    try {
      p = doc.Paragraphs.Item(i);
    } catch (e) {
      continue;
    }

    var sn = paragraphStyleName(p);
    var txt = "";
    try { txt = String(p.Range.Text); } catch (e1) {}

    if (/^(标题|Heading)\s*[12]$/i.test(sn)) {
      headingCount++;
      if (headingCount <= 10) {
        var listStr = "";
        var listType = "";
        try { listStr = String(p.Range.ListFormat.ListString); } catch (e2) {}
        try { listType = String(p.Range.ListFormat.ListType); } catch (e3) {}
        log("INFO", "标题段 #" + i + " 样式=" + sn +
          " 自动编号=" + (listStr ? listStr : "(空)") +
          " ListType=" + listType +
          " 文本=" + esc(clip(txt, 50)));
      }
    }

    if (txt.indexOf(OBJ_CHAR) >= 0) {
      objParaCount++;
      if (objParaCount <= 10) {
        log("INFO", "含对象字符段 #" + i + " 样式=" + sn + " 文本=" + esc(clip(txt, 80)));
      }
    }
  }

  log("INFO", "扫描结果：一/二级标题段 " + headingCount + " 个，含嵌入式对象字符的段落 " + objParaCount + " 个");

  readValue("样式「标题 1」", function () { return doc.Styles.Item("标题 1").NameLocal; });
  readValue("样式「Heading 1」", function () { return doc.Styles.Item("Heading 1").NameLocal; });
  readValue("样式「标题 2」", function () { return doc.Styles.Item("标题 2").NameLocal; });
}

/* ============================================================
 * 八、文档自定义属性（只读）
 * ============================================================ */
function probeCustomProps(doc) {
  section("八、文档自定义属性（只读）");

  var count = -1;
  var cr = run(function () { return doc.CustomDocumentProperties.Count; });
  if (cr.ok) {
    count = Number(cr.value);
    log("INFO", "自定义属性共 " + count + " 个");
  } else {
    log("FAIL", "无法读取 CustomDocumentProperties.Count：" + cr.error);
    return;
  }

  var limit = count > 0 ? Math.min(count, 20) : 0;
  for (var i = 1; i <= limit; i++) {
    try {
      var prop = doc.CustomDocumentProperties.Item(i);
      log("INFO", "属性：" + prop.Name + " = " + esc(clip(String(prop.Value), 60)));
    } catch (e) {}
  }

  readValue("CustomDocumentProperties.Add 方法", function () {
    return typeof doc.CustomDocumentProperties.Add;
  });
}

/* ============================================================
 * 九、光标与选区（只读）
 * ============================================================ */
function probeSelection() {
  section("九、光标与选区（只读）");

  var sr = run(function () { return Application.Selection; });
  if (!sr.ok || !sr.value) {
    log("FAIL", "无法获取 Selection：" + (sr.error || "返回空"));
    return;
  }

  var sel = sr.value;
  var type = readValue("Selection.Type", function () { return sel.Type; });
  log("INFO", "Selection.Type 说明 = " + (SELECTION_TYPE_NAMES[type] || "未知"));

  readValue("选区起点 Start", function () { return sel.Range.Start; });
  readValue("选区终点 End", function () { return sel.Range.End; });
  readValue("选区长度", function () { return sel.Range.End - sel.Range.Start; });
  readValue("选区文本", function () { return esc(clip(String(sel.Range.Text), 80)); });
  readValue("选区段落数", function () { return sel.Range.Paragraphs.Count; });
  readValue("Selection.InlineShapes 数", function () { return sel.InlineShapes.Count; });
  readValue("选区 Range.InlineShapes 数", function () { return sel.Range.InlineShapes.Count; });
  readValue("选区表格数", function () { return sel.Range.Tables.Count; });
  readValue("选区域数", function () { return sel.Range.Fields.Count; });
  readValue("选区 OMaths 数", function () { return sel.Range.OMaths.Count; });

  readValue("Range.Areas 是否可用", function () {
    return "typeof=" + typeof sel.Range.Areas + "，数量=" + sel.Range.Areas.Count;
  });
}

/* ============================================================
 * 十、写入能力（临时文档，不保存）
 * ============================================================ */
function testSeqFields(d) {
  try {
    docEndRange(d).InsertAfter("SEQ测试：");
    for (var k = 1; k <= 2; k++) {
      d.Fields.Add(docEndRange(d), -1, " SEQ EQNUM_TEST \\* ARABIC ", false);
      docEndRange(d).InsertAfter("  ");
    }
    // 测试 \r 1 重置
    docEndRange(d).InsertAfter("重置：");
    d.Fields.Add(docEndRange(d), -1, " SEQ EQNUM_TEST \\r 1 \\* ARABIC ", false);

    d.Fields.Update();

    var vals = [];
    var codes = [];
    for (var i = 1; i <= d.Fields.Count; i++) {
      var f = d.Fields.Item(i);
      var code = "";
      try { code = String(f.Code.Text); } catch (e1) {}
      if (code.indexOf("EQNUM_TEST") >= 0) {
        var res = "";
        try { res = String(f.Result.Text); } catch (e2) {}
        vals.push(res);
        codes.push(code);
      }
    }

    if (vals.length >= 2 && vals[0] === "1" && vals[1] === "2") {
      log("PASS", "SEQ 域可用：连续两个 SEQ 结果为 1、2");
    } else if (vals.length > 0) {
      log("WARN", "SEQ 域已插入但结果异常：[" + vals.join(" | ") + "]");
    } else {
      log("FAIL", "SEQ 域插入后未找到对应域");
    }

    if (vals.length >= 3) {
      log("INFO", "SEQ 重置测试结果：" + vals[2] + (vals[2] === "1" ? "（\\r 1 可用）" : "（\\r 1 未生效）"));
    }
    log("INFO", "SEQ 域代码：" + esc(codes.join(" / ")));
  } catch (e) {
    log("FAIL", "SEQ 域测试失败：" + errText(e));
  }
}

function testBookmarkAndRef(d) {
  try {
    docEndRange(d).InsertAfter("\r参考测试：");
    var start = d.Content.End - 1;
    docEndRange(d).InsertAfter("2.1");
    d.Bookmarks.Add("EQNUM_BM_TEST", d.Range(start, start + 3));

    var exists = false;
    try { exists = d.Bookmarks.Exists("EQNUM_BM_TEST"); } catch (e1) {}
    if (exists) {
      log("PASS", "书签创建成功：EQNUM_BM_TEST");
    } else {
      log("FAIL", "书签创建后 Bookmarks.Exists 返回 false");
    }

    docEndRange(d).InsertAfter(" → ");
    var refField = d.Fields.Add(docEndRange(d), -1, " REF EQNUM_BM_TEST \\h ", false);
    d.Fields.Update();

    var res = "";
    try { res = String(refField.Result.Text); } catch (e2) {}
    if (res.indexOf("2.1") >= 0) {
      log("PASS", "REF 域可用：引用书签显示 " + esc(res));
    } else {
      log("FAIL", "REF 域结果异常：" + esc(res));
    }

    var code = "";
    try { code = String(refField.Code.Text); } catch (e3) {}
    log("INFO", "REF 域代码：" + esc(code));
  } catch (e) {
    log("FAIL", "书签/REF 域测试失败：" + errText(e));
  }
}

function testCustomPropsWrite(d) {
  try {
    d.CustomDocumentProperties.Add("EQNUM_CFG_TEST", false, 4, "mode=chapter;sub=on");
    var value = "";
    try {
      value = String(d.CustomDocumentProperties.Item("EQNUM_CFG_TEST").Value);
    } catch (e1) {
      for (var i = 1; i <= d.CustomDocumentProperties.Count; i++) {
        var p = d.CustomDocumentProperties.Item(i);
        if (String(p.Name) === "EQNUM_CFG_TEST") {
          value = String(p.Value);
          break;
        }
      }
    }

    if (value.indexOf("mode=chapter") >= 0) {
      log("PASS", "文档自定义属性可读写：" + value);
    } else {
      log("FAIL", "文档自定义属性写入后读取不一致：" + value);
    }
  } catch (e) {
    log("FAIL", "文档自定义属性测试失败：" + errText(e));
  }
}

function testTable(d) {
  try {
    docEndRange(d).InsertAfter("\r");
    var tbl = d.Tables.Add(docEndRange(d), 2, 2);

    tbl.Borders.Enable = 0;
    var borderEnabled = "";
    try { borderEnabled = String(tbl.Borders.Enable); } catch (e1) {}

    tbl.Cell(1, 1).Range.Text = "公式";
    tbl.Cell(1, 2).Range.Text = "(2.1)";

    var vertical = "";
    try {
      tbl.Cell(1, 2).VerticalAlignment = 1; // 垂直居中
      vertical = String(tbl.Cell(1, 2).VerticalAlignment);
    } catch (e2) {}

    var cellsBefore = 0;
    try { cellsBefore = tbl.Range.Cells.Count; } catch (e3) {}

    tbl.Cell(1, 2).Merge(tbl.Cell(2, 2));
    var cellsAfter = 0;
    try { cellsAfter = tbl.Range.Cells.Count; } catch (e4) {}

    log("INFO", "表格测试：边框 Enable=" + borderEnabled +
      "，单元格垂直对齐=" + vertical +
      "，合并前单元格=" + cellsBefore +
      "，合并后单元格=" + cellsAfter);

    if (cellsAfter === cellsBefore - 1) {
      log("PASS", "单元格合并可用（右列可做垂直居中的主号）");
    } else {
      log("FAIL", "单元格合并没有生效");
    }

    try {
      tbl.PreferredWidthType = 2; // 百分比
      tbl.PreferredWidth = 100;
      log("PASS", "表格宽度可设为 100%（对应版心宽）");
    } catch (e5) {
      log("FAIL", "表格宽度设置失败：" + errText(e5));
    }

    try {
      tbl.TopPadding = 0;
      tbl.BottomPadding = 0;
      tbl.LeftPadding = 0;
      tbl.RightPadding = 0;
      log("PASS", "单元格内边距可设置为 0");
    } catch (e6) {
      log("FAIL", "单元格内边距设置失败：" + errText(e6));
    }
  } catch (e) {
    log("FAIL", "表格测试失败：" + errText(e));
  }
}

function testHeadingStyle(d) {
  try {
    docEndRange(d).InsertAfter("\r测试标题");
    var p = d.Paragraphs.Item(d.Paragraphs.Count);

    var applied = "";
    try {
      p.Style = "标题 1";
      applied = "标题 1";
    } catch (e1) {
      try {
        p.Style = "Heading 1";
        applied = "Heading 1";
      } catch (e2) {}
    }

    var sn = paragraphStyleName(p);
    var listStr = "";
    var listType = "";
    try { listStr = String(p.Range.ListFormat.ListString); } catch (e3) {}
    try { listType = String(p.Range.ListFormat.ListType); } catch (e4) {}

    if (sn !== "读取失败") {
      log("PASS", "标题样式可设置：应用 " + applied + "，读取 NameLocal=" + sn +
        "，ListString=" + (listStr ? listStr : "(空)") +
        "，ListType=" + listType);
    } else {
      log("FAIL", "标题样式设置失败");
    }
  } catch (e) {
    log("FAIL", "标题样式测试失败：" + errText(e));
  }
}

function testFieldUpdate(d) {
  try {
    d.Fields.Update();
    log("PASS", "Document.Fields.Update() 可调用");
  } catch (e) {
    log("FAIL", "Document.Fields.Update() 不可用：" + errText(e));
  }
}

function probeWrite() {
  section("十、写入能力（临时文档，不保存）");

  var d = addDocument();
  if (!d) {
    log("FAIL", "无法创建临时文档，跳过写入能力测试");
    return;
  }

  log("INFO", "临时文档已创建：" + toText(d.Name));
  try {
    testSeqFields(d);
    testBookmarkAndRef(d);
    testCustomPropsWrite(d);
    testTable(d);
    testHeadingStyle(d);
    testFieldUpdate(d);
  } finally {
    var name = "";
    try { name = String(d.Name); } catch (e) {}
    var cr = run(function () {
      d.Close(0); // 0 = 不保存
      return true;
    });
    if (cr.ok) log("INFO", "临时文档已关闭（未保存）：" + name);
    else log("FAIL", "关闭临时文档失败：" + cr.error);
  }
}

/* ============================================================
 * 汇总与报告
 * ============================================================ */
function finish() {
  section("十一、汇总");

  log("INFO", "PASS=" + passCount + "，FAIL=" + failCount + "，WARN=" + warnCount + "，INFO=" + infoCount);
  log("INFO", "说明：FAIL 仅表示该项 API 不可用或行为不符，不代表脚本执行失败。");

  var text = "EQNUM 探测报告\r\n时间：" + new Date().toString() + "\r\n" + reportLines.join("\r\n");
  var reportDoc = addDocument();
  if (reportDoc) {
    try {
      reportDoc.Content.Text = text;
      try { reportDoc.Content.Font.Name = "Consolas"; } catch (e1) {}
      try { reportDoc.Content.Font.Size = 10; } catch (e2) {}
      log("INFO", "探测报告已写入新建文档，可直接复制内容反馈");
    } catch (e) {
      log("WARN", "写入报告文档失败：" + errText(e));
    }
  } else {
    log("WARN", "无法创建报告文档，请直接复制宏编辑器日志");
  }

  try {
    if (typeof alert === "function") {
      alert("EQNUM 探测完成：PASS " + passCount + "，FAIL " + failCount + "，WARN " + warnCount +
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
  log("INFO", "本脚本不修改当前文档；写入类测试全部在临时新文档中进行，结束后不保存关闭。");

  probeEnvironment();

  var docR = run(function () { return Application.ActiveDocument; });
  var doc = (docR.ok && docR.value) ? docR.value : null;

  if (doc) {
    probeDocumentOverview(doc);
    probeInlineShapes(doc);
    probeShapes(doc);
    probeFields(doc);
    probeBookmarks(doc);
    probeParagraphs(doc);
    probeCustomProps(doc);
  } else {
    log("WARN", "没有打开的活动文档，跳过所有文档相关探测（选区与写入能力测试仍会执行）");
  }

  probeSelection();
  probeWrite();
  finish();
}

  try {
    main();
  } catch (e) {
    consoleLog("脚本执行异常：" + errText(e));
    try {
      if (typeof alert === "function") {
        alert("脚本执行异常：" + errText(e));
      }
    } catch (e2) {}
  }
}
