/*
 * 制表位版式原型（探针，会修改文档，请用副本）
 *
 * 目的：不用表格，改成“公式段 + 右对齐制表位”，测量单条公式各阶段耗时，
 *       与表格版的日志对比（表格版：旧号 5–27s，转表/排版/写号各约 0.1s）。
 *
 * 运行：把光标放到目标章节任意位置，运行 EQNUM_TABSTOP_TEST()。
 * 处理顺序与正式流程一致：从章末向前。
 */

function EQNUM_TABSTOP_TEST() {
  var doc = null;
  try { doc = Application.ActiveDocument; } catch (e0) {}
  if (!doc) { alert("请先打开文档"); return; }

  var n = 8;
  if (typeof InputBox === "function") {
    var ans = InputBox("测试多少个公式？建议 8（从当前章节末尾向前取）", "EQNUM 制表位原型", "8");
    if (ans === null || ans === undefined) return;
    var num = Number(String(ans).replace(/\s/g, ""));
    if (num > 0) n = Math.floor(num);
  }
  if (n > 40) n = 40;

  var curPara = eqnumTabsCursorParagraph(doc);
  var start = eqnumTabsChapterStart(doc, curPara);
  var end = eqnumTabsChapterEnd(doc, start);

  var cands = [];
  for (var i = end; i >= start && cands.length < n; i--) {
    var p = null;
    try { p = doc.Paragraphs.Item(i); } catch (e1) { continue; }

    var omCount = -1;
    try { omCount = Number(p.Range.OMaths.Count); } catch (e2) { continue; }
    if (omCount !== 1) continue;

    var np = null;
    try { np = doc.Paragraphs.Item(i + 1); } catch (e3) { continue; }
    if (!np) continue;

    var nt = "";
    try { nt = String(np.Range.Text); } catch (e4) {}
    if (!/^[\s]*[\(（]\s*\d+(?:\.\d+)*\s*[\)）][\s]*$/.test(nt)) continue;

    cands.push(i);
  }

  if (cands.length === 0) {
    alert("当前章节没有找到“公式段 + 独立编号段”的候选");
    return;
  }

  var right = eqnumTabsTextWidth(doc);
  var tAll = Date.now();
  var totalDel = 0, totalTab = 0, totalIns = 0, totalWrite = 0;
  var lines = [];

  for (var k = 0; k < cands.length; k++) {
    var pi = cands[k];
    var tItem = Date.now();
    var para = null;
    try { para = doc.Paragraphs.Item(pi); } catch (e5) { continue; }

    // 1) 删除旧编号段落（与正式流程相同的开销点）
    var t0 = Date.now();
    var okDel = false;
    try { doc.Paragraphs.Item(pi + 1).Range.Delete(); okDel = true; } catch (e6) {}
    var dDel = Date.now() - t0;

    // 2) 右对齐制表位
    var t1 = Date.now();
    var tabOk = false;
    try {
      para.Range.ParagraphFormat.TabStops.Add(right, 2, 0);
      tabOk = true;
    } catch (e7) {}
    var dTab = Date.now() - t1;

    // 3) 公式后插入制表符
    var t2 = Date.now();
    var insOk = false;
    try {
      var om = para.Range.OMaths.Item(1);
      var omEnd = Number(om.Range.End);
      doc.Range(omEnd, omEnd).InsertAfter("\t");
      insOk = true;
    } catch (e8) {}
    var dIns = Date.now() - t2;

    // 4) 写入 "( STYLEREF 1 \n . SEQ 公式 \s 1 )"
    var t3 = Date.now();
    var writeOk = false;
    try {
      var pe = Number(para.Range.End) - 1;
      doc.Range(pe, pe).InsertAfter("(");

      pe = Number(para.Range.End) - 1;
      doc.Fields.Add(doc.Range(pe, pe), -1, " STYLEREF 1 \\n ", false).Update();

      pe = Number(para.Range.End) - 1;
      doc.Range(pe, pe).InsertAfter(".");

      pe = Number(para.Range.End) - 1;
      doc.Fields.Add(doc.Range(pe, pe), -1, " SEQ 公式 \\* ARABIC \\s 1 ", false).Update();

      pe = Number(para.Range.End) - 1;
      doc.Range(pe, pe).InsertAfter(")");
      writeOk = true;
    } catch (e9) {}
    var dWrite = Date.now() - t3;

    totalDel += dDel; totalTab += dTab; totalIns += dIns; totalWrite += dWrite;

    var line = "第 " + (k + 1) + "/" + cands.length + " 段#" + pi +
      " 用时 " + (Date.now() - tItem) + "ms（删号 " + dDel +
      "，制表位 " + dTab + "，插制表符 " + dIns + "，写号 " + dWrite + "）" +
      (okDel ? "" : " 删号失败") + (tabOk ? "" : " 制表位失败") +
      (insOk ? "" : " 插入失败") + (writeOk ? "" : " 写号失败");
    lines.push(line);
    try { console.log("[TABS] " + line); } catch (eA) {}
    if (typeof Application.StatusBar !== "undefined") {
      try { Application.StatusBar = "制表位原型：" + (k + 1) + "/" + cands.length; } catch (eB) {}
    }
  }

  try { Application.StatusBar = ""; } catch (eC) {}

  var count = cands.length;
  var summary = "制表位原型完成，共 " + count + " 个，总用时 " + (Date.now() - tAll) + "ms\r" +
    "平均：删号 " + Math.round(totalDel / count) + "ms，" +
    "制表位 " + Math.round(totalTab / count) + "ms，" +
    "插制表符 " + Math.round(totalIns / count) + "ms，" +
    "写号 " + Math.round(totalWrite / count) + "ms\r\r" +
    "右对齐制表位位置：" + right + "pt\r\r" +
    lines.join("\r");

  alert(summary);
}

function eqnumTabsCursorParagraph(doc) {
  var start = 0;
  try { start = Number(Application.Selection.Range.Start); } catch (e) {}
  try { return Number(doc.Range(0, start).Paragraphs.Count) + 1; } catch (e2) { return 1; }
}

function eqnumTabsChapterStart(doc, curPara) {
  var total = 1;
  try { total = Number(doc.Paragraphs.Count); } catch (e) {}
  if (!(curPara >= 1) || curPara > total) curPara = 1;

  for (var i = curPara; i >= 1; i--) {
    var p = null;
    try { p = doc.Paragraphs.Item(i); } catch (e1) { continue; }
    var sn = "";
    try { sn = String(p.Style.NameLocal); } catch (e2) {}
    if (/^(标题|Heading)\s*1$/i.test(sn)) return i;
  }
  return 1;
}

function eqnumTabsChapterEnd(doc, start) {
  var total = 1;
  try { total = Number(doc.Paragraphs.Count); } catch (e) {}
  for (var i = start + 1; i <= total; i++) {
    var p = null;
    try { p = doc.Paragraphs.Item(i); } catch (e1) { continue; }
    var sn = "";
    try { sn = String(p.Style.NameLocal); } catch (e2) {}
    if (/^(标题|Heading)\s*1$/i.test(sn)) return i - 1;
  }
  return total;
}

function eqnumTabsTextWidth(doc) {
  try {
    var ps = doc.PageSetup;
    return Number(ps.PageWidth) - Number(ps.LeftMargin) - Number(ps.RightMargin);
  } catch (e) {
    return 415.3;
  }
}
