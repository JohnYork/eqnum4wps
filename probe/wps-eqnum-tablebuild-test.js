/*
 * “搬公式进表格”探针（会改文档，务必用副本）
 *
 * 背景：ConvertToTable("\t") 会把公式的数学区按制表符切开，
 *       右单元格残留裸 m:r，编号被写进数学区，引用显示为对象，
 *       保存后数学区还会吞掉后续正文。
 *
 * 目标：验证不经过 ConvertToTable 的另一条路——
 *       先建 1x2 空表，再把公式搬进左单元格。
 *
 * 运行：打开原始文档的副本，光标放在一个“独占段落的显示公式”上，
 *       运行 EQNUM_TABLEBUILD_TEST()，把输出发回。
 */

function EQNUM_TABLEBUILD_TEST() {
  var doc = null;
  try { doc = Application.ActiveDocument; } catch (e0) {}
  if (!doc) { alert("请先打开文档"); return; }

  var lines = [];
  function P(s) { lines.push(s); try { console.log("[TBLBUILD] " + s); } catch (e) {} }
  function err(e) {
    if (!e) return "未知";
    try { if (e.message) return String(e.message); } catch (e1) {}
    try { if (e.description) return String(e.description); } catch (e2) {}
    return "无法读取";
  }

  var p = null;
  try { p = Application.Selection.Range.Paragraphs.Item(1); } catch (e1) {}
  if (!p) { alert("取不到光标段落"); return; }

  var om = null;
  try { om = p.Range.OMaths.Item(1); } catch (e2) {}
  if (!om) { alert("光标段落里没有公式，请把光标放进公式段"); return; }

  var pStart = -1, pEnd = -1;
  try { pStart = Number(p.Range.Start); pEnd = Number(p.Range.End); } catch (e3) {}
  P("公式段 " + pStart + ".." + pEnd + " 段落内公式数=" + Number(p.Range.OMaths.Count));

  var tbl = null;
  var added = null;
  try {
    added = doc.Tables.Add(doc.Range(pEnd, pEnd), 1, 2);
    tbl = added;
    P("空表创建成功");
  } catch (e4) {
    P("Tables.Add 失败：" + err(e4));
    alert(lines.join("\r"));
    return;
  }

  function cellStat(tag, cell) {
    var mc = -1, ish = -1, txt = "?";
    try { mc = Number(cell.Range.OMaths.Count); } catch (e) {}
    try { ish = Number(cell.Range.InlineShapes.Count); } catch (e) {}
    try { txt = String(cell.Range.Text).replace(/[\r\n\u0007]/g, "\\n"); } catch (e2) {}
    if (txt.length > 30) txt = txt.substring(0, 30) + "…";
    P(tag + " -> OMaths=" + mc + " InlineShapes=" + ish + " 文本=" + txt);
    return mc;
  }

  cellStat("空表 左单元格", tbl.Cell(1, 1));
  cellStat("空表 右单元格", tbl.Cell(1, 2));

  // 路径 A：FormattedText
  var okA = false;
  try {
    tbl.Cell(1, 1).Range.FormattedText = om.Range.FormattedText;
    okA = true;
    P("A FormattedText 赋值完成");
  } catch (e5) { P("A FormattedText 失败：" + err(e5)); }
  if (okA) cellStat("A 后 左单元格", tbl.Cell(1, 1));

  // 路径 B：复制粘贴（A 无效时）
  var mcA = -1;
  try { mcA = Number(tbl.Cell(1, 1).Range.OMaths.Count); } catch (e6) {}
  if (!(mcA > 0)) {
    try {
      om.Range.Copy();
      tbl.Cell(1, 1).Range.Paste();
      P("B Copy/Paste 完成");
    } catch (e7) { P("B Copy/Paste 失败：" + err(e7)); }
    cellStat("B 后 左单元格", tbl.Cell(1, 1));
  }

  // 右单元格写纯文本编号，验证不会被吸进数学区
  try { tbl.Cell(1, 2).Range.Text = "(1.1)"; } catch (e8) { P("写编号失败：" + err(e8)); }
  var mcR = cellStat("写号后 右单元格", tbl.Cell(1, 2));

  // 书签范围自检
  try {
    doc.Bookmarks.Add("TBLTEST_NUM", tbl.Cell(1, 2).Range);
    var b = doc.Bookmarks.Item("TBLTEST_NUM");
    P("书签自检 文本=" + String(b.Range.Text).replace(/[\r\n\u0007]/g, "\\n") +
      " 公式数=" + Number(b.Range.OMaths.Count));
  } catch (e9) { P("书签自检失败：" + err(e9)); }

  P("");
  P("结论：" + (mcR === 0 ? "右单元格无公式，方案可行" : "右单元格仍含公式，方案不可行") +
    "；左单元格公式数需 >0");

  alert("搬公式探针完成（请勿保存，或保存到副本）\r\r" + lines.join("\r"));
}
