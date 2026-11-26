/*
 * 编号单元格“裸 m:r”清理探针（会改文档，请用副本）
 *
 * 背景：编号单元格里没有 m:oMath 容器，只有 ConvertToTable 切分留下的
 *       裸 m:r 数学运行，导致 OMaths.Count=1 但 Range.Delete 无效，
 *       编号被写进数学区，引用显示为对象。
 *
 * 运行：打开一份“已被本工具编号过、且编号单元格 OMaths.Count>0”的副本，
 *       运行 EQNUM_CELLFIX_PROBE()，把输出发回。
 */

function EQNUM_CELLFIX_PROBE() {
  var doc = null;
  try { doc = Application.ActiveDocument; } catch (e0) {}
  if (!doc) { alert("请先打开文档"); return; }

  var lines = [];
  function P(s) { lines.push(s); try { console.log("[CELLFIX] " + s); } catch (e) {} }
  function err(e) {
    if (!e) return "未知";
    try { if (e.message) return String(e.message); } catch (e1) {}
    try { if (e.description) return String(e.description); } catch (e2) {}
    return "无法读取";
  }

  // 找到含 EQNUM_NUM_ 书签的表格
  var target = null;
  var tn = 0;
  try { tn = Number(doc.Tables.Count); } catch (eT) {}
  for (var i = 1; i <= tn; i++) {
    var tb = null;
    try { tb = doc.Tables.Item(i); } catch (e1) { continue; }
    var cell = null;
    try { cell = tb.Cell(1, 2); } catch (e2) { continue; }
    if (String(cell.Range.Text).indexOf(")") >= 0 || Number(cell.Range.OMaths.Count) > 0) {
      target = tb;
      break;
    }
  }
  if (!target) { alert("没找到两栏公式表格，请确认文档已被本工具编号过"); return; }
  P("目标表格 #" + i);

  var c = target.Cell(1, 2);

  function stat(tag) {
    var cnt = -1, txt = "?";
    try { cnt = Number(c.Range.OMaths.Count); } catch (e) {}
    try { txt = String(c.Range.Text).replace(/[\r\n]/g, "\\n"); } catch (e2) {}
    if (txt.length > 30) txt = txt.substring(0, 30) + "…";
    P(tag + " -> OMaths=" + cnt + " 文本=" + txt + " 内联=" + inlineCount(c));
  }

  function inlineCount(cell) {
    var n = -1;
    try { n = Number(cell.Range.InlineShapes.Count); } catch (e) { return "ERR"; }
    return n;
  }

  stat("初始");

  // S1: OMath.Range.Delete
  try { c.Range.OMaths.Item(1).Range.Delete(); } catch (e) { P("S1 ERR " + err(e)); }
  stat("S1 OMath.Range.Delete");

  // S2: OMath.Range.Text = ""
  try { c.Range.OMaths.Item(1).Range.Text = ""; } catch (e) { P("S2 ERR " + err(e)); }
  stat("S2 OMath.Range.Text=''");

  // S3: 整段内容删除（排除单元格标记）
  try {
    var r = c.Range;
    r.MoveEnd(1, -1);
    r.Delete();
  } catch (e) { P("S3 ERR " + err(e)); }
  stat("S3 Range.Delete");

  // S4: 新增一段后删首段
  try {
    c.Range.InsertParagraphAfter();
    var ps = c.Range.Paragraphs;
    P("  S4 段落数=" + Number(ps.Count));
    ps.Item(1).Range.Delete();
  } catch (e) { P("S4 ERR " + err(e)); }
  stat("S4 新增段+删首段");

  // S5: 写占位再清空
  try { c.Range.Text = "x"; } catch (e) { P("S5a ERR " + err(e)); }
  try { c.Range.Text = ""; } catch (e) { P("S5b ERR " + err(e)); }
  stat("S5 Text='x' 再清空");

  // S6: 段落样式改正文后再删
  try { c.Range.Paragraphs.Item(1).Style = "正文"; } catch (e) { P("S6a ERR " + err(e)); }
  try {
    var r6 = c.Range;
    r6.MoveEnd(1, -1);
    r6.Delete();
  } catch (e) { P("S6b ERR " + err(e)); }
  stat("S6 样式正文+删除");

  alert("编号单元格清理探针完成\r\r" + lines.join("\r"));
}
