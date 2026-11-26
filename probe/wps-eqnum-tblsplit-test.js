/*
 * “段落转 1x1 表 + 单元格拆分”探针（会改文档，请用副本）
 *
 * 思路：不插制表符（避免把数学区切开），
 *       用 ConvertToTable(0)（按段落分隔）把公式段变成 1x1 表，
 *       再把唯一单元格拆成两列，公式自然留在左栏。
 *
 * 运行：光标放进独占段落的显示公式，运行 EQNUM_TBLSPLIT_TEST()
 */

function EQNUM_TBLSPLIT_TEST() {
  var doc = null;
  try { doc = Application.ActiveDocument; } catch (e0) {}
  if (!doc) { alert("请先打开文档"); return; }

  var lines = [];
  function P(s) { lines.push(s); try { console.log("[TBLSPLIT] " + s); } catch (e) {} }
  function err(e) {
    if (!e) return "未知";
    try { if (e.message) return String(e.message); } catch (e1) {}
    try { if (e.description) return String(e.description); } catch (e2) {}
    return "无法读取";
  }

  var p = null;
  try { p = Application.Selection.Range.Paragraphs.Item(1); } catch (e1) {}
  if (!p) { P("E 取不到段落"); alert(lines.join("\r")); return; }

  var om = null;
  try { om = p.Range.OMaths.Item(1); } catch (e2) {}
  if (!om) { P("E 段落内没有公式"); alert(lines.join("\r")); return; }
  P("0 定位到公式段");

  // 1) 按段落分隔转 1x1 表（不插制表符）
  var tbl = null;
  try { tbl = p.Range.ConvertToTable(0); }
  catch (e3) { P("E ConvertToTable(0) 失败：" + err(e3)); alert(lines.join("\r")); return; }

  var cells = -1;
  try { cells = Number(tbl.Range.Cells.Count); } catch (e4) {}
  var rows = -1;
  try { rows = Number(tbl.Rows.Count); } catch (e5) {}
  var mc1 = -1;
  try { mc1 = Number(tbl.Cell(1, 1).Range.OMaths.Count); } catch (e6) {}
  P("1 转表完成：行=" + rows + " 单元格=" + cells + " 左栏公式数=" + mc1);
  if (cells !== 1) {
    P("W 不是 1x1 表，后续拆分可能不符合预期，仍继续尝试");
  }

  // 2) 拆分单元格为 1 行 2 列
  var t2 = Date.now();
  try { tbl.Cell(1, 1).Split(1, 2); }
  catch (e7) { P("E Cell.Split 失败：" + err(e7)); alert(lines.join("\r")); return; }
  P("2 拆分完成，用时 " + (Date.now() - t2) + "ms");

  var cells2 = -1;
  try { cells2 = Number(tbl.Range.Cells.Count); } catch (e8) {}
  var mL = -1, mR = -1;
  try { mL = Number(tbl.Cell(1, 1).Range.OMaths.Count); } catch (e9) {}
  try { mR = Number(tbl.Cell(1, 2).Range.OMaths.Count); } catch (e10) {}
  P("3 拆分后：单元格=" + cells2 + " 左栏公式数=" + mL + " 右栏公式数=" + mR);

  // 4) 右栏写纯文本编号并校验
  try { tbl.Cell(1, 2).Range.Text = "(1.1)"; } catch (e11) { P("写编号失败：" + err(e11)); }
  var mR2 = -1, txtR = "?";
  try { mR2 = Number(tbl.Cell(1, 2).Range.OMaths.Count); } catch (e12) {}
  try { txtR = String(tbl.Cell(1, 2).Range.Text).replace(/[\r\n\u0007]/g, "\\n"); } catch (e13) {}
  P("4 右栏写入后：公式数=" + mR2 + " 文本=" + txtR);

  var ok = (cells2 === 2 && mL > 0 && mR2 === 0);
  P("");
  P("结论：" + (ok
    ? "可行——不插制表符、不用剪贴板"
    : "不可行（单元格=" + cells2 + "，左栏公式=" + mL + "，右栏公式=" + mR2 + "）"));

  alert("段落转表+拆分探针完成（请用副本，勿保存原文档）\r\r" + lines.join("\r"));
}
