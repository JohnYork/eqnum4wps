/*
 * InsertXML 搬公式探针（会改文档，请用副本）
 *
 * 目的：验证能否用 Range.XML + Range.InsertXML 搬运公式，
 *       从而完全不占用剪贴板。
 *
 * 运行：光标放进独占段落的显示公式，运行 EQNUM_INSERTXML_TEST()
 */

function EQNUM_INSERTXML_TEST() {
  var doc = null;
  try { doc = Application.ActiveDocument; } catch (e0) {}
  if (!doc) { alert("请先打开文档"); return; }

  var lines = [];
  function P(s) { lines.push(s); try { console.log("[INSXML] " + s); } catch (e) {} }
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

  // 1) 读取 XML
  var xml = null;
  try { xml = String(om.Range.XML); }
  catch (e3) { P("E Range.XML 不可用：" + err(e3)); alert(lines.join("\r")); return; }
  P("1 Range.XML 读取成功，长度=" + xml.length);
  P("  片段：" + xml.substring(0, 120));

  // 2) 建空表
  var pe = -1;
  try { pe = Number(p.Range.End); } catch (e4) {}
  var tbl = null;
  try { tbl = doc.Tables.Add(doc.Range(pe, pe), 1, 2); }
  catch (e5) { P("E Tables.Add 失败：" + err(e5)); alert(lines.join("\r")); return; }
  P("2 空表已建");

  // 3) InsertXML 到左栏
  var t = Date.now();
  try { tbl.Cell(1, 1).Range.InsertXML(xml); }
  catch (e6) { P("E InsertXML 失败：" + err(e6)); alert(lines.join("\r")); return; }
  P("3 InsertXML 完成，用时 " + (Date.now() - t) + "ms");

  // 4) 校验
  var mc = -1, txt = "?";
  try { mc = Number(tbl.Cell(1, 1).Range.OMaths.Count); } catch (e7) {}
  try { txt = String(tbl.Cell(1, 1).Range.Text).replace(/[\r\n\u0007]/g, "\\n"); } catch (e8) {}
  P("4 左栏 OMaths=" + mc + " 文本=" + (txt.length > 40 ? txt.substring(0, 40) + "…" : txt));

  // 5) 右栏写纯文本编号
  try { tbl.Cell(1, 2).Range.Text = "(1.1)"; } catch (e9) { P("写编号失败：" + err(e9)); }
  var mr = -1;
  try { mr = Number(tbl.Cell(1, 2).Range.OMaths.Count); } catch (e10) {}
  P("5 右栏 OMaths=" + mr);

  P("");
  P("结论：" + (mc > 0 && mr === 0
    ? "InsertXML 可用，可替代剪贴板"
    : "InsertXML 结果不符合预期，需继续用剪贴板"));

  alert("InsertXML 探针完成（请勿保存原文档，用副本）\r\r" + lines.join("\r"));
}
