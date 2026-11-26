/*
 * XML 搬运探针（会改文档，请用副本）
 *
 * 上一轮发现 WPS 把 Range.XML 实现成方法（打印出来是 function 源码），
 * 本轮改为“调用”它，并配合 InsertXML 把公式插进空表左栏，
 * 右栏保持表格自带的普通段落（避免 Split 克隆出数学段落属性）。
 *
 * 运行：光标放进独占段落的显示公式，运行 EQNUM_XMLMOVE_PROBE()
 */

function EQNUM_XMLMOVE_PROBE() {
  var doc = null;
  try { doc = Application.ActiveDocument; } catch (e0) {}
  if (!doc) { alert("请先打开文档"); return; }

  var lines = [];
  function P(s) { lines.push(s); try { console.log("[XMLMOVE] " + s); } catch (e) {} }
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

  // 1) Range.XML：判断是属性还是方法
  var xml = null;
  var kind = "";
  try {
    var x = om.Range.XML;
    if (typeof x === "function") {
      kind = "方法";
      xml = String(x.call(om.Range));
    } else {
      kind = typeof x;
      xml = String(x);
    }
  } catch (e3) {
    P("E Range.XML 调用失败：" + err(e3));
  }
  if (xml === null) { alert(lines.join("\r")); return; }
  P("1 Range.XML 类型=" + kind + " 长度=" + xml.length);
  P("  片段：" + xml.substring(0, 100));
  var looksXml = /<\s*[a-zA-Z]/.test(xml);
  P("  是否像 XML：" + (looksXml ? "是" : "否"));

  // 2) WordOpenXML（如果也实现成方法，同样调用）
  var wv = null;
  try {
    var w = om.Range.WordOpenXML;
    wv = (typeof w === "function") ? String(w.call(om.Range)) : String(w);
    P("2 WordOpenXML 长度=" + wv.length + " 片段：" + wv.substring(0, 60));
  } catch (e4) {
    P("2 WordOpenXML 不可用：" + err(e4));
  }

  if (!looksXml && !wv) {
    P("");
    P("结论：拿不到 XML 字符串，改用其他方案");
    alert(lines.join("\r"));
    return;
  }

  // 3) 建空表（两栏都是普通段落），把公式片段插进左栏
  var pe = -1;
  try { pe = Number(p.Range.End); } catch (e5) {}
  var tbl = null;
  try { tbl = doc.Tables.Add(doc.Range(pe, pe), 1, 2); }
  catch (e6) { P("E Tables.Add 失败：" + err(e6)); alert(lines.join("\r")); return; }
  P("3 空表已建");

  // 先试范围片段 WordOpenXML（整文档 XML 已被证明会被忽略）
  var used = "";
  var t = Date.now();
  if (wv) {
    try { tbl.Cell(1, 1).Range.InsertXML(wv); used = "WordOpenXML"; }
    catch (e7) { P("E InsertXML(WordOpenXML) 失败：" + err(e7)); }
  }
  P("4 InsertXML(" + (used || "无") + ") 完成，用时 " + (Date.now() - t) + "ms");

  var mcL = -1, txtL = "?";
  try { mcL = Number(tbl.Cell(1, 1).Range.OMaths.Count); } catch (e8) {}
  try { txtL = String(tbl.Cell(1, 1).Range.Text).replace(/[\r\n\u0007]/g, "\\n"); } catch (e9) {}
  if (txtL.length > 30) txtL = txtL.substring(0, 30) + "…";
  P("5 左栏 公式数=" + mcL + " 文本=" + txtL);

  if (!(mcL > 0) && looksXml) {
    var t2 = Date.now();
    try { tbl.Cell(1, 1).Range.InsertXML(xml); used = "Range.XML"; }
    catch (e7b) { P("E InsertXML(Range.XML) 失败：" + err(e7b)); }
    try { mcL = Number(tbl.Cell(1, 1).Range.OMaths.Count); } catch (e8b) {}
    P("5b 改用整文档 XML 再试：公式数=" + mcL + "，用时 " + (Date.now() - t2) + "ms");
  }

  // 4) 右栏写编号并检查是否被动成数学区
  try { tbl.Cell(1, 2).Range.Text = "(1.1)"; } catch (e10) { P("写编号失败：" + err(e10)); }
  var mcR = -1, txtR = "?";
  try { mcR = Number(tbl.Cell(1, 2).Range.OMaths.Count); } catch (e11) {}
  try { txtR = String(tbl.Cell(1, 2).Range.Text).replace(/[\r\n\u0007]/g, "\\n"); } catch (e12) {}
  P("6 右栏 公式数=" + mcR + " 文本=" + txtR);

  // 5) 尝试用 WordOpenXML 检查右栏是否含数学段落属性
  try {
    var rx = tbl.Cell(1, 2).Range.WordOpenXML;
    var rv = (typeof rx === "function") ? String(rx.call(tbl.Cell(1, 2).Range)) : String(rx);
    P("7 右栏 XML 含 m:rPr：" + (rv.indexOf("<m:rPr") >= 0 ? "是（数学段）" : "否"));
    P("  右栏 XML 含 w:oMath：" + (rv.indexOf("w:oMath") >= 0 ? "是" : "否"));
  } catch (e13) {
    P("7 右栏 XML 不可读：" + err(e13));
  }

  P("");
  P("结论：" + (mcL > 0 ? "公式已用 InsertXML 搬入左栏；请看第 6、7 行判断右栏是否干净" : "InsertXML 没搬进公式"));
  alert("XML 搬运探针完成\r\r" + lines.join("\r"));
}
