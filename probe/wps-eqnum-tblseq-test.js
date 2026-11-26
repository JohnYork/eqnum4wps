/*
 * 建表全流程逐步骤探针（会改文档，务必用副本）
 *
 * 目的：定位 v23 新流程里哪一步让 WPS 卡死。宏编辑器里 console.log 实时输出，
 *       卡住时最后一行就是卡住的位置。
 *
 * 步骤：找公式段 → 裁剪空白 → 建空表 → 搬公式 → 删原段落
 * 运行：光标放进独占段落的显示公式，运行 EQNUM_TBLSEQ_TEST()
 */

function EQNUM_TBLSEQ_TEST() {
  var doc = null;
  try { doc = Application.ActiveDocument; } catch (e0) {}
  if (!doc) { alert("请先打开文档"); return; }

  function P(s) {
    try { console.log("[TBLSEQ] " + s); } catch (e) {}
    try { Application.StatusBar = "[TBLSEQ] " + s; } catch (e2) {}
  }

  var t0 = Date.now();
  P("0 开始");

  var p = null;
  try { p = Application.Selection.Range.Paragraphs.Item(1); } catch (e1) {}
  if (!p) { P("E 取不到段落"); return; }

  var om = null;
  try { om = p.Range.OMaths.Item(1); } catch (e2) {}
  if (!om) { P("E 段落内没有公式"); return; }
  P("1 已定位公式段，用时 " + (Date.now() - t0) + "ms");

  var pStart = -1, pEnd = -1, omStart = -1, omEnd = -1;
  try {
    pStart = Number(p.Range.Start); pEnd = Number(p.Range.End);
    omStart = Number(om.Range.Start); omEnd = Number(om.Range.End);
  } catch (e3) { P("E 读取位置失败"); return; }
  P("2 位置 p=" + pStart + ".." + pEnd + " om=" + omStart + ".." + omEnd);

  // 裁剪公式前后空白（与主流程一致）
  if (omStart > pStart) {
    try { doc.Range(pStart, omStart).Delete(); } catch (e4) { P("E 裁剪前部失败"); }
    P("3 已裁剪前部空白");
  } else {
    P("3 前部无空白，跳过");
  }
  var shift = Math.max(0, omStart - pStart);
  var pEnd2 = pEnd - shift;
  var omEnd2 = omEnd - shift;
  if (pEnd2 - 1 > omEnd2) {
    try { doc.Range(omEnd2, pEnd2 - 1).Delete(); } catch (e5) { P("E 裁剪尾部失败"); }
    P("4 已裁剪尾部空白");
  } else {
    P("4 尾部无空白，跳过");
  }

  var t1 = Date.now();
  var pe = -1;
  try { pe = Number(p.Range.End); } catch (e6) { P("E 读取段尾失败"); return; }
  P("5 段尾位置 " + pe);

  var tbl = null;
  try { tbl = doc.Tables.Add(doc.Range(pe, pe), 1, 2); }
  catch (e7) { P("E Tables.Add 失败：" + (e7 && e7.message)); return; }
  P("6 空表已建，用时 " + (Date.now() - t1) + "ms");

  var t2 = Date.now();
  try { tbl.Cell(1, 1).Range.FormattedText = om.Range.FormattedText; }
  catch (e8) { P("E FormattedText 失败：" + (e8 && e8.message)); return; }
  var mc = -1;
  try { mc = Number(tbl.Cell(1, 1).Range.OMaths.Count); } catch (e9) {}
  P("7 公式已搬入左栏 OMaths=" + mc + "，用时 " + (Date.now() - t2) + "ms");

  var t3 = Date.now();
  try { p.Range.Delete(); }
  catch (e10) { P("E 删除原段落失败：" + (e10 && e10.message)); return; }
  P("8 原公式段落已删除，用时 " + (Date.now() - t3) + "ms");

  P("9 全部完成，总用时 " + (Date.now() - t0) + "ms");
  alert("建表流程完成，未卡死。\r\r请查看 [TBLSEQ] 各步耗时。\r\r" +
    "若目视发现公式或正文异常，请勿保存，直接关闭副本。");
}
