/*
 * 删号开销拆解探针（会修改文档，请用副本）
 *
 * 目的：把“删旧号”拆成三段分别计时，确认瓶颈是
 *   1) doc.Paragraphs.Item(序号) 按序号取段落
 *   2) 读段落文本
 *   3) Range.Delete()（分“删整段”和“删跨段范围”两种）
 *
 * 运行：光标放到第 1 章任意位置，运行 EQNUM_DELETE_COST_TEST()
 * 全程 ScreenUpdating=false，与正式流程一致。
 */

function EQNUM_DELETE_COST_TEST() {
  var doc = null;
  try { doc = Application.ActiveDocument; } catch (e0) {}
  if (!doc) { alert("请先打开文档"); return; }

  var n = 6;
  if (typeof InputBox === "function") {
    var ans = InputBox("测试几个公式？建议 6（前 3 个删整段，后 3 个删跨段范围）", "EQNUM 删号拆解", "6");
    if (ans === null || ans === undefined) return;
    var num = Number(String(ans).replace(/\s/g, ""));
    if (num > 0) n = Math.floor(num);
  }
  if (n < 4) n = 4;
  if (n > 12) n = 12;

  var curPara = dctCursorParagraph(doc);
  var start = dctChapterStart(doc, curPara);
  var end = dctChapterEnd(doc, start);

  var cands = [];
  for (var i = end; i >= start && cands.length < n; i--) {
    var p = null;
    try { p = doc.Paragraphs.Item(i); } catch (e1) { continue; }
    var c = -1;
    try { c = Number(p.Range.OMaths.Count); } catch (e2) { continue; }
    if (c !== 1) continue;

    var np = null;
    try { np = doc.Paragraphs.Item(i + 1); } catch (e3) { continue; }
    var nt = "";
    try { nt = String(np.Range.Text); } catch (e4) {}
    if (!/^[\s]*[\(（]\s*\d+(?:\.\d+)*\s*[\)）][\s]*$/.test(nt)) continue;

    cands.push(i);
  }

  if (cands.length === 0) {
    alert("当前章节没有找到候选");
    return;
  }

  var oldScreen = true;
  try { oldScreen = !!Application.ScreenUpdating; } catch (eS) {}
  try { Application.ScreenUpdating = false; } catch (eS2) {}

  var lines = [];
  var sumLook = 0, sumText = 0, sumDelWhole = 0, sumDelRange = 0;
  var cntWhole = 0, cntRange = 0;

  for (var k = 0; k < cands.length; k++) {
    var pi = cands[k];
    var useRange = (k % 2 === 1); // 交替：偶数删整段，奇数删跨段范围

    var t1 = Date.now();
    var np = null;
    try { np = doc.Paragraphs.Item(pi + 1); } catch (e5) {}
    var dLook = Date.now() - t1;

    var t2 = Date.now();
    var txt = "";
    try { txt = String(np.Range.Text); } catch (e6) {}
    var dText = Date.now() - t2;

    var dDel = 0;
    var ok = false;

    if (useRange) {
      var t3 = Date.now();
      try {
        var para = doc.Paragraphs.Item(pi);
        var om = para.Range.OMaths.Item(1);
        var omEnd = Number(om.Range.End);
        var npEnd = Number(np.Range.End) - 1;
        doc.Range(omEnd, npEnd).Delete();
        ok = true;
      } catch (e7) {}
      dDel = Date.now() - t3;
      sumDelRange += dDel; cntRange++;
    } else {
      var t4 = Date.now();
      try {
        np.Range.Delete();
        ok = true;
      } catch (e8) {}
      dDel = Date.now() - t4;
      sumDelWhole += dDel; cntWhole++;
    }

    sumLook += dLook;
    sumText += dText;

    var line = "第 " + (k + 1) + "/" + cands.length + " 段#" + pi +
      (useRange ? " [删跨段范围]" : " [删整段]") +
      " 取段落 " + dLook + "ms，读文本 " + dText + "ms，删除 " + dDel + "ms" +
      (ok ? "" : " 删除失败");
    lines.push(line);
    try { console.log("[DELCOST] " + line); } catch (e9) {}
  }

  try { Application.ScreenUpdating = oldScreen; } catch (eA) {}
  try { Application.StatusBar = ""; } catch (eB) {}

  var cnt = cands.length;
  var summary = "删号拆解完成，共 " + cnt + " 个（ScreenUpdating=false）\r\r" +
    "平均：取段落 " + Math.round(sumLook / cnt) + "ms，读文本 " + Math.round(sumText / cnt) + "ms\r" +
    "删整段平均：" + (cntWhole ? Math.round(sumDelWhole / cntWhole) : "-") + "ms（" + cntWhole + " 个）\r" +
    "删跨段范围平均：" + (cntRange ? Math.round(sumDelRange / cntRange) : "-") + "ms（" + cntRange + " 个）\r\r" +
    lines.join("\r");

  alert(summary);
}

function dctCursorParagraph(doc) {
  var start = 0;
  try { start = Number(Application.Selection.Range.Start); } catch (e) {}
  try { return Number(doc.Range(0, start).Paragraphs.Count) + 1; } catch (e2) { return 1; }
}

function dctChapterStart(doc, curPara) {
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

function dctChapterEnd(doc, start) {
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
