/*
 * “隐藏空段”方案探针（会修改文档，请用副本）
 *
 * 入口：
 *   EQNUM_HIDEPARA_TEST()   隐藏空段原型计时（删文字 + 设隐藏 + 制表位 + 写号）
 *   EQNUM_HIDEBATCH_TEST()  “先隐藏加书签、编号后统一删除”方案验证
 *
 * EQNUM_HIDEBATCH_TEST 分两段，各取一半候选：
 *   阶段 B：隐藏旧编号段 + 加校验书签 → 在该范围内用 Find/Replace（格式=隐藏）
 *           尝试一次批量删除，看书签是否全部消失、耗时多少
 *   阶段 A：隐藏旧编号段 + 加 HPROBE_DEL_ 书签 → 立刻按书签逐条删除，记单条耗时
 *   判定：阶段 A 单条远低于 1 秒才值得改主流程；阶段 B 若可行则更省
 *
 * 运行：光标放第 1 章，运行入口宏，全程 ScreenUpdating=false。
 * 注意：阶段 B 的批量替换会删掉该范围内的所有隐藏文字（务必用副本）。
 */

function EQNUM_HIDEPARA_TEST() {
  var doc = null;
  try { doc = Application.ActiveDocument; } catch (e0) {}
  if (!doc) { alert("请先打开文档"); return; }

  var n = 6;
  if (typeof InputBox === "function") {
    var ans = InputBox("测试几个公式？建议 6", "EQNUM 隐藏空段原型", "6");
    if (ans === null || ans === undefined) return;
    var num = Number(String(ans).replace(/\s/g, ""));
    if (num > 0) n = Math.floor(num);
  }
  if (n < 2) n = 2;
  if (n > 12) n = 12;

  var curPara = hpCursorParagraph(doc);
  var start = hpChapterStart(doc, curPara);
  var end = hpChapterEnd(doc, start);

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

  var right = hpTextWidth(doc);
  var lines = [];
  var sDel = 0, sHide = 0, sTab = 0, sWrite = 0;
  var hiddenOk = 0;

  for (var k = 0; k < cands.length; k++) {
    var pi = cands[k];
    var tItem = Date.now();

    var para = null;
    try { para = doc.Paragraphs.Item(pi); } catch (e5) { continue; }
    var numPara = null;
    try { numPara = doc.Paragraphs.Item(pi + 1); } catch (e6) { continue; }

    // 1) 只删编号文字（保留段落标记）
    var t1 = Date.now();
    var delOk = false;
    try {
      var ns = Number(numPara.Range.Start);
      var ne = Number(numPara.Range.End) - 1;
      if (ne > ns) doc.Range(ns, ne).Delete();
      delOk = true;
    } catch (e7) {}
    var dDel = Date.now() - t1;

    // 2) 把空段落设为隐藏文字
    var t2 = Date.now();
    var hideOk = false;
    try {
      numPara.Range.Font.Hidden = true;
      hideOk = !!numPara.Range.Font.Hidden;
      if (hideOk) hiddenOk++;
    } catch (e8) {}
    var dHide = Date.now() - t2;

    // 3) 公式后插制表符 + 右对齐制表位
    var t3 = Date.now();
    var insOk = false;
    try { para.Range.ParagraphFormat.TabStops.Add(right, 2, 0); } catch (e9) {}
    try {
      var om = para.Range.OMaths.Item(1);
      var omEnd = Number(om.Range.End);
      doc.Range(omEnd, omEnd).InsertAfter("\t");
      insOk = true;
    } catch (e10) {}
    var dTab = Date.now() - t3;

    // 4) 写编号字段
    var t4 = Date.now();
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
    } catch (e11) {}
    var dWrite = Date.now() - t4;

    sDel += dDel; sHide += dHide; sTab += dTab; sWrite += dWrite;

    var line = "第 " + (k + 1) + "/" + cands.length + " 段#" + pi +
      " 用时 " + (Date.now() - tItem) + "ms（删文字 " + dDel +
      "，设隐藏 " + dHide + "，制表位+插入 " + dTab + "，写号 " + dWrite + "）" +
      (delOk ? "" : " 删文字失败") + (hideOk ? "" : " 隐藏未生效") +
      (insOk ? "" : " 插入失败") + (writeOk ? "" : " 写号失败");
    lines.push(line);
    try { console.log("[HIDEPARA] " + line); } catch (e12) {}
  }

  try { Application.ScreenUpdating = oldScreen; } catch (eA) {}
  try { Application.StatusBar = ""; } catch (eB) {}

  var cnt = cands.length;
  var summary = "隐藏空段原型完成，共 " + cnt + " 个（ScreenUpdating=false）\r\r" +
    "平均：删文字 " + Math.round(sDel / cnt) + "ms，设隐藏 " + Math.round(sHide / cnt) +
    "ms，制表位+插入 " + Math.round(sTab / cnt) + "ms，写号 " + Math.round(sWrite / cnt) + "ms\r" +
    "隐藏生效：" + hiddenOk + " / " + cnt + "\r\r" +
    lines.join("\r");

  alert(summary);
}

function hpErr(e) {
  if (!e) return "未知错误";
  try { if (e.message) return String(e.message); } catch (e1) {}
  try { if (e.description) return String(e.description); } catch (e2) {}
  try { return String(e); } catch (e3) { return "无法读取的错误"; }
}

function hpAddBookmark(doc, name, rng) {
  try { if (doc.Bookmarks.Exists(name)) doc.Bookmarks.Item(name).Delete(); } catch (e0) {}
  try { doc.Bookmarks.Add(name, rng); return true; } catch (e1) { return false; }
}

function hpDropBookmark(doc, name) {
  try { if (doc.Bookmarks.Exists(name)) doc.Bookmarks.Item(name).Delete(); } catch (e0) {}
}

// 在指定范围内尝试“格式=隐藏”的批量替换，删除隐藏文字
function hpTryHiddenReplace(doc, rng) {
  var res = { ok: false, elapsed: 0, how: "", error: "" };

  var find = null;
  try { find = rng.Find; } catch (e0) {
    res.error = "拿不到 Range.Find";
    return res;
  }
  if (!find) {
    res.error = "Range.Find 为空";
    return res;
  }

  try { find.ClearFormatting(); } catch (e1) {}
  try { find.Replacement.ClearFormatting(); } catch (e2) {}
  try {
    find.Text = "";
    find.Replacement.Text = "";
  } catch (e3) {}

  try { find.Font.Hidden = true; } catch (e4) {
    res.error = "Find.Font.Hidden 不支持";
    return res;
  }
  try { find.Format = true; } catch (e5) {}

  var t0 = Date.now();
  var ok = false;

  try {
    ok = !!find.Execute();
    res.how = "Execute()";
  } catch (e6) {
    res.error = "Execute() 失败：" + hpErr(e6);
  }

  // 属性式调用不通时，退回 Word 的 15 参数式
  if (!ok) {
    try {
      ok = !!find.Execute("", false, false, false, false, false, true, 1, true, "", 2);
      res.how = "Execute(15 参数)";
      res.error = "";
    } catch (e7) {
      res.error = (res.error ? res.error + "；" : "") + "参数式 Execute 失败：" + hpErr(e7);
    }
  }

  res.elapsed = Date.now() - t0;
  res.ok = ok;
  return res;
}

function EQNUM_HIDEBATCH_TEST() {
  var doc = null;
  try { doc = Application.ActiveDocument; } catch (e0) {}
  if (!doc) { alert("请先打开文档"); return; }

  var n = 6;
  if (typeof InputBox === "function") {
    var ans = InputBox("每段测试几个公式？建议 6（共取 2 倍候选，前后各一半）",
      "EQNUM 隐藏+书签统一删除", "6");
    if (ans === null || ans === undefined) return;
    var num = Number(String(ans).replace(/\s/g, ""));
    if (num > 0) n = Math.floor(num);
  }
  if (n < 2) n = 2;
  if (n > 10) n = 10;

  var curPara = hpCursorParagraph(doc);
  var start = hpChapterStart(doc, curPara);
  var end = hpChapterEnd(doc, start);

  // 从章末向前取 2n 个候选（公式段 + 紧跟的纯编号段）
  var cands = [];
  for (var i = end; i >= start && cands.length < n * 2; i--) {
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

  if (cands.length < 4) {
    alert("当前章节候选太少（" + cands.length + " 个），换一章再试");
    return;
  }

  // cands 按段号降序：前一半靠后（阶段 B），后一半靠前（阶段 A）
  var half = Math.floor(cands.length / 2);
  var bList = cands.slice(0, half);
  var aList = cands.slice(half);

  var oldScreen = true;
  try { oldScreen = !!Application.ScreenUpdating; } catch (eS) {}
  try { Application.ScreenUpdating = false; } catch (eS2) {}

  var lines = [];
  var chkNames = [];
  var delNames = [];

  /* ---------- 阶段 B：隐藏 + Find/Replace 批量删除 ---------- */
  var bHidden = 0;
  for (var bi = 0; bi < bList.length; bi++) {
    var bNum = null;
    try { bNum = doc.Paragraphs.Item(bList[bi] + 1); } catch (eB1) { continue; }
    var bOk = false;
    try {
      bNum.Range.Font.Hidden = true;
      bOk = !!bNum.Range.Font.Hidden;
    } catch (eB2) {}
    if (!bOk) continue;

    var cname = "HPROBE_CHK_" + (bHidden + 1);
    if (hpAddBookmark(doc, cname, bNum.Range)) {
      chkNames.push(cname);
      bHidden++;
    }
  }

  var bRange = null;
  try {
    var rs = Number(doc.Paragraphs.Item(bList[bList.length - 1]).Range.Start);
    var re = Number(doc.Paragraphs.Item(bList[0] + 1).Range.End);
    if (re > rs) bRange = doc.Range(rs, re);
  } catch (eB3) {}

  var bRes = { ok: false, elapsed: 0, how: "", error: "无可用范围" };
  if (bRange) bRes = hpTryHiddenReplace(doc, bRange);

  var alive = [];
  for (var ci = 0; ci < chkNames.length; ci++) {
    var exists = true;
    try { exists = !!doc.Bookmarks.Exists(chkNames[ci]); } catch (eB4) {}
    if (exists) alive.push(chkNames[ci]);
  }

  lines.push("【阶段B】隐藏 " + bHidden + " 个编号段；批量替换用时 " + bRes.elapsed +
    "ms，返回 " + (bRes.ok ? "true" : "false") +
    (bRes.how ? "（" + bRes.how + "）" : "") +
    (bRes.error ? "，错误：" + bRes.error : ""));
  lines.push("  校验书签剩余 " + alive.length + " / " + chkNames.length +
    (chkNames.length === 0 ? "（无样本）" :
      (alive.length === 0 ? " → 隐藏内容已被批量删除" :
        " → 仍残留：" + alive.slice(0, 3).join("、"))));

  /* ---------- 阶段 A：隐藏 + 书签，逐条删除计时 ---------- */
  var aHidden = 0;
  var aTotal = 0;
  var aCount = 0;

  for (var ai = 0; ai < aList.length; ai++) {
    var pi = aList[ai];
    var aNum = null;
    try { aNum = doc.Paragraphs.Item(pi + 1); } catch (eA1) { continue; }

    var tH = Date.now();
    var aOk = false;
    try {
      aNum.Range.Font.Hidden = true;
      aOk = !!aNum.Range.Font.Hidden;
    } catch (eA2) {}
    var dH = Date.now() - tH;
    if (!aOk) {
      lines.push("  #" + pi + " 隐藏失败，跳过");
      continue;
    }
    aHidden++;

    var dname = "HPROBE_DEL_" + (aHidden);
    if (!hpAddBookmark(doc, dname, aNum.Range)) {
      lines.push("  #" + pi + " 加书签失败，跳过删除计时");
      continue;
    }
    delNames.push(dname);

    var tD = Date.now();
    var dOk = false;
    var dErr = "";
    try {
      var bm = doc.Bookmarks.Item(dname);
      bm.Range.Delete();
      dOk = true;
    } catch (eA3) { dErr = "：" + hpErr(eA3); }
    var dD = Date.now() - tD;

    aTotal += dD;
    aCount++;
    lines.push("  #" + pi + "：隐藏 " + dH + "ms，按书签删段 " + dD + "ms" +
      (dOk ? "" : " 删除失败" + dErr));
  }

  // 清理仍未用到的书签（阶段 B 批量删除失败时留下的校验书签）
  for (var qi = 0; qi < chkNames.length; qi++) hpDropBookmark(doc, chkNames[qi]);
  for (var qj = 0; qj < delNames.length; qj++) hpDropBookmark(doc, delNames[qj]);

  try { Application.ScreenUpdating = oldScreen; } catch (eA) {}
  try { Application.StatusBar = ""; } catch (eB) {}

  var aAvg = aCount > 0 ? Math.round(aTotal / aCount) : -1;
  var verdict;
  if (aAvg >= 0 && aAvg < 200) {
    verdict = "结论：按书签删隐藏段落约 " + aAvg + "ms/条 → 延后统一删除方案可行";
  } else if (aAvg >= 0) {
    verdict = "结论：按书签删隐藏段落约 " + aAvg + "ms/条 → 仍过慢，建议保持永久隐藏";
  } else {
    verdict = "结论：没有完成任何删除计时，无法判断";
  }
  lines.push("");
  lines.push(verdict);
  if (alive.length === 0 && chkNames.length > 0) {
    lines.push("补充：Find/Replace（格式=隐藏）已能批量删除隐藏内容，可优先采用");
  }

  try { console.log("[HIDEBATCH]\r" + lines.join("\r")); } catch (eC) {}

  alert("隐藏 + 书签统一删除验证完成\r\r" + lines.join("\r"));
}

function hpCursorParagraph(doc) {
  var start = 0;
  try { start = Number(Application.Selection.Range.Start); } catch (e) {}
  try { return Number(doc.Range(0, start).Paragraphs.Count) + 1; } catch (e2) { return 1; }
}

function hpChapterStart(doc, curPara) {
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

function hpChapterEnd(doc, start) {
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

function hpTextWidth(doc) {
  try {
    var ps = doc.PageSetup;
    return Number(ps.PageWidth) - Number(ps.LeftMargin) - Number(ps.RightMargin);
  } catch (e) { return 415.3; }
}
