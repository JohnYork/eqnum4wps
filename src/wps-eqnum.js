/*
 * WPS 公式自动编号（题注版）
 *
 * 入口：
 *   EQNUM_CURRENT()     编号光标所在公式（不限 Type）
 *   EQNUM_RENUMBER()    全文重编号并刷新引用（仅公式后已有编号的公式）
 *   EQNUM_REFRESH()     刷新已编号公式（当前章节 / 全文）
 *   EQNUM_SETTINGS()    查看/修改设置
 *   EQNUM_INSERT_REF()  打开交叉引用对话框（失败回退输入框）
 *   EQNUM_DEBUG_SCAN()  只读诊断：打印前 20 个候选公式及旧编号位置
 *
 * 编号实现：
 *   - 右单元格题注字段：按章 ( STYLEREF <层级> \n . SEQ 公式 \s <层级> )
 *   - 连续模式：( SEQ 公式 )
 *   - 原生题注引用选「整项题注」可显示带括号编号
 *   - 保留 EQNUM_NUM_xxxx 书签作为内部标识
 *
 * 范围：
 *   - 编号当前：Type=0 / Type=1 均可；混排会先拆段
 *   - 全文重编号：只处理“公式后面已有编号”的公式；
 *     允许空格、制表符、软换行、硬回车，不跨分页/分节符
 *   - 同段多公式、表格内非本工具公式：跳过
 *   - MathType 公式不处理，保留原有编号
 *
 * 安全：会修改当前文档，执行前请先备份。
 */

function EQNUM_CURRENT() { return EQNUM_MAIN("current"); }
function EQNUM_RENUMBER() { return EQNUM_MAIN("renumber"); }
function EQNUM_REFRESH() { return EQNUM_MAIN("refresh"); }
function EQNUM_SETTINGS() { return EQNUM_MAIN("settings"); }
function EQNUM_INSERT_REF() { return EQNUM_MAIN("insertref"); }
function EQNUM_UPDATE_REFS() { return EQNUM_MAIN("updaterefs"); }
function EQNUM_DEBUG_SCAN() { return EQNUM_MAIN("debugscan"); }

function EQNUM_MAIN(mode) {
var PROGRESS_STEP = 500;
var BM_PREFIX = "EQNUM_NUM_";
var CFG = {
  mode: "chapter",      // continuous | chapter
  level: 1,             // 1 | 2
  start: 1,
  refPrefix: "式",
  refreshRefs: "ours",  // all | ours | none
  sub: 0,               // v1 仅保存，不启用
  subStyle: "alpha"
};

var state = {
  doc: null,
  info: 0,
  warn: 0,
  fail: 0,
  messages: []
};

function consoleLog(s) {
  try {
    if (typeof console !== "undefined" && console && console.log) console.log(s);
  } catch (e) {}
}

function log(level, msg) {
  if (level === "FAIL") state.fail++;
  else if (level === "WARN") state.warn++;
  else state.info++;
  state.messages.push("[" + level + "] " + msg);
  consoleLog("[" + level + "] " + msg);
}

function section(title) {
  state.messages.push("");
  state.messages.push("=== " + title + " ===");
  consoleLog("");
  consoleLog("=== " + title + " ===");
}

function errText(e) {
  if (!e) return "未知错误";
  try { if (e.message) return String(e.message); } catch (e1) {}
  try { if (e.description) return String(e.description); } catch (e2) {}
  try { return String(e); } catch (e3) { return "无法读取的错误"; }
}

function run(fn) {
  try { return { ok: true, value: fn() }; }
  catch (e) { return { ok: false, error: errText(e) }; }
}

// 让出一次线程：加载项页面里 setTimeout 可用，可实时刷新面板并响应中止；
// 宏环境没有 setTimeout 时直接继续，不改变原有同步行为。
function yieldToUi() {
  return new Promise(function (resolve) {
    if (typeof setTimeout === "function") setTimeout(resolve, 0);
    else resolve();
  });
}

function cancelRequested() {
  try {
    if (typeof window !== "undefined" && window && window.EQNUM_CANCEL) return true;
  } catch (e) {}
  return false;
}

// 让 WPS 主线程有机会重绘文档：临时打开屏幕更新并停留一小段时间。
// 间隔由面板设置 window.EQNUM_UI_GAP（毫秒），0 表示不等待。
function eqnumUiGap() {
  var gap = 0;
  try {
    if (typeof window !== "undefined" && window && typeof window.EQNUM_UI_GAP === "number") {
      gap = window.EQNUM_UI_GAP;
    }
  } catch (e0) {}

  if (!(gap > 0) || typeof setTimeout !== "function") return yieldToUi();

  return new Promise(function (resolve) {
    var restored = false;
    try { Application.ScreenUpdating = true; restored = true; } catch (e1) {}
    // 尽量强制重绘一次；接口不存在时静默跳过
    try { Application.ScreenRefresh(); } catch (eR1) {}
    try { Application.ActiveWindow.ScreenRefresh(); } catch (eR2) {}
    setTimeout(function () {
      if (restored) {
        try { Application.ScreenUpdating = false; } catch (e2) {}
      }
      resolve();
    }, gap);
  });
}

// 确定=继续，取消=放弃；不支持 confirm 时用默认值为 YES 的 InputBox
function askConfirm(msg, title) {
  if (typeof confirm === "function") return confirm(msg);

  if (typeof InputBox === "function") {
    var a = InputBox(msg, title || "EQNUM 确认", "YES");
    if (a === null || a === undefined || a === false) return false;
    return String(a).replace(/\s/g, "") !== "";
  }

  return false;
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
  try { t = String(v); } catch (e) { t = "<无法转字符串>"; }
  return esc(clip(t, 120));
}

function status(msg) {
  try { Application.StatusBar = msg; } catch (e) {}
}

// 进度回调：加载项面板注册了 window.EQNUM_PROGRESS 时实时更新进度条；
// 宏环境没有该回调，静默跳过。total <= 0 表示阶段未知。
function reportProgress(done, total, text) {
  try {
    if (typeof window !== "undefined" && window && typeof window.EQNUM_PROGRESS === "function") {
      window.EQNUM_PROGRESS(Number(done) || 0, Number(total) || 0, String(text || ""));
    }
  } catch (e) {}
}

function styleName(p) {
  try { return String(p.Style.NameLocal); } catch (e1) {}
  try { return String(p.Style); } catch (e2) {}
  return "读取失败";
}

function isHeadingLevel(style, level) {
  if (level === 1) return /^(标题|Heading)\s*1$/i.test(style);
  if (level === 2) return /^(标题|Heading)\s*2$/i.test(style);
  return false;
}

function headingNumber(p) {
  var s = "";
  try { s = String(p.Range.ListFormat.ListString); } catch (e1) {}
  s = s.replace(/^\s+|\s+$/g, "");
  if (s) return s;
  var t = "";
  try { t = String(p.Range.Text); } catch (e2) {}
  var m = t.match(/^\s*(\d+(?:\.\d+)*)/);
  return m ? m[1] : "";
}

// 去掉空白、制表位、段落标记、单元格标记和旧编号
function cleanAllowed(s) {
  s = String(s);
  s = s.replace(/[\(（]\s*\d+(?:\s*[-.]\s*\d+)*\s*[\)）]/g, "");
  s = s.replace(/[\s\u0007]/g, "");
  return s;
}

// 只识别“公式后面紧跟的尾部编号”，避免误伤正文引用
function trailingNumber(s) {
  s = String(s);
  s = s.replace(/[\r\u0007]/g, "");
  var m = s.match(/^[\s]*([\(（]\s*\d+(?:\.\d+)*\s*[\)）])[\s]*$/);
  return m ? m[1] : "";
}

function hasPageBreakText(s) {
  return String(s).indexOf("\f") >= 0;
}

// 判断段落是否只剩空白和段落标记
function paragraphIsEmpty(p) {
  var raw = "";
  try { raw = String(p.Range.Text); } catch (e0) { return false; }
  if (raw.indexOf("\f") >= 0) return false;

  // 空白 + 段落/单元格标记 + StyleSeparator 等非打印字符一律忽略
  var clean = raw.replace(/[\r\n\t\v\u00A0\u3000\u0007\u001E\u001F\u2028\u2029 ]/g, "");
  if (clean.length > 0) return false;

  try { if (Number(p.Range.Fields.Count) > 0) return false; } catch (e1) { return false; }
  try { if (Number(p.Range.OMaths.Count) > 0) return false; } catch (e2) { return false; }
  try { if (Number(p.Range.InlineShapes.Count) > 0) return false; } catch (e3) { return false; }
  try { if (Number(p.Range.Tables.Count) > 0) return false; } catch (e4) { return false; }

  return true;
}

// 判断段落是否只有这个编号（允许前后空白）
function paragraphOnlyNumber(p, numberText) {
  var raw = "";
  try { raw = String(p.Range.Text); } catch (e0) { return false; }
  if (raw.indexOf("\f") >= 0) return false;

  var clean = raw.replace(/[\r\n\t\v\u00A0\u3000\u0007\u001E\u001F\u2028\u2029 ]/g, "");
  var num = String(numberText).replace(/[\r\n\t\v\u00A0\u3000\u0007\u001E\u001F\u2028\u2029 ]/g, "");

  return clean.length > 0 && clean === num;
}

// 判断段落是否只剩旧编号的残片，例如 "(1.1)" 被删成 "1.1)"
function paragraphHasNumberRemnant(p, numberText) {
  var raw = "";
  try { raw = String(p.Range.Text); } catch (e0) { return false; }
  if (raw.indexOf("\f") >= 0) return false;

  var clean = raw.replace(/[\r\n\t\v\u00A0\u3000\u0007\u001E\u001F\u2028\u2029 ]/g, "");
  var num = String(numberText).replace(/[\r\n\t\v\u00A0\u3000\u0007\u001E\u001F\u2028\u2029 ]/g, "");
  var noOpen = num.replace(/^[\(（]/, "");

  return clean.length > 0 &&
    (clean === num || clean === noOpen || clean === ")" || clean === "）");
}

function stripLeadingAllowed(s) {
  var m = String(s).match(/^[ \t\v\n\r\u00A0\u3000\u001E\u001F\u2028\u2029]*/);
  return m ? m[0].length : 0;
}

// 只包含空白和 StyleSeparator 等非打印字符
function onlySeparators(s) {
  return String(s).replace(/[ \t\v\n\r\u00A0\u3000\u001E\u001F\u2028\u2029]/g, "").length === 0;
}

function matchLeadingNumber(s) {
  var m = String(s).match(/^[\(（][ \t\u00A0\u3000\u001E\u001F\u2028\u2029]*[0-9]+(?:[.\-][0-9]+)*[ \t\u00A0\u3000\u001E\u001F\u2028\u2029]*[\)）]/);
  return m ? m[0] : "";
}

// 取段落之后的下一段。不依赖段落序号，编辑后仍然有效。
// 依次尝试 Paragraph.Next、段尾折叠区域；拿到的仍是本段则返回 null。
function nextParagraphAfter(doc, p) {
  var ps = -1;
  try { ps = Number(p.Range.Start); } catch (ePS) { return null; }

  var cands = [];
  try {
    var nx = p.Next;
    if (nx) cands.push(nx);
  } catch (eNx) {}

  try {
    var pe = Number(p.Range.End);
    cands.push(doc.Range(pe, pe).Paragraphs.Item(1));
  } catch (ePe) {}

  for (var i = 0; i < cands.length; i++) {
    var c = cands[i];
    var cs = -1;
    try { cs = Number(c.Range.Start); } catch (eC) { continue; }
    if (cs > ps) return c;
  }
  return null;
}

// 在公式后面查找“已有编号”。
// 允许空格、制表符、软换行、硬回车、StyleSeparator；遇到分页/分节符立即停止。
// paraIndex 为扫描时的段落序号，仅在无法用对象定位下一段时兜底，可为 0。
function findTrailingNumber(doc, p, om, paraIndex) {
  var pos = -1;
  try { pos = Number(om.Range.End); } catch (e0) { return { found: false }; }

  var curPara = p;
  var baseIndex = (paraIndex > 0) ? paraIndex : 0;
  var curIndex = baseIndex;
  var maxLook = 4;

  for (var step = 0; step < maxLook; step++) {
    var fullEnd = -1;
    var segEnd = -1;
    try {
      fullEnd = Number(curPara.Range.End);
      segEnd = fullEnd - 1;
    } catch (e1) { return { found: false }; }
    if (segEnd < pos) segEnd = pos;

    var fullText = "";
    var segText = "";
    try { fullText = String(doc.Range(pos, fullEnd).Text); } catch (e2) {}
    try { segText = String(doc.Range(pos, segEnd).Text); } catch (e3) {}

    if (hasPageBreakText(fullText)) return { found: false, blocked: true };

    var wsLen = stripLeadingAllowed(segText);
    var rest = segText.substring(wsLen);
    var restPos = pos + wsLen;

    if (rest.length > 0) {
      var num = matchLeadingNumber(rest);
      if (!num) return { found: false };

      // 编号后面必须只剩空白/StyleSeparator，否则是“式(1.14)中……”这类正文引用
      var afterNum = rest.substring(num.length);
      if (!onlySeparators(afterNum)) return { found: false };

      return {
        found: true,
        text: num,
        range: doc.Range(restPos, restPos + num.length),
        paraIndex: curIndex,
        sameParagraph: (step === 0),
        nextPara: null
      };
    }

    var nextPara = nextParagraphAfter(doc, curPara);
    if (!nextPara && baseIndex > 0) {
      // 仅扫描阶段（有可靠段落序号）才用序号兜底，避免编辑后序号错位误判
      try { nextPara = doc.Paragraphs.Item(curIndex + 1); } catch (e4) {}
    }
    if (!nextPara) return { found: false };

    var ntCount = -1;
    try { ntCount = Number(nextPara.Range.Tables.Count); } catch (e5) { ntCount = -1; }
    if (ntCount > 0) return { found: false };

    var ntStart = -1;
    try { ntStart = Number(nextPara.Range.Start); } catch (e6) { return { found: false }; }

    var ntFull = "";
    try { ntFull = String(doc.Range(ntStart, Number(nextPara.Range.End)).Text); } catch (e7) {}
    if (hasPageBreakText(ntFull)) return { found: false, blocked: true };

    var ntOmCount = 0;
    try { ntOmCount = Number(nextPara.Range.OMaths.Count); } catch (e8) {}
    var ntShapeCount = 0;
    try { ntShapeCount = Number(nextPara.Range.InlineShapes.Count); } catch (e9) {}

    // 下一段如果有公式/嵌入对象，只接受“编号 + 非打印字符 + 对象”的结构
    var objectStart = -1;
    if (ntOmCount > 0) {
      try { objectStart = Number(nextPara.Range.OMaths.Item(1).Range.Start); } catch (e10) {}
    } else if (ntShapeCount > 0) {
      try { objectStart = Number(nextPara.Range.InlineShapes.Item(1).Range.Start); } catch (e11) {}
    }

    var ntText = "";
    if (objectStart > ntStart) {
      try { ntText = String(doc.Range(ntStart, objectStart).Text); } catch (e12) {}
    } else if (objectStart === -1) {
      try { ntText = String(doc.Range(ntStart, Number(nextPara.Range.End) - 1).Text); } catch (e13) {}
    } else {
      // 段落一开头就是公式/对象，没有前导编号
      return { found: false };
    }

    var ntWs = stripLeadingAllowed(ntText);
    var ntRest = ntText.substring(ntWs);
    var ntNum = matchLeadingNumber(ntRest);
    if (ntNum) {
      var ntAfter = ntRest.substring(ntNum.length);
      if (onlySeparators(ntAfter)) {
        return {
          found: true,
          text: ntNum,
          range: doc.Range(ntStart + ntWs, ntStart + ntWs + ntNum.length),
          paraIndex: curIndex + 1,
          sameParagraph: false,
          nextPara: nextPara
        };
      }
      // 编号后还有正文，视为正文引用
      return { found: false };
    }

    // 下一段有公式/对象但没有前导编号，不能继续往后借用编号
    if (objectStart >= 0) return { found: false };

    // 下一段只有空白/StyleSeparator 才继续往后找
    if (!onlySeparators(ntText)) return { found: false };

    curPara = nextPara;
    curIndex = curIndex + 1;
    pos = ntStart;
  }

  return { found: false };
}

// 公式前后是否还有可见正文（旧编号不算正文）
function formulaMixed(doc, p, om, trailing) {
  var pStart = -1, pEnd = -1, omStart = -1, omEnd = -1;
  try {
    pStart = Number(p.Range.Start);
    pEnd = Number(p.Range.End);
    omStart = Number(om.Range.Start);
    omEnd = Number(om.Range.End);
  } catch (e0) { return true; }

  var before = "";
  var after = "";
  try { before = String(doc.Range(pStart, omStart).Text); } catch (e1) {}
  try { after = String(doc.Range(omEnd, Math.max(omEnd, pEnd - 1)).Text); } catch (e2) {}

  if (hasPageBreakText(before) || hasPageBreakText(after)) return true;

  var b = before.replace(/[ \t\v\n\r\u00A0\u3000\u001E\u001F\u2028\u2029]/g, "");
  var a = after.replace(/[ \t\v\n\r\u00A0\u3000\u001E\u001F\u2028\u2029]/g, "");

  if (trailing && trailing.found && trailing.sameParagraph && trailing.text) {
    var idx = a.indexOf(trailing.text);
    if (idx >= 0) a = a.substring(0, idx) + a.substring(idx + trailing.text.length);
  }

  return (b.length > 0 || a.length > 0);
}

// 把公式从混排段落中拆出来，使其独占一个段落
function isolateFormula(doc, p, om) {
  var raw = "";
  try { raw = String(p.Range.Text); } catch (eRaw) {}
  if (hasPageBreakText(raw)) {
    return { ok: false, reason: "段落中含分页/分节符" };
  }

  var info = standaloneInfo(doc, p, om);
  if (info.standalone) return { ok: true, p: p, om: om };

  var pStart = -1, pEnd = -1, omStart = -1, omEnd = -1;
  try {
    pStart = Number(p.Range.Start);
    pEnd = Number(p.Range.End);
    omStart = Number(om.Range.Start);
    omEnd = Number(om.Range.End);
  } catch (e0) { return { ok: false, reason: "无法读取公式范围" }; }

  var before = "";
  var after = "";
  try { before = String(doc.Range(pStart, omStart).Text); } catch (e1) {}
  try { after = String(doc.Range(omEnd, Math.max(omEnd, pEnd - 1)).Text); } catch (e2) {}

  if (hasPageBreakText(before) || hasPageBreakText(after)) {
    return { ok: false, reason: "公式前后存在分页/分节符" };
  }

  // 删除紧跟在公式后面的旧编号，避免拆分后残留
  var wsLen = stripLeadingAllowed(after);
  var rest = after.substring(wsLen);
  var num = matchLeadingNumber(rest);
  if (num) {
    var afterNum = rest.substring(num.length);
    if (onlySeparators(afterNum)) {
      var numStart = omEnd + wsLen;
      run(function () { doc.Range(numStart, numStart + num.length).Delete(); return true; });
      after = after.substring(0, wsLen) + after.substring(wsLen + num.length);
    }
  }

  var bClean = before.replace(/[ \t\v\n\r\u00A0\u3000\u001E\u001F\u2028\u2029]/g, "");
  var aClean = after.replace(/[ \t\v\n\r\u00A0\u3000\u001E\u001F\u2028\u2029]/g, "");
  var needBefore = bClean.length > 0;
  var needAfter = aClean.length > 0;

  if (needBefore) {
    var r1 = run(function () { om.Range.InsertParagraphBefore(); return true; });
    if (!r1.ok) return { ok: false, reason: "插入前分段失败：" + r1.error };
  }
  if (needAfter) {
    var r2 = run(function () { om.Range.InsertParagraphAfter(); return true; });
    if (!r2.ok) return { ok: false, reason: "插入后分段失败：" + r2.error };
  }

  var p2 = null;
  var om2 = null;
  try { p2 = om.Range.Paragraphs.Item(1); } catch (e3) {}
  if (!p2) return { ok: false, reason: "拆分后无法读取段落" };
  try { om2 = p2.Range.OMaths.Item(1); } catch (e4) {}
  if (!om2) return { ok: false, reason: "拆分后无法读取公式" };

  var info2 = standaloneInfo(doc, p2, om2);
  if (!info2.standalone) {
    return { ok: false, reason: "拆分后公式仍与正文混排" };
  }

  return { ok: true, p: p2, om: om2 };
}

function getTextWidth(doc) {
  var r = run(function () {
    var ps = doc.PageSetup;
    return Number(ps.PageWidth) - Number(ps.LeftMargin) - Number(ps.RightMargin);
  });
  if (r.ok && Number(r.value) > 0) return Number(r.value);
  return 426.1;
}

function getBodyFontSize(doc) {
  var r = run(function () { return Number(doc.Styles.Item("正文").Font.Size); });
  if (r.ok && Number(r.value) > 0) return Number(r.value);
  return 10.5;
}

function numberWidthFor(maxChars, fontSize) {
  var w = maxChars * fontSize * 0.55 + 8;
  if (w < 32) w = 32;
  if (w > 64) w = 64;
  return Math.round(w * 100) / 100;
}

function readProp(name, def) {
  try {
    var doc = state.doc || Application.ActiveDocument;
    var props = doc.CustomDocumentProperties;
    for (var i = 1; i <= props.Count; i++) {
      var p = props.Item(i);
      if (String(p.Name) === name) return String(p.Value);
    }
  } catch (e) {}
  return def;
}

function writeProp(name, value) {
  var props = null;
  try {
    var doc = state.doc || Application.ActiveDocument;
    props = doc.CustomDocumentProperties;
  } catch (e1) { return false; }

  for (var i = 1; i <= props.Count; i++) {
    var p = null;
    try { p = props.Item(i); } catch (e2) { continue; }
    if (String(p.Name) === name) {
      var r1 = run(function () { p.Value = String(value); return true; });
      return r1.ok;
    }
  }

  var r2 = run(function () {
    props.Add(name, false, 4, String(value));
    return true;
  });
  return r2.ok;
}

function loadConfig() {
  var m = readProp("EQNUM_MODE", CFG.mode);
  CFG.mode = (m === "continuous") ? "continuous" : "chapter";
  CFG.level = readProp("EQNUM_LEVEL", "1") === "2" ? 2 : 1;
  CFG.start = Number(readProp("EQNUM_START", "1"));
  if (!(CFG.start >= 1)) CFG.start = 1;
  CFG.refPrefix = readProp("EQNUM_REF_PREFIX", "式");
  var rr = readProp("EQNUM_REFRESH_REFS", "ours");
  CFG.refreshRefs = (rr === "all" || rr === "none") ? rr : "ours";
  CFG.sub = readProp("EQNUM_SUB", "0") === "1" ? 1 : 0;
  CFG.subStyle = readProp("EQNUM_SUB_STYLE", "alpha");
}

function saveConfig() {
  writeProp("EQNUM_MODE", CFG.mode);
  writeProp("EQNUM_LEVEL", String(CFG.level));
  writeProp("EQNUM_START", String(CFG.start));
  writeProp("EQNUM_REF_PREFIX", CFG.refPrefix);
  writeProp("EQNUM_REFRESH_REFS", CFG.refreshRefs);
  writeProp("EQNUM_SUB", String(CFG.sub));
  writeProp("EQNUM_SUB_STYLE", CFG.subStyle);
}

function ensureDocument() {
  // 选择 OMaths 最多的文档，避免误改探测报告等文档
  var best = null;
  var bestCount = -1;
  var total = 0;
  try { total = Number(Application.Documents.Count); } catch (e0) {}

  for (var i = 1; i <= total; i++) {
    var d = null;
    try { d = Application.Documents.Item(i); } catch (e1) { continue; }
    var c = -1;
    try { c = Number(d.OMaths.Count); } catch (e2) {}
    if (c > bestCount) {
      bestCount = c;
      best = d;
    }
  }

  if (best && bestCount > 0) {
    run(function () { best.Activate(); return true; });
    state.doc = best;
    return best;
  }

  var r = run(function () { return Application.ActiveDocument; });
  if (!r.ok || !r.value) {
    alert("没有打开的文档");
    return null;
  }
  state.doc = r.value;
  return r.value;
}

function showSummary(title) {
  status("");
  var text = title + "\r\r" + "INFO=" + state.info + "，WARN=" + state.warn + "，FAIL=" + state.fail;
  if (state.fail > 0) {
    text += "\r\r失败信息：\r";
    for (var i = 0; i < state.messages.length; i++) {
      if (state.messages[i].indexOf("[FAIL]") === 0) text += state.messages[i] + "\r";
    }
  }
  try { alert(text); } catch (e) {
    consoleLog(text);
  }
}

function settingsSummary() {
  var modeText = (CFG.mode === "chapter")
    ? "按章（标题 " + CFG.level + "）"
    : "连续";

  return "编号模式：" + modeText +
    "\r起始值：" + CFG.start + (CFG.mode === "chapter" ? "（按章模式忽略）" : "") +
    "\r引用前缀：" + (CFG.refPrefix ? CFG.refPrefix : "（无）") +
    "\r刷新引用：" + (CFG.refreshRefs === "all" ? "全部" : (CFG.refreshRefs === "none" ? "不刷新" : "仅本工具")) +
    "\r子公式编号：" + (CFG.sub ? "开启（预留）" : "关闭");
}

function parseSettingsLine(line) {
  var text = String(line).replace(/；/g, ";");
  var parts = text.split(";");
  var values = {};
  var errors = [];

  for (var i = 0; i < parts.length; i++) {
    var part = String(parts[i]).replace(/^\s+|\s+$/g, "");
    if (!part) continue;

    var eq = part.indexOf("=");
    if (eq < 0) eq = part.indexOf(":");
    if (eq < 0) eq = part.indexOf("：");
    if (eq < 0) {
      errors.push("缺少等号：" + part);
      continue;
    }

    var key = part.substring(0, eq).replace(/^\s+|\s+$/g, "");
    var val = part.substring(eq + 1).replace(/^\s+|\s+$/g, "");
    var k = key.toLowerCase();

    if (key === "编号模式" || k === "mode") {
      if (val === "按章" || /chapter/i.test(val)) values.mode = "chapter";
      else if (val === "连续" || /continuous/i.test(val)) values.mode = "continuous";
      else errors.push("编号模式只能是“按章”或“连续”");
    } else if (key === "章层级" || k === "level") {
      if (/^(1|一级|标题1|h1)$/i.test(val)) values.level = 1;
      else if (/^(2|二级|标题2|h2)$/i.test(val)) values.level = 2;
      else errors.push("章层级只能是 1 或 2");
    } else if (key === "起始值" || k === "start") {
      if (/^\d+$/.test(val)) {
        var n = parseInt(val, 10);
        if (n >= 1 && n <= 999) values.start = n;
        else errors.push("起始值需在 1-999 之间");
      } else {
        errors.push("起始值必须是整数");
      }
    } else if (key === "引用前缀" || k === "prefix") {
      values.refPrefix = val;
    } else if (key === "刷新引用" || k === "refs") {
      if (/^(全部|all)$/i.test(val)) values.refreshRefs = "all";
      else if (/^(仅本工具|本工具|ours)$/i.test(val)) values.refreshRefs = "ours";
      else if (/^(不刷新|无|none|off)$/i.test(val)) values.refreshRefs = "none";
      else errors.push("刷新引用只能是“全部/仅本工具/不刷新”");
    } else if (key === "子公式" || k === "sub") {
      if (/^(开|开启|是|on|true|1)$/i.test(val)) values.sub = 1;
      else if (/^(关|关闭|否|off|false|0)$/i.test(val)) values.sub = 0;
      else errors.push("子公式只能是“开”或“关”");
    } else {
      errors.push("未知设置项：" + key);
    }
  }

  if (errors.length > 0) return { ok: false, error: errors.join("\r") };
  return { ok: true, values: values };
}

function runSettings(doc) {
  loadConfig();

  var current =
    "编号模式=" + (CFG.mode === "chapter" ? "按章" : "连续") +
    "; 章层级=" + CFG.level +
    "; 起始值=" + CFG.start +
    "; 引用前缀=" + CFG.refPrefix +
    "; 刷新引用=" + (CFG.refreshRefs === "all" ? "全部" : (CFG.refreshRefs === "none" ? "不刷新" : "仅本工具")) +
    "; 子公式=" + (CFG.sub ? "开" : "关");

  if (typeof InputBox !== "function") {
    alert("当前宏环境不支持 InputBox。\r\r当前设置：\r" + settingsSummary());
    return;
  }

  var help =
    "一次编辑全部设置，各项用分号分隔，键和值用等号连接。\r" +
    "可只保留需要修改的项，未出现的项保持原值。\r\r" +
    "可用键：\r" +
    "  编号模式 = 按章 / 连续\r" +
    "  章层级 = 1 / 2\r" +
    "  起始值 = 1-999（仅连续模式）\r" +
    "  引用前缀 = 任意文本，可留空\r" +
    "  刷新引用 = 全部 / 仅本工具 / 不刷新（默认仅本工具，速度最快）\r" +
    "  子公式 = 开 / 关（预留）\r\r" +
    "当前值已填入输入框，改完点确定：";

  var input = InputBox(help, "EQNUM 编号设置", current);
  if (input === null || input === undefined) return;

  var text = String(input).replace(/^\s+|\s+$/g, "");
  if (text === "") return;

  var parsed = parseSettingsLine(text);
  if (!parsed.ok) {
    alert("设置未保存：\r" + parsed.error + "\r\r请使用“键=值”的格式，各项用分号分隔。");
    return;
  }

  var v = parsed.values;
  if (v.mode !== undefined) CFG.mode = v.mode;
  if (v.level !== undefined) CFG.level = v.level;
  if (v.start !== undefined) CFG.start = v.start;
  if (v.refPrefix !== undefined) CFG.refPrefix = v.refPrefix;
  if (v.sub !== undefined) CFG.sub = v.sub;

  saveConfig();
  alert("设置已保存：\r" + settingsSummary());
}

/* ============================================================
 * 扫描与识别
 * ============================================================ */
function chapterValue(h1, h2) {
  var c = (CFG.level === 2 && h2) ? h2 : h1;
  return c || "0";
}

function chapterFor(h1, h2) {
  if (CFG.mode === "continuous") return "";
  return chapterValue(h1, h2);
}

function cellBookmarkName(cell) {
  try {
    var bms = cell.Range.Bookmarks;
    for (var i = 1; i <= bms.Count; i++) {
      var n = String(bms.Item(i).Name);
      if (n.indexOf(BM_PREFIX) === 0) return n;
    }
  } catch (e) {}
  return "";
}

// 判断段落是否属于本工具生成的两栏编号表
function getOurTableInfo(p, doc) {
  var t = null;
  try {
    if (p.Range.Tables.Count < 1) return null;
    t = p.Range.Tables.Item(1);
  } catch (e1) { return null; }

  var rows = -1;
  var cols = -1;
  try { rows = Number(t.Rows.Count); cols = Number(t.Columns.Count); } catch (e2) {}
  if (rows !== 1 || cols !== 2) return null;

  var cell2 = null;
  try { cell2 = t.Cell(1, 2); } catch (e3) { return null; }

  var bm = cellBookmarkName(cell2);
  if (!bm) return null;

  var om = null;
  try {
    if (Number(t.Cell(1, 1).Range.OMaths.Count) < 1) return null;
    om = t.Cell(1, 1).Range.OMaths.Item(1);
  } catch (e4) { return null; }

  return { table: t, bookmarkName: bm, om: om };
}

function standaloneInfo(doc, p, om) {
  var pStart = -1;
  var pEnd = -1;
  var omStart = -1;
  var omEnd = -1;

  try { pStart = Number(p.Range.Start); pEnd = Number(p.Range.End); } catch (e1) { return { standalone: false }; }
  try { omStart = Number(om.Range.Start); omEnd = Number(om.Range.End); } catch (e2) { return { standalone: false }; }

  var before = "";
  var after = "";
  try { before = String(doc.Range(pStart, omStart).Text); } catch (e3) {}
  try { after = String(doc.Range(omEnd, pEnd).Text); } catch (e4) {}

  var cleanB = cleanAllowed(before);
  var cleanA = cleanAllowed(after);

  return {
    standalone: (cleanB.length === 0 && cleanA.length === 0),
    before: before,
    after: after,
    trailing: trailingNumber(after)
  };
}

function headingLevelOf(sn) {
  if (isHeadingLevel(sn, 1)) return 1;
  if (isHeadingLevel(sn, 2)) return 2;
  return 0;
}

function chapterBounds(doc, curPara) {
  var total = 1;
  var tr = run(function () { return doc.Paragraphs.Count; });
  if (tr.ok) total = Number(tr.value);
  if (!(curPara >= 1) || curPara > total) curPara = 1;

  var start = 1;
  for (var i = curPara; i >= 1; i--) {
    var p = null;
    try { p = doc.Paragraphs.Item(i); } catch (e1) { continue; }
    var lv = headingLevelOf(styleName(p));
    if (lv > 0 && lv <= CFG.level) { start = i; break; }
  }

  var end = total;
  for (var j = start + 1; j <= total; j++) {
    var p2 = null;
    try { p2 = doc.Paragraphs.Item(j); } catch (e2) { continue; }
    var lv2 = headingLevelOf(styleName(p2));
    if (lv2 > 0 && lv2 <= CFG.level) { end = j - 1; break; }
  }

  var h1 = "";
  var h2 = "";
  for (var k = start; k >= 1; k--) {
    var pk = null;
    try { pk = doc.Paragraphs.Item(k); } catch (e3) { continue; }
    var lvk = headingLevelOf(styleName(pk));
    if (lvk === 1) {
      h1 = headingNumber(pk);
      break;
    }
    if (lvk === 2 && CFG.level === 2 && !h2) h2 = headingNumber(pk);
  }
  if (CFG.level === 1) h2 = "";

  return { start: start, end: end, h1: h1, h2: h2 };
}

function collectFormulas(doc, options) {
  options = options || {};
  section("扫描公式");

  var total = 0;
  var tr = run(function () { return doc.Paragraphs.Count; });
  if (!tr.ok) {
    log("FAIL", "无法读取段落总数：" + tr.error);
    return null;
  }
  total = Number(tr.value);

  var startPara = Number(options.startPara);
  if (!(startPara >= 1)) startPara = 1;
  var endPara = Number(options.endPara);
  if (!(endPara >= startPara) || endPara > total) endPara = total;
  if (startPara > 1 || endPara < total) {
    log("INFO", "扫描范围：段 #" + startPara + " – #" + endPara);
  }

  var tCol = Date.now();

  var out = {
    formulas: [],
    skipNoNumber: 0,
    skipBreak: 0,
    skipMulti: 0,
    skipTable: 0,
    mixedCount: 0,
    inlineCount: 0,
    crossParaCount: 0,
    existingTables: 0,
    h1: 0,
    h2: 0
  };

  var h1 = options.initH1 || "";
  var h2 = options.initH2 || "";

  var span = endPara - startPara + 1;
  for (var i = startPara; i <= endPara; i++) {
    var scanned = i - startPara + 1;
    if (scanned % PROGRESS_STEP === 0) {
      status("EQNUM 扫描中：" + scanned + " / " + span);
      log("INFO", "已扫描段落 " + i + " / " + endPara);
    }

    var p = null;
    try { p = doc.Paragraphs.Item(i); } catch (e1) { continue; }

    var sn = styleName(p);
    if (isHeadingLevel(sn, 1)) {
      out.h1++;
      var n1 = headingNumber(p);
      if (n1) { h1 = n1; h2 = ""; }
    } else if (isHeadingLevel(sn, 2)) {
      out.h2++;
      var n2 = headingNumber(p);
      if (n2) h2 = n2;
    }

    var omCount = -1;
    var cr = run(function () { return p.Range.OMaths.Count; });
    if (cr.ok) omCount = Number(cr.value);
    if (omCount <= 0) continue;

    var allOms = [];
    for (var j = 1; j <= omCount; j++) {
      var omItem = null;
      try { omItem = p.Range.OMaths.Item(j); } catch (e2) { continue; }
      allOms.push(omItem);
    }
    if (allOms.length === 0) continue;

    var tableCount = -1;
    var tcr = run(function () { return p.Range.Tables.Count; });
    if (tcr.ok) tableCount = Number(tcr.value);

    if (tableCount > 0) {
      var ours = getOurTableInfo(p, doc);
      if (ours) {
        out.formulas.push({
          kind: "table",
          para: i,
          p: p,
          table: ours.table,
          bookmarkName: ours.bookmarkName,
          om: ours.om,
          h1: h1,
          h2: h2,
          trailing: "",
          trailingRange: null,
          trailingSame: true,
          trailingPara: i,
          mixed: false,
          inline: false
        });
        out.existingTables++;
      } else {
        out.skipTable += allOms.length;
      }
      continue;
    }

    if (allOms.length > 1) {
      out.skipMulti += allOms.length;
      continue;
    }

    var om = allOms[0];
    var ty = 0;
    try { ty = Number(om.Type); } catch (e3) {}

    var trail = findTrailingNumber(doc, p, om, i);
    if (!trail.found) {
      if (trail.blocked) out.skipBreak++;
      else out.skipNoNumber++;
      continue;
    }

    var mixed = formulaMixed(doc, p, om, trail);
    out.formulas.push({
      kind: "plain",
      para: i,
      p: p,
      table: null,
      bookmarkName: "",
      om: om,
      h1: h1,
      h2: h2,
      trailing: trail.text,
      trailingRange: trail.range,
      trailingSame: trail.sameParagraph,
      trailingPara: trail.paraIndex,
      mixed: mixed,
      inline: (ty !== 0)
    });
    if (mixed) out.mixedCount++;
    if (ty !== 0) out.inlineCount++;
    if (!trail.sameParagraph) out.crossParaCount++;
  }

  status("");
  log("INFO", "带编号候选 " + out.formulas.length +
    "（已有编号表 " + out.existingTables +
    "，跨段旧号 " + out.crossParaCount +
    "；混排 " + out.mixedCount +
    "，行内 " + out.inlineCount + "）");
  log("INFO", "无编号跳过 " + out.skipNoNumber +
    "，分页/分节阻断 " + out.skipBreak +
    "，同段多公式跳过 " + out.skipMulti +
    "，表格内跳过 " + out.skipTable);
  log("INFO", "标题 1=" + out.h1 + "，标题 2=" + out.h2);
  log("INFO", "扫描完成，用时 " + (Date.now() - tCol) + "ms");

  return out;
}

/* ============================================================
 * 表格排版与编号写入
 * ============================================================ */
function maxBookmarkId(doc) {
  var max = 0;
  try {
    var bms = doc.Bookmarks;
    for (var i = 1; i <= bms.Count; i++) {
      var n = String(bms.Item(i).Name);
      if (n.indexOf(BM_PREFIX) === 0) {
        var num = Number(n.substring(BM_PREFIX.length));
        if (num > max) max = num;
      }
    }
  } catch (e) {}
  return max;
}

function makeBookmarkName(id) {
  var s = String(id);
  while (s.length < 4) s = "0" + s;
  return BM_PREFIX + s;
}

function clearCell(cell) {
  var r = run(function () {
    cell.Range.Text = "";
    return String(cell.Range.Text);
  });
  if (r.ok && String(r.value).replace(/[\s\u0007]/g, "") === "") return true;

  var r2 = run(function () {
    var rr = cell.Range;
    rr.MoveEnd(1, -1); // 1 = wdCharacter，排除单元格标记
    rr.Delete();
    return true;
  });
  return r2.ok;
}

function applyTableLayout(doc, table, maxChars, layout) {
  var textWidth = (layout && layout.textWidth > 0) ? layout.textWidth : getTextWidth(doc);
  var fontSize = (layout && layout.fontSize > 0) ? layout.fontSize : getBodyFontSize(doc);
  var numWidth = numberWidthFor(maxChars, fontSize);
  var leftWidth = Math.round((textWidth - numWidth) * 100) / 100;

  run(function () { table.AllowAutoFit = false; table.Borders.Enable = 0; return true; });
  run(function () { table.PreferredWidthType = 2; table.PreferredWidth = 100; return true; });

  run(function () { table.TopPadding = 0; table.BottomPadding = 0; return true; });
  run(function () { table.LeftPadding = 0; table.RightPadding = 0; return true; });

  // 表级段落格式，替代左右单元格各设一遍
  run(function () {
    var pf = table.Range.ParagraphFormat;
    pf.LeftIndent = 0;
    pf.RightIndent = 0;
    pf.FirstLineIndent = 0;
    pf.SpaceBefore = 0;
    pf.SpaceAfter = 0;
    return true;
  });

  run(function () { table.Cell(1, 1).Range.ParagraphFormat.Alignment = 1; return true; });
  run(function () { table.Cell(1, 2).Range.ParagraphFormat.Alignment = 2; return true; });

  // 批量垂直居中，失败再逐单元格回退
  var va = run(function () { table.Range.Cells.VerticalAlignment = 1; return true; });
  if (!va.ok) {
    run(function () { table.Cell(1, 1).VerticalAlignment = 1; return true; });
    run(function () { table.Cell(1, 2).VerticalAlignment = 1; return true; });
  }

  run(function () { table.Cell(1, 2).Width = numWidth; return true; });
  run(function () { table.Cell(1, 1).Width = leftWidth; return true; });

  // 两列合计已等于版心宽，再确认表格保持 100%
  run(function () { table.PreferredWidthType = 2; table.PreferredWidth = 100; return true; });

  return { numWidth: numWidth, leftWidth: leftWidth, textWidth: textWidth };
}

// 已编号表只更新列宽，避免重复设置边框、内边距、对齐等触发整页重排
function applyTableWidths(table, maxChars, layout) {
  var textWidth = (layout && layout.textWidth > 0) ? layout.textWidth : 426.1;
  var fontSize = (layout && layout.fontSize > 0) ? layout.fontSize : 10.5;
  var numWidth = numberWidthFor(maxChars, fontSize);
  var leftWidth = Math.round((textWidth - numWidth) * 100) / 100;

  run(function () { table.Cell(1, 2).Width = numWidth; return true; });
  run(function () { table.Cell(1, 1).Width = leftWidth; return true; });
  run(function () { table.PreferredWidthType = 2; table.PreferredWidth = 100; return true; });
}

function deleteBookmarkIfExists(doc, name) {
  run(function () {
    var b = findBookmark(doc, name);
    if (b) b.Delete();
    return true;
  });
}

// 在右单元格写入题注字段 "( STYLEREF . SEQ 公式 )"
function cellParaEnd(cell) {
  try {
    var ps = cell.Range.Paragraphs;
    var p = ps.Item(ps.Count);
    return Number(p.Range.End) - 1;
  } catch (e) { return -1; }
}

function addFieldInCellAtEnd(doc, cell, code, doUpdate) {
  var end = cellParaEnd(cell);
  if (end < 0) return null;

  var fr = run(function () {
    return doc.Fields.Add(doc.Range(end, end), -1, code, false);
  });
  if (!fr.ok || !fr.value) return null;

  var f = fr.value;
  if (doUpdate !== false) {
    var t0 = Date.now();
    var up = run(function () { f.Update(); return true; });
    var dt = Date.now() - t0;
    if (!up.ok) log("WARN", "域更新失败：" + up.error);
    else if (dt > 1000) log("WARN", "域更新耗时 " + dt + "ms：" + code);
  }
  return f;
}

function removeOldNumber(doc, trail, label) {
  if (!trail || !trail.range) return false;

  var removedWhole = false;
  var tag = label ? "（" + label + "）" : "";

  // 旧编号所在段落：优先用实时段落对象，避免编辑后段落序号失效
  function numberParagraph() {
    if (trail.nextPara) return trail.nextPara;
    if (trail.paraIndex > 0) {
      try { return doc.Paragraphs.Item(trail.paraIndex); } catch (eNP0) {}
    }
    return null;
  }

  // 旧编号单独成段：直接删整段，避免域/书签导致只删掉一部分
  if (!trail.sameParagraph) {
    var np0 = numberParagraph();

    if (np0 && paragraphOnlyNumber(np0, trail.text)) {
      var wp = run(function () { np0.Range.Delete(); return true; });
      if (wp.ok) {
        removedWhole = true;
        log("INFO", "已删除旧编号所在段落：" + trail.text + tag);
      } else {
        log("WARN", "删除旧编号段落失败：" + wp.error);
      }
    }
  }

  if (!removedWhole) {
    var del = run(function () { trail.range.Delete(); return true; });
    if (del.ok) log("INFO", "已删除旧编号：" + trail.text + tag);
    else log("WARN", "删除旧编号失败：" + del.error);

    if (del.ok && !trail.sameParagraph) {
      var np = numberParagraph();

      if (np && (paragraphIsEmpty(np) || paragraphHasNumberRemnant(np, trail.text))) {
        // 用“最后一注”的实际段尾判断，避免删掉文档最后的段落标记
        var isLast = false;
        try {
          var npc = Number(doc.Paragraphs.Count);
          var lastP = doc.Paragraphs.Item(npc);
          isLast = (Number(np.Range.End) >= Number(lastP.Range.End));
        } catch (eLast) {}

        if (!isLast) {
          var dp = run(function () { np.Range.Delete(); return true; });
          if (dp.ok) log("INFO", "已删除旧编号所在的空段落" + tag);
          else log("WARN", "删除空段落失败：" + dp.error);
        } else {
          log("INFO", "旧编号段落是文档最后一段，保留段落标记");
        }
      }
    }
  }

  return removedWhole;
}

function writeNumberCell(doc, table, name, level, startValue, deferUpdate) {
  var cell = null;
  try { cell = table.Cell(1, 2); } catch (e0) { return null; }

  var oldSize = 0;
  try { oldSize = Number(cell.Range.Font.Size); } catch (eSize) {}
  if (!(oldSize > 0)) {
    try { oldSize = getBodyFontSize(doc); } catch (eSize2) {}
  }
  if (!(oldSize > 0)) oldSize = 10.5;

  deleteBookmarkIfExists(doc, name);
  if (!clearCell(cell)) {
    log("FAIL", "清空编号单元格失败：" + name);
    return null;
  }

  var start = -1;
  try { start = Number(cell.Range.Start); } catch (e1) { return null; }

  run(function () { doc.Range(start, start).InsertAfter("("); return true; });

  if (CFG.mode === "chapter") {
    var f1 = addFieldInCellAtEnd(doc, cell, " STYLEREF " + level + " \\n ", !deferUpdate);
    if (!f1) {
      log("FAIL", "插入 STYLEREF 域失败：" + name);
      return null;
    }

    var end1 = cellParaEnd(cell);
    run(function () { if (end1 >= 0) doc.Range(end1, end1).InsertAfter("."); return true; });
  }

  var seqCode = " SEQ 公式 " +
    (startValue > 0 ? "\\r " + startValue + " " : "") +
    "\\* ARABIC " +
    (CFG.mode === "chapter" ? "\\s " + level + " " : "");
  var seq = addFieldInCellAtEnd(doc, cell, seqCode, !deferUpdate);
  if (!seq) {
    log("FAIL", "插入 SEQ 公式 域失败：" + name);
    return null;
  }

  var end2 = cellParaEnd(cell);
  run(function () { if (end2 >= 0) doc.Range(end2, end2).InsertAfter(")"); return true; });

  run(function () {
    var p = null;
    try { p = cell.Range.Paragraphs.Item(1); } catch (e4) { return false; }

    var sn = styleName(p);
    if (sn !== "题注" && sn !== "Caption") {
      try { p.Style = "题注"; } catch (e5) {
        try { p.Style = "Caption"; } catch (e6) {}
      }
    }

    var curSize = 0;
    try { curSize = Number(cell.Range.Font.Size); } catch (e8) {}
    if (!(curSize > 0) || Math.abs(curSize - oldSize) > 0.01) {
      try { cell.Range.Font.Size = oldSize; } catch (e7) {}
    }
    return true;
  });

  var bStart = start;
  var bEnd = (end2 >= 0) ? (end2 + 1) : -1;
  if (bEnd < 0) {
    try { bEnd = Number(cell.Range.End) - 2; } catch (e8) {}
  }

  if (bStart >= 0 && bEnd > bStart) {
    var br = run(function () {
      doc.Bookmarks.Add(name, doc.Range(bStart, bEnd));
      return true;
    });
    if (!br.ok) log("WARN", "书签创建失败 " + name + "：" + br.error);
  }

  run(function () { cell.Range.ParagraphFormat.Alignment = 2; return true; });
  return seq;
}

// 把独立段落整理成“公式 + 制表符”，再转两栏表格
function preparePlainFormula(doc, f) {
  var p = f.p;
  var om = f.om;

  var rawText = "";
  try { rawText = String(p.Range.Text); } catch (eRaw) {}
  if (hasPageBreakText(rawText)) {
    log("FAIL", "段 #" + f.para + " 含分页/分节符，跳过");
    return null;
  }

  var preInfo = standaloneInfo(doc, p, om);
  if (!preInfo.standalone) {
    log("FAIL", "段 #" + f.para + " 公式未与正文隔离，跳过");
    return null;
  }

  var ty0 = null;
  try { ty0 = Number(om.Type); } catch (eTy) {}
  if (ty0 !== 0) {
    var tr = run(function () { om.Type = 0; return true; });
    if (tr.ok) {
      try {
        p = om.Range.Paragraphs.Item(1);
        om = p.Range.OMaths.Item(1);
        f.p = p;
        f.om = om;
      } catch (eTy2) {}
    }
  }

  var pStart = -1;
  var pEnd = -1;
  var omStart = -1;
  var omEnd = -1;
  try {
    pStart = Number(p.Range.Start);
    pEnd = Number(p.Range.End);
    omStart = Number(om.Range.Start);
    omEnd = Number(om.Range.End);
  } catch (e0) { return null; }

  // 清掉公式前后可能残留的空白/旧编号，只留公式本身
  var shift = 0;
  if (omStart > pStart) {
    run(function () { doc.Range(pStart, omStart).Delete(); return true; });
    shift = omStart - pStart;
  }

  var p2End = pEnd - shift;
  var om2End = omEnd - shift;
  if (p2End - 1 > om2End) {
    run(function () { doc.Range(om2End, p2End - 1).Delete(); return true; });
  }

  var p2 = f.p;
  if (!p2) return null;

  // 制表符插在“段落标记之前”，确保落在公式外面。
  // 用 OMath.Range.End 定位时，WPS 可能把制表符算进公式，转表结果会嵌到公式内部。
  var pe = -1;
  try { pe = Number(p2.Range.End) - 1; } catch (ePe) { return null; }
  if (pe < 0) return null;

  run(function () { doc.Range(pe, pe).InsertAfter("\t"); return true; });

  var ptxt = "";
  try { ptxt = String(p2.Range.Text); } catch (ePt) {}
  if (ptxt.indexOf("\t") < 0) {
    log("FAIL", "段 #" + f.para + " 插入制表符失败");
    return null;
  }

  var tbl = null;
  var cr = run(function () { return p2.Range.ConvertToTable("\t"); });
  if (cr.ok && cr.value) tbl = cr.value;
  else {
    try { tbl = doc.Tables.Item(doc.Tables.Count); } catch (e3) {}
  }
  if (!tbl) {
    log("FAIL", "段落转两栏表格失败（段 #" + f.para + "）");
    return null;
  }

  // 校验：必须是两栏，且公式在第一栏；否则说明表格被嵌进了公式对象
  var cellCount = -1;
  try { cellCount = Number(tbl.Range.Cells.Count); } catch (eCC) {}
  var omInCell = -1;
  try { omInCell = Number(tbl.Cell(1, 1).Range.OMaths.Count); } catch (eOC) {}

  if (omInCell === 0 || (cellCount > 0 && cellCount !== 2)) {
    log("FAIL", "段 #" + f.para + " 转表结果异常（栏数=" + cellCount +
      "，第一栏公式数=" + omInCell + "），疑似表格被嵌入公式，请手动检查");
    return null;
  }

  return tbl;
}

function findBookmark(doc, name) {
  try {
    if (doc.Bookmarks.Exists(name)) return doc.Bookmarks.Item(name);
  } catch (e1) {}
  try {
    var bms = doc.Bookmarks;
    for (var i = 1; i <= bms.Count; i++) {
      var b = bms.Item(i);
      if (String(b.Name) === name) return b;
    }
  } catch (e2) {}
  return null;
}

function updateOurSeqFields(doc) {
  var items = [];
  try {
    var bms = doc.Bookmarks;
    for (var i = 1; i <= bms.Count; i++) {
      var b = bms.Item(i);
      var n = String(b.Name);
      if (n.indexOf(BM_PREFIX) !== 0) continue;
      var st = -1;
      try { st = Number(b.Range.Start); } catch (e1) {}
      items.push({ name: n, start: st });
    }
  } catch (e2) {}

  items.sort(function (a, b) { return a.start - b.start; });

  var tSeq = Date.now();
  log("INFO", "开始更新编号域：共 " + items.length + " 个");

  var ok = 0;
  var fail = 0;
  for (var k = 0; k < items.length; k++) {
    if (k > 0 && k % 20 === 0) {
      log("INFO", "更新编号域 " + k + " / " + items.length);
    }

    var t0 = Date.now();
    var r = run(function () {
      var b = findBookmark(doc, items[k].name);
      if (!b) throw new Error("书签不存在");
      var fs = b.Range.Fields;
      var fn = Number(fs.Count);
      if (fn < 1) throw new Error("书签内没有域");
      for (var j = 1; j <= fn; j++) fs.Item(j).Update();
      return true;
    });
    var dt = Date.now() - t0;

    if (r.ok) ok++;
    else {
      fail++;
      log("WARN", "更新编号域失败 " + items[k].name + "：" + r.error);
    }
    if (dt > 1000) {
      log("WARN", "更新编号域耗时 " + dt + "ms：" + items[k].name);
    }
  }

  log("INFO", "编号域更新完成：成功 " + ok + "，失败 " + fail +
    "，用时 " + (Date.now() - tSeq) + "ms");
  return { ok: ok, fail: fail };
}

function updateOurRefBookmarks(doc) {
  var tr = run(function () { return doc.Bookmarks.Count; });
  if (!tr.ok) return 0;

  var count = Number(tr.value);
  var updated = 0;
  var t0 = Date.now();
  log("INFO", "更新本工具引用书签：书签总数 " + count);

  for (var i = 1; i <= count; i++) {
    var b = null;
    try { b = doc.Bookmarks.Item(i); } catch (e1) { continue; }
    var name = "";
    try { name = String(b.Name); } catch (e2) { continue; }
    if (name.indexOf("EQNUM_REF_") !== 0) continue;

    var ur = run(function () { b.Range.Fields.Item(1).Update(); return true; });
    if (ur.ok) updated++;
    else log("WARN", "更新引用书签失败：" + name + "：" + ur.error);
  }

  log("INFO", "引用域更新完成：更新 " + updated + " 个，用时 " + (Date.now() - t0) + "ms");
  if (updated === 0) {
    log("INFO", "未找到 EQNUM_REF_ 书签；旧引用请用 EQNUM_UPDATE_REFS 或手动 F9");
  }
  return updated;
}

function updateRefFields(doc, mode) {
  mode = mode || CFG.refreshRefs || "ours";
  if (mode === "none") {
    log("INFO", "引用刷新：已设置为不刷新，跳过");
    return 0;
  }

  // 仅本工具：只更新带 EQNUM_REF_ 书签的引用，不扫描全部域
  if (mode === "ours") {
    return updateOurRefBookmarks(doc);
  }

  var wantCaption = true;
  var total = 0;
  var tr = run(function () { return doc.Fields.Count; });
  if (!tr.ok) return 0;
  total = Number(tr.value);

  var tRef = Date.now();
  log("INFO", "开始更新引用域：总域数 " + total +
    "，模式=" + (wantCaption ? "全部" : "仅本工具"));

  var updated = 0;
  for (var i = 1; i <= total; i++) {
    if (i % 2000 === 0) {
      status("刷新引用：" + i + " / " + total);
      log("INFO", "刷新引用 " + i + " / " + total);
    }
    var f = null;
    try { f = doc.Fields.Item(i); } catch (e1) { continue; }
    var code = "";
    try { code = String(f.Code.Text); } catch (e2) { continue; }

    var isOurs = (code.indexOf(BM_PREFIX) >= 0);
    var isCaption = (code.indexOf("_Ref") >= 0);
    if (/^\s*REF\b/i.test(code) && (isOurs || (wantCaption && isCaption))) {
      var t0 = Date.now();
      var r = run(function () { f.Update(); return true; });
      var dt = Date.now() - t0;
      if (r.ok) updated++;
      else log("WARN", "引用域更新失败：" + r.error);
      if (dt > 1000) log("WARN", "引用域更新耗时 " + dt + "ms：" + clip(code, 80));
    }
  }

  log("INFO", "引用域更新完成：更新 " + updated + " 个，用时 " + (Date.now() - tRef) + "ms");
  status("");
  return updated;
}

function listOurBookmarks(doc) {
  var items = [];
  try {
    var bms = doc.Bookmarks;
    for (var i = 1; i <= bms.Count; i++) {
      var b = bms.Item(i);
      var n = String(b.Name);
      if (n.indexOf(BM_PREFIX) !== 0) continue;
      var txt = "";
      var st = -1;
      try { txt = String(b.Range.Text); } catch (e1) {}
      try { st = Number(b.Range.Start); } catch (e2) {}
      items.push({ name: n, text: txt.replace(/[\r\u0007]/g, ""), start: st });
    }
  } catch (e3) {}
  items.sort(function (a, b) { return a.start - b.start; });
  return items;
}

// 向前查找当前公式所属的章号
// forScope=true 时忽略连续/按章模式，始终返回标题章号，用于“当前章节”范围判断
function chapterAt(doc, p, forScope) {
  if (!forScope && CFG.mode === "continuous") return "";

  var pStart = -1;
  try { pStart = Number(p.Range.Start); } catch (e0) {
    return forScope ? chapterValue("", "") : chapterFor("", "");
  }

  var idx = 0;
  try { idx = Number(doc.Range(0, pStart).Paragraphs.Count) + 1; } catch (e1) { idx = 0; }
  if (idx <= 0) return forScope ? chapterValue("", "") : chapterFor("", "");

  var h1 = "";
  var h2 = "";
  for (var k = idx; k >= 1; k--) {
    var q = null;
    try { q = doc.Paragraphs.Item(k); } catch (e2) { continue; }
    var sn = styleName(q);
    if (isHeadingLevel(sn, 1)) {
      var n1 = headingNumber(q);
      if (n1 && !h1) h1 = n1;
      break;
    }
    if (isHeadingLevel(sn, 2) && !h2) {
      var n2 = headingNumber(q);
      if (n2) h2 = n2;
    }
  }
  return forScope ? chapterValue(h1, h2) : chapterFor(h1, h2);
}

// 收集本工具已经生成的编号表，可按章过滤
function collectOurTables(doc, scopeChapter) {
  var items = listOurBookmarks(doc);
  var out = [];
  var skipped = 0;

  for (var i = 0; i < items.length; i++) {
    var st = items[i].start;
    if (st < 0) { skipped++; continue; }

    var p = null;
    try { p = doc.Range(st, st).Paragraphs.Item(1); } catch (e1) { skipped++; continue; }

    var info = getOurTableInfo(p, doc);
    if (!info) { skipped++; continue; }

    var ch = chapterAt(doc, p, true);
    if (scopeChapter && ch !== scopeChapter) continue;

    out.push({
      kind: "table",
      para: 0,
      p: p,
      table: info.table,
      bookmarkName: info.bookmarkName,
      om: info.om,
      h1: "",
      h2: "",
      chapter: ch,
      trailing: "",
      trailingRange: null,
      trailingSame: true,
      trailingPara: 0,
      mixed: false,
      inline: false
    });
  }

  return { formulas: out, skipped: skipped };
}

// 处理期间临时关闭会触发整篇重排/拼写的后台行为。
// 每一项都读回校验，不支持的 WPS 版本自动跳过，并在日志里说明。
function pauseHeavyUpdates() {
  var st = { pagination: null, spell: null, grammar: null, view: null, applied: [] };

  var o = null;
  try { o = Application.Options; } catch (e0) {}

  if (o) {
    try {
      st.pagination = !!o.Pagination;
      o.Pagination = false;
      st.applied.push(!!o.Pagination === false ? "自动分页=关闭" : "自动分页=未生效");
    } catch (e1) {
      st.pagination = null;
      st.applied.push("自动分页=不支持");
    }

    try {
      st.spell = !!o.CheckSpellingAsYouType;
      o.CheckSpellingAsYouType = false;
      st.applied.push(!!o.CheckSpellingAsYouType === false ? "拼写检查=关闭" : "拼写检查=未生效");
    } catch (e2) {
      st.spell = null;
      st.applied.push("拼写检查=不支持");
    }

    try {
      st.grammar = !!o.CheckGrammarAsYouType;
      o.CheckGrammarAsYouType = false;
      st.applied.push(!!o.CheckGrammarAsYouType === false ? "语法检查=关闭" : "语法检查=未生效");
    } catch (e3) {
      st.grammar = null;
      st.applied.push("语法检查=不支持");
    }
  } else {
    st.applied.push("Application.Options=不可用");
  }

  var win = null;
  try { win = Application.ActiveWindow; } catch (e4) {}
  if (win) {
    try {
      var vt = Number(win.View.Type);
      st.view = vt;
      if (vt === 3) {
        st.applied.push("视图=已是普通");
      } else {
        win.View.Type = 3;
        st.applied.push(Number(win.View.Type) === 3 ? "视图=已切普通" : "视图=切换未生效");
      }
    } catch (e5) {
      st.view = null;
      st.applied.push("视图=不支持");
    }
  } else {
    st.applied.push("ActiveWindow=不可用");
  }

  return st;
}

function resumeHeavyUpdates(st) {
  if (!st) return;

  var o = null;
  try { o = Application.Options; } catch (e0) {}
  if (o) {
    try { if (st.pagination !== null) o.Pagination = st.pagination; } catch (e1) {}
    try { if (st.spell !== null) o.CheckSpellingAsYouType = st.spell; } catch (e2) {}
    try { if (st.grammar !== null) o.CheckGrammarAsYouType = st.grammar; } catch (e3) {}
  }

  if (st.view !== null && st.view > 0) {
    try { Application.ActiveWindow.View.Type = st.view; } catch (e4) {}
  }
}

async function renumberAll(doc, collected, options) {
  options = options || {};
  if (!collected) collected = collectFormulas(doc);
  if (!collected) return { ok: false, count: 0, failed: 0, refs: 0 };
  var list = collected.formulas;

  if (options.scopeChapter) {
    var filtered = [];
    for (var q = 0; q < list.length; q++) {
      var fc = (list[q].chapter !== undefined) ? list[q].chapter : chapterValue(list[q].h1, list[q].h2);
      if (fc === options.scopeChapter) filtered.push(list[q]);
    }
    list = filtered;
  }

  if (list.length === 0) {
    log("WARN", "没有找到需要刷新的已编号公式");
    return { ok: true, count: 0, failed: 0, refs: 0 };
  }

  if (!options.skipConfirm && !collected.confirmed &&
      (collected.mixedCount > 0 || collected.inlineCount > 0 || collected.crossParaCount > 0)) {
    var msg = "将编号 " + list.length + " 个带编号公式：\r" +
      "混排需拆分 " + collected.mixedCount + " 个，行内公式 " + collected.inlineCount +
      " 个，跨段旧号 " + collected.crossParaCount + " 个。\r\r" +
      "拆分会在公式前后插入段落标记，改变原有排版。\r\r点“确定”继续，点“取消”放弃。";

    if (!askConfirm(msg, "EQNUM 确认")) {
      log("WARN", "用户取消编号");
      return { ok: false, cancelled: true, count: 0, failed: 0, refs: 0 };
    }
  }

  var nextId = maxBookmarkId(doc) + 1;
  for (var i = 0; i < list.length; i++) {
    if (!list[i].bookmarkName) list[i].bookmarkName = makeBookmarkName(nextId++);
  }

  var maxChars = 10;
  var allBms = listOurBookmarks(doc);
  for (var b = 0; b < allBms.length; b++) {
    if (allBms[b].text.length > maxChars) maxChars = allBms[b].text.length;
  }

  if (CFG.mode === "continuous" && !options.scopeChapter && list.length > 0) {
    list[0].startValue = CFG.start;
  }

  var done = 0;
  var failed = 0;
  var cancelled = false;
  var tAll = Date.now();
  var layout = { textWidth: getTextWidth(doc), fontSize: getBodyFontSize(doc) };

  log("INFO", "=== 全文重编号开始：候选 " + list.length +
    "，模式=" + CFG.mode + "，层级=" + CFG.level +
    "，刷新引用=" + CFG.refreshRefs + " ===");
  if (CFG.mode === "chapter") {
    log("INFO", "章节联动：STYLEREF " + CFG.level + " \\n + SEQ 公式 \\* ARABIC \\s " + CFG.level);
  } else {
    log("INFO", "连续编号：SEQ 公式 \\* ARABIC（首项按起始值重置）");
  }

  var oldScreen = true;
  try {
    var sv = Application.ScreenUpdating;
    if (sv !== undefined && sv !== null) oldScreen = !!sv;
  } catch (eScreen0) {}
  try { Application.ScreenUpdating = false; } catch (eScreen1) {}
  log("INFO", "屏幕更新已关闭");

  reportProgress(0, list.length, "准备编号（共 " + list.length + " 个）");

  var heavy = pauseHeavyUpdates();
  log("INFO", "性能开关：" + (heavy.applied.length ? heavy.applied.join("，") : "无"));

  try {
    // 倒序处理：前面的表格转换不会影响后面已捕获的 Range/OMath 对象。
    // 正序实测更慢（2556s vs 1883s），因为后续段落整体位移需要临时书签锚点，
    // 且每次编辑仍会触发整篇重排，已恢复倒序。
    for (var idx = list.length - 1; idx >= 0; idx--) {
      var f = list[idx];
      var seqNo = list.length - idx;
      var table = f.table;
      var tItem = Date.now();
      var tPrep = 0;
      var tLayout = 0;
      var tWrite = 0;
      var tDel = 0;

      log("INFO", "编号 " + seqNo + "/" + list.length + "：段 #" + f.para +
        "，类型=" + (f.kind || "plain") + " 开始");

      reportProgress(seqNo - 1, list.length,
        "第 " + seqNo + "/" + list.length + " 个：段 #" + f.para);

      if (f.kind === "plain") {
        if (f.trailingRange) {
          var tD = Date.now();
          // 旧编号一律删除，不做“合并进公式段”：WPS 上文本替换会留下残片/破坏结构
          status("段 #" + f.para + "：删除旧编号");
          removeOldNumber(doc, {
            range: f.trailingRange,
            text: f.trailing || "",
            sameParagraph: !!f.trailingSame,
            paraIndex: f.trailingPara || 0
          }, "重编号");
          try { f.p = f.om.Range.Paragraphs.Item(1); } catch (eRe) {}
          tDel = Date.now() - tD;
        }

        if (f.mixed || f.inline) {
          status("段 #" + f.para + "：拆分混排公式");
          log("INFO", "  段 #" + f.para + "：拆分开始");
          var iso = isolateFormula(doc, f.p, f.om);
          if (!iso.ok) {
            failed++;
            log("FAIL", "段 #" + f.para + " 拆分失败：" + iso.reason +
              "，用时 " + (Date.now() - tItem) + "ms");
            continue;
          }
          f.p = iso.p;
          f.om = iso.om;
        }

        status("段 #" + f.para + "：转两栏表格");
        log("INFO", "  段 #" + f.para + "：转表格开始");
        var t1 = Date.now();
        table = preparePlainFormula(doc, f);
        tPrep = Date.now() - t1;
        if (!table) {
          failed++;
          log("FAIL", "段 #" + f.para + " 转为两栏表格失败，转表用时 " + tPrep + "ms");
          continue;
        }
      }

      status("段 #" + f.para + "：排版");
      log("INFO", "  段 #" + f.para + "：排版开始");
      var t2 = Date.now();
      if (f.kind === "table") {
        applyTableWidths(table, maxChars, layout);
      } else {
        applyTableLayout(doc, table, maxChars, layout);
      }
      tLayout = Date.now() - t2;

      status("段 #" + f.para + "：写入编号");
      log("INFO", "  段 #" + f.para + "：写号开始");
      var t3 = Date.now();
      var field = writeNumberCell(doc, table, f.bookmarkName, CFG.level, f.startValue || 0, false);
      tWrite = Date.now() - t3;
      if (!field) {
        failed++;
        log("FAIL", "段 #" + f.para + " 写入编号失败，写号用时 " + tWrite + "ms");
        continue;
      }
      done++;

      log("INFO", "编号 " + seqNo + "/" + list.length + "：段 #" + f.para +
        " 完成，用时 " + (Date.now() - tItem) + "ms（旧号 " + tDel +
        "，转表 " + tPrep + "，排版 " + tLayout + "，写号 " + tWrite + "）");

      await eqnumUiGap();
      if (cancelRequested()) {
        log("WARN", "收到中止请求，停止后续编号；已完成 " + done + " 个");
        cancelled = true;
        break;
      }
    }

    status("");
    log("INFO", "=== 编号循环结束：成功 " + done + "，失败 " + failed +
      "，用时 " + (Date.now() - tAll) + "ms ===");

    if (cancelled) {
      reportProgress(done + failed, list.length, "已中止，仍在更新编号域");
    }

    reportProgress(list.length, list.length, "更新编号域");
    log("INFO", "=== 开始更新编号域 ===");
    updateOurSeqFields(doc);

    reportProgress(list.length, list.length, "刷新引用");
    log("INFO", "=== 开始更新引用域 ===");
    var refs = updateRefFields(doc, CFG.refreshRefs);

    log("INFO", "=== 全文重编号结束：成功 " + done + "，失败 " + failed +
      "，引用 " + refs + "，总用时 " + (Date.now() - tAll) + "ms ===");
    reportProgress(list.length, list.length,
      (cancelled ? "已中止：完成 " : "完成：成功 ") + done + " 个");
    return { ok: true, count: done, failed: failed, refs: refs, cancelled: cancelled };
  } finally {
    resumeHeavyUpdates(heavy);
    try { Application.ScreenUpdating = oldScreen; } catch (eScreen2) {}
    status("");
  }
}

function runDebugCandidates(doc) {
  section("候选诊断（只读）");

  var total = 0;
  try { total = Number(doc.Paragraphs.Count); } catch (e0) { return; }
  log("INFO", "开始诊断：段落总数 " + total + "，最多输出前 20 个候选");

  var found = 0;
  var scanned = 0;

  for (var i = 1; i <= total && found < 20; i++) {
    if (i % 500 === 0) {
      status("候选诊断：" + i + " / " + total);
      log("INFO", "已扫描段落 " + i + " / " + total + "，候选 " + found);
    }

    var p = null;
    try { p = doc.Paragraphs.Item(i); } catch (e1) { continue; }
    scanned++;

    var omCount = 0;
    try { omCount = Number(p.Range.OMaths.Count); } catch (e2) { continue; }
    if (omCount !== 1) continue;

    var om = null;
    try { om = p.Range.OMaths.Item(1); } catch (e3) { continue; }

    var trail = findTrailingNumber(doc, p, om, i);
    if (!trail.found) continue;

    found++;

    var ty = -1;
    try { ty = Number(om.Type); } catch (e4) {}
    var mixed = formulaMixed(doc, p, om, trail);

    log("INFO", "候选 #" + found + " 段 #" + i +
      " Type=" + ty +
      " 混排=" + (mixed ? "是" : "否") +
      " 旧号=" + trail.text +
      " 同段=" + (trail.sameParagraph ? "是" : "否") +
      " 编号段=#" + trail.paraIndex);

    var ptext = "";
    try { ptext = String(p.Range.Text); } catch (e5) {}
    log("INFO", "  本段=" + clip(ptext.replace(/[\r\u0007]/g, ""), 160));

    if (!trail.sameParagraph) {
      var tp = null;
      try { tp = doc.Paragraphs.Item(trail.paraIndex); } catch (e6) {}
      if (tp) {
        var ttext = "";
        try { ttext = String(tp.Range.Text); } catch (e7) {}
        var tOm = 0;
        var tShape = 0;
        try { tOm = Number(tp.Range.OMaths.Count); } catch (e8) {}
        try { tShape = Number(tp.Range.InlineShapes.Count); } catch (e9) {}
        log("INFO", "  编号段文本=" + clip(ttext.replace(/[\r\u0007]/g, ""), 160) +
          "（OMaths=" + tOm + "，InlineShapes=" + tShape + "）");
      }
    }
  }

  status("");
  log("INFO", "诊断完成：扫描 " + scanned + " 段，输出 " + found + " 个候选");
  alert("候选诊断完成，详见宏日志。\r最多输出前 20 个候选。");
}

async function runRenumber(doc) {
  reportProgress(0, 0, "扫描全文候选（此阶段界面可能短暂无响应）");
  var collected = collectFormulas(doc);
  if (!collected) return;

  var list = collected.formulas;
  if (list.length === 0) {
    alert("没有找到“公式后面带编号”的公式。\r当前规则只重编号已有编号的公式。");
    return;
  }

  var samePara = list.length - collected.existingTables - collected.crossParaCount;
  if (samePara < 0) samePara = 0;

  var msg = "共找到 " + list.length + " 个带编号的公式：\r" +
    "已有编号表 " + collected.existingTables + " 个；同段旧号 " + samePara +
    " 个；跨段旧号 " + collected.crossParaCount + " 个。\r" +
    "混排需拆分 " + collected.mixedCount + " 个，行内公式 " + collected.inlineCount + " 个。\r\r" +
    "将重建这些公式的编号，建议先备份。\r\r点“确定”继续，点“取消”放弃。";

  var preview = "";
  var pv = Math.min(5, list.length);
  for (var pi = 0; pi < pv; pi++) {
    var ff = list[pi];
    preview += "\r  段 #" + ff.para + " " + (ff.inline ? "行内" : "显示") +
      " 旧号 " + (ff.trailing || "(空)") +
      (ff.trailingSame ? " 同段" : " 跨段");
  }
  if (preview) {
    msg += "\r\r前 " + pv + " 个候选：" + preview;
    log("INFO", "候选预览（前 " + pv + " 个）：" + preview.replace(/\r/g, " | "));
  }

  if (list.length > 200 && collected.crossParaCount === list.length && collected.inlineCount === list.length) {
    msg += "\r\r注意：本次候选全部是行内公式且旧编号都在后续段落，可能存在误判，请先核对上面的候选。";
  } else if (list.length > 500) {
    msg += "\r\r候选数量较多（" + list.length + "），预计耗时较长，建议先备份或先只刷新当前章节。";
  }

  if (!askConfirm(msg, "EQNUM 全文重编号")) {
    alert("已取消");
    return;
  }

  collected.confirmed = true;
  var res = await renumberAll(doc, collected);
  if (res.cancelled) return;

  alert("全文重编号完成：\r编号 " + res.count + " 个，失败 " + res.failed + " 个，刷新引用 " + res.refs + " 个。");
}

function runUpdateRefs(doc) {
  if (!askConfirm("将刷新全文引用域，可能耗时较长。\r建议先备份。\r\r点“确定”继续，点“取消”放弃。",
      "EQNUM 刷新引用")) {
    return;
  }

  var n = updateRefFields(doc, "all");
  alert("引用刷新完成：" + n + " 个");
}

async function runRefresh(doc) {
  var curPara = null;
  var curChapter = "";
  try {
    var sel = Application.Selection;
    curPara = sel.Range.Paragraphs.Item(1);
    curChapter = chapterAt(doc, curPara, true);
  } catch (e0) {}

  if (typeof InputBox !== "function") {
    alert("当前宏环境不支持 InputBox，无法选择刷新范围");
    return;
  }

  var msg = "请选择刷新范围：\r1 = 只刷新当前章节已有编号" +
    (curChapter ? "（第 " + curChapter + " 章）" : "") +
    "\r2 = 刷新全文已有编号" +
    "\r3 = 扫描并重编号当前章节（首次处理/性能测试）" +
    "\r\r输入 1、2 或 3：";
  var ans = String(InputBox(msg, "EQNUM 刷新已编号公式", "1")).replace(/\s/g, "");

  var scope = "";
  if (ans === "1") scope = "chapter";
  else if (ans === "2") scope = "all";
  else if (ans === "3") scope = "chapterScan";
  else {
    alert("已取消");
    return;
  }

  if ((scope === "chapter" || scope === "chapterScan") && !curChapter) {
    alert("无法确定当前章节。\r请把光标放到章节内的公式上，或选择刷新全文。");
    return;
  }

  if (scope === "chapterScan") {
    reportProgress(0, 0, "扫描当前章节候选（此阶段界面可能短暂无响应）");
    var selScan = Application.Selection;
    var curPara = 1;
    try { curPara = Number(doc.Range(0, selScan.Range.Start).Paragraphs.Count) + 1; } catch (eSel) {}

    var bounds = chapterBounds(doc, curPara);
    log("INFO", "第 " + curChapter + " 章范围：段 #" + bounds.start + " – #" + bounds.end +
      "（章号上下文 h1=" + bounds.h1 + "，h2=" + bounds.h2 + "）");

    var scanned = collectFormulas(doc, {
      startPara: bounds.start,
      endPara: bounds.end,
      initH1: bounds.h1,
      initH2: bounds.h2
    });
    if (!scanned) return;

    var nChap = 0;
    for (var qi = 0; qi < scanned.formulas.length; qi++) {
      var qc = (scanned.formulas[qi].chapter !== undefined)
        ? scanned.formulas[qi].chapter
        : chapterValue(scanned.formulas[qi].h1, scanned.formulas[qi].h2);
      if (qc === curChapter) nChap++;
    }

    log("INFO", "第 " + curChapter + " 章候选 " + nChap + " 个，开始重编号（含未编号公式）");
    var resScan = await renumberAll(doc, scanned, {
      skipConfirm: true,
      scopeChapter: curChapter
    });
    if (resScan.cancelled) return;

    alert("当前章节重编号完成：\r章节：" + curChapter +
      "\r公式 " + resScan.count + " 个，失败 " + resScan.failed +
      " 个，刷新引用 " + resScan.refs + " 个。");
    return;
  }

  var collected = collectOurTables(doc, scope === "chapter" ? curChapter : "");
  var list = collected.formulas;
  if (list.length === 0) {
    alert(scope === "chapter"
      ? "当前章节没有本工具生成的编号公式"
      : "文档中没有本工具生成的编号公式");
    return;
  }

  var options = { skipConfirm: true };
  if (scope === "chapter") {
    options.scopeChapter = curChapter;
  }

  collected.confirmed = true;
  var res = await renumberAll(doc, collected, options);
  if (res.cancelled) return;

  alert("刷新完成：\r范围：" + (scope === "chapter" ? "当前章节 " + curChapter : "全文") +
    "\r公式 " + res.count + " 个，失败 " + res.failed + " 个，刷新引用 " + res.refs + " 个。");
}

function runCurrent(doc) {
  var sr = run(function () { return Application.Selection; });
  if (!sr.ok || !sr.value) {
    alert("无法获取光标位置");
    return;
  }
  var sel = sr.value;

  var selStart = 0;
  try { selStart = Number(sel.Range.Start); } catch (e0) {}

  function pickNearest(list, pos) {
    if (!list || list.length === 0) return null;
    var best = list[0];
    var bestD = 1e12;
    for (var k = 0; k < list.length; k++) {
      var st = 0, en = 0;
      try { st = Number(list[k].Range.Start); en = Number(list[k].Range.End); } catch (e) {}
      var d = (pos < st) ? (st - pos) : (pos > en ? pos - en : 0);
      if (d < bestD) { bestD = d; best = list[k]; }
    }
    return best;
  }

  var oms = [];
  var r = run(function () { return sel.Range.OMaths; });
  if (r.ok) {
    for (var i = 1; i <= r.value.Count; i++) {
      var om = null;
      try { om = r.value.Item(i); } catch (e1) { continue; }
      oms.push(om);
    }
  }

  if (oms.length === 0) {
    var pr = run(function () { return sel.Paragraphs.Item(1).Range.OMaths; });
    if (pr.ok) {
      for (i = 1; i <= pr.value.Count; i++) {
        var om2 = null;
        try { om2 = pr.value.Item(i); } catch (e3) { continue; }
        oms.push(om2);
      }
    }
  }

  if (oms.length === 0) {
    alert("光标附近没有公式");
    return;
  }

  var target = pickNearest(oms, selStart);
  if (!target) {
    alert("无法确定目标公式");
    return;
  }

  var p = null;
  try { p = target.Range.Paragraphs.Item(1); } catch (e5) {
    alert("无法读取公式所在段落");
    return;
  }

  var tCount = run(function () { return p.Range.Tables.Count; });
  if (tCount.ok && Number(tCount.value) > 0) {
    var ours = getOurTableInfo(p, doc);
    if (ours) {
      var b = findBookmark(doc, ours.bookmarkName);
      var txt = "";
      try { if (b) txt = String(b.Range.Text).replace(/[\r\u0007]/g, ""); } catch (e6) {}
      alert("该公式已编号：" + txt + "\r书签：" + ours.bookmarkName);
    } else {
      alert("该公式位于表格内，第一版暂不处理");
    }
    return;
  }

  // 删除公式后面已有的旧编号（同段或后续段落，不跨分页/分节符）
  var paraIndex = -1;
  try { paraIndex = Number(doc.Range(0, p.Range.Start).Paragraphs.Count) + 1; } catch (eIdx) {}

  if (paraIndex > 0) {
    var trail = findTrailingNumber(doc, p, target, paraIndex);
    if (trail.found && trail.range) {
      removeOldNumber(doc, trail, "当前");
      try { p = target.Range.Paragraphs.Item(1); } catch (eRe) {}
    }
  }

  var info = standaloneInfo(doc, p, target);
  if (!info.standalone) {
    var msg = "该公式与正文混排，编号会在公式前后插入段落标记，把它拆成独立段落并生成两栏表格。\r\r点“确定”继续，点“取消”放弃。";
    if (!askConfirm(msg, "EQNUM 编号当前")) return;

    var iso = isolateFormula(doc, p, target);
    if (!iso.ok) {
      alert("公式拆分失败：" + iso.reason);
      return;
    }
    p = iso.p;
    target = iso.om;
  }

  try {
    if (Number(target.Type) !== 0) {
      target.Type = 0;
      p = target.Range.Paragraphs.Item(1);
      target = p.Range.OMaths.Item(1);
    }
  } catch (eType) {}

  var f = { para: 0, p: p, om: target, kind: "plain", h1: "", h2: "" };
  var table = preparePlainFormula(doc, f);
  if (!table) {
    alert("公式转为两栏表格失败，请查看宏日志");
    return;
  }

  var maxChars = 10;
  var items = listOurBookmarks(doc);
  for (var k = 0; k < items.length; k++) {
    if (items[k].text.length > maxChars) maxChars = items[k].text.length;
  }

  applyTableLayout(doc, table, maxChars);

  var nextId = maxBookmarkId(doc) + 1;
  var name = makeBookmarkName(nextId);
  var startValue = 0;
  if (CFG.mode === "continuous" && listOurBookmarks(doc).length === 0) {
    startValue = CFG.start;
  }
  var field = writeNumberCell(doc, table, name, CFG.level, startValue);
  if (!field) {
    alert("写入编号失败，请查看宏日志。\r表格已生成但右列为空，可删除后重试。");
    return;
  }

  var shown = "";
  try { shown = String(table.Cell(1, 2).Range.Text).replace(/[\r\u0007]/g, ""); } catch (eShow) {}

  log("INFO", "当前公式已编号：" + shown + "，书签：" + name);
}

function findControlInCollection(controls, depth) {
  if (!controls || depth > 3) return null;

  var n = -1;
  try { n = Number(controls.Count); } catch (e0) { return null; }
  if (n <= 0) return null;

  for (var i = 1; i <= n; i++) {
    var c = null;
    try { c = controls.Item(i); } catch (e1) { continue; }

    var id = -1;
    try { id = Number(c.Id); } catch (e2) {}
    var cap = "";
    try { cap = String(c.Caption); } catch (e3) {}

    if (id === 775 || /交叉引用/.test(cap)) return c;

    var sub = null;
    try { sub = c.Controls; } catch (e4) {}
    if (sub) {
      var found = findControlInCollection(sub, depth + 1);
      if (found) return found;
    }
  }

  return null;
}

// 查找 WPS「插入 → 引用 → 交叉引用」菜单命令
function findCrossRefControl() {
  var cb = null;
  try { cb = Application.CommandBars; } catch (e0) { return null; }
  if (!cb) return null;

  // 优先扫描菜单栏 #41
  var preferred = [41];
  for (var k = 0; k < preferred.length; k++) {
    var bar = null;
    try { bar = cb.Item(preferred[k]); } catch (e1) {}
    if (!bar) continue;

    var c = findControlInCollection(bar.Controls, 0);
    if (c) return c;
  }

  // 备用：扫描全部 CommandBars（较慢）
  var count = -1;
  try { count = Number(cb.Count); } catch (e2) { return null; }

  for (var i = 1; i <= count; i++) {
    if (i === 41) continue;

    var b = null;
    try { b = cb.Item(i); } catch (e3) { continue; }

    var c2 = findControlInCollection(b.Controls, 0);
    if (c2) return c2;
  }

  return null;
}

function nextOurRefName() {
  var seq = 0;
  var tr = run(function () { return readProp("EQNUM_REF_SEQ", "0"); });
  if (tr.ok) seq = Number(tr.value);
  if (!(seq >= 0)) seq = 0;
  seq++;
  writeProp("EQNUM_REF_SEQ", String(seq));
  return "EQNUM_REF_" + pad4(seq);
}

function runInsertRef(doc) {
  // 优先打开 WPS 原生交叉引用对话框
  var nativeCtrl = findCrossRefControl();
  if (nativeCtrl) {
    var nr = run(function () { nativeCtrl.Execute(); return true; });
    if (nr.ok) {
      log("INFO", "已打开原生交叉引用对话框，请选择「公式」和整项题注");
      return;
    }
    log("WARN", "打开原生交叉引用对话框失败，回退到输入框：" + nr.error);
  } else {
    log("WARN", "未找到原生交叉引用命令，回退到输入框");
  }

  var items = listOurBookmarks(doc);
  if (items.length === 0) {
    alert("文档中还没有已编号的公式");
    return;
  }

  if (typeof InputBox !== "function") {
    alert("当前宏环境不支持 InputBox，无法选择引用");
    return;
  }

  var hint = "可引用的编号：\r";
  for (var i = 0; i < items.length && i < 40; i++) {
    hint += (i + 1) + ". " + items[i].text + "\r";
  }
  if (items.length > 40) hint += "……\r";
  hint += "\r请输入编号，例如 2.1 或 (2.1)";

  var ans = InputBox(hint, "EQNUM 插入引用", "");
  if (ans === null || String(ans).replace(/\s/g, "") === "") return;

  var want = String(ans).replace(/\s/g, "");
  var target = null;
  for (i = 0; i < items.length; i++) {
    var t = items[i].text.replace(/\s/g, "");
    var bare = t.replace(/[\(\)（）]/g, "");
    var wantBare = want.replace(/[\(\)（）]/g, "");
    if (t === want || t === "(" + want + ")" || bare === wantBare) {
      target = items[i];
      break;
    }
  }

  if (!target) {
    alert("没有找到编号：" + want);
    return;
  }

  var sel = Application.Selection;
  var start = Number(sel.Range.Start);
  var prefix = CFG.refPrefix || "";

  if (prefix) sel.Range.InsertAfter(prefix);

  var pos = start + prefix.length;
  var f = doc.Fields.Add(doc.Range(pos, pos), -1, " REF " + target.name + " \\h ", false);
  var ur = run(function () { f.Update(); return true; });
  if (!ur.ok) log("WARN", "引用域更新失败：" + ur.error);

  var refName = nextOurRefName();
  var br = run(function () { doc.Bookmarks.Add(refName, f.Range); return true; });
  if (br.ok) {
    log("INFO", "引用已加书签：" + refName);
  } else {
    log("WARN", "引用书签创建失败：" + br.error);
  }

  alert("已插入引用：" + prefix + target.text);
}


/* ============================================================
 * 入口分发
 * ============================================================ */
async function main() {
  // 加载项页面也会加载本文件，此时没有宏入口传入的 mode，直接返回
  if (typeof mode === "undefined" || !mode) return;

  log("INFO", "EQNUM 版本：2026-11-24-v11（删除旧编号不再合并；制表符插段尾，表格外包公式）");

  if (mode === "settings") {
    var d0 = ensureDocument();
    if (d0) {
      loadConfig();
      runSettings(d0);
    }
    return;
  }

  var doc = ensureDocument();
  if (!doc) return;
  loadConfig();

  try {
    if (mode === "current") return await runCurrent(doc);
    else if (mode === "renumber") return await runRenumber(doc);
    else if (mode === "refresh") return await runRefresh(doc);
    else if (mode === "insertref") return await runInsertRef(doc);
    else if (mode === "updaterefs") return await runUpdateRefs(doc);
    else if (mode === "debugscan") return await runDebugCandidates(doc);
    else alert("未知操作：" + mode);
  } catch (e) {
    alert("执行异常：" + errText(e));
    consoleLog("执行异常：" + errText(e));
  }
}

return main();
}
