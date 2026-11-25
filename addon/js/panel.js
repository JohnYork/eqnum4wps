/*
 * EQNUM 任务窗格逻辑
 *
 * 页面加载顺序（见 panel.html）：
 *   shim.js       → 提供 Application 别名、拦截 confirm/alert/console
 *   wps-eqnum.js  → 宏版核心逻辑（EQNUM_CURRENT / EQNUM_RENUMBER / ...）
 *   panel.js      → 本文件，负责界面、设置存储、动作调度、日志展示
 */

var EQNUM_PANEL_VERSION = "0.2.8";
var MAX_LOG_LINES = 1500;
var g_busy = false;
var g_progText = "";

function uiApp() { return window.Application; }

function uiDoc() {
  try {
    var d = uiApp().ActiveDocument;
    return d ? d : null;
  } catch (e) {
    return null;
  }
}

function uiErr(e) {
  if (!e) return "未知错误";
  try { if (e.message) return String(e.message); } catch (e1) {}
  try { if (e.description) return String(e.description); } catch (e2) {}
  return String(e);
}

function $(id) { return document.getElementById(id); }

function bindDrawer(drawerId, toggleId, storageKey) {
  var drawer = $(drawerId);
  var toggle = $(toggleId);
  if (!toggle || !drawer) return;
  toggle.onclick = function () {
    drawer.classList.toggle("collapsed");
    try { localStorage.setItem(storageKey, String(!drawer.classList.contains("collapsed"))); } catch (e) {}
  };
  try {
    var open = localStorage.getItem(storageKey);
    if (open === "true") drawer.classList.remove("collapsed");
  } catch (e) {}
}

function propGet(doc, name, def) {
  try {
    var props = doc.CustomDocumentProperties;
    for (var i = 1; i <= props.Count; i++) {
      var p = props.Item(i);
      if (String(p.Name) === name) return String(p.Value);
    }
  } catch (e) {}
  return def;
}

function propSet(doc, name, value) {
  try {
    var props = doc.CustomDocumentProperties;
    for (var i = 1; i <= props.Count; i++) {
      var p = props.Item(i);
      if (String(p.Name) === name) {
        p.Value = String(value);
        return true;
      }
    }
    props.Add(name, false, 4, String(value));
    return true;
  } catch (e) {
    appendLog({ text: "保存属性失败 " + name + "：" + uiErr(e), level: "err" });
    return false;
  }
}

function appendLog(item) {
  var area = $("logArea");
  if (!area) return;
  var line = document.createElement("div");
  line.className = item.level || "log";
  line.textContent = item.text;
  area.appendChild(line);
  while (area.childNodes.length > MAX_LOG_LINES) {
    area.removeChild(area.firstChild);
  }
  area.scrollTop = area.scrollHeight;
}

// 核心逻辑（wps-eqnum.js）通过该回调上报进度：done/total 为已处理数/总数，
// total <= 0 表示当前阶段总量未知，只在文字上提示。
function setProgress(done, total, text) {
  g_progText = String(text || "");

  var pct = 0;
  if (total > 0) {
    pct = Math.round(Number(done) / Number(total) * 100);
    if (!(pct >= 0)) pct = 0;
    if (pct > 100) pct = 100;
  }

  var bar = $("progBar");
  if (bar) {
    bar.style.width = pct + "%";
    bar.className = (pct >= 100) ? "bar-in done" : "bar-in";
  }

  var txt = $("progText");
  if (txt) {
    txt.textContent = g_progText +
      (total > 0 ? "（" + done + "/" + total + "，" + pct + "%）" : "");
  }
}

window.EQNUM_PROGRESS = setProgress;

function setBusy(busy, label) {
  g_busy = !!busy;
  var ids = ["btnCurrent", "btnChapter", "btnRenumber", "btnInsertRef", "btnUpdateRefs", "btnSave", "btnStats"];
  for (var i = 0; i < ids.length; i++) {
    var el = $(ids[i]);
    if (el) el.disabled = g_busy;
  }
  var cancel = $("btnCancel");
  if (cancel) cancel.disabled = !g_busy;
  var st = $("stateText");
  if (st) {
    st.textContent = g_busy ? ("运行中：" + (label || "")) : "空闲";
    st.className = g_busy ? "state busy" : "state";
  }
}

function refreshDocInfo() {
  var doc = uiDoc();
  var el = $("docText");
  if (!el) return;
  if (!doc) {
    el.textContent = "没有打开的文档";
    return;
  }
  var name = "";
  try { name = String(doc.Name); } catch (e) {}
  el.textContent = "当前文档：" + name;
}

// 统计本工具生成的编号书签数（EQNUM_NUM_），用于判断文档是否已处理过
function countOurBookmarks(doc) {
  var n = 0;
  try {
    var bms = doc.Bookmarks;
    for (var i = 1; i <= bms.Count; i++) {
      if (String(bms.Item(i).Name).indexOf("EQNUM_NUM_") === 0) n++;
    }
  } catch (e) {
    return -1;
  }
  return n;
}

function refreshStats() {
  var doc = uiDoc();
  function set(id, v) { var el = $(id); if (el) el.textContent = v; }
  if (!doc) {
    set("stPara", "-"); set("stOm", "-"); set("stField", "-");
    set("stBm", "-"); set("stTable", "-"); set("stProp", "-"); set("stOurs", "-");
    return;
  }
  function count(fn) {
    try { var n = fn(); return (n === null || n === undefined) ? "-" : String(n); } catch (e) { return "-"; }
  }
  set("stPara", count(function () { return doc.Paragraphs.Count; }));
  set("stOm", count(function () { return doc.OMaths.Count; }));
  set("stField", count(function () { return doc.Fields.Count; }));
  set("stBm", count(function () { return doc.Bookmarks.Count; }));
  set("stTable", count(function () { return doc.Tables.Count; }));
  set("stProp", count(function () { return doc.CustomDocumentProperties.Count; }));

  var ours = countOurBookmarks(doc);
  set("stOurs", ours < 0 ? "-" : String(ours));

  var hint = $("hintOurs");
  if (hint) {
    if (ours === 0) {
      hint.textContent = "当前文档还没有本工具生成的编号：请点「本章重编号」或「全文重编号」。";
    } else if (ours > 0) {
      hint.textContent = "本工具已生成 " + ours + " 个编号；「本章重编号」会扫描并重排该章节。";
    } else {
      hint.textContent = "";
    }
  }
}

function loadSettings() {
  var doc = uiDoc();
  if (!doc) return;

  var mode = propGet(doc, "EQNUM_MODE", "chapter");
  var level = propGet(doc, "EQNUM_LEVEL", "1");
  var start = propGet(doc, "EQNUM_START", "1");
  var prefix = propGet(doc, "EQNUM_REF_PREFIX", "式");
  var refresh = propGet(doc, "EQNUM_REFRESH_REFS", "ours");

  $("selMode").value = (mode === "continuous") ? "continuous" : "chapter";
  $("selLevel").value = (level === "2") ? "2" : "1";
  $("inpStart").value = String(Number(start) >= 1 ? Number(start) : 1);
  $("inpRefPrefix").value = prefix;
  $("selRefreshRefs").value = (refresh === "all" || refresh === "none") ? refresh : "ours";
}

function saveSettings() {
  var doc = uiDoc();
  if (!doc) {
    appendLog({ text: "请先打开文档", level: "err" });
    return;
  }

  var start = Number($("inpStart").value);
  if (!(start >= 1)) start = 1;
  $("inpStart").value = String(start);

  var ok = true;
  ok = propSet(doc, "EQNUM_MODE", $("selMode").value) && ok;
  ok = propSet(doc, "EQNUM_LEVEL", $("selLevel").value) && ok;
  ok = propSet(doc, "EQNUM_START", String(start)) && ok;
  ok = propSet(doc, "EQNUM_REF_PREFIX", $("inpRefPrefix").value) && ok;
  ok = propSet(doc, "EQNUM_REFRESH_REFS", $("selRefreshRefs").value) && ok;

  if (ok) {
    appendLog({
      text: "设置已保存：模式=" + $("selMode").value +
        "，层级=" + $("selLevel").value +
        "，起始值=" + start +
        "，引用前缀=" + $("inpRefPrefix").value +
        "，刷新引用=" + $("selRefreshRefs").value,
      level: "info"
    });
  }
}

function actionLabel(kind, scope) {
  if (kind === "current") return "编号当前公式";
  if (kind === "renumber") return "全文重编号";
  if (kind === "refresh") {
    if (String(scope) === "3") return "本章重编号";
    if (String(scope) === "2") return "刷新全文已编号";
    return "刷新当前章节已编号";
  }
  if (kind === "insertref") return "插入引用";
  if (kind === "updaterefs") return "刷新引用";
  return kind;
}

function runAction(kind, scope) {
  if (g_busy) {
    appendLog({ text: "上一个任务还在执行，请稍候", level: "warn" });
    return;
  }

  var doc = uiDoc();
  if (!doc) {
    appendLog({ text: "请先打开文档", level: "err" });
    return;
  }

  if (kind === "refresh" && (scope === "1" || scope === "2") && countOurBookmarks(doc) === 0) {
    appendLog({
      text: "提示：「刷新」只处理本工具已生成编号的公式，当前文档还没有；首次请选「扫描并重编号当前章节」。",
      level: "warn"
    });
  }

  var label = actionLabel(kind, scope);
  setBusy(true, label);
  setProgress(0, 0, "准备：" + label);
  appendLog({ text: "=== " + label + " 开始 ===", level: "info" });

  window.EQNUM_CANCEL = false;

  function finish() {
    appendLog({ text: "=== " + label + " 结束 ===", level: "info" });
    window.EQNUM_CANCEL = false;
    setBusy(false);
    // 核心没有上报过任何进度时，给一个结束提示，避免一直停在“准备”
    if (/^准备/.test(g_progText)) setProgress(0, 0, "结束：" + label);
    refreshDocInfo();
    refreshStats();
  }

  // 先让界面把“运行中”画出来，再进入 COM 调用
  setTimeout(function () {
    var ret = null;
    try {
      if (kind === "current") ret = window.EQNUM_CURRENT();
      else if (kind === "renumber") ret = window.EQNUM_RENUMBER();
      else if (kind === "refresh") {
        window.__EQNUM_INPUT_QUEUE.push(String(scope || "1"));
        ret = window.EQNUM_REFRESH();
      }
      else if (kind === "insertref") ret = window.EQNUM_INSERT_REF();
      else if (kind === "updaterefs") ret = window.EQNUM_UPDATE_REFS();
    } catch (e) {
      appendLog({ text: "执行异常：" + uiErr(e), level: "err" });
      finish();
      return;
    }

    if (ret && typeof ret.then === "function") {
      ret.then(finish, function (e) {
        appendLog({ text: "执行异常：" + uiErr(e), level: "err" });
        finish();
      });
    } else {
      finish();
    }
  }, 30);
}

// 功能区按钮通过 PluginStorage 投递动作，面板轮询后执行
function pollAction() {
  if (g_busy) return;

  var raw = null;
  try { raw = uiApp().PluginStorage.getItem("EQNUM_ACTION"); } catch (e) { return; }
  if (!raw) return;
  try { uiApp().PluginStorage.setItem("EQNUM_ACTION", ""); } catch (e) {}

  var req = null;
  try { req = JSON.parse(raw); } catch (e) { return; }
  if (!req || !req.action) return;
  if (req.action === "settings") return;

  runAction(req.action, req.scope);
}

function bind() {
  // 设置抽屉切换
  bindDrawer("drawerSettings", "toggleSettings", "EQNUM_DRAWER_OPEN");
  bindDrawer("drawerStats",   "toggleStats",   "EQNUM_DRAWER_STATS");
  bindDrawer("drawerLog",     "toggleLog",     "EQNUM_DRAWER_LOG");

  $("btnSave").onclick = function () { saveSettings(); };
  $("btnReload").onclick = function () {
    loadSettings();
    appendLog({ text: "已重新读取文档中的设置", level: "info" });
  };
  $("btnStats").onclick = function () {
    refreshDocInfo();
    refreshStats();
    appendLog({ text: "统计已刷新", level: "info" });
  };
  $("btnClearLog").onclick = function () {
    $("logArea").innerHTML = "";
  };

  // 5 个主操作按钮
  $("btnCurrent").onclick = function () { runAction("current"); };
  $("btnChapter").onclick = function () { runAction("refresh", "3"); };  // 扫描并重编号当前章节
  $("btnRenumber").onclick = function () { runAction("renumber"); };
  $("btnInsertRef").onclick = function () { runAction("insertref"); };
  $("btnUpdateRefs").onclick = function () { runAction("updaterefs"); };

  $("btnCancel").onclick = function () {
    if (!g_busy) return;
    window.EQNUM_CANCEL = true;
    setProgress(0, 0, "已请求中止：当前公式处理完后停止");
    appendLog({ text: "已请求中止：当前公式处理完后停止", level: "warn" });
  };

  $("chkAutoConfirm").onchange = function () {
    window.__EQNUM_AUTO_CONFIRM = !!$("chkAutoConfirm").checked;
    appendLog({
      text: "执行前确认框：" + (window.__EQNUM_AUTO_CONFIRM ? "已关闭" : "已开启"),
      level: "info"
    });
  };

  $("selGap").onchange = function () {
    var gap = Number($("selGap").value);
    window.EQNUM_UI_GAP = gap;
    try { localStorage.setItem("EQNUM_UI_GAP", String(gap)); } catch (e) {}
    appendLog({ text: "界面响应间隔：" + gap + " ms", level: "info" });
  };
}

// 界面响应间隔：每条公式后给 WPS 主线程留出的重绘时间
function initGapSetting() {
  var gap = 50;
  try {
    var saved = localStorage.getItem("EQNUM_UI_GAP");
    if (saved !== null && saved !== "") gap = Number(saved);
  } catch (e) {}
  if (!(gap >= 0)) gap = 50;

  var sel = $("selGap");
  if (sel) {
    var found = false;
    for (var i = 0; i < sel.options.length; i++) {
      if (Number(sel.options[i].value) === gap) { found = true; break; }
    }
    if (found) sel.value = String(gap);
    else sel.value = "50";
  }
  window.EQNUM_UI_GAP = Number(sel ? sel.value : gap);
}

function init() {
  $("verText").textContent = "v" + EQNUM_PANEL_VERSION;

  if (window.EQNUM_BUS) {
    for (var i = 0; i < EQNUM_BUS.lines.length; i++) appendLog(EQNUM_BUS.lines[i]);
    EQNUM_BUS.on(function (item) { appendLog(item); });
  }

  bind();
  initGapSetting();
  window.__EQNUM_AUTO_CONFIRM = !!$("chkAutoConfirm").checked;

  refreshDocInfo();
  loadSettings();
  refreshStats();
  appendLog({ text: "面板就绪，核心版本见下方日志", level: "info" });

  setInterval(pollAction, 700);
}

init();
