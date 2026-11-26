/*
 * EQNUM WPS 加载项入口
 *
 * 约定（同官方模板）：index.html 引入 main.js；对象模型走 window.Application。
 * 功能区只负责打开任务窗格并投递动作，真正的编号逻辑在 ui/panel.html 里运行。
 */

var EQNUM_ADDON_VERSION = "0.2.8";
var g_ribbonUI = null;

// 诊断：每个功能区回调都先 alert 一下，便于定位问题（稳定后关闭）
var EQNUM_DIAG = false;

// 加载项根目录 URL，用于拼 ShowDialog / CreateTaskPane 的相对路径
function GetUrlPath() {
  var e = document.location.toString();
  e = decodeURI(e);
  if (e.indexOf("/") !== -1) e = e.substring(0, e.lastIndexOf("/"));
  return e;
}

function OnAddinLoad(ribbonUI) {
  try {
    g_ribbonUI = ribbonUI;
    try { window.Application.ribbonUI = ribbonUI; } catch (e) {}
    try { window.Application.PluginStorage.setItem("EQNUM_LOADED", EQNUM_ADDON_VERSION); } catch (e2) {}
    try { console.log("[EQNUM] 加载项已载入，版本 " + EQNUM_ADDON_VERSION + "，路径 " + GetUrlPath()); } catch (e3) {}
  } catch (eAll) {
    try { alert("EQNUM OnAddinLoad 异常：" + eqnumErr(eAll)); } catch (ex) {}
  }
  return true;
}

function OnGetEnabled(control) { return true; }
function OnGetVisible(control) { return true; }

// 功能区按钮图标。WPS 只认 getImage 回调，不认 ribbon.xml 的 image 属性。
// 返回相对路径时，根路径是加载项根目录。
function GetImage(control) {
  var id = "";
  try { id = String(control.Id); } catch (e) {}
  var file = "";
  if (id === "btnCurrent") file = "images/eqnum-current.png";
  else if (id === "btnInsertRef") file = "images/eqnum-insert-ref.png";
  if (!file) return "";
  try { console.log("[EQNUM] GetImage " + id + " -> " + file); } catch (e2) {}
  return file;
}

function eqnumApp() { return window.Application; }

function eqnumErr(e) {
  if (!e) return "未知错误";
  try { if (e.message) return String(e.message); } catch (e1) {}
  try { if (e.description) return String(e.description); } catch (e2) {}
  return String(e);
}

// 打开（或显示）任务窗格
function eqnumShowPanel() {
  try {
    var app = eqnumApp();
    if (!app) { alert("eqnumShowPanel: window.Application 为空"); return null; }

    var id = null;
    try { id = app.PluginStorage.getItem("EQNUM_PANEL_ID"); } catch (eP) {}

    if (id) {
      try {
        var tp = app.GetTaskPane(id);
        if (tp) {
          tp.Visible = true;
          return tp;
        }
      } catch (eG) {
        try { alert("GetTaskPane(" + id + ") 失败：" + eqnumErr(eG)); } catch (ea) {}
      }
    }

    var url = GetUrlPath() + "/ui/panel.html";
    try { console.log("[EQNUM] CreateTaskPane url=" + url); } catch (el) {}
    var pane = app.CreateTaskPane(url);
    try { app.PluginStorage.setItem("EQNUM_PANEL_ID", pane.ID); } catch (eS) {}
    try { pane.DockPosition = 2; } catch (eD) {}
    pane.Visible = true;
    return pane;
  } catch (eAll) {
    try { alert("eqnumShowPanel 异常：" + eqnumErr(eAll)); } catch (ex) {}
    return null;
  }
}

// 投递动作给面板，由面板在自己的页面里执行（进度/日志都在面板里）
function eqnumQueue(action, scope) {
  try {
    var app = eqnumApp();
    if (!app) { alert("eqnumQueue: window.Application 为空"); return; }
    var payload = JSON.stringify({ action: action, scope: scope || "", ts: Date.now() });
    try { app.PluginStorage.setItem("EQNUM_ACTION", payload); } catch (e) {
      alert("无法把操作传给面板：" + eqnumErr(e));
    }
    eqnumShowPanel();
  } catch (eAll) {
    try { alert("eqnumQueue 异常：" + eqnumErr(eAll)); } catch (ex) {}
  }
}

function OnBtnCurrent() {
  try { if (EQNUM_DIAG) alert("[EQNUM] 点击：编号当前公式"); } catch (e) {}
  eqnumQueue("current");
}
function OnBtnChapter() {
  try { if (EQNUM_DIAG) alert("[EQNUM] 点击：本章重编号"); } catch (e) {}
  eqnumQueue("refresh", "3");
}
function OnBtnRenumber() {
  try { if (EQNUM_DIAG) alert("[EQNUM] 点击：全文重编号"); } catch (e) {}
  eqnumQueue("renumber");
}
function OnBtnInsertRef() {
  try { if (EQNUM_DIAG) alert("[EQNUM] 点击：插入引用"); } catch (e) {}
  eqnumQueue("insertref");
}
function OnBtnUpdateRefs() {
  try { if (EQNUM_DIAG) alert("[EQNUM] 点击：刷新引用"); } catch (e) {}
  eqnumQueue("updaterefs");
}
function OnBtnSettings() {
  try { if (EQNUM_DIAG) alert("[EQNUM] 点击：编号设置"); } catch (e) {}
  eqnumShowPanel();
}
