/*
 * WPS 交叉引用命令探测（CommandBars.ExecuteMso）
 *
 * 目的：尝试通过功能区命令打开交叉引用对话框。
 * 备选命令：
 *   InsertCrossReference
 *   CrossReference
 *   InsertCrossRef
 * 再备选：
 *   Application.Run("InsertCrossReference")
 *
 * 安全：不修改文档；打开对话框可能需要你手动关闭。
 */

function EQNUM_COMMANDBAR_TEST() {
var reportLines = [];

function consoleLog(s) {
  try {
    if (typeof console !== "undefined" && console && console.log) console.log(s);
  } catch (e) {}
}

function log(level, msg) {
  var line = "[" + level + "] " + msg;
  reportLines.push(line);
  consoleLog(line);
}

function section(title) {
  reportLines.push("");
  reportLines.push("=== " + title + " ===");
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

function addDocument() {
  var before = -1;
  try { before = Application.Documents.Count; } catch (e) {}
  var r = run(function () { return Application.Documents.Add(); });
  if (r.ok && r.value) return r.value;
  var after = -1;
  try { after = Application.Documents.Count; } catch (e2) {}
  if (before >= 0 && after === before + 1) {
    try { return Application.ActiveDocument; } catch (e3) {}
  }
  return null;
}

function confirmGo(msg) {
  if (typeof confirm === "function") return confirm(msg);
  if (typeof InputBox === "function") {
    var a = InputBox(msg + "\r\r输入 YES 继续", "EQNUM 命令测试", "YES");
    return String(a).replace(/\s/g, "").toUpperCase() === "YES";
  }
  return false;
}

function main() {
  section("零、脚本信息");
  log("INFO", "开始时间：" + new Date().toString());
  log("INFO", "本脚本不修改文档；可能会打开对话框，需要你关闭。");

  section("一、CommandBars 是否可用");
  var cb = run(function () { return Application.CommandBars; });
  if (!cb.ok) {
    log("FAIL", "Application.CommandBars 不可用：" + cb.error);
    writeReport();
    return;
  }
  log("PASS", "Application.CommandBars 可用");

  var count = -1;
  var cr = run(function () { return Number(Application.CommandBars.Count); });
  if (cr.ok) count = cr.value;
  log("INFO", "CommandBars.Count = " + count);

  var em = run(function () { return typeof Application.CommandBars.ExecuteMso; });
  log("INFO", "ExecuteMso 类型 = " + (em.ok ? em.value : em.error));

  section("二、依次尝试五种方式");
  var attempts = [
    { id: 1, label: 'ExecuteMso("InsertCrossReference")', fn: function () { Application.CommandBars.ExecuteMso("InsertCrossReference"); } },
    { id: 2, label: 'ExecuteMso("CrossReference")', fn: function () { Application.CommandBars.ExecuteMso("CrossReference"); } },
    { id: 3, label: 'ExecuteMso("InsertCrossRef")', fn: function () { Application.CommandBars.ExecuteMso("InsertCrossRef"); } },
    { id: 4, label: 'Application.Run("InsertCrossReference")', fn: function () { Application.Run("InsertCrossReference"); } },
    { id: 5, label: 'Application.Run("InsertCrossRef")', fn: function () { Application.Run("InsertCrossRef"); } }
  ];

  if (!em.ok || em.value !== "function") {
    log("WARN", "ExecuteMso 不是函数，只尝试 Application.Run");
  }

  if (!confirmGo("将依次尝试 5 种方式。\r\r如果某一步弹出了「交叉引用」对话框，请先关闭它，脚本才会继续。\r最后会询问是哪一种方式真正弹出了对话框。\r\r点“确定”继续。")) {
    log("INFO", "用户取消测试");
  } else {
    for (var i = 0; i < attempts.length; i++) {
      if (attempts[i].id <= 3 && (!em.ok || em.value !== "function")) continue;

      log("INFO", "尝试 " + attempts[i].label + " ...");
      var r = run(attempts[i].fn);
      if (r.ok) log("INFO", attempts[i].label + " 调用返回成功");
      else log("WARN", attempts[i].label + " 报错：" + r.error);
    }

    var fb = "";
    if (typeof InputBox === "function") {
      fb = InputBox("刚才哪一种方式真正弹出了「交叉引用」对话框？\r" +
        "1 = ExecuteMso(InsertCrossReference)\r" +
        "2 = ExecuteMso(CrossReference)\r" +
        "3 = ExecuteMso(InsertCrossRef)\r" +
        "4 = Run(InsertCrossReference)\r" +
        "5 = Run(InsertCrossRef)\r" +
        "0 = 都没有\r\r请输入序号：",
        "EQNUM 结果确认", "0");
    } else if (typeof confirm === "function") {
      fb = confirm("刚才是否弹出了交叉引用对话框？") ? "1" : "0";
    }
    log("INFO", "用户反馈：弹窗方式序号=" + String(fb));
  }

  section("三、结论");
  log("INFO", "如果反馈某个序号真的弹出了「交叉引用」对话框，就可以把该命令接入 EQNUM_INSERT_REF。");
  log("INFO", "如果序号为 0，说明这些命令 ID 在 WPS 中都不生效，应放弃原生对话框。");

  writeReport();

  try {
    if (typeof alert === "function") alert("CommandBars 测试完成，详见报告文档。");
  } catch (e9) {}
}

function writeReport() {
  var text = "EQNUM CommandBars 测试报告\r时间：" + new Date().toString() + "\r" + reportLines.join("\r");
  var rd = addDocument();
  if (rd) {
    try {
      rd.Content.Text = text;
      try { rd.Content.Font.Name = "Consolas"; } catch (e1) {}
      try { rd.Content.Font.Size = 10; } catch (e2) {}
    } catch (e3) {}
  }
}

try {
  main();
} catch (e) {
  consoleLog("脚本执行异常：" + errText(e));
  try {
    if (typeof alert === "function") alert("脚本执行异常：" + errText(e));
  } catch (e2) {}
}
}
