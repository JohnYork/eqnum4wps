/*
 * WPS CommandBars 控件扫描：查找“交叉引用”命令
 *
 * 目的：
 *   遍历 Application.CommandBars 的控件，按 Caption/Tag/OnAction 查找
 *   “交叉引用 / Cross Reference”菜单项，并尝试执行。
 *
 * 安全：不修改文档；执行命令可能会打开对话框，需要手动关闭。
 */

function EQNUM_COMMANDBAR_SCAN() {
var reportLines = [];
var scannedBars = 0;
var scannedControls = 0;
var matches = [];

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
    var a = InputBox(msg + "\r\r输入 YES 继续", "EQNUM 控件扫描", "YES");
    return String(a).replace(/\s/g, "").toUpperCase() === "YES";
  }
  return false;
}

function safeStr(obj, prop) {
  try {
    var v = obj[prop];
    if (v === undefined || v === null) return "";
    return String(v);
  } catch (e) { return ""; }
}

function safeNum(obj, prop) {
  try {
    var v = obj[prop];
    if (v === undefined || v === null) return -1;
    return Number(v);
  } catch (e) { return -1; }
}

function scanControls(controls, barLabel, depth) {
  if (!controls || depth > 2) return;

  var n = -1;
  try { n = Number(controls.Count); } catch (e0) { return; }
  if (n <= 0) return;

  for (var i = 1; i <= n; i++) {
    if (scannedControls >= 8000) return;

    var c = null;
    try { c = controls.Item(i); } catch (e1) { continue; }
    scannedControls++;

    var caption = safeStr(c, "Caption");
    var tag = safeStr(c, "Tag");
    var onAction = safeStr(c, "OnAction");
    var id = safeNum(c, "Id");
    var type = safeNum(c, "Type");

    var text = caption + " | " + tag + " | " + onAction;
    if (/交叉|引用|Cross|Reference/i.test(text)) {
      var m = {
        bar: barLabel,
        caption: caption,
        tag: tag,
        onAction: onAction,
        id: id,
        type: type,
        control: c
      };
      matches.push(m);
      log("INFO", "命中：" + barLabel +
        " 控件#" + i +
        " 标题=" + caption +
        " Id=" + id +
        " Type=" + type +
        " Tag=" + tag +
        " OnAction=" + onAction);
    }

    // 有子控件的继续扫描（弹出菜单等）
    var sub = null;
    try { sub = c.Controls; } catch (e2) {}
    if (sub) {
      var sn = -1;
      try { sn = Number(sub.Count); } catch (e3) {}
      if (sn > 0) scanControls(sub, barLabel + " / " + caption, depth + 1);
    }
  }
}

function main() {
  section("零、脚本信息");
  log("INFO", "开始时间：" + new Date().toString());
  log("INFO", "本脚本不修改文档；扫描可能较慢。");

  section("一、扫描 CommandBars 控件");
  var cb = run(function () { return Application.CommandBars; });
  if (!cb.ok) {
    log("FAIL", "Application.CommandBars 不可用：" + cb.error);
    writeReport();
    return;
  }

  var barCount = -1;
  try { barCount = Number(cb.value.Count); } catch (e0) {}
  log("INFO", "CommandBars.Count = " + barCount);

  if (barCount <= 0) {
    log("FAIL", "没有可扫描的 CommandBars");
    writeReport();
    return;
  }

  for (var i = 1; i <= barCount; i++) {
    if (scannedControls >= 8000) {
      log("WARN", "控件数量达到上限，停止扫描");
      break;
    }
    if (i % 25 === 0) log("INFO", "已扫描 CommandBars " + i + " / " + barCount);

    var bar = null;
    try { bar = cb.value.Item(i); } catch (e1) { continue; }
    scannedBars++;

    var name = safeStr(bar, "Name");
    var nameLocal = safeStr(bar, "NameLocal");
    var visible = safeStr(bar, "Visible");
    var label = "#" + i + " " + (nameLocal || name);

    var controls = null;
    try { controls = bar.Controls; } catch (e2) { continue; }
    scanControls(controls, label, 0);
  }

  log("INFO", "扫描完成：CommandBars=" + scannedBars + "，控件=" + scannedControls +
    "，命中=" + matches.length);

  section("二、命中结果");
  if (matches.length === 0) {
    log("FAIL", "没有找到“交叉引用”相关控件");
  } else {
    for (var k = 0; k < matches.length; k++) {
      var m = matches[k];
      log("INFO", "#" + (k + 1) + " " + m.bar +
        " 标题=" + m.caption +
        " Id=" + m.id +
        " Type=" + m.type +
        " OnAction=" + m.onAction);
    }
  }

  section("三、执行真正的“交叉引用”命令");
  var target = null;
  for (var k = 0; k < matches.length; k++) {
    if (matches[k].type !== 10 && /交叉引用/.test(matches[k].caption)) {
      target = matches[k];
      break;
    }
  }
  if (!target) {
    for (k = 0; k < matches.length; k++) {
      if (matches[k].type !== 10) { target = matches[k]; break; }
    }
  }
  if (!target && matches.length > 0) target = matches[0];

  var fc = run(function () { return Application.CommandBars.FindControl(0, 775); });
  if (fc.ok && fc.value) log("PASS", "CommandBars.FindControl(0, 775) 可用");
  else log("WARN", "CommandBars.FindControl(0, 775) 不可用：" + (fc.error || "返回空"));

  if (target) {
    var msg = "将执行：\r" + target.bar + "\r标题=" + target.caption +
      "\rId=" + target.id + "\rType=" + target.type +
      "\r\r可能会打开「交叉引用」对话框，请检查引用类型并关闭它。\r点“确定”执行。";

    if (confirmGo(msg)) {
      var exec = run(function () { target.control.Execute(); return true; });
      if (exec.ok) {
        log("INFO", "控件 Execute() 调用成功");
      } else {
        log("WARN", "控件 Execute() 失败：" + exec.error);
      }

      if (!exec.ok && target.onAction) {
        var ar = run(function () { Application.Run(target.onAction); return true; });
        if (ar.ok) log("INFO", "Application.Run(OnAction) 调用成功");
        else log("WARN", "Application.Run(OnAction) 失败：" + ar.error);
      }

      var fb = "";
      if (typeof InputBox === "function") {
        fb = InputBox("刚才是否弹出了「交叉引用」对话框？\r1=是\r0=否", "EQNUM 结果确认", "0");
      } else if (typeof confirm === "function") {
        fb = confirm("刚才是否弹出了交叉引用对话框？") ? "1" : "0";
      }
      log("INFO", "用户反馈：是否弹窗=" + String(fb));
    } else {
      log("INFO", "用户取消执行");
    }
  } else {
    log("INFO", "没有候选，跳过执行");
  }

  section("四、结论");
  log("INFO", "如果命中项执行后弹出了交叉引用对话框，就可以把该控件的 Caption/Id/OnAction 接入 EQNUM_INSERT_REF。");
  log("INFO", "如果没有命中或执行无效，说明 WPS 没有把交叉引用命令暴露到 CommandBars。");

  writeReport();

  try {
    if (typeof alert === "function") alert("CommandBars 扫描完成，详见报告文档。");
  } catch (e9) {}
}

function writeReport() {
  var text = "EQNUM CommandBars 扫描报告\r时间：" + new Date().toString() + "\r" + reportLines.join("\r");
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
