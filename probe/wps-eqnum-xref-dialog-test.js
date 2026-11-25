/*
 * WPS 交叉引用对话框探测
 *
 * 目的：确认 JSA 是否能：
 *   1. 访问 Application.Dialogs
 *   2. 找到“交叉引用”对话框
 *   3. 设置引用类型为“公式”
 *   4. 打开该对话框
 *
 * 安全：不修改文档；只有在用户确认后才会打开对话框。
 */

function EQNUM_XREF_DIALOG_TEST() {
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

function dialogName(dlg) {
  var name = "";
  try { name = String(dlg.Name); } catch (e1) {}
  return name;
}

function dialogType(dlg) {
  var t = "";
  try { t = String(dlg.Type); } catch (e1) {}
  return t;
}

function main() {
  section("零、脚本信息");
  log("INFO", "开始时间：" + new Date().toString());
  log("INFO", "本脚本不修改文档。");

  section("一、Application.Dialogs 是否可用");
  var dr = run(function () { return Application.Dialogs; });
  if (!dr.ok) {
    log("FAIL", "Application.Dialogs 不可用：" + dr.error);
    writeReport();
    return;
  }
  log("PASS", "Application.Dialogs 可用");

  var count = -1;
  var cr = run(function () { return Number(Application.Dialogs.Count); });
  if (cr.ok) count = cr.value;
  log("INFO", "Dialogs.Count = " + count);

  if (count <= 0) {
    log("WARN", "Dialogs 数量为 0 或无法读取，无法按名称枚举");
  }

  section("二、枚举对话框名称");
  var found = -1;
  var foundName = "";
  var limit = (count > 0) ? Math.min(count, 80) : 0;

  for (var i = 1; i <= limit; i++) {
    var ir = run(function () { return Application.Dialogs.Item(i); });
    if (!ir.ok) continue;

    var name = dialogName(ir.value);
    var type = dialogType(ir.value);

    if (name || type) {
      log("INFO", "#" + i + " 名称=" + name + " 类型=" + type);
    }

    if (/交叉|引用|Cross|Reference/i.test(name)) {
      found = i;
      foundName = name;
    }
  }

  if (found > 0) {
    log("PASS", "找到可能的交叉引用对话框：#" + found + " " + foundName);
  } else {
    log("WARN", "没有从名称中找到交叉引用对话框");
  }

  section("三、已知候选索引 41（Word 的 InsertCrossReference）");
  var candidate = -1;
  var c41 = run(function () { return Application.Dialogs.Item(41); });
  if (c41.ok) {
    candidate = 41;
    log("INFO", "#41 名称=" + dialogName(c41.value) + " 类型=" + dialogType(c41.value));
  } else {
    log("WARN", "Dialogs.Item(41) 不可用：" + c41.error);
  }

  if (candidate < 0 && found > 0) candidate = found;

  if (candidate > 0) {
    var cd = run(function () { return Application.Dialogs.Item(candidate); });
    if (cd.ok) {
      var dlg = cd.value;

      var setType = run(function () { dlg.ReferenceType = "公式"; return true; });
      log(setType.ok ? "PASS" : "WARN",
        "设置 ReferenceType=\"公式\"：" + (setType.ok ? "成功" : setType.error));

      var setKind = run(function () { dlg.ReferenceKind = 1; return true; });
      log(setKind.ok ? "PASS" : "WARN",
        "设置 ReferenceKind=1：" + (setKind.ok ? "成功" : setKind.error));

      var setItem = run(function () { dlg.ReferenceItem = 1; return true; });
      log(setItem.ok ? "PASS" : "WARN",
        "设置 ReferenceItem=1：" + (setItem.ok ? "成功" : setItem.error));
    }
  }

  section("四、打开对话框（需要你确认）");
  if (candidate <= 0) {
    log("WARN", "没有可用候选，跳过打开测试");
  } else {
    var msg = "找到候选对话框 #" + candidate + "（" + foundName + "）。\r" +
      "是否现在打开它测试？\r\r点“确定”打开，点“取消”跳过。";
    var ans = "";
    if (typeof confirm === "function") {
      ans = confirm(msg) ? "YES" : "";
    } else if (typeof InputBox === "function") {
      ans = InputBox(msg, "EQNUM 交叉引用对话框测试", "YES");
    }

    if (String(ans).replace(/\s/g, "").toUpperCase() === "YES") {
      var sd = run(function () {
        var d = Application.Dialogs.Item(candidate);
        return d.Show();
      });
      if (sd.ok) {
        log("PASS", "对话框已打开并已关闭，Show 返回=" + String(sd.value));
      } else {
        log("FAIL", "打开对话框失败：" + sd.error);
      }
    } else {
      log("INFO", "用户跳过打开测试");
    }
  }

  section("五、结论");
  log("INFO", "如果 Dialogs 可用且能打开 #41/找到的索引，EQNUM_INSERT_REF 可以改为打开原生对话框。");
  log("INFO", "ReferenceType 设置成功则能预选「公式」；失败则只能依赖对话框记住上次选择。");

  writeReport();

  try {
    if (typeof alert === "function") alert("交叉引用对话框测试完成，详见报告文档。");
  } catch (e9) {}
}

function writeReport() {
  var text = "EQNUM 交叉引用对话框测试报告\r时间：" + new Date().toString() + "\r" + reportLines.join("\r");
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
