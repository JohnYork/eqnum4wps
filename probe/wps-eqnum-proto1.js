/*
 * WPS 公式自动编号 —— 单公式原型验证（只在新文档中执行）
 *
 * 验证链路：
 *   1. 用 OMaths.Add 创建 OMML 公式，并设为显示公式
 *   2. 公式所在段落转换为两栏无边框表格
 *   3. 右列插入 SEQ 域作为编号
 *   4. 设置表格宽度、边框、对齐、垂直居中
 *   5. 编号加书签，正文插 REF 域并更新
 *
 * 安全约定：
 *   - 不读取、不修改当前文档，全部操作在新建文档中完成。
 *   - 新文档保留打开，方便肉眼检查排版，不需要保存。
 *
 * 使用：在宏编辑器中运行入口函数 EQNUM_PROTO1()。
 */

function EQNUM_PROTO1() {
var reportLines = [];
var passCount = 0;
var failCount = 0;

function consoleLog(s) {
  try {
    if (typeof console !== "undefined" && console && console.log) {
      console.log(s);
    }
  } catch (e) {}
}

function log(level, msg) {
  if (level === "PASS") passCount++;
  else if (level === "FAIL") failCount++;

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
  try {
    return { ok: true, value: fn() };
  } catch (e) {
    return { ok: false, error: errText(e) };
  }
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
  try {
    t = String(v);
  } catch (e) {
    t = "<无法转字符串>";
  }
  return esc(clip(t, 120));
}

function addDocument() {
  var beforeCount = -1;
  try { beforeCount = Application.Documents.Count; } catch (e) {}

  var r = run(function () { return Application.Documents.Add(); });
  if (r.ok && r.value) return r.value;

  var afterCount = -1;
  try { afterCount = Application.Documents.Count; } catch (e2) {}
  if (beforeCount >= 0 && afterCount === beforeCount + 1) {
    try { return Application.ActiveDocument; } catch (e3) {}
  }
  return null;
}

function endRange(d) {
  var end = 0;
  try { end = d.Content.End - 1; } catch (e) { end = 0; }
  if (end < 0) end = 0;
  return d.Range(end, end);
}

function cellOMathCount(cell) {
  var r = run(function () { return cell.Range.OMaths.Count; });
  return r.ok ? Number(r.value) : -1;
}

/* ============================================================
 * 1. 创建公式
 * ============================================================ */
function createEquation(d) {
  section("一、创建 OMML 公式");

  var before = 0;
  try { before = d.OMaths.Count; } catch (e0) {}

  // 新文档只有一个空段落，把选区放到段首再插入公式
  d.Range(0, 0).Select();

  var addR = run(function () { return d.OMaths.Add(d.Range(0, 0)); });
  log("INFO", "OMaths.Add " + (addR.ok ? "调用成功" : "调用失败：" + addR.error) +
    "，返回值类型=" + typeof addR.value + "，值=" + toText(addR.value));

  var after = -1;
  try { after = d.OMaths.Count; } catch (e1) {}
  log("INFO", "公式数量：调用前 " + before + "，调用后 " + after);

  if (after <= 0) {
    log("FAIL", "OMaths.Add 未产生公式对象");
    return null;
  }

  // 注意：WPS 的 Add 返回值不是 OMath 对象，必须按索引重新取
  var om = run(function () { return d.OMaths.Item(after); }).value;
  if (!om) {
    log("FAIL", "无法按索引取到 OMath 对象");
    return null;
  }

  var type = "";
  try { type = String(om.Type); } catch (e2) {}
  log("INFO", "取到公式对象，Type=" + type);

  var content = "";
  try { content = String(om.Range.Text); } catch (e3) {}
  log("INFO", "公式初始内容=" + esc(clip(content, 60)) + "，长度=" + content.length);

  // 内容为空：写入线性公式文本
  if (content.length === 0) {
    var setR = run(function () {
      om.Range.Text = "x^2+y^2=1";
      return String(om.Range.Text);
    });
    if (setR.ok) {
      log("PASS", "已向公式写入线性文本，读回=" + esc(clip(String(setR.value), 60)));
    } else {
      var insR = run(function () {
        om.Range.InsertAfter("x^2+y^2=1");
        return String(om.Range.Text);
      });
      if (insR.ok) {
        log("PASS", "InsertAfter 写入公式文本成功，读回=" + esc(clip(String(insR.value), 60)));
      } else {
        log("FAIL", "无法向空公式写入内容：" + setR.error + " / " + insR.error);
      }
    }
  }

  // 线性格式转专业格式
  var bu = run(function () { d.OMaths.BuildUp(); return true; });
  if (bu.ok) {
    log("PASS", "OMaths.BuildUp() 可调用");
  } else {
    var bu2 = run(function () { om.BuildUp(); return true; });
    if (bu2.ok) log("PASS", "OMath.BuildUp() 可调用");
    else log("WARN", "BuildUp 不可用（公式可能仍为线性格式）");
  }

  // BuildUp 后重新取对象，避免引用失效
  var om2 = run(function () { return d.OMaths.Item(d.OMaths.Count); }).value;
  if (om2) om = om2;

  content = "";
  try { content = String(om.Range.Text); } catch (e4) {}
  log("INFO", "公式最终内容=" + esc(clip(content, 60)) + "，长度=" + content.length);

  if (content.length === 0) {
    log("FAIL", "公式仍为空，无法继续验证排版链路");
    return null;
  }

  var tr = run(function () { om.Type = 0; return om.Type; });
  if (tr.ok) {
    log("PASS", "公式可设为显示公式，Type=" + String(tr.value));
  } else {
    log("WARN", "设置 Type=0 失败：" + tr.error);
  }

  // 重新取一次，Type 变更后 Range 可能变化
  var om3 = run(function () { return d.OMaths.Item(d.OMaths.Count); }).value;
  if (om3) om = om3;

  var pText = "";
  try { pText = String(om.Range.Paragraphs.Item(1).Range.Text); } catch (e5) {}
  log("INFO", "公式所在段落文本=" + esc(clip(pText, 80)));

  return om;
}

/* ============================================================
 * 2. 段落转两栏表格
 * ============================================================ */
function convertToTwoColumnTable(d, om) {
  section("二、段落转为两栏表格");

  var converted = null;

  // 在公式后插入制表符，作为两栏分隔
  var ir = run(function () {
    var endPos = om.Range.End;
    d.Range(endPos, endPos).InsertAfter("\t");
    return true;
  });
  if (ir.ok) log("PASS", "公式后插入制表符");
  else log("FAIL", "插入制表符失败：" + ir.error);

  var paraText = "";
  try { paraText = String(om.Range.Paragraphs.Item(1).Range.Text); } catch (e) {}
  log("INFO", "待转换段落文本=" + esc(clip(paraText, 80)));

  var tblBefore = -1;
  try { tblBefore = d.Tables.Count; } catch (e1) {}

  var cr = run(function () {
    var paraRange = om.Range.Paragraphs.Item(1).Range;
    return paraRange.ConvertToTable("\t");
  });

  if (cr.ok && cr.value) {
    converted = cr.value;
    log("PASS", "Range.ConvertToTable 返回了表格对象");
  } else {
    log("FAIL", "ConvertToTable 直接调用失败：" + (cr.error || "返回空"));
    var tblAfter = -1;
    try { tblAfter = d.Tables.Count; } catch (e2) {}
    if (tblAfter > tblBefore) {
      converted = run(function () { return d.Tables.Item(tblAfter); }).value;
      if (converted) log("PASS", "ConvertToTable 未返回对象，但表格数量增加，已取到最后一个表格");
    }
  }

  if (!converted) {
    log("WARN", "ConvertToTable 不可用，尝试 FormattedText 迁移作为备选方案");

    var alt = run(function () {
      var target = endRange(d);
      var t = d.Tables.Add(target, 1, 2);
      var paraRange = om.Range.Paragraphs.Item(1).Range;
      t.Cell(1, 1).Range.FormattedText = paraRange.FormattedText;
      return t;
    });

    if (alt.ok && alt.value) {
      converted = alt.value;
      log("PASS", "FormattedText 迁移可调用（原段落未删除，仅供能力验证）");
      log("INFO", "备选表格左单元格 OMaths 数=" + cellOMathCount(converted.Cell(1, 1)));
      return converted;
    }

    log("FAIL", "两种转换方式都不可用：" + (alt.error || "返回空"));
    return null;
  }

  var rows = -1;
  var cols = -1;
  try { rows = converted.Rows.Count; } catch (e3) {}
  try { cols = converted.Columns.Count; } catch (e4) {}
  log("INFO", "表格结构：行=" + rows + "，列=" + cols);

  log("INFO", "左单元格 OMaths 数=" + cellOMathCount(converted.Cell(1, 1)));
  log("INFO", "右单元格 OMaths 数=" + cellOMathCount(converted.Cell(1, 2)));

  if (cellOMathCount(converted.Cell(1, 1)) === 1) {
    log("PASS", "公式完整保留在左单元格");
  } else {
    log("FAIL", "公式未保留在左单元格");
  }

  return converted;
}

/* ============================================================
 * 3. 右列插入 SEQ 编号
 * ============================================================ */
function insertSeqNumber(d, tbl) {
  section("三、右列插入 SEQ 编号");

  var field = null;
  var fr = run(function () {
    return d.Fields.Add(tbl.Cell(1, 2).Range, -1, " SEQ EQNUM_PROTO \\* ARABIC ", false);
  });

  if (!fr.ok) {
    log("FAIL", "在单元格内插入 SEQ 域失败：" + fr.error);
    return null;
  }
  field = fr.value;
  log("PASS", "单元格内可插入 SEQ 域");

  var ur = run(function () {
    d.Fields.Update();
    return true;
  });
  if (ur.ok) log("PASS", "Fields.Update 可调用");
  else log("FAIL", "Fields.Update 失败：" + ur.error);

  var result = "";
  var code = "";
  try { result = String(field.Result.Text); } catch (e1) {}
  try { code = String(field.Code.Text); } catch (e2) {}
  log("INFO", "SEQ 域代码=" + esc(code) + "，结果=" + esc(result));

  if (result.replace(/\s/g, "") === "1") {
    log("PASS", "SEQ 编号生成正确：1");
  } else {
    log("FAIL", "SEQ 编号结果异常：" + esc(result));
  }

  return field;
}

/* ============================================================
 * 4. 表格排版
 * ============================================================ */
function formatTable(tbl) {
  section("四、两栏表格排版");

  var br = run(function () {
    tbl.Borders.Enable = 0;
    return tbl.Borders.Enable;
  });
  if (br.ok) log("PASS", "表格边框已关闭，Enable=" + String(br.value));
  else log("FAIL", "关闭边框失败：" + br.error);

  var wr = run(function () {
    tbl.PreferredWidthType = 2;
    tbl.PreferredWidth = 100;
    return tbl.PreferredWidthType + "/" + tbl.PreferredWidth;
  });
  if (wr.ok) log("PASS", "表格宽度设为 100%，读回=" + String(wr.value));
  else log("FAIL", "设置表格宽度失败：" + wr.error);

  var vr = run(function () {
    tbl.Cell(1, 1).VerticalAlignment = 1;
    tbl.Cell(1, 2).VerticalAlignment = 1;
    return tbl.Cell(1, 1).VerticalAlignment + "/" + tbl.Cell(1, 2).VerticalAlignment;
  });
  if (vr.ok) log("PASS", "单元格垂直居中，读回=" + String(vr.value));
  else log("FAIL", "设置垂直居中失败：" + vr.error);

  var ar = run(function () {
    tbl.Cell(1, 1).Range.ParagraphFormat.Alignment = 1;
    tbl.Cell(1, 2).Range.ParagraphFormat.Alignment = 2;
    return tbl.Cell(1, 1).Range.ParagraphFormat.Alignment + "/" + tbl.Cell(1, 2).Range.ParagraphFormat.Alignment;
  });
  if (ar.ok) log("PASS", "左列居中、右列右对齐，读回=" + String(ar.value));
  else log("FAIL", "设置水平对齐失败：" + ar.error);

  var pr = run(function () {
    tbl.TopPadding = 0;
    tbl.BottomPadding = 0;
    tbl.LeftPadding = 0;
    tbl.RightPadding = 0;
    return true;
  });
  if (pr.ok) log("PASS", "单元格内边距设为 0");
  else log("FAIL", "设置内边距失败：" + pr.error);

  var cr = run(function () {
    tbl.Cell(1, 2).Width = 70;
    return tbl.Cell(1, 2).Width;
  });
  if (cr.ok) log("INFO", "右列宽度可设置，读回=" + String(cr.value));
  else log("WARN", "右列宽度设置失败：" + cr.error);
}

/* ============================================================
 * 5. 书签与 REF 引用
 * ============================================================ */
function testReference(d, field) {
  section("五、书签与 REF 引用");

  // 记录编号域的起止位置，在域前后插入括号，再把整段 "(1)" 作为书签
  var seqStart = -1;
  var seqEnd = -1;
  try {
    seqStart = Number(field.Result.Start);
    seqEnd = Number(field.Result.End);
  } catch (e0) {}

  if (seqStart < 0 || seqEnd < seqStart) {
    log("FAIL", "无法取得 SEQ 域结果范围");
    return;
  }
  log("INFO", "SEQ 域结果范围=" + seqStart + "-" + seqEnd);

  var p1 = run(function () {
    d.Range(seqStart, seqStart).InsertBefore("(");
    return true;
  });
  if (p1.ok) log("PASS", "编号左侧括号插入成功");
  else log("FAIL", "编号左侧括号插入失败：" + p1.error);

  var p2 = run(function () {
    d.Range(seqEnd + 1, seqEnd + 1).InsertBefore(")");
    return true;
  });
  if (p2.ok) log("PASS", "编号右侧括号插入成功");
  else log("FAIL", "编号右侧括号插入失败：" + p2.error);

  var bm = run(function () {
    return d.Bookmarks.Add("EQNUM_PROTO_NUM", d.Range(seqStart, seqEnd + 2));
  });
  if (bm.ok) log("PASS", "编号书签创建成功（覆盖左右括号）");
  else {
    log("FAIL", "编号书签创建失败：" + bm.error);
    return;
  }

  var exists = run(function () {
    return d.Bookmarks.Exists("EQNUM_PROTO_NUM");
  });
  log(exists.ok && exists.value ? "PASS" : "FAIL", "书签存在性检查=" + (exists.ok ? String(exists.value) : exists.error));

  d.Content.InsertAfter("\r正文引用测试：式 ");

  var refField = null;
  var fr = run(function () {
    return d.Fields.Add(endRange(d), -1, " REF EQNUM_PROTO_NUM \\h ", false);
  });
  if (!fr.ok) {
    log("FAIL", "插入 REF 域失败：" + fr.error);
    return;
  }
  refField = fr.value;
  log("PASS", "REF 域插入成功");

  // 只更新 REF 域，避免重新更新 SEQ 导致书签失效
  var ur = run(function () {
    refField.Update();
    return true;
  });
  if (ur.ok) {
    log("PASS", "Field.Update 可调用");
  } else {
    var ur2 = run(function () {
      d.Fields.Update();
      return true;
    });
    if (ur2.ok) log("WARN", "Field.Update 不可用，已回退到 Fields.Update");
    else log("FAIL", "引用更新失败：" + ur2.error);
  }

  var exists2 = run(function () {
    return d.Bookmarks.Exists("EQNUM_PROTO_NUM");
  });
  log("INFO", "更新后书签是否存在=" + (exists2.ok ? String(exists2.value) : exists2.error));

  var result = "";
  var code = "";
  try { result = String(refField.Result.Text); } catch (e1) {}
  try { code = String(refField.Code.Text); } catch (e2) {}
  log("INFO", "REF 域代码=" + esc(code) + "，结果=" + esc(result));

  if (result.replace(/\s/g, "") === "(1)") {
    log("PASS", "REF 引用显示正确：(1)");
  } else {
    log("FAIL", "REF 引用结果异常：" + esc(result));
  }
}

/* ============================================================
 * 6. 把结果写入原型文档
 * ============================================================ */
function writeReport(d) {
  section("六、汇总");

  var summary = "PASS=" + passCount + "，FAIL=" + failCount;
  log("INFO", summary);
  log("INFO", "说明：FAIL 仅表示该步骤不可用，不代表脚本执行失败。");
  log("INFO", "本文档是临时原型，不需要保存。");

  var br = run(function () {
    endRange(d).InsertBreak(7); // 分页符
    return true;
  });
  if (!br.ok) {
    try { d.Content.InsertAfter("\f"); } catch (e) {}
  }

  var text = "===== 原型验证结果 =====\r时间：" + new Date().toString() + "\r" + reportLines.join("\r");
  var ir = run(function () {
    d.Content.InsertAfter("\r" + text);
    return true;
  });
  if (!ir.ok) log("FAIL", "写入结果文本失败：" + ir.error);

  try {
    var all = d.Content;
    all.Font.Name = "Consolas";
    all.Font.Size = 10;
  } catch (e2) {}
}

/* ============================================================
 * 主流程
 * ============================================================ */
function main() {
  section("零、脚本信息");
  log("INFO", "开始时间：" + new Date().toString());
  log("INFO", "原型只在新文档中执行，不修改当前文档。");

  var d = addDocument();
  if (!d) {
    log("FAIL", "无法创建临时文档，原型无法执行");
    return;
  }

  log("INFO", "临时文档名称=" + (run(function () { return d.Name; }).value || ""));

  var om = createEquation(d);
  if (!om) {
    writeReport(d);
    return;
  }

  var tbl = convertToTwoColumnTable(d, om);
  if (!tbl) {
    writeReport(d);
    return;
  }

  var field = insertSeqNumber(d, tbl);
  formatTable(tbl);
  if (field) testReference(d, field);

  writeReport(d);

  try {
    if (typeof alert === "function") {
      alert("EQNUM 单公式原型完成：PASS " + passCount + "，FAIL " + failCount +
        "。\r\n请查看新建的原型文档，重点看公式是否居中、编号是否在右列。");
    }
  } catch (e) {}
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
