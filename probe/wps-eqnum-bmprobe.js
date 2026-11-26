/*
 * 书签 API 探针（只读，不修改文档）
 *
 * 目的：确认 WPS 文字 Linux 下 Bookmarks 集合的正确用法。
 * 背景：collectFormulas / runCurrent 用 bookmarksInRange 读旧编号上的
 *       _RefXXXX 书签，但实测一个都读不到，导致 REF 域无法改指向。
 *
 * 运行：打开含公式编号与交叉引用的文档，光标放在任意位置，
 *       运行 EQNUM_BM_PROBE()，把弹窗内容发回。
 */

function EQNUM_BM_PROBE() {
  var doc = null;
  try { doc = Application.ActiveDocument; } catch (e0) {}
  if (!doc) { alert("请先打开文档"); return; }

  var lines = [];

  function P(s) {
    lines.push(s);
    try { console.log("[BMPROBE] " + s); } catch (e) {}
  }

  // 统一取错：返回字符串
  function err(e) {
    if (!e) return "未知";
    try { if (e.message) return String(e.message); } catch (e1) {}
    try { if (e.description) return String(e.description); } catch (e2) {}
    try { return String(e); } catch (e3) { return "无法读取"; }
  }

  function T(label, fn) {
    try {
      var v = fn();
      P(label + " -> " + (v === null ? "null" : (typeof v === "undefined" ? "undefined" : String(v))));
      return v;
    } catch (e) {
      P(label + " -> ERR " + err(e));
      return undefined;
    }
  }

  // ---------- 1. 集合基本能力 ----------
  P("=== 1. Bookmarks 集合 ===");
  var cnt = T("doc.Bookmarks.Count", function () { return doc.Bookmarks.Count; });
  var n = Number(cnt);
  if (!(n >= 0)) n = 0;
  var show = n > 40 ? 40 : n;
  P("（共 " + n + " 个，最多列出 " + show + " 个）");

  for (var i = 1; i <= show; i++) {
    var line = "#" + i;
    var bm = null;
    try { bm = doc.Bookmarks.Item(i); } catch (e1) { line += " Item(i) ERR " + err(e1); P(line); continue; }

    try { line += " name=" + String(bm.Name); } catch (e2) { line += " name=ERR"; }
    try { line += " start=" + Number(bm.Start) + " end=" + Number(bm.End); }
    catch (e3) {
      line += " Start/End=ERR";
      try { line += " range=" + Number(bm.Range.Start) + ".." + Number(bm.Range.End); }
      catch (e4) { line += " Range=ERR"; }
    }
    try {
      var rt = String(bm.Range.Text);
      rt = rt.replace(/[\r\n]/g, "\\n");
      if (rt.length > 24) rt = rt.substring(0, 24) + "…";
      line += " text=" + rt;
    } catch (e5) { line += " text=ERR"; }
    P(line);
  }

  // ---------- 2. Item(name) / Exists ----------
  P("");
  P("=== 2. 按名访问 ===");
  var firstName = "";
  try { firstName = String(doc.Bookmarks.Item(1).Name); } catch (e6) {}
  if (!firstName) {
    P("没有可用书签，跳过");
  } else {
    T("Item(\"" + firstName + "\").Name", function () { return doc.Bookmarks.Item(firstName).Name; });
    T("Exists(\"" + firstName + "\")", function () { return doc.Bookmarks.Exists(firstName); });
    T("Exists(\"_不存在的书签_\")", function () { return doc.Bookmarks.Exists("_不存在的书签_"); });
  }

  // ---------- 3. Range.Bookmarks ----------
  P("");
  P("=== 3. Range.Bookmarks ===");
  var sel = null;
  try { sel = Application.Selection.Range; } catch (e7) {}
  if (!sel) {
    P("无法取得 Selection.Range");
  } else {
    var p = null;
    try { p = sel.Paragraphs.Item(1); } catch (e8) {}
    if (p) {
      var ps = -1, pe = -1;
      try { ps = Number(p.Range.Start); pe = Number(p.Range.End); } catch (e9) {}
      P("光标段落 range = " + ps + ".." + pe);

      T("p.Range.Bookmarks.Count", function () { return p.Range.Bookmarks.Count; });
      T("p.Range.Bookmarks(1).Name", function () { return p.Range.Bookmarks.Item(1).Name; });

      // 逐个数全文档书签，用 InRange 判断是否在当前段
      if (n > 0) {
        var hit = [];
        for (var j = 1; j <= n; j++) {
          try {
            var b = doc.Bookmarks.Item(j);
            var nm = String(b.Name);
            var inr = p.Range.InRange(b.Range);
            if (inr) hit.push(nm);
          } catch (e10) {}
        }
        P("InRange 命中当前段：" + (hit.length ? hit.join(", ") : "（无）"));
      }
    } else {
      P("无法取得光标段落");
    }
  }

  // ---------- 4. 域里的书签引用 ----------
  P("");
  P("=== 4. 域代码抽样 ===");
  var fc = T("doc.Fields.Count", function () { return doc.Fields.Count; });
  var fn = Number(fc);
  if (!(fn >= 0)) fn = 0;
  var fs = fn > 20 ? 20 : fn;
  for (var k = 1; k <= fs; k++) {
    try {
      var f = doc.Fields.Item(k);
      var code = String(f.Code.Text).replace(/[\r\n]/g, " ");
      if (code.length > 60) code = code.substring(0, 60) + "…";
      P("field#" + k + " type=" + f.Type + " code=" + code);
    } catch (e11) { P("field#" + k + " ERR " + err(e11)); }
  }

  alert("书签 API 探针完成（" + lines.length + " 行）\r\r" +
    "如窗口太窄，可打开 ALT+F12 控制台查看 [BMPROBE] 输出。\r\r" +
    lines.slice(0, 30).join("\r"));
}

/*
 * 第二轮：隐藏书签（_Ref / _Toc）是否被集合排除，能否按名访问
 * 只读，不改文档。运行 EQNUM_BM_PROBE2()。
 */
function EQNUM_BM_PROBE2() {
  var doc = null;
  try { doc = Application.ActiveDocument; } catch (e0) {}
  if (!doc) { alert("请先打开文档"); return; }

  var lines = [];
  function P(s) { lines.push(s); try { console.log("[BMPROBE2] " + s); } catch (e) {} }
  function err(e) {
    if (!e) return "未知";
    try { if (e.message) return String(e.message); } catch (e1) {}
    try { if (e.description) return String(e.description); } catch (e2) {}
    return "无法读取";
  }
  function T(label, fn) {
    try { var v = fn(); P(label + " -> " + String(v)); return v; }
    catch (e) { P(label + " -> ERR " + err(e)); return undefined; }
  }

  P("=== 1. ShowHidden ===");
  T("Bookmarks.ShowHidden（读）", function () { return doc.Bookmarks.ShowHidden; });
  T("设为 true", function () { doc.Bookmarks.ShowHidden = true; return doc.Bookmarks.ShowHidden; });
  T("Count（设为 true 后）", function () { return doc.Bookmarks.Count; });

  P("");
  P("=== 2. 从 REF 域收集书签名 ===");
  var names = [];
  var seen = {};
  var fc = 0;
  try { fc = Number(doc.Fields.Count); } catch (eF) {}
  for (var i = 1; i <= fc; i++) {
    try {
      var code = String(doc.Fields.Item(i).Code.Text);
      var m = code.match(/\bREF\s+([^\s\\]+)/i);
      if (m && !seen[m[1]]) { seen[m[1]] = true; names.push(m[1]); }
    } catch (eG) {}
  }
  P("REF 域引用的书签：" + (names.length ? names.join(", ") : "（无）"));

  P("");
  P("=== 3. 按名访问 ===");
  for (var k = 0; k < names.length && k < 5; k++) {
    (function (nm) {
      T("Exists(" + nm + ")", function () { return doc.Bookmarks.Exists(nm); });
      T("Item(" + nm + ").Name", function () { return doc.Bookmarks.Item(nm).Name; });
      T("Item(" + nm + ").Start/End", function () {
        var b = doc.Bookmarks.Item(nm);
        return Number(b.Start) + "/" + Number(b.End);
      });
      T("Item(" + nm + ").Range.Text", function () {
        var t = String(doc.Bookmarks.Item(nm).Range.Text).replace(/[\r\n]/g, "\\n");
        return t.length > 24 ? t.substring(0, 24) + "…" : t;
      });
    })(names[k]);
  }

  P("");
  P("=== 4. 集合枚举（ShowHidden=true 后） ===");
  var c2 = 0;
  try { c2 = Number(doc.Bookmarks.Count); } catch (eC) {}
  var lim = c2 > 20 ? 20 : c2;
  for (var j = 1; j <= lim; j++) {
    try {
      var b = doc.Bookmarks.Item(j);
      P("#" + j + " " + String(b.Name) + " " + Number(b.Start) + ".." + Number(b.End));
    } catch (eH) { P("#" + j + " ERR " + err(eH)); }
  }

  P("");
  P("=== 5. StoryRanges ===");
  T("StoryRanges.Count", function () { return doc.StoryRanges.Count; });
  T("StoryRanges.Item(1).Bookmarks.Count", function () {
    return doc.StoryRanges.Item(1).Bookmarks.Count;
  });

  alert("隐藏书签探针完成（" + lines.length + " 行）\r\r" + lines.join("\r"));
}
