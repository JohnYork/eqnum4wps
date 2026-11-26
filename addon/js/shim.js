/*
 * 加载项运行环境适配层
 *
 * 作用：让 src/wps-eqnum.js（原本给 JSA 宏用）能直接在加载项页面里运行。
 * - 宏里直接写 Application，这里补一个全局别名指向 window.Application
 * - confirm / alert / InputBox 改为投递到面板日志，避免阻塞式弹窗
 * - console.log 汇总到 EQNUM_BUS，供面板展示
 */

var Application = window.Application;

var EQNUM_BUS = {
  lines: [],
  handlers: [],
  push: function (text, level) {
    var item = { text: String(text), level: level || "log", time: Date.now() };
    this.lines.push(item);
    if (this.lines.length > 4000) this.lines.splice(0, this.lines.length - 4000);
    for (var i = 0; i < this.handlers.length; i++) {
      try { this.handlers[i](item); } catch (e) {}
    }
  },
  on: function (fn) { this.handlers.push(fn); }
};

function eqnumEmit(text, level) { EQNUM_BUS.push(text, level); }

// 面板里用队列回答宏的 InputBox 询问
window.__EQNUM_INPUT_QUEUE = [];

function InputBox(msg, title, def) {
  if (window.__EQNUM_INPUT_QUEUE.length > 0) {
    return window.__EQNUM_INPUT_QUEUE.shift();
  }
  eqnumEmit("[提示] " + (title || "") + "：" + String(msg).replace(/\r/g, " "), "warn");
  return def;
}

// 加载项里由用户点击按钮即视为确认
window.confirm = function (msg) {
  eqnumEmit("[确认] " + String(msg).replace(/\r/g, " "), "warn");
  return true;
};

window.alert = function (msg) {
  eqnumEmit(String(msg), "alert");
};

(function () {
  var orig = null;
  try { orig = console.log ? console.log.bind(console) : null; } catch (e) {}
  console.log = function () {
    var parts = [];
    for (var i = 0; i < arguments.length; i++) parts.push(String(arguments[i]));
    eqnumEmit(parts.join(" "), "log");
    if (orig) {
      try { orig.apply(console, arguments); } catch (e) {}
    }
  };
})();
