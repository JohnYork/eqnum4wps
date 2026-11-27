# eqnum 工程状态（2026-11-26）

## 1. 项目目标

在 WPS 文字（Linux）中为公式自动编号，当前主形态是 JS 加载项（任务窗格 + 功能区），`src/wps-eqnum.js` 同时保留宏入口：

- 任务窗格提供「编号当前公式」「本章重编号」「全文重编号」「插入引用」「刷新引用」，设置 / 文档统计 / 日志均为可收起抽屉
- 功能区并入内置「插入」标签的「公式编号」分组
- 支持全文连续编号 `(1)` 与按章编号 `(2.1)`，章层级可配置
- 按章编号必须与章节号联动：`STYLEREF` + `SEQ 公式 \s`
- 显示公式用无边框两栏表格排版：左列公式居中，右列编号右对齐，表格对齐版心
- 引用使用原生题注引用或本工具书签 REF，重编号后自动刷新（重建同名书签）

## 2. 环境与探测结论

**运行环境**
- WPS 文字 12.0 / 12.1.2.29183，Linux，`/usr/lib/office6`
- 宏编辑器为 JSA：模块顶层只能放函数声明，不能有变量初始化或函数调用
- 所有脚本都必须把常量和辅助函数放在入口函数内部

**已验证可用**
- `Document.OMaths`、`OMath.Type`（0=显示，1=行内）
- `InlineShapes` + `OLEFormat.ProgID`
- `Fields.Add(..., -1, " SEQ/STYLEREF/REF ... ", ...)`、`Field.Update()`、`Fields.Update()`
- `Bookmarks` 创建 / `Exists` / 删除
- `CustomDocumentProperties` 读写
- `Range.ConvertToTable("\t")` 保留 OMML 公式
- 表格：边框关闭、宽度百分比、`Cell.Width`、垂直居中、单元格合并、内边距 0
- `ListFormat.ListString` 读取标题自动编号
- `InsertCaption("公式", ...)` 可用，会生成 `STYLEREF 1 \s` + `SEQ 公式 \s 1`
- 原生交叉引用（类型=1 整项题注）会生成 `REF _Ref...`，**结果带外层括号**
- 表格单元格内的题注字段可被原生交叉引用识别
- `STYLEREF 2 \n` / `STYLEREF 2 \w` / `STYLEREF "标题 2" \n` 都能返回 `1.1`、`1.2`
- `SEQ 公式 \s 2` 按标题 2 重置；`SEQ EQNUM` 与 `SEQ 公式` 是独立计数器
- `CommandBars` 可定位「插入 → 引用 → 交叉引用」（Id=775），`Execute()` 能打开原生对话框

**已验证不可用 / 有限制**
- `Range.Areas` 不存在：Writer 不支持不连续多选
- 单元格 / 列的 `PreferredWidthType = 1`（自动宽度）不生效
- `AutoFitBehavior(1)`（按内容）会让表格总宽缩成内容宽
- 原生交叉引用对话框无法通过 API 预选「公式」，只能打开后手动选一次
- `Dialogs.Item(41).Show()` 报错；`ExecuteMso("InsertCrossReference")` 是空操作
- `CommandBars.FindControl(0, 775)` 返回空，必须扫描菜单栏定位

## 3. 目标文档与扫描数据

- `可配置型卫星定位接收机软件架构设计方案.docm`
- 7556 段落、8217 OMaths（显示 548 / 行内 7669）、71 InlineShapes、7363 域、16 书签、31 表格
- 标题 1 共 9 个，标题 2 共 52 个；`ListString` 可读

**显示公式分类（早期只统计 Type=0）**

| 分类 | 数量 |
|---|---|
| 可编号（独占段落、非表格） | 100 |
| 正文混排 | 231 |
| 表格内 | 74 |
| 同段多公式 | 143 |
| 读取失败 / 其他对象 | 0 |

**全文重编号候选**
- 候选 **988 个**，全部为 `Type=1` 行内公式、独占段落、编号在下一段
- 已确认这些是合法公式编号，不是 `(1.14)中得到：` 这类正文引用
- 例：段 #252 公式段只有公式，段 #253 只有 `(1.1)`
- 已修复误判：编号后面必须是纯非打印字符；跨段下一段若含公式/对象，只接受“编号 + 非打印字符 + 对象”
- StyleSeparator 允许字符集合：`\v \n \r \u001E \u001F \u2028 \u2029`（不含分页符 `\f`）

**现有编号**
- 显示公式段落中旧编号 103 处，混有正文引用（如“式(2.85)”“由(1.9)可知”）
- MathType 公式 7 个（显示 3、行内 4、表格内 3）
- MathType 域：`ZEqnNum` 66、`SEQ MTEqn/MTSec` 12、`MTPlaceRef` 3，其中 38 个结果显示“错误”

## 4. 已确认的需求决策

- **编号当前**：不限制公式 Type；混排时先确认再拆成独立段落；只编当前公式，不触发全文重排
- **全文重编号**：只处理“公式后面已有编号”的公式；允许空格、制表符、软/硬换行、StyleSeparator，不跨分页/分节符；编号后若还有正文，不算公式编号
- **多公式段**：暂不处理（2C）
- **旧编号清理**：只清公式尾部编号，正文引用不动（3A）；跨段旧号也识别并删除
- **MathType**：保留原有编号，不纳入新编号体系
- **编号外观**：半角括号；连续 `(1)`；按章 `(2.1)`；层级可配置
- **章节联动**：按章始终使用 `STYLEREF <层级> \n` + `SEQ 公式 \s <层级>`，不使用“普通文本章号 + 显式 `SEQ \r`”方案
- **版式**：无边框两栏表格，表宽 = 版心宽 100%，右列宽按最长编号计算，左列 = 版心宽 − 右列宽
- **引用**：原生题注引用（整项题注，带括号）或本工具 `EQNUM_REF_` 书签 REF；前缀默认“式”
- **设置存储**：文档自定义属性
- **刷新引用**：`全部 / 仅本工具 / 不刷新`，默认“仅本工具”
- **子公式编号**：设置项保留，暂不生效

## 5. 宽度方案（已实测）

- 版心宽 = `PageSetup.PageWidth - LeftMargin - RightMargin`，目标文档约 415.3pt / 14.65cm
- 表格 `PreferredWidthType = 2`、`PreferredWidth = 100`
- 右列 `Cell(1,2).Width = 编号宽度`
- 左列 `Cell(1,1).Width = 版心宽 - 编号宽度`
- 编号宽度公式：`字符数 × 字号 × 0.55 + 8pt`，限制 32–64pt
  - 五号字下 `(2.1)` ≈ 36pt / 1.27cm
  - `(10.123)` ≈ 48pt / 1.69cm
- 之后再设回表格 100% 不会重置列宽

## 6. 文件清单

| 路径 | 用途 |
|---|---|
| `src/wps-eqnum.js` | 主功能，全部入口宏 |
| `src/wps-eqnum-scan.js` | 早期只读扫描与编号预览 |
| `probe/wps-eqnum-probe.js` | 第一轮 API 能力探测 |
| `probe/wps-eqnum-probe2.js` | 第二轮 OMaths / 现有编号探测 |
| `probe/wps-eqnum-proto1.js` | 单公式排版链路原型（20 项 PASS） |
| `probe/wps-eqnum-width-test*.js` | 编号列宽度试验 1–3（3 为最终方案） |
| `probe/wps-eqnum-styleref-test.js` | STYLEREF 兼容性探测 |
| `probe/wps-eqnum-caption-test*.js` | 题注、交叉引用、表格内题注探测 |
| `probe/wps-eqnum-caption-paren-test.js` | 原生题注引用是否带括号 |
| `probe/wps-eqnum-caption-level2*.js` | 二级章 `STYLEREF` / `SEQ \s` 探测 |
| `probe/wps-eqnum-xref-dialog-test.js` | 原生交叉引用对话框 API 探测 |
| `probe/wps-eqnum-commandbar-test.js` | `CommandBars.ExecuteMso` 探测 |
| `probe/wps-eqnum-commandbar-scan.js` | 扫描菜单栏，定位交叉引用 Id=775 |
| `可配置型卫星定位接收机软件架构设计方案.docm` | 目标文档；测试一律使用副本 |

## 7. 主脚本实现状态

**入口宏（`src/wps-eqnum.js`）**
- `EQNUM_CURRENT()`：编号光标所在公式，不重排其他公式
- `EQNUM_RENUMBER()`：全文重编号
- `EQNUM_REFRESH()`：
  - `1` = 只刷新当前章节已有编号
  - `2` = 刷新全文已有编号
  - `3` = 扫描并重编号当前章节（只扫描章节段落范围）
- `EQNUM_SETTINGS()`：编号模式、章层级、起始值、引用前缀、刷新引用策略
- `EQNUM_INSERT_REF()`：打开 WPS 原生交叉引用对话框（CommandBars Id=775），失败回退输入框 + 书签 REF
- `EQNUM_UPDATE_REFS()`：单独刷新全文引用（`all` 模式，较慢）

**核心机制**
- 公式识别：全文重编号按“公式后面已有编号”筛选；编号当前不限 Type；已生成表格通过右单元格书签 `EQNUM_NUM_xxxx` 识别
- 建表（v30）：建 **1×2 空表**（`Tables.Add(段尾, 1, 2)`）→ 取公式片段的 `Range.WordOpenXML()` → `Cell(1,1).Range.InsertXML(xml)` → 删原段落
  - 不能用 `ConvertToTable("\t")`：制表符会被吸进数学区，右栏出现裸 `m:r`，引用被当成公式对象
  - 不能用 `ConvertToTable(0)` + `Cell.Split`：右栏克隆 `<m:rPr/>`，编号仍落在数学区
  - 不能用 `Cell.Range.FormattedText` 赋值：加载项环境下跨进程死锁，必卡死
  - 不能用剪贴板：可用，但每处理一个公式就覆盖一次用户剪贴板
  - 插入后校验：左栏 `OMaths > 0`，右栏必须为 0，否则记 WARN / FAIL
- 编号：右单元格写入 `( STYLEREF <层级> \n . SEQ 公式 \s <层级> )`；连续模式为 `( SEQ 公式 )`；套用「题注」样式并恢复字号
- 章节联动：按章模式保留 `STYLEREF` + `SEQ \s`，性能优化不改变编号结构
- 书签：`EQNUM_NUM_0001` 起，覆盖整段编号
- 重排：扫描 → **倒序**删除旧号/写题注字段 → 更新书签内全部域 → 按 `CFG.refreshRefs` 刷新引用
  - 倒序理由：前面的表格化不会让后面已捕获的 Range/OMath 失效
  - 2026-11-24 试过正序 + 临时书签 `EQNUM_ANC_` 锚点：更慢（2556s vs 1883s），已回退
  - 结论：换顺序不是性能杠杆，单条开销来自每次编辑触发的整篇重排；正序还需额外维护锚点
  - 下一段定位改为 `nextParagraphAfter`（`Paragraph.Next` → 段尾折叠区域），不再依赖段落序号；扫描阶段仍保留序号兜底
  - `removeOldNumber` 支持传入实时段落对象（`trail.nextPara`），倒序时仍走 `paraIndex` 兜底
- 旧编号处理（v11）：`EQNUM_CURRENT` / `EQNUM_REFRESH` / `EQNUM_RENUMBER` 共用 `removeOldNumber`
  - 旧编号单独成段时整段删除；范围删除后只剩 `1.1)`、`)` 等残片也删整段
  - 同段旧号由建表清理逻辑处理
  - **已彻底移除** `mergeTrailingIntoFormula`：WPS 上跨 OMath 的文本替换会留残片、破坏结构
  - 回归理由：用户要求文档内容干净，接受删除开销（隐藏方案会让旧编号永久留在文件里）
- 旧编号引用（v31）：编号写完后用 `readdRefBookmarks` 在**新编号范围上重建同名 `_Ref…` 书签**，正文原生 REF 域一个字符都不改，F9 即显示新编号
  - 重建失败的书签才用 `repointRefFields` 改写域代码指向新书签；域数 > 4000 时跳过兜底，避免 O(n²) 卡死
  - **不再**询问“更新 / 忽略”，也**不再**替换正文里手打的纯文本引用（如 `式(1.1)`）；面板上的单选组与“执行前不再弹确认框”已移除
  - 纯文本引用不复写是刻意取舍：无法区分“手动引用”和“公式编号文本”，替换会误伤正文
- 排版：表级内边距、表级段落格式、`Cells.VerticalAlignment` 批量垂直居中（失败回退逐单元格）
- 引用书签：`EQNUM_INSERT_REF` 插入 REF 后用 `EQNUM_REF_xxxx` 包住域
- 引用刷新默认 `ours`：只遍历 `EQNUM_REF_` 书签，不扫描 7486 个域
- 章节扫描：`chapterBounds` 定位当前章节起止段 → `collectFormulas` 只扫该范围

## 8. 性能实测与预估

**第 1 章 123 个公式实测：倒序 vs 正序**

| 版本 | 循环总用时 | 单条耗时 |
|---|---|---|
| 倒序（v2） | 1883 秒 | 2.6s → 26s 递增 |
| 正序 + 书签锚点（v3） | 2556 秒 | 7s → 30s |

- 正序的单条耗时几乎全部落在「定位删号」：13–28 秒；转表/排版/写号只有几十毫秒
- 说明真正的瓶颈是**每次编辑触发的整篇重排**，与处理方向无关；换顺序无法解决
- 编号域更新仅 3.3 秒（123 个）、引用 `ours` 24ms，都不是瓶颈
- 引用 `all` 模式 874 个 `_Ref` 曾耗时 914 秒，默认已改为 `ours`

**历史尝试（均已否决，勿重复）**
- `mergeTrailingIntoFormula`（v6）：把旧号段并进公式段再转表，省一次删除；WPS 上跨 OMath 的文本替换会留残片、破坏结构，**已彻底移除**
- `pauseHeavyUpdates` / `resumeHeavyUpdates`（v5）：实测对本项目无收益，代码保留
- 写号 `deferUpdate=true`：收益仅约 1–2 秒，风险不值得，已放弃
- 隐藏旧编号（v9）：旧编号会以隐藏文字永久留在文件里，且按书签统一删除依旧过慢，已回退（见第 8 节末）
- 制表位替代表格：删号耗时不变（表格不是元凶），写号反而从 ~0.1s 涨到 4.2s，已否决

**章节模式第 3 项**
- 只扫描当前章节范围，预计 1–2 分钟，而不是全文 10 分钟
- 章号仍由 `STYLEREF` 解析，不受扫描范围影响

## 9. 已知限制与风险

- 全文重编号只处理“后面已有编号”的公式；没有编号的公式不会自动编号；同段多公式仍跳过；非本工具表格内公式不处理
- 题注章号依赖 `STYLEREF`：公式位于第一个标题之前时可能显示“未找到关联章节样式”
- 旧 `SEQ EQNUM` 域不参与 `SEQ 公式` 计数；切换后需一次全文重编号
- 文档中其他手动「公式」题注会共用 `SEQ 公式` 计数器
- 连续模式起始值只作用于全文重编号的第一个公式
- 正文已有 `式(2.85)` 是普通文本，按设计**不会**被改写；重编号后可能与新编号不一致，需人工核对
- MathType 公式保留旧编号；38 个损坏域不处理
- `Cell.Range.Text = ""` 清空编号单元格已验证；建表已改为 `Tables.Add` + `InsertXML(WordOpenXML())`，不再依赖 `ConvertToTable`
- 原生交叉引用对话框无法预选「公式」，需手动选一次
- 刷新引用默认只刷 `EQNUM_REF_` 书签；原生 `_Ref` 需 F9 / 打印，或建表后重建的同名书签
- 旧版本插入的引用没有 `EQNUM_REF_` 书签，`ours` 模式会提示未找到
- 若中途失败，文档可能处于部分修改状态；必须使用副本并先备份

## 10. 回归测试顺序

1. 关闭 WPS，复制目标 `.docm` 作为测试副本（目录里有锁文件 `.~...docm`）
2. 加载项已部署到 `jsaddons/eqnum_0.2.9`；只需在「插入」标签点「公式编号」分组打开面板
3. 在设置抽屉里确认：编号模式、章层级、刷新引用=仅本工具
4. 光标放到第 1 章，点「本章重编号」
   - 日志开头版本应为 `2026-11-26-v32`
   - 检查表格右列编号、公式在左列、表宽等于版心
   - 旧编号与新编号不同的公式，正文原生引用按 F9 后应显示新编号，且光标停在章号上不再弹公式编辑器
5. 检查域代码：`STYLEREF <层级> \n` + `SEQ 公式 \s <层级>`
6. 点「插入引用」，确认原生交叉引用对话框能打开，类型选「公式」，整项题注显示 `(2.1)`
7. 单章稳定后再跑「全文重编号」；监控 `旧号/转表/排版/写号/编号域/引用域` 日志
8. 宏版回归：把 `src/wps-eqnum.js` 粘贴到宏编辑器，运行同一批 `EQNUM_*` 入口
9. 性能仍受限于「每次编辑触发整篇重排」，暂不再做单项优化；如需提速只能减少编辑次数

## 11. 运行方式与安全约定

- 宏编辑器 → 新建脚本 → 粘贴完整文件 → 运行对应 `EQNUM_*` 入口
- 可在 WPS 自定义工具栏/快速访问栏绑定这些宏
- 主脚本会修改当前文档，执行前必须备份，优先使用副本
- 探测/宽度/题注试验只在新文档或只读模式下运行
- 目标 `.docm` 是否已被修改需在下次开始时确认；不要直接在唯一正式文档上跑全文重编号

## 12. 加载项版本（2026-11-24 起）

**决定**：宏版性能优化暂停（倒序/正序、自动分页、合并旧号均无实质收益），改为加载项版，好处是不随文档保存、UI 可自定义。

**环境确认**
- WPS 12.1.2.29183，`/usr/lib/office6`，自带 `addons/jsapi`、`kcef`、`kwpsjsrunner`
- 加载项目录：`~/.local/share/Kingsoft/wps/jsaddons`（已存在）
- 官方约定：WPS 启动后读取 `jsaddons/publish.xml`，再加载对应组件目录下的 `ribbon.xml` + `main.js`
- **必须自带 `index.html`**（官方模板也带），内容只需一行引入 `./main.js`；文档里“WPS 自动创建 index.html”的说法不可依赖
- **对象模型入口是 `window.Application`**，不是 `wps.WpsApplication()`；直接写 `typeof wps.X` 会抛 ReferenceError，导致回调静默失败（v0.1.0 的故障原因）
- 常用 API（来自官方模板 `@wps-jsapi/wps`）：
  - `Application.ShowDialog(url, title, w, h, modal)`
  - `Application.CreateTaskPane(url)` / `GetTaskPane(id)` / `PluginStorage` / `ribbonUI.InvalidateControl(id)`
  - `GetUrlPath()` 用 `document.location` 推算加载项根路径
  - `alert()` 可用
- 离线安装格式（源自 wpsjs 2.2.3 `src/lib/build.js`）：
  ```xml
  <jsplugins>
    <jsplugin name="eqnum" type="wps" url="eqnum_0.1.1" version="0.1.1"
              enable="enable_dev" install="null" customDomain=""/>
  </jsplugins>
  ```
  目录为 `jsaddons/eqnum_0.1.1/`，`url` 就是该相对目录名
- WPS 首次加载后会在 jsaddons 下生成 `authaddin.json`（记录 name/path/md5/enable/isload）与 `jsaddinblockhost.ini`；未签名加载项在 `enable_dev` 下可以加载
- 调试：加载项页面按 `ALT + F12` 打开调试器，`console.log` 输出在 Console

**已创建文件**
| 路径 | 用途 |
|---|---|
| `addon/index.html` | 加载项入口，唯一职责是引入 `./main.js` |
| `addon/ribbon.xml` | 「插入」标签 → 「公式编号」分组：编号当前公式(large) / 本章重编号 / 全文重编号 / 插入引用(large) / 刷新引用 / 编号设置 |
| `addon/main.js` | 功能区回调：打开任务窗格 + 通过 `PluginStorage` 投递动作 |
| `addon/ui/panel.html` | 任务窗格：5 个等宽竖排按钮 + 进度条/中止/界面响应；设置、文档统计、日志为可收起抽屉 |
| `addon/images/eqnum-current.png` | 功能区「编号当前公式」大按钮图标（32×32，x + 蓝色 (1)） |
| `addon/images/eqnum-insert-ref.png` | 功能区「插入引用」大按钮图标（32×32，文档 + 箭头 + 橙色 (1)） |
| `addon/js/shim.js` | 适配层：`Application` 别名、`confirm/alert/InputBox` 转到面板日志、`console.log` 汇总到 `EQNUM_BUS` |
| `addon/js/panel.js` | 面板逻辑：读写 `CustomDocumentProperties`、调度动作、日志渲染、轮询动作队列 |
| `addon/install.sh` | 复制到 jsaddons 并维护 publish.xml；**把 `src/wps-eqnum.js` 复制为 `js/wps-eqnum.js`**（单一实现来源） |
| `addon/package.sh` | 构建发布包：`dist/eqnum-<版本>-build<N>-linux.{tar.gz,zip}`（版本/build 取自 main.js） |
| `addon/pkg/` | 发布包模板：`install.sh` / `uninstall.sh`（自动识别 jsaddons 路径）/ `README.txt` |
| `dist/eqnum-0.2.9-build1-linux.tar.gz`、`.zip` | 可直接分发的安装包，解压后 `bash install.sh` |
| `addon/uninstall.sh` | 删除目录并移除 publish.xml 条目 |

**架构：核心逻辑复用宏脚本**
- 面板页按 `shim.js → wps-eqnum.js → panel.js` 顺序加载，直接调用全局入口 `EQNUM_CURRENT/EQNUM_RENUMBER/EQNUM_REFRESH/EQNUM_INSERT_REF/EQNUM_UPDATE_REFS`
- `addon/js/wps-eqnum.js` 只是运行时副本：`install.sh` 从 `src` 复制到安装目录，`package.sh` 打包前也用 `src` 覆盖，**禁止以它为改动来源**（build 1 曾因打包用了旧副本导致修复未生效）
- `wps-eqnum.js` 里所有辅助函数与 `CFG/state` 都在 `EQNUM_MAIN()` 内部，加载时不会执行任何逻辑
- `main()` 增加 `if (typeof mode === "undefined" || !mode) return;` 作为保险
- 面板自己读写文档属性（键名与宏版一致：`EQNUM_MODE/LEVEL/START/REF_PREFIX/REFRESH_REFS`），不调用 `EQNUM_SETTINGS`
- 进度/中止：`renumberAll` 为 `async`，每条公式后 `await eqnumUiGap()` 让出线程，面板可实时刷新日志与进度条

**阶段 0 验证结果（已通过）**
- 功能区出现「公式编号」选项卡
- 环境自检输出：WPS 文字 12.0（Build 12.1.2.29183）、7556 段、8217 公式、7363 域、16 书签、31 表格、自定义属性 5、视图类型 3
- `CreateTaskPane`、`ShowDialog` 均为 function
- **`Options.Pagination` 本来就是 false**，宏版“关闭自动分页”优化从一开始就是空操作

**阶段 1（0.2.1）**
1. 功能区按钮 → 打开右侧任务窗格（已可用）
2. 面板显示文档名、设置、统计、日志；新增「本工具编号」计数
3. 「刷新范围」默认选「扫描并重编号当前章节」，并根据计数给出黄色提示
4. 语义澄清：**「刷新」只处理本工具已生成的编号表**（`EQNUM_NUM_` 书签），
   原始文档首次必须用「扫描并重编号当前章节」或「全文重编号」，否则提示“没有本工具生成的编号公式”

**阶段 2（0.2.2）**
- `renumberAll` 改为 `async`，每条公式处理完 `await yieldToUi()`（内部 `setTimeout(0)`），让面板能实时重绘
- 中止：面板「中止」按钮置 `window.EQNUM_CANCEL = true`，循环在下一个让出点检查并 `break`，
  `renumberAll` 返回 `{cancelled:true}`，仍会执行编号域与引用更新，保证已处理部分一致
- 调用链：`EQNUM_* → EQNUM_MAIN → main() → runRenumber/runRefresh → await renumberAll`
- 面板按返回值是否 thenable 决定何时恢复按钮；动作标签带上范围（扫描并重编号 / 刷新当前章节 / 刷新全文）
- 兼容：宏环境没有 `setTimeout` 时 `yieldToUi` 立即 resolve，仍为同步执行
- 风险：JSA 若不支持 `async/await` 语法，宏编辑器会解析失败；加载项（Chromium）无此问题

**后续阶段**
1. 阶段 3：插入引用改为面板内选择编号列表，不再依赖 `InputBox` / 原生交叉引用对话框
2. 阶段 4：性能与体验收尾（批量排版、操作可取消），必要时兼容 `wpsjs build/publish`

**阶段 2b（0.2.3）：文档界面重绘**
- 原因：全程 `ScreenUpdating=false`，且 `setTimeout(0)` 的间隙太短，WPS 主线程来不及重绘，所以文档区看起来“卡死”
- 做法：新增 `eqnumUiGap()`，每条公式后临时把 `ScreenUpdating` 置 true，停留 `window.EQNUM_UI_GAP` 毫秒再关掉
- 面板新增「界面响应」下拉（0 / 50 / 100 / 200 ms，默认 50），存 `localStorage`
- 预期效果：从“整篇全黑卡死”变成“每条公式卡一下、中间能看到文档刷新”
- 局限：单条公式内部的编辑（旧号删除 5–27 秒）仍会占用 WPS 主线程，无法通过让出线程消除；总时长增加约 123×间隔

**阶段 2c（0.2.4）：间隙强制重绘**
- `eqnumUiGap()` 打开 `ScreenUpdating` 后，额外尝试 `Application.ScreenRefresh()` 与
  `Application.ActiveWindow.ScreenRefresh()`，不存在时静默跳过
- 配合面板「界面响应」100/200ms 观察是否真的能看到文档刷新

**阶段 2d（0.2.6）：进度条 + 中止**
- 核心新增 `reportProgress(done, total, text)`，通过 `window.EQNUM_PROGRESS` 回调给面板
  - `renumberAll` 每条公式前上报 `序号-1 / 总数`，结尾上报「更新编号域」「刷新引用」「完成」
  - 全文重编号 / 章节扫描开始前上报阶段文字（该阶段是同步调用，界面可能短暂无响应）
- 面板新增进度文字 + 进度条（`progText` / `progBar`），100% 变绿；「中止」立即改进度文字
- 每公式后的界面响应间隔默认 **50ms**，实现位置仍是 `eqnumUiGap()`（临时开 `ScreenUpdating`）
- 中止语义：置 `window.EQNUM_CANCEL`，当前公式处理完后跳出循环，**仍会**更新已编号公式的域

**制表位原型（probe/wps-eqnum-tabstop-test.js）**
- 目的：验证“不要表格、改右对齐制表位”后，**删旧号那一步是否变便宜**（表格版中转表+排版+写号仅约 284ms，旧号却要 5–27s）
- 做法：从当前章节末尾向前取 N 个“公式段 + 独立编号段”候选，依次
  删旧号段 → 设右对齐制表位（版心宽，`TabStops.Add(pos, 2, 0)`）→ 公式后插 `\t` → 写 `( STYLEREF 1 \n . SEQ 公式 \s 1 )`
- 输出：每条的分段耗时与平均值；`TabStops` 不支持时会标“制表位失败”
- 注意：探针不建书签、不处理引用，仅用于计时对比；必须用副本

**制表位原型结论（2026-11-24，已否决）**
- 8 个样本：总 561 秒；平均 删号 65872ms、制表位 2ms、插制表符 7ms、写号 4185ms
- 删号依旧 52–142 秒 → **表格不是元凶**，去掉表格不会变快
- 写号从表格版的 ~0.1s 涨到 4.2s（字段直接插正文，每次 Update 触发重排）
- 注意：探针未设 `ScreenUpdating=false`，与正式流程不可直接比，但“表格无关”这一结论成立
- 下一步：用 `probe/wps-eqnum-delete-cost-test.js` 拆解删号开销（按序号取段落 / 读文本 / 删整段 vs 删跨段范围）

**删号拆解结论（2026-11-24）**
- 平均：取段落 62ms、读文本 4ms、删整段 51714ms、删跨段范围 55488ms
- 结论：**段落删除本身**（约 50 秒/次）才是瓶颈，与按序号取段落、删除方式无关
- 下一步候选：不删段落，改为“删编号文字 + 空段设为隐藏文字”，见 `probe/wps-eqnum-hidepara-test.js`

**隐藏空段原型结论（2026-11-24）**
- 探针 6 个样本：删编号文字 52304ms、**设隐藏 3ms**、制表位+插入 16ms、写号 3563ms，`Font.Hidden` 生效 6/6
- 用户目视确认：隐藏旧编号段后**没有多余空行**，编号右对齐正常
- v9 曾按此实现（整段含段落标记设 `Font.Hidden = true`，`ensureHiddenTextOff()` 关闭隐藏文字显示）
- 代价：旧编号文本以隐藏形式永久留在文档里
- 待验证（`probe/wps-eqnum-hidepara-test.js` 的 `EQNUM_HIDEBATCH_TEST`）：
  先隐藏 + 加 `HPROBE_DEL_` 书签，编号结束后按书签统一删除。关键是**删除已隐藏段落**
  是否仍要 ~50 秒/条；同探针还会试 `Find/Replace`（格式=隐藏）能否一次批量删除

**隐藏方案最终结论（2026-11-24，已放弃）**
- 用户实测：按书签统一删除仍然过慢，说明删除开销与“是否隐藏”无关
- 用户决策：为保持文档内容干净，接受删除开销 → 主脚本回退到隐藏方案之前的删除逻辑（v10）
- `EQNUM_HIDEBATCH_TEST` 探针保留，供以后复查 WPS 行为

**注意**
- 加载项与宏共用同一套对象模型，性能瓶颈（每次编辑整篇重排）不会因改加载项而消失
- 宏版仍是可用版本，加载项未稳定前不要删 `src/wps-eqnum.js`

**阶段 5（0.2.9 build 1，当前状态）**
- 面板：5 个等宽竖排按钮（编号当前公式 / 本章重编号 / 全文重编号 / 插入引用 / 刷新引用）；设置、文档统计、日志均为可收起抽屉；进度条 + 中止 + 界面响应间隔
- 功能区：并入内置「插入」标签的「公式编号」分组。WPS 不支持挂接「公式助手」等上下文标签，A–E 五个 idMso 候选全部落到自定义标签，该方案已作废
- 建表（v30）：1×2 空表 + `Cell(1,1).Range.InsertXML(公式.WordOpenXML())`；右栏无数学属性，引用点击不再弹公式编辑器
- 旧编号引用（v31）：只重建同名 `_Ref` 书签；删除面板「更新/忽略」单选与「执行前不再弹确认框」，不再替换纯文本引用
- 调试代码（v32）：移除分阶段停点（`stageStop` / `EQNUM_STAGE`）、`EQNUM_DEBUG_SCAN` 入口与面板调试输入框；`probe/` 目录探针保留备查
- 多文档与设置（v33）：`ensureDocument` 改为跟随活动窗口文档（不再按 OMaths 数量激活其它文档）；`Selection` 优先取 `ActiveWindow.Selection`；面板执行前把界面设置写入 `window.EQNUM_SETTINGS_OVERRIDE`，核心以它覆盖文档属性，修复“选连续仍按章编号”
- 混排拆分（v34）：行内公式先转显示公式；拆分后按“原对象 → 位置 → 附近 → 全文最近”逐级找回公式；清理零宽字符/对象占位符，避免把成功拆分误判为混排
- 版本约定：功能未稳定前版本号保持 0.2.9，只递增 `EQNUM_ADDON_BUILD`（面板标题、安装日志、安装包名均带 build）
- 拆分校验（v36）：拆段前记录 OMath 的 `WordOpenXML` 与文本；拆段后先在同段±2 段内按 XML/文本核对，找不到才全文搜索（>4000 公式跳过）。搬运公式直接使用拆段前记录的 XML，避免拆碎后只搬到碎片；若碎片段落已无正文则自动清理
- 选中公式与拆段（v41）：参照公式在拆段前记录（文本 + oMath 本体 XML）；**正文内不做行内→显示转换**（WPS 转换会插入软回车，导致公式看似独立成行却仍在同段），改为在“公式前正文末字符之后 / 公式后正文首字符之前”插入段落标记，失败才回退 `OMath.InsertParagraphBefore/After`，再把残留软回车替换为 `\r`；显示转换推迟到公式进入表格左栏之后执行。拆段后继续用同一 OMath 对象读 `Range` 定位，按本体 XML/文本核对，搜索仅作兜底
- 安装位置：`~/.local/share/Kingsoft/wps/jsaddons/eqnum_0.2.9`，`publish.xml` 条目 `version="0.2.9"`
- 已验证：单公式编号全阶段、本章重编号批量路径、引用 F9 更新、引用点击不再弹公式编辑器

