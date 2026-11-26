eqnum 公式编号（WPS Linux 加载项）v__VERSION__
================================================

在 WPS 文字中为公式自动编号，支持按章编号 (2.1) 与连续编号 (1)，
重编号后自动修复正文里的原生交叉引用。

一、环境要求
------------
- WPS 文字 Linux 版（12.x），或兼容的 WPS Office for Linux
- bash、python3（用于更新 publish.xml）
- 以当前登录用户身份安装，不要用 sudo

二、安装
--------
1. 解压后进入本目录；
2. 运行：
       bash install.sh
   （已赋予执行权限时也可用 ./install.sh）
3. 脚本会自动定位 WPS 的 jsaddons 目录并安装，最后打印安装路径；
4. 完全退出并重启 WPS；
5. 打开任意文档，在「插入」标签页找到「公式编号」分组。

安装位置默认是：
    ~/.local/share/Kingsoft/wps/jsaddons/eqnum___VERSION__

三、自动识别规则
----------------
脚本按以下顺序确定 jsaddons 目录：
1. 环境变量 WPS_JSADDONS_DIR（手动指定，优先级最高）
2. 已知路径中已存在 publish.xml 的目录
3. 在 ~/.local/share、~/.config、~/.kingsoft 下搜索名为 jsaddons 的目录
4. 以上都没有时，若检测到 WPS 安装（/usr/lib/office6、/opt/kingsoft/wps-office
   或 PATH 中的 wps），则创建 ~/.local/share/Kingsoft/wps/jsaddons

指定自定义目录的示例：
    WPS_JSADDONS_DIR="$HOME/.local/share/Kingsoft/wps/jsaddons" bash install.sh

四、卸载
--------
    bash uninstall.sh

脚本会删除 eqnum_* 目录并从 publish.xml 中移除对应条目，重启 WPS 生效。

五、常见问题
------------
1. 重启后功能区没有「公式编号」分组
   删除 ~/.local/share/Kingsoft/wps/jsaddons/authaddin.json 后再次重启 WPS。

2. 分组出现但按钮没有图标
   按 ALT+F12 打开调试器，查看 Console 是否有
   [EQNUM] GetImage ... 输出；如有异常请联系开发者。

3. 找不到加载项目录
   用 WPS_JSADDONS_DIR 显式指定，例如：
   WPS_JSADDONS_DIR="$HOME/.local/share/Kingsoft/wps/jsaddons" bash install.sh

六、目录结构
------------
    install.sh     安装脚本（自动识别路径）
    uninstall.sh   卸载脚本
    README.txt     本文件
    addon/         加载项文件
      index.html   入口页
      main.js      功能区回调
      ribbon.xml   功能区定义
      js/          核心逻辑与面板脚本
      ui/          任务窗格页面
      images/      按钮图标

注意：本工具会修改当前文档中的公式与编号，正式使用前请先备份文档。
