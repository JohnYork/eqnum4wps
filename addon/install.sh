#!/usr/bin/env bash
# 把 addon/ 安装到 WPS 加载项目录，并维护 publish.xml
set -euo pipefail

NAME="eqnum"
VERSION="0.2.9"
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
JSADDONS="${WPS_JSADDONS_DIR:-$HOME/.local/share/Kingsoft/wps/jsaddons}"
TARGET="$JSADDONS/${NAME}_${VERSION}"

mkdir -p "$JSADDONS" "$TARGET"

# 删除本加载项的其他版本目录，避免 WPS 加载到旧文件
find "$JSADDONS" -maxdepth 1 -type d -name "${NAME}_*" ! -name "${NAME}_${VERSION}" -print -exec rm -rf {} + 2>/dev/null || true

cp "$ROOT/index.html" "$TARGET/"
cp "$ROOT/ribbon.xml" "$TARGET/"
cp "$ROOT/main.js" "$TARGET/"

rm -rf "$TARGET/js" "$TARGET/ui" "$TARGET/images"
mkdir -p "$TARGET/js" "$TARGET/ui" "$TARGET/images"
cp "$ROOT/js/"*.js "$TARGET/js/"
cp "$ROOT/ui/"* "$TARGET/ui/"
cp "$ROOT/images/"* "$TARGET/images/"

# 核心逻辑直接复用宏版脚本，避免两份实现分叉
cp "$ROOT/../src/wps-eqnum.js" "$TARGET/js/wps-eqnum.js"

python3 - "$JSADDONS/publish.xml" "$NAME" "$VERSION" <<'PY'
import os
import sys
import xml.etree.ElementTree as ET

path, name, version = sys.argv[1], sys.argv[2], sys.argv[3]

root = None
if os.path.exists(path) and os.path.getsize(path) > 0:
    try:
        root = ET.parse(path).getroot()
    except Exception:
        root = None

if root is None or root.tag != "jsplugins":
    root = ET.Element("jsplugins")

for el in list(root):
    if el.get("name") == name:
        root.remove(el)

ET.SubElement(root, "jsplugin", {
    "name": name,
    "type": "wps",
    "url": "%s_%s" % (name, version),
    "version": version,
    "enable": "enable_dev",
    "install": "null",
    "customDomain": "",
})

try:
    ET.indent(root, space="  ")
except AttributeError:
    pass

ET.ElementTree(root).write(path, encoding="UTF-8", xml_declaration=True)
print("[install] publish.xml 已更新: %s" % path)
PY

echo "[install] 已安装到 $TARGET"
echo "[install] 重启 WPS 后，功能区「公式编号」按钮会打开右侧任务窗格"
