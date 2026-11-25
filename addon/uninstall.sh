#!/usr/bin/env bash
# 移除已安装的 EQNUM 加载项，并从 publish.xml 中删除条目
set -euo pipefail

NAME="eqnum"
JSADDONS="${WPS_JSADDONS_DIR:-$HOME/.local/share/Kingsoft/wps/jsaddons}"

find "$JSADDONS" -maxdepth 1 -type d -name "${NAME}_*" -print -exec rm -rf {} + 2>/dev/null || true

if [ -f "$JSADDONS/publish.xml" ]; then
  python3 - "$JSADDONS/publish.xml" "$NAME" <<'PY'
import sys
import xml.etree.ElementTree as ET

path, name = sys.argv[1], sys.argv[2]
try:
    root = ET.parse(path).getroot()
except Exception:
    sys.exit(0)

removed = 0
for el in list(root):
    if el.get("name") == name:
        root.remove(el)
        removed += 1

if removed:
    try:
        ET.indent(root, space="  ")
    except AttributeError:
        pass
    ET.ElementTree(root).write(path, encoding="UTF-8", xml_declaration=True)
    print("[uninstall] 已移除 publish.xml 中的 %d 个条目" % removed)
PY
fi

echo "[uninstall] 完成，重启 WPS 生效"
