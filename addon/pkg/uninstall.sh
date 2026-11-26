#!/usr/bin/env bash
# eqnum 卸载脚本（发布包内使用）
set -euo pipefail

NAME="eqnum"
log() { printf '[eqnum] %s\n' "$*"; }
die() { printf '[eqnum] 错误：%s\n' "$*" >&2; exit 1; }

detect_jsaddons() {
  local data="${XDG_DATA_HOME:-$HOME/.local/share}"
  local c found

  if [ -n "${WPS_JSADDONS_DIR:-}" ] && [ -d "${WPS_JSADDONS_DIR}" ]; then
    printf '%s\n' "${WPS_JSADDONS_DIR}"; return 0
  fi

  local cands=(
    "$data/Kingsoft/wps/jsaddons"
    "$data/kingsoft/wps/jsaddons"
    "$data/Kingsoft/WPS/jsaddons"
    "$HOME/.kingsoft/wps/jsaddons"
    "$HOME/.config/Kingsoft/wps/jsaddons"
  )
  for c in "${cands[@]}"; do
    [ -f "$c/publish.xml" ] && { printf '%s\n' "$c"; return 0; }
  done

  found="$(find "$data" "$HOME/.config" "$HOME/.kingsoft" -maxdepth 5 -type d -name jsaddons 2>/dev/null | head -n 1 || true)"
  [ -n "$found" ] && { printf '%s\n' "$found"; return 0; }
  return 1
}

JSADDONS="$(detect_jsaddons)" || die "未找到 WPS 加载项目录，可用 WPS_JSADDONS_DIR 指定。"

removed=0
while IFS= read -r -d '' d; do
  log "删除 $d"
  rm -rf "$d"
  removed=1
done < <(find "$JSADDONS" -maxdepth 1 -type d -name "${NAME}_*" -print0 2>/dev/null)

if [ "$removed" -eq 0 ]; then
  log "没有找到已安装的 $NAME"
fi

if [ -f "$JSADDONS/publish.xml" ] && command -v python3 >/dev/null 2>&1; then
  cp "$JSADDONS/publish.xml" "$JSADDONS/publish.xml.bak"
  python3 - "$JSADDONS/publish.xml" "$NAME" <<'PY'
import sys
import xml.etree.ElementTree as ET

path, name = sys.argv[1], sys.argv[2]
try:
    root = ET.parse(path).getroot()
except Exception:
    sys.exit(0)

for el in list(root):
    if el.get("name") == name:
        root.remove(el)

try:
    ET.indent(root, space="  ")
except AttributeError:
    pass
ET.ElementTree(root).write(path, encoding="UTF-8", xml_declaration=True)
print("[eqnum] publish.xml 已移除条目")
PY
fi

log "卸载完成，重启 WPS 生效。"
