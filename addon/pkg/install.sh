#!/usr/bin/env bash
# eqnum 安装脚本（发布包内使用）
# 自动识别 WPS 的 jsaddons 目录，安装后重启 WPS 生效。
set -euo pipefail

NAME="eqnum"
VERSION="__VERSION__"
BUILD="__BUILD__"
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SRC="$HERE/addon"

log()  { printf '[eqnum] %s\n' "$*"; }
warn() { printf '[eqnum] 注意：%s\n' "$*" >&2; }
die()  { printf '[eqnum] 错误：%s\n' "$*" >&2; exit 1; }

[ -f "$SRC/main.js" ] || die "安装包不完整，缺少 $SRC/main.js"

if [ "$(id -u)" -eq 0 ]; then
  warn "正在以 root 身份运行。加载项属于当前登录用户，建议用普通用户执行本脚本。"
fi

# 依次尝试：环境变量 → 已知路径（优先已有 publish.xml）→ 搜索已有 jsaddons → 按 WPS 安装情况兜底
detect_jsaddons() {
  local data="${XDG_DATA_HOME:-$HOME/.local/share}"
  local c found

  if [ -n "${WPS_JSADDONS_DIR:-}" ]; then
    [ -d "${WPS_JSADDONS_DIR}" ] && { printf '%s\n' "${WPS_JSADDONS_DIR}"; return 0; }
    warn "WPS_JSADDONS_DIR 指向的目录不存在：${WPS_JSADDONS_DIR}"
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

  if [ -d /usr/lib/office6 ] || [ -d /opt/kingsoft/wps-office ] || command -v wps >/dev/null 2>&1; then
    printf '%s\n' "$data/Kingsoft/wps/jsaddons"
    return 0
  fi
  return 1
}

JSADDONS="$(detect_jsaddons)" || die "未找到 WPS 加载项目录。可用 WPS_JSADDONS_DIR=/path/to/jsaddons 手动指定。"

[ -d "$JSADDONS" ] || mkdir -p "$JSADDONS"
TARGET="$JSADDONS/${NAME}_${VERSION}"

WPS_BIN="$(command -v wps || command -v wps-office || true)"
[ -n "$WPS_BIN" ] && log "检测到 WPS 可执行文件：$WPS_BIN"
log "加载项目录：$JSADDONS"
log "安装目标：$TARGET"

# 删除本加载项的旧版本目录
find "$JSADDONS" -maxdepth 1 -type d -name "${NAME}_*" ! -name "${NAME}_${VERSION}" -print -exec rm -rf {} + 2>/dev/null || true

rm -rf "$TARGET"
mkdir -p "$TARGET/js" "$TARGET/ui" "$TARGET/images"
cp "$SRC/index.html" "$SRC/main.js" "$SRC/ribbon.xml" "$TARGET/"
cp "$SRC/js/"*.js "$TARGET/js/"
cp "$SRC/ui/"* "$TARGET/ui/"
cp "$SRC/images/"* "$TARGET/images/"

# 维护 publish.xml
command -v python3 >/dev/null 2>&1 || die "需要 python3 来更新 publish.xml"
[ -f "$JSADDONS/publish.xml" ] && cp "$JSADDONS/publish.xml" "$JSADDONS/publish.xml.bak"

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
print("[eqnum] publish.xml 已更新")
PY

for f in index.html main.js ribbon.xml js/wps-eqnum.js js/panel.js js/shim.js ui/panel.html images/eqnum-current.png images/eqnum-insert-ref.png; do
  [ -f "$TARGET/$f" ] || die "安装校验失败，缺少 $TARGET/$f"
done

log "安装完成（版本 $VERSION build $BUILD）"
log "下一步：完全退出并重启 WPS，然后在「插入」标签的「公式编号」分组里使用。"
log "若功能区未出现，请删除 $JSADDONS/authaddin.json 后再次重启 WPS。"
