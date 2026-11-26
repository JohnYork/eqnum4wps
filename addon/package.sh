#!/usr/bin/env bash
# 构建可分发安装包：dist/eqnum-<版本>-linux.{tar.gz,zip}
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJ="$(cd "$ROOT/.." && pwd)"
VERSION="$(sed -n 's/.*EQNUM_ADDON_VERSION = "\([^"]*\)".*/\1/p' "$ROOT/main.js" | head -n 1)"
[ -n "$VERSION" ] || { echo "无法从 main.js 读取版本号" >&2; exit 1; }

NAME="eqnum-${VERSION}"
DIST="$PROJ/dist"
STAGE="$(mktemp -d)"
PKG="$STAGE/$NAME"

cleanup() { rm -rf "$STAGE"; }
trap cleanup EXIT

mkdir -p "$PKG/addon/js" "$PKG/addon/ui" "$PKG/addon/images"
cp "$ROOT/index.html" "$ROOT/main.js" "$ROOT/ribbon.xml" "$PKG/addon/"
cp "$ROOT/js/"*.js "$PKG/addon/js/"
cp "$ROOT/ui/"* "$PKG/addon/ui/"
cp "$ROOT/images/"*.png "$PKG/addon/images/"

sed "s/__VERSION__/${VERSION}/g" "$ROOT/pkg/install.sh"   > "$PKG/install.sh"
sed "s/__VERSION__/${VERSION}/g" "$ROOT/pkg/uninstall.sh" > "$PKG/uninstall.sh"
sed "s/__VERSION__/${VERSION}/g" "$ROOT/pkg/README.txt"   > "$PKG/README.txt"
chmod +x "$PKG/install.sh" "$PKG/uninstall.sh"

bash -n "$PKG/install.sh"
bash -n "$PKG/uninstall.sh"

mkdir -p "$DIST"
rm -f "$DIST/${NAME}-linux.tar.gz" "$DIST/${NAME}-linux.zip"

tar -C "$STAGE" -czf "$DIST/${NAME}-linux.tar.gz" "$NAME"

python3 - "$STAGE" "$NAME" "$DIST/${NAME}-linux.zip" <<'PY'
import os
import sys
import zipfile

stage, name, out = sys.argv[1], sys.argv[2], sys.argv[3]
base = os.path.join(stage, name)

with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    for root, dirs, files in os.walk(base):
        dirs.sort()
        for f in sorted(files):
            full = os.path.join(root, f)
            rel = os.path.join(name, os.path.relpath(full, base))
            zi = zipfile.ZipInfo(rel)
            zi.compress_type = zipfile.ZIP_DEFLATED
            mode = 0o755 if f.endswith(".sh") else 0o644
            zi.external_attr = (mode | 0o100000) << 16
            with open(full, "rb") as fh:
                z.writestr(zi, fh.read())
print("[package] zip: %s" % out)
PY

echo "[package] 版本：$VERSION"
ls -lh "$DIST/${NAME}-linux.tar.gz" "$DIST/${NAME}-linux.zip"
