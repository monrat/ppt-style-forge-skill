#!/usr/bin/env bash
# End-to-end smoke test for the PPT Style Forge execution layer.
#
# Builds the synthetic fixture, runs the full pipeline (inventory -> generate ->
# validate), and asserts that each step succeeds and validation PASSES. Run
# from the repo root:
#
#   bash tests/smoke_test.sh
#
# Exits non-zero on any failure. Requires python-pptx (generate_deck.py only);
# validate_deck.py is stdlib-only. render_qa.py is intentionally NOT asserted
# here because it needs a GUI renderer that CI may not have.
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"
# Prefer python3, fall back to python (Windows). Override with $PYTHON.
if [ -n "${PYTHON:-}" ]; then
  PY="$PYTHON"
elif command -v python3 >/dev/null 2>&1; then
  PY=python3
else
  PY=python
fi
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT

FIXTURE="tests/fixtures/sample.pptx"
SPEC="$TMP/smoke.deck-spec.json"
DECK="$TMP/smoke.pptx"

echo "== 0. fixture exists =="
[ -f "$FIXTURE" ] || { echo "FAIL: $FIXTURE missing"; exit 1; }
echo "ok"

echo "== 1. inventory runs (stdlib) =="
$PY scripts/extract_pptx_inventory.py "$FIXTURE" > "$TMP/inventory.txt"
grep -q "LAYOUTS:" "$TMP/inventory.txt" || { echo "FAIL: no LAYOUTS in inventory"; exit 1; }
echo "ok"

echo "== 2. generate runs (python-pptx) =="
cat > "$SPEC" <<'JSON'
{
  "slides": [
    {"layout": "Title Slide", "fields": {"ctrTitle": "Smoke Cover", "subTitle": "e2e test"}},
    {"layout": "Title and Content", "fields": {"title": "Point", "body": "one\ntwo"}}
  ]
}
JSON
$PY scripts/generate_deck.py "$SPEC" --template "$FIXTURE" --out "$DECK" > "$TMP/gen.log"
[ -f "$DECK" ] || { echo "FAIL: deck not produced"; exit 1; }
echo "ok"

echo "== 3. validate PASSES (stdlib, no python-pptx needed) =="
# Block python-pptx to prove validate_deck is truly stdlib-only.
$PY - "$DECK" "$SPEC" "$FIXTURE" <<'PYEOF' > "$TMP/val.log"
import sys, runpy
sys.modules['pptx'] = None  # ensure stdlib-only path
sys.argv = ['validate_deck.py', sys.argv[1], '--spec', sys.argv[2], '--template', sys.argv[3]]
runpy.run_path('scripts/validate_deck.py', run_name='__main__')
PYEOF
grep -q "PASS" "$TMP/val.log" || { echo "FAIL: validation did not PASS"; cat "$TMP/val.log"; exit 1; }
echo "ok"

echo "== 4. chart dangling-rel regression =="
# Inject a <c:chart r:id="rId99"> into the first active slide with NO matching
# relationship, then assert _empty_graphics detects it (regression for a bug
# where chart detection looked at the wrong element and missed all charts).
DANGLING="$TMP/dangling.pptx"
$PY - "$DECK" "$DANGLING" <<'PYEOF'
import sys, zipfile
from pathlib import Path
src, dst = Path(sys.argv[1]), Path(sys.argv[2])
# Locate the first active slide part so we inject where the validator reads.
sys.path.insert(0, "scripts")
import pptx_io
with zipfile.ZipFile(src) as zf:
    target = pptx_io.active_slide_parts(zf)[0]
gf = (
    '<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    ' xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<p:nvGraphicFramePr><p:cNvPr id="99" name="Chart 1"/><p:cNvGraphicFramePr/>'
    '<p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="100000" y="100000"/>'
    '<a:ext cx="5000000" cy="3000000"/></p:xfrm>'
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
    '<c:chart r:id="rId99"/></a:graphicData></a:graphic></p:graphicFrame>'
)
with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == target:
            txt = data.decode("utf-8")
            data = txt.replace("</p:spTree>", gf + "</p:spTree>").encode("utf-8")
        zout.writestr(item, data)
PYEOF
$PY - "$DANGLING" <<'PYEOF' > "$TMP/chart.log"
import sys, zipfile
sys.path.insert(0, "scripts")
import validate_deck as v
with zipfile.ZipFile(sys.argv[1]) as zf:
    res = v._empty_graphics(zf)
print("empty_graphics:", res)
assert any("rId99" in r for r in res), "REGRESSION: dangling chart not detected"
print("REGRESSION_OK")
PYEOF
grep -q "REGRESSION_OK" "$TMP/chart.log" || { echo "FAIL: chart regression"; cat "$TMP/chart.log"; exit 1; }
echo "ok"

echo
echo "ALL SMOKE TESTS PASSED"
