#!/usr/bin/env bash
# Regenerate the README gallery images in docs/gallery/.
# Requires the repo's own dependencies plus ImageMagick.
set -euo pipefail
cd "$(dirname "$0")/.."
ROOT="$PWD"
GAL="$ROOT/docs/gallery"
TMP="${TMPDIR:-/tmp}/tablepngs_gal"
mkdir -p "$TMP"
mkdir -p "$GAL"
rm -f "$TMP"/*.png

page () {  # page <pdf> <pagenum> <out.png>
  rm -f "$TMP"/pg-*.png
  mkdir -p "$TMP"
  pdftoppm -png -r 100 -f "$2" -l "$2" "$1" "$TMP/pg"
  mv "$TMP"/pg-*.png "$3"
}

pair () {  # pair <left> <right> <out> <leftlabel> <rightlabel>
  local F=/System/Library/Fonts/Supplemental/Arial.ttf
  [ -f "$F" ] || F=""
  local FA=(); [ -n "$F" ] && FA=(-font "$F")
  magick "$1" -resize 700x -bordercolor '#1f5fa8' -border 3 \
    \( -size 706x34 -background '#1f5fa8' -fill white -pointsize 17 \
       "${FA[@]}" -gravity center label:"$4" \) +swap -append "$TMP/l.png"
  magick "$2" -resize 700x -bordercolor '#c1651b' -border 3 \
    \( -size 706x34 -background '#c1651b' -fill white -pointsize 17 \
       "${FA[@]}" -gravity center label:"$5" \) +swap -append "$TMP/r.png"
  magick "$TMP/l.png" "$TMP/r.png" -background white -gravity north \
    -splice 16x0 +append -bordercolor white -border 12 "$3"
}

echo ">> 1/3 before-and-after page"
cd "$ROOT/examples/01_basic_booktabs"
pdflatex -interaction=nonstopmode main.tex >/dev/null
pdflatex -interaction=nonstopmode main.tex >/dev/null
python3 "$ROOT/tablepngs.py" main.tex --compare >/dev/null
page main.pdf 2 "$TMP/before.png"
page main_tablepngs.pdf 2 "$TMP/after.png"
pair "$TMP/before.png" "$TMP/after.png" "$GAL/before_after_page.png" \
  "ORIGINAL PDF  (tables are live text)" "FLATTENED PDF  (tables are images)"
cp main_tablepngs/_compare/t01-compare.png "$GAL/visual_check_pass.png"

echo ">> 2/3 landscape longtable chunks"
cd "$ROOT/examples/04_rotated"
python3 "$ROOT/tablepngs.py" main.tex >/dev/null
page main_tablepngs.pdf 5 "$TMP/l1.png"
page main_tablepngs.pdf 6 "$TMP/l2.png"
pair "$TMP/l1.png" "$TMP/l2.png" "$GAL/landscape_longtable.png" \
  "landscape longtable, page 1" "page 2 - clean break between images"

echo ">> 3/3 negative control (visual check catching a mangle)"
cd "$ROOT/tests/fixtures/mangle_control"
python3 "$ROOT/tablepngs.py" main.tex --compare >/dev/null 2>&1 || true
cp main_tablepngs/_compare/t01-compare.png "$GAL/visual_check_catches_mangle.png"

echo ">> cleaning generated outputs"
cd "$ROOT"
python3 tests/run_tests.py --no-compare 01 >/dev/null 2>&1 || true
rm -rf tests/fixtures/mangle_control/main_tablepngs* \
       tests/fixtures/mangle_control/main.{aux,log,pdf} \
       examples/04_rotated/main_tablepngs* examples/04_rotated/main.{aux,log,pdf} 2>/dev/null || true
rm -f "$TMP"/*.png
echo "gallery written to $GAL"
