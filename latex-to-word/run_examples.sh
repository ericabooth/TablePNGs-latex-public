#!/usr/bin/env bash
# The two conversions this skill was built from, end to end.
set -euo pipefail
VENV=./venv                       # python3 -m venv venv; venv/bin/pip install python-docx
SOF="/Applications/LibreOffice.app/Contents/MacOS/soffice"

# ---------- Stata Journal article ----------
SJ=/path/to/StataJournal_GoogleTools
latexpand "$SJ/manuscript/googletools_sj.tex" > flat.tex
python3 scripts/prep_sj.py flat.tex prepped.tex "$SJ/figures" "$SJ/manuscript/googletools_sj.aux"
$VENV/bin/python scripts/make_ref.py ref_default.docx ref_sj.docx 8 11 0.9
pandoc prepped.tex -f latex -t docx -o googletools_sj.docx \
    --extract-media=media --reference-doc=ref_sj.docx \
    --citeproc --bibliography=enriched.bib

# ---------- kaobook manuscript ----------
BK=/path/to/LaTeXBookCode
# ref_default.docx ships with the scripts; regenerate it if pandoc changed:
#   pandoc --print-default-data-file reference.docx > scripts/ref_default.docx
# The .aux argument to render_tikz.py resolves \ref inside diagrams; without
# it, any in-diagram cross-reference ships as "??" baked into the PNG.
python3 scripts/render_tikz.py "$BK/src/main_assembled.tex" "$BK/src/00_preamble.tex" tikz "$BK/main.aux"
python3 scripts/prep_book.py "$BK/src/main_assembled.tex" prepped.tex tikz "$BK/images" "$BK/main.aux"
$VENV/bin/python scripts/make_ref.py scripts/ref_default.docx ref_book.docx 8.5 11 1.0
pandoc prepped.tex -f latex -t docx -o main_raw.docx \
    --extract-media=media --reference-doc=ref_book.docx \
    --citeproc --bibliography="$BK/main.bib"
$VENV/bin/python scripts/polish_docx.py main_raw.docx main.docx
$VENV/bin/python scripts/verify_docx.py main.docx prepped.tex "$BK/main.pdf"

# ---------- verify ----------
for f in googletools_sj main; do
  "$SOF" --headless --convert-to pdf $f.docx --outdir . >/dev/null
  echo "$f: $(pdfinfo $f.pdf | awk '/^Pages/{print $2}') pages"
  pdftotext $f.pdf - | grep -oE '\\[a-zA-Z]{2,}' | sort | uniq -c | sort -rn | head
done
