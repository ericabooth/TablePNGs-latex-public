---
name: tablepngs
description: >
  Flatten LaTeX tables into high-resolution PNGs so PDF-to-Word conversion
  (Adobe Acrobat export) leaves them intact. Use whenever the user wants to
  convert a LaTeX-produced PDF to Word without tables breaking, mentions
  tables getting mangled/reformatted in Word, asks to rasterize/flatten
  tables in a .tex document, or asks to run tablepngs. Drives the
  tablepngs.py script — never hand-roll the conversion.
---

# tablepngs — flatten LaTeX tables for clean PDF-to-Word conversion

You drive `tablepngs.py`. The script does ALL the work: scanning the
document, re-typesetting each table with the document's own preamble,
rasterizing, rewriting, compiling, and verifying. Do not reimplement any of
those steps by hand (no ad-hoc standalone documents, no manual pdftoppm
calls, no manual edits to swap tables for images), even if a run fails —
diagnose and re-run the script instead.

## Locating the script

1. If this repo is TablePNGs-latex-public, it is `./tablepngs.py`.
2. Otherwise check `~/.claude/skills/tablepngs/tablepngs.py` (bundled copy,
   may exist if the user copied the skill).
3. Otherwise search: `ls ~/Documents/GitHub/TablePNGs-latex-public/tablepngs.py`.
4. If not found anywhere, offer to clone/download it, or ask the user where
   it lives.

Set `SCRIPT` to the found path; run everything with `python3` (`py` on
Windows).

## Standard workflow

1. **Doctor first** (fast, catches most problems):
   `python3 $SCRIPT --check`
   If required tools are missing, the output includes the exact install
   command for this OS. Offer to run the install (brew/apt/choco), then
   re-run the doctor to confirm.

2. **Dry run** to show the user what will be flattened:
   `python3 $SCRIPT path/to/main.tex --list`
   Report the table list (index, environment, line, caption). If the user
   only wants some tables, note the indices for `--only`/`--skip`.

3. **Convert, with the visual check on**:
   `python3 $SCRIPT path/to/main.tex --compare`
   Always pass `--compare` unless the user asks you not to (or ImageMagick is
   unavailable) — it is what proves the images are not mangled. Add
   `--engine xelatex|lualatex` only if auto-detection picked wrong (the script
   auto-detects from `% !TEX program` magic comments and fontspec/polyglossia
   usage). Add `--shell-escape` only if the original document needs it. Use
   `--dpi 600` if the user wants print-grade images.

4. **Read the result**. Success looks like:
   - exit code 0,
   - `VERIFY PASS — 0/N table-text probes leaked; M caption(s) confirmed as live text`,
   - `VISUAL PASS — every flattened table matches its in-document rendering`,
   - outputs listed: `<stem>_tablepngs.tex`, `<stem>_tablepngs.pdf`,
     `<stem>_tablepngs/` with the PNGs.

5. **Look at the comparison sheets yourself** (see below), then tell the user:
   open `<stem>_tablepngs.pdf` in Acrobat and export to Word as usual; the
   original .tex was not modified.

## The visual verification routine (do this every run)

`--compare` writes one side-by-side sheet per table to
`<stem>_tablepngs/_compare/tNN-compare.png`. Each sheet has the table as it is
typeset in the real document on the **left, blue-bordered** ("REFERENCE") and
the flattened PNG that went into the output on the **right, orange-bordered**
("FLATTENED").

The script screens these automatically two ways:

- **content check** (all tables): every literal word and number written in the
  table source must appear in the rasterized image's text layer. Catches
  dropped rows and cells. Page-break invariant, so it works for longtables.
- **pixel check** (tables rendered via `preview`, i.e. everything except
  longtables): the flattened PNG must match the in-context rendering exactly.
  A clean run reports `pixels 0.0000`. Longtables are page-cropped by a
  different path, so their pixel score is reported as `n/a (see sheet)` and
  the content check plus your own eyes are authoritative there.

**After the run, Read the sheets** with the Read tool — you can see images, so
use that. For each one confirm: same numbers, same column structure, same
row count, nothing clipped at an edge, no `?` marks where citations or
references should be. Report anything that looks off, naming the table.
This is the step that catches problems the text-layer check cannot: a table
whose snippet compiled cleanly but rendered the wrong content (for example a
macro redefined mid-document, where the isolated snippet still sees the
preamble's older definition).

If the script reports `MISMATCH`:
- **content tokens missing** — the flattened image genuinely lost content.
  Read the sheet, identify what is missing, and look for the cause in the
  source (usually a macro or definition that only exists mid-document).
  Report it; do not silently ship the output.
- **pixel difference above threshold** — compare the two halves of the sheet
  visually. If they look identical, it may be a harmless rendering
  difference; say so explicitly and show your reasoning. If they differ,
  treat it as a real defect.

The repo ships a negative control at `tests/fixtures/mangle_control/` that is
built to be mangled; running the script on it MUST report a MISMATCH. Use it
to sanity-check that the visual verification is working if you ever doubt it.

## Interpreting exit codes and failures

- **exit 1, "the ORIGINAL document failed to compile"** — the user's
  document is broken independent of tablepngs. Show the log excerpt, help
  fix the document, then re-run.
- **exit 1, a specific table "failed to compile — leaving this table
  as-is"** — re-run with `--keep-build`, then read
  `<stem>_tablepngs/_build/tNN/tNN.log`. Common causes: the table body uses
  a package the preamble loads conditionally, or a `\caption` variant the
  scanner lifted incorrectly. Workarounds: fix the document, or exclude that
  table with `--skip N` and tell the user which table was left as text.
- **exit 2, VERIFY FAILED (leaked probes)** — some table text is still in
  the final PDF's text layer. Usually part of a table sat outside the
  detected environment (e.g. notes typed after `\end{tabular}` but outside
  any environment). Inspect the named table in the source; widen the
  environment or accept and explain.
- **exit 3, visual check FAILED** — the flattened image does not match the
  in-document rendering. See the visual verification section above; this is
  a content problem, not a cosmetic one, so investigate before shipping.
- **"--compare needs ImageMagick"** — install it (`brew install imagemagick`
  / `sudo apt install imagemagick` / `choco install imagemagick`), or run
  without `--compare` and tell the user the visual check was skipped.
- **"caption text not found in text layer"** — a caption ended up inside an
  image (multi-caption float) or was dropped. Check the warning list; the
  multi-caption case is a documented limitation (captions baked in, counter
  still advanced).
- **Python traceback** — a genuine tablepngs bug. Capture the traceback and
  the minimal document, report it to the repo, and as a stopgap try
  `--skip` on the offending table.

## Dependency troubleshooting

- **No rasterizer**: install poppler (`brew install poppler` /
  `sudo apt install poppler-utils` / `choco install poppler`). ImageMagick
  and Ghostscript are automatic fallbacks.
- **ImageMagick "not authorized" on PDF (Linux)**: the distro's
  `/etc/ImageMagick-*/policy.xml` blocks PDF input. Do NOT edit system
  policy; install poppler instead — the script prefers pdftoppm when
  present.
- **pdfcrop missing**: only multi-page longtables need it
  (`tlmgr install pdfcrop`, or `sudo apt install texlive-extra-utils`).
  Without it the script still works but longtable page images keep
  full-page margins.
- **fontspec document failing under pdflatex**: the engine was forced
  incorrectly; drop `--engine` or use `--engine xelatex`.
- **Windows**: use `py` instead of `python3`; ensure MiKTeX/TeX Live bin dir
  is on PATH (`where pdflatex`).

## What to tell the user about limitations

Only when relevant: beamer frames are out of scope; multi-caption floats
bake captions into the image; `\footnote` inside a table body is baked into
the image; biblatex (not bibtex/natbib) citations inside cells render as
`[?]` in the image.
