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

Run this and use the first path it prints:

```bash
for p in ./tablepngs.py ~/.claude/skills/tablepngs/tablepngs.py \
         ~/Documents/GitHub/TablePNGs-latex-public/tablepngs.py; do
  [ -f "$p" ] && echo "$p" && break
done
```

If none exists, search wider (`find ~ -name tablepngs.py -maxdepth 6 2>/dev/null`),
and if it is genuinely absent, tell the user and stop — do not hand-roll the
conversion. Refer to the found path as `$SCRIPT` below. Use `python3`
(`py` on Windows).

## Where to run it, and paths

Run the script from the directory containing the document, passing the bare
filename (`cd /path/to/doc && python3 $SCRIPT main.tex`). The document's own
relative paths — `\usepackage{../shared/style}`, `\includegraphics{../figures/x}`
— are resolved relative to the document, so a document that builds in place
will build here too. Outputs are written next to the document.

If you copy a document elsewhere to experiment, copy the whole tree its
relative paths reach (style files AND figure directories), or the baseline
compile will fail on a missing file that has nothing to do with tablepngs.

## Standard workflow

Budget the time: a short paper takes seconds, but a 100-page report with
20-odd tables compiles the document many times and can run for several
minutes. Give the command a generous timeout (10 minutes or more) and do not
kill it early — a half-finished run leaves confusing artifacts.

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
   You rarely need `--engine`: detection reads a `% !TEX program` comment,
   then the engine recorded in a previous build's `.log`/`.xdv`, then the
   preamble and any local `.sty`/`.cls` files it loads, and it retries other
   engines by itself if the baseline fails on an engine mismatch. Pass it only
   when you know detection was wrong. Add `--shell-escape` only if the
   original document needs it, and `--dpi 600` for print-grade images.

4. **Read the result**. Success looks like:
   - exit code 0,
   - `VERIFY PASS — 0/N table-text probes leaked; M caption(s) confirmed as live text`,
   - `VISUAL PASS — every flattened table matches its in-document rendering`,
   - outputs listed: `<stem>_tablepngs.tex`, `<stem>_tablepngs.pdf`,
     `<stem>_tablepngs/` with the PNGs.

   Note that warnings and per-table MISMATCH lines go to **stderr** while the
   per-table `[ok]` lines go to stdout, so a bare `... | tail` can hide
   failures. Capture both (`2>&1`), or you may see a run that looks clean but
   exited 3.

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

- **exit 1, "the ORIGINAL document failed to compile"** — tablepngs never got
  as far as flattening anything. Read the log excerpt before concluding the
  document is broken:
  - *File ... not found* — you are running from the wrong directory, or the
    tree was copied without the figure/style directories it references.
  - *fontspec / Unicode / engine mismatch* — the script retries other engines
    automatically for this class of error; if it still fails, pass
    `--engine xelatex` (or lualatex) explicitly.
  - anything else — the document genuinely does not build; fix that first,
    then re-run. Confirm by building it yourself the way the user does.
- **exit 1, "the flattened document ... failed to compile"** — the flattening
  itself produced LaTeX that does not build. The original is untouched. The
  message names the likely fixes; `--skip <n>` for the table named in the
  error is the reliable fallback. Report which table you had to skip.
- **"is <X>. tablepngs will NOT flatten it"** — a table in an environment the
  script does not support. That table stays live text and Word may still
  reflow it. Tell the user explicitly which table and offer to convert that
  environment to a supported one (plain `tabular`/`longtable`).
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
