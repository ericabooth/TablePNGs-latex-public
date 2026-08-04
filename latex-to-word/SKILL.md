---
name: latex-to-word
description: Convert a LaTeX manuscript (book, journal article, or report) into a clean, genuinely editable Word .docx, with verification. Use when a co-author, editor, or publisher who does not use LaTeX needs a Word version, or when an existing .docx conversion is garbled.
---

# LaTeX to Word

Goal: a `.docx` a human can actually edit. Real Word heading styles so the navigation
pane works, real footnotes, real tables, embedded figures, resolved cross-references,
and a formatted bibliography.

---

## Rule zero

**Convert from the LaTeX source, never from the PDF.**

A PDF-derived `.docx` looks plausible and is unusable. Tells: the table of contents
appears as literal text with rows of dots; author names run together
(`Eric BoothElizabeth Teas`); every line is its own paragraph; no heading styles, so
Word's navigation pane is empty; running headers merge into sentences
(`Teas 2 3 Interacting with a web API`).

If handed a PDF, find the `.tex`. If it truly does not exist, say so rather than
shipping a PDF conversion.

---

## Step 1: DECIDE — inspect before you write any code

Run this triage first. The answers determine everything downstream.

```bash
# A. What class is it? (drives which prep script to start from)
head -5 main.tex | grep documentclass

# B. What macros exist, by frequency? Anything with a real count needs a decision.
grep -ohE '\\[a-zA-Z]+' flat.tex | sort | uniq -c | sort -rn | head -40

# C. What environments exist? Custom ones will be silently dropped.
grep -ohE '\\begin\{[a-zA-Z*]+\}' flat.tex | sort | uniq -c | sort -rn

# D. How wide is the widest code line? (sets the monospace font size)
python3 -c "import re;b=re.findall(r'\\\\begin\{verbatim\}(.*?)\\\\end',open('flat.tex').read(),re.S);print(max(len(l) for x in b for l in x.split(chr(10))))"

# E. Are there drawn diagrams? (TikZ/pgfplots must be pre-rendered)
grep -c 'begin{tikzpicture}' flat.tex

# F. Where does the bibliography live? .bib, .bbl, or both?
ls *.bib *.bbl 2>/dev/null; grep -c 'url{' *.bbl 2>/dev/null
```

### Decision table

| What you find | What to do |
|---|---|
| Custom class (`kaobook`, `statapress`, in-house) | Expect 10–30 unknown macros. Budget for a prep script; do not try raw pandoc. |
| Standard `article`/`report`/`book`, no custom envs | Raw pandoc may be enough. Still run the verification in Step 5. |
| `tikzpicture` present | Pre-render each to PNG (`render_tikz.py`). Word cannot draw TikZ. |
| Margin notes (`\marginnote`, `\sidenote`, tufte-style) | Map to `\footnote`. Then set `Footnote Text` explicitly — see Step 3. |
| Custom callout boxes (tcolorbox etc.) | Bold title paragraph + `quote` block. |
| Longest code line ≤ 75 chars | Monospace can be 10pt. Comfortable. |
| Longest code line 75–95 chars | Monospace must be 8–8.5pt or it wraps mid-word. |
| Longest code line > 95 chars | Consider landscape sections, or reflow the source. |
| `.bbl` has URLs the `.bib` lacks | Harvest URLs from `.bbl` into a working `.bib` copy. |
| Verbatim inside a `figure` float | **Unwrap it.** See "The code-in-figure trap". |
| Wide floats (`figure*`, `table*`) | Flatten to `figure`/`table`; Word has no wide-float concept. |

### Which prep script to start from

- **Journal article, `\input` fragments, syntax diagrams, console logs** → `prep_sj.py`
- **Book, chapters, margin notes, callout boxes, parts, index** → `prep_book.py`

Neither is universal. Copy the closer one and adapt using the Step-1 inventory.

---

## Step 2: PREPARE the source

```bash
latexpand main.tex > flat.tex            # resolves \input/\include, NOT the bibliography
python3 scripts/prep_<kind>.py flat.tex prepped.tex FIGDIR AUX
```

### Construct mapping

| LaTeX | Word | Why |
|---|---|---|
| `\marginnote`, `\sidenote` | `\footnote` | Word has no margin column |
| tcolorbox callout | bold title + `quote` | survives as styled, editable text |
| `tikzpicture` | pre-rendered PNG | Word cannot draw TikZ |
| `\index{...}` | delete | meaningless without a generated index |
| `\ref`, float numbers | resolved from `.aux` | pandoc cannot number floats |
| `\pageref` | a figure/section reference | page numbers do not survive reflow |
| `figure*` / `table*` | `figure` / `table` | no wide floats in Word |
| console box-drawing macros | Unicode `│ ─ ┌ ┐ └ ┘ ├ ┤` | keeps the box editable |
| `\optional{X}` (SJ) | `[X]` | that is what it means |
| `\underbar{sh}eet` (SJ) | `\underline{sh}eet` | marks minimum abbreviation |

Numbers come from the `.aux`:
```python
labs = {m.group(1): m.group(2) for m in
        re.finditer(r'\\newlabel\{([^}]+)\}\{\{([^}]*)\}\{([^}]*)\}', aux)}
```

### Traps that cost real time

**`\\` immediately before `[` eats the bracket.** LaTeX reads `\\[2em]` as a line
break with optional vertical space, so `\\` followed by `[options]` silently deletes
the options. Stata syntax diagrams are full of this.
```python
src = re.sub(r'\\\\(\s*\n\s*)\[', r'\\\\{}\1[', src)
```

**The code-in-figure trap.** A `verbatim` block inside a `figure` float is dropped,
and pandoc renders the wreckage as a **malformed Word table**. On the reference book
this cost 50 lines of Stata code and produced 7 phantom tables — and the table count
looked *higher*, not lower, so it read as success. Unwrap any figure containing
verbatim before converting. Check with:
```bash
python3 -c "import re;t=open('prepped.tex').read();print(sum(1 for f in re.findall(r'\\\\begin\{figure\}.*?\\\\end\{figure\}',t,re.S) if 'begin{verbatim}' in f))"
```

**Unknown macros vanish mid-word.** `\underbar{sh}eet(tabname)` → `eet(tabname)`;
`\hskip 4em` → the literal text `4em`. Invisible to word counts; visible on a page.

**Silent whole-element loss.** A `\keywords{...}` line disappeared entirely. Only the
sentence-level diff in Step 5 catches this.

**Macros with several optional arguments.** `\setpartpreamble[uc][.75\textwidth]{...}`
needs `(?:\[[^\]]*\])*` and must be unwrapped as a balanced brace group, or you
delete the opener and orphan its `}`.

**`\resizebox{w}{h}{$math$}`** defeats pandoc's math reader. Unwrap all three
arguments, keep the third, then collapse any `\[$...$\]` left behind.

**Escapes in `re.sub` replacements.** `re.sub(p, '\\caption{...}', s)` raises
`bad escape \c`. Use a lambda for any replacement containing backslashes.

---

## Step 3: STYLE the reference document

```bash
pandoc --print-default-data-file reference.docx > ref_default.docx
venv/bin/python scripts/make_ref.py ref_default.docx ref.docx MONO_PT BODY_PT MARGIN_IN
```

**Size the monospace font by measurement, not taste.** Text width in points ÷ longest
line ÷ 0.6 = the largest font that will not wrap. A 92-char Stata log in a 468pt
measure needs 8pt; a 74-char book listing can take 10pt.

**`Verbatim Char` must have NO explicit size.** An absolute size on the inline-code
character style wins inside footnotes and table cells, so `` `code` `` renders *larger*
than the 9pt footnote text around it. Clear it so it inherits; size only the
`Source Code` paragraph style, which governs blocks.

**Set `Footnote Text` explicitly.** A tufte-style source can produce hundreds of
footnotes (333 on the reference book). Left at default they compete with body text.
9pt against an 11pt body reads as clearly subordinate.

---

## Step 4: CONVERT

```bash
pandoc prepped.tex -f latex -t docx -o out.docx \
  --extract-media=media --reference-doc=ref.docx \
  --citeproc --bibliography=refs.bib
venv/bin/python scripts/polish_docx.py out.docx final.docx
```

`--citeproc` is **required**: pandoc will not resolve `\citep` against an inlined
`thebibliography`, and citations silently render as nothing.

- **Placement.** citeproc appends the list at the very end. `\hypertarget{refs}{}`
  becomes a `div#refs` and puts it where `\bibliography` was.
- **Missing URLs.** If the `.bbl` is richer than the `.bib`, harvest `\url{}` per
  `\bibitem` key and inject `url = {...}` into a working copy. This recovered 9 of
  19 URLs on the reference article.

`polish_docx.py` fixes what is easier to correct in the finished file: table cells
default to bottom alignment, so a wrapping header drops the other headers and the row
looks broken. It top-aligns all cells, bolds header rows, and marks them to repeat
across page breaks.

---

## Step 5: VERIFY — every item, every time

```bash
soffice --headless --convert-to pdf final.docx --outdir .
```

1. **Zero pandoc warnings.** Each one is a dropped or mangled element.
2. **Word retention** vs the original PDF: expect **94–97%**. The shortfall is the
   original's running headers and TOC. Below ~90% means real loss.
3. **Sentence diff.** Normalize and list original sentences absent from the output.
   Most hits are page furniture; read them anyway — that is how the missing
   `\keywords` line surfaced.
4. **Code fidelity, line by line.** Compare every `Source Code` paragraph against the
   source `verbatim` blocks. This is what caught the code-in-figure trap:
   ```python
   src_lines = [l for b in verbatim_blocks for l in b.split('\n') if l.strip()]
   doc_lines = [l for p in docx_code_paras for l in p.split('\n') if l.strip()]
   missing = [l for l in src_lines if l not in doc_lines]   # must be empty
   ```
5. **Element counts match the source:** `Source Code` paragraphs == verbatim blocks;
   tables == `table` environments; images == figures. A count that is *higher* than
   the source is as suspicious as one that is lower.
6. **LaTeX leakage:** `pdftotext out.pdf - | grep -oE '\\[a-zA-Z]{2,}'` should return
   only genuine content (Windows paths, prose that discusses LaTeX).
7. **Heading hierarchy:** Part→H1, Chapter→H2, Section→H3, Subsection→H4.
8. **Look at rendered pages.** Front matter, a figure page, a table page, the
   code-heaviest page, the footnote-heaviest page, and the references. Find those
   pages programmatically rather than sampling blindly:
   ```python
   # score each page for code lines / footnote markers / table captions, take the max
   ```
   Several defects here were invisible to every automated check.

---

## Surgical edits after delivery

When one figure changes and the recipient may already be editing the `.docx`, patch
in place instead of rebuilding.

```bash
unzip -q delivered.docx -d work && cd work
# identify which media file is which: compare each against the source figures
magick compare -metric RMSE -resize 200x200! word/media/rIdNN.png source_fig.png null:
```
An RMSE near zero means unchanged; a large value means that is the one that moved.
Then replace `word/media/rIdNN.png`, and **correct the drawing extent** in
`word/document.xml` or Word will stretch the new image to the old aspect ratio:

```python
new_cy = round(cx * new_h / new_w)      # keep width, recompute height (EMU)
```
Re-zip, then open with python-docx and render to PDF to confirm.

---

## How the two reference documents differed

| | Stata Journal article | kaobook manuscript |
|---|---|---|
| Class | `statapress.cls` | `kaobook` (tufte-style) |
| Size | 38pp → 31pp docx | 330pp → 288pp docx |
| Hard part | syntax diagrams, console box art | 35 TikZ diagrams, 333 margin notes |
| Fragments | 23 `\input` files → `latexpand` | already assembled by a build script |
| Diagrams | none (figures already PNG) | `render_tikz.py`, sized by aspect ratio |
| Mono size | **8pt** (92-char logs) | **10pt** (74-char listings) |
| Bibliography | `.bbl` richer than `.bib`; URLs harvested | `.bib` sufficient |
| Unique fix | `\\`+`[` guard; box-drawing → Unicode | code-in-figure unwrap; footnote sizing |
| Result | 96.9% words, 19/19 URLs, 0 warnings | 95.2% words, 67 images, 33 tables, 0 warnings |

The lesson: **the class determines the work.** Same pipeline, different prep script,
different font size, different traps. Always run Step 1 before assuming.

---

## Setup

```bash
brew install --cask libreoffice          # headless docx rendering + verification
python3 -m venv venv && venv/bin/pip install python-docx lxml Pillow
```
`pandoc`, `latexpand`, `pdflatex`, `pdftoppm`, `pdfinfo`, and ImageMagick come with
TeX Live and Homebrew. macOS `pip` refuses system installs (PEP 668); use a venv.

## Scripts

- `scripts/prep_sj.py` — Stata Journal: `stlog`, `stsyntax`, `\optional`,
  `\underbar`, `\hangpara`, console box-drawing.
- `scripts/prep_book.py` — kaobook: margin notes, callouts, wide floats,
  `\setpartpreamble`, index, TikZ substitution, code-in-figure unwrap.
- `scripts/render_tikz.py` — compiles each `tikzpicture` standalone with the real
  preamble colours/styles, crops, exports 300dpi PNG.
- `scripts/make_ref.py` — builds the reference.docx (fonts, sizes, margins,
  footnote and block-text styling).
- `scripts/polish_docx.py` — post-conversion table fixes.
