#!/usr/bin/env python3
"""Render every tikzpicture in a LaTeX document to a cropped PNG.

Word cannot draw TikZ. Rather than lose the book's diagrams, compile each one
on its own with the book's real colour and style definitions, crop it, and hand
back a PNG that the .docx can embed in place of the picture.

Usage: render_tikz.py ASSEMBLED.tex PREAMBLE.tex OUTDIR [AUX]

Pass the document's .aux whenever any diagram contains \ref or \pageref:
the standalone compile has no label table, so without it every in-diagram
cross-reference renders as "??" and ships that way into the .docx. The
manifest still hashes the RAW block text, so aux substitution does not
disturb the order-integrity guard prep_book.py checks.
"""
import re, subprocess, sys, pathlib, shutil

src_path, preamble_path, outdir = sys.argv[1], sys.argv[2], pathlib.Path(sys.argv[3])
aux_path = sys.argv[4] if len(sys.argv) > 4 else ''
outdir.mkdir(parents=True, exist_ok=True)
src = pathlib.Path(src_path).read_text()
pre = pathlib.Path(preamble_path).read_text()

# Label table from the .aux, for resolving \ref inside diagrams.
labs = {}
if aux_path:
    aux = pathlib.Path(aux_path).read_text(errors='replace')
    for m in re.finditer(r'\\newlabel\{([^}]+)\}\{\{([^{}]*)\}', aux):
        labs[m.group(1)] = m.group(2)

def resolve_refs(body):
    """Replace \\ref{x} (and \\S\\ref{x}) with the aux number, for the
    standalone compile only. Unknown labels are left alone, so the "??"
    stays visible instead of being papered over with wrong text."""
    if not labs:
        if re.search(r'\\ref\{', body):
            sys.stderr.write('WARNING: diagram contains \\ref but no AUX was '
                             'given; it will render as "??"\n')
        return body
    return re.sub(r'\\ref\{([^}]+)\}',
                  lambda m: labs.get(m.group(1), m.group(0)), body)

# Pull only what a standalone tikz compile needs: colours and the tikzset block.
colors = '\n'.join(re.findall(r'\\definecolor\{[^}]*\}\{[^}]*\}\{[^}]*\}', pre))
m = re.search(r'\\tikzset\{.*?\n\}', pre, re.S)
tikzset = m.group(0) if m else ''
libs = '\n'.join(re.findall(r'\\usetikzlibrary\{[^}]*\}', pre))

TEMPLATE = r"""\documentclass[border=4pt]{standalone}
\usepackage[T1]{fontenc}
\usepackage{lmodern}
\usepackage{amsmath,amssymb}
\usepackage{xcolor}
\usepackage{tikz}
%(libs)s
\usetikzlibrary{arrows.meta,positioning,shapes.geometric,fit,backgrounds,calc,shadows.blur}
%(colors)s
%(tikzset)s
\newcommand{\marginnote}[2][]{}
\begin{document}
%(body)s
\end{document}
"""

blocks = []
for mm in re.finditer(r'\\begin\{tikzpicture\}.*?\\end\{tikzpicture\}', src, re.S):
    blocks.append((mm.start(), mm.end(), mm.group(0)))

print(f'found {len(blocks)} tikzpicture blocks')
work = outdir / '_work'
work.mkdir(exist_ok=True)
ok, fail = [], []
for i, (a, b, body) in enumerate(blocks, 1):
    stem = f'tikz{i:02d}'
    tex = work / f'{stem}.tex'
    tex.write_text(TEMPLATE % dict(libs=libs, colors=colors, tikzset=tikzset,
                                   body=resolve_refs(body)))
    r = subprocess.run(['pdflatex', '-interaction=nonstopmode', '-halt-on-error', f'{stem}.tex'],
                       cwd=work, capture_output=True, text=True)
    pdf = work / f'{stem}.pdf'
    if r.returncode != 0 or not pdf.exists():
        fail.append((i, (r.stdout or '')[-400:]))
        continue
    png = outdir / f'{stem}.png'
    subprocess.run(['pdftoppm', '-png', '-r', '300', '-singlefile', str(pdf), str(png)[:-4]],
                   capture_output=True)
    if png.exists():
        ok.append(stem)
    else:
        fail.append((i, 'pdftoppm produced nothing'))

# Manifest: hash of each block's source, so prep_book.py can verify that the
# document has not changed between rendering and substitution. Without this, an
# edit that adds or removes a tikzpicture silently shifts every later diagram
# into the wrong figure.
import hashlib, json
manifest = {f'tikz{i:02d}': hashlib.md5(b[2].encode()).hexdigest()
            for i, b in enumerate(blocks, 1)}
(outdir / 'tikz_manifest.json').write_text(json.dumps(manifest, indent=1))

print(f'rendered {len(ok)} / {len(blocks)}')
for i, why in fail[:5]:
    print(f'  FAILED #{i}: {why[:300]}')
