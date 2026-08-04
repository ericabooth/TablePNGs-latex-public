#!/usr/bin/env python3
"""Render every tikzpicture in a LaTeX document to a cropped PNG.

Word cannot draw TikZ. Rather than lose the book's diagrams, compile each one
on its own with the book's real colour and style definitions, crop it, and hand
back a PNG that the .docx can embed in place of the picture.

Usage: render_tikz.py ASSEMBLED.tex PREAMBLE.tex OUTDIR
"""
import re, subprocess, sys, pathlib, shutil

src_path, preamble_path, outdir = sys.argv[1], sys.argv[2], pathlib.Path(sys.argv[3])
outdir.mkdir(parents=True, exist_ok=True)
src = pathlib.Path(src_path).read_text()
pre = pathlib.Path(preamble_path).read_text()

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
    tex.write_text(TEMPLATE % dict(libs=libs, colors=colors, tikzset=tikzset, body=body))
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

print(f'rendered {len(ok)} / {len(blocks)}')
for i, why in fail[:5]:
    print(f'  FAILED #{i}: {why[:300]}')
