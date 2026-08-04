#!/usr/bin/env python3
"""Build a pandoc reference.docx tuned for a LaTeX-derived manuscript.

The two things that matter for a converted technical document:
  1. Monospaced blocks must be small enough that the widest captured console
     line fits the text measure without wrapping mid-word.
  2. Body text must be a real serif at a readable size, because the whole
     point of the .docx is that a human edits it.
"""
import sys, docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_LINE_SPACING

ref_in, ref_out = sys.argv[1], sys.argv[2]
mono_pt   = float(sys.argv[3]) if len(sys.argv) > 3 else 8.0
body_pt   = float(sys.argv[4]) if len(sys.argv) > 4 else 11.0
margin_in = float(sys.argv[5]) if len(sys.argv) > 5 else 1.0

d = docx.Document(ref_in)

BODY_FONT = 'Cambria'      # ships with Word on macOS and Windows
MONO_FONT = 'Consolas'     # ships with Word; falls back gracefully
HEAD_FONT = 'Calibri'

def setfont(style, name=None, size=None, bold=None, color=None):
    try:
        f = style.font
    except Exception:
        return
    if name:  f.name = name
    if size == 'inherit':
        f.size = None            # clear any explicit size so it inherits
    elif size:
        f.size = Pt(size)
    if bold is not None: f.bold = bold
    if color: f.color.rgb = color

for s in d.styles:
    n = s.name
    if n in ('Normal', 'Body Text', 'First Paragraph', 'Compact', 'Abstract'):
        setfont(s, BODY_FONT, body_pt)
    elif n.startswith('Heading'):
        setfont(s, HEAD_FONT, None, True, RGBColor(0x1F, 0x36, 0x64))
    elif n in ('Title',):
        setfont(s, HEAD_FONT, body_pt + 8, True, RGBColor(0x1F, 0x36, 0x64))
    elif n in ('Author', 'Date', 'Subtitle'):
        setfont(s, BODY_FONT, body_pt)
    elif n == 'Verbatim Char':
        # Inline code. Deliberately NO explicit size: an absolute size here wins
        # inside footnotes and table cells, so `code` would render larger than the
        # 9pt footnote text around it. Inheriting keeps it in scale everywhere.
        setfont(s, MONO_FONT, 'inherit')
    elif n == 'Source Code':
        setfont(s, MONO_FONT, mono_pt)
    elif 'Caption' in n:
        setfont(s, BODY_FONT, body_pt - 1)
    elif n == 'Footnote Text':
        # A tufte-style source turns margin notes into footnotes, so there can be
        # hundreds. Set them explicitly: clearly subordinate, still readable.
        setfont(s, BODY_FONT, body_pt - 2)
    elif n == 'Block Text':
        # Callout boxes land here; keep them a touch smaller than body so they
        # read as set-off material rather than as more prose.
        setfont(s, BODY_FONT, body_pt - 0.5)

# Source Code may not exist as a paragraph style in the default reference;
# pandoc creates it from Verbatim Char. Add it so the size is honoured.
try:
    sc = d.styles['Source Code']
except KeyError:
    from docx.enum.style import WD_STYLE_TYPE
    sc = d.styles.add_style('Source Code', WD_STYLE_TYPE.PARAGRAPH)
    sc.base_style = d.styles['Normal']
setfont(sc, MONO_FONT, mono_pt)
sc.paragraph_format.space_before = Pt(4)
sc.paragraph_format.space_after  = Pt(4)
sc.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

try:
    ft = d.styles['Footnote Text']
    ft.paragraph_format.space_after = Pt(2)
    ft.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
except KeyError:
    pass
try:
    bt = d.styles['Block Text']
    bt.paragraph_format.left_indent = Inches(0.3)
    bt.paragraph_format.space_before = Pt(4)
    bt.paragraph_format.space_after = Pt(6)
except KeyError:
    pass

for sec in d.sections:
    sec.left_margin = sec.right_margin = Inches(margin_in)
    sec.top_margin  = sec.bottom_margin = Inches(margin_in)

d.save(ref_out)
print(f'wrote {ref_out}: body {BODY_FONT} {body_pt}pt, mono {MONO_FONT} {mono_pt}pt, margins {margin_in}in')
