#!/usr/bin/env python3
"""Preprocess a kaobook manuscript so pandoc can produce a clean, editable .docx.

Usage: prep_book.py ASSEMBLED.tex OUT.tex TIKZDIR FIGDIR AUX

kaobook is a tufte-style class: margin notes, wide floats, custom tcolorbox
callouts, TikZ diagrams, an index. None of it exists in Word. Each construct is
rewritten into the nearest thing Word does have, chosen so an editor can still
read and change the content:

  margin note / sidenote  -> footnote        (Word's own margin-ish apparatus)
  tcolorbox callout       -> bold title + block quote
  tikzpicture             -> pre-rendered PNG
  index entry             -> dropped (meaningless without a generated index)
  \\ref / \\caption number  -> resolved from the .aux
"""
import re, sys, pathlib

src   = pathlib.Path(sys.argv[1]).read_text()
OUT   = sys.argv[2]
TIKZ  = sys.argv[3]
FIGD  = sys.argv[4]
AUX   = sys.argv[5] if len(sys.argv) > 5 else ''


def strip_braced(text, macro):
    """Delete \\macro{...} entirely, honouring nested braces."""
    out, i, tag = [], 0, '\\' + macro + '{'
    while True:
        j = text.find(tag, i)
        if j < 0:
            out.append(text[i:]); return ''.join(out)
        out.append(text[i:j])
        k, depth = j + len(tag), 1
        while k < len(text) and depth:
            if text[k] == '{': depth += 1
            elif text[k] == '}':
                depth -= 1
                if depth == 0: break
            k += 1
        i = k + 1


def rewrite_braced(text, macro, open_s, close_s, optional=False):
    """Replace \\macro[opt]{...} with open_s + body + close_s."""
    out, i = [], 0
    pat = re.compile(r'\\' + macro + (r'(?:\[[^\]]*\])*' if optional else '') + r'\{%?\s*')
    while True:
        m = pat.search(text, i)
        if not m:
            out.append(text[i:]); return ''.join(out)
        out.append(text[i:m.start()])
        k, depth = m.end(), 1
        while k < len(text) and depth:
            if text[k] == '{': depth += 1
            elif text[k] == '}':
                depth -= 1
                if depth == 0: break
            k += 1
        out.append(open_s + text[m.end():k] + close_s)
        i = k + 1


# ----------------------------------------------- 0. body only, drop preamble
b = src.find('\\begin{document}')
e = src.rfind('\\end{document}')
body = src[b + len('\\begin{document}'):e]

# --------------------------------------- 0b. kaobook front-matter plumbing
# These exist only to drive kaobook's page furniture. Unwrap the ones that hold
# real text (the colophon) and delete the rest.
body = rewrite_braced(body, 'lowertitleback', '\n\n', '\n\n')
body = rewrite_braced(body, 'uppertitleback', '\n\n', '\n\n')
body = re.sub(r'\\title\[[^\]]*\]', r'\\title', body)
# \author{A \hspace{..} B}: the spacer is the only separator, and stripping it
# would run the names together. Replace it with a word before it is removed.
body = re.sub(r'(\\author\{[^}]*?)\\hspace\*?\{[^}]*\}\s*', r'\1 and ', body)
body = re.sub(r'\\markboth\{[^}]*\}\{[^}]*\}', '', body)
for mac in ('setchapterstyle', 'chapterstyle', 'vspace', 'hspace', 'vspace*', 'hspace*'):
    body = strip_braced(body, mac)
body = re.sub(r'\\(makeatletter|makeatother|nopagebreak|bigskip|medskip|smallskip|'
              r'noindent|clearpage|newpage|normalsize|small|footnotesize|centering|'
              r'raggedright|raggedleft|par)\b', ' ', body)
# \setpartpreamble wraps a part's blurb; keep the blurb as ordinary text.
body = rewrite_braced(body, 'setpartpreamble', '\n\n', '\n\n', optional=True)

# ---------------------------------------------------------- 1. TikZ -> PNG
# Blocks are numbered in document order by render_tikz.py, so replace in order.
counter = [0]

# Refuse to substitute against a stale render. tikz_manifest.json records the
# hash of each block at render time. Verify against the RAW source here, before
# any preprocessing touches the text: the cleanup passes below edit macros
# inside tikz blocks too, so hashing after them would always mismatch.
import hashlib, json
_mpath = pathlib.Path(TIKZ) / 'tikz_manifest.json'
if _mpath.exists():
    _manifest = json.loads(_mpath.read_text())
    _raw_blocks = re.findall(r'\\begin\{tikzpicture\}.*?\\end\{tikzpicture\}', src, re.S)
    if len(_raw_blocks) != len(_manifest):
        sys.exit(f'FATAL: {len(_raw_blocks)} tikzpictures in source but manifest has '
                 f'{len(_manifest)}. The source changed after render_tikz.py ran; re-render first.')
    for _i, _b in enumerate(_raw_blocks, 1):
        if _manifest.get(f'tikz{_i:02d}') != hashlib.md5(_b.encode()).hexdigest():
            sys.exit(f'FATAL: tikz{_i:02d} does not match the rendered manifest. '
                     f'The source changed after render_tikz.py ran; re-render first.')
else:
    sys.stderr.write('WARNING: no tikz_manifest.json; diagram order unverified\n')

def png_size(path):
    """Width/height from the PNG IHDR chunk, no image library needed."""
    try:
        d = pathlib.Path(path).read_bytes()[16:24]
        return int.from_bytes(d[:4], 'big'), int.from_bytes(d[4:], 'big')
    except Exception:
        return None, None

MAX_W, MAX_H = 6.0, 6.5     # inches inside a 1in-margin letter page

def tikz_sub(m):
    counter[0] += 1
    f = '%s/tikz%02d.png' % (TIKZ, counter[0])
    w, h = png_size(f)
    width = MAX_W
    if w and h:
        # a tall diagram at full width would run off the page; cap its height
        width = min(MAX_W, MAX_H * w / h)
    return '\\includegraphics[width=%.2fin]{%s}' % (width, f)
body = re.sub(r'\\begin\{tikzpicture\}.*?\\end\{tikzpicture\}', tikz_sub, body, flags=re.S)

# ------------------------------------------- 2. margin notes -> real footnotes
body = rewrite_braced(body, 'marginnote', '\\footnote{', '}', optional=True)
body = rewrite_braced(body, 'sidenote',   '\\footnote{', '}', optional=True)

# ------------------------------------------------------- 3. index / layout
for mac in ('index', 'glsadd', 'pagelayout', 'labch', 'margintoc'):
    body = strip_braced(body, mac)
body = re.sub(r'\\(pagelayout|margintoc|blindtext|listoffigures|listoftables|'
              r'printindex|printglossar\w*|tableofcontents|mainmatter|frontmatter|'
              r'backmatter|appendixpage|addcontentsline\{[^}]*\}\{[^}]*\})\b', '', body)
body = strip_braced(body, 'addcontentsline')

# \resizebox{w}{h}{ $math$ } defeats pandoc's math reader. Drop the wrapper and
# its two size arguments, keeping the maths itself.
def unwrap_resizebox(text):
    out, i = [], 0
    while True:
        j = text.find('\\resizebox', i)
        if j < 0:
            out.append(text[i:]); return ''.join(out)
        out.append(text[i:j])
        k = j + len('\\resizebox')
        for _ in range(3):                       # skip {w}, {h}, then take {body}
            while k < len(text) and text[k] != '{':
                k += 1
            depth, start = 1, k + 1
            k += 1
            while k < len(text) and depth:
                if text[k] == '{': depth += 1
                elif text[k] == '}': depth -= 1
                k += 1
            last = text[start:k-1]
        out.append(last)                          # third group is the content
        i = k
body = unwrap_resizebox(body)
# Unwrapping can leave $...$ nested inside \[...\]; collapse the inner pair.
body = re.sub(r'\\\[\s*\$(.*?)\$\s*\\\]', lambda m: '\\[' + m.group(1) + '\\]', body, flags=re.S)

# ------------------------------------- 4. callout boxes -> title + blockquote
def box_to_quote(text, env, default_title, key=None):
    """\\begin{env}[title] ... \\end{env}  ->  bold title paragraph + quote."""
    out, i = [], 0
    pat = re.compile(r'\\begin\{' + env + r'\}(\[(.*?)\])?', re.S)
    while True:
        m = pat.search(text, i)
        if not m:
            out.append(text[i:]); return ''.join(out)
        k = text.find('\\end{' + env + '}', m.end())
        if k < 0:
            out.append(text[i:]); return ''.join(out)
        title = (m.group(2) or default_title).strip()
        if key and title.startswith(key):
            title = title[len(key):].strip().lstrip('=').strip()
        title = re.sub(r'^frametitle\s*=\s*', '', title)
        inner = text[m.end():k]
        out.append(text[i:m.start()])
        out.append('\n\n\\textbf{%s}\n\n\\begin{quote}\n%s\n\\end{quote}\n\n' % (title, inner))
        i = k + len('\\end{' + env + '}')

body = box_to_quote(body, 'appliedexample', 'Applied Example')
body = box_to_quote(body, 'kaobox',         'Note', key='frametitle')
body = box_to_quote(body, 'vizcallout',     'Visualization to build')
body = box_to_quote(body, 'restson',        'What this depends on')
body = box_to_quote(body, 'vizbox',         'Visualization')
body = box_to_quote(body, 'restsonbox',     'What this depends on')
body = box_to_quote(body, 'mdframed',       'Note')

# --------------------------------------------------------------- 5. parts
body = re.sub(r'\\addpart\{', r'\\part{', body)

# ------------------------------------------- 6. references from the .aux
labs = {}
if AUX:
    a = pathlib.Path(AUX).read_text()
    labs = {mm.group(1): mm.group(2)
            for mm in re.finditer(r'\\newlabel\{([^}]+)\}\{\{([^}]*)\}\{([^}]*)\}', a)}
body = re.sub(r'\\ref\{([^}]*)\}',     lambda m: labs.get(m.group(1), '?'), body)
body = re.sub(r'\\nameref\{([^}]*)\}', lambda m: labs.get(m.group(1), '?'), body)
# Page numbers do not survive a reflow into Word. Turn "the figure on page~X"
# into a figure reference, which stays true whatever the pagination.
body = re.sub(r'the figure on page~?\\pageref\{([^}]*)\}',
              lambda m: 'Figure ' + labs.get(m.group(1), '?'), body)
body = re.sub(r'\s*(on|in)\s+pages?~?\\pageref\{[^}]*\}', '', body)
body = re.sub(r'\\pageref\{[^}]*\}', '', body)

# --------------------------------------- 6b. numbered captions for Word
def number_captions(text, envname, word):
    out, i, tag = [], 0, '\\begin{%s}' % envname
    while True:
        j = text.find(tag, i)
        if j < 0:
            out.append(text[i:]); return ''.join(out)
        k = text.find('\\end{%s}' % envname, j)
        if k < 0:
            out.append(text[i:]); return ''.join(out)
        blk = text[j:k]
        lm = re.search(r'\\label\{([^}]+)\}', blk)
        num = labs.get(lm.group(1)) if lm else None
        if num:
            blk = re.sub(r'\\caption(\[[^\]]*\])?\{',
                         lambda _m, w=word, n=num: '\\caption{%s %s: ' % (w, n), blk, count=1)
        out.append(text[i:j]); out.append(blk); i = k
for env, w in (('figure*', 'Figure'), ('figure', 'Figure'),
               ('table*', 'Table'), ('table', 'Table')):
    body = number_captions(body, env, w)
# strip any surviving short-caption brackets
body = re.sub(r'\\caption\[[^\]]*\]\{', r'\\caption{', body)

# --------------------------------------------- 7. wide floats -> plain floats
body = body.replace('\\begin{figure*}', '\\begin{figure}').replace('\\end{figure*}', '\\end{figure}')
body = body.replace('\\begin{table*}',  '\\begin{table}').replace('\\end{table*}',  '\\end{table}')

# ------------------------- 7b. figures that wrap a code listing, not an image
# pandoc's figure reader expects an image plus a caption. A verbatim block inside
# a figure float is silently DROPPED, which cost 50 lines of Stata code before
# this was caught. Unwrap those floats so the listing survives as ordinary
# content, keeping the caption as an italic line beneath it.
def unwrap_code_figures(text):
    out, i, tag = [], 0, '\\begin{figure}'
    while True:
        j = text.find(tag, i)
        if j < 0:
            out.append(text[i:]); return ''.join(out)
        k = text.find('\\end{figure}', j)
        if k < 0:
            out.append(text[i:]); return ''.join(out)
        blk = text[j:k]
        if '\\begin{verbatim}' in blk:
            body = re.sub(r'^\\begin\{figure\}(\[[^\]]*\])?', '', blk)
            body = re.sub(r'\\caption\{(.*?)\}\s*$',
                          lambda m: '\n\n\\textit{%s}\n' % m.group(1),
                          body, flags=re.S)
            out.append(text[i:j]); out.append('\n\n' + body + '\n\n')
        else:
            out.append(text[i:j]); out.append(blk + '\\end{figure}')
        i = k + len('\\end{figure}')
body = unwrap_code_figures(body)

# ----------------------------------------------------- 8. figure file paths
def figpath(m):
    opt, path = m.group(1) or '', m.group(2)
    if path.startswith('/'):
        return m.group(0)                      # already absolute (rendered tikz)
    path = re.sub(r'^images/', '', path)        # graphicspath already points at images/
    return '\\includegraphics%s{%s/%s}' % (opt, FIGD, path)
body = re.sub(r'\\includegraphics(\[[^\]]*\])?\{([^}]*)\}', figpath, body)

# \input pulls in a fragment (a generated table). Inline it so nothing is lost.
def inline_input(m):
    name = m.group(1)
    if not name.endswith('.tex'):
        name += '.tex'
    for base in (pathlib.Path(sys.argv[1]).parent, pathlib.Path(sys.argv[1]).parent.parent):
        f = base / name
        if f.exists():
            return f.read_text()
    return ''
body = re.sub(r'\\input\{([^}]*)\}', inline_input, body)

# ------------------------------------------------ 9. bibliography placeholder
body = re.sub(r'\\printbibliography(\[[^\]]*\])?',
              lambda _m: '\\chapter*{References}\n\\hypertarget{refs}{}', body)

body = re.sub(r'\\label\{[^}]*\}', '', body)
body = re.sub(r'\n{4,}', '\n\n\n', body)

pathlib.Path(OUT).write_text(body)
print('wrote', OUT, len(body.split('\n')), 'lines;', counter[0], 'tikz images placed')
