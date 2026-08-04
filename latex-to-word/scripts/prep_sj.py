#!/usr/bin/env python3
"""Preprocess a Stata Journal (statapress.cls) manuscript so pandoc can produce
a clean, editable .docx.

Usage: prep_sj.py FLAT.tex OUT.tex [FIGDIR] [AUX]

The SJ classes define many macros pandoc has never heard of. Left alone, pandoc
silently drops them: that is how "\\underbar{sh}eet(tabname)" becomes
"eet(tabname)" and "\\hskip 4em" becomes the literal text "4em". Every macro
handled below was found by inventorying the real source, not guessed at.
"""
import re, sys, pathlib


def replace_braced(text, macro, open_s, close_s):
    """Replace \\macro{...} with open_s + ... + close_s, matching braces properly."""
    out, i, tag = [], 0, '\\' + macro + '{'
    while True:
        j = text.find(tag, i)
        if j < 0:
            out.append(text[i:])
            return ''.join(out)
        out.append(text[i:j])
        k, depth = j + len(tag), 1
        while k < len(text) and depth:
            if text[k] == '{':
                depth += 1
            elif text[k] == '}':
                depth -= 1
                if depth == 0:
                    break
            k += 1
        out.append(open_s + text[j + len(tag):k] + close_s)
        i = k + 1


src = pathlib.Path(sys.argv[1]).read_text()
FIGDIR = sys.argv[3] if len(sys.argv) > 3 else ''
AUX = sys.argv[4] if len(sys.argv) > 4 else ''

# ---------------------------------------------------------------- 1. preamble
src = re.sub(r'\\sjsetissue\{[^}]*\}\{[^}]*\}\{[^}]*\}\{[^}]*\}\s*', '', src)
src = re.sub(r'\\inserttype(\[[^\]]*\])?\{[^}]*\}\s*', '', src)

m = re.search(r'\\author\{[^}]*\}\{(.*?)\n*\}\s*\n', src, re.S)
if m:
    authors = m.group(1).replace('\\and', '\\\\').replace('\\\\', ', ')
    authors = ' '.join(authors.split()).replace(' ,', ',').replace(', ,', ',')
    src = src[:m.start()] + '\\author{' + authors + '}\n' + src[m.end():]
src = re.sub(r'\\title\[[^\]]*\]%?\s*', r'\\title', src, flags=re.S)

# \keywords{...} is an SJ macro pandoc drops silently. Keep it as a real
# paragraph after the abstract, where the journal prints it.
src = re.sub(r'\\inserttag\b', 'gc1', src)
src = re.sub(r'\\keywords\{(.*?)\}\s*',
             lambda mm: '\n\n\\noindent\\textbf{Keywords:} ' + ' '.join(mm.group(1).split()) + '\n\n',
             src, flags=re.S)

# ------------------------------------------- 2. Stata console box-drawing art
# stata.sty draws results boxes from macros. Map them onto the Unicode glyphs
# they imitate, so the box survives as text a reader can still edit.
BOX = {'TLC': '\u250c', 'TRC': '\u2510', 'BLC': '\u2514', 'BRC': '\u2518',
       'LFTT': '\u251c', 'RGTT': '\u2524', 'VBAR': '\u2502'}
for name, glyph in BOX.items():
    src = re.sub(r'\{\\' + name + r'\}|\\' + name + r'\b', glyph, src)
src = re.sub(r'\\HLI\{(\d+)\}', lambda mm: '\u2500' * int(mm.group(1)), src)
src = re.sub(r'\{\\smallskip\}', '', src)        # blank-line marker inside stlog

# ------------------------------------------------------------ 3. environments
src = re.sub(r'\\begin\{stlog\}', r'\\begin{verbatim}', src)
src = re.sub(r'\\end\{stlog\}', r'\\end{verbatim}', src)
# stsyntax holds the syntax diagram; an indented block keeps the italics and the
# minimum-abbreviation underlines that a code block would destroy.
src = re.sub(r'\\begin\{stsyntax\}', r'\\begin{quote}', src)
src = re.sub(r'\\end\{stsyntax\}', r'\\end{quote}', src)
src = re.sub(r'\\begin\{aboutauthors\}', r'\\section*{About the authors}', src)
src = re.sub(r'\\end\{aboutauthors\}', '', src)

# ------------------------------------------------------------- 4. text macros
src = re.sub(r'\{\\it\s+(.*?)\\/?\}', lambda mm: '\\textit{%s}' % mm.group(1), src, flags=re.S)
src = re.sub(r'\{\\it\s+(.*?)\}', lambda mm: '\\textit{%s}' % mm.group(1), src, flags=re.S)
src = re.sub(r'\{\\tt\s+(.*?)\}', lambda mm: '\\texttt{%s}' % mm.group(1), src, flags=re.S)
src = replace_braced(src, 'underbar', '\\underline{', '}')   # min. abbreviation
src = replace_braced(src, 'optional', '[', ']')              # Stata syntax [opt]
src = replace_braced(src, 'stcmd', '\\texttt{', '}')
src = replace_braced(src, 'stresultsgroup', '\\textbf{', '}')
src = re.sub(r'\\tytilde\b', '~', src)
src = re.sub(r'\\hangpara\b', '\n\n', src)                   # one para per option
src = re.sub(r'\\hskip\s*[\d.]+\s*(em|pc|pt|in|cm|ex)', '    ', src)  # syntax-diagram indent
# A line break immediately followed by "[" is read by LaTeX (and pandoc) as
# \\[len], the optional vertical-space argument, which silently EATS the
# bracketed text. Stata syntax diagrams are full of "\\\n  [options]", so guard
# every line break with an empty group.
src = re.sub(r'\\\\(\s*\n\s*)\[', r'\\\\{}\1[', src)
for junk in (r'\\allowbreak', r'\\smallskip', r'\\medskip', r'\\bigskip',
             r'\\clearpage', r'\\newpage', r'\\noindent', r'\\footnotesize',
             r'\\centering', r'\\vsp'):
    src = re.sub(junk + r'\b', '', src)

# ------------------------------------------- 5. cross-references from the .aux
labs = {}
if AUX:
    a = pathlib.Path(AUX).read_text()
    labs = {mm.group(1): mm.group(2)
            for mm in re.finditer(r'\\newlabel\{([^}]+)\}\{\{([^}]*)\}\{([^}]*)\}', a)}
src = re.sub(r'\\ref\{([^}]*)\}', lambda mm: labs.get(mm.group(1), '?'), src)
src = re.sub(r'\\pageref\{[^}]*\}', '', src)

# ----------------------------------------------------------------- 6. figures
if FIGDIR:
    src = re.sub(r'\\includegraphics(\[[^\]]*\])?\{([^}]*)\}',
                 lambda mm: '\\includegraphics%s{%s/%s}' % (mm.group(1) or '', FIGDIR, mm.group(2)),
                 src)

# ------------------------------------- 6b. numbered captions ("Figure 1: ...")
# LaTeX numbers floats automatically; Word does not. Inject the real numbers
# from the .aux so captions read "Figure 1: ..." as they do in the PDF.
def number_captions(text, envname, word):
    out, i = [], 0
    tag = '\\begin{%s}' % envname
    while True:
        j = text.find(tag, i)
        if j < 0:
            out.append(text[i:]); return ''.join(out)
        k = text.find('\\end{%s}' % envname, j)
        block = text[j:k]
        lm = re.search(r'\\label\{([^}]+)\}', block)
        num = labs.get(lm.group(1)) if lm else None
        if num:
            block = re.sub(r'\\caption\{', lambda _m, w=word, n=num: '\\caption{%s %s: ' % (w, n), block, count=1)
        out.append(text[i:j]); out.append(block); i = k

src = number_captions(src, 'figure', 'Figure')
src = number_captions(src, 'table', 'Table')

# ------------------------------------------------- 6c. bibliography placement
# pandoc --citeproc appends the reference list at the very end unless a div with
# id "refs" says otherwise. \hypertarget{refs}{} becomes exactly that div, so the
# references land where LaTeX had \bibliography, ahead of "About the authors".
src = re.sub(r'\\bibliographystyle\{[^}]*\}\s*', '', src)
src = re.sub(r'\\bibliography\{[^}]*\}',
             lambda _m: '\\section*{References}\n\\hypertarget{refs}{}', src)

src = re.sub(r'\\label\{[^}]*\}', '', src)

pathlib.Path(sys.argv[2]).write_text(src)
print('wrote', sys.argv[2], len(src.split('\n')), 'lines')
