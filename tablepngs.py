#!/usr/bin/env python3
"""
tablepngs — flatten every table in a LaTeX document into a high-resolution PNG.

Why: when Adobe Acrobat converts a PDF to Word, it detects text-based tables and
tries to reconstruct them as Word tables, which usually mangles the layout. A
flattened raster image is left alone. tablepngs rewrites your document so each
table is rendered as a crisp PNG (captions and labels stay as live text, so
numbering, \\ref and the List of Tables keep working), producing a companion
document whose PDF survives PDF-to-Word conversion with tables intact.

How it works, in one breath: it compiles your original document once (so the
.aux exists), extracts each table's body, re-typesets that body in a tiny
companion document that reuses YOUR preamble (same class, packages, macros,
fonts, \\textwidth) cropped tight via the `preview` package, rasterizes the
result at high DPI, and splices `\\includegraphics` calls back into a copy of
your document. Multi-page longtables are compiled with your page geometry,
split into one PNG per page, and re-inserted as a stack of page-fitting images
that break cleanly between pages. A verification pass then extracts the text
layer of the final PDF and confirms the table contents are gone from it (and
the captions are still there).

Requires: a TeX distribution (pdflatex / xelatex / lualatex), plus ONE of
pdftoppm (poppler), ImageMagick, or Ghostscript for PDF->PNG. `pdfcrop` (ships
with TeX Live/MiKTeX) is used for longtable page cropping. Run
`tablepngs.py --check` for a full dependency report with install instructions.

Author: Eric Booth (with Claude). License: MIT.
"""

import argparse
import glob as globmod
import json
import os
import platform
import re
import shutil
import struct
import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path

__version__ = "1.0.0"

# --------------------------------------------------------------------------
# Environment families
# --------------------------------------------------------------------------
FLOAT_ENVS = ["table", "table*", "sidewaystable", "sidewaystable*"]
LONG_ENVS = ["longtable", "longtabu", "xltabular", "supertabular", "longtblr"]
BARE_ENVS = ["tabular", "tabular*", "tabularx", "tabulary", "tabu", "NiceTabular",
             "tblr", "talltblr"]

# Table-ish environments tablepngs does NOT flatten, and what to tell the user.
# Being explicit beats silently leaving a table as live text: a silent skip
# looks identical to success right up until Word mangles that one table.
UNSUPPORTED_ENVS = {
    "tabbing": "a \\tabbing block, which is laid out as text rather than as "
               "a table; tablepngs cannot isolate it reliably",
    "deluxetable": "an AASTeX deluxetable, which carries its own caption and "
                   "float machinery",
    "deluxetable*": "an AASTeX deluxetable",
    "supertabular*": "supertabular* (the unstarred supertabular is supported)",
    "xtabular": "xtabular from the xtab package",
    "ctable": "a ctable, which builds its own float",
    "TAB": "a TAB environment from the easytable/tabls family",
}
VERBATIM_ENVS = {
    "verbatim", "verbatim*", "Verbatim", "Verbatim*", "lstlisting",
    "minted", "alltt", "comment", "filecontents", "filecontents*",
}

TABULAR_CONTENT_RE = re.compile(
    r"\\begin\s*\{(?:tabular|tabular\*|tabularx|tabulary|tabu|longtable|longtabu"
    r"|xltabular|supertabular|NiceTabular|threeparttable)\}"
)


# --------------------------------------------------------------------------
# Small utilities
# --------------------------------------------------------------------------
class TPError(Exception):
    pass


def info(msg):
    print(f"[tablepngs] {msg}")


def warn(msg):
    print(f"[tablepngs] WARNING: {msg}", file=sys.stderr)


def die(msg, code=1):
    print(f"[tablepngs] ERROR: {msg}", file=sys.stderr)
    sys.exit(code)


def which(cmd):
    return shutil.which(cmd)


def run(cmd, cwd=None, timeout=300, env=None):
    """Run a command; return (returncode, stdout+stderr as text)."""
    try:
        p = subprocess.run(
            cmd, cwd=cwd, timeout=timeout, env=env,
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        )
        return p.returncode, p.stdout.decode("utf-8", errors="replace")
    except subprocess.TimeoutExpired:
        return 124, f"TIMEOUT after {timeout}s: {' '.join(map(str, cmd))}"
    except FileNotFoundError:
        return 127, f"command not found: {cmd[0]}"


def read_text_guess(path):
    """Read a file as utf-8, falling back to latin-1. Returns (text, encoding)."""
    data = Path(path).read_bytes()
    for enc in ("utf-8", "latin-1"):
        try:
            return data.decode(enc), enc
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace"), "utf-8"


def texpath(p):
    """Path string safe for use inside a .tex file (forward slashes)."""
    return str(p).replace("\\", "/")


# --------------------------------------------------------------------------
# LaTeX-aware text scanning
# --------------------------------------------------------------------------
def mask_comments(text):
    """Return a same-length copy of `text` in which comments, \\verb args and
    verbatim-environment bodies are replaced by spaces (newlines preserved).
    Scanning the masked text with regexes is then safe; offsets map 1:1 back
    to the original."""
    out = list(text)
    i, n = 0, len(text)
    while i < n:
        c = text[i]
        if c == "\\":
            # a control sequence: \% \\ \verb \begin etc.
            j = i + 1
            if j < n and text[j].isalpha():
                while j < n and text[j].isalpha():
                    j += 1
                name = text[i + 1:j]
                if name == "verb":
                    # \verb* possibly
                    if j < n and text[j] == "*":
                        j += 1
                    if j < n:
                        delim = text[j]
                        k = j + 1
                        while k < n and text[k] != delim and text[k] != "\n":
                            out[k] = " "
                            k += 1
                        i = k + 1
                        continue
                elif name == "begin":
                    m = re.match(r"\s*\{([^{}]*)\}", text[j:])
                    if m and m.group(1) in VERBATIM_ENVS:
                        env = m.group(1)
                        body_start = j + m.end()
                        endpat = re.compile(
                            r"\\end\s*\{" + re.escape(env).replace(r"\*", r"\*") + r"\}"
                        )
                        em = endpat.search(text, body_start)
                        stop = em.start() if em else n
                        for k in range(body_start, stop):
                            if text[k] != "\n":
                                out[k] = " "
                        i = em.end() if em else n
                        continue
                i = j
                continue
            else:
                i += 2  # \%  \{  \\ etc: skip escaped char
                continue
        if c == "%":
            k = i
            while k < n and text[k] != "\n":
                out[k] = " "
                k += 1
            i = k
            continue
        i += 1
    return "".join(out)


def match_brace(text, i):
    """text[i] must be '{'. Return index just past the matching '}'.
    Understands escaped braces."""
    assert text[i] == "{"
    depth = 0
    n = len(text)
    while i < n:
        c = text[i]
        if c == "\\":
            i += 2
            continue
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0:
                return i + 1
        i += 1
    raise TPError("unbalanced braces")


def match_bracket(text, i):
    """text[i] must be '['. Return index just past the matching ']'."""
    assert text[i] == "["
    depth = 0
    n = len(text)
    while i < n:
        c = text[i]
        if c == "\\":
            i += 2
            continue
        if c == "[":
            depth += 1
        elif c == "]":
            depth -= 1
            if depth == 0:
                return i + 1
        elif c == "{":
            i = match_brace(text, i)
            continue
        i += 1
    raise TPError("unbalanced brackets")


INPUT_RE = re.compile(r"\\(input|include)\s*\{([^{}]+)\}")


def inline_inputs(text, basedir, max_depth=8):
    """Recursively replace \\input{file} / \\include{file} with the file's
    contents (comment-aware), so tables living in \\input-ed files are
    scanned and flattened too. Missing files and piped input are left
    untouched. Returns (new_text, files_inlined)."""
    total = 0
    for _ in range(max_depth):
        masked = mask_comments(text)
        out, pos, changed = [], 0, 0
        for m in INPUT_RE.finditer(masked):
            name = m.group(2).strip()
            if name.startswith("|") or "\\" in name:
                continue
            p = Path(basedir) / name
            if p.suffix == "":
                p = p.with_suffix(".tex")
            if not p.is_file():
                continue
            content, _ = read_text_guess(p)
            # honor \endinput: drop everything after it
            em = re.search(r"\\endinput(?![a-zA-Z])", mask_comments(content))
            if em:
                content = content[:em.start()]
            content = f"% tablepngs: inlined {name}\n{content}\n"
            if m.group(1) == "include":
                content = f"\\clearpage\n{content}\\clearpage\n"
            out.append(text[pos:m.start()])
            out.append(content)
            pos = m.end()
            changed += 1
        out.append(text[pos:])
        text = "".join(out)
        total += changed
        if not changed:
            break
    return text, total


@dataclass
class EnvSpan:
    name: str
    start: int          # offset of the backslash of \begin
    end: int            # offset just past \end{name}
    body_start: int     # just past \begin{name} and any [..] option
    body_end: int       # offset of the backslash of \end
    opt: str = ""       # the raw [...] option incl. brackets, if any


def find_environments(masked, names):
    """Find all outer-level spans of the named environments in masked text.
    Nesting of the *same* environment name is handled with a counter."""
    nameset = set(names)
    tok_re = re.compile(r"\\(begin|end)\s*\{([^{}]*)\}")
    spans = []
    stack = {}  # name -> list of (start, after_begin)
    for m in tok_re.finditer(masked):
        kind, name = m.group(1), m.group(2)
        if name not in nameset:
            continue
        if kind == "begin":
            stack.setdefault(name, []).append((m.start(), m.end()))
        else:
            if stack.get(name):
                start, after_begin = stack[name].pop()
                if not stack[name]:  # outermost of this name
                    # parse optional [..] after \begin{name}
                    body_start = after_begin
                    opt = ""
                    k = body_start
                    while k < len(masked) and masked[k] in " \t\n":
                        k += 1
                    if k < len(masked) and masked[k] == "[":
                        k2 = match_bracket(masked, k)
                        opt = masked[k:k2]
                        body_start = k2
                    spans.append(EnvSpan(name, start, m.end(), body_start, m.start(), opt))
    spans.sort(key=lambda s: s.start)
    return spans


def find_command_spans(masked, command, star_ok=True, opt_ok=True):
    """Find spans of \\command[opt]{arg} in masked text.
    Returns list of dicts with full span, star flag, opt span, arg span."""
    out = []
    pat = re.compile(r"\\" + command + r"(?![a-zA-Z])")
    for m in pat.finditer(masked):
        i = m.end()
        star = False
        if star_ok and i < len(masked) and masked[i] == "*":
            star = True
            i += 1
        while i < len(masked) and masked[i] in " \t\n":
            i += 1
        opt_span = None
        if opt_ok and i < len(masked) and masked[i] == "[":
            j = match_bracket(masked, i)
            opt_span = (i, j)
            i = j
            while i < len(masked) and masked[i] in " \t\n":
                i += 1
        if i < len(masked) and masked[i] == "{":
            j = match_brace(masked, i)
            out.append({
                "start": m.start(), "end": j, "star": star,
                "opt": opt_span, "arg": (i + 1, j - 1),
            })
    return out


# --------------------------------------------------------------------------
# Table target model
# --------------------------------------------------------------------------
STYLE_BOUNDARY_RE = re.compile(
    r"\n[ \t]*\n"                       # paragraph break
    r"|\\begingroup(?![a-zA-Z])|\\bgroup(?![a-zA-Z])"
    r"|\\begin\s*\{[^{}]*\}|\\end\s*\{[^{}]*\}"
    r"|\\(?:par|clearpage|newpage|pagebreak|noindent)(?![a-zA-Z])"
    r"|\\(?:section|subsection|subsubsection|paragraph|chapter)\*?\s*"
    r"(?:\[[^\]]*\])?\s*\{")


# Declarations that legitimately style a table from outside it. Commands with
# arguments must be listed explicitly, because carrying a command without its
# arguments would break the snippet.
STYLE_WITH_ARGS = {
    "setlength": 2, "addtolength": 2, "renewcommand": 2, "def": 0,
    "definecolor": 3, "captionsetup": 1, "rowcolors": 3, "arrayrulecolor": 1,
    "setstretch": 1, "fontsize": 2, "selectfont": 0,
}
# Never carry these across: they are structure or content, not styling.
STYLE_BLACKLIST = {
    "begin", "end", "item", "caption", "captionof", "label", "ref", "cite",
    "par", "newpage", "clearpage", "pagebreak", "footnote", "footnotetext",
    "includegraphics", "input", "include", "hline", "toprule", "midrule",
    "bottomrule", "cmidrule", "multicolumn", "multirow", "section",
    "subsection", "subsubsection", "paragraph", "chapter", "textbf", "textit",
    "emph", "text", "url", "href", "noindent", "indent",
}


def extract_style_declarations(chunk):
    """Pull the styling declarations out of a run of LaTeX source.

    Handles the two shapes that actually occur before a table:
        \\begingroup\\footnotesize\\setlength{\\tabcolsep}{3pt}
        {\\tabfigures\\footnotesize
    The second leaves an unbalanced brace, so we keep the commands and drop
    the braces -- in a snippet the group scope is the whole document anyway.
    An unrecognised no-argument macro is kept (it is typically a preamble
    switch such as \\tabfigures, and the snippet reuses that preamble); an
    unrecognised macro that takes arguments is dropped, since keeping the
    name without its arguments would break the compile."""
    masked = mask_comments(chunk)
    out = []
    consumed = 0        # end of the last command WITH its arguments
    for m in re.finditer(r"\\([a-zA-Z@]+)\*?", masked):
        if m.start() < consumed:
            continue    # inside an argument we already took verbatim
        name = m.group(1)
        if name in STYLE_BLACKLIST:
            continue
        i = m.end()
        nargs = STYLE_WITH_ARGS.get(name)
        if nargs:
            j = i
            try:
                for _ in range(nargs):
                    while j < len(masked) and masked[j] in " \t\n":
                        j += 1
                    if j < len(masked) and masked[j] == "{":
                        j = match_brace(masked, j)
                    elif j < len(masked) and masked[j] == "\\":
                        k = re.match(r"\\[a-zA-Z@]+", masked[j:])
                        j += k.end() if k else 1
                    else:
                        raise TPError("bad arg")
            except (TPError, AssertionError):
                continue
            out.append(chunk[m.start():j])
            consumed = j
        else:
            # keep only if not immediately followed by a brace argument
            j = i
            while j < len(masked) and masked[j] in " \t\n":
                j += 1
            if j < len(masked) and masked[j] == "{" and nargs is None:
                continue
            out.append(chunk[m.start():m.end()])
    return "".join(out)


def leading_style_context(text, masked, start, max_chars=600):
    """Formatting that applies to a table but is written OUTSIDE its
    environment, e.g.

        \\begingroup\\footnotesize\\setlength{\\tabcolsep}{3pt}
        \\begin{longtable}{...}

    Without this the snippet would typeset the table at the wrong size, which
    changes column widths and (because the image then exceeds \\linewidth)
    silently shrinks the table in the output. Only a run consisting purely of
    commands is accepted -- if there is any prose in the gap, we take nothing.
    """
    window_start = max(0, start - max_chars)
    boundary = window_start
    for m in STYLE_BOUNDARY_RE.finditer(masked, window_start, start):
        boundary = m.end()
    # A closing brace before the table ends any group that could have styled
    # it, so nothing before that point applies.
    close = masked.rfind("}", boundary, start)
    if close != -1:
        # ...unless it closes a group opened after the boundary (e.g. the
        # argument of \setlength). Only trust it if braces are unbalanced.
        depth = 0
        for m in re.finditer(r"\\.|\{|\}", masked[boundary:start], re.S):
            if m.group() == "{":
                depth += 1
            elif m.group() == "}":
                depth -= 1
                if depth < 0:
                    boundary = boundary + m.end()
                    depth = 0
    return extract_style_declarations(text[boundary:start])


@dataclass
class Target:
    index: int                 # 1-based, document order
    kind: str                  # 'float' | 'long' | 'bare'
    env: str
    span: EnvSpan              # in the main document text
    line: int
    caption_raw: str = ""      # full \caption[..]{..} source, "" if none
    caption_text: str = ""     # caption argument (raw latex)
    caption_pos: str = "above"  # 'above' | 'below'
    labels_raw: list = field(default_factory=list)
    render_content: str = ""   # latex to typeset for the image
    in_landscape: bool = False
    multi_caption: bool = False
    n_captions: int = 0
    caps_list: list = field(default_factory=list)  # all captions (multi-caption bake)
    images: list = field(default_factory=list)   # relative png paths
    snippet_pdf: str = ""      # the PDF the PNGs were rasterized from
    widths_pt: list = field(default_factory=list)
    probes: list = field(default_factory=list)
    method: str = ""           # 'preview' | 'pagecrop'
    ok: bool = False
    note: str = ""
    style_prefix: str = ""     # formatting applied from outside the environment
    caps_before: int = 0       # table numbers consumed earlier (this section)
    counter_seed: dict = field(default_factory=dict)  # counter -> value


def line_of(text, off):
    return text.count("\n", 0, off) + 1


def strip_caption_and_labels(body_raw, body_masked, for_longtable=False):
    """Remove \\caption... and \\label{...} from a body. Returns
    (content, captions, labels) where captions is a list of dicts
    (raw, text, offset) and labels a list of raw \\label commands.
    For longtables, a row-terminating \\\\ that immediately follows a
    removed caption (+labels) is removed with it."""
    captions = find_command_spans(body_masked, "caption")
    labels = find_command_spans(body_masked, "label", star_ok=False, opt_ok=False)
    # a \label nested INSIDE a caption argument (the \caption{...\label{x}}
    # idiom) travels with the caption's raw text; never cut it separately
    labels = [l for l in labels
              if not any(c["start"] <= l["start"] and l["end"] <= c["end"]
                         for c in captions)]
    cuts = []  # (start, end)
    cap_out, lab_out = [], []
    label_spans = {(l["start"], l["end"]) for l in labels}
    for c in captions:
        s, e = c["start"], c["end"]
        own_labels = []
        # swallow trailing whitespace + labels + (for longtable) one row-end \\
        j = e
        n = len(body_masked)
        while True:
            k = j
            while k < n and body_masked[k] in " \t\n":
                k += 1
            took_label = False
            for l in labels:
                if l["start"] == k:
                    j = l["end"]
                    label_spans.discard((l["start"], l["end"]))
                    own_labels.append(body_raw[l["start"]:l["end"]])
                    took_label = True
                    break
            if not took_label:
                j = k
                break
        if for_longtable and j + 1 < len(body_masked) and body_masked[j] == "\\" and body_masked[j + 1] == "\\":
            j += 2
            m = re.match(r"\s*\[[^\]]*\]", body_masked[j:])
            if m:
                j += m.end()
        cuts.append((s, j))
        cap_out.append({
            "raw": body_raw[s:e],
            "text": body_raw[c["arg"][0]:c["arg"][1]],
            "offset": s,
            "star": c["star"],
            "labels": own_labels,
        })
    for (s, e) in sorted(label_spans):
        cuts.append((s, e))
        lab_out.append(body_raw[s:e])
    # splice out the cuts
    cuts.sort()
    content, pos = [], 0
    for (s, e) in cuts:
        content.append(body_raw[pos:s])
        pos = e
    content.append(body_raw[pos:])
    return "".join(content), cap_out, lab_out


def longtable_caption_position(body_masked, cap_offset):
    """A longtable caption placed in the \\endfoot / \\endlastfoot block is
    rendered BELOW the table; anything earlier renders above."""
    term_re = re.compile(r"\\end(firsthead|head|foot|lastfoot)(?![a-zA-Z])")
    for m in term_re.finditer(body_masked):
        if m.start() > cap_offset:
            return "below" if m.group(1) in ("foot", "lastfoot") else "above"
    return "above"


def scan_targets(text, masked):
    """Locate every table-like target in document order."""
    doc_m = re.search(r"\\begin\s*\{document\}", masked)
    if not doc_m:
        raise TPError(r"no \begin{document} found — is this a full LaTeX document?")
    body_off = doc_m.end()

    floats = [s for s in find_environments(masked, FLOAT_ENVS) if s.start > body_off]
    longs = [s for s in find_environments(masked, LONG_ENVS) if s.start > body_off]
    landscapes = [s for s in find_environments(masked, ["landscape"]) if s.start > body_off]
    bares_all = [s for s in find_environments(masked, BARE_ENVS) if s.start > body_off]
    bares = [
        s for s in bares_all
        if not any(f.start < s.start and s.end <= f.end for f in floats)
        and not any(l.start < s.start and s.end <= l.end for l in longs)
        and not any(b is not s and b.start < s.start and s.end <= b.end for b in bares_all)
    ]

    # ltablex turns tabularx into a page-breaking environment
    preamble_masked = masked[:doc_m.start()]
    if re.search(r"\\usepackage(\[[^\]]*\])?\s*\{[^{}]*ltablex[^{}]*\}", preamble_masked):
        promote = [b for b in bares if b.name == "tabularx"]
        bares = [b for b in bares if b.name != "tabularx"]
        longs += promote
        longs.sort(key=lambda s: s.start)

    def in_landscape(s):
        return any(L.start < s.start and s.end <= L.end for L in landscapes)

    targets = []
    for s in floats:
        targets.append(("float", s))
    for s in longs:
        targets.append(("long", s))
    for s in bares:
        targets.append(("bare", s))
    targets.sort(key=lambda t: t[1].start)

    # Table-like environments we knowingly cannot handle. Report them by name
    # and line so the user is never left assuming a table was flattened when
    # it was quietly left as live text.
    unsupported = []
    for m in re.finditer(r"\\begin\s*\{([A-Za-z*]+)\}", masked[body_off:]):
        env = m.group(1)
        if env in UNSUPPORTED_ENVS:
            unsupported.append((env, line_of(text, body_off + m.start())))

    out, notes = [], []
    idx = 0
    for kind, s in targets:
        body_raw = text[s.body_start:s.body_end]
        body_masked = masked[s.body_start:s.body_end]
        if kind == "float" and not (
                TABULAR_CONTENT_RE.search(body_masked)
                or re.search(r"\\(input|include)\s*\{|\\halign", body_masked)):
            notes.append(f"skipping {s.name} float at line {line_of(text, s.start)}: "
                         "no tabular content (image or text-only float)")
            continue
        idx += 1
        t = Target(index=idx, kind=kind, env=s.name, span=s,
                   line=line_of(text, s.start), in_landscape=in_landscape(s))
        if kind == "float":
            content, caps, labs = strip_caption_and_labels(body_raw, body_masked)
            t.n_captions = len(caps)
            if len(caps) > 1:
                # unusual: several captioned tables inside one float.
                # Bake everything (captions included) into the image; the
                # rewrite re-steps the counter and re-plants the labels.
                t.multi_caption = True
                t.render_content = body_raw
                t.caps_list = caps
                t.labels_raw = []
                t.note = f"{len(caps)} captions in one float: captions baked into image"
            else:
                t.render_content = content
                t.labels_raw = (caps[0]["labels"] if caps else []) + labs
                if caps:
                    t.caption_raw = caps[0]["raw"]
                    t.caption_text = caps[0]["text"]
                    core = re.sub(r"\\centering|\\small|\\footnotesize|\\scriptsize|\\normalsize",
                                  "", body_masked[:caps[0]["offset"]])
                    m = TABULAR_CONTENT_RE.search(body_masked)
                    first_material = m.start() if m else len(body_masked)
                    t.caption_pos = "above" if caps[0]["offset"] < first_material else "below"
        elif kind == "long":
            # keep the full environment (incl. colspec) minus captions/labels
            env_raw = text[s.start:s.end]
            env_masked = masked[s.start:s.end]
            content, caps, labs = strip_caption_and_labels(env_raw, env_masked, for_longtable=True)
            t.render_content = content
            t.labels_raw = (caps[0]["labels"] if caps else []) + labs
            t.n_captions = len(caps)
            if caps:
                t.caption_raw = caps[0]["raw"]
                t.caption_text = caps[0]["text"]
                # keep only the first "real" caption; \caption[]{...} continued
                # headers were stripped along with it, which is what we want
                t.caption_pos = longtable_caption_position(env_masked, caps[0]["offset"])
        else:  # bare
            t.render_content = text[s.start:s.end]
        t.style_prefix = leading_style_context(text, masked, s.start)
        if t.style_prefix:
            t.render_content = t.style_prefix + "\n" + t.render_content
        out.append(t)
    # ---- table-number bookkeeping -----------------------------------------
    # The table counter is reset by a parent counter: \chapter in book/report,
    # or whatever \counterwithin{table}{...} names (common in reports that
    # number tables 2.1, 2.2, ...). Seeding the snippet with the right parent
    # value is what makes \thetable inside a continued header correct.
    parent = None
    cw = re.search(r"\\(?:counterwithin|numberwithin)\*?\s*\{table\}\s*\{([a-zA-Z]+)\}",
                   masked)
    if cw:
        parent = cw.group(1)
    elif re.search(r"\\chapter(?!\*)\s*[\[{]", masked):
        parent = "chapter"

    # positions of each counter-stepping command, outermost first
    chain = []          # [(counter_name, [positions])]
    if parent == "chapter":
        chain = [("chapter", None)]
    elif parent:
        if re.search(r"\\chapter(?!\*)\s*[\[{]", masked):
            chain = [("chapter", None), (parent, None)]
        else:
            chain = [(parent, None)]
    chain = [(name, [m.start() for m in
                     re.finditer(r"\\" + name + r"(?!\*)(?![a-zA-Z])\s*(?:\[|\{)", masked)])
             for name, _ in chain]

    seg, caps = None, 0
    for t in out:
        seed = {}
        for name, marks in chain:
            seed[name] = sum(1 for c in marks if c < t.span.start)
        # the table counter restarts whenever any parent in the chain advances
        key = tuple(seed.get(name, 0) for name, _ in chain)
        if chain and key != seg:
            seg, caps = key, 0
        t.counter_seed = seed
        t.caps_before = caps
        caps += max(t.n_captions, 1) if t.kind == "long" else t.n_captions
    return out, notes, unsupported


# --------------------------------------------------------------------------
# Engines & external tools
# --------------------------------------------------------------------------
STYLE_LOAD_RE = re.compile(
    r"\\(?:usepackage|RequirePackage)\s*(?:\[[^\]]*\])?\s*\{([^{}]*)\}"
    r"|\\documentclass\s*(?:\[[^\]]*\])?\s*\{([^{}]*)\}")


def collect_local_styles(preamble, basedir, depth=4, _seen=None):
    """Text of every LOCAL .sty/.cls file the preamble pulls in, following
    nested \\RequirePackage chains. Engine-defining packages such as fontspec
    are routinely buried inside a house style file rather than named in the
    document, so detection has to look there too."""
    if _seen is None:
        _seen = set()
    if depth <= 0:
        return ""
    chunks = []
    for m in STYLE_LOAD_RE.finditer(mask_comments(preamble)):
        names = (m.group(1) or m.group(2) or "")
        for name in names.split(","):
            name = name.strip()
            if not name:
                continue
            for ext in (".sty", ".cls", ""):
                p = (Path(basedir) / (name + ext))
                if p.is_file() and p.suffix in (".sty", ".cls"):
                    rp = str(p.resolve())
                    if rp in _seen:
                        break
                    _seen.add(rp)
                    body, _ = read_text_guess(p)
                    chunks.append(body)
                    chunks.append(collect_local_styles(body, p.parent,
                                                       depth - 1, _seen))
                    break
    return "\n".join(chunks)


def engine_from_log(main_tex):
    """How this document was last actually built. A .log or .xdv left by the
    user's own toolchain is the most reliable signal there is."""
    log = main_tex.with_suffix(".log")
    if log.is_file():
        try:
            head = log.read_text(errors="replace")[:4000]
        except OSError:
            head = ""
        if re.search(r"This is XeTeX", head):
            return "xelatex"
        if re.search(r"This is LuaHBTeX|This is LuaTeX", head):
            return "lualatex"
        if re.search(r"This is pdfTeX", head):
            return "pdflatex"
    if main_tex.with_suffix(".xdv").is_file():
        return "xelatex"
    return None


def detect_engine(text, masked, basedir=None, main_tex=None):
    m = re.search(r"^%\s*!\s*TEX\s+(?:TS-)?program\s*=\s*(\S+)", text,
                  re.IGNORECASE | re.MULTILINE)
    if m:
        prog = m.group(1).lower()
        for e in ("xelatex", "lualatex", "pdflatex"):
            if e in prog:
                return e, "magic comment"
    if main_tex is not None:
        eng = engine_from_log(main_tex)
        if eng:
            return eng, "previous build artifacts"
    doc = re.search(r"\\begin\s*\{document\}", masked)
    pre = masked[:doc.start()] if doc else masked
    hay = pre
    if basedir is not None:
        hay = pre + "\n" + collect_local_styles(pre, basedir)
    if re.search(r"\\(?:usepackage|RequirePackage)(\[[^\]]*\])?\s*\{[^{}]*"
                 r"(luacode|luatexja|luaotfload)[^{}]*\}|\\directlua", hay):
        return "lualatex", "lua-only packages"
    if re.search(r"\\(?:usepackage|RequirePackage)(\[[^\]]*\])?\s*\{[^{}]*"
                 r"(fontspec|polyglossia|unicode-math|mathspec|xeCJK)[^{}]*\}", hay):
        return "xelatex", "fontspec/unicode packages"
    return "pdflatex", "default"


@dataclass
class Tools:
    engine: str = ""
    raster: str = ""        # pdftoppm | magick | gs
    raster_path: str = ""
    pdfcrop: str = ""
    gs: str = ""
    text_extract: str = ""  # pdftotext | gs
    pdfinfo: str = ""


def discover_tools(engine):
    t = Tools()
    t.engine = which(engine) or ""
    t.gs = which("gs") or which("gswin64c") or which("gswin32c") or ""
    for name in ("pdftoppm", "magick", "gs"):
        p = which(name) or (t.gs if name == "gs" else "")
        if p:
            t.raster, t.raster_path = name, p
            break
    t.pdfcrop = which("pdfcrop") or ""
    if which("pdftotext"):
        t.text_extract = "pdftotext"
    elif t.gs:
        t.text_extract = "gs"
    t.pdfinfo = which("pdfinfo") or ""
    return t


INSTALL_HINTS = {
    "Darwin": {
        "tex": "install MacTeX:  brew install --cask mactex-no-gui   (or https://tug.org/mactex/)",
        "pdftoppm": "brew install poppler",
        "pdftotext": "brew install poppler",
        "magick": "brew install imagemagick",
        "gs": "brew install ghostscript",
        "pdfcrop": "ships with MacTeX; if missing:  sudo tlmgr install pdfcrop",
    },
    "Linux": {
        "tex": "sudo apt install texlive-latex-extra texlive-extra-utils   (Debian/Ubuntu)",
        "pdftoppm": "sudo apt install poppler-utils",
        "pdftotext": "sudo apt install poppler-utils",
        "magick": "sudo apt install imagemagick   (note: PDF may be blocked by /etc/ImageMagick-*/policy.xml — prefer poppler)",
        "gs": "sudo apt install ghostscript",
        "pdfcrop": "sudo apt install texlive-extra-utils",
    },
    "Windows": {
        "tex": "install MiKTeX (https://miktex.org) or TeX Live; ensure it is on PATH",
        "pdftoppm": "choco install poppler   (or: scoop install poppler)",
        "pdftotext": "choco install poppler",
        "magick": "choco install imagemagick",
        "gs": "choco install ghostscript",
        "pdfcrop": "MiKTeX installs pdfcrop on first use; TeX Live:  tlmgr install pdfcrop",
    },
}


def doctor():
    osname = platform.system()
    hints = INSTALL_HINTS.get(osname, INSTALL_HINTS["Linux"])
    print(f"tablepngs {__version__} dependency check  (platform: {osname})\n")
    ok = True

    def row(name, path, required, hint_key):
        nonlocal ok
        status = path or "MISSING"
        mark = "ok " if path else ("REQ" if required else "opt")
        print(f"  [{mark}] {name:<12} {status}")
        if not path:
            print(f"         -> {hints.get(hint_key, hints['tex'])}")
            if required:
                ok = False

    print("LaTeX engines (need at least the one your document uses):")
    engines = {e: which(e) for e in ("pdflatex", "xelatex", "lualatex")}
    for e, p in engines.items():
        print(f"  [{'ok ' if p else '-- '}] {e:<12} {p or 'not found'}")
    if not any(engines.values()):
        print(f"         -> {hints['tex']}")
        ok = False

    print("\nPDF -> PNG rasterizer (need ONE; checked in order of preference):")
    row("pdftoppm", which("pdftoppm"), False, "pdftoppm")
    row("magick", which("magick"), False, "magick")
    gs = which("gs") or which("gswin64c")
    row("gs", gs, False, "gs")
    if not (which("pdftoppm") or which("magick") or gs):
        print("  !! no rasterizer found — install one of the above")
        ok = False

    print("\nLongtable page cropping (recommended for multi-page tables):")
    row("pdfcrop", which("pdfcrop"), False, "pdfcrop")
    print("\nVerification text extraction (need one; gs also works):")
    row("pdftotext", which("pdftotext"), False, "pdftotext")

    print(f"\n{'All good — ready to run.' if ok else 'Missing required tools — see hints above.'}")
    return 0 if ok else 1


# --------------------------------------------------------------------------
# LaTeX compilation
# --------------------------------------------------------------------------
def run_latex(engine, texfile, outdir, cwd, passes=1, shell_escape=False, timeout=300):
    """Compile texfile (absolute path) with output into outdir, cwd set to the
    main document's directory so relative \\input/graphics paths resolve.
    Returns (ok, log_excerpt, pdf_path)."""
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    cmd = [engine, "-interaction=nonstopmode", "-halt-on-error",
           "-file-line-error", f"-output-directory={outdir}"]
    if shell_escape:
        cmd.append("-shell-escape")
    cmd.append(str(texfile))
    log = ""
    for _ in range(passes):
        rc, log = run(cmd, cwd=cwd, timeout=timeout)
        if rc != 0:
            break
    pdf = outdir / (Path(texfile).stem + ".pdf")
    ok = rc == 0 and pdf.exists()
    if not ok:
        log = latex_error_excerpt(log)
    return ok, log, pdf


def latex_error_excerpt(log, context=18):
    """The part of a TeX log a human actually needs. TeX reports errors as a
    line starting with '!' (or 'file:line: message' with -file-line-error);
    everything before that is noise, and the tail of the log is usually the
    unrelated file list."""
    lines = log.splitlines()
    hits = [i for i, l in enumerate(lines)
            if l.startswith("! ") or l.startswith("!pdfTeX")
            or re.match(r"^[^\s:]+\.(?:tex|sty|cls|def|ltx):\d+:\s", l)]
    if hits:
        start = max(0, hits[0] - 2)
        chunk = lines[start:hits[0] + context]
    else:
        # no recognizable error: show the tail, minus the file-list noise
        useful = [l for l in lines if l.strip() and not l.startswith("(")]
        chunk = useful[-context:] if useful else lines[-context:]
    out = "\n".join(chunk).strip()
    if len(hits) > 1:
        out += f"\n... and {len(hits) - 1} more LaTeX error(s) in the log"
    return out


def pdf_page_count(pdf, tools):
    if tools.pdfinfo:
        rc, out = run([tools.pdfinfo, str(pdf)])
        m = re.search(r"^Pages:\s+(\d+)", out, re.M)
        if rc == 0 and m:
            return int(m.group(1))
    try:
        data = Path(pdf).read_bytes()
        counts = re.findall(rb"/Type\s*/Page[^s]", data)
        if counts:
            return len(counts)
    except OSError:
        pass
    return None


# --------------------------------------------------------------------------
# Rasterization
# --------------------------------------------------------------------------
def png_dims(path):
    with open(path, "rb") as f:
        head = f.read(24)
    if len(head) < 24 or head[:8] != b"\x89PNG\r\n\x1a\n":
        raise TPError(f"not a PNG: {path}")
    w, h = struct.unpack(">II", head[16:24])
    return w, h


def pdf_to_pngs(pdf, out_dir, stem, dpi, tools):
    """Rasterize every page of pdf to out_dir/<stem>-pNN.png. Returns list of paths."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    tmp_prefix = out_dir / f"__tp_{stem}"
    for old in out_dir.glob(f"__tp_{stem}-*.png"):
        old.unlink()
    if tools.raster == "pdftoppm":
        rc, out = run([tools.raster_path, "-png", "-r", str(dpi),
                       str(pdf), str(tmp_prefix)])
    elif tools.raster == "magick":
        rc, out = run([tools.raster_path, "-density", str(dpi), str(pdf),
                       "-background", "white", "-alpha", "remove",
                       "-colorspace", "sRGB", str(tmp_prefix) + "-%03d.png"])
    else:  # gs
        rc, out = run([tools.raster_path, "-dSAFER", "-dBATCH", "-dNOPAUSE",
                       "-sDEVICE=png16m", f"-r{dpi}",
                       "-dTextAlphaBits=4", "-dGraphicsAlphaBits=4",
                       "-o", str(tmp_prefix) + "-%03d.png", str(pdf)])
    if rc != 0:
        raise TPError(f"{tools.raster} failed on {pdf}:\n{out[-800:]}")
    produced = sorted(
        out_dir.glob(f"__tp_{stem}-*.png"),
        key=lambda p: int(re.search(r"-(\d+)\.png$", p.name).group(1)),
    )
    if not produced:
        raise TPError(f"{tools.raster} produced no PNGs for {pdf}")
    final = []
    for i, p in enumerate(produced, 1):
        dest = out_dir / f"{stem}-p{i:02d}.png"
        if dest.exists():
            dest.unlink()
        p.rename(dest)
        final.append(dest)
    return final


def crop_pdf(pdf, tools, margin_pt=3):
    """Crop each page of pdf to its ink bounding box. Returns path of cropped
    pdf (may equal input if no cropper available)."""
    pdf = Path(pdf)
    if tools.pdfcrop:
        out = pdf.with_name(pdf.stem + "-crop.pdf")
        rc, log = run([tools.pdfcrop, "--margins", str(margin_pt), str(pdf), str(out)],
                      cwd=pdf.parent)
        if rc == 0 and out.exists():
            return out
        warn(f"pdfcrop failed ({log.splitlines()[-1] if log.splitlines() else rc}); using uncropped pages")
    else:
        warn("pdfcrop not found; longtable page images keep full-page margins "
             "(install pdfcrop: tlmgr install pdfcrop)")
    return pdf


# --------------------------------------------------------------------------
# Snippet documents
# --------------------------------------------------------------------------
def _aux_hook(aux_name):
    # \makeatletter is needed twice: outside the hook so \@input itself
    # tokenizes, and inside so the aux file's \@writefile etc. tokenize
    # when the hook runs at \begin{document}.
    return (f"\\makeatletter\\AtBeginDocument{{\\makeatletter\\@input{{{aux_name}}}\\makeatother}}\\makeatother\n"
            if aux_name else "")


def _counter_lines(table_counter, parent_seed):
    """Restore the counters \\thetable depends on, so a number printed inside
    the table (a continued header, say) matches the real document."""
    lines = ""
    for name, value in (parent_seed or {}).items():
        # guard: the counter may not exist in this class
        lines += (f"\\makeatletter\\@ifundefined{{c@{name}}}{{}}"
                  f"{{\\setcounter{{{name}}}{{{value}}}}}\\makeatother%\n")
    lines += f"\\setcounter{{table}}{{{table_counter}}}%\n"
    return lines


def build_preview_snippet(preamble, content, aux_name, table_counter=0,
                          parent_seed=None, wrap_captions=False):
    aux = _aux_hook(aux_name)
    if wrap_captions:
        # multi-caption bake: \caption cannot appear outside a float, so
        # route it through \captionof inside a full-width minipage
        content = (f"\\begin{{minipage}}{{\\textwidth}}\n"
                   f"\\makeatletter\\def\\@captype{{table}}\\makeatother%\n"
                   f"{content}\n\\end{{minipage}}")
    return (
        f"{preamble}"
        f"\\usepackage[active,tightpage]{{preview}}\n"
        f"\\setlength\\PreviewBorder{{1.5pt}}\n"
        f"\\makeatletter\\@ifundefined{{captionof}}{{\\usepackage{{capt-of}}}}{{}}\\makeatother\n"
        f"{aux}"
        f"\\pagestyle{{empty}}\n"
        f"\\begin{{document}}\n"
        f"{_counter_lines(table_counter, parent_seed)}"
        f"\\begin{{preview}}%\n"
        f"{content}%\n"
        f"\\end{{preview}}\n"
        f"\\end{{document}}\n"
    )


def build_page_snippet(preamble, content, aux_name, landscape=False,
                       table_counter=0, parent_seed=None, wrap_captions=False):
    aux = _aux_hook(aux_name)
    body = content
    if wrap_captions:
        body = (f"\\begin{{minipage}}{{\\textwidth}}\n"
                f"\\makeatletter\\def\\@captype{{table}}\\makeatother%\n"
                f"{body}\n\\end{{minipage}}")
    if landscape:
        body = f"\\begin{{landscape}}\n{body}\n\\end{{landscape}}"
    return (
        f"{preamble}"
        f"{aux}"
        f"\\pagestyle{{empty}}\n"
        f"\\begin{{document}}\n"
        f"\\thispagestyle{{empty}}\n"
        f"{_counter_lines(table_counter, parent_seed)}"
        f"{body}\n"
        f"\\end{{document}}\n"
    )


INJECT_BLOCK = r"""
%% ---- injected by tablepngs v{version} (do not edit) ----------------------
\makeatletter
\@ifpackageloaded{{graphicx}}{{}}{{\usepackage{{graphicx}}}}
\@ifundefined{{captionof}}{{\usepackage{{capt-of}}}}{{}}
% multi-page tables are re-emitted as a longtable of page images
\@ifpackageloaded{{longtable}}{{}}{{\usepackage{{longtable}}}}
\@ifundefined{{tablepngs@nat}}{{%
  \newlength{{\tablepngs@nat}}\newlength{{\tablepngs@cap}}}}{{}}
% \tablepngsincl[<max width>]{{<image>}}{{<natural width>}}: include at natural
% size, never wider than <max width> (default \linewidth) nor taller than
% {maxh} of \textheight. Used for ordinary table floats.
\providecommand{{\tablepngsincl}}[3][\linewidth]{{%
  \setlength{{\tablepngs@nat}}{{#3}}%
  \setlength{{\tablepngs@cap}}{{#1}}%
  \ifdim\tablepngs@nat>\tablepngs@cap\setlength{{\tablepngs@nat}}{{\tablepngs@cap}}\fi
  \includegraphics[width=\tablepngs@nat,height={maxh}\textheight,keepaspectratio]{{#2}}%
}}
% Page chunks of a multi-page table were typeset to fill the text block, so
% they are included at up to {pageh} of \textheight -- shrinking them further
% would render the table smaller than it is in the original document.
\providecommand{{\tablepngspage}}[3][\linewidth]{{%
  \setlength{{\tablepngs@nat}}{{#3}}%
  \setlength{{\tablepngs@cap}}{{#1}}%
  \ifdim\tablepngs@nat>\tablepngs@cap\setlength{{\tablepngs@nat}}{{\tablepngs@cap}}\fi
  \includegraphics[width=\tablepngs@nat,height={pageh}\textheight,keepaspectratio]{{#2}}%
}}
% Same, but leaving room for a caption sharing the page.
\providecommand{{\tablepngspagecap}}[3][\linewidth]{{%
  \setlength{{\tablepngs@nat}}{{#3}}%
  \setlength{{\tablepngs@cap}}{{#1}}%
  \ifdim\tablepngs@nat>\tablepngs@cap\setlength{{\tablepngs@nat}}{{\tablepngs@cap}}\fi
  \includegraphics[width=\tablepngs@nat,height={caph}\textheight,keepaspectratio]{{#2}}%
}}
% Glue used around a flattened table chunk.
\providecommand{{\tablepngsgap}}{{\par\addvspace{{\medskipamount}}}}
% Forbids a page break at this point, so a caption can never be separated
% from the image it belongs to. A minipage would do it too, but would also
% trap any \footnote in the caption inside the box.
\providecommand{{\tablepngsnobreak}}{{\nopagebreak[4]\par\nopagebreak[4]}}
\makeatother
%% ---- end tablepngs -------------------------------------------------------
"""


# --------------------------------------------------------------------------
# Replacement text
# --------------------------------------------------------------------------
def float_replacement(t, img_relpaths, widths_pt):
    cap = "\\textheight" if t.env.startswith("sideways") else "\\linewidth"
    incl = "\n".join(
        f"\\tablepngsincl[{cap}]{{{p}}}{{{w:.2f}pt}}"
        for p, w in zip(img_relpaths, widths_pt)
    )
    caption_block = ""
    if t.caption_raw:
        caption_block = t.caption_raw + "".join(t.labels_raw)
    elif t.labels_raw:
        caption_block = "".join(t.labels_raw)
    parts = [f"\\begin{{{t.env}}}{t.opt_str()}", "\\centering"]
    if t.multi_caption:
        parts.append(incl)
        # re-step the counter once per baked (non-starred) caption and
        # re-plant each caption's labels so \ref still resolves correctly
        for cap in t.caps_list:
            if cap["star"]:
                continue
            parts.append("\\refstepcounter{table}" + "".join(cap["labels"]) + "%")
    elif t.caption_raw and t.caption_pos == "above":
        parts += [caption_block, incl]
    elif caption_block:
        parts += [incl, caption_block]
    else:
        parts.append(incl)
    parts.append(f"\\end{{{t.env}}}")
    return "\n".join(parts)


def defuse_footnotes(caption_text):
    """Split \\footnote{...} out of a caption.

    Inside a float, \\caption defers its argument and \\footnote survives. A
    longtable caption has to be re-emitted with \\captionof, whose argument is
    rescanned by the caption package -- and a \\footnote there fails with
    "Argument of \\caption@ydblarg has an extra }". Replacing each footnote
    with \\footnotemark and emitting \\footnotetext afterwards produces the
    same printed result without the fragile argument."""
    masked = mask_comments(caption_text)
    spans, notes = [], []
    for m in re.finditer(r"\\footnote(?![a-zA-Z])\s*", masked):
        i = m.end()
        if i < len(masked) and masked[i] == "[":
            try:
                i = match_bracket(masked, i)
            except TPError:
                continue
        if i >= len(masked) or masked[i] != "{":
            continue
        try:
            j = match_brace(masked, i)
        except TPError:
            continue
        spans.append((m.start(), j))
        notes.append(caption_text[i + 1:j - 1])
    if not spans:
        return caption_text, []
    out, pos = [], 0
    for s, e in spans:
        out.append(caption_text[pos:s])
        out.append("\\footnotemark{}")
        pos = e
    out.append(caption_text[pos:])
    return "".join(out), notes


def long_replacement(t, img_relpaths, widths_pt):
    """Re-emit a multi-page table as a one-column longtable of images.

    Using a real longtable rather than \\captionof matters for two reasons.
    First, the caption is reproduced VERBATIM inside the same kind of
    environment it came from, so constructs the document already proved it can
    typeset -- a \\footnote or a nested \\label in the caption, both of which
    break \\captionof -- keep working. Second, longtable rows break across
    pages natively, so the page chunks flow one after another without floating
    away from their anchor.
    """
    last = len(img_relpaths) - 1
    rows = []
    caption_row = None
    if t.caption_raw:
        caption_row = t.caption_raw + "".join(t.labels_raw)
    for i, (p, w) in enumerate(zip(img_relpaths, widths_pt)):
        shares = (caption_row is not None
                  and ((i == 0 and t.caption_pos == "above")
                       or (i == last and t.caption_pos == "below")))
        macro = "tablepngspagecap" if shares else "tablepngspage"
        rows.append(f"\\{macro}{{{p}}}{{{w:.2f}pt}}")
    body = []
    if caption_row is not None and t.caption_pos == "above":
        body.append(caption_row + "\\\\*")
    body += [r + "\\\\" for r in rows]
    if caption_row is not None and t.caption_pos == "below":
        body.append(caption_row + "\\\\")
    if caption_row is None and t.labels_raw:
        # keep \ref targets working even with no caption
        body.append("".join(t.labels_raw) + "%")
    return ("\\begin{longtable}{@{}c@{}}\n" + "\n".join(body)
            + "\n\\end{longtable}")


def bare_replacement(t, img_relpaths, widths_pt):
    return "".join(
        f"\\tablepngsincl{{{p}}}{{{w:.2f}pt}}"
        for p, w in zip(img_relpaths, widths_pt)
    )


def _opt_str(self):
    return self.span.opt or ""


Target.opt_str = _opt_str


# --------------------------------------------------------------------------
# Verification
# --------------------------------------------------------------------------
def extract_pdf_text(pdf, tools, first=None, last=None):
    if tools.text_extract == "pdftotext":
        cmd = ["pdftotext", "-enc", "UTF-8"]
        if first is not None:
            cmd += ["-f", str(first), "-l", str(last or first)]
        rc, _ = run(cmd + [str(pdf), str(pdf) + ".txt"])
        if rc == 0:
            txt, _ = read_text_guess(str(pdf) + ".txt")
            return txt
    if tools.gs:
        outtxt = str(pdf) + ".gs.txt"
        rc, _ = run([tools.gs, "-dSAFER", "-dBATCH", "-dNOPAUSE",
                     "-sDEVICE=txtwrite", "-o", outtxt, str(pdf)])
        if rc == 0 and Path(outtxt).exists():
            txt, _ = read_text_guess(outtxt)
            return txt
    return None


def norm_alnum(s):
    return re.sub(r"[^a-z0-9.]", "", s.lower())


def strip_latex(s):
    s = re.sub(r"\\[a-zA-Z@]+\s*(\[[^\]]*\])?", " ", s)
    s = re.sub(r"[{}&~$^_%]", " ", s)
    s = s.replace("\\\\", " ")
    return s


def pick_probes(content, rest_of_doc_norm, k=6):
    """Distinctive tokens from a table body that appear nowhere else in the
    document source — used to prove the table text left the PDF text layer."""
    cleaned = strip_latex(content)
    toks = re.split(r"[\s,;()\[\]]+", cleaned)
    seen, nums, words = set(), [], []
    for tok in toks:
        tok = tok.strip(".-–—")
        if not tok or tok in seen:
            continue
        seen.add(tok)
        tn = norm_alnum(tok)
        if len(tn) < 4 or tn in rest_of_doc_norm:
            continue
        if re.fullmatch(r"\d[\d.,]{3,}", tok):
            nums.append(tn)
        elif re.fullmatch(r"[A-Za-z][A-Za-z'-]{4,}", tok):
            words.append(tn)
    probes = []
    for a, b in zip(nums + [None] * k, words + [None] * k):
        if a:
            probes.append(a)
        if b:
            probes.append(b)
        if len(probes) >= k:
            break
    return probes[:k]


def caption_probe(caption_text):
    """A literal run of caption text to look for in the output.

    The probe must avoid macros: a caption like "How the \\dcUniverseN{}
    operating sites" prints an expanded number that the source does not
    contain, so a probe spanning the macro would never match. Taking the
    longest macro-free stretch keeps the check honest. Returns '' when there
    is no stretch long enough to be meaningful, in which case the caller
    skips the check rather than guessing."""
    text = re.sub(r"\\\\|\\[ ,;:!]", " ", caption_text)
    segments = re.split(r"\\[a-zA-Z@]+\*?(?:\[[^\]]*\])?(?:\{[^{}]*\})?|[{}$&~^_]",
                        text)
    best = ""
    for seg in segments:
        n = norm_alnum(seg)
        if len(n) > len(best):
            best = n
    return best[:24] if len(best) >= 10 else ""


def verify(targets, final_pdf, tools):
    text = extract_pdf_text(final_pdf, tools)
    if text is None:
        warn("no text extractor available (install poppler for pdftotext, or ghostscript); skipping text-layer verification")
        return None
    hay = norm_alnum(text)
    all_ok = True
    results = []
    for t in targets:
        if not t.ok:
            results.append((t, None, None))
            continue
        leaked = [p for p in t.probes if p and p in hay]
        cap_probe = caption_probe(t.caption_text) if t.caption_text and not t.multi_caption else ""
        cap_found = (cap_probe in hay) if cap_probe else None
        if leaked or cap_found is False:
            all_ok = False
        results.append((t, leaked, cap_found))
    return all_ok, results


# --------------------------------------------------------------------------
# Visual verification (--compare)
# --------------------------------------------------------------------------
ENV_HEADER_ARGS = {  # mandatory {..} groups after \begin{env} (colspecs etc.)
    "tabular": 1, "tabular*": 2, "tabularx": 2, "tabulary": 2, "tabu": 1,
    "longtable": 1, "longtabu": 1, "xltabular": 2, "supertabular": 1,
    "NiceTabular": 1, "minipage": 1, "wraptable": 2,
}


def strip_env_headers(content):
    """Remove \\begin{env}[opt]{colspec}... header arguments so column specs
    like {lrr} or {p{3cm}} never masquerade as table data tokens."""
    masked = mask_comments(content)
    cuts = []
    for m in re.finditer(r"\\begin\s*\{([A-Za-z*]+)\}", masked):
        nargs = ENV_HEADER_ARGS.get(m.group(1))
        if not nargs:
            continue
        i = m.end()
        try:
            while i < len(masked) and masked[i] in " \t\n":
                i += 1
            if i < len(masked) and masked[i] == "[":
                i = match_bracket(masked, i)
            for _ in range(nargs):
                while i < len(masked) and masked[i] in " \t\n":
                    i += 1
                if i < len(masked) and masked[i] == "{":
                    i = match_brace(masked, i)
        except (TPError, AssertionError):
            continue
        cuts.append((m.end(), i))
    out, pos = [], 0
    for s, e in cuts:
        out.append(content[pos:s])
        pos = e
    out.append(content[pos:])
    return "".join(out)


def longtable_rows_region(content):
    """The data-row region of a longtable: everything after the last
    \\endfirsthead/\\endhead/\\endfoot/\\endlastfoot spec terminator. The spec
    blocks contain text that legitimately may not render (a one-page table
    never shows its 'continued' foot), so they are excluded from the
    completeness check."""
    masked = mask_comments(content)
    last = None
    for m in re.finditer(r"\\end(firsthead|head|foot|lastfoot)(?![a-zA-Z])", masked):
        last = m.end()
    return content[last:] if last is not None else content


def source_literal_tokens(content, kind="float"):
    """Literal data tokens written in the table source (words and numbers that
    are not macro-generated). Every one of them must show up in the rendered
    image's text layer, or the render dropped content."""
    if kind == "long":
        content = longtable_rows_region(content)
    # comments never render, so their words must not count as table content
    content = "".join(c if mc != " " or c.isspace() else " "
                      for c, mc in zip(content, mask_comments(content)))
    content = strip_env_headers(content)
    # declarations whose arguments are values, not text
    content = re.sub(
        r"\\(?:renewcommand|newcommand|providecommand|setlength|addtolength"
        r"|definecolor|arrayrulewidth|setcounter|settowidth)\s*\*?"
        r"(?:\{[^{}]*\}|\\[a-zA-Z@]+)(?:\s*\[[^\]]*\])?\s*(?:\{[^{}]*\})?"
        r"(?:\s*\{[^{}]*\})?", " ", content)
    # rules and spans: the column/range arguments are not content
    content = re.sub(r"\\(?:cmidrule|cline)\s*(?:\([^)]*\))?\s*\{[^{}]*\}",
                     " ", content)
    content = re.sub(r"\\multicolumn\s*\{[^{}]*\}\s*\{[^{}]*\}", " ", content)
    content = re.sub(r"\\multirow\s*\*?(?:\[[^\]]*\])?\s*\{[^{}]*\}"
                     r"(?:\s*\[[^\]]*\])?\s*\{[^{}]*\}", " ", content)
    # arguments of these commands are keys/parameters, not typeset text
    content = re.sub(r"\\rowcolors\*?\s*\{[^{}]*\}\s*\{[^{}]*\}\s*\{[^{}]*\}",
                     " ", content)
    content = re.sub(
        r"\\(?:ref|pageref|eqref|autoref|[cC]ref|vref|label|cite[a-zA-Z]*"
        r"|rowcolors\*?|rowcolor|cellcolor|arrayrulecolor|columncolor"
        r"|hyperref|hypertarget|hyperlink|includegraphics|input|graphicspath)"
        r"\s*(?:\[[^\]]*\])?\s*\{[^{}]*\}(?:\s*\{[^{}]*\})?",
        " ", content)
    # \textcolor{name}{TEXT}: drop the color name, keep the text
    content = re.sub(r"\\(?:textcolor|colorbox)\s*(?:\[[^\]]*\])?\{[^{}]*\}", " ", content)
    content = re.sub(r"\\fcolorbox\s*\{[^{}]*\}\s*\{[^{}]*\}", " ", content)
    content = re.sub(r"\\(?:begin|end)\s*\{[^{}]*\}", " ", content)
    s = strip_latex(content)
    out, seen = [], set()
    for tok in re.findall(r"[A-Za-z][A-Za-z'’-]{2,}|\d[\d.,]*\d|\d{3,}", s):
        n = norm_alnum(tok)
        if len(n) >= 3 and n not in seen:
            seen.add(n)
            out.append(n)
    return out


def pages_containing(pdf, probes, tools):
    """Page numbers of `pdf` whose text layer contains any of `probes`.
    One extraction for the whole file: pdftotext separates pages with a form
    feed, so per-page scanning costs nothing extra and there is no page cap."""
    if not probes:
        return []
    txt = extract_pdf_text(pdf, tools)
    if txt is None:
        return []
    return [i for i, page in enumerate(txt.split("\f"), 1)
            if any(pr in norm_alnum(page) for pr in probes)]


def build_reference_doc(text, masked, targets, preamble_end, aux_name=None):
    """Build a copy of the whole document in which each rendered target is
    replaced by `\\begin{preview}<its render content>\\end{preview}`, with the
    preview package active. Compiling this yields ONE tightly cropped page per
    target, showing that table typeset with every bit of its real in-document
    context (macros defined mid-document, counters, changed lengths).

    Comparing those pages against the PNGs produced from the isolated snippets
    is a genuine differential test of the isolation: if the snippet is missing
    context, the two renderings differ."""
    new = text
    for t in sorted(targets, key=lambda t: t.span.start, reverse=True):
        body = t.render_content
        if t.multi_caption:
            body = ("\\makeatletter\\def\\@captype{table}\\makeatother%\n" + body)
        # a float's content is not a float here, so preview can box it safely
        new = (new[:t.span.start]
               + f"\\begin{{preview}}%\n{body}%\n\\end{{preview}}"
               + new[t.span.end:])
    # neutralize landscape wrappers: with `active` preview only the preview
    # boxes are output, and removing the wrapper keeps reference pages
    # unrotated so they align with the candidate PNGs
    new = re.sub(r"\\(begin|end)\s*\{landscape\}", "", new)
    # import the main .aux so \ref and bibtex \cite inside cells resolve to
    # the same text the candidate snippet produced (otherwise they render as
    # "?" here and the pixel comparison flags a difference that is not real)
    inject = ("\\usepackage[active,tightpage]{preview}\n"
              "\\setlength\\PreviewBorder{1.5pt}\n"
              "\\makeatletter\\@ifundefined{captionof}"
              "{\\usepackage{capt-of}}{}\\makeatother\n"
              + _aux_hook(aux_name))
    m = re.search(r"\\begin\s*\{document\}", mask_comments(new))
    return new[:m.start()] + inject + new[m.start():]


def montage_vertical(pngs, out, tools_magick):
    """Stack PNGs vertically into one image (for multi-page longtables)."""
    rc, log = run([tools_magick, *[str(p) for p in pngs],
                   "-background", "white", "-gravity", "center",
                   "-append", str(out)])
    return rc == 0 and Path(out).exists()


FONT_CANDIDATES = [
    "/System/Library/Fonts/Supplemental/Arial.ttf",       # macOS
    "/System/Library/Fonts/Helvetica.ttc",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",    # Debian/Ubuntu
    "/usr/share/fonts/dejavu/DejaVuSans.ttf",             # Fedora
    "/usr/share/fonts/TTF/DejaVuSans.ttf",                # Arch
    "C:/Windows/Fonts/arial.ttf",                         # Windows
    "C:/Windows/Fonts/segoeui.ttf",
]


def find_label_font(magick):
    """ImageMagick often has no font registered (macOS especially), which
    makes `label:` fail. Find a font file we can pass explicitly; return None
    if labelling is not possible (then sheets get colored borders instead)."""
    for f in FONT_CANDIDATES:
        if Path(f).is_file():
            rc, _ = run([magick, "-background", "black", "-fill", "white",
                         "-pointsize", "16", "-font", f, "label:test",
                         "null:"])
            if rc == 0:
                return f
    rc, _ = run([magick, "-background", "black", "-fill", "white",
                 "-pointsize", "16", "label:test", "null:"])
    return "" if rc == 0 else None   # "" = default font works


def compare_images(ref, cand, out_sidebyside, magick, label_ref, label_cand,
                   font=None):
    """Resize the reference to the candidate's width, compute a normalized
    RMSE difference, and write a labelled side-by-side contact sheet.
    Returns the normalized score (0 = identical, 1 = maximally different),
    or None if the comparison could not be made."""
    # Trim surrounding whitespace from BOTH first: the in-context reference
    # and the flattened PNG are cropped by different mechanisms (preview vs
    # pdfcrop), so raw framing differs even when the tables are identical.
    stem = Path(out_sidebyside).stem
    tref = Path(out_sidebyside).with_name(f"_tr_{stem}.png")
    tcand = Path(out_sidebyside).with_name(f"_tc_{stem}.png")
    for src, dest in ((ref, tref), (cand, tcand)):
        rc, _ = run([magick, str(src), "-background", "white", "-alpha", "remove",
                     "-fuzz", "2%", "-trim", "+repage", "-colorspace", "sRGB",
                     str(dest)])
        if rc != 0 or not dest.exists():
            shutil.copy(src, dest)
    ref, cand = tref, tcand

    w, h = png_dims(cand)
    tmp_ref = Path(out_sidebyside).with_name(f"_rs_{stem}.png")
    rc, _ = run([magick, str(ref), "-resize", f"{w}x{h}!",
                 "-colorspace", "sRGB", "-alpha", "remove", str(tmp_ref)])
    if rc != 0:
        return None
    rc, out = run([magick, "compare", "-metric", "RMSE", str(tmp_ref),
                   str(cand), "null:"])
    score = None
    m = re.search(r"\(([0-9.eE+-]+)\)", out)
    if m:
        try:
            score = float(m.group(1))
        except ValueError:
            score = None
    # labelled side-by-side for human / Claude visual review
    def labelled(src, text, color, dest):
        cmd = [magick, str(src), "-resize", "700x1400>",
               "-bordercolor", color, "-border", "3"]
        if font is not None:
            cmd += ["-background", color, "-fill", "white", "-pointsize", "18"]
            if font:
                cmd += ["-font", font]
            cmd += ["label:" + text, "+swap", "-gravity", "center", "-append"]
        cmd.append(str(dest))
        rc, _ = run(cmd)
        if rc != 0:  # last resort: plain bordered copy
            run([magick, str(src), "-resize", "700x1400>",
                 "-bordercolor", color, "-border", "3", str(dest)])

    a = Path(out_sidebyside).with_name(f"_a_{Path(out_sidebyside).stem}.png")
    b = Path(out_sidebyside).with_name(f"_b_{Path(out_sidebyside).stem}.png")
    labelled(ref, label_ref, "#1f5fa8", a)      # blue  = reference
    labelled(cand, label_cand, "#c1651b", b)    # orange = flattened
    rc, _ = run([magick, str(a), str(b), "-background", "white",
                 "-splice", "16x0", "+append", "-bordercolor", "white",
                 "-border", "10", str(out_sidebyside)])
    for p in (tmp_ref, a, b, tref, tcand):
        p.unlink(missing_ok=True)
    return score


def visual_compare(done, text, masked, preamble_end, build, imgdir, maindir,
                   engine, aux_name, encoding, args, tools, orig_pdf):
    """Render every flattened table a second time IN CONTEXT and compare it
    against the PNG that went into the document."""
    magick = which("magick") or which("compare")
    if not magick:
        warn("--compare needs ImageMagick (brew install imagemagick / "
             "apt install imagemagick / choco install imagemagick); skipping "
             "visual comparison")
        return None
    cmpdir = imgdir / "_compare"
    cmpdir.mkdir(exist_ok=True)
    refdir = build / "reference"
    refdir.mkdir(exist_ok=True)
    if aux_name:
        shutil.copy(build / aux_name, refdir / aux_name)

    font = find_label_font(magick)

    # --- reference A: preview-rendered targets get an exact in-context
    # differential (same cropping path as the candidate, so pixels line up).
    # Longtables are excluded: boxing a multi-page table into a single
    # preview page overflows TeX's maximum dimension.
    prev_targets = [t for t in done if t.method == "preview"]
    ref_pages = []
    if prev_targets:
        info(f"visual check: re-rendering {len(prev_targets)} table(s) in "
             "their real document context...")
        ref_tex = refdir / "reference.tex"
        ref_tex.write_text(
            build_reference_doc(text, masked, prev_targets, preamble_end,
                                aux_name=aux_name),
            encoding=encoding)
        ok, log, ref_pdf = run_latex(engine, ref_tex, refdir, maindir, passes=2,
                                     shell_escape=args.shell_escape,
                                     timeout=args.timeout)
        if not ok:
            warn("visual check: the in-context reference document failed to "
                 f"compile; falling back to original-PDF pages.\n--- log ---\n{log}")
        else:
            try:
                ref_pages = pdf_to_pngs(ref_pdf, refdir, "ref", args.dpi, tools)
            except TPError as e:
                warn(f"visual check: could not rasterize the reference ({e})")
            if ref_pages and len(ref_pages) != len(prev_targets):
                warn(f"visual check: reference produced {len(ref_pages)} page(s) "
                     f"for {len(prev_targets)} table(s); falling back to "
                     "original-PDF pages")
                ref_pages = []
    ref_map = dict(zip([t.index for t in prev_targets], ref_pages)) if ref_pages else {}

    orig_raw = None
    results = []
    for t in done:
        # --- content completeness (all table kinds, page-break invariant):
        # every literal token written in the table source must appear in the
        # text layer of the PDF we rasterized.
        missing, n_src = None, 0
        cand_txt = (extract_pdf_text(Path(t.snippet_pdf), tools)
                    if t.snippet_pdf and Path(t.snippet_pdf).exists() else None)
        if cand_txt is not None:
            hay = norm_alnum(cand_txt)
            src = source_literal_tokens(t.render_content, t.kind)
            n_src = len(src)
            missing = [tok for tok in src if tok not in hay]

        # --- candidate image (stack the pages of a multi-page table)
        cand = imgdir / Path(t.images[0]).name
        if len(t.images) > 1:
            merged = cmpdir / f"t{t.index:02d}-candidate-stacked.png"
            if montage_vertical([imgdir / Path(p).name for p in t.images],
                                merged, magick):
                cand = merged

        # --- reference image: in-context preview render, else the pages of
        # the ORIGINAL pre-conversion PDF that contain this table
        ref = ref_map.get(t.index)
        ref_label = f"REFERENCE  t{t.index:02d}  (typeset in the document)"
        if ref is None:
            pgs = pages_containing(orig_pdf, t.probes, tools)
            if pgs:
                if orig_raw is None:
                    orig_raw = pdf_to_pngs(orig_pdf, refdir, "origall",
                                           max(110, args.dpi // 3), tools)
                raw = orig_raw
                picked = [raw[p - 1] for p in pgs if p - 1 < len(raw)]
                merged_ref = cmpdir / f"t{t.index:02d}-reference-origpages.png"
                if picked and montage_vertical(picked, merged_ref, magick):
                    ref = merged_ref
                    ref_label = (f"REFERENCE  t{t.index:02d}  (original PDF "
                                 f"page{'s' if len(picked) > 1 else ''} "
                                 f"{','.join(map(str, pgs))})")
        if ref is None:
            results.append((t, None, None, "none", missing, n_src))
            continue

        ref_kind = "preview" if t.index in ref_map else "origpages"
        side = cmpdir / f"t{t.index:02d}-compare.png"
        score = compare_images(ref, cand, side, magick, ref_label,
                               f"FLATTENED  t{t.index:02d}  (PNG in the output)",
                               font=font)
        results.append((t, score, side, ref_kind, missing, n_src))
    return results


# --------------------------------------------------------------------------
# Main pipeline
# --------------------------------------------------------------------------
def process(args):
    main_tex = Path(args.texfile).resolve()
    if not main_tex.exists():
        die(f"file not found: {main_tex}")
    maindir = main_tex.parent
    stem = main_tex.stem

    text, encoding = read_text_guess(main_tex)

    # refuse to flatten our own output (running twice would rasterize the
    # PNGs again and duplicate the macro block)
    if "Generated by tablepngs" in text[:400] or "\\tablepngsincl" in text:
        die(f"{main_tex.name} looks like tablepngs output already "
            "(found the tablepngs header / \\tablepngsincl). Re-run tablepngs "
            "on your ORIGINAL document instead.")

    # pull \input/\include-ed files into the working text so tables that live
    # in separate fragment files are found and flattened too
    text, n_inlined = inline_inputs(text, maindir)
    if n_inlined:
        info(f"inlined {n_inlined} \\input/\\include file(s) for scanning")
    masked = mask_comments(text)

    engine = args.engine
    auto_engine = engine == "auto"
    if auto_engine:
        engine, why = detect_engine(text, masked, basedir=maindir,
                                    main_tex=main_tex)
        info(f"engine auto-detected: {engine} (from {why})")
    tools = discover_tools(engine)
    if not tools.engine:
        die(f"LaTeX engine '{engine}' not found on PATH — run with --check for install help")
    if not tools.raster:
        die("no PDF->PNG tool found (need pdftoppm, magick, or gs) — run with --check")
    info(f"rasterizer: {tools.raster} @ {args.dpi} dpi")

    # ---- scan --------------------------------------------------------------
    targets, scan_notes, unsupported = scan_targets(text, masked)
    for note in scan_notes:
        info(note)
    for env, line in unsupported:
        warn(f"line {line}: \\begin{{{env}}} is {UNSUPPORTED_ENVS[env]}. "
             f"tablepngs will NOT flatten it — that table stays as live text "
             f"and Word may still reflow it.")
    if args.no_bare:
        targets = [t for t in targets if t.kind != "bare"]
    if args.only:
        keep = {int(x) for x in args.only.split(",")}
        targets = [t for t in targets if t.index in keep]
    if args.skip:
        drop = {int(x) for x in args.skip.split(",")}
        targets = [t for t in targets if t.index not in drop]
    if not targets:
        info("no tables found — nothing to do")
        return 0

    n_float = sum(1 for t in targets if t.kind == "float")
    n_long = sum(1 for t in targets if t.kind == "long")
    n_bare = sum(1 for t in targets if t.kind == "bare")
    info(f"found {len(targets)} table target(s): {n_float} float(s), "
         f"{n_long} longtable(s), {n_bare} bare tabular(s)")
    for t in targets:
        preview = re.sub(r"\\(?:label|cite[a-zA-Z]*|ref|nonumber)\s*"
                         r"(?:\[[^\]]*\])?\{[^{}]*\}", " ", t.caption_text)
        cap = (re.sub(r"\s+", " ", strip_latex(preview)).strip()[:48]
               or "(no caption)")
        extra = " [landscape]" if t.in_landscape else ""
        print(f"    t{t.index:02d}  {t.env:<14} line {t.line:>5}  {cap}{extra}")
    if args.list:
        return 0

    # citations/refs inside tables need the aux trick; biblatex is a limitation
    for t in targets:
        if re.search(r"\\cite[a-zA-Z]*\b", t.render_content):
            info(f"t{t.index:02d} contains \\cite — resolving via the main document's .aux")
            break

    # ---- workspace ---------------------------------------------------------
    imgdir_name = args.imgdir or f"{stem}_tablepngs"
    imgdir = maindir / imgdir_name
    imgdir.mkdir(exist_ok=True)
    build = imgdir / "_build"
    if build.exists():
        shutil.rmtree(build)
    build.mkdir(parents=True)

    # ---- compile the original once (validates it + creates .aux) -----------
    info(f"compiling original with {engine} (baseline)...")
    ok, log, orig_pdf = run_latex(engine, main_tex, build / "orig", maindir,
                                  passes=2, shell_escape=args.shell_escape,
                                  timeout=args.timeout)
    # Retry other engines only when the failure actually looks like an engine
    # mismatch. A missing graphics file or an undefined macro fails identically
    # under every engine, and retrying just buries the real error.
    ENGINE_MISMATCH = (
        r"cannot-use-pdftex|cannot be used with|only be used with|"
        r"requires (?:XeTeX|LuaTeX|xelatex|lualatex)|"
        r"Unicode character|invalid in PDF mode|\\XeTeXversion|"
        r"fontspec|polyglossia|unicode-math|luatex(?:base)?\.sty")
    if not ok and auto_engine and re.search(ENGINE_MISMATCH, log, re.I):
        hint = " (the log points at an engine/font mismatch)"
        for alt in [e for e in ("xelatex", "lualatex", "pdflatex") if e != engine]:
            if not which(alt):
                continue
            info(f"{engine} could not build this document{hint}; "
                 f"retrying with {alt}...")
            ok2, log2, pdf2 = run_latex(alt, main_tex, build / "orig", maindir,
                                        passes=2, shell_escape=args.shell_escape,
                                        timeout=args.timeout)
            if ok2:
                info(f"{alt} succeeded — using {alt} for this run "
                     f"(pass --engine {alt} to skip this retry next time)")
                engine, ok, log, orig_pdf = alt, ok2, log2, pdf2
                tools = discover_tools(engine)
                break
        if not ok:
            warn("no available LaTeX engine could build this document; "
                 "reporting the first engine's error below")
    if not ok:
        if not args.keep_build:
            shutil.rmtree(build, ignore_errors=True)
            try:
                imgdir.rmdir()  # remove only if we just created it empty
            except OSError:
                pass
        die(f"the ORIGINAL document failed to compile with {engine}, so there "
            f"is nothing tablepngs can safely flatten.\n"
            f"  * If the document builds for you with a different engine, say so: "
            f"--engine xelatex|lualatex|pdflatex\n"
            f"  * If it needs shell escape (minted, gnuplot): --shell-escape\n"
            f"  * Otherwise fix the LaTeX error below and re-run.\n"
            f"--- log excerpt ---\n{log}")
    orig_pages = pdf_page_count(orig_pdf, tools)

    # prefer a user-compiled aux (it has bibliography info); else ours
    aux_src = main_tex.with_suffix(".aux")
    if not aux_src.exists():
        aux_src = build / "orig" / f"{stem}.aux"
    aux_name = None
    if aux_src.exists():
        aux_name = "tablepngs_main_aux"  # extensionless: \@input appends nothing
        shutil.copy(aux_src, build / (aux_name + ".tex"))
        aux_name += ".tex"

    doc_m = re.search(r"\\begin\s*\{document\}", masked)
    preamble = text[:doc_m.start()]

    # ---- render each target ------------------------------------------------
    for t in targets:
        stem_t = f"t{t.index:02d}"
        snipdir = build / stem_t
        snipdir.mkdir()
        if aux_name:
            shutil.copy(build / aux_name, snipdir / aux_name)
        # \thetable value to seed inside the snippet so continued heads and
        # any in-body \thetable print the number this table really had:
        #  - long/multi-caption: the env or the first \captionof will step it,
        #    so seed at caps_before;
        #  - single above-caption float: caption already stepped it -> +1;
        #  - single below-caption float: not yet stepped -> caps_before.
        if t.kind == "long" or t.multi_caption:
            ctr = t.caps_before
        else:
            ctr = t.caps_before + (1 if (t.n_captions == 1 and t.caption_pos == "above") else 0)
        seed = t.counter_seed
        if t.kind == "long":
            snippet = build_page_snippet(preamble, t.render_content, aux_name,
                                         landscape=t.in_landscape, table_counter=ctr,
                                         parent_seed=seed)
            passes, t.method = 3, "pagecrop"
        else:
            snippet = build_preview_snippet(preamble, t.render_content, aux_name,
                                            table_counter=ctr, parent_seed=seed,
                                            wrap_captions=t.multi_caption)
            passes, t.method = 2, "preview"
        snip_tex = snipdir / f"{stem_t}.tex"
        snip_tex.write_text(snippet, encoding=encoding)
        ok, log, pdf = run_latex(engine, snip_tex, snipdir, maindir,
                                 passes=passes, shell_escape=args.shell_escape,
                                 timeout=args.timeout)
        if not ok and t.method == "preview":
            # fallback: some content (e.g. stray \\pagebreak) rejects preview
            snippet = build_page_snippet(preamble, t.render_content, aux_name,
                                         landscape=t.in_landscape, table_counter=ctr,
                                         parent_seed=seed, wrap_captions=t.multi_caption)
            snip_tex.write_text(snippet, encoding=encoding)
            ok, log, pdf = run_latex(engine, snip_tex, snipdir, maindir,
                                     passes=2, shell_escape=args.shell_escape,
                                     timeout=args.timeout)
            t.method = "pagecrop-fallback"
        if not ok:
            warn(f"t{t.index:02d} ({t.env}, line {t.line}) failed to compile — "
                 f"leaving this table as-is.\n--- log excerpt ---\n{log}\n")
            t.note = "snippet compile failed; left unflattened"
            continue
        if t.method != "preview":
            pdf = crop_pdf(pdf, tools)
        t.snippet_pdf = str(pdf)
        try:
            pngs = pdf_to_pngs(pdf, imgdir, stem_t, args.dpi, tools)
        except TPError as e:
            warn(f"t{t.index:02d}: {e}; leaving this table as-is")
            t.note = "rasterization failed; left unflattened"
            continue
        for p in pngs:
            w_px, h_px = png_dims(p)
            t.widths_pt.append(w_px * 72.0 / args.dpi)
            t.images.append(f"{imgdir_name}/{p.name}")
        t.ok = True
        # exclusion haystack: everything OUTSIDE this table, plus its caption
        # and labels (those stay live text in the output on purpose)
        rest = norm_alnum(strip_latex(
            text[:t.span.start] + text[t.span.end:]
            + " " + t.caption_text + " " + " ".join(t.labels_raw)))
        t.probes = pick_probes(t.render_content, rest)
        sizes = ", ".join(f"{png_dims(imgdir / Path(p).name)[0]}x{png_dims(imgdir / Path(p).name)[1]}px"
                          for p in t.images[:3])
        more = f" (+{len(t.images)-3} more)" if len(t.images) > 3 else ""
        info(f"t{t.index:02d} -> {len(t.images)} PNG(s) [{t.method}] {sizes}{more}")

    done = [t for t in targets if t.ok]
    if not done:
        die("no tables could be rendered — nothing to rewrite", 1)

    # ---- rewrite the document ---------------------------------------------
    out_tex = maindir / f"{stem}{args.suffix}.tex"
    newtext = text
    for t in sorted(done, key=lambda t: t.span.start, reverse=True):
        if t.kind == "float":
            repl = float_replacement(t, t.images, t.widths_pt)
        elif t.kind == "long":
            repl = long_replacement(t, t.images, t.widths_pt)
        else:
            repl = bare_replacement(t, t.images, t.widths_pt)
        newtext = newtext[:t.span.start] + repl + newtext[t.span.end:]
    try:
        caph = f"{max(0.50, float(args.page_height) - 0.10):.2f}"
    except ValueError:
        die(f"--page-height must be a number, got {args.page_height!r}")
    inject = INJECT_BLOCK.format(version=__version__, maxh=args.max_height,
                                 pageh=args.page_height, caph=caph)
    m = re.search(r"\\begin\s*\{document\}", mask_comments(newtext))
    newtext = newtext[:m.start()] + inject + newtext[m.start():]
    header = (f"% Generated by tablepngs v{__version__} from {main_tex.name} — "
              f"tables flattened to PNGs in {imgdir_name}/\n")
    out_tex.write_text(header + newtext, encoding=encoding)
    info(f"wrote {out_tex.name}")

    # ---- compile the flattened document ------------------------------------
    info(f"compiling {out_tex.name} with {engine}...")
    outbuild = build / "final"
    outbuild.mkdir()
    for ext in (".bbl", ".ind", ".gls"):
        side = main_tex.with_suffix(ext)
        if side.exists():
            shutil.copy(side, outbuild / (out_tex.stem + ext))
    ok, log, final_pdf = run_latex(engine, out_tex, outbuild, maindir,
                                   passes=2, shell_escape=args.shell_escape,
                                   timeout=args.timeout)
    if not ok:
        die(f"the flattened document ({out_tex.name}) failed to compile, so no "
            f"PDF was produced.\n"
            f"  The original document is untouched and still fine.\n"
            f"  The generated .tex was kept so you can inspect it, and the "
            f"PNGs in {imgdir_name}/ are usable on their own.\n"
            f"  Most likely one table's replacement does not fit where the "
            f"table used to be. Try re-running with --skip <n> for the table "
            f"named in the error, or --page-height 0.9 if the error mentions "
            f"an overfull \\vbox.\n"
            f"  If the error looks like a tablepngs defect, please report it "
            f"with the excerpt below.\n--- log ---\n{log}", 1)
    deliver_pdf = maindir / f"{out_tex.stem}.pdf"
    shutil.copy(final_pdf, deliver_pdf)
    final_pages = pdf_page_count(deliver_pdf, tools)
    pg = (f"{orig_pages} -> {final_pages} pages" if orig_pages and final_pages
          else "")
    info(f"wrote {deliver_pdf.name} {pg}")
    if orig_pages and final_pages and final_pages > orig_pages * 1.10:
        grew = final_pages - orig_pages
        warn(f"the flattened document grew by {grew} page(s) "
             f"({orig_pages} -> {final_pages}). A rasterized table cannot be "
             f"split across a page the way a live longtable can, so a large "
             f"table that used to start mid-page now moves to the next page. "
             f"If that matters, --page-height 0.99 packs them tighter, or "
             f"--skip <n> leaves a particular table as live text.")

    # ---- verify ------------------------------------------------------------
    rc = 0
    if not args.no_verify:
        v = verify(done, final_pdf, tools)
        if v is not None:
            all_ok, results = v
            leaks = 0
            for t, leaked, cap_found in results:
                if leaked is None:
                    continue
                if leaked:
                    leaks += len(leaked)
                    warn(f"t{t.index:02d}: table text leaked into the PDF text layer: "
                         f"{leaked} — flattening may have missed part of this table")
                if cap_found is False:
                    warn(f"t{t.index:02d}: caption text not found in text layer "
                         "(captions should remain live text)")
            n_probes = sum(len(t.probes) for t in done)
            n_caps = sum(1 for t in done if t.caption_text and not t.multi_caption)
            if all_ok:
                info(f"VERIFY PASS — 0/{n_probes} table-text probes leaked; "
                     f"{n_caps} caption(s) confirmed as live text")
            else:
                warn("VERIFY FAILED — see above")
                rc = 2

    # ---- visual comparison (opt-in) ----------------------------------------
    if args.compare:
        cmp_results = visual_compare(done, text, masked, doc_m.start(), build,
                                     imgdir, maindir, engine, aux_name,
                                     encoding, args, tools, orig_pdf)
        if cmp_results:
            bad = 0
            for t, score, side, ref_kind, missing, n_src in cmp_results:
                # content completeness is authoritative and page-break safe;
                # the pixel score only means something when the reference came
                # through the SAME preview cropping path as the candidate
                if missing is None:
                    content = "content n/a"
                    content_bad = False
                else:
                    # A word hyphenated across lines inside a narrow cell can
                    # be unrecoverable from the text layer, because extraction
                    # interleaves neighbouring columns between the fragments.
                    # Tolerate a few such tokens; a genuinely dropped row costs
                    # far more than that and still fails.
                    tol = max(3, int(0.03 * n_src))
                    content_bad = len(missing) > tol
                    if not missing:
                        content = f"content {n_src}/{n_src} ok"
                    elif content_bad:
                        content = f"content {n_src - len(missing)}/{n_src} tokens"
                    else:
                        content = (f"content {n_src - len(missing)}/{n_src} "
                                   f"(~hyphenation)")
                pixel_valid = ref_kind == "preview"
                pixel_bad = (pixel_valid and score is not None
                             and score > args.compare_threshold)
                if score is None or not pixel_valid:
                    pix = ("pixels n/a" if ref_kind == "none"
                           else "pixels n/a (see sheet)" if not pixel_valid
                           else "pixels n/a")
                else:
                    pix = f"pixels {score:.4f}"
                tag = "MISMATCH" if (content_bad or pixel_bad) else "ok"
                if tag == "MISMATCH":
                    bad += 1
                where = (f"{side.parent.name}/{side.name}" if side is not None
                         else "(no reference image)")
                line = f"    t{t.index:02d}  {content}  {pix}  [{tag}]  {where}"
                if tag == "MISMATCH":
                    warn(line.strip())
                    if missing:
                        warn(f"         missing from the flattened image: "
                             f"{missing[:8]}{' ...' if len(missing) > 8 else ''}")
                else:
                    print(line)
                    if missing:
                        print(f"         not found in the text layer (within "
                              f"tolerance, check the sheet if unsure): "
                              f"{missing[:6]}")
            info(f"visual check: {len(cmp_results)} side-by-side sheet(s) in "
                 f"{imgdir.name}/_compare/ (blue = as typeset in your document, "
                 f"orange = the flattened PNG) — open them to confirm nothing "
                 f"was mangled")
            if bad:
                warn(f"visual check FAILED for {bad} table(s) — inspect the "
                     "flagged sheets above")
                rc = rc or 3
            else:
                info("VISUAL PASS — every flattened table matches its "
                     "in-document rendering")

    skipped = [t for t in targets if not t.ok]
    for t in skipped:
        warn(f"t{t.index:02d} was NOT flattened: {t.note}")

    if not args.keep_build:
        shutil.rmtree(build, ignore_errors=True)
    else:
        info(f"build artifacts kept in {build}")

    info(f"done: {out_tex.name} + {deliver_pdf.name} + "
         f"{sum(len(t.images) for t in done)} PNG(s) in {imgdir_name}/")
    return rc


# --------------------------------------------------------------------------
def main(argv=None):
    ap = argparse.ArgumentParser(
        prog="tablepngs",
        description="Flatten LaTeX tables into high-res PNGs so PDF->Word "
                    "conversion leaves them intact.",
        epilog="Typical use:  python3 tablepngs.py main.tex "
               "   ->  main_tablepngs.tex, main_tablepngs.pdf, main_tablepngs/*.png",
    )
    ap.add_argument("texfile", nargs="?", help="main .tex document")
    ap.add_argument("--engine", default="auto",
                    choices=["auto", "pdflatex", "xelatex", "lualatex"],
                    help="LaTeX engine (default: auto-detect from magic "
                         "comments / fontspec)")
    ap.add_argument("--dpi", type=int, default=300,
                    help="PNG resolution (default 300; use 600 for print)")
    ap.add_argument("--suffix", default="_tablepngs",
                    help="suffix for the output .tex/.pdf (default _tablepngs)")
    ap.add_argument("--imgdir", default=None,
                    help="image folder name (default <stem>_tablepngs)")
    ap.add_argument("--max-height", default="0.85", metavar="FRAC",
                    help="cap table-float image height at FRAC of \\textheight "
                         "(default 0.85)")
    ap.add_argument("--page-height", default="0.96", metavar="FRAC",
                    help="cap multi-page-table chunk height at FRAC of "
                         "\\textheight (default 0.96; these chunks were "
                         "typeset to fill the page, so shrinking them makes "
                         "the table smaller than the original)")
    ap.add_argument("--only", default=None, metavar="N,M",
                    help="process only these table indices (see --list)")
    ap.add_argument("--skip", default=None, metavar="N,M",
                    help="skip these table indices")
    ap.add_argument("--no-bare", action="store_true",
                    help="ignore tabulars that are not inside a float/longtable")
    ap.add_argument("--list", action="store_true",
                    help="list detected tables and exit (dry run)")
    ap.add_argument("--no-verify", action="store_true",
                    help="skip the text-layer verification pass")
    ap.add_argument("--compare", action="store_true",
                    help="visual check: re-render every table in its real "
                         "document context and compare it against the "
                         "flattened PNG, writing side-by-side sheets to "
                         "<imgdir>/_compare/ (needs ImageMagick)")
    ap.add_argument("--compare-threshold", type=float, default=0.06,
                    metavar="X",
                    help="normalized RMSE above which --compare flags a table "
                         "as possibly mangled (default 0.06)")
    ap.add_argument("--shell-escape", action="store_true",
                    help="pass -shell-escape to the engine")
    ap.add_argument("--keep-build", action="store_true",
                    help="keep intermediate build files for debugging")
    ap.add_argument("--timeout", type=int, default=300,
                    help="per-compile timeout in seconds (default 300)")
    ap.add_argument("--check", action="store_true",
                    help="check dependencies and print install instructions")
    ap.add_argument("--version", action="version",
                    version=f"tablepngs {__version__}")
    args = ap.parse_args(argv)

    if args.check:
        sys.exit(doctor())
    if not args.texfile:
        ap.error("texfile is required (or use --check)")
    try:
        sys.exit(process(args))
    except TPError as e:
        die(str(e))


if __name__ == "__main__":
    main()
