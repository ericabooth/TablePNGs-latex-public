#!/usr/bin/env python3
"""
Test harness for tablepngs: runs every example in examples/ on every engine
its manifest declares, and checks that
  1. tablepngs exits 0 (which already implies the built-in text-layer
     verification passed),
  2. the expected number of table targets was found,
  3. every target was actually flattened (none skipped),
  4. the flattened .tex, .pdf and PNGs exist.

It also runs the negative control in tests/fixtures/mangle_control, which is
built to defeat naive flattening: that case MUST fail the visual check. If it
ever passes, the visual verification has stopped working.

Usage:
  python3 tests/run_tests.py                # full matrix
  python3 tests/run_tests.py 03 08          # only examples whose dir starts 03/08
  python3 tests/run_tests.py --engine xelatex
  python3 tests/run_tests.py --keep         # keep generated outputs for inspection
  python3 tests/run_tests.py --no-compare   # skip the (slower) visual pass
"""
import argparse
import json
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPT = ROOT / "tablepngs.py"
EXAMPLES = ROOT / "examples"
TEXBIN = shutil.which("pdflatex")
BIBTEX = shutil.which("bibtex")

CLEAN_PATTERNS = [
    "main_tablepngs", 
    "main_tablepngs*", "*.lot", "*.aux", "*.log", "*.out", "*.toc", "*.lot", "*.lof",
    "*.bbl", "*.blg", "*.fls", "*.fdb_latexmk", "main.pdf", "*.synctex.gz",
]


def clean(d):
    for pat in CLEAN_PATTERNS:
        for p in d.glob(pat):
            if p.is_dir():
                shutil.rmtree(p, ignore_errors=True)
            else:
                p.unlink(missing_ok=True)


def sh(cmd, cwd, timeout=600):
    p = subprocess.run(cmd, cwd=cwd, timeout=timeout,
                       stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    return p.returncode, p.stdout.decode("utf-8", errors="replace")


def precompile_bibtex(d, engine):
    """For needs_bibtex examples: build main.aux/main.bbl in-place so
    tablepngs can resolve \\cite via the aux trick."""
    for cmd in ([engine, "-interaction=nonstopmode", "main.tex"],
                [BIBTEX, "main"],
                [engine, "-interaction=nonstopmode", "main.tex"],
                [engine, "-interaction=nonstopmode", "main.tex"]):
        rc, out = sh(cmd, d)
        if rc != 0 and cmd[0] != BIBTEX:
            return False, out[-1500:]
    return True, ""


def run_case(d, engine, manifest, keep, compare=True):
    clean(d)
    try:
        if manifest.get("needs_bibtex"):
            ok, log = precompile_bibtex(d, engine)
            if not ok:
                return "PRECOMPILE-FAIL", log
        cmd = [sys.executable, str(SCRIPT), "main.tex", "--engine", engine]
        if compare:
            cmd.append("--compare")
        t0 = time.time()
        rc, out = sh(cmd, d)
        dt = time.time() - t0
        problems = []
        if rc != 0:
            problems.append(f"exit {rc}")
        m = re.search(r"found (\d+) table target", out)
        found = int(m.group(1)) if m else -1
        if found != manifest["expected_tables"]:
            problems.append(f"found {found} targets, expected {manifest['expected_tables']}")
        if "was NOT flattened" in out:
            problems.append("some targets skipped")
        if "VERIFY PASS" not in out:
            problems.append("no VERIFY PASS")
        if compare and "VISUAL PASS" not in out:
            problems.append("no VISUAL PASS")
        if not (d / "main_tablepngs.tex").exists() or not (d / "main_tablepngs.pdf").exists():
            problems.append("missing output tex/pdf")
        pngs = list((d / "main_tablepngs").glob("*.png")) if (d / "main_tablepngs").exists() else []
        if len(pngs) < manifest["expected_tables"]:
            problems.append(f"only {len(pngs)} PNGs")
        if problems:
            return "FAIL: " + "; ".join(problems), out[-3000:]
        return f"PASS ({dt:.0f}s, {len(pngs)} png)", ""
    finally:
        if not keep:
            clean(d)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("filters", nargs="*", help="example dir prefixes to run")
    ap.add_argument("--engine", default=None)
    ap.add_argument("--keep", action="store_true")
    ap.add_argument("--no-compare", action="store_true",
                    help="skip the visual comparison pass (faster)")
    args = ap.parse_args()
    compare = not args.no_compare

    cases = []
    for mf in sorted(EXAMPLES.glob("*/manifest.json")):
        manifest = json.loads(mf.read_text())
        d = mf.parent
        if args.filters and not any(d.name.startswith(f) for f in args.filters):
            continue
        for engine in manifest["engines"]:
            if args.engine and engine != args.engine:
                continue
            cases.append((d, engine, manifest))

    if not cases:
        print("no cases matched")
        return 1

    results, failed = [], []

    # negative control: this fixture MUST fail the visual check
    ctrl = ROOT / "tests" / "fixtures" / "mangle_control"
    if compare and ctrl.is_dir() and not args.filters:
        print("... mangle_control (negative control)", flush=True)
        clean(ctrl)
        rc, out = sh([sys.executable, str(SCRIPT), "main.tex", "--compare"], ctrl)
        caught = rc == 3 and "MISMATCH" in out and "VISUAL PASS" not in out
        status = ("PASS (mangle correctly caught)" if caught else
                  f"FAIL: negative control NOT caught (rc={rc}) — "
                  "the visual check has lost its teeth")
        results.append(("mangle_control × pdflatex", status))
        if not caught:
            failed.append(("mangle_control", status, out[-2000:]))
        print(f"    {status}", flush=True)
        if not args.keep:
            clean(ctrl)

    for d, engine, manifest in cases:
        label = f"{d.name} × {engine}"
        print(f"... {label}", flush=True)
        status, detail = run_case(d, engine, manifest, args.keep, compare)
        results.append((label, status))
        if not status.startswith("PASS"):
            failed.append((label, status, detail))
        print(f"    {status}", flush=True)

    print("\n" + "=" * 64)
    print(f"{'case':<40} result")
    print("-" * 64)
    for label, status in results:
        print(f"{label:<40} {status}")
    print("=" * 64)
    npass = sum(1 for _, s in results if s.startswith("PASS"))
    print(f"{npass}/{len(results)} passed")
    for label, status, detail in failed:
        print(f"\n--- {label}: {status} ---\n{detail}")
    return 0 if not failed else 1


if __name__ == "__main__":
    sys.exit(main())
