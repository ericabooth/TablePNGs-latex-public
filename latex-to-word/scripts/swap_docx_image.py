#!/usr/bin/env python3
"""Replace one image inside a .docx without rebuilding the document.

Use when a single figure changed and the recipient may already be editing the file.
Also corrects the drawing extent so Word does not stretch the new image to the old
aspect ratio.

Usage: swap_docx_image.py IN.docx OUT.docx word/media/rIdNN.png NEW_IMAGE.png
"""
import re, shutil, subprocess, sys, pathlib, tempfile, os

src, dst, target, newimg = sys.argv[1:5]

def png_size(p):
    d = pathlib.Path(p).read_bytes()[16:24]
    return int.from_bytes(d[:4], 'big'), int.from_bytes(d[4:], 'big')

work = tempfile.mkdtemp()
subprocess.run(['unzip', '-q', src, '-d', work], check=True)

old = os.path.join(work, target)
ow, oh = png_size(old)
nw, nh = png_size(newimg)
print(f'old {ow}x{oh} (aspect {ow/oh:.4f}) -> new {nw}x{nh} (aspect {nw/nh:.4f})')
shutil.copy(newimg, old)

rid = pathlib.Path(target).stem                      # e.g. rId47
doc = os.path.join(work, 'word', 'document.xml')
d = pathlib.Path(doc).read_text(encoding='utf8')
i = d.find(rid)
if i < 0:
    sys.exit(f'{rid} not referenced in document.xml')
start = max(0, i - 1500)
window = d[start:i + 300]
m = re.search(r'cx="(\d+)" cy="(\d+)"', window)
cx, cy = int(m.group(1)), int(m.group(2))
new_cy = round(cx * nh / nw)
print(f'extent {cx}x{cy} -> {cx}x{new_cy} EMU')
d = d[:start] + window.replace(f'cx="{cx}" cy="{cy}"', f'cx="{cx}" cy="{new_cy}"') + d[start + len(window):]
pathlib.Path(doc).write_text(d, encoding='utf8')

out = os.path.abspath(dst)
if os.path.exists(out):
    os.remove(out)
subprocess.run(['zip', '-q', '-r', '-X', out, '.'], cwd=work, check=True)
print('wrote', out)
