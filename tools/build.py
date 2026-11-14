# -*- coding: utf-8 -*-
"""
Pravi instalacione pakete u korenu repozitorijuma:

  translit.plugin        – ONLYOFFICE (sadržaj foldera translit/)
  translit.oxt           – LibreOffice (sadržaj foldera libreoffice/)
  translit-msoffice.zip  – MS Office (folder msoffice/ sa install.cmd)

    python tools/build.py
"""
import os
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SKIP_DIRS = {"__pycache__", ".git"}
SKIP_EXT = {".pyc", ".pyo"}


def pack(src_dir, out_name, prefix=""):
    src = os.path.join(ROOT, src_dir)
    out = os.path.join(ROOT, out_name)
    count = 0
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for base, dirs, files in os.walk(src):
            dirs[:] = sorted(d for d in dirs if d not in SKIP_DIRS)
            for name in sorted(files):
                if os.path.splitext(name)[1] in SKIP_EXT:
                    continue
                path = os.path.join(base, name)
                z.write(path, prefix + os.path.relpath(path, src).replace(os.sep, "/"))
                count += 1
    print("%-22s %3d fajlova, %6d bajtova" % (out_name, count, os.path.getsize(out)))


def check_vba_sources():
    """VBA uvozi samo ASCII sa CRLF krajevima redova."""
    src = os.path.join(ROOT, "msoffice", "src")
    for name in sorted(os.listdir(src)):
        raw = open(os.path.join(src, name), "rb").read()
        if any(b > 127 for b in raw) or raw.count(b"\n") != raw.count(b"\r\n"):
            raise SystemExit("%s mora biti ASCII sa CRLF (pokrenite tests/vba_syntax_test.py)" % name)


if __name__ == "__main__":
    pack("translit", "translit.plugin")
    pack("libreoffice", "translit.oxt")
    check_vba_sources()
    pack("msoffice", "translit-msoffice.zip", prefix="translit-msoffice/")
