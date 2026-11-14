# -*- coding: utf-8 -*-
"""
Provera sintakse VBA modula bez Office-a: svaki modul se prevede u VBScript
(uklone se tipovi, oznake i "On Error GoTo"), pa ga cscript.exe samo učita.
VBScript prijavljuje sintaksne greške (nezatvoren If/For/Do/Select, pogrešan
red...) pre izvršavanja; imena iz Word/Excel/PowerPoint biblioteka se ne proveravaju.

Proverava i da su moduli čist ASCII sa CRLF krajevima redova (to VBA uvoz traži).

    python tests/vba_syntax_test.py
"""
import glob
import os
import re
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "msoffice", "src")

TYPE_RE = re.compile(r"\s+As\s+(?:New\s+)?[A-Za-z_][A-Za-z0-9_.]*(\(\))?")
LABEL_RE = re.compile(r"^\s*[A-Za-z_]\w*:\s*$")


def to_vbscript(src):
    out = []
    skip = False
    for line in src.split("\r\n"):
        s = line.strip()
        if s.startswith("Attribute ") or s.startswith("#"):
            skip = s.startswith("#If") or (skip and not s.startswith("#End If"))
            if s.startswith("#End If"):
                skip = False
            continue
        if skip or "Declare " in s:
            continue
        if LABEL_RE.match(line):
            out.append("")
            continue
        line = re.sub(r"On Error GoTo (?!0\b)\w+", "On Error Resume Next", line)
        line = re.sub(r"^(\s*)Resume \w+\s*$", r"\1' resume", line)
        line = TYPE_RE.sub(lambda m: m.group(1) or "", line)
        out.append(line)
    return "\r\n".join(out)


def main():
    failures = 0
    tmp = tempfile.mkdtemp(prefix="vbasyntax")
    for path in sorted(glob.glob(os.path.join(SRC, "*.bas"))):
        name = os.path.basename(path)
        raw = open(path, "rb").read()
        if any(b > 127 for b in raw):
            print("GREŠKA %s: nije čist ASCII" % name)
            failures += 1
            continue
        if raw.count(b"\n") != raw.count(b"\r\n"):
            print("GREŠKA %s: krajevi redova nisu CRLF" % name)
            failures += 1
            continue
        if not raw.startswith(b'Attribute VB_Name = "%s"' % os.path.splitext(name)[0].encode()):
            print("GREŠKA %s: prvi red mora biti Attribute VB_Name" % name)
            failures += 1
            continue
        vbs = os.path.join(tmp, name + ".vbs")
        with open(vbs, "w", encoding="ascii", newline="") as f:
            f.write(to_vbscript(raw.decode("ascii")))
        res = subprocess.run(["cscript", "//nologo", vbs], stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        text = res.stdout.decode("cp1252", "replace").strip()
        if res.returncode != 0 or "error" in text.lower():
            failures += 1
            print("GREŠKA %s:\n  %s" % (name, text.replace("\n", "\n  ")))
        else:
            print("ok     %s" % name)
    print("\n%s" % ("SVE U REDU" if not failures else "%d modula sa greškom" % failures))
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
