# -*- coding: utf-8 -*-
"""
Testira VBA jezgro (msoffice/src/TranslitCore.bas) bez Office-a: modul se pretvori
u VBScript (uklone se "As Tip" deklaracije) i izvrši preko cscript.exe, pa se
rezultati porede sa translit.js (ONLYOFFICE) i srtranslit.py (LibreOffice).

    python tests/vba_core_test.py        (potrebni Windows i Node.js)
"""
import os
import random
import re
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(ROOT, "libreoffice", "pythonpath"))
import srtranslit  # noqa: E402
from parity_test import EXPECTED, OPTION_SETS, ALPHABET, SNIPPETS, run_js  # noqa: E402

BAS = os.path.join(ROOT, "msoffice", "src", "TranslitCore.bas")
TYPE_RE = re.compile(r"\s+As\s+(?:String|Long|Boolean|Object|Variant|Integer|Double)\b(\(\))?")

HARNESS = r'''
' --- zamena za TranslitNative.TrNormalizeCluster (VBScript ne moze da zove Windows API):
' samo slova koja test koristi
Function TrNormalizeCluster(s)
    Dim m
    TrNormalizeCluster = s
    For Each m In Array("c30C:10D", "s30C:161", "z30C:17E", "C30C:10C", "S30C:160", "Z30C:17D", "c301:107", "C301:106")
        If s = Left(m, 1) & ChrW(CLng("&H" & Mid(m, 2, 3))) Then TrNormalizeCluster = ChrW(CLng("&H" & Mid(m, 6)))
    Next
End Function

Function FromHex(s)
    Dim parts, i, r
    r = ""
    If s <> "" Then
        parts = Split(s, " ")
        For i = 0 To UBound(parts)
            r = r & ChrW(CLng("&H" & parts(i)))
        Next
    End If
    FromHex = r
End Function

Function ToHex(s)
    Dim i, r, c
    r = ""
    For i = 1 To Len(s)
        c = AscW(Mid(s, i, 1))
        If c < 0 Then c = c + 65536
        If i > 1 Then r = r & " "
        r = r & Hex(c)
    Next
    ToHex = r
End Function

Dim fso, inp, outp, line, f, opts, n, i, sp, text
Set fso = CreateObject("Scripting.FileSystemObject")
Set inp = fso.OpenTextFile(WScript.Arguments(0), 1)
Set outp = fso.CreateTextFile(WScript.Arguments(1), True)
Set opts = CreateObject("Scripting.Dictionary")
Do Until inp.AtEndOfStream
    line = inp.ReadLine
    f = Split(line, vbTab)
    If f(0) = "OPT" Then
        opts.Add f(1), f
    Else
        sp = opts(f(1))
        TrSetOptions sp(2) = "1", sp(3) = "1", sp(4) = "1", sp(5) = "1", sp(6) = "1", FromHex(sp(7))
        text = FromHex(f(2))
        n = TrSpans(text, f(0))
        line = ToHex(TrApplySpans(text)) & vbTab
        For i = 0 To n - 1
            line = line & TrA(i) & "," & TrB(i) & "," & ToHex(TrOut(i)) & ";"
        Next
        outp.WriteLine line
    End If
Loop
outp.Close
'''


def to_vbscript(src):
    out = []
    for line in src.replace("\r\n", "\n").split("\n"):
        if line.startswith("Attribute "):
            continue
        out.append(TYPE_RE.sub(lambda m: m.group(1) or "", line))
    return "\n".join(out) + "\n" + HARNESS


def hexs(s):
    return " ".join("%X" % ord(c) for c in s)


def unhex(h):
    return "".join(chr(int(x, 16)) for x in h.split(" ") if x)


def opt_flags(o):
    n = srtranslit.normalize_opts(o)
    exc = o.get("exceptions") or ""
    return ["1" if n[k] else "0" for k in ("skipForeign", "skipRoman", "skipUrls", "skipUnits", "djToDj")] + [hexs(exc)]


def main():
    src = open(BAS, encoding="ascii").read()  # VBA modul mora biti čist ASCII
    rng = random.Random(77)
    alphabet = [c for c in ALPHABET if c not in "́̌"]
    option_sets = OPTION_SETS + [o for _, _, _, o in EXPECTED]
    cases = []
    for mode, text, _, o in EXPECTED:
        cases.append((mode, text, option_sets.index(o)))
    for extra in ("čas Šuma žena ćup", "Lj́x"):
        cases.append(("cyr", extra, 0))
    for _ in range(1500):
        parts = []
        for _ in range(rng.randint(1, 25)):
            parts.append(rng.choice(SNIPPETS) if rng.random() < 0.3 else rng.choice(alphabet))
        text = "".join(parts)
        for mode in srtranslit.MODES:
            cases.append((mode, text, rng.randrange(len(OPTION_SETS))))

    tmp = tempfile.mkdtemp(prefix="vbatest")
    vbs = os.path.join(tmp, "core.vbs")
    inp = os.path.join(tmp, "in.txt")
    outp = os.path.join(tmp, "out.txt")
    with open(vbs, "w", encoding="ascii", newline="\r\n") as f:
        f.write(to_vbscript(src))
    with open(inp, "w", encoding="ascii", newline="\r\n") as f:
        for i, o in enumerate(option_sets):
            f.write("\t".join(["OPT", str(i)] + opt_flags(o)) + "\n")
        for mode, text, oi in cases:
            f.write("%s\t%d\t%s\n" % (mode, oi, hexs(text)))
    res = subprocess.run(["cscript", "//nologo", vbs, inp, outp], stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    if res.returncode != 0 or not os.path.exists(outp):
        print(res.stdout.decode("cp1252", "replace"))
        sys.exit("VBScript nije uspeo")
    rows = open(outp, encoding="ascii").read().splitlines()

    js = run_js([[m, t, option_sets[oi]] for m, t, oi in cases])
    failures = 0
    for (mode, text, oi), row, js_out in zip(cases, rows, js):
        o = option_sets[oi]
        res_hex, span_str = row.split("\t")
        vba_out = unhex(res_hex)
        vba_spans = []
        for part in span_str.split(";"):
            if part:
                a, b, h = part.split(",")
                vba_spans.append((int(a), int(b), unhex(h)))
        py_spans = srtranslit.transliterate_spans(text, mode, o)
        want = srtranslit.transliterate(text, mode, o)
        problems = []
        if vba_out != want:
            problems.append("tekst: vba %r, python %r" % (vba_out, want))
        if vba_out != js_out and "̌" not in text and "́" not in text:
            problems.append("tekst: vba %r, js %r" % (vba_out, js_out))
        if vba_spans != py_spans:
            problems.append("izmene: vba %r, python %r" % (vba_spans, py_spans))
        if problems:
            failures += 1
            if failures <= 12:
                print("RAZLIKA [%s] %r opcije=%r\n  %s" % (mode, text, o, "\n  ".join(problems)))
    print("VBA jezgro: %d razlika u %d slučajeva" % (failures, len(cases)))
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
