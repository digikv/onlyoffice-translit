# -*- coding: utf-8 -*-
"""
Proverava da Python jezgro (LibreOffice) i JavaScript jezgro (ONLYOFFICE)
daju isti rezultat, i da su poznati primeri ispravni.

    python tests/parity_test.py          (potreban je Node.js)
"""
import json
import os
import random
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "libreoffice", "pythonpath"))
import srtranslit  # noqa: E402

JS = os.path.join(ROOT, "translit", "scripts", "translit.js")

# (smer, ulaz, očekivano, opcije)
EXPECTED = [
    ("cyr", "Ljubav LJUBAV ljubav", "Љубав ЉУБАВ љубав", {}),
    ("cyr", "Džep DŽEP Njegoš NJEGOŠ", "Џеп ЏЕП Његош ЊЕГОШ", {}),
    ("cyr", "injekcija, konjugacija, vanjezički, podžupan, nadživeti",
     "инјекција, конјугација, ванјезички, поджупан, надживети", {}),
    ("cyr", "5 km i 90 km/h, 22 °C, 64 GB, 20 m²", "5 km и 90 km/h, 22 °C, 64 GB, 20 m²", {}),
    ("cyr", "s njim u kući", "с њим у кући", {}),
    ("cyr", "5 km", "5 км", {"skipUnits": False}),
    ("cyr", "www.example.com i a@b.rs", "www.example.com и a@b.rs", {}),
    ("cyr", "Poglavlje XIV, godina MMXXVI", "Поглавље XIV, година MMXXVI", {}),
    ("cyr", "Windows i Excel, Word i Paint", "Windows и Excel, Word и Паинт", {}),
    ("cyr", "Microsoftov Word", "Microsoftov Word", {"exceptions": "Microsoft*\nWord"}),
    ("cyr", "odjednom Djordje", "одједном Ђорђе", {"djToDj": True}),
    ("cyr", "Djordje", "Дјордје", {}),
    # "I" kao veznik, a ne rimski broj
    ("cyr", "I tako je bilo.", "И тако је било.", {}),
    ("cyr", "Kraj. I onda", "Крај. И онда", {}),
    ("cyr", "„I tako“", "„И тако“", {}),
    ("cyr", "MAMA I TATA", "МАМА И ТАТА", {}),
    ("cyr", "Petar I Karađorđević", "Петар I Карађорђевић", {}),
    ("cyr", "I. Uvod", "I. Увод", {}),
    ("cyr", "I i II kategorija", "I и II категорија", {}),
    ("cyr", "klasa I", "класа I", {}),
    ("cyr", "I", "I", {}),
    ("cyr", "MI SMO DOŠLI, VI STE OTIŠLI, DA LI JE", "МИ СМО ДОШЛИ, ВИ СТЕ ОТИШЛИ, ДА ЛИ ЈЕ", {}),
    ("cyr", "GLAVA VI", "ГЛАВА VI", {}),
    ("cyr", "XX vek", "XX век", {}),
    ("cyr", "I tako", "I тако", {"skipRoman": True, "exceptions": "I"}),
    ("cyr", "čaj", "чај", {}),
    ("lat", "Љубав ЉУБАВ Џеп ЏЕП Његош", "Ljubav LJUBAV Džep DŽEP Njegoš", {}),
    ("lat", "Њ је слово, ЊЕ", "Nj je slovo, NJE", {}),
    ("yuLat", "~a{a [uma @aba \\or|e ]uk ^ep", "čaša Šuma Žaba Đorđe Ćuk Čep", {}),
    ("yuLat", "niz [1] i (x) @ ^", "niz [1] i (x) @ ^", {}),
    ("yuCyr", "Qubav Wegox Xep ~a{a", "Љубав Његоџ Џеп чаша", {}),
]

OPTION_SETS = [
    {},
    {"skipForeign": False, "skipRoman": False, "skipUrls": False, "skipUnits": False},
    {"djToDj": True},
    {"exceptions": "Microsoft*\nWi-Fi\nOnly Office\nkm*"},
]

ALPHABET = (list("abcčćddđefghijklmnoprsštuvzžqwxy") + list("ABCČĆDĐEFGHIJKLMNOPRSŠTUVZŽQWXY")
            + list("абвгдђежзијклљмнњопрстћуфхцчџшАБВГДЂЕЖЗИЈКЛЉМНЊОПРСТЋУФХЦЧЏШ")
            + list("0123456789") + [" "] * 8 + list(".,;:!?-–()[]{}@`\\|^~/°²„“\"'\n\t")
            + [" ", "́", "̌", "µ", "Ω"])
SNIPPETS = ["lj", "nj", "dž", "Lj", "NJ", "dj", "km/h", " m²", "5 km", "°C", "kWh", " I ", "MI ", "VI ",
            "LI ", "XIV", "www.a.rs", "a@b.com", "http://x.y/z", "injekc", "podžup", "odj", "Microsoft",
            "Wi-Fi", ". I tako", "MAMA I TATA", "I i II", "3,5 kg", "64 GB", "10 MPa", "Only Office"]


def random_text(rng):
    parts = []
    for _ in range(rng.randint(1, 25)):
        if rng.random() < 0.3:
            parts.append(rng.choice(SNIPPETS))
        else:
            parts.append(rng.choice(ALPHABET))
    return "".join(parts)


def run_js(cases):
    script = r"""
const T = require(process.argv[1]);
const fns = {cyr: T.toCyrillic, lat: T.toLatin, yuLat: T.yusciiToLatin, yuCyr: T.yusciiToCyrillic};
let data = "";
process.stdin.setEncoding("utf8");
process.stdin.on("data", d => data += d);
process.stdin.on("end", () => {
  const cases = JSON.parse(data);
  process.stdout.write(JSON.stringify(cases.map(c => fns[c[0]](c[1], c[2]))));
});
"""
    res = subprocess.run(["node", "-e", script, JS], input=json.dumps(cases).encode("utf-8"),
                         stdout=subprocess.PIPE, check=True)
    return json.loads(res.stdout.decode("utf-8"))


def main():
    failures = 0
    for mode, src, want, opts in EXPECTED:
        got = srtranslit.transliterate(src, mode, opts)
        if got != want:
            failures += 1
            print("POGREŠNO [%s] %r\n  očekivano %r\n  dobijeno  %r" % (mode, src, want, got))

    rng = random.Random(20261113)
    cases = [[m, s, o] for m, s, _, o in EXPECTED]
    for _ in range(4000):
        text = random_text(rng)
        for mode in srtranslit.MODES:
            cases.append([mode, text, rng.choice(OPTION_SETS)])
    js = run_js(cases)
    mismatches = 0
    for (mode, text, opts), js_out in zip(cases, js):
        py_out = srtranslit.transliterate(text, mode, opts)
        if py_out != js_out:
            mismatches += 1
            if mismatches <= 15:
                print("RAZLIKA [%s] %r opcije=%r\n  js %r\n  py %r" % (mode, text, opts, js_out, py_out))

    # spans moraju dati isti tekst kao transliterate (za tekst bez kombinujućih znakova)
    span_errors = 0
    for mode, text, opts in cases:
        out = list(text)
        for a, b, rep in sorted(srtranslit.transliterate_spans(text, mode, opts), reverse=True):
            out[a:b] = [rep]
        rebuilt = "".join(out)
        want = srtranslit.transliterate(text, mode, opts)
        if rebuilt != want:
            import unicodedata
            if unicodedata.normalize("NFC", rebuilt) != unicodedata.normalize("NFC", want):
                span_errors += 1
                if span_errors <= 5:
                    print("SPANS [%s] %r\n  %r\n  %r" % (mode, text, rebuilt, want))

    print("primeri: %d grešaka; poređenje JS/Python: %d razlika u %d slučajeva; spans: %d grešaka"
          % (failures, mismatches, len(cases), span_errors))
    sys.exit(1 if failures or mismatches or span_errors else 0)


if __name__ == "__main__":
    main()
