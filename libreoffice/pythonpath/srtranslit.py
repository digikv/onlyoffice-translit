# -*- coding: utf-8 -*-
"""
Srpska transliteracija latinica <-> ćirilica i stari YUSCII.

Nezavisno od LibreOffice-a. Pravila su ista kao u translit/scripts/translit.js
(ONLYOFFICE dodatak) – kad menjate jedno, promenite i drugo
(tools/parity_test.py proverava da daju isti rezultat).

Funkcije *_pieces vraćaju listu (dužina_izvora, izlaz): svaki deo kaže koliko
znakova ulaza je zamenjeno kojim tekstom. Tako LibreOffice deo može da menja
tekst segment po segment i da sačuva formatiranje.
"""
import re
import unicodedata

LAT2CYR = {
    "a": "а", "b": "б", "c": "ц", "č": "ч", "ć": "ћ", "d": "д", "đ": "ђ",
    "e": "е", "f": "ф", "g": "г", "h": "х", "i": "и", "j": "ј", "k": "к",
    "l": "л", "m": "м", "n": "н", "o": "о", "p": "п", "r": "р", "s": "с",
    "š": "ш", "t": "т", "u": "у", "v": "в", "z": "з", "ž": "ж",
}

DIGRAPHS = {"lj": "љ", "nj": "њ", "dž": "џ"}

CYR2LAT = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "ђ": "đ", "е": "e",
    "ж": "ž", "з": "z", "и": "i", "ј": "j", "к": "k", "л": "l", "љ": "lj",
    "м": "m", "н": "n", "њ": "nj", "о": "o", "п": "p", "р": "r", "с": "s",
    "т": "t", "ћ": "ć", "у": "u", "ф": "f", "х": "h", "ц": "c", "ч": "č",
    "џ": "dž", "ш": "š",
}

# Reči u kojima se slova digrafa izgovaraju odvojeno.
# "|" označava mesto razdvajanja, "^" da izuzetak važi samo na početku reči.
EXCEPTIONS = [
    # nj -> нј
    "in|jek", "in|junk", "kon|jug", "kon|junk", "kon|jukt", "van|jez", "van|jug",
    "tran|jug", "an|jon",
    # dž -> дж
    "^nad|žanj", "^nad|želj", "^nad|živ", "^nad|žnje", "^nad|žup",
    "^pod|žanr", "^pod|želud", "^pod|žup", "^pod|živ", "^pod|žnje",
    "^od|živ", "^od|žvaka",
]

# Kada je uključena opcija dj -> đ, ovi prefiksi ostaju "dj" (odjednom, podjednako...)
DJ_EXCEPTIONS = ["^od|j", "^nad|j", "^pod|j", "^pred|j", "^ispod|j", "^iznad|j"]

ROMAN_RE = re.compile(r"M{0,4}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})")
# Srpske reči koje izgledaju kao rimski brojevi kada su pisane verzalom
COLLIDING_ROMAN = ("MI", "VI", "LI", "DIV")
FOREIGN_RE = re.compile(r"[qwxyQWXY]")
LATIN_LETTER = "A-Za-zČčĆćĐđŠšŽžÀ-ÖØ-öø-ÿ"
CYR_LETTER = "Ѐ-ӿ"
WORD_RE = re.compile("[" + LATIN_LETTER + CYR_LETTER + "0-9_]+|[^" + LATIN_LETTER + CYR_LETTER + "0-9_]+")
SR_LATIN_RE = re.compile("[A-Za-zČčĆćĐđŠšŽž]")
SENTENCE_END_RE = re.compile("[.!?…:;\n]")
SPACE_ONLY_RE = re.compile("[   \t]+")
# \s iz JavaScript-a (Python-ov \s se malo razlikuje)
JS_SPACE = "\t\n\u000B\f\r    -     　﻿"

URL_RE = re.compile(
    r"(?:https?://|ftp://|www\.)[^" + JS_SPACE + r"<>\"]+"
    r"|[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}"
    r"|\b[A-Za-z0-9\-]+\.(?:com|org|net|rs|ba|me|hr|info|edu|gov|io)(?:/[^" + JS_SPACE + r"<>\"]*)?\b",
    re.ASCII)

# Merne jedinice (SI i uobičajene). Jednoslovne jedinice (m, s, g, l, h, t, V, A...)
# štite se samo iza broja, da se ne bi preskakali predlozi "s", "u", "a"...
UNITS = sorted([
    # dužina, površina, zapremina
    "nm", "µm", "μm", "mm", "cm", "dm", "m", "km", "ft", "ha",
    "ml", "cl", "dl", "l", "L", "hl", "cc",
    # masa
    "µg", "mg", "g", "dag", "kg", "t", "lb", "oz", "ct",
    # vreme, frekvencija
    "ns", "µs", "ms", "s", "min", "h", "Hz", "kHz", "MHz", "GHz", "THz", "rpm",
    # elektrika, energija, snaga
    "mA", "A", "kA", "mV", "V", "kV", "MV", "W", "kW", "MW", "GW", "mW",
    "Wh", "kWh", "MWh", "GWh", "VA", "kVA", "MVA", "VAr", "kVAr", "Ah", "mAh",
    "Ω", "kΩ", "MΩ", "F", "µF", "nF", "pF", "H", "mH", "C", "T", "Wb", "S",
    "J", "kJ", "MJ", "cal", "kcal", "eV", "keV", "MeV", "KS", "hp",
    # sila, pritisak, temperatura
    "N", "kN", "MN", "Pa", "hPa", "kPa", "MPa", "GPa", "bar", "mbar", "atm", "psi", "K",
    # svetlo, zvuk, hemija, računari
    "cd", "lm", "lx", "dB", "dBm", "mol", "mmol", "ppm", "pH",
    "b", "B", "kb", "kB", "KB", "Mb", "MB", "Gb", "GB", "TB", "PB",
    "bps", "kbps", "Mbps", "Gbps", "dpi", "px", "pt",
], key=len, reverse=True)
UNIT = "(?:" + "|".join(re.escape(u) for u in UNITS) + ")"
UNIT_EXPR = UNIT + "(?:[²³23])?(?:[" + JS_SPACE + "]?[/·⋅*][" + JS_SPACE + "]?" + UNIT + "(?:[²³23])?)*"
LETTER_CLASS = "[" + LATIN_LETTER + CYR_LETTER + "µμΩ]"
UNITS_RE = re.compile(
    # iza broja: "5 km", "10kg", "3,5 m²", "100 km/h"
    "[0-9](?:[" + JS_SPACE + "]| | )?" + UNIT_EXPR + "(?!" + LETTER_CLASS + ")"
    # složene jedinice bilo gde: "km/h", "m/s²", "kWh/m2"
    "|(?<!" + LETTER_CLASS + ")" + UNIT + "(?:[²³23])?/" + UNIT + "(?:[²³23])?(?!" + LETTER_CLASS + ")"
    # stepeni: "°C", "°F"
    "|°[" + JS_SPACE + "]?[CFK](?!" + LETTER_CLASS + ")")


def build_exceptions_re(lst):
    """Korisnički izuzeci: jedna reč ili izraz po redu, "*" na kraju = bilo koji nastavak reči."""
    if not lst:
        return None
    if isinstance(lst, str):
        lst = re.split(r"\r?\n", lst)
    parts = []
    for s in lst:
        s = s.strip()
        if not s or s == "*":
            continue
        star = s.endswith("*")
        if star:
            s = s[:-1]
        parts.append(re.escape(s) + (LETTER_CLASS + "*" if star else ""))
    if not parts:
        return None
    parts.sort(key=len, reverse=True)
    return re.compile("(?<!" + LETTER_CLASS + ")(?:" + "|".join(parts) + ")(?!" + LETTER_CLASS + ")",
                      re.IGNORECASE)


def is_upper(ch):
    return ch != ch.lower() and ch == ch.upper()


def is_letter(ch):
    return bool(ch) and ch.lower() != ch.upper()


def is_roman_form(w):
    return bool(w) and ROMAN_RE.fullmatch(w) is not None


def split_positions(lower_word, lst):
    found = set()
    for pat in lst:
        anchored = pat.startswith("^")
        if anchored:
            pat = pat[1:]
        k = pat.index("|")
        plain = pat.replace("|", "")
        if anchored:
            if lower_word.startswith(plain):
                found.add(k)
        else:
            pos = lower_word.find(plain)
            while pos != -1:
                found.add(pos + k)
                pos = lower_word.find(plain, pos + 1)
    return found


def apply_case(src, dst):
    return dst.upper() if is_upper(src) else dst


# ---------- rimski brojevi u kontekstu ----------

def prev_word(tokens, i):
    if i < 2 or SENTENCE_END_RE.search(tokens[i - 1]):
        return None
    return tokens[i - 2]


def next_word(tokens, i):
    if i + 2 >= len(tokens) or not SPACE_ONLY_RE.fullmatch(tokens[i + 1]):
        return None
    return tokens[i + 2]


def is_caps_word(w):
    return (bool(w) and len(w) >= 2 and not re.search("[0-9_]", w)
            and w == w.upper() and w != w.lower() and not is_roman_form(w))


def is_roman(tokens, i):
    """Rimski broj, ili reč koja samo tako izgleda ("I tako...", "MAMA I TATA", "MI SMO")?"""
    w = tokens[i]
    if not is_roman_form(w):
        return False
    nxt = next_word(tokens, i)
    if w == "I":
        if not nxt or not is_letter(nxt[0]):
            return True
        if nxt.lower() in ("i", "ili", "do") and is_roman_form(next_word(tokens, i + 2)):
            return True
        prev = prev_word(tokens, i)
        return not (prev is None or is_caps_word(prev))
    if w in COLLIDING_ROMAN:
        return not is_caps_word(nxt)
    return True


# ---------- latinica -> ćirilica ----------

def word_to_cyrillic(tokens, i, opts):
    """Vraća listu delova (dužina_izvora, izlaz) za token tokens[i]."""
    word = tokens[i]
    if (opts["skipForeign"] and FOREIGN_RE.search(word)) \
            or (opts["skipRoman"] and is_roman(tokens, i)) \
            or not SR_LATIN_RE.search(word):
        return [(len(word), word)]

    lower = word.lower()
    splits = split_positions(lower, EXCEPTIONS)
    dj_splits = split_positions(lower, DJ_EXCEPTIONS) if opts["djToDj"] else set()
    out = []
    i = 0
    n = len(word)
    while i < n:
        ch = word[i]
        pair = lower[i:i + 2]
        if pair in DIGRAPHS and (i + 1) not in splits:
            out.append((2, apply_case(ch, DIGRAPHS[pair])))
            i += 2
            continue
        if opts["djToDj"] and pair == "dj" and (i + 1) not in dj_splits:
            out.append((2, apply_case(ch, "ђ")))
            i += 2
            continue
        lc = lower[i]
        out.append((1, apply_case(ch, LAT2CYR[lc]) if lc in LAT2CYR else ch))
        i += 1
    return out


def cyr_chunk(opts):
    def fn(chunk):
        tokens = WORD_RE.findall(chunk)
        out = []
        for i in range(len(tokens)):
            out.extend(word_to_cyrillic(tokens, i, opts))
        return out
    return fn


# ---------- ćirilica -> latinica ----------

def latin_chunk(text):
    out = []
    for i, ch in enumerate(text):
        lc = ch.lower()
        lat = CYR2LAT.get(lc)
        if not lat:
            out.append((1, ch))
        elif not is_upper(ch):
            out.append((1, lat))
        elif len(lat) == 1:
            out.append((1, lat.upper()))
        else:
            # Љ, Њ, Џ: "LJ" u reči pisanoj verzalom, inače "Lj"
            nxt = text[i + 1] if i + 1 < len(text) else ""
            prev = text[i - 1] if i > 0 else ""
            all_caps = is_upper(nxt) if is_letter(nxt) else (is_letter(prev) and is_upper(prev))
            out.append((1, lat.upper() if all_caps else lat[0].upper() + lat[1:]))
    return out


# ---------- YUSCII (JUS I.B1.002 latinica, JUS I.B1.003 ćirilica) ----------

YU_LAT = {
    "@": "Ž", "`": "ž", "[": "Š", "{": "š", "\\": "Đ",
    "|": "đ", "]": "Ć", "}": "ć", "^": "Č", "~": "č",
}
YU_CYR = {
    "A": "А", "B": "Б", "C": "Ц", "D": "Д", "E": "Е", "F": "Ф", "G": "Г", "H": "Х",
    "I": "И", "J": "Ј", "K": "К", "L": "Л", "M": "М", "N": "Н", "O": "О", "P": "П",
    "Q": "Љ", "R": "Р", "S": "С", "T": "Т", "U": "У", "V": "В", "W": "Њ", "X": "Џ",
    "Z": "З", "@": "Ж", "[": "Ш", "\\": "Ђ", "]": "Ћ", "^": "Ч",
}
for _k in list(YU_CYR):
    _lk = {"@": "`", "[": "{", "\\": "|", "]": "}", "^": "~"}.get(_k, _k.lower())
    YU_CYR[_lk] = YU_CYR[_k].lower()
YU_SYMBOLS = set("@`[]{}\\|^~")


def _in_word(text, i):
    def wordish(c):
        return is_letter(c) or c in YU_SYMBOLS

    def at(k):
        return text[k] if 0 <= k < len(text) else ""
    # gleda i preko susednih YUSCII znakova ("[}" i sl.)
    j = i - 1
    while j >= 0 and text[j] in YU_SYMBOLS:
        j -= 1
    k = i + 1
    while k < len(text) and text[k] in YU_SYMBOLS:
        k += 1
    return (is_letter(at(j)) or is_letter(at(k))
            or (wordish(at(i - 1)) and wordish(at(i + 1))))


def yuscii_chunk(table):
    # Znak YUSCII-ja (npr. "[" ili "@") menja se samo kada je deo reči,
    # da bi zagrade, "@" u tekstu i slično ostali netaknuti.
    def fn(text):
        out = []
        for i, ch in enumerate(text):
            m = table.get(ch)
            if not m or (ch in YU_SYMBOLS and not _in_word(text, i)):
                out.append((1, ch))
            else:
                out.append((1, m))
        return out
    return fn


# ---------- zajedničko ----------

def map_unprotected(text, regexes, fn):
    """Delovi teksta koje hvata bilo koji od regexa ostaju netaknuti, ostatak prolazi kroz fn."""
    if not regexes:
        return fn(text) if text else []
    rx, rest = regexes[0], regexes[1:]
    out = []
    last = 0
    for m in rx.finditer(text):
        if m.end() == m.start():
            continue
        out.extend(map_unprotected(text[last:m.start()], rest, fn))
        out.append((m.end() - m.start(), m.group(0)))
        last = m.end()
    out.extend(map_unprotected(text[last:], rest, fn))
    return out


def protectors(opts, direction):
    lst = []
    if opts["exceptionsRe"] is not None:
        lst.append(opts["exceptionsRe"])
    if direction == "cyr":
        if opts["skipUrls"]:
            lst.append(URL_RE)
        if opts["skipUnits"]:
            lst.append(UNITS_RE)
    elif direction in ("yuLat", "yuCyr"):
        if opts["skipUrls"]:
            lst.append(URL_RE)
    return lst


_exceptions_cache = {"key": None, "re": None}


def normalize_opts(o=None):
    o = o or {}
    if o.get("_normalized"):
        return o
    exc = o.get("exceptions") or ""
    key = "\n".join(exc) if isinstance(exc, (list, tuple)) else exc
    if key != _exceptions_cache["key"]:
        _exceptions_cache["key"] = key
        _exceptions_cache["re"] = build_exceptions_re(key)
    return {
        "_normalized": True,
        "skipForeign": o.get("skipForeign", True) is not False,
        "skipRoman": o.get("skipRoman", True) is not False,
        "skipUrls": o.get("skipUrls", True) is not False,
        "skipUnits": o.get("skipUnits", True) is not False,
        "djToDj": bool(o.get("djToDj", False)),
        "exceptionsRe": _exceptions_cache["re"],
    }


MODES = ("cyr", "lat", "yuLat", "yuCyr")


def transliterate_pieces(text, mode, opts=None):
    """Delovi (dužina_izvora, izlaz) za tekst koji je već u NFC obliku."""
    opts = normalize_opts(opts)
    if not text:
        return []
    if mode == "cyr":
        fn = cyr_chunk(opts)
    elif mode == "lat":
        fn = latin_chunk
    elif mode == "yuLat":
        fn = yuscii_chunk(YU_LAT)
    elif mode == "yuCyr":
        fn = yuscii_chunk(YU_CYR)
    else:
        raise ValueError("Nepoznat smer: %r" % (mode,))
    return map_unprotected(text, protectors(opts, mode), fn)


def transliterate(text, mode, opts=None):
    if not text:
        return text
    if mode == "cyr":
        text = unicodedata.normalize("NFC", text)
    return "".join(out for _, out in transliterate_pieces(text, mode, opts))


def to_cyrillic(text, opts=None):
    return transliterate(text, "cyr", opts)


def to_latin(text, opts=None):
    return transliterate(text, "lat", opts)


def yuscii_to_latin(text, opts=None):
    return transliterate(text, "yuLat", opts)


def yuscii_to_cyrillic(text, opts=None):
    return transliterate(text, "yuCyr", opts)


# ---------- preslovljavanje sa mapom položaja (za tekst sa formatiranjem) ----------

_COMBINING_RE = re.compile("[̀-ͯ]")


def _nfc_with_offsets(text):
    """NFC oblik teksta i za svaki njegov položaj odgovarajući položaj u originalu.

    Spaja samo osnovno slovo sa kombinujućim znacima koji slede (c + ̌ -> č),
    pa je mapa uvek monotona."""
    if not _COMBINING_RE.search(text):
        return text, None
    parts = []
    offs = []
    i = 0
    n = len(text)
    while i < n:
        j = i + 1
        while j < n and _COMBINING_RE.match(text[j]):
            j += 1
        cluster = text[i:j]
        norm = unicodedata.normalize("NFC", cluster) if j - i > 1 else cluster
        for k in range(len(norm)):
            offs.append(i if k == 0 else j)
        parts.append(norm)
        i = j
    offs.append(n)
    return "".join(parts), offs


def transliterate_spans(text, mode, opts=None):
    """Lista (početak, kraj, izlaz) u koordinatama originalnog teksta, bez delova koji se ne menjaju."""
    if not text:
        return []
    if mode == "cyr":
        norm, offs = _nfc_with_offsets(text)
    else:
        norm, offs = text, None
    spans = []
    pos = 0
    for length, out in transliterate_pieces(norm, mode, opts):
        a, b = pos, pos + length
        pos = b
        if offs is not None:
            a, b = offs[a], offs[b]
        if text[a:b] != out:
            spans.append((a, b, out))
    return spans
