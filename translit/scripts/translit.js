/*
 * Srpska transliteracija latinica <-> ćirilica
 * (po uzoru na ooOTranslit za LibreOffice/OpenOffice)
 */
(function (root) {
    "use strict";

    var LAT2CYR = {
        "a": "а", "b": "б", "c": "ц", "č": "ч", "ć": "ћ", "d": "д", "đ": "ђ",
        "e": "е", "f": "ф", "g": "г", "h": "х", "i": "и", "j": "ј", "k": "к",
        "l": "л", "m": "м", "n": "н", "o": "о", "p": "п", "r": "р", "s": "с",
        "š": "ш", "t": "т", "u": "у", "v": "в", "z": "з", "ž": "ж"
    };

    var DIGRAPHS = { "lj": "љ", "nj": "њ", "dž": "џ" };

    var CYR2LAT = {
        "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "ђ": "đ", "е": "e",
        "ж": "ž", "з": "z", "и": "i", "ј": "j", "к": "k", "л": "l", "љ": "lj",
        "м": "m", "н": "n", "њ": "nj", "о": "o", "п": "p", "р": "r", "с": "s",
        "т": "t", "ћ": "ć", "у": "u", "ф": "f", "х": "h", "ц": "c", "ч": "č",
        "џ": "dž", "ш": "š"
    };

    // Reči u kojima se slova digrafa izgovaraju odvojeno.
    // "|" označava mesto razdvajanja, "^" da izuzetak važi samo na početku reči.
    var EXCEPTIONS = [
        // nj -> нј
        "in|jek", "in|junk", "kon|jug", "kon|junk", "kon|jukt", "van|jez", "van|jug",
        "tran|jug", "an|jon",
        // dž -> дж
        "^nad|žanj", "^nad|želj", "^nad|živ", "^nad|žnje", "^nad|žup",
        "^pod|žanr", "^pod|želud", "^pod|žup", "^pod|živ", "^pod|žnje",
        "^od|živ", "^od|žvaka"
    ];

    // Kada je uključena opcija dj -> đ, ovi prefiksi ostaju "dj" (odjednom, podjednako...)
    var DJ_EXCEPTIONS = ["^od|j", "^nad|j", "^pod|j", "^pred|j", "^ispod|j", "^iznad|j"];

    var ROMAN_RE = /^M{0,4}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$/;
    // Srpske reči koje izgledaju kao rimski brojevi kada su pisane verzalom
    var COLLIDING_ROMAN = ["MI", "VI", "LI", "DIV"];
    var SENTENCE_END_RE = /[.!?…:;\n]/;
    var SPACE_ONLY_RE = /^[   \t]+$/;
    var FOREIGN_RE = /[qwxyQWXY]/;
    var LATIN_LETTER = "A-Za-zČčĆćĐđŠšŽžÀ-ÖØ-öø-ÿ";
    var CYR_LETTER = "Ѐ-ӿ";
    var WORD_RE = new RegExp("[" + LATIN_LETTER + CYR_LETTER + "0-9_]+|[^" + LATIN_LETTER + CYR_LETTER + "0-9_]+", "g");
    var URL_RE = /(?:https?:\/\/|ftp:\/\/|www\.)[^\s<>"]+|[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}|\b[A-Za-z0-9\-]+\.(?:com|org|net|rs|ba|me|hr|info|edu|gov|io)(?:\/[^\s<>"]*)?\b/g;

    // Merne jedinice (SI i uobičajene). Jednoslovne jedinice (m, s, g, l, h, t, V, A...)
    // štite se samo iza broja, da se ne bi preskakali predlozi "s", "u", "a"...
    var UNITS = [
        // dužina, površina, zapremina
        "nm", "µm", "μm", "mm", "cm", "dm", "m", "km", "ft", "ha",
        "ml", "cl", "dl", "l", "L", "hl", "cc",
        // masa
        "µg", "mg", "g", "dag", "kg", "t", "lb", "oz", "ct",
        // vreme, frekvencija
        "ns", "µs", "ms", "s", "min", "h", "Hz", "kHz", "MHz", "GHz", "THz", "rpm",
        // elektrika, energija, snaga
        "mA", "A", "kA", "mV", "V", "kV", "MV", "W", "kW", "MW", "GW", "mW",
        "Wh", "kWh", "MWh", "GWh", "VA", "kVA", "MVA", "VAr", "kVAr", "Ah", "mAh",
        "Ω", "kΩ", "MΩ", "F", "µF", "nF", "pF", "H", "mH", "C", "T", "Wb", "S",
        "J", "kJ", "MJ", "cal", "kcal", "eV", "keV", "MeV", "KS", "hp",
        // sila, pritisak, temperatura
        "N", "kN", "MN", "Pa", "hPa", "kPa", "MPa", "GPa", "bar", "mbar", "atm", "psi", "K",
        // svetlo, zvuk, hemija, računari
        "cd", "lm", "lx", "dB", "dBm", "mol", "mmol", "ppm", "pH",
        "b", "B", "kb", "kB", "KB", "Mb", "MB", "Gb", "GB", "TB", "PB",
        "bps", "kbps", "Mbps", "Gbps", "dpi", "px", "pt"
    ].sort(function (a, b) { return b.length - a.length; });
    var UNIT = "(?:" + UNITS.map(escapeRe).join("|") + ")";
    var UNIT_EXPR = UNIT + "(?:[²³23])?(?:\\s?[\\/·⋅*]\\s?" + UNIT + "(?:[²³23])?)*";
    var LETTER_CLASS = "[" + LATIN_LETTER + CYR_LETTER + "µμΩ]";
    var UNITS_RE = new RegExp(
        // iza broja: "5 km", "10kg", "3,5 m²", "100 km/h"
        "\\d(?:\\s|\\u00A0|\\u202F)?" + UNIT_EXPR + "(?!" + LETTER_CLASS + ")" +
        // složene jedinice bilo gde: "km/h", "m/s²", "kWh/m2"
        "|(?<!" + LETTER_CLASS + ")" + UNIT + "(?:[²³23])?\\/" + UNIT + "(?:[²³23])?(?!" + LETTER_CLASS + ")" +
        // stepeni: "°C", "°F"
        "|°\\s?[CFK](?!" + LETTER_CLASS + ")",
        "g");

    function escapeRe(s) { return s.replace(/[.*+?^${}()|[\]\\\/]/g, "\\$&"); }

    // Korisnički izuzeci: jedna reč ili izraz po redu, "*" na kraju = bilo koji nastavak reči.
    function buildExceptionsRe(list) {
        if (!list) return null;
        if (typeof list === "string") list = list.split(/\r?\n/);
        var parts = list.map(function (s) { return s.trim(); })
            .filter(function (s) { return s && s !== "*"; })
            .map(function (s) {
                var star = s.charAt(s.length - 1) === "*";
                if (star) s = s.substring(0, s.length - 1);
                return escapeRe(s) + (star ? LETTER_CLASS + "*" : "");
            });
        if (!parts.length) return null;
        parts.sort(function (a, b) { return b.length - a.length; });
        return new RegExp("(?<!" + LETTER_CLASS + ")(?:" + parts.join("|") + ")(?!" + LETTER_CLASS + ")", "gi");
    }

    function isUpper(ch) { return ch !== ch.toLowerCase() && ch === ch.toUpperCase(); }
    function isLower(ch) { return ch !== ch.toUpperCase() && ch === ch.toLowerCase(); }
    function isLetter(ch) { return !!ch && ch.toLowerCase() !== ch.toUpperCase(); }

    function splitPositions(lowerWord, list) {
        var set = {};
        list.forEach(function (pat) {
            var anchored = pat.charAt(0) === "^";
            if (anchored) pat = pat.substring(1);
            var k = pat.indexOf("|");
            var plain = pat.replace("|", "");
            if (anchored) {
                if (lowerWord.indexOf(plain) === 0) set[k] = true;
            } else {
                var pos = lowerWord.indexOf(plain);
                while (pos !== -1) {
                    set[pos + k] = true;
                    pos = lowerWord.indexOf(plain, pos + 1);
                }
            }
        });
        return set;
    }

    function applyCase(src, dst) {
        return isUpper(src.charAt(0)) ? dst.toUpperCase() : dst;
    }

    // ---------- rimski brojevi u kontekstu ----------
    // tokens naizmenično sadrži reči i ono između njih (WORD_RE)

    function isRomanForm(w) { return !!w && ROMAN_RE.test(w); }

    function prevWord(tokens, i) {
        if (i < 2 || SENTENCE_END_RE.test(tokens[i - 1])) return null;
        return tokens[i - 2];
    }

    function nextWord(tokens, i) {
        if (i + 2 >= tokens.length || !SPACE_ONLY_RE.test(tokens[i + 1])) return null;
        return tokens[i + 2];
    }

    function isCapsWord(w) {
        return !!w && w.length >= 2 && !/[0-9_]/.test(w) &&
            w === w.toUpperCase() && w !== w.toLowerCase() && !isRomanForm(w);
    }

    // Rimski broj, ili reč koja samo tako izgleda ("I tako...", "MAMA I TATA", "MI SMO")?
    function isRoman(tokens, i) {
        var w = tokens[i];
        if (!isRomanForm(w)) return false;
        var next = nextWord(tokens, i);
        if (w === "I") {
            if (!next || !isLetter(next.charAt(0))) return true;
            var ln = next.toLowerCase();
            if ((ln === "i" || ln === "ili" || ln === "do") && isRomanForm(nextWord(tokens, i + 2))) return true;
            var prev = prevWord(tokens, i);
            return !(prev === null || isCapsWord(prev));
        }
        if (COLLIDING_ROMAN.indexOf(w) !== -1) return !isCapsWord(next);
        return true;
    }

    function wordToCyrillic(tokens, idx, opts) {
        var word = tokens[idx];
        if (opts.skipForeign && FOREIGN_RE.test(word)) return word;
        if (opts.skipRoman && isRoman(tokens, idx)) return word;
        if (!/[A-Za-zČčĆćĐđŠšŽž]/.test(word)) return word;

        var lower = word.toLowerCase();
        var splits = splitPositions(lower, EXCEPTIONS);
        var djSplits = opts.djToDj ? splitPositions(lower, DJ_EXCEPTIONS) : {};
        var out = "";
        for (var i = 0; i < word.length; i++) {
            var ch = word.charAt(i);
            var pair = lower.substr(i, 2);
            if (DIGRAPHS[pair] && !splits[i + 1]) {
                out += applyCase(ch, DIGRAPHS[pair]);
                i++;
                continue;
            }
            if (opts.djToDj && pair === "dj" && !djSplits[i + 1]) {
                out += applyCase(ch, "ђ");
                i++;
                continue;
            }
            var lc = lower.charAt(i);
            out += LAT2CYR[lc] ? applyCase(ch, LAT2CYR[lc]) : ch;
        }
        return out;
    }

    function protectors(opts, direction) {
        var list = [];
        if (opts.exceptionsRe) list.push(opts.exceptionsRe);
        if (direction === "cyr") {
            if (opts.skipUrls) list.push(URL_RE);
            if (opts.skipUnits) list.push(UNITS_RE);
        } else if (direction === "yu") {
            if (opts.skipUrls) list.push(URL_RE);
        }
        return list;
    }

    function toCyrillic(text, opts) {
        opts = normalizeOpts(opts);
        if (!text) return text;
        text = text.normalize ? text.normalize("NFC") : text;
        return mapUnprotected(text, protectors(opts, "cyr"), function (chunk) {
            var tokens = chunk.match(WORD_RE) || [], out = "";
            for (var i = 0; i < tokens.length; i++) out += wordToCyrillic(tokens, i, opts);
            return out;
        });
    }

    function toLatin(text, opts) {
        opts = normalizeOpts(opts);
        if (!text) return text;
        return mapUnprotected(text, protectors(opts, "lat"), latinChunk);
    }

    function latinChunk(text) {
        var out = "";
        for (var i = 0; i < text.length; i++) {
            var ch = text.charAt(i);
            var lc = ch.toLowerCase();
            var lat = CYR2LAT[lc];
            if (!lat) { out += ch; continue; }
            if (!isUpper(ch)) { out += lat; continue; }
            if (lat.length === 1) { out += lat.toUpperCase(); continue; }
            // Љ, Њ, Џ: "LJ" u reči pisanoj verzalom, inače "Lj"
            var next = text.charAt(i + 1), prev = text.charAt(i - 1);
            var allCaps = isLetter(next) ? isUpper(next) : (isLetter(prev) && isUpper(prev));
            out += allCaps ? lat.toUpperCase() : lat.charAt(0).toUpperCase() + lat.substring(1);
        }
        return out;
    }

    // Delovi teksta koje hvata bilo koji od regexa ostaju netaknuti, ostatak prolazi kroz fn.
    function mapUnprotected(text, regexes, fn) {
        if (!regexes.length) return fn(text);
        var re = regexes[0], rest = regexes.slice(1);
        var out = "", last = 0, m;
        re.lastIndex = 0;
        while ((m = re.exec(text)) !== null) {
            if (m[0] === "") { re.lastIndex++; continue; }
            out += mapUnprotected(text.substring(last, m.index), rest, fn) + m[0];
            last = m.index + m[0].length;
        }
        return out + mapUnprotected(text.substring(last), rest, fn);
    }

    var exceptionsCache = { key: null, re: null };

    function normalizeOpts(o) {
        o = o || {};
        if (o._normalized) return o;
        var key = Array.isArray(o.exceptions) ? o.exceptions.join("\n") : (o.exceptions || "");
        if (key !== exceptionsCache.key)
            exceptionsCache = { key: key, re: buildExceptionsRe(key) };
        return {
            _normalized: true,
            skipForeign: o.skipForeign !== false,
            skipRoman: o.skipRoman !== false,
            skipUrls: o.skipUrls !== false,
            skipUnits: o.skipUnits !== false,
            djToDj: !!o.djToDj,
            exceptionsRe: exceptionsCache.re
        };
    }

    // ---------- YUSCII (JUS I.B1.002 latinica, JUS I.B1.003 ćirilica) ----------

    var YU_LAT = {
        "@": "Ž", "`": "ž", "[": "Š", "{": "š", "\\": "Đ",
        "|": "đ", "]": "Ć", "}": "ć", "^": "Č", "~": "č"
    };
    var YU_CYR = {
        "A": "А", "B": "Б", "C": "Ц", "D": "Д", "E": "Е", "F": "Ф", "G": "Г", "H": "Х",
        "I": "И", "J": "Ј", "K": "К", "L": "Л", "M": "М", "N": "Н", "O": "О", "P": "П",
        "Q": "Љ", "R": "Р", "S": "С", "T": "Т", "U": "У", "V": "В", "W": "Њ", "X": "Џ",
        "Z": "З", "@": "Ж", "[": "Ш", "\\": "Ђ", "]": "Ћ", "^": "Ч"
    };
    Object.keys(YU_CYR).forEach(function (k) {
        var lk = { "@": "`", "[": "{", "\\": "|", "]": "}", "^": "~" }[k] || k.toLowerCase();
        YU_CYR[lk] = YU_CYR[k].toLowerCase();
    });
    var YU_SYMBOL_RE = /[@`\[\]{}\\|^~]/;

    // Znak YUSCII-ja (npr. "[" ili "@") menja se samo kada je deo reči,
    // da bi zagrade, "@" u tekstu i slično ostali netaknuti.
    function yusciiChunk(map) {
        return function (text) {
            var out = "";
            for (var i = 0; i < text.length; i++) {
                var ch = text.charAt(i), m = map[ch];
                if (!m) { out += ch; continue; }
                if (YU_SYMBOL_RE.test(ch) && !inWord(text, i)) { out += ch; continue; }
                out += m;
            }
            return out;
        };
    }

    function inWord(text, i) {
        function wordish(c) { return isLetter(c) || (!!c && YU_SYMBOL_RE.test(c)); }
        // gleda i preko susednih YUSCII znakova ("[}" i sl.)
        var j = i - 1;
        while (j >= 0 && YU_SYMBOL_RE.test(text.charAt(j))) j--;
        var k = i + 1;
        while (k < text.length && YU_SYMBOL_RE.test(text.charAt(k))) k++;
        return isLetter(text.charAt(j)) || isLetter(text.charAt(k)) ||
            (wordish(text.charAt(i - 1)) && wordish(text.charAt(i + 1)));
    }

    function yusciiToLatin(text, opts) {
        opts = normalizeOpts(opts);
        if (!text) return text;
        return mapUnprotected(text, protectors(opts, "yu"), yusciiChunk(YU_LAT));
    }

    function yusciiToCyrillic(text, opts) {
        opts = normalizeOpts(opts);
        if (!text) return text;
        return mapUnprotected(text, protectors(opts, "yu"), yusciiChunk(YU_CYR));
    }

    var api = {
        toCyrillic: toCyrillic, toLatin: toLatin,
        yusciiToLatin: yusciiToLatin, yusciiToCyrillic: yusciiToCyrillic
    };
    if (typeof module !== "undefined" && module.exports) module.exports = api;
    else root.SrTranslit = api;
})(this);
