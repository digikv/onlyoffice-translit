(function (window) {
    "use strict";

    // Privatni Unicode znak kao zamena za tabulator, da se pri zameni ne pomeša sa pravim razmacima
    var TAB = "";
    var OPTION_IDS = ["skipForeign", "skipRoman", "skipUrls", "skipUnits", "djToDj"];
    var busy = false;

    function $(id) { return document.getElementById(id); }

    function setStatus(msg) { $("status").textContent = msg || ""; }

    function getOptions() {
        var o = {};
        OPTION_IDS.forEach(function (id) { o[id] = $(id).checked; });
        o.exceptions = $("exceptions").value;
        return o;
    }

    function getScope() {
        var r = document.querySelector("input[name=scope]:checked");
        return r ? r.value : "selection";
    }

    function saveOptions() {
        try { localStorage.setItem("translit-options", JSON.stringify(getOptions())); } catch (e) {}
    }

    function loadOptions() {
        try {
            var o = JSON.parse(localStorage.getItem("translit-options") || "null");
            if (o) OPTION_IDS.forEach(function (id) { if (id in o) $(id).checked = !!o[id]; });
            if (o && typeof o.exceptions === "string") $("exceptions").value = o.exceptions;
        } catch (e) {}
    }

    function makeConverter(direction) {
        var opts = getOptions();
        var fn = {
            cyr: SrTranslit.toCyrillic,
            lat: SrTranslit.toLatin,
            yuLat: SrTranslit.yusciiToLatin,
            yuCyr: SrTranslit.yusciiToCyrillic
        }[direction];
        return function (s) { return fn(s, opts); };
    }

    function done(msg) {
        busy = false;
        setStatus(msg);
    }

    // ---------- Tekst dokument i prezentacija: označeni tekst ----------

    function convertSelection(convert) {
        window.Asc.plugin.executeMethod("GetSelectedText", [{
            Numbering: false,
            Math: false,
            TableCellSeparator: "\n",
            ParaSeparator: "\n",
            TabSymbol: TAB
        }], function (text) {
            if (!text) return done("Nije označen tekst.");
            var paras = text.replace(/\r/g, "").split("\n");
            if (paras.length > 1 && paras[paras.length - 1] === "") paras.pop();
            var converted = paras.map(convert);
            var changed = converted.some(function (p, i) { return p !== paras[i]; });
            if (!changed) return done("Nema teksta za preslovljavanje.");
            window.Asc.plugin.executeMethod("ReplaceTextSmart", [converted, TAB, "\n"], function () {
                done("Gotovo.");
            });
        });
    }

    function selectWholeDocument(callback) {
        window.Asc.plugin.callCommand(function () {
            var doc = Api.GetDocument();
            var range = doc.GetRange();
            if (range && range.Select) { range.Select(); return true; }
            return false;
        }, false, false, callback);
    }

    // ---------- Prezentacija: ceo dokument (prolaz kroz sve oblike) ----------

    // Ista funkcija se izvršava dva puta: "read" skuplja tekst, "write" upisuje preslovljeni.
    // Mora biti samostalna jer je OnlyOffice serijalizuje i izvršava u editoru.
    function presentationRunsCommand() {
        var mode = Asc.scope.mode, texts = Asc.scope.texts || [], out = [], idx = 0;
        var pres = Api.GetPresentation();

        function handleContent(content) {
            if (!content || !content.GetElementsCount) return;
            for (var p = 0; p < content.GetElementsCount(); p++) {
                var para = content.GetElement(p);
                if (!para || !para.GetClassType || para.GetClassType() !== "paragraph") continue;
                for (var r = 0; r < para.GetElementsCount(); r++) {
                    var run = para.GetElement(r);
                    if (!run || !run.GetClassType || run.GetClassType() !== "run" || !run.GetText) continue;
                    var t = run.GetText();
                    if (typeof t !== "string" || t === "" || /[\t\r\n]/.test(t)) continue;
                    if (mode === "read") {
                        out.push(t);
                    } else {
                        var nt = texts[idx++];
                        if (typeof nt === "string" && nt !== t) {
                            run.ClearContent();
                            run.AddText(nt);
                        }
                    }
                }
            }
        }

        function handleShapes(shapes) {
            for (var i = 0; i < shapes.length; i++) {
                var sh = shapes[i];
                if (sh && sh.GetDocContent) handleContent(sh.GetDocContent());
            }
        }

        for (var s = 0; s < pres.GetSlidesCount(); s++) {
            var slide = pres.GetSlideByIndex(s);
            if (slide && slide.GetAllShapes) handleShapes(slide.GetAllShapes());
        }
        return mode === "read" ? JSON.stringify(out) : "";
    }

    function convertPresentation(convert) {
        Asc.scope.mode = "read";
        Asc.scope.texts = [];
        window.Asc.plugin.callCommand(presentationRunsCommand, false, false, function (res) {
            var texts = [];
            try { texts = JSON.parse(res || "[]"); } catch (e) {}
            if (!texts.length) return done("Nema teksta za preslovljavanje.");
            Asc.scope.mode = "write";
            Asc.scope.texts = texts.map(convert);
            window.Asc.plugin.callCommand(presentationRunsCommand, false, true, function () {
                done("Gotovo.");
            });
        });
    }

    // ---------- Tabela ----------

    function convertSpreadsheet(convert, whole) {
        Asc.scope.whole = whole;
        window.Asc.plugin.callCommand(function () {
            var sheet = Api.GetActiveSheet();
            var range = Asc.scope.whole ? sheet.GetUsedRange() : sheet.GetSelection();
            var out = [];
            if (!range) return "[]";
            range.ForEach(function (cell) {
                var f = cell.GetFormula ? cell.GetFormula() : "";
                if (typeof f === "string" && f.charAt(0) === "=") return;
                var v = cell.GetValue();
                if (typeof v === "string" && v !== "" && isNaN(Number(v)))
                    out.push([cell.GetAddress(false, false, "xlA1", false), v]);
            });
            return JSON.stringify(out);
        }, false, false, function (res) {
            var cells = [];
            try { cells = JSON.parse(res || "[]"); } catch (e) {}
            var changed = [];
            cells.forEach(function (c) {
                var nv = convert(c[1]);
                if (nv !== c[1]) changed.push([c[0], nv]);
            });
            if (!changed.length) return done("Nema teksta za preslovljavanje.");
            Asc.scope.cells = changed;
            window.Asc.plugin.callCommand(function () {
                var sheet = Api.GetActiveSheet();
                var cells = Asc.scope.cells;
                for (var i = 0; i < cells.length; i++)
                    sheet.GetRange(cells[i][0]).SetValue(cells[i][1]);
            }, false, true, function () {
                done("Gotovo (" + changed.length + " ćelija).");
            });
        });
    }

    // ---------- Pokretanje ----------

    function run(direction) {
        if (busy) return;
        busy = true;
        setStatus("Radim…");
        saveOptions();

        var convert = makeConverter(direction);
        var whole = getScope() === "all";
        var editor = window.Asc.plugin.info.editorType;

        try {
            if (editor === "cell") {
                convertSpreadsheet(convert, whole);
            } else if (editor === "slide") {
                if (whole) convertPresentation(convert);
                else convertSelection(convert);
            } else {
                if (whole) {
                    selectWholeDocument(function (ok) {
                        if (ok === false) return done("Nije moguće označiti ceo dokument.");
                        convertSelection(convert);
                    });
                } else {
                    convertSelection(convert);
                }
            }
        } catch (e) {
            done("Greška: " + (e && e.message ? e.message : e));
        }
    }

    window.Asc.plugin.init = function () {
        loadOptions();
        $("toCyr").onclick = function () { run("cyr"); };
        $("toLat").onclick = function () { run("lat"); };
        $("yuLat").onclick = function () { run("yuLat"); };
        $("yuCyr").onclick = function () { run("yuCyr"); };
        OPTION_IDS.forEach(function (id) { $(id).onchange = saveOptions; });
        $("exceptions").oninput = saveOptions;
    };

    window.Asc.plugin.button = function () {
        this.executeCommand("close", "");
    };

    window.Asc.plugin.onThemeChanged = function (theme) {
        window.Asc.plugin.onThemeChangedBase(theme);
        document.body.style.color = theme["text-normal"] || "";
    };

})(window);
