# -*- coding: utf-8 -*-
"""
Translit za LibreOffice – preslovljavanje srpskog teksta latinica <-> ćirilica i YUSCII.

UNO komponenta koju pozivaju meni i traka sa alatkama (Addons.xcu) preko URL-a
"service:org.digikv.translit.Translit?<komanda>". Pravila preslovljavanja su u
pythonpath/srtranslit.py.

Tekst se menja segment po segment (text portion), tako da formatiranje
(podebljano, kurziv, boje, linkovi...) ostaje sačuvano.
"""
import json
import os
import traceback

import uno
import unohelper
from com.sun.star.task import XJobExecutor
from com.sun.star.awt import XActionListener

import srtranslit

IMPL_NAME = "org.digikv.translit.Translit"
SERVICE_NAMES = (IMPL_NAME,)

TITLES = {
    "cyr": "Latinica → Ћирилица",
    "lat": "Ћирилица → Latinica",
    "yuLat": "YUSCII → Latinica",
    "yuCyr": "YUSCII → Ћирилица",
}

OPTION_KEYS = ("skipForeign", "skipRoman", "skipUrls", "skipUnits", "djToDj")
DEFAULTS = {
    "skipForeign": True,
    "skipRoman": True,
    "skipUrls": True,
    "skipUnits": True,
    "djToDj": False,
    "exceptions": "",
    "scope": "auto",  # "auto": označeno, a ako ništa nije označeno – ceo dokument; "all": uvek ceo dokument
}

OBJ = "￼"  # zamena za polja, fusnote i sl. unutar pasusa (nije slovo, deli reči)


# ---------------------------------------------------------------- podešavanja

def settings_path(ctx):
    subst = ctx.ServiceManager.createInstanceWithContext("com.sun.star.util.PathSubstitution", ctx)
    url = subst.substituteVariables("$(user)/translit.json", True)
    return uno.fileUrlToSystemPath(url)


def load_settings(ctx):
    s = dict(DEFAULTS)
    try:
        with open(settings_path(ctx), encoding="utf-8") as f:
            data = json.load(f)
        for k in DEFAULTS:
            if k in data and isinstance(data[k], type(DEFAULTS[k])):
                s[k] = data[k]
    except Exception:
        pass
    return s


def save_settings(ctx, s):
    try:
        path = settings_path(ctx)
        with open(path, "w", encoding="utf-8") as f:
            json.dump({k: s[k] for k in DEFAULTS}, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ---------------------------------------------------------------- pomoćno

def supports(obj, service):
    try:
        return obj is not None and obj.supportsService(service)
    except Exception:
        return False


def prop(obj, name, default=None):
    try:
        return obj.getPropertyValue(name)
    except Exception:
        return default


def iterate(enum_access):
    enum = enum_access.createEnumeration()
    while enum.hasMoreElements():
        yield enum.nextElement()


def message_box(ctx, frame, text, error=False):
    smgr = ctx.ServiceManager
    toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    parent = frame.getContainerWindow() if frame is not None else toolkit.getDesktopWindow()
    kind = uno.Enum("com.sun.star.awt.MessageBoxType", "ERRORBOX" if error else "INFOBOX")
    buttons = uno.getConstantByName("com.sun.star.awt.MessageBoxButtons.BUTTONS_OK")
    box = toolkit.createMessageBox(parent, kind, buttons, "Translit", text)
    box.execute()
    box.dispose()


# ---------------------------------------------------------------- izmena teksta

class Converter(object):
    def __init__(self, mode, settings, frame=None):
        self.mode = mode
        self.opts = srtranslit.normalize_opts({k: settings[k] for k in OPTION_KEYS + ("exceptions",)})
        self.changed = 0
        self.in_delete = False
        self.frame = frame
        self.indicator = None
        self.done = 0
        self.total = 0

    def spans(self, text):
        return srtranslit.transliterate_spans(text, self.mode, self.opts)

    # traka napretka u statusnoj liniji, samo za veće dokumente
    def start_progress(self, total):
        if self.frame is None or total < 300 or self.indicator is not None:
            return
        try:
            self.indicator = self.frame.createStatusIndicator()
            self.indicator.start("Translit: " + TITLES[self.mode], total)
            self.total = total
        except Exception:
            self.indicator = None

    def tick(self):
        if self.indicator is not None:
            self.done += 1
            if self.done % 50 == 0:
                self.indicator.setValue(min(self.done, self.total))

    def end_progress(self):
        if self.indicator is not None:
            self.indicator.end()
            self.indicator = None


def replace_in_portion(xtext, portion, old, new):
    """Zamenjuje tekst jednog segmenta tako da novi tekst preuzme njegovo formatiranje.

    Novi tekst se umeće UNUTAR segmenta (iza bar jednog njegovog starog znaka), pa
    tek onda briše stari – tako dobija atribute tog segmenta, a ne susednog.
    Brisanja idu od kraja ka početku: kada je uključeno praćenje izmena, obrisani
    znaci ostaju u tekstu, pa položaji posle njih ne bi bili tačni."""
    n = len(old)
    pre = 0
    lim = min(n, len(new))
    while pre < lim and old[pre] == new[pre]:
        pre += 1
    suf = 0
    while suf < lim - pre and old[n - 1 - suf] == new[len(new) - 1 - suf]:
        suf += 1
    a, b, ins = pre, n - suf, new[pre:len(new) - suf]
    if a == b:  # čisto umetanje – proširi na jedan stari znak
        if b < n:
            ins, b = ins + old[a], b + 1
        else:
            a, ins = a - 1, old[a - 1] + ins

    anchor = xtext.createTextCursorByRange(portion.getStart())

    def cursor(offset, select=0):
        c = xtext.createTextCursorByRange(anchor)
        if offset:
            c.goRight(offset, False)
        if select:
            c.goRight(select, True)
        return c

    if a > 0:
        if ins:
            xtext.insertString(cursor(a), ins, False)
        cursor(a + len(ins), b - a).setString("")
    elif not ins:
        cursor(0, b).setString("")
    else:
        xtext.insertString(cursor(1), ins, False)
        if b > 1:
            cursor(1 + len(ins), b - 1).setString("")
        cursor(0, 1).setString("")


def distribute(portions, spans, length):
    """Raspoređuje izmene (početak, kraj, tekst) na segmente; vraća novi tekst svakog segmenta."""
    owner = [-1] * length
    for idx, (_, start, s) in enumerate(portions):
        owner[start:start + len(s)] = [idx] * len(s)
    new = [list(s) for (_, _, s) in portions]
    for a, b, out in reversed(spans):
        if a == b:
            o = owner[a] if a < length else -1
            if o < 0 and a > 0:
                o = owner[a - 1]
            if o >= 0:
                lp = a - portions[o][1]
                new[o][lp:lp] = [out]
            continue
        touched = []
        i = a
        while i < b:
            o = owner[i]
            j = i
            while j < b and owner[j] == o:
                j += 1
            if o >= 0:
                start = portions[o][1]
                touched.append((o, i - start, j - start))
            i = j
        for t, (o, lo, hi) in enumerate(touched):
            new[o][lo:hi] = [out] if t == 0 else []
    return ["".join(x) for x in new]


def process_paragraph(para, conv):
    conv.tick()
    xtext = para.getText()
    portions = []  # (segment, početak u tekstu pasusa, tekst)
    parts = []
    pos = 0
    for por in iterate(para):
        kind = prop(por, "TextPortionType", "Text")
        if kind == "Text":
            s = por.getString()
            if s and not conv.in_delete:
                portions.append((por, pos, s))
                parts.append(s)
                pos += len(s)
        elif kind == "Redline":
            if prop(por, "RedlineType") == "Delete":
                conv.in_delete = bool(prop(por, "IsStart", False))
        else:
            try:
                s = por.getString()
            except Exception:
                s = ""
            if s:
                parts.append(OBJ)
                pos += 1
    if not portions:
        return
    spans = conv.spans("".join(parts))
    if not spans:
        return
    new_texts = distribute(portions, spans, pos)
    for (por, _, old), new in reversed(list(zip(portions, new_texts))):
        if new != old:
            replace_in_portion(xtext, por, old, new)
            conv.changed += 1


def process_text(xtext, conv):
    """Svi pasusi (i tabele) nekog teksta: telo dokumenta, ćelija, okvir, oblik..."""
    conv.in_delete = False
    for el in iterate(xtext):
        if supports(el, "com.sun.star.text.TextTable"):
            process_writer_table(el, conv)
        elif hasattr(el, "createEnumeration"):
            process_paragraph(el, conv)


def process_writer_table(table, conv, names=None):
    for name in names if names is not None else table.getCellNames():
        cell = table.getCellByName(name)
        if cell is not None:
            process_text(cell, conv)
    conv.in_delete = False


def shape_type(shape):
    try:
        return shape.getShapeType()
    except Exception:
        return ""


def process_shape(shape, conv, writer=False):
    if supports(shape, "com.sun.star.drawing.GroupShape"):
        for i in range(shape.getCount()):
            process_shape(shape.getByIndex(i), conv, writer)
        return
    if writer and (supports(shape, "com.sun.star.text.TextFrame") or prop(shape, "TextBox", False)):
        return  # okviri i tekst-polja se obrađuju preko TextFrames
    if shape_type(shape) == "com.sun.star.drawing.TableShape":
        table = prop(shape, "Model")
        if table is not None:
            for r in range(table.getRows().getCount()):
                for c in range(table.getColumns().getCount()):
                    try:
                        cell = table.getCellByPosition(c, r)
                    except Exception:
                        continue
                    if cell.getString():
                        process_text(cell, conv)
        return
    if supports(shape, "com.sun.star.drawing.Text"):
        try:
            if shape.getString():
                process_text(shape, conv)
        except Exception:
            pass


def process_shapes(container, conv, writer=False):
    for i in range(container.getCount()):
        process_shape(container.getByIndex(i), conv, writer)


def process_range(rng, conv):
    """Označeni deo teksta: pasusi u opsegu (prvi i poslednji samo delimično)."""
    cur = rng.getText().createTextCursorByRange(rng)
    conv.in_delete = False
    for el in iterate(cur):
        if supports(el, "com.sun.star.text.TextTable"):
            process_writer_table(el, conv)
        else:
            process_paragraph(el, conv)


# ---------------------------------------------------------------- Writer

def writer_all(doc, conv):
    process_text(doc.getText(), conv)

    frames = doc.getTextFrames()
    for name in frames.getElementNames():
        process_text(frames.getByName(name).getText(), conv)

    for notes in (doc.getFootnotes(), doc.getEndnotes()):
        for i in range(notes.getCount()):
            process_text(notes.getByIndex(i), conv)

    seen = []
    styles = doc.getStyleFamilies().getByName("PageStyles")
    for name in styles.getElementNames():
        style = styles.getByName(name)
        for on, keys in (("HeaderIsOn", ("HeaderText", "HeaderTextLeft", "HeaderTextFirst")),
                         ("FooterIsOn", ("FooterText", "FooterTextLeft", "FooterTextFirst"))):
            if not prop(style, on, False):
                continue
            for key in keys:
                t = prop(style, key)
                if t is not None and not any(t == x for x in seen):
                    seen.append(t)
                    process_text(t, conv)

    process_shapes(doc.getDrawPage(), conv, writer=True)


def cell_names_in_range(rng):
    """Imena ćelija u opsegu tabele Writer-a (i kada tabela ima spojene ćelije)."""
    try:
        data = rng.getDataArray()
        rows, cols = len(data), max(len(row) for row in data)
    except Exception:
        rows = cols = None
    names = []
    r = 0
    while rows is None or r < rows:
        c = 0
        found = False
        while cols is None or c < cols:
            try:
                cell = rng.getCellByPosition(c, r)
            except Exception:
                if cols is None:
                    break
                c += 1
                continue
            found = True
            name = prop(cell, "CellName")
            if name and name not in names:
                names.append(name)
            c += 1
        if rows is None and not found:
            break
        r += 1
    return names


def writer_selection(doc, conv):
    """Vraća False ako ništa nije označeno."""
    sel = doc.getCurrentSelection()
    if sel is None:
        return False
    if supports(sel, "com.sun.star.text.TextRanges"):
        ranges = [sel.getByIndex(i) for i in range(sel.getCount())]
        ranges = [r for r in ranges if r.getString()]
        if not ranges:
            return False
        for r in ranges:
            process_range(r, conv)
        return True
    if supports(sel, "com.sun.star.text.TextTableCursor"):
        table = prop(doc.getCurrentController().getViewCursor(), "TextTable")
        if table is None:
            return False
        process_writer_table(table, conv, cell_names_in_range(table.getCellRangeByName(sel.getRangeName())))
        return True
    if supports(sel, "com.sun.star.text.TextFrame"):
        process_text(sel.getText(), conv)
        return True
    if hasattr(sel, "getCount") and hasattr(sel, "getByIndex"):  # označeni oblici
        if sel.getCount() == 0:
            return False
        process_shapes(sel, conv, writer=False)
        return True
    return False


# ---------------------------------------------------------------- Calc

def process_cell(cell, conv):
    conv.tick()
    old = cell.getString()
    spans = conv.spans(old) if old else None
    if not spans:
        return
    paras = list(iterate(cell))
    if len(paras) == 1:
        pors = list(iterate(paras[0]))
        if len(pors) == 1 and prop(pors[0], "TextPortionType", "Text") == "Text":
            # obična ćelija bez formatiranja unutar teksta – brža zamena
            cell.setString(_apply_spans(old, spans))
            conv.changed += 1
            return
    conv.in_delete = False
    for para in paras:
        process_paragraph(para, conv)


def _apply_spans(text, spans):
    out = []
    last = 0
    for a, b, rep in spans:
        out.append(text[last:a])
        out.append(rep)
        last = b
    out.append(text[last:])
    return "".join(out)


def cell_count(ranges):
    n = 0
    try:
        for a in ranges.getRangeAddresses():
            n += (a.EndRow - a.StartRow + 1) * (a.EndColumn - a.StartColumn + 1)
    except Exception:
        pass
    return n


def process_cell_ranges(ranges, conv):
    if ranges is None:
        return
    for cell in iterate(ranges.getCells()):
        process_cell(cell, conv)


def calc_convert(doc, conv, whole):
    string_flag = uno.getConstantByName("com.sun.star.sheet.CellFlags.STRING")
    if not whole:
        sel = doc.getCurrentSelection()
        if supports(sel, "com.sun.star.sheet.SheetCellRange") or supports(sel, "com.sun.star.sheet.SheetCellRanges"):
            ranges = sel.queryContentCells(string_flag)
            conv.start_progress(cell_count(ranges))
            process_cell_ranges(ranges, conv)
            return
        if hasattr(sel, "getCount") and hasattr(sel, "getByIndex") and sel.getCount() > 0:
            process_shapes(sel, conv)
            return
        if hasattr(sel, "getString") and hasattr(sel, "getText") and sel.getString():
            process_range(sel, conv)
            return
    sheets = doc.getSheets()
    targets = [sheets.getByIndex(i) for i in range(sheets.getCount())]
    targets = [(sheet, sheet.queryContentCells(string_flag)) for sheet in targets]
    conv.start_progress(sum(cell_count(r) for _, r in targets))
    for sheet, ranges in targets:
        process_cell_ranges(ranges, conv)
        process_shapes(sheet.getDrawPage(), conv)


# ---------------------------------------------------------------- Impress / Draw

def draw_selection(doc, conv):
    sel = doc.getCurrentSelection()
    if sel is None:
        return False
    if hasattr(sel, "getCount") and hasattr(sel, "getByIndex"):
        if sel.getCount() == 0:
            return False
        process_shapes(sel, conv)
        return True
    if hasattr(sel, "getString") and hasattr(sel, "getText"):  # označen tekst unutar oblika
        if not sel.getString():
            return False
        process_range(sel, conv)
        return True
    return False


def draw_all(doc, conv):
    pages = doc.getDrawPages()
    for i in range(pages.getCount()):
        page = pages.getByIndex(i)
        process_shapes(page, conv)
        notes = None
        try:
            notes = page.getNotesPage()
        except Exception:
            pass
        if notes is not None:
            process_shapes(notes, conv)


# ---------------------------------------------------------------- pokretanje

def convert_document(doc, mode, settings, frame=None):
    """Preslovljava dokument; vraća broj izmenjenih segmenata/ćelija ili None ako vrsta dokumenta nije podržana."""
    conv = Converter(mode, settings, frame)
    whole = settings.get("scope") == "all"

    if supports(doc, "com.sun.star.text.TextDocument"):
        kind = "writer"
    elif supports(doc, "com.sun.star.sheet.SpreadsheetDocument"):
        kind = "calc"
    elif supports(doc, "com.sun.star.drawing.GenericDrawingDocument"):
        kind = "draw"
    else:
        return None

    # Označeno se čita pre zaključavanja prikaza
    undo = None
    try:
        undo = doc.getUndoManager()
        undo.enterUndoContext("Translit: " + TITLES[mode])
    except Exception:
        undo = None
    doc.lockControllers()
    auto_calc = None
    try:
        if kind == "writer":
            if whole or not writer_selection(doc, conv):
                conv.start_progress(prop(doc, "ParagraphCount", 0) or 0)
                writer_all(doc, conv)
        elif kind == "calc":
            try:
                auto_calc = doc.isAutomaticCalculationEnabled()
                doc.enableAutomaticCalculation(False)
            except Exception:
                auto_calc = None
            calc_convert(doc, conv, whole)
        else:
            if whole or not draw_selection(doc, conv):
                draw_all(doc, conv)
    finally:
        conv.end_progress()
        if auto_calc is not None:
            doc.enableAutomaticCalculation(auto_calc)
        doc.unlockControllers()
        if undo is not None:
            undo.leaveUndoContext()
    return conv.changed


# ---------------------------------------------------------------- dijalog

class _EndDialog(unohelper.Base, XActionListener):
    def __init__(self, dialog, code):
        self.dialog = dialog
        self.code = code

    def actionPerformed(self, ev):
        self.dialog.endDialog(self.code)

    def disposing(self, ev):
        pass


DIALOG_MODES = ("cyr", "lat", "yuLat", "yuCyr")


def show_dialog(ctx, frame, settings):
    """Vraća (komanda, podešavanja); komanda je smer preslovljavanja, "save" ili None."""
    smgr = ctx.ServiceManager
    model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    W = 236
    model.Width = W
    model.Height = 292
    model.Title = "Translit – preslovljavanje"

    def add(kind, name, x, y, w, h, **props):
        m = model.createInstance("com.sun.star.awt.UnoControl" + kind + "Model")
        m.PositionX, m.PositionY, m.Width, m.Height = x, y, w, h
        for k, v in props.items():
            setattr(m, k, v)
        model.insertByName(name, m)
        return m

    full = W - 12
    add("FixedLine", "lScope", 6, 4, full, 10, Label="Opseg / Опсег")
    add("RadioButton", "scopeAuto", 12, 16, full - 6, 10,
        Label="Označeni tekst (ako ništa nije označeno – ceo dokument)",
        State=0 if settings["scope"] == "all" else 1)
    add("RadioButton", "scopeAll", 12, 28, full - 6, 10, Label="Uvek ceo dokument",
        State=1 if settings["scope"] == "all" else 0)

    add("FixedLine", "lOpts", 6, 42, full, 10, Label="Opcije")
    labels = (
        ("skipForeign", "Preskoči strane reči (q, w, x, y)"),
        ("skipRoman", "Preskoči rimske brojeve (XIV, MMXXVI)"),
        ("skipUrls", "Preskoči veb adrese i e‑poštu"),
        ("skipUnits", "Ne preslovljavaj merne jedinice (km, kg, m², km/h, °C…)"),
        ("djToDj", "„dj“ → „ђ“ (osim odjednom, podjednako…)"),
    )
    y = 54
    for key, label in labels:
        add("CheckBox", key, 12, y, full - 6, 10, Label=label, State=1 if settings[key] else 0)
        y += 12

    add("FixedLine", "lExc", 6, y + 2, full, 10, Label="Izuzeci")
    y += 14
    add("Edit", "exceptions", 12, y, full - 6, 50, MultiLine=True, VScroll=True,
        Text=settings["exceptions"].replace("\r\n", "\n").replace("\n", "\r\n"),
        HelpText="Jedna reč ili izraz po redu")
    y += 52
    add("FixedText", "hintExc", 12, y, full - 6, 18, MultiLine=True,
        Label="Jedna reč ili izraz po redu; ne preslovljavaju se ni u jednom smeru. "
              "Zvezdica na kraju znači „i svi oblici“: Microsoft* obuhvata i Microsoftov.")
    y += 22

    add("FixedLine", "lRun", 6, y, full, 10, Label="Preslovi")
    y += 12
    half = (full - 6 - 4) // 2
    help_yu = "Stari YUSCII tekst, npr. ~a{a → čaša. Posle konverzije promenite YU font (npr. „YU Times“) u običan."
    buttons = (
        ("bCyr", 12, y, TITLES["cyr"], ""),
        ("bLat", 12 + half + 4, y, TITLES["lat"], ""),
        ("bYuLat", 12, y + 16, TITLES["yuLat"], help_yu),
        ("bYuCyr", 12 + half + 4, y + 16, TITLES["yuCyr"], help_yu),
    )
    for name, x, by, label, helptext in buttons:
        add("Button", name, x, by, half, 14, Label=label, HelpText=helptext)
    y += 36

    add("FixedLine", "lEnd", 6, y, full, 6)
    y += 8
    ok_type = uno.Enum("com.sun.star.awt.PushButtonType", "OK")
    cancel_type = uno.Enum("com.sun.star.awt.PushButtonType", "CANCEL")
    add("Button", "bSave", W - 6 - 50 - 4 - 50, y, 50, 14, Label="Sačuvaj", PushButtonType=ok_type,
        DefaultButton=True)
    add("Button", "bClose", W - 6 - 50, y, 50, 14, Label="Zatvori", PushButtonType=cancel_type)
    model.Height = y + 20

    dialog = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dialog.setModel(model)
    toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    parent = frame.getContainerWindow() if frame is not None else toolkit.getDesktopWindow()
    dialog.createPeer(toolkit, parent)
    for code, (name, _, _, _, _) in enumerate(buttons):
        dialog.getControl(name).addActionListener(_EndDialog(dialog, 10 + code))

    result = dialog.execute()

    new = dict(settings)
    for key, _ in labels:
        new[key] = model.getByName(key).State == 1
    new["exceptions"] = model.getByName("exceptions").Text.replace("\r\n", "\n")
    new["scope"] = "all" if model.getByName("scopeAll").State == 1 else "auto"
    dialog.dispose()

    if result >= 10:
        return DIALOG_MODES[result - 10], new
    if result == 1:
        return "save", new
    return None, settings


# ---------------------------------------------------------------- UNO komponenta

class Translit(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        self.ctx = ctx

    def _current(self):
        desktop = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
        frame = desktop.getCurrentFrame()
        doc = None
        if frame is not None:
            controller = frame.getController()
            if controller is not None:
                doc = controller.getModel()
        if doc is None:
            doc = desktop.getCurrentComponent()
        return frame, doc

    def trigger(self, command):
        frame = None
        try:
            frame, doc = self._current()
            settings = load_settings(self.ctx)
            if command == "dialog":
                command, new = show_dialog(self.ctx, frame, settings)
                if command is None:
                    return
                save_settings(self.ctx, new)
                settings = new
                if command == "save":
                    return
            if command not in TITLES:
                return
            if doc is None:
                return
            changed = convert_document(doc, command, settings, frame)
            if changed is None:
                message_box(self.ctx, frame, "Preslovljavanje radi u Writer-u, Calc-u, Impress-u i Draw-u.")
            elif changed == 0:
                message_box(self.ctx, frame, "Nema teksta za preslovljavanje.")
        except Exception:
            try:
                message_box(self.ctx, frame, "Greška pri preslovljavanju:\n\n" + traceback.format_exc(), error=True)
            except Exception:
                traceback.print_exc()


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(Translit, IMPL_NAME, SERVICE_NAMES)
