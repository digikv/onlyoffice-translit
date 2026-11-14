# -*- coding: utf-8 -*-
"""
Test preslovljavanja u pravim LibreOffice dokumentima.

Pokrenite LibreOffice da sluša na portu 2083 (zaseban profil, bez prozora):
  soffice -env:UserInstallation=file:///C:/tmp/lo_profile --headless --norestore
          "--accept=socket,host=127.0.0.1,port=2083;urp;"
pa ovaj skript LibreOffice-ovim Python-om:
  "C:\\Program Files\\LibreOffice\\program\\python.exe" tests/lo_test.py
"""
import os
import sys
import time

import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.awt.FontSlant import ITALIC, NONE

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "libreoffice", "pythonpath"))
sys.path.insert(0, os.path.join(ROOT, "libreoffice"))
import Translit  # noqa: E402

HIDDEN = False
PORT = int(os.environ.get("LO_PORT", "2083"))
failures = []


def check(name, got, want):
    if got != want:
        failures.append(name)
        print("GREŠKA %s\n  očekivano %r\n  dobijeno  %r" % (name, want, got))
    else:
        print("ok     %s" % name)


def connect():
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(120):
        try:
            return resolver.resolve("uno:socket,host=127.0.0.1,port=%d;urp;StarOffice.ComponentContext" % PORT)
        except Exception:
            time.sleep(0.5)
    raise SystemExit("LibreOffice ne odgovara na portu %d" % PORT)


def pv(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


ctx = connect()
smgr = ctx.ServiceManager
desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)


def new_doc(kind):
    # Skriveni dokumenti nemaju označeni deo (getCurrentSelection vraća None), pa se otvaraju "vidljivi";
    # uz --headless ionako nema prozora.
    return desktop.loadComponentFromURL("private:factory/" + kind, "_blank", 0, (pv("Hidden", HIDDEN),))


def shape_type(shape):
    try:
        return shape.getShapeType()
    except Exception:
        return ""


def append_bold(xtext, plain, bold):
    """Dodaje običan pa podebljan tekst (radi i za ćelije Calc-a i oblike)."""
    xtext.insertString(xtext.getEnd(), plain, False)
    xtext.insertString(xtext.getEnd(), bold, False)
    c = xtext.createTextCursor()
    c.gotoEnd(False)
    c.goLeft(len(bold), True)
    c.CharWeight = 150.0


def settings(**kw):
    s = dict(Translit.DEFAULTS)
    s.update(kw)
    return s


def append(text, s, bold=False, italic=False, url=None):
    c = text.createTextCursorByRange(text.getEnd())
    text.insertString(c, s, True)
    c.CharWeight = 150.0 if bold else 100.0
    c.CharPosture = ITALIC if italic else NONE
    if url:
        c.HyperLinkURL = url
    return c


def portions(xtext):
    out = []
    for para in Translit.iterate(xtext):
        if not hasattr(para, "createEnumeration") or Translit.supports(para, "com.sun.star.text.TextTable"):
            continue
        row = []
        for por in Translit.iterate(para):
            if Translit.prop(por, "TextPortionType", "Text") != "Text":
                continue
            row.append((por.getString(), por.CharWeight > 100, por.CharPosture == ITALIC,
                        Translit.prop(por, "HyperLinkURL", "")))
        out.append(row)
    return out


# ------------------------------------------------------------------ Writer
def test_writer():
    doc = new_doc("swriter")
    text = doc.getText()
    append(text, "Normalan tekst, ")
    append(text, "ljubav", bold=True)
    append(text, " i ")
    append(text, "Njegoš", italic=True)
    append(text, " – brzina 90 km/h, 5 km, sajt ")
    append(text, "www.primer.rs", url="http://www.primer.rs")
    append(text, " i ")
    append(text, "L", bold=True)
    append(text, "jubica. ")
    append(text, "I tako je bilo u ")
    append(text, "XIV", italic=True)
    append(text, " veku.")
    # fusnota i polje usred pasusa
    fn = doc.createInstance("com.sun.star.text.Footnote")
    text.insertTextContent(text.getEnd(), fn, False)
    fn.setString("Napomena o džepu.")
    append(text, " Strana ")
    field = doc.createInstance("com.sun.star.text.TextField.PageNumber")
    field.NumberingType = 4  # ARABIC
    field.SubType = uno.Enum("com.sun.star.text.PageNumberType", "CURRENT")
    text.insertTextContent(text.getEnd(), field, False)
    append(text, " kraj.")
    text.insertControlCharacter(text.getEnd(), 0, False)  # PARAGRAPH_BREAK
    append(text, "Drugi pasus: čaša vode.")

    table = doc.createInstance("com.sun.star.text.TextTable")
    table.initialize(2, 2)
    text.insertTextContent(text.getEnd(), table, False)
    table.getCellByName("A1").setString("Snaga")
    table.getCellByName("B1").setString("10 kW")
    table.getCellByName("A2").setString("Napon")
    table.getCellByName("B2").setString("230 V")

    frame = doc.createInstance("com.sun.star.text.TextFrame")
    frame.Width, frame.Height = 4000, 2000
    text.insertTextContent(text.getEnd(), frame, False)
    frame.getText().setString("Tekst u okviru")

    style = doc.getStyleFamilies().getByName("PageStyles").getByName("Standard")
    style.HeaderIsOn = True
    style.HeaderText.setString("Zaglavlje dokumenta")

    shape = doc.createInstance("com.sun.star.drawing.RectangleShape")
    shape.Size = uno.createUnoStruct("com.sun.star.awt.Size", 3000, 1500)
    text.insertTextContent(text.getEnd(), shape, False)
    shape.setString("Oblik sa tekstom")

    original = text.getString()

    # --- ceo dokument u ćirilicu
    n = Translit.convert_document(doc, "cyr", settings(scope="all"))
    p = portions(text)
    check("writer: formatiranje prvog pasusa", p[0][:4], [
        ("Нормалан текст, ", False, False, ""),
        ("љубав", True, False, ""),
        (" и ", False, False, ""),
        ("Његош", False, True, ""),
    ])
    first = "".join(x[0] for x in p[0])
    check("writer: jedinice, URL, rimski broj, I", first,
          "Нормалан текст, љубав и Његош – брзина 90 km/h, 5 km, сајт www.primer.rs и Љубица. "
          "И тако је било у XIV веку. Страна  крај.")
    link = [x for x in p[0] if x[3]]
    check("writer: link ostao", link, [("www.primer.rs", False, False, "http://www.primer.rs")])
    k = [x[0] for x in p[0]].index("Љ")
    check("writer: digraf preko granice formatiranja (L podebljano + j)",
          (p[0][k][1], p[0][k + 1][0][:7], p[0][k + 1][1]), (True, "убица. ", False))
    check("writer: drugi pasus", "".join(x[0] for x in p[1]), "Други пасус: чаша воде.")
    check("writer: tabela", [table.getCellByName(c).getString() for c in ("A1", "B1", "A2", "B2")],
          ["Снага", "10 kW", "Напон", "230 V"])
    check("writer: okvir", frame.getText().getString(), "Текст у оквиру")
    check("writer: zaglavlje", style.HeaderText.getString(), "Заглавље документа")
    check("writer: fusnota", fn.getString(), "Напомена о џепу.")
    check("writer: oblik", shape.getString(), "Облик са текстом")
    check("writer: polje strane ostalo", field.getPresentation(False), "1")
    print("       izmenjeno segmenata:", n)

    # --- jedan korak poništavanja
    doc.getUndoManager().undo()
    check("writer: jedno poništavanje vraća sve", text.getString(), original)

    # --- označeni deo
    vc = doc.getCurrentController().getViewCursor()
    c = text.createTextCursor()
    c.gotoStart(False)
    c.goRight(len("Normalan tekst, "), False)
    c.goRight(len("ljubav i Njeg"), True)
    doc.getCurrentController().select(c)
    Translit.convert_document(doc, "cyr", settings())
    check("writer: samo označeno", text.getString().split("\n")[0][:40], "Normalan tekst, љубав и Његoš – brzina 9")
    p = portions(text)
    check("writer: označeno – formatiranje", p[0][:5], [
        ("Normalan tekst, ", False, False, ""),
        ("љубав", True, False, ""),
        (" и ", False, False, ""),
        ("Његoš", False, True, ""),
        (" – brzina 90 km/h, 5 km, sajt ", False, False, ""),
    ])
    doc.getUndoManager().undo()

    # --- bez oznake = ceo dokument; pa nazad u latinicu
    vc.gotoStart(False)
    Translit.convert_document(doc, "cyr", settings())
    Translit.convert_document(doc, "lat", settings())
    check("writer: tamo-nazad daje isti tekst", text.getString(), original)
    doc.close(True)


def test_writer_tracked_changes():
    doc = new_doc("swriter")
    text = doc.getText()
    text.setString("Prva reč obrisana druga reč.")
    doc.RecordChanges = True
    c = text.createTextCursor()
    c.gotoStart(False)
    c.goRight(len("Prva reč "), False)
    c.goRight(len("obrisana "), True)
    c.setString("")
    Translit.convert_document(doc, "cyr", settings(scope="all"))
    doc.RecordChanges = False
    check("writer/praćenje izmena: obrisani tekst nije preslovljen", "obrisana " in text.getString(), True)
    dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
    dispatcher.executeDispatch(doc.getCurrentController().getFrame(), ".uno:AcceptAllTrackedChanges", "", 0, ())
    check("writer/praćenje izmena: posle prihvatanja izmena", text.getString(), "Прва реч друга реч.")
    doc.close(True)


def test_writer_table_selection():
    doc = new_doc("swriter")
    text = doc.getText()
    table = doc.createInstance("com.sun.star.text.TextTable")
    table.initialize(2, 3)
    text.insertTextContent(text.getEnd(), table, False)
    for name in table.getCellNames():
        table.getCellByName(name).setString("reč " + name)
    sel = table.createCursorByCellName("A1")
    sel.gotoCellByName("B2", True)
    doc.getCurrentController().select(sel)
    Translit.convert_document(doc, "cyr", settings())
    check("writer: označene ćelije tabele",
          [table.getCellByName(n).getString() for n in ("A1", "B1", "A2", "B2", "C1", "C2")],
          ["реч А1", "реч Б1", "реч А2", "реч Б2", "reč C1", "reč C2"])
    doc.close(True)


def test_writer_combining():
    doc = new_doc("swriter")
    text = doc.getText()
    append(text, "c\u030ca", bold=True)
    append(text, "j i s\u030cec\u0301er")
    Translit.convert_document(doc, "cyr", settings(scope="all"))
    check("writer: rastavljena slova (c + ˇ)", [(x[0], x[1]) for x in portions(text)[0]],
          [("ча", True), ("ј и шећер", False)])
    doc.close(True)


# ------------------------------------------------------------------ Calc
def test_calc():
    doc = new_doc("scalc")
    sheet = doc.getSheets().getByIndex(0)
    sheet.getCellByPosition(0, 0).setString("Ljubav i džep")
    sheet.getCellByPosition(0, 1).setFormula('="formula tekst"')
    sheet.getCellByPosition(0, 2).setValue(42)
    sheet.getCellByPosition(0, 3).setString("Snaga 10 kW")
    rich = sheet.getCellByPosition(0, 4)
    append_bold(rich.getText(), "obično ", "podebljano")
    check("calc: priprema ćelije sa formatiranjem", [(x[0], x[1]) for x in portions(rich)[0]],
          [("obično ", False), ("podebljano", True)])
    sheet.getCellByPosition(1, 0).setString("druga kolona")
    doc.getSheets().insertNewByName("Drugi", 1)
    doc.getSheets().getByName("Drugi").getCellByPosition(0, 0).setString("drugi list")

    # označen samo A1:A5
    doc.getCurrentController().select(sheet.getCellRangeByName("A1:A5"))
    Translit.convert_document(doc, "cyr", settings())
    check("calc: tekstualne ćelije", sheet.getCellByPosition(0, 0).getString(), "Љубав и џеп")
    check("calc: formula netaknuta", sheet.getCellByPosition(0, 1).getFormula(), '="formula tekst"')
    check("calc: broj netaknut", sheet.getCellByPosition(0, 2).getValue(), 42.0)
    check("calc: jedinica", sheet.getCellByPosition(0, 3).getString(), "Снага 10 kW")
    check("calc: van oznake netaknuto", sheet.getCellByPosition(1, 0).getString(), "druga kolona")
    check("calc: formatiranje unutar ćelije", [(x[0], x[1]) for x in portions(rich)[0]],
          [("обично ", False), ("подебљано", True)])

    Translit.convert_document(doc, "cyr", settings(scope="all"))
    check("calc: ceo dokument – svi listovi",
          (sheet.getCellByPosition(1, 0).getString(),
           doc.getSheets().getByName("Drugi").getCellByPosition(0, 0).getString()),
          ("друга колона", "други лист"))
    doc.close(True)


# ------------------------------------------------------------------ Impress
def test_impress():
    doc = new_doc("simpress")
    page = doc.getDrawPages().getByIndex(0)
    shape = doc.createInstance("com.sun.star.drawing.RectangleShape")
    shape.Size = uno.createUnoStruct("com.sun.star.awt.Size", 8000, 3000)
    page.add(shape)
    append_bold(shape.getText(), "Naslov ", "slajda")
    other = doc.createInstance("com.sun.star.drawing.RectangleShape")
    other.Size = uno.createUnoStruct("com.sun.star.awt.Size", 8000, 3000)
    page.add(other)
    other.setString("Drugi oblik 5 m")

    table = doc.createInstance("com.sun.star.drawing.TableShape")
    page.add(table)
    table.Size = uno.createUnoStruct("com.sun.star.awt.Size", 8000, 2000)
    table.Position = uno.createUnoStruct("com.sun.star.awt.Point", 1000, 9000)
    model = table.Model
    if model.getRows().getCount() == 0:
        model.getRows().insertByIndex(0, 1)
        model.getColumns().insertByIndex(0, 1)
    model.getCellByPosition(0, 0).setString("ćelija tabele")

    notes = page.getNotesPage()
    for i in range(notes.getCount()):
        s = notes.getByIndex(i)
        if shape_type(s) == "com.sun.star.presentation.NotesShape":
            s.setString("Beleške predavača")

    # označen samo prvi oblik
    shapes = smgr.createInstanceWithContext("com.sun.star.drawing.ShapeCollection", ctx)
    shapes.add(shape)
    doc.getCurrentController().select(shapes)
    Translit.convert_document(doc, "cyr", settings())
    check("impress: označeni oblik + formatiranje", [(x[0], x[1]) for x in portions(shape)[0]],
          [("Наслов ", False), ("слајда", True)])
    check("impress: neoznačeni oblik netaknut", other.getString(), "Drugi oblik 5 m")

    Translit.convert_document(doc, "cyr", settings(scope="all"))
    check("impress: ceo dokument", other.getString(), "Други облик 5 m")
    check("impress: tabela", model.getCellByPosition(0, 0).getString(), "ћелија табеле")
    texts = [notes.getByIndex(i).getString() for i in range(notes.getCount())
             if shape_type(notes.getByIndex(i)) == "com.sun.star.presentation.NotesShape"]
    check("impress: beleške", texts, ["Белешке предавача"])
    doc.close(True)


def test_yuscii():
    doc = new_doc("swriter")
    text = doc.getText()
    append(text, "~a{a ", bold=True)
    append(text, "[uma i niz [1]")
    Translit.convert_document(doc, "yuLat", settings(scope="all"))
    check("writer: YUSCII → latinica", [(x[0], x[1]) for x in portions(text)[0]],
          [("čaša ", True), ("Šuma i niz [1]", False)])
    doc.close(True)


def test_speed():
    doc = new_doc("swriter")
    text = doc.getText()
    para = "Ovo je jedan duži pasus sa ljubavlju, njegom i džepovima; brzina 90 km/h. "
    c = text.createTextCursor()
    for i in range(300):
        text.insertString(text.getEnd(), para * 3, False)
        cc = text.createTextCursorByRange(text.getEnd())
        text.insertString(cc, "podebljano", True)
        cc.CharWeight = 150.0
        text.insertControlCharacter(text.getEnd(), 0, False)
    t = time.time()
    Translit.convert_document(doc, "cyr", settings(scope="all"))
    dt = time.time() - t
    print("       300 pasusa preko mreže (sporije nego u samom LibreOffice-u): %.1f s" % dt)
    check("writer: veliki dokument", text.getString()[:20], "Ово је један дужи па")
    doc.close(True)


TESTS = [test_writer, test_writer_tracked_changes, test_writer_table_selection, test_writer_combining,
         test_calc, test_impress, test_yuscii, test_speed]
if len(sys.argv) > 1:  # npr. lo_test.py test_calc
    TESTS = [t for t in TESTS if t.__name__ in sys.argv[1:]]
for test in TESTS:
    try:
        test()
    except Exception:
        import traceback
        traceback.print_exc()
        failures.append(test.__name__)

print("\n%s" % ("SVE U REDU" if not failures else "NEUSPELO: " + ", ".join(failures)))
sys.exit(1 if failures else 0)
