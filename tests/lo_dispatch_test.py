# -*- coding: utf-8 -*-
"""
Proverava INSTALIRANU ekstenziju: poziva komande kao meni/traka
("service:org.digikv.translit.Translit?..."), u samom LibreOffice procesu.
Pokretanje kao za tests/lo_test.py, posle "unopkg add translit.oxt" u isti profil.
"""
import os
import sys
import time

import uno
from com.sun.star.beans import PropertyValue

PORT = int(os.environ.get("LO_PORT", "2083"))
failures = []


def check(name, got, want):
    if got != want:
        failures.append(name)
        print("GREŠKA %s\n  očekivano %r\n  dobijeno  %r" % (name, want, got))
    else:
        print("ok     %s" % name)


local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
for _ in range(120):
    try:
        ctx = resolver.resolve("uno:socket,host=127.0.0.1,port=%d;urp;StarOffice.ComponentContext" % PORT)
        break
    except Exception:
        time.sleep(0.5)
smgr = ctx.ServiceManager
desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)

check("komponenta je registrovana", smgr.createInstanceWithContext("org.digikv.translit.Translit", ctx) is not None, True)


def run(doc, command):
    frame = doc.getCurrentController().getFrame()
    frame.activate()
    t = time.time()
    dispatcher.executeDispatch(frame, "service:org.digikv.translit.Translit?" + command, "", 0, ())
    return time.time() - t


def open_doc(kind):
    return desktop.loadComponentFromURL("private:factory/" + kind, "_blank", 0, ())


# Writer
doc = open_doc("swriter")
text = doc.getText()
text.setString("Ljubav i džep, 90 km/h. I tako.")
run(doc, "cyr")
check("writer preko menija: latinica → ćirilica", text.getString(), "Љубав и џеп, 90 km/h. И тако.")
run(doc, "lat")
check("writer preko menija: ćirilica → latinica", text.getString(), "Ljubav i džep, 90 km/h. I tako.")

# brzina u samom LibreOffice-u
para = "Ovo je jedan duži pasus sa ljubavlju, njegom i džepovima; brzina 90 km/h. " * 3
text.setString("\n".join([para] * 1000))
dt = run(doc, "cyr")
print("       1000 pasusa (~%d reči) u samom LibreOffice-u: %.2f s" % (len(text.getString().split()), dt))
check("writer: veliki dokument", text.getString()[:12], "Ово је један")
doc.close(True)

# Calc
doc = open_doc("scalc")
sheet = doc.getSheets().getByIndex(0)
for r in range(2000):
    sheet.getCellByPosition(0, r).setString("ćelija broj %d, snaga 5 kW" % r)
doc.getCurrentController().select(sheet.getCellRangeByName("A1:A2000"))
dt = run(doc, "cyr")
print("       2000 ćelija: %.2f s" % dt)
check("calc preko menija", sheet.getCellByPosition(0, 1999).getString(), "ћелија број 1999, снага 5 kW")
doc.close(True)

print("\n%s" % ("SVE U REDU" if not failures else "NEUSPELO: " + ", ".join(failures)))
sys.exit(1 if failures else 0)
