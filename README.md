# Translit za ONLYOFFICE, LibreOffice i MS Office

Dodatak koji preslovljava srpski tekst latinica ↔ ćirilica, po uzoru na ooOTranslit za LibreOffice/OpenOffice –
u tri verzije sa istim pravilima: za **ONLYOFFICE**, **LibreOffice** i **Microsoft Office** (Word, Excel, PowerPoint).

*Serbian Latin ↔ Cyrillic (and legacy YUSCII) transliteration for ONLYOFFICE, LibreOffice and Microsoft Office, inspired by ooOTranslit.*

## Mogućnosti

- **Latinica → Ћирилица** i **Ћирилица → Latinica**
- **YUSCII → Latinica** (JUS I.B1.002: `^ ~ ] } [ { @ \` \ |` → Č č Ć ć Š š Ž ž Đ đ)
- **YUSCII → Ћирилица** (JUS I.B1.003: latinična slova → ćirilica, `Q W X` → `Љ Њ Џ`)
  – YUSCII znaci se menjaju samo unutar reči, pa zagrade i sl. ostaju netaknuti
- Opseg: označeni tekst ili ceo dokument
- Tekst dokumenti (formatiranje se zadržava), tabele (ćelije sa formulama se preskaču), prezentacije
- Ispravno obrađuje digrafe `lj`, `nj`, `dž` (i verzal: `LJ` / `Lj`)
- Izuzeci gde digraf nije jedno slovo: *injekcija, konjugacija, konjunktura, vanjezički, podžupan, nadživeti*…
- „I“ na početku rečenice („I tako…“) i u tekstu pisanom verzalom („MAMA I TATA“, „MI SMO“) je veznik/reč,
  a ne rimski broj; „Petar I“, „I. Uvod“, „klasa I“ ostaju rimski brojevi
- Opcije:
  - preskakanje stranih reči (sa q, w, x, y)
  - preskakanje rimskih brojeva (XIV, MMXXVI)
  - preskakanje veb adresa i e‑pošte
  - `dj` → `ђ` (uz izuzetke *odjednom, podjednako, nadjačati*…)
  - **ne preslovljavaj merne jedinice** – `5 km`, `90 km/h`, `20 m²`, `22 °C`, `64 GB` ostaju latinicom
    (jednoslovne jedinice poput `m`, `s`, `t` samo iza broja, da predlozi „s“, „u“ ostanu ispravni)
- **Korisnički izuzeci** – reči i izrazi koji se nikada ne preslovljavaju, jedan po redu;
  `*` na kraju obuhvata sve oblike (`Microsoft*` → i *Microsoftov*). Važe u oba smera.

### Šta LibreOffice verzija ima, a ooOTranslit nema

ooOTranslit 0.4.1 menja slova slepo (pretraga i zamena), pa preslovljava i ono što ne treba. Translit za LibreOffice:

| | ooOTranslit 0.4.1 | Translit |
|---|---|---|
| Merne jedinice (`5 km`, `90 km/h`, `22 °C`, `64 GB`) | preslovljava | ostaju latinicom |
| Veb adrese i e‑pošta | preslovljava | ostaju |
| Strane reči (q, w, x, y), rimski brojevi | preslovljava | ostaju |
| Korisnički izuzeci, `dj` → `ђ` | – | ima |
| Formatiranje pri preslovljavanju označenog teksta | gubi se | čuva se (i kada se digraf prostire preko dva formata) |
| Impress i Draw | – | oblici, tabele, beleške |
| Fusnote, zaglavlja, okviri, oblici u Writer-u | delimično | da |
| Praćenje izmena | – | obrisani tekst se ne dira |
| Poništavanje | po jedno slovo | jedan korak za celo preslovljavanje |

## Instalacija

Preuzmite pakete sa stranice [Releases](https://github.com/digikv/onlyoffice-translit/releases/latest)
ili ih napravite sami (vidi dole).

### ONLYOFFICE – `translit.plugin`

**ONLYOFFICE Desktop Editors:** Dodaci → Upravljanje dodacima → Dodaj (Add plugin) → izaberite `translit.plugin`.

Ručno: iskopirajte sadržaj foldera `translit` u
`%LOCALAPPDATA%\ONLYOFFICE\DesktopEditors\data\sdkjs-plugins\{7c2f4b1e-8d3a-4f6e-9b21-5a0c3e8d7f14}\`
i ponovo otvorite dokument.

**Document Server:** kopirajte folder u `/var/www/onlyoffice/documentserver/sdkjs-plugins/{7c2f4b1e-8d3a-4f6e-9b21-5a0c3e8d7f14}/`.

### LibreOffice – `translit.oxt`

Dvoklik na `translit.oxt`, ili **Alatke → Upravljač proširenjima → Dodaj**, pa ponovo pokrenite LibreOffice.
Iz komandne linije (LibreOffice mora biti zatvoren): `unopkg add translit.oxt`.

Radi u LibreOffice 6.0 i novijim (Writer, Calc, Impress, Draw), na Windows-u, Linux-u i macOS-u;
koristi Python koji dolazi uz LibreOffice. Može se instalirati uporedo sa ooOTranslit-om.

Upotreba:

- meni **Alatke → Translit** i traka sa alatkama **Translit** (Pogled → Trake sa alatkama → Translit)
- preslovljava se **označeni tekst**; ako ništa nije označeno – **ceo dokument**
  (u Calc-u se uvek preslovljavaju označene ćelije; ceo dokument = svi listovi uz opciju „Uvek ceo dokument“)
- **Podešavanja i izuzeci…** otvara dijalog sa opcijama, izuzecima i dugmadima za preslovljavanje;
  podešavanja se čuvaju u `translit.json` u korisničkom profilu LibreOffice-a
- celo preslovljavanje se poništava jednim **Ctrl+Z**

### Microsoft Office – `translit-msoffice.zip`

Za Word, Excel i PowerPoint 2010 ili novije (i Microsoft 365) na Windows-u.

1. Raspakujte `translit-msoffice.zip`.
2. Zatvorite Word, Excel i PowerPoint.
3. Pokrenite `install.cmd` (dvoklik).

Instalacija na tom računaru pravi `Translit.dotm` (Word, folder STARTUP), `Translit.xlam` i `Translit.ppam`
(Excel i PowerPoint, folder `%APPDATA%\Microsoft\AddIns`) i uključuje ih. Da bi uvezla VBA kod, privremeno
uključuje opciju *Trust access to the VBA project object model* i posle je vraća kako je bila.
Obe lokacije su podrazumevano „pouzdane“, pa makroi rade bez upozorenja. Samo jedan program:
`install.cmd -Apps Word`. Uklanjanje: `uninstall.cmd`.

Upotreba – kartica **Translit** na traci:

- **Preslovi**: Latinica → Ћирилица, Ћирилица → Latinica, YUSCII → Latinica, YUSCII → Ћирилица
- preslovljava se **označeni tekst** (ili označeni oblici); ako ništa nije označeno – **ceo dokument**
  (Word: telo, tabele, zaglavlja i podnožja, fusnote, tekst-polja; PowerPoint: svi slajdovi, tabele, grupe i beleške).
  U Excel-u se preslovljavaju označene ćelije sa tekstom (formule i brojevi se preskaču), a uz opciju
  **Uvek ceo dokument** svi listovi
- **Opcije** su potvrdni okviri na kartici i važe za sva tri programa
- **Izuzeci…** otvara spisak u Notepad-u (`%APPDATA%\Translit\izuzeci.txt`, jedna reč ili izraz po redu)
- formatiranje se čuva; u Word-u se celo preslovljavanje poništava jednim **Ctrl+Z**, a u Excel-u
  **Poništi Translit** vraća prethodni tekst ćelija (bez formatiranja unutar ćelije)
- makroi za prečice sa tastature: `TranslitLatinicaUCirilicu`, `TranslitCirilicaULatinicu`,
  `TranslitYusciiULatinicu`, `TranslitYusciiUCirilicu`

Pravila koriste `VBScript.RegExp`, koji je deo Windows-a (u Microsoft 365 je ugrađen u VBA).

## Struktura

```
translit/                    ONLYOFFICE dodatak
  config.json                opis dodatka
  index.html                 panel
  scripts/translit.js        pravila preslovljavanja (nezavisno od ONLYOFFICE-a)
  scripts/code.js            veza sa ONLYOFFICE API-jem
  resources/                 ikonice
libreoffice/                 LibreOffice proširenje
  description.xml            opis proširenja
  Addons.xcu                 meni Alatke → Translit, traka sa alatkama, ikonice
  Office/UI/*WindowState.xcu naziv trake sa alatkama
  Translit.py                UNO komponenta: dijalog, Writer/Calc/Impress/Draw
  pythonpath/srtranslit.py   pravila preslovljavanja – Python verzija translit.js
  icons/                     ikonice
msoffice/                    Microsoft Office dodatak (VBA)
  install.ps1, install.cmd   pravi i instalira Translit.dotm / .xlam / .ppam
  uninstall.cmd              uklanja ih
  ribbon/customUI14.xml      kartica Translit i ikonice
  src/TranslitCore.bas       pravila preslovljavanja – VBA verzija translit.js
  src/TranslitWord.bas       Word        src/TranslitExcel.bas       Excel
  src/TranslitPowerPoint.bas PowerPoint  src/TranslitShapes.bas      tekst u oblicima
  src/TranslitSettings.bas   opcije, izuzeci, dugmad kartice
  src/TranslitNative.bas     NFC normalizacija (Windows API)
tests/
  parity_test.py             poredi translit.js i srtranslit.py (+ poznati primeri)
  vba_core_test.py           poredi TranslitCore.bas (preko VBScript-a) sa njima
  vba_syntax_test.py         sintaksa svih VBA modula, ASCII i CRLF
  install_test.ps1           install.ps1: sintaksa i ubacivanje kartice u paket
  lo_test.py                 test u pravim LibreOffice dokumentima
  lo_dispatch_test.py        test instaliranog proširenja (komande menija)
tools/
  build.py                   pravi translit.plugin, translit.oxt i translit-msoffice.zip
  make_icons.ps1             pravi ikonice
```

Pravila preslovljavanja postoje u tri kopije – `translit/scripts/translit.js`,
`libreoffice/pythonpath/srtranslit.py` i `msoffice/src/TranslitCore.bas`. Kada menjate jednu, promenite
i ostale i pokrenite `python tests/parity_test.py` i `python tests/vba_core_test.py` (potreban Node.js):
proveravaju poznate primere i porede sve tri verzije na nekoliko hiljada nasumičnih tekstova.
VBA moduli moraju ostati čist ASCII sa CRLF krajevima redova (ćirilica se piše kao `ChrW`/`\uXXXX`).

ONLYOFFICE dodatak koristi plugin SDK (`../v1/plugins.js`) koji je već ugrađen u ONLYOFFICE, pa radi i bez interneta.

## Pravljenje paketa

```
python tools/build.py
```

pravi `translit.plugin` (ZIP sadržaja foldera `translit`), `translit.oxt` (ZIP sadržaja foldera `libreoffice`)
i `translit-msoffice.zip` (folder `msoffice`).

## Testovi

```
python tests/parity_test.py
python tests/vba_core_test.py
python tests/vba_syntax_test.py
powershell -ExecutionPolicy Bypass -File tests/install_test.ps1
```

VBA jezgro se testira bez Office-a, kroz Windows-ov VBScript (isti `RegExp`). Delovi koji rade sa
Word-om, Excel-om i PowerPoint-om (i sama instalacija u Office) nemaju automatske testove – njih treba
isprobati u Office-u.

Testovi u LibreOffice-u koriste zaseban profil i LibreOffice bez prozora:

```
soffice -env:UserInstallation=file:///C:/tmp/lo_test --headless --norestore "--accept=socket,host=127.0.0.1,port=2083;urp;"
"C:\Program Files\LibreOffice\program\python.exe" tests/lo_test.py
```

`tests/lo_dispatch_test.py` proverava instalirano proširenje (prethodno
`unopkg add -env:UserInstallation=file:///C:/tmp/lo_test translit.oxt`).

## Licenca

[MIT](LICENSE)
