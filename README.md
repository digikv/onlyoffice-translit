# Translit za ONLYOFFICE

Dodatak za ONLYOFFICE koji preslovljava srpski tekst latinica ↔ ćirilica, po uzoru na ooOTranslit za LibreOffice/OpenOffice.

*Serbian Latin ↔ Cyrillic (and legacy YUSCII) transliteration plugin for ONLYOFFICE, inspired by ooOTranslit for LibreOffice/OpenOffice.*

## Mogućnosti

- **Latinica → Ћирилица** i **Ћирилица → Latinica**
- **YUSCII → Latinica** (JUS I.B1.002: `^ ~ ] } [ { @ \` \ |` → Č č Ć ć Š š Ž ž Đ đ)
- **YUSCII → Ћирилица** (JUS I.B1.003: latinična slova → ćirilica, `Q W X` → `Љ Њ Џ`)
  – YUSCII znaci se menjaju samo unutar reči, pa zagrade i sl. ostaju netaknuti
- Opseg: označeni tekst ili ceo dokument
- Tekst dokumenti (formatiranje se zadržava), tabele (ćelije sa formulama se preskaču), prezentacije
- Ispravno obrađuje digrafe `lj`, `nj`, `dž` (i verzal: `LJ` / `Lj`)
- Izuzeci gde digraf nije jedno slovo: *injekcija, konjugacija, konjunktura, vanjezički, podžupan, nadživeti*…
- Opcije:
  - preskakanje stranih reči (sa q, w, x, y)
  - preskakanje rimskih brojeva (XIV, MMXXVI)
  - preskakanje veb adresa i e‑pošte
  - `dj` → `ђ` (uz izuzetke *odjednom, podjednako, nadjačati*…)
  - **ne preslovljavaj merne jedinice** – `5 km`, `90 km/h`, `20 m²`, `22 °C`, `64 GB` ostaju latinicom
    (jednoslovne jedinice poput `m`, `s`, `t` samo iza broja, da predlozi „s“, „u“ ostanu ispravni)
- **Korisnički izuzeci** – reči i izrazi koji se nikada ne preslovljavaju, jedan po redu;
  `*` na kraju obuhvata sve oblike (`Microsoft*` → i *Microsoftov*). Važe u oba smera.

## Instalacija

Preuzmite `translit.plugin` sa stranice [Releases](https://github.com/digikv/onlyoffice-translit/releases/latest).

**ONLYOFFICE Desktop Editors:** Dodaci → Upravljanje dodacima → Dodaj (Add plugin) → izaberite `translit.plugin`.

Ručno: iskopirajte sadržaj foldera `translit` u
`%LOCALAPPDATA%\ONLYOFFICE\DesktopEditors\data\sdkjs-plugins\{7c2f4b1e-8d3a-4f6e-9b21-5a0c3e8d7f14}\`
i ponovo otvorite dokument.

**Document Server:** kopirajte folder u `/var/www/onlyoffice/documentserver/sdkjs-plugins/{7c2f4b1e-8d3a-4f6e-9b21-5a0c3e8d7f14}/`.

## Struktura

```
translit/
  config.json          opis dodatka
  index.html           panel
  scripts/translit.js  pravila preslovljavanja (nezavisno od ONLYOFFICE-a)
  scripts/code.js      veza sa ONLYOFFICE API-jem
  resources/           ikonice
```

Dodatak koristi ONLYOFFICE plugin SDK (`../v1/plugins.js`) koji je već ugrađen u ONLYOFFICE, pa radi i bez interneta.

## Pravljenje paketa

`.plugin` fajl je običan ZIP sadržaja foldera `translit`:

```powershell
Compress-Archive -Path translit\* -DestinationPath translit.zip; Rename-Item translit.zip translit.plugin
```

## Licenca

[MIT](LICENSE)
