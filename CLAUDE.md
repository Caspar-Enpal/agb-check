# AGB-Check: Hinweise für Claude Code

Dieses Repo prüft, welcher AGB-Stand (`*_AGB_Stromtarif_Enpal.OnePlus.pdf`) in Enpal.One+ Kundenordnern liegt, liest die Kunden-E-Mail aus `*_Auftrag_Stromtarif_Enpal.OnePlus.pdf` und erstellt eine Excel-Datei. Details in README.md.

## Standard-Auftrag

Typische Anfrage: „Hier ist mein Pfad zu meinem Ordner für alle Kunden, Stichtag ist der 12.05.2026, bitte erstell mir die Excel.“

Dann genau so vorgehen, ohne Rückfragen, solange Pfad und Stichtag angegeben sind:

1. **Einrichtung prüfen (nur beim ersten Mal nötig):**
   - `py --version` muss Python 3.11 oder neuer zeigen. Fehlt Python, abbrechen und den User bitten, Python zu installieren.
   - `py -m pip install --user -r requirements.txt`
2. **Quelle bestimmen:** Der User gibt entweder einen lokalen Ordnerpfad oder einen SharePoint-Link. Beides direkt an das Programm übergeben, das Programm löst SharePoint-Links selbst in den synchronisierten Ordner auf.
3. **Stichtag übernehmen:** Das Datum des Users als `--stichtag TT.MM.JJJJ` übergeben. Gemeint ist: Geprüft werden alle Kunden, deren AGB-Datei **ab** diesem Datum in SharePoint hochgeladen bzw. geändert wurde. Ältere werden nur gezählt. Ohne Angabe gilt der 12.05.2026.
4. **Ausführen** (Pfad oder Link in Anführungszeichen):
   ```
   py -P agb_check.py "<Pfad oder Link>" --stichtag <TT.MM.JJJJ> --oeffnen
   ```
   Bei großen Ordnern (Tausende Kunden) im Hintergrund laufen lassen. Der erste Lauf dauert einige Minuten, weil OneDrive die PDFs herunterlädt.
5. **Ergebnis melden**, kurz und ohne Kundendaten:
   - Ordner gesamt, vor Stichtag (nicht geprüft), ohne AGB-Datei, geprüft
   - neu / alt / nicht lesbar und die Quote
   - wie viele E-Mail-Adressen gefunden wurden
   - welche alten AGB-Stände vorkommen (steht im Blatt „Übersicht“)
   - Pfad der Excel-Datei (liegt in `output/`)

Die Excel-Datei hat die Blätter „Übersicht“, „Kunden geprüft“, „Alte AGB“, „Ohne AGB“ und bei Bedarf „Nicht lesbar“. Die Spalte E-Mail ist standardmäßig dabei. Nur wenn der User ausdrücklich keine E-Mails will: `--ohne-email`.

## Wenn etwas nicht klappt

- **„kein synchronisierter Ordner gefunden“:** Der User muss den Ordner in SharePoint öffnen und auf **Synchronisieren** klicken (oder einen übergeordneten Ordner), danach erneut starten. Es gibt keinen Zugriff über die SharePoint-API. Das Programm liest nur den lokal synchronisierten Ordner. Keine Umwege über Tokens, Browser-Cookies oder andere Zugänge bauen.
- **Alle Ordner landen in „Ohne AGB“:** Der Ordner ist anders aufgebaut als erwartet (ein Unterordner pro Kunde mit den PDFs darin). Struktur prüfen, ohne Kundennamen auszugeben, z. B. nur Anzahl Ordner und Dateinamen-Muster, und den User fragen.
- **Viele „nicht lesbar“:** Stichprobe ansehen, ob der Text „Stand: TT.MM.JJJJ“ anders formatiert ist oder die PDFs gescannt sind.

## Fachliche Regeln (nur nach Rücksprache ändern)

- Stichtag: Datei in SharePoint geändert ab **12.05.2026** (`--stichtag`). Das Datum im Ordnernamen ist bei Upselling-Verträgen die Erstunterschrift und wird nicht verwendet.
- Neue AGB-Stände: **12.05.2026, 31.07.2026** (`--neue-staende`). Alles andere ist „alt“.
- Bei mehreren AGB-Dateien im Ordner zählt die zuletzt geänderte.
- Die E-Mail steht im Auftrag oben links auf Seite 1. Die Enpal-Adresse im Fuß wird ignoriert.

## Datenschutz

- Excel-Ausgaben (`output/`, `*.xlsx`) enthalten Kundennamen und E-Mail-Adressen und dürfen nie committet werden.
- E-Mail-Adressen und Kundennamen nie in den Chat ausgeben. Prüfungen nur als Zählungen (gefunden / nicht gefunden).
- Keine Kundendaten, Ordnernamen oder PDF-Inhalte in Commits, Issues oder Chats außerhalb des Teams.
