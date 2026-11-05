# AGB-Check Enpal.One+ Stromtarif

Prüft für alle Kundenordner in „01_Auftragsübersichten Enpal.One+“, welcher Stand der Datei `*_AGB_Stromtarif_Enpal.OnePlus.pdf` abgelegt ist, und erstellt eine Excel-Datei mit der Neu-/Alt-Quote.

## So funktioniert die Prüfung

1. Jeder Unterordner ist ein Kunde.
2. Maßgeblich ist das Datum, an dem die AGB-Datei **in SharePoint zuletzt geändert/hochgeladen** wurde. Das Datum im Ordnernamen wird nicht verwendet, weil es bei Upselling-Verträgen das Datum der Erstunterschrift ist.
3. AGB-Dateien, die vor dem Stichtag hochgeladen wurden (Standard: **12.05.2026**), werden übersprungen.
4. Bei allen anderen liest das Programm auf Seite 1 oben links den Text „Stand: TT.MM.JJJJ“.
   - **12.05.2026** oder **31.07.2026** → `neu`
   - jedes andere Datum → `alt`
   - kein Stand gefunden → `nicht lesbar`
5. Liegen mehrere AGB-Dateien in einem Ordner, zählt die zuletzt hochgeladene.

Quote = neu / (neu + alt). Nicht lesbare Dateien fließen nicht in die Quote ein.

## Einrichtung (einmalig)

1. Python 3.11 oder neuer installieren.
2. In SharePoint den Ordner `14_Vertragsunterlagen/01_Auftragsübersichten Enpal.One+` öffnen und auf **Synchronisieren** klicken. Er erscheint dann z. B. unter
   `C:\Users\<name>\Enpal B.V\Venture Development - 01_Auftragsübersichten Enpal.One+`
3. Abhängigkeiten installieren:
   ```
   py -m pip install --user -r requirements.txt
   ```

## Ausführen

```
py agb_check.py --root "C:\Users\<name>\Enpal B.V\Venture Development - 01_Auftragsübersichten Enpal.One+"
```

Optionen:

| Option | Standard | Bedeutung |
|---|---|---|
| `--stichtag` | `12.05.2026` | Nur AGB-Dateien prüfen, die ab diesem Datum hochgeladen wurden |
| `--neue-staende` | `12.05.2026,31.07.2026` | AGB-Stände, die als „neu“ gelten |
| `--ausgabe` | `AGB_Check_<heute>.xlsx` | Pfad der Excel-Datei |
| `--threads` | `8` | Parallele Downloads |

Kommt ein neuer AGB-Stand dazu, einfach `--neue-staende` erweitern.

## Excel-Ausgabe

| Blatt | Inhalt |
|---|---|
| Übersicht | Zahlen, Quote, Verteilung der alten Stände |
| Kunden geprüft | Alle Kunden ab Stichtag mit Upload-Datum, AGB-Stand, Status |
| Alte AGB | Nur die Kunden mit altem Stand |
| Nicht lesbar | Nur vorhanden, wenn ein Stand nicht gelesen werden konnte |
| Ohne AGB | Kundenordner ohne AGB_Stromtarif-Datei |

## Hinweise

- **Datenschutz:** Die Excel-Datei enthält Kundennamen (Ordnernamen). Sie darf nicht ins Repository und nur intern weitergegeben werden. `*.xlsx` und `output/` stehen deshalb in der `.gitignore`.
- Bei „Dateien bei Bedarf“ lädt OneDrive die geprüften PDFs herunter. Danach kann man über Rechtsklick auf den Ordner → **Speicherplatz freigeben** wieder Platz schaffen.
