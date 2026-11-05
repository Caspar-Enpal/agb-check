# AGB-Check: Hinweise für Claude Code

Dieses Repo prüft, welcher AGB-Stand (`*_AGB_Stromtarif_Enpal.OnePlus.pdf`) in Enpal.One+ Kundenordnern liegt. Details in README.md.

## Wenn der User einen SharePoint-Link schickt

1. Abhängigkeiten sicherstellen: `py -m pip install --user -r requirements.txt`
2. Ausführen (Link in Anführungszeichen, `-P` damit kein Code aus dem aktuellen Ordner geladen wird):
   `py -P agb_check.py "<SharePoint-Link>" --oeffnen`
3. Meldet das Programm „kein synchronisierter Ordner gefunden“: Der User muss den Ordner in SharePoint öffnen und auf **Synchronisieren** klicken (oder einen übergeordneten Ordner). Danach erneut starten. Es gibt keinen Zugriff über die SharePoint-API, das Programm liest nur den lokal synchronisierten Ordner.
4. Ergebnis zusammenfassen: geprüfte Kunden, neu/alt/nicht lesbar, Quote, Pfad der Excel-Datei. Keine Kundennamen in den Chat schreiben, die stehen in der Excel-Datei.

## Fachliche Regeln (nur nach Rücksprache ändern)

- Stichtag: Datei in SharePoint geändert ab **12.05.2026** (`--stichtag`). Das Datum im Ordnernamen ist bei Upselling-Verträgen die Erstunterschrift und wird nicht verwendet.
- Neue AGB-Stände: **12.05.2026, 31.07.2026** (`--neue-staende`). Alles andere ist „alt“.
- Bei mehreren AGB-Dateien im Ordner zählt die zuletzt geänderte.

## Datenschutz

- Excel-Ausgaben (`output/`, `*.xlsx`) enthalten Kundennamen und dürfen nie committet werden.
- Keine Kundendaten, Ordnernamen oder PDF-Inhalte in Commits, Issues oder Chats außerhalb des Teams.
