"""
AGB-Check Enpal.One+ Stromtarif
================================

Prueft, welcher AGB-Stand in den Kundenordnern abgelegt ist, und schreibt das
Ergebnis in eine Excel-Datei.

Logik
-----
1. Jeder Unterordner von --root ist ein Kunde.
2. Im Ordner wird die Datei *AGB_Stromtarif*.pdf gesucht.
3. Massgeblich fuer "seit wann ist das Dokument da" ist das Aenderungsdatum der
   Datei in SharePoint (= "Geaendert" in SharePoint bzw. LastWriteTime im
   synchronisierten OneDrive-Ordner). Das Datum im Ordnernamen wird NICHT
   verwendet, weil es bei Upselling-Vertraegen die Erstunterschrift ist.
4. Dateien, die vor --stichtag zuletzt geaendert wurden, werden nicht gelesen
   (sie koennen die neuen AGB noch nicht haben).
5. Bei allen anderen wird auf Seite 1 der Text "Stand: TT.MM.JJJJ" gelesen.
   Steht dort eines der --neue-staende, gilt der Kunde als "neu", sonst "alt".

Datenschutz: Das Programm arbeitet nur lokal. Die Excel-Datei enthaelt
Kundenordner-Namen und gehoert NICHT ins Git-Repository.

Aufruf (Beispiel)
-----------------
    py agb_check.py --root "C:\\Users\\<name>\\Enpal B.V\\Venture Development - 01_Auftragsübersichten Enpal.One+"

Alle Optionen: py agb_check.py --help
"""

from __future__ import annotations

import argparse
import datetime as dt
import os
import re
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from urllib.parse import parse_qs, unquote, urlparse

import pymupdf
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

AGB_PATTERN = "AGB_Stromtarif"
AUFTRAG_PATTERN = "Auftrag_Stromtarif"
STAND_RE = re.compile(r"Stand\s*:?\s*(\d{1,2})\.(\d{1,2})\.(\d{4})")
KUNDEN_ID_RE = re.compile(r"^(DE\d+)_")
EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+")

STATUS_NEU = "neu"
STATUS_ALT = "alt"
STATUS_UNKLAR = "nicht lesbar"


@dataclass
class AgbDatei:
    pfad: str
    name: str
    geaendert: dt.datetime
    stand: str | None = None
    fehler: str | None = None


@dataclass
class Kunde:
    ordner: str
    pfad: str
    agb_dateien: list[AgbDatei] = field(default_factory=list)
    letzte_aenderung_ordner: dt.datetime | None = None
    auftrag_pfad: str | None = None
    auftrag_geaendert: dt.datetime | None = None
    email: str = ""
    email_hinweis: str = ""

    @property
    def aktuelle_agb(self) -> AgbDatei | None:
        """Die zuletzt geaenderte AGB-Datei ist die massgebliche."""
        if not self.agb_dateien:
            return None
        return max(self.agb_dateien, key=lambda a: a.geaendert)

    @property
    def kunden_id(self) -> str:
        for a in self.agb_dateien:
            m = KUNDEN_ID_RE.match(a.name)
            if m:
                return m.group(1)
        return ""


def parse_datum(s: str) -> dt.date:
    for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(s.strip(), fmt).date()
        except ValueError:
            pass
    raise argparse.ArgumentTypeError(f"Datum nicht erkannt: {s!r} (erwartet TT.MM.JJJJ)")


def _sync_basisordner() -> list[str]:
    """Ordner, unter denen OneDrive SharePoint-Bibliotheken synchronisiert (z. B. 'C:\\Users\\x\\Enpal B.V')."""
    home = os.path.expanduser("~")
    basen = []
    with os.scandir(home) as it:
        for e in it:
            if e.is_dir() and ("enpal" in e.name.lower() or e.name.lower().startswith("onedrive")):
                basen.append(e.path)
    return basen


def sharepoint_link_aufloesen(link: str) -> str:
    """Findet zu einem SharePoint-Ordnerlink den lokal synchronisierten Ordner.

    Voraussetzung: Der Ordner (oder ein übergeordneter Ordner) ist per OneDrive
    'Synchronisieren' bzw. 'Verknüpfung zu Meine Dateien hinzufügen' eingebunden.
    """
    url = urlparse(link)
    query = parse_qs(url.query)
    # Browser-Ansichtslinks (.../Forms/AllItems.aspx?id=/sites/...) tragen den Ordner im Parameter 'id'.
    pfad = query["id"][0] if "id" in query else url.path
    # Freigabelinks: /:f:/r/sites/<site>/<bibliothek>/<ordner>/...
    pfad = re.sub(r"^/:[a-z]:/[a-z]/", "/", pfad)
    segmente = [unquote(s) for s in pfad.split("/") if s]
    # Ansicht einer ganzen Bibliothek: .../<bibliothek>/Forms/AllItems.aspx -> nur <bibliothek>
    if "Forms" in segmente:
        segmente = segmente[: segmente.index("Forms")]
    if "sites" in segmente:
        segmente = segmente[segmente.index("sites") + 2:]  # 'sites' und Site-Name weg
    if not segmente:
        raise ValueError(f"Kein Ordnerpfad im Link erkannt: {link}")

    basen = _sync_basisordner()
    # Vom tiefsten Segment nach oben suchen, welches lokal synchronisiert ist.
    for i in range(len(segmente) - 1, -1, -1):
        name = segmente[i]
        for basis in basen:
            with os.scandir(basis) as it:
                for e in it:
                    if e.is_dir() and (e.name == name or e.name.endswith(" - " + name)):
                        ziel = os.path.join(e.path, *segmente[i + 1:])
                        if os.path.isdir(ziel):
                            return ziel
    raise FileNotFoundError(
        "Zu diesem SharePoint-Link wurde kein synchronisierter Ordner gefunden.\n"
        "Bitte den Ordner in SharePoint öffnen und auf 'Synchronisieren' klicken, "
        f"dann erneut starten. Gesuchte Ordnernamen: {', '.join(segmente)}\n"
        f"Durchsuchte Orte: {', '.join(basen) or '(keine)'}"
    )


def _norm(name: str) -> str:
    """'AGB-Stromtarif', 'AGB Stromtarif', 'agb_stromtarif' -> 'agb_stromtarif'."""
    return re.sub(r"[\s\-.]+", "_", name.lower())


def _dateityp(dateiname: str) -> str | None:
    if not dateiname.lower().endswith(".pdf"):
        return None
    n = _norm(dateiname)
    if _norm(AGB_PATTERN) in n:
        return "agb"
    if _norm(AUFTRAG_PATTERN) in n:
        return "auftrag"
    return None


def ordner_scannen(root: str) -> list[Kunde]:
    """Findet Kundenordner in beliebiger Tiefe.

    Ein Ordner gilt als Kunde, wenn direkt darin eine AGB_Stromtarif- oder
    Auftrag_Stromtarif-PDF liegt. Darunter wird nicht weiter gesucht.
    Ein Ordner ganz unten in der Struktur (ohne Unterordner) ohne diese PDFs
    gilt als Kunde ohne AGB. Ordner dazwischen (z. B. Monats- oder
    Batch-Ordner) sind nur Sortierebenen und werden durchsucht.
    Der Kundenname in der Excel ist der Pfad relativ zum Startordner.
    """
    root = os.path.abspath(root)
    root_name = os.path.basename(root.rstrip("\\/"))
    kunden = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = sorted(d for d in dirnames if not d.startswith("."))
        typen = {f: _dateityp(f) for f in filenames}
        ist_kunde = any(typen.values())
        ist_blatt = not dirnames and dirpath != root
        if not ist_kunde and not ist_blatt:
            continue
        rel = os.path.relpath(dirpath, root)
        k = Kunde(ordner=root_name if rel == "." else rel, pfad=dirpath)
        letzte = None
        for f in filenames:
            pfad = os.path.join(dirpath, f)
            try:
                mtime = dt.datetime.fromtimestamp(os.stat(pfad).st_mtime)
            except OSError:
                continue
            letzte = mtime if letzte is None or mtime > letzte else letzte
            if typen[f] == "agb":
                k.agb_dateien.append(AgbDatei(pfad=pfad, name=f, geaendert=mtime))
            elif typen[f] == "auftrag":
                # Bei mehreren Auftragsdateien zählt die zuletzt geänderte.
                if k.auftrag_geaendert is None or mtime > k.auftrag_geaendert:
                    k.auftrag_pfad, k.auftrag_geaendert = pfad, mtime
        k.letzte_aenderung_ordner = letzte
        kunden.append(k)
        if ist_kunde:
            dirnames[:] = []  # Unterordner eines Kunden nicht als eigene Kunden zählen
    return kunden


def struktur_info(kunden: list[Kunde], root: str) -> dict[int, int]:
    """Anzahl Kundenordner je Ordnerebene unter dem Startordner (für die Übersicht)."""
    ebenen: dict[int, int] = {}
    for k in kunden:
        e = 0 if os.path.abspath(k.pfad) == os.path.abspath(root) else os.path.relpath(k.pfad, root).count(os.sep) + 1
        ebenen[e] = ebenen.get(e, 0) + 1
    return dict(sorted(ebenen.items()))


def datei_lesen(pfad: str) -> bytes:
    # Bei OneDrive "Dateien bei Bedarf" laedt das Lesen die Datei herunter.
    with open(pfad, "rb") as fh:
        return fh.read()


def stand_auslesen(pdf_bytes: bytes) -> str | None:
    with pymupdf.open(stream=pdf_bytes, filetype="pdf") as doc:
        if doc.page_count == 0:
            return None
        seite = doc[0]
        r = seite.rect
        # Zuerst nur oben links (dort steht der Stand), dann ganze Seite 1.
        for clip in (pymupdf.Rect(0, 0, r.width / 2, r.height / 4), None):
            text = seite.get_text("text", clip=clip)
            m = STAND_RE.search(text)
            if m:
                d, mo, y = m.groups()
                return f"{int(d):02d}.{int(mo):02d}.{y}"
    return None


def pruefen(dateien: list[AgbDatei], threads: int) -> None:
    gesamt = len(dateien)
    fertig = 0
    with ThreadPoolExecutor(max_workers=threads) as pool:
        futures = {pool.submit(datei_lesen, a.pfad): a for a in dateien}
        for fut in as_completed(futures):
            a = futures[fut]
            try:
                # PyMuPDF wird nur im Haupt-Thread benutzt (nicht thread-sicher).
                a.stand = stand_auslesen(fut.result())
                if a.stand is None:
                    a.fehler = "Kein 'Stand: TT.MM.JJJJ' auf Seite 1 gefunden"
            except Exception as ex:  # noqa: BLE001
                a.fehler = f"{type(ex).__name__}: {ex}"
            fertig += 1
            if fertig % 100 == 0 or fertig == gesamt:
                print(f"  {fertig}/{gesamt} PDFs gelesen", flush=True)


def email_auslesen(pdf_bytes: bytes) -> str | None:
    """Kunden-E-Mail aus dem Auftrag: steht oben links auf Seite 1.

    Unten im Fuß steht die Enpal-Adresse, die wird ignoriert.
    """
    with pymupdf.open(stream=pdf_bytes, filetype="pdf") as doc:
        if doc.page_count == 0:
            return None
        seite = doc[0]
        r = seite.rect
        for clip in (pymupdf.Rect(0, 0, r.width / 2, r.height / 3), None):
            for m in EMAIL_RE.findall(seite.get_text("text", clip=clip)):
                if "enpal" not in m.lower().split("@")[-1]:
                    return m.strip(".")
    return None


def emails_lesen(kunden: list[Kunde], threads: int) -> None:
    mit_auftrag = [k for k in kunden if k.auftrag_pfad]
    for k in kunden:
        if not k.auftrag_pfad:
            k.email_hinweis = "Keine Auftrag_Stromtarif-Datei im Ordner"
    gesamt = len(mit_auftrag)
    fertig = 0
    with ThreadPoolExecutor(max_workers=threads) as pool:
        futures = {pool.submit(datei_lesen, k.auftrag_pfad): k for k in mit_auftrag}
        for fut in as_completed(futures):
            k = futures[fut]
            try:
                k.email = email_auslesen(fut.result()) or ""
                if not k.email:
                    k.email_hinweis = "Keine E-Mail im Auftrag gefunden"
            except Exception as ex:  # noqa: BLE001
                k.email_hinweis = f"{type(ex).__name__}: {ex}"
            fertig += 1
            if fertig % 100 == 0 or fertig == gesamt:
                print(f"  {fertig}/{gesamt} Aufträge gelesen", flush=True)


def status(a: AgbDatei, neue: set[str]) -> str:
    if a.stand is None:
        return STATUS_UNKLAR
    return STATUS_NEU if a.stand in neue else STATUS_ALT


# ---------------------------------------------------------------- Excel ----

KOPF_FONT = Font(bold=True, color="FFFFFF")
KOPF_FILL = PatternFill("solid", fgColor="1F3864")
FILL = {
    STATUS_NEU: PatternFill("solid", fgColor="C6EFCE"),
    STATUS_ALT: PatternFill("solid", fgColor="FFC7CE"),
    STATUS_UNKLAR: PatternFill("solid", fgColor="FFEB9C"),
}


def tabelle(ws, kopf: list[str], zeilen: list[list], status_spalte: int | None = None):
    ws.append(kopf)
    for c in ws[1]:
        c.font = KOPF_FONT
        c.fill = KOPF_FILL
    for z in zeilen:
        ws.append(z)
        if status_spalte is not None:
            ws.cell(ws.max_row, status_spalte + 1).fill = FILL.get(z[status_spalte], PatternFill())
    ws.freeze_panes = "A2"
    if zeilen:
        ws.auto_filter.ref = ws.dimensions
    for i, k in enumerate(kopf, 1):
        breite = max([len(str(k))] + [len(str(z[i - 1])) for z in zeilen[:2000] if z[i - 1] is not None])
        ws.column_dimensions[get_column_letter(i)].width = min(max(breite + 2, 10), 70)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            if isinstance(c.value, dt.datetime):
                c.number_format = "DD.MM.YYYY HH:MM"


def excel_schreiben(pfad: str, kunden: list[Kunde], args, neue: set[str]) -> dict:
    stichtag = dt.datetime.combine(args.stichtag, dt.time.min)
    geprueft, vor_stichtag, ohne_agb = [], [], []
    for k in kunden:
        a = k.aktuelle_agb
        if a is None:
            ohne_agb.append(k)
        elif a.geaendert < stichtag:
            vor_stichtag.append(k)
        else:
            geprueft.append(k)

    zaehler = {STATUS_NEU: 0, STATUS_ALT: 0, STATUS_UNKLAR: 0}
    zeilen = []
    for k in sorted(geprueft, key=lambda k: k.aktuelle_agb.geaendert, reverse=True):
        a = k.aktuelle_agb
        s = status(a, neue)
        zaehler[s] += 1
        alle_staende = ", ".join(sorted({x.stand or "?" for x in k.agb_dateien})) if len(k.agb_dateien) > 1 else ""
        hinweis = "; ".join(h for h in (a.fehler, k.email_hinweis if args.mit_email else "") if h)
        zeilen.append([k.ordner, k.kunden_id, k.email, a.name, a.geaendert, a.stand or "", s,
                       len(k.agb_dateien), alle_staende, hinweis])
    STATUS_SPALTE = 6
    if not args.mit_email:
        zeilen = [z[:2] + z[3:] for z in zeilen]
        STATUS_SPALTE = 5

    bewertbar = zaehler[STATUS_NEU] + zaehler[STATUS_ALT]
    quote_neu = zaehler[STATUS_NEU] / bewertbar if bewertbar else 0
    quote_alt = zaehler[STATUS_ALT] / bewertbar if bewertbar else 0

    wb = Workbook()
    ws = wb.active
    ws.title = "Übersicht"
    ws["A1"] = "AGB-Check Enpal.One+ Stromtarif"
    ws["A1"].font = Font(bold=True, size=14)
    info = [
        ("Erstellt am", dt.datetime.now().strftime("%d.%m.%Y %H:%M")),
        ("Quellordner", args.root),
        ("Stichtag (Datei in SharePoint geändert ab)", args.stichtag.strftime("%d.%m.%Y")),
        ("Als 'neu' gewertete AGB-Stände", ", ".join(sorted(neue))),
        ("", ""),
        ("Kundenordner gesamt", len(kunden)),
        ("Kundenordner je Ordnerebene", ", ".join(f"Ebene {e}: {n}" for e, n in getattr(args, "struktur", {}).items())),
        ("davon AGB vor Stichtag hochgeladen (nicht geprüft)", len(vor_stichtag)),
        ("davon ohne AGB_Stromtarif-Datei", len(ohne_agb)),
        ("davon geprüft (AGB ab Stichtag)", len(geprueft)),
        ("", ""),
        ("Neue AGB", zaehler[STATUS_NEU]),
        ("Alte AGB", zaehler[STATUS_ALT]),
        ("Nicht lesbar", zaehler[STATUS_UNKLAR]),
        ("", ""),
        ("Quote neue AGB (neu / (neu + alt))", quote_neu),
        ("Quote alte AGB (alt / (neu + alt))", quote_alt),
    ]
    for i, (k, v) in enumerate(info, start=3):
        ws.cell(i, 1, k)
        c = ws.cell(i, 2, v)
        if isinstance(v, float):
            c.number_format = "0.0%"
            ws.cell(i, 1).font = Font(bold=True)
            c.font = Font(bold=True)
    # Verteilung der alten Stände
    alt_staende: dict[str, int] = {}
    for z in zeilen:
        if z[STATUS_SPALTE] == STATUS_ALT:
            stand = z[STATUS_SPALTE - 1]
            alt_staende[stand] = alt_staende.get(stand, 0) + 1
    r = len(info) + 5
    ws.cell(r, 1, "Gefundene alte Stände").font = Font(bold=True)
    for stand, n in sorted(alt_staende.items(), key=lambda x: -x[1]):
        r += 1
        ws.cell(r, 1, f"Stand {stand}")
        ws.cell(r, 2, n)
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 60

    kopf = ["Kundenordner", "Kunden-ID", "E-Mail", "AGB-Datei", "In SharePoint geändert am",
            "AGB-Stand", "Status", "Anzahl AGB-Dateien", "Alle Stände (bei mehreren)", "Hinweis"]
    if not args.mit_email:
        kopf.remove("E-Mail")
    S = STATUS_SPALTE
    tabelle(wb.create_sheet("Kunden geprüft"), kopf, zeilen, status_spalte=S)
    tabelle(wb.create_sheet("Alte AGB"), kopf, [z for z in zeilen if z[S] == STATUS_ALT], status_spalte=S)
    if zaehler[STATUS_UNKLAR]:
        tabelle(wb.create_sheet("Nicht lesbar"), kopf, [z for z in zeilen if z[S] == STATUS_UNKLAR], status_spalte=S)
    ohne_kopf = ["Kundenordner", "E-Mail", "Letzte Änderung im Ordner", "Im Prüfzeitraum", "Hinweis"]
    ohne_zeilen = [[k.ordner, k.email, k.letzte_aenderung_ordner or "",
                    "ja" if k.letzte_aenderung_ordner and k.letzte_aenderung_ordner >= stichtag else "nein",
                    k.email_hinweis]
                   for k in sorted(ohne_agb, key=lambda k: k.ordner)]
    if not args.mit_email:
        ohne_kopf = [ohne_kopf[0]] + ohne_kopf[2:4]
        ohne_zeilen = [[z[0]] + z[2:4] for z in ohne_zeilen]
    tabelle(wb.create_sheet("Ohne AGB"), ohne_kopf, ohne_zeilen)
    wb.save(pfad)
    return {"gesamt": len(kunden), "vor_stichtag": len(vor_stichtag), "ohne_agb": len(ohne_agb),
            "geprueft": len(geprueft), **zaehler, "quote_neu": quote_neu, "quote_alt": quote_alt,
            "mit_email": sum(1 for k in geprueft if k.email)}


def main() -> int:
    p = argparse.ArgumentParser(description="Prüft den AGB-Stand (Stromtarif) in Enpal.One+ Kundenordnern.")
    p.add_argument("quelle", nargs="?", help="SharePoint-Ordnerlink ODER lokaler Pfad zum synchronisierten Ordner")
    p.add_argument("--root", help="Wie 'quelle' (alternative Schreibweise)")
    p.add_argument("--stichtag", type=parse_datum, default=dt.date(2026, 5, 12),
                   help="Nur AGB-Dateien prüfen, die ab diesem Datum in SharePoint geändert wurden (Default 12.05.2026)")
    p.add_argument("--neue-staende", default="12.05.2026,31.07.2026",
                   help="Kommagetrennte AGB-Stände, die als 'neu' gelten (Default 12.05.2026,31.07.2026)")
    p.add_argument("--ausgabe", default=None, help="Pfad der Excel-Datei (Default: output/AGB_Check_<datum_uhrzeit>.xlsx)")
    p.add_argument("--oeffnen", action="store_true", help="Excel-Datei nach dem Lauf direkt öffnen")
    p.add_argument("--threads", type=int, default=8, help="Parallele Downloads/Lesevorgänge (Default 8)")
    p.add_argument("--ohne-email", dest="mit_email", action="store_false",
                   help="Keine Kunden-E-Mails aus dem Auftrag auslesen (z. B. wenn nur die Quote gebraucht wird)")
    args = p.parse_args()
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass

    quelle = args.quelle or args.root
    if not quelle:
        p.error("Bitte einen SharePoint-Link oder lokalen Ordnerpfad angeben.")
    if quelle.lower().startswith("http"):
        try:
            args.root = sharepoint_link_aufloesen(quelle)
        except (FileNotFoundError, ValueError) as ex:
            print(ex, file=sys.stderr)
            return 1
        print(f"SharePoint-Link -> lokaler Ordner: {args.root}")
    else:
        args.root = quelle

    if not os.path.isdir(args.root):
        print(f"Ordner nicht gefunden: {args.root}", file=sys.stderr)
        return 1
    neue = {parse_datum(s).strftime("%d.%m.%Y") for s in args.neue_staende.split(",") if s.strip()}
    if args.ausgabe:
        ausgabe = args.ausgabe
    else:
        out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
        os.makedirs(out_dir, exist_ok=True)
        ausgabe = os.path.join(out_dir, f"AGB_Check_{dt.datetime.now():%Y-%m-%d_%H%M}.xlsx")

    print("Scanne Kundenordner ...", flush=True)
    kunden = ordner_scannen(args.root)
    args.struktur = struktur_info(kunden, args.root)
    print("Kundenordner je Ebene: " + ", ".join(f"Ebene {e}: {n}" for e, n in args.struktur.items()), flush=True)
    if not any(k.agb_dateien or k.auftrag_pfad for k in kunden):
        print("\nWARNUNG: In keinem Ordner wurde eine AGB_Stromtarif- oder Auftrag_Stromtarif-PDF gefunden.\n"
              "Stimmt der Ordner? Ist er vollständig synchronisiert? Die Excel wird trotzdem erstellt.\n", flush=True)
    stichtag = dt.datetime.combine(args.stichtag, dt.time.min)
    zu_pruefen = [k.aktuelle_agb for k in kunden if k.aktuelle_agb and k.aktuelle_agb.geaendert >= stichtag]
    # Bei mehreren AGB-Dateien im Ordner auch die übrigen lesen (für die Spalte "Alle Stände").
    zu_pruefen += [a for k in kunden if len(k.agb_dateien) > 1 and k.aktuelle_agb.geaendert >= stichtag
                   for a in k.agb_dateien if a is not k.aktuelle_agb]
    print(f"{len(kunden)} Ordner gefunden, {len(zu_pruefen)} AGB-PDFs ab {args.stichtag:%d.%m.%Y} werden gelesen.", flush=True)
    pruefen(zu_pruefen, args.threads)

    if args.mit_email:
        # Nur für Kunden, die in der Excel auftauchen: geprüfte und Ordner ohne AGB.
        relevant = [k for k in kunden if k.aktuelle_agb is None or k.aktuelle_agb.geaendert >= stichtag]
        print(f"Lese E-Mail-Adressen aus {sum(1 for k in relevant if k.auftrag_pfad)} Aufträgen ...", flush=True)
        emails_lesen(relevant, args.threads)

    r = excel_schreiben(ausgabe, kunden, args, neue)
    print()
    print(f"Ordner gesamt:            {r['gesamt']}")
    print(f"AGB vor Stichtag:         {r['vor_stichtag']} (nicht geprüft)")
    print(f"Ohne AGB-Datei:           {r['ohne_agb']}")
    print(f"Geprüft:                  {r['geprueft']}")
    print(f"  neu:                    {r[STATUS_NEU]}")
    print(f"  alt:                    {r[STATUS_ALT]}")
    print(f"  nicht lesbar:           {r[STATUS_UNKLAR]}")
    print(f"Quote neu / alt:          {r['quote_neu']:.1%} / {r['quote_alt']:.1%}")
    if args.mit_email:
        print(f"E-Mail gefunden:          {r['mit_email']} von {r['geprueft']} geprüften Kunden")
    print(f"Excel: {os.path.abspath(ausgabe)}")
    if args.oeffnen and hasattr(os, "startfile"):
        os.startfile(os.path.abspath(ausgabe))
    return 0


if __name__ == "__main__":
    sys.exit(main())
