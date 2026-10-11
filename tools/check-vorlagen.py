#!/usr/bin/env python3
"""Prueft, ob Vorlagen und Add-In zusammenpassen.

    python3 tools/check-vorlagen.py

Verglichen wird beides gegeneinander:
  * welche Platzhalter {{...}} in den .docx unter vorlagen/ wirklich stehen
  * welche Felder taskpane.html je Vorlage anbietet

Weicht das voneinander ab, bleibt entweder ein Platzhalter im fertigen
Dokument stehen oder ein Eingabefeld laeuft ins Leere. Beides faellt sonst
erst auf, wenn jemand in Word damit arbeitet.

Geprueft wird ausserdem, ob jedes in mc:Ignorable aufgezaehlte Namensraum-
Praefix an der Wurzel auch deklariert ist. Fehlt eines, ist die Datei formal
ungueltig: Word bricht beim Einsetzen mit GeneralException ab, waehrend
LibreOffice sie klaglos oeffnet - der Fehler faellt sonst erst in Word auf.

Und die Sperrliste aus clean-vorlagen.py laeuft nochmals ueber die
eingecheckten Dateien - das Repository ist oeffentlich.
"""

import pathlib
import re
import sys
import zipfile

ROOT = pathlib.Path(__file__).resolve().parent.parent
VORLAGEN_DIR = ROOT / "vorlagen"
TASKPANE = ROOT / "taskpane.html"

sys.path.insert(0, str(ROOT / "tools"))
try:
    from importlib import import_module
    SPERRLISTE = import_module("clean-vorlagen".replace("-", "_")).SPERRLISTE
except Exception:  # Modulname mit Bindestrich laesst sich nicht importieren
    SPERRLISTE = [
        "Sina Joachim", "Buttisholz", "Dorf 27", "Isabella", "Wüest",
        "Barbara", "Wüthrich", "SSBL", "Schwerbehinderte", "Rathausen",
        "p:\\", "P:\\", "Sterbebegleitung", "Palliativ", "Palliative",
        "Personaladministration", "Eingangsrechnungen", "Zahlungsauftrag",
    ]

PLATZHALTER = re.compile(r"\{\{([A-ZÄÖÜ_]+)\}\}")
# Setzt das Add-In selbst, dafuer gibt es kein Eingabefeld: das Datum und die
# Grussanrede des Briefs (aus Anrede und Name).
AUTOMATISCH = {"DATUM", "BRIEFANREDE"}


def platzhalter_in_docx(pfad: pathlib.Path) -> set:
    """Platzhalter finden - auch wenn Word den Text auf mehrere Laeufe verteilt."""
    gefunden = set()
    z = zipfile.ZipFile(pfad)
    for name in z.namelist():
        if not re.match(r"word/(document|header\d+|footer\d+)\.xml$", name):
            continue
        xml = z.read(name).decode("utf-8", "replace")
        nur_text = re.sub(r"<[^>]+>", "", xml)
        gefunden |= set(PLATZHALTER.findall(nur_text))
    return gefunden


WURZEL = re.compile(r"<w:(?:document|hdr|ftr)\b[^>]*>")


def namensraeume_pruefen(pfad: pathlib.Path) -> list:
    """Jedes Praefix aus mc:Ignorable muss an der Wurzel deklariert sein."""
    probleme = []
    z = zipfile.ZipFile(pfad)
    for name in z.namelist():
        if not re.match(r"word/(document|header\d+|footer\d+)\.xml$", name):
            continue
        xml = z.read(name).decode("utf-8", "replace")
        m = WURZEL.search(xml)
        if not m:
            probleme.append("%s: kein Wurzel-Element" % name)
            continue
        wurzel = m.group(0)
        deklariert = set(re.findall(r'xmlns:([A-Za-z0-9]+)=', wurzel))
        ign = re.search(r'mc:Ignorable="([^"]*)"', wurzel)
        for praefix in (ign.group(1).split() if ign else []):
            if praefix not in deklariert:
                probleme.append("%s: mc:Ignorable nennt %r, aber xmlns:%s fehlt"
                                % (name, praefix, praefix))
    return probleme


def brief_layout_pruefen(pfad: pathlib.Path) -> list:
    """Vorgaben der Linde fuer den Brief: Absatzabstand 0, Zeilenabstand
    einfach - in Rumpf, Kopf, Fuss und den Formatvorlagen - und das Logo
    mittig in der Kopfzeile."""
    probleme = []
    z = zipfile.ZipFile(pfad)
    for name in ("word/document.xml", "word/styles.xml", "word/header1.xml", "word/footer1.xml"):
        xml = z.read(name).decode("utf-8", "replace")
        for sp in re.findall(r"<w:spacing[^>]*/>", xml):
            if (re.search(r'w:(?:before|after)="(?!0")\d+"', sp)
                    or re.search(r'w:line="(?!240")\d+"', sp)
                    or ("lineRule" in sp and 'lineRule="auto"' not in sp)):
                probleme.append("%s: Absatzabstand nicht 0 / Zeilenabstand nicht einfach: %s"
                                % (name, sp))
    kopf = z.read("word/header1.xml").decode("utf-8", "replace")
    if "<w:drawing" not in kopf or '<w:jc w:val="center"' not in kopf:
        probleme.append("header1.xml: Logo fehlt oder ist nicht zentriert")
    return probleme


def felder_aus_taskpane() -> dict:
    """Aus dem VORLAGEN-Array je Datei die angebotenen Felder lesen."""
    html = TASKPANE.read_text(encoding="utf-8")
    block = re.search(r"var VORLAGEN = \[(.*?)\n\];", html, re.S)
    if not block:
        print("VORLAGEN-Array in taskpane.html nicht gefunden", file=sys.stderr)
        sys.exit(1)
    out = {}
    for chunk in block.group(1).split("{ id:")[1:]:
        m = re.search(r"datei:'(vorlagen/[^']+)'", chunk)
        if not m:
            continue          # der Zweig ohne eigene Datei
        out[m.group(1)] = set(re.findall(r"ph:'([A-ZÄÖÜ_]+)'", chunk))
    return out


def main() -> int:
    felder = felder_aus_taskpane()
    dateien = sorted(VORLAGEN_DIR.glob("*.docx"))
    if not dateien:
        print("vorlagen/ ist leer - zuerst tools/clean-vorlagen.py laufen lassen",
              file=sys.stderr)
        return 1

    fehler = []
    for pfad in dateien:
        rel = "vorlagen/" + pfad.name
        im_docx = platzhalter_in_docx(pfad)
        angeboten = felder.get(rel)
        if angeboten is None:
            fehler.append("%s: liegt in vorlagen/, wird aber im Add-In nicht angeboten" % rel)
            continue
        fehlt = im_docx - angeboten - AUTOMATISCH
        zuviel = angeboten - im_docx
        for ph in sorted(fehlt):
            fehler.append("%s: {{%s}} steht im Dokument, hat aber kein Eingabefeld" % (rel, ph))
        for ph in sorted(zuviel):
            fehler.append("%s: Eingabefeld {{%s}} kommt im Dokument nicht vor" % (rel, ph))
        print("%-34s Platzhalter: %s" % (
            rel, ", ".join(sorted(im_docx)) or "keine"))

        text = " ".join(
            zipfile.ZipFile(pfad).read(n).decode("utf-8", "replace")
            for n in zipfile.ZipFile(pfad).namelist()
            if n.endswith((".xml", ".rels"))
        ).lower()
        for wort in SPERRLISTE:
            if wort.lower() in text:
                fehler.append("%s: Sperrliste - %r steckt noch drin" % (rel, wort))

        for problem in namensraeume_pruefen(pfad):
            fehler.append("%s: %s" % (rel, problem))

        if pfad.name == "brief.docx":
            for problem in brief_layout_pruefen(pfad):
                fehler.append("%s: %s" % (rel, problem))

    for rel in felder:
        if not (ROOT / rel).exists():
            fehler.append("%s: im Add-In angeboten, Datei fehlt aber" % rel)

    if fehler:
        print("\nFehler:", file=sys.stderr)
        for f in fehler:
            print("  " + f, file=sys.stderr)
        return 1
    print("\nVorlagen und Add-In passen zusammen, Namensraeume vollstaendig, "
          "Sperrliste sauber.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
