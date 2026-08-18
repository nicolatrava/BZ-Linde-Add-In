#!/usr/bin/env python3
"""Macht aus den ausgefuellten Beispiel-Dokumenten leere Vorlagen.

    python3 tools/clean-vorlagen.py <Ordner-mit-Originalen>

Die Originale, die aus dem Betrieb kommen, enthalten echte Inhalte und echte
Personendaten (Empfaengeradressen, Namen von Mitarbeitenden, Dateipfade aus
dem P-Laufwerk). Dieses Repository ist oeffentlich - die Originale duerfen
darum NICHT eingecheckt werden. Hier entsteht daraus je Vorlage eine leere
Fassung, die nach vorlagen/ geschrieben wird.

Entfernt werden:
  * die Beispiel-Inhalte (Regeln unten, je Vorlage einzeln aufgefuehrt)
  * Autor, Bearbeiter, Titel und Firma aus den Dokumenteigenschaften
  * die benutzerdefinierten Eigenschaften (Doc.*) der Protokoll-Vorlage
  * zwischengespeicherte Feldergebnisse mit P-Laufwerk-Pfaden
  * eingebettete Zeichnungen im Rumpf, wo sie zum Beispiel gehoeren
  * Rueckstaende der Fremdvorlage, aus der die Protokoll-Vorlage stammt
    (SSBL Luzern, Standort Rathausen) - auch in settings.xml und theme1.xml

Zum Schluss laeuft eine Sperrliste ueber jede erzeugte Datei. Taucht dort ein
Name oder Pfad aus den Originalen auf, bricht das Skript ab.
"""

import pathlib
import re
import shutil
import sys
import xml.etree.ElementTree as ET
import zipfile

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "vorlagen"

DOC = "word/document.xml"

# Namen und Pfade, die in keiner erzeugten Datei mehr vorkommen duerfen.
SPERRLISTE = [
    "Sina Joachim", "Joachim", "Buttisholz", "Dorf 27",
    "Isabella", "Wüest", "Barbara", "Wüthrich",
    "SSBL", "Schwerbehinderte", "Rathausen",
    "p:\\", "P:\\",
    "Sterbebegleitung", "Palliativ", "Palliative",
    "Personaladministration", "Eingangsrechnungen", "Zahlungsauftrag",
]

# Wird in JEDEM XML-Teil jeder Vorlage ersetzt. Diese Namen stecken nicht nur
# im Text, sondern auch in settings.xml, theme1.xml und den benutzer-
# definierten Eigenschaften - einzelne Absatzregeln erwischen sie nicht.
GLOBAL_ERSETZEN = [
    ("Stiftung für Schwerbehinderte Luzern SSBL", ""),
    ("Schwerbehinderte", ""),
    ("Rathausen", "Grosswangen"),
    ("SSBL", ""),
    # Farbschema und Design der Fremdvorlage heissen nach ihr bzw. nach einer
    # Person - beides auf einen neutralen Namen ziehen.
    ('name="ssbl"', 'name="Linde"'),
    ('name="Larissa"', 'name="Linde"'),
    ("Barbara Wüthrich", ""),
    ("Wüthrich", ""),
    ("Barbara", ""),
    ("Isabella Wüest", ""),
    ("Wüest Isabella", ""),
    ("Wüest", ""),
]

# Je Quelldatei: Zielname und die Bearbeitungsschritte.
#   run   (teil, absatz, alt, neu)  Textlauf im Absatz ersetzen
#   nach  (teil, absatz)            alles nach dem Doppelpunkt-Lauf leeren
#   leer  (teil, absatz)            ganzen Absatz leeren
#   ab    (teil, absatz)            diesen und alle folgenden Absaetze leeren
#   ohne_bilder (teil)              Zeichnungen im Teil entfernen
VORLAGEN = [
    {
        "quelle": "Brief-Vorlage NEU.docx",
        "ziel": "brief.docx",
        "titel": "Brief",
        "schritte": [
            ("run", DOC, 6, "Frau", "{{ANREDE}}"),
            ("run", DOC, 7, "Sina Joachim", "{{EMPFAENGER}}"),
            ("run", DOC, 7, "Dorf 27", "{{STRASSE}}"),
            ("run", DOC, 8, "6", ""),
            ("run", DOC, 8, "018 Buttisholz", "{{PLZORT}}"),
            ("run", DOC, 13, "18. August 2026", "{{DATUM}}"),
            ("run", DOC, 17, "Betreff:", "Betreff: {{BETREFF}}"),
            ("run", DOC, 41, "Vorname, Name, Bezeichnung/Abteilung", "{{ABSENDER}}"),
        ],
    },
    {
        "quelle": "Protokoll-Vorlage.docx",
        "ziel": "protokoll.docx",
        "titel": "Protokoll",
        # Der Ort im Datumsfeld ist ein DOCPROPERTY der Fremdvorlage und wird
        # von GLOBAL_ERSETZEN auf Grosswangen gezogen; das Datum ist ein
        # DATE-Feld und aktualisiert sich selbst.
        "schritte": [
            ("run", DOC, 0, ": Betreff", ": {{BETREFF}}"),
        ],
    },
    {
        "quelle": "Prozess-Steckbrief - Vorlage Neu.docx",
        "ziel": "prozess-steckbrief.docx",
        "titel": "Prozess-Steckbrief",
        "schritte": [
            ("run", DOC, 11, "B Kernprozesse", "{{PROZESSGRUPPE}}"),
            ("run", DOC, 11, "B1 Pflege und Betreuung", "{{PROZESSBEREICH}}"),
            ("run", DOC, 13, "B1.", ""),
            ("run", DOC, 13, "6 Sterbebegleitung", "{{PROZESSNR}} {{PROZESSNAME}}"),
            ("nach", DOC, 15), ("nach", DOC, 17), ("nach", DOC, 19),
            ("nach", DOC, 21), ("nach", DOC, 23), ("nach", DOC, 24),
            ("nach", DOC, 26), ("nach", DOC, 28),
            ("run", DOC, 30, "PDL", "{{VERANTWORTLICH}}"),
        ],
    },
    {
        "quelle": "Prozess-Vorlage NEU.docx",
        "ziel": "prozess.docx",
        "titel": "Prozess",
        "schritte": [
            ("run", "word/header1.xml", 0,
             "C Unterstützungsprozesse / C3 Sekretariat", "{{PROZESSGRUPPE}}"),
            ("run", "word/header1.xml", 0,
             "C3.3.1 Personaladministration", "{{PROZESSNR}} {{PROZESSNAME}}"),
            ("run", "word/header1.xml", 1, "Isabella Wüest", "{{VERANTWORTUNG}}"),
            ("run", "word/footer1.xml", 1, "Erstelldatum: 02.12.2014", "Erstelldatum: {{DATUM}}"),
            ("run", "word/footer1.xml", 2, "30.11.2023", ""),
            # Der ganze Rumpf ab der zweiten Tabellenzeile ist das Beispiel.
            ("ab", DOC, 4),
            ("ohne_bilder", DOC),
        ],
    },
    {
        "quelle": "QMS Vorlage Hilfsmittel.docx",
        "ziel": "qms-hilfsmittel.docx",
        "titel": "QMS-Hilfsmittel",
        "schritte": [
            ("run", DOC, 0, "Pflege und Betreuung", "{{BEREICH}}"),
            ("run", DOC, 2, "B1.6", "{{NUMMER}}"),
            ("run", DOC, 4, "Richtlinie", "{{DOKUMENTART}}"),
            ("run", DOC, 5, "Palliativ", "{{TITEL}}"),
            ("run", DOC, 5, "care", ""),
            ("leer", DOC, 10), ("leer", DOC, 11),
            ("leer", DOC, 14),
            ("leer", DOC, 18), ("leer", DOC, 19), ("leer", DOC, 20),
            ("leer", DOC, 24), ("leer", DOC, 25), ("leer", DOC, 26),
        ],
    },
]


ROOT_TAG = re.compile(rb"<(w:document|w:hdr|w:ftr|w:settings)\b[^>]*>")


def serialisieren(root, original: bytes) -> bytes:
    """Baum zurueckschreiben und die Namensraeume beider Seiten vereinen.

    ElementTree gibt nur die Namensraeume aus, die im Baum wirklich benutzt
    werden. Eine Word-Datei deklariert an der Wurzel aber rund drei Dutzend
    und zaehlt einen Teil davon in mc:Ignorable auf. Faellt eine Deklaration
    weg, die dort noch steht, ist die Datei formal ungueltig: Word bricht mit
    GeneralException ab, waehrend LibreOffice sie klaglos oeffnet.

    Umgekehrt vergibt ElementTree fuer Namensraeume, die es nicht kennt,
    eigene Praefixe (ns0, ns1 ...) und deklariert sie an SEINER Wurzel. Das
    Wurzel-Tag einfach durch das alte zu ersetzen, reisst die wieder weg -
    dann ist die Datei an anderer Stelle kaputt ("unbound prefix").

    Darum: das Wurzel-Tag von ElementTree behalten und nur die Deklarationen
    ergaenzen, die im Original standen und dort fehlen.
    """
    m_alt = ROOT_TAG.search(original)
    if not m_alt:
        raise SystemExit("Wurzel-Element nicht gefunden")
    alt_tag = m_alt.group(0).decode("utf-8")

    neu = ET.tostring(root, encoding="unicode")
    m_neu = re.match(r"<[^>]+>", neu)
    neu_tag = m_neu.group(0)

    schon_da = set(re.findall(r'xmlns:([A-Za-z0-9]+)=', neu_tag))
    ergaenzen = [d for d in re.findall(r'xmlns:[A-Za-z0-9]+="[^"]*"', alt_tag)
                 if re.match(r'xmlns:([A-Za-z0-9]+)=', d).group(1) not in schon_da]
    if ergaenzen:
        neu_tag = neu_tag[:-1].rstrip() + " " + " ".join(ergaenzen) + ">"

    kopf = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    return (kopf + neu_tag + neu[m_neu.end():]).encode("utf-8")


def register_namespaces(xml: bytes) -> None:
    """Namensraum-Praefixe erhalten, sonst schreibt ElementTree ns0, ns1 ..."""
    for m in re.finditer(rb'xmlns:([A-Za-z0-9]+)="([^"]+)"', xml):
        ET.register_namespace(m.group(1).decode(), m.group(2).decode())


def paragraphs(root):
    return list(root.iter(W + "p"))


def texts(p):
    return [t for t in p.iter(W + "t")]


def set_text(t, value):
    t.text = value
    if value != value.strip():
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def op_run(root, idx, alt, neu, protokoll):
    ps = paragraphs(root)
    if idx >= len(ps):
        raise SystemExit("Absatz %d gibt es nicht" % idx)
    for t in texts(ps[idx]):
        if t.text == alt:
            set_text(t, neu)
            protokoll.append("  Absatz %-3d %r -> %r" % (idx, alt, neu))
            return
    raise SystemExit("Absatz %d: Textlauf %r nicht gefunden" % (idx, alt))


def op_nach(root, idx, protokoll):
    """Alles nach dem Lauf mit dem Doppelpunkt leeren - die Beschriftung bleibt."""
    ps = paragraphs(root)
    ts = texts(ps[idx])
    treffer = next((i for i, t in enumerate(ts) if t.text and ":" in t.text), None)
    if treffer is None:
        raise SystemExit("Absatz %d: keine Beschriftung mit Doppelpunkt" % idx)
    weg = 0
    for t in ts[treffer + 1:]:
        if t.text:
            weg += 1
        t.text = ""
    protokoll.append("  Absatz %-3d Beschriftung %r behalten, %d Wert-Laeufe geleert"
                     % (idx, ts[treffer].text, weg))


def op_leer(root, idx, protokoll):
    ps = paragraphs(root)
    alt = "".join(t.text or "" for t in texts(ps[idx]))
    for t in texts(ps[idx]):
        t.text = ""
    protokoll.append("  Absatz %-3d geleert (war %r)" % (idx, alt[:60]))


def op_ab(root, start, protokoll):
    ps = paragraphs(root)
    n = 0
    for i in range(start, len(ps)):
        for t in texts(ps[i]):
            if t.text:
                n += 1
            t.text = ""
    protokoll.append("  Absatz %d bis %d geleert (%d Laeufe)" % (start, len(ps) - 1, n))


def op_ohne_bilder(root, protokoll):
    n = 0
    for parent in root.iter():
        for kind in list(parent):
            if kind.tag in (W + "drawing", W + "pict", W + "object"):
                parent.remove(kind)
                n += 1
    protokoll.append("  %d Zeichnung(en) aus dem Rumpf entfernt" % n)


def scrub_feldergebnisse(root, protokoll):
    """Zwischengespeicherte Feldergebnisse mit P-Laufwerk-Pfaden leeren."""
    n = 0
    for t in root.iter(W + "t"):
        if t.text and re.search(r"[pP]:\\", t.text):
            t.text = ""
            n += 1
    if n:
        protokoll.append("  %d zwischengespeicherte(s) Feldergebnis(se) geleert" % n)


def scrub_docprops(zin, name, data, protokoll):
    s = data.decode("utf-8")
    if name == "docProps/core.xml":
        for tag in ("dc:creator", "cp:lastModifiedBy", "dc:title", "dc:subject",
                    "cp:keywords", "dc:description", "cp:category"):
            s2 = re.sub(r"<%s>.*?</%s>" % (tag, tag), "<%s></%s>" % (tag, tag), s, flags=re.S)
            if s2 != s:
                protokoll.append("  %s: %s geleert" % (name, tag))
                s = s2
    elif name == "docProps/app.xml":
        for tag in ("Company", "Manager"):
            s2 = re.sub(r"<%s>.*?</%s>" % (tag, tag), "<%s></%s>" % (tag, tag), s, flags=re.S)
            if s2 != s:
                protokoll.append("  %s: %s geleert" % (name, tag))
                s = s2
    elif name == "docProps/custom.xml":
        s2 = re.sub(r"<vt:lpwstr>[^<]*</vt:lpwstr>", "<vt:lpwstr></vt:lpwstr>", s)
        if s2 != s:
            protokoll.append("  %s: Werte der benutzerdefinierten Eigenschaften geleert" % name)
            s = s2
    return s.encode("utf-8")


DOCVARS = re.compile(r"<w:docVars>.*?</w:docVars>", re.S)


def strip_docvars(data: bytes, protokoll):
    """settings.xml der Protokoll-Vorlage traegt ein komplettes Profil der
    Fremdorganisation mit (Adresse, Telefon, Logopfade, Mitarbeitendensatz).
    Dokumentvariablen braucht keine der Vorlagen - sie fliegen ganz raus."""
    s = data.decode("utf-8", "replace")
    s2 = DOCVARS.sub("", s)
    if s2 != s:
        protokoll.append("  word/settings.xml: Dokumentvariablen entfernt")
    return s2.encode("utf-8")


def global_ersetzen(data: bytes):
    s = data.decode("utf-8", "replace")
    n = 0
    for alt, neu in GLOBAL_ERSETZEN:
        n += s.count(alt)
        s = s.replace(alt, neu)
    return s.encode("utf-8"), n


def clean(src: pathlib.Path, spec: dict) -> pathlib.Path:
    protokoll = []
    zin = zipfile.ZipFile(src)
    # Welche XML-Teile werden angefasst?
    parts = {}
    for schritt in spec["schritte"]:
        parts.setdefault(schritt[1], None)
    for name in zin.namelist():
        if re.match(r"word/(document|header\d+|footer\d+)\.xml$", name):
            parts.setdefault(name, None)

    roh = {}
    for part in list(parts):
        raw = zin.read(part)
        roh[part] = raw
        register_namespaces(raw)
        parts[part] = ET.fromstring(raw)

    for schritt in spec["schritte"]:
        art, part = schritt[0], schritt[1]
        root = parts[part]
        if art == "run":
            op_run(root, schritt[2], schritt[3], schritt[4], protokoll)
        elif art == "nach":
            op_nach(root, schritt[2], protokoll)
        elif art == "leer":
            op_leer(root, schritt[2], protokoll)
        elif art == "ab":
            op_ab(root, schritt[2], protokoll)
        elif art == "ohne_bilder":
            op_ohne_bilder(root, protokoll)
        else:
            raise SystemExit("unbekannter Schritt: " + art)

    for root in parts.values():
        scrub_feldergebnisse(root, protokoll)

    ziel = OUT / spec["ziel"]
    with zipfile.ZipFile(ziel, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename in parts:
                data = serialisieren(parts[item.filename], roh[item.filename])
            else:
                data = zin.read(item.filename)
            if item.filename == "word/settings.xml":
                data = strip_docvars(data, protokoll)
            if item.filename.startswith("docProps/"):
                data = scrub_docprops(zin, item.filename, data, protokoll)
            if item.filename.endswith((".xml", ".rels")):
                data, n = global_ersetzen(data)
                if n:
                    protokoll.append("  %s: %d Rueckstand/Rueckstaende ersetzt"
                                     % (item.filename, n))
            zout.writestr(item, data)

    print("%s  ->  vorlagen/%s" % (spec["quelle"], spec["ziel"]))
    for zeile in protokoll:
        print(zeile)
    return ziel


def pruefe(ziel: pathlib.Path) -> list:
    """Sperrliste ueber alle Textteile der erzeugten Datei."""
    treffer = []
    z = zipfile.ZipFile(ziel)
    for name in z.namelist():
        if not (name.endswith(".xml") or name.endswith(".rels")):
            continue
        text = z.read(name).decode("utf-8", "replace")
        for wort in SPERRLISTE:
            if wort.lower() in text.lower():
                treffer.append("%s: %r in %s" % (ziel.name, wort, name))
    return treffer


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    quelle = pathlib.Path(sys.argv[1])
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir()

    alle_treffer = []
    for spec in VORLAGEN:
        src = quelle / spec["quelle"]
        if not src.exists():
            print("fehlt: " + str(src), file=sys.stderr)
            return 1
        ziel = clean(src, spec)
        alle_treffer += pruefe(ziel)
        print()

    if alle_treffer:
        print("ABBRUCH - Sperrliste hat angeschlagen:", file=sys.stderr)
        for t in alle_treffer:
            print("  " + t, file=sys.stderr)
        return 1
    print("Sperrliste: sauber, keine Personendaten und keine P-Pfade mehr.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
