#!/usr/bin/env python3
"""Setzt das Feedback der Linde vom September 2026 in den Vorlagen um.

    python3 tools/feedback-anpassungen.py

Laeuft auf den bereinigten Dateien unter vorlagen/ (nach clean-vorlagen.py
und modernisiere-qms.py) und ist wiederholbar: eine bereits angepasste
Vorlage wird erkannt und uebersprungen. Die QMS-Hilfsmittel-Vorlage bleibt
unveraendert - sie war laut Rueckmeldung "sehr gut".

Brief
  * Logo oben mittig (eigene Kopfzeile), Adressblock darunter - Aufbau und
    Leerzeilen wie in Isabellas freigegebenem Muster (Brief Test3)
  * ohne Absatzabstaende, alles auf einer linken Fluchtlinie, Aptos 11,
    Betreff 12 fett, alles schwarz - ueber eigene Formatvorlagen "Linde
    Brief/Betreff/Kopfzeile/Fusszeile", damit Word beim Einsetzen nicht die
    Standard-Vorlagen des Zieldokuments darueberlegt
  * Grussformel mit {{BRIEFANREDE}} (setzt das Add-In aus Anrede und Name),
    Unterschrift mit {{ABSENDER}} und {{EMAIL}}, weniger Leerraum davor
  * Fusszeile: Betagtenzentrum Linde | Dorfstrasse 6e | 6022 Grosswangen |
    041 984 29 29 | www.bz-linde.ch

Protokoll
  * Datums-/Entwurfsfelder der Fremdvorlage aus den Kopfzeilen entfernt
  * DOCPROPERTY-Felder der Fremdvorlage durch ihren Text ersetzt - sonst
    zeigt Word nach F9 "Fehler! Unbekannter Name fuer Dokumenteigenschaft"
  * Abstand Titel zu Datum/Zeit verkleinert, "Alle Teilnehmende" korrigiert
  * Traktanden-Zellen auf einen Absatz, der Cursor sitzt oben in der Zelle
  * {{PROTOKOLLFUEHRUNG}} in der Kopftabelle und bei der Unterschrift,
    Abstaende am Schluss verkleinert, Datum als {{DATUM}}

Prozess
  * Logo links in der Kopfzeile, Angaben rechts auf drei Zeilen

Prozess-Steckbrief
  * Titel im grauen Balken in Schriftgroesse 11

Alle vier Vorlagen nutzen dasselbe Logo assets/logo-linde.jpg (farbtreu
ueber das ICC-Profil aus dem 1-MB-CMYK-Original gewandelt, 57 KB).
"""

import pathlib
import re
import sys
import zipfile

ROOT = pathlib.Path(__file__).resolve().parent.parent
VORLAGEN = ROOT / "vorlagen"
LOGO = ROOT / "assets" / "logo-linde.jpg"
LOGO_RATIO = 1600 / 701                    # Seitenverhaeltnis des Logos

DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
REL_HDR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
REL_IMG = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
CT_HDR = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"

APTOS = '<w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:cs="Aptos"/>'
FUSSZEILE_BRIEF = ("Betagtenzentrum Linde | Dorfstrasse 6e | 6022 Grosswangen | "
                   "041 984 29 29 | www.bz-linde.ch")
LOGO_DESCR = "Logo Betagtenzentrum Linde"


# ─── Paket lesen und schreiben ──────────────────────────────────────────────

def lesen(pfad):
    z = zipfile.ZipFile(pfad)
    return ({i.filename: z.read(i.filename) for i in z.infolist()},
            {i.filename: i for i in z.infolist()})


def schreiben(pfad, teile, infos):
    with zipfile.ZipFile(pfad, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in teile.items():
            z.writestr(infos.get(name, name), data)


def text(teile, name):
    return teile[name].decode("utf-8")


def setze(teile, name, xml):
    teile[name] = xml.encode("utf-8")


def wurzel_tag(xml):
    """Das Wurzelelement (<w:hdr ...>, <w:ftr ...>) mit allen Namensraeumen."""
    return re.search(r"<w:(?:document|hdr|ftr)\b[^>]*>", xml).group(0)


def hdr_wurzel():
    """Vollstaendiger Namensraum-Satz einer Word-Kopfzeile mit Zeichnung."""
    teile, _ = lesen(VORLAGEN / "protokoll.docx")
    tag = wurzel_tag(text(teile, "word/header2.xml"))
    if "xmlns:pic=" not in tag or "xmlns:a=" not in tag:
        raise SystemExit("protokoll header2.xml ohne Zeichnungs-Namensraeume")
    return tag


def rels_ergaenzen(teile, name, rid, typ, ziel):
    if name in teile:
        rels = text(teile, name)
        if 'Id="%s"' % rid in rels:
            return
    else:
        rels = DECL + '<Relationships xmlns="%s"></Relationships>' % RELS_NS
    rels = rels.replace("</Relationships>",
                        '<Relationship Id="%s" Type="%s" Target="%s"/></Relationships>'
                        % (rid, typ, ziel))
    setze(teile, name, rels)


def content_types(teile, override=None, jpeg=True):
    ct = text(teile, "[Content_Types].xml")
    if jpeg and 'Extension="jpeg"' not in ct:
        ct = ct.replace("</Types>",
                        '<Default Extension="jpeg" ContentType="image/jpeg"/></Types>')
    if override and 'PartName="%s"' % override[0] not in ct:
        ct = ct.replace("</Types>",
                        '<Override PartName="%s" ContentType="%s"/></Types>' % override)
    setze(teile, "[Content_Types].xml", ct)


# ─── OOXML-Bausteine ────────────────────────────────────────────────────────

def rpr(fett=False, sz=22, farbe=None):
    return ('<w:rPr>' + APTOS + ('<w:b/><w:bCs/>' if fett else '') +
            ('<w:color w:val="%s"/>' % farbe if farbe else '') +
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (sz, sz))


def absatz(inhalt="", fett=False, sz=22, ind=0, stil=None, jc=None, farbe=None):
    """Ein Absatz ohne Abstaende; inhalt ist Text oder fertige Laeufe."""
    ppr = ('<w:pPr>' + ('<w:pStyle w:val="%s"/>' % stil if stil else '') +
           '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>' +
           ('<w:ind w:left="%d"/>' % ind if ind else '') +
           ('<w:jc w:val="%s"/>' % jc if jc else '') +
           rpr(fett, sz, farbe) + '</w:pPr>')
    if inhalt.startswith("<w:r"):
        laeufe = inhalt
    elif inhalt:
        laeufe = '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr(fett, sz, farbe), inhalt)
    else:
        laeufe = ""
    return "<w:p>" + ppr + laeufe + "</w:p>"


def pic(rid, cx, cy, doc_id):
    return ('<wp:docPr id="%d" name="Logo" descr="%s"/>'
            '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/>'
            '</wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.'
            'openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr>'
            '<pic:cNvPr id="0" name="Logo"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch>'
            '</pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/>'
            '<a:ext cx="%d" cy="%d"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/>'
            '</a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
            % (doc_id, LOGO_DESCR, rid, cx, cy))


def logo_inline(rid, cy):
    """Logo im Textfluss - so steht es im Protokoll."""
    cx = int(cy * LOGO_RATIO)
    return ('<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>'
            '<wp:inline distT="0" distB="0" distL="0" distR="0">'
            '<wp:extent cx="%d" cy="%d"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            % (cx, cy) + pic(rid, cx, cy, 1) + '</wp:inline></w:drawing></w:r>')


def logo_schwebend(rid, cy):
    """Logo am linken Rand, der Text des Absatzes laeuft daneben rechts."""
    cx = int(cy * LOGO_RATIO)
    return ('<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>'
            '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
            'relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" '
            'allowOverlap="1"><wp:simplePos x="0" y="0"/>'
            '<wp:positionH relativeFrom="margin"><wp:align>left</wp:align></wp:positionH>'
            '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset>'
            '</wp:positionV><wp:extent cx="%d" cy="%d"/>'
            '<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
            % (cx, cy) + pic(rid, cx, cy, 1) + '</wp:anchor></w:drawing></w:r>')


LOGO_CY_GROSS = 1213215      # 3.37 cm - wie bisher im Protokoll
LOGO_CY_BRIEF = 1155960      # 3.2 cm, mittig - wie in Isabellas Muster (Brief Test3)
LOGO_CY_PROZESS = 720000     # 2.0 cm
SCHWARZ = "000000"

# Eigene Formatvorlagen fuer den Brief. Word ersetzt beim Einsetzen einer
# Datei die Formatvorlagen gleichen Namens (Standard, Ueberschrift 1, Kopf-
# und Fusszeile) durch die des Zieldokuments - so kamen dessen 8 pt Absatz-
# abstand, der 1.16-fache Zeilenabstand und das Blau von Ueberschrift 1 in
# den Brief. Unbekannte Formatvorlagen nimmt Word dagegen mit.
LINDE_STYLES = (
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="LindeBrief">'
    '<w:name w:val="Linde Brief"/><w:basedOn w:val="Standard"/><w:qFormat/>'
    '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    '<w:rPr>' + APTOS + '<w:color w:val="000000"/><w:sz w:val="22"/><w:szCs w:val="22"/>'
    '</w:rPr></w:style>'
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="LindeBetreff">'
    '<w:name w:val="Linde Betreff"/><w:basedOn w:val="LindeBrief"/>'
    '<w:next w:val="LindeBrief"/><w:qFormat/>'
    '<w:rPr><w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>'
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="LindeKopf">'
    '<w:name w:val="Linde Kopfzeile"/><w:basedOn w:val="LindeBrief"/>'
    '<w:pPr><w:jc w:val="center"/></w:pPr></w:style>'
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="LindeFuss">'
    '<w:name w:val="Linde Fusszeile"/><w:basedOn w:val="LindeBrief"/>'
    '<w:pPr><w:jc w:val="center"/></w:pPr>'
    '<w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:style>'
)


# ─── Felder der Fremdvorlage ────────────────────────────────────────────────

KIND = re.compile(r'<w:r\b[^>]*/>|<w:r\b[^>]*(?<!/)>.*?</w:r>|<w:fldSimple\b[^>]*>.*?</w:fldSimple>'
                  r'|<w:hyperlink\b[^>]*>.*?</w:hyperlink>|<w:[A-Za-z]+\b[^>]*/>', re.S)


def _fldtyp(lauf):
    m = re.search(r'<w:fldChar w:fldCharType="(\w+)"', lauf)
    return m.group(1) if m else None


def felder_glaetten(xml):
    """DOCPROPERTY-Felder durch ihr zwischengespeichertes Ergebnis ersetzen.

    Verschachtelte oder andere Felder (IF, PAGE, FILENAME) bleiben stehen.
    Ein Feld ohne Ergebnis verschwindet ganz.
    """
    def absatz_bearbeiten(m):
        p = m.group(0)
        kopf_ende = p.find("</w:pPr>")
        kopf_ende = kopf_ende + len("</w:pPr>") if kopf_ende >= 0 else p.find(">") + 1
        kopf, rest = p[:kopf_ende], p[kopf_ende:-len("</w:p>")]
        kinder = KIND.findall(rest)
        if "".join(kinder) != rest or "fldCharType" not in rest:
            return p                        # nichts zu tun oder nicht sauber zerlegbar
        out, i = [], 0
        while i < len(kinder):
            if _fldtyp(kinder[i]) != "begin":
                out.append(kinder[i]); i += 1
                continue
            tiefe, j, sep, verschachtelt = 0, i, None, False
            while j < len(kinder):
                t = _fldtyp(kinder[j])
                if t == "begin":
                    tiefe += 1
                    verschachtelt = verschachtelt or tiefe > 1
                elif t == "separate" and tiefe == 1:
                    sep = j
                elif t == "end":
                    tiefe -= 1
                    if tiefe == 0:
                        break
                j += 1
            if tiefe != 0:
                return p                    # Feld ohne Ende - lieber nichts anfassen
            instr = "".join(re.findall(r"<w:instrText[^>]*>([^<]*)</w:instrText>",
                                       "".join(kinder[i:(sep if sep is not None else j)])))
            if not verschachtelt and re.match(r"\s*DOCPROPERTY\b", instr):
                if sep is not None:
                    out.extend(kinder[sep + 1:j])
            else:
                out.extend(kinder[i:j + 1])
            i = j + 1
        return kopf + "".join(out) + "</w:p>"

    return re.sub(r"<w:p\b[^>]*(?<!/)>.*?</w:p>", absatz_bearbeiten, xml, flags=re.S)


def bloecke(body):
    """Oberste Ebene des Rumpfs: Absaetze, Tabellen, sectPr - in Reihenfolge."""
    if body.count("<w:tbl>") != body.count("</w:tbl>"):
        raise SystemExit("Tabellen nicht paarig")
    out = []
    for m in re.finditer(r"<w:tbl>.*?</w:tbl>|<w:p\b[^>]*/>|<w:p\b[^>]*(?<!/)>.*?</w:p>"
                         r"|<w:sectPr\b.*?</w:sectPr>|<w:bookmark(?:Start|End)\b[^>]*/>|\s+",
                         body, re.S):
        out.append(m.group(0))
    if "".join(out) != body:
        raise SystemExit("Rumpf nicht lueckenlos zerlegbar")
    for b in out:
        if b.startswith("<w:tbl>") and b.count("<w:tbl>") > 1:
            raise SystemExit("verschachtelte Tabelle - nicht vorgesehen")
    return out


def nur_text(xml):
    return "".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", xml))



def logo_descr_bereinigen(xml):
    """Laufwerkspfade aus den Bildbeschreibungen der Originale entfernen."""
    return re.sub(r'descr="[^"]*\\[^"]*"', 'descr="%s"' % LOGO_DESCR, xml)


# ─── Brief ──────────────────────────────────────────────────────────────────

def brief(hdr_tag):
    """Rumpf, Kopf, Fuss und Formatvorlagen werden jedes Mal neu aufgebaut;
    geschrieben wird nur, wenn sich dabei etwas aendert."""
    pfad = VORLAGEN / "brief.docx"
    teile, infos = lesen(pfad)
    vorher = dict(teile)

    doc = text(teile, "word/document.xml")
    kopf = doc[:doc.find("<w:body>")]
    sect = re.search(r"<w:sectPr\b.*?</w:sectPr>", doc, re.S).group(0)
    sect = re.sub(r"<w:sectPr\b[^>]*>", "<w:sectPr>", sect)
    if "rIdHdrLogo" not in sect:
        sect = re.sub(r"<w:footerReference\b",
                      '<w:headerReference w:type="default" r:id="rIdHdrLogo"/><w:footerReference',
                      sect, count=1)
    sect = re.sub(r"<w:pgMar[^>]*/>",
                  '<w:pgMar w:top="2943" w:right="1418" w:bottom="1134" w:left="1418" '
                  'w:header="567" w:footer="567" w:gutter="0"/>', sect)

    ADR = 5040                                  # Adressfeld rechts, wie bisher

    def z(inhalt="", **kw):
        return absatz(inhalt, stil="LindeBrief", farbe=SCHWARZ, **kw)

    # Aufbau und Leerzeilen genau wie in Isabellas freigegebenem Muster.
    rumpf = "".join([
        z("", ind=ADR),
        z("{{ANREDE}}", ind=ADR),
        z("{{EMPFAENGER}}", ind=ADR),
        z("{{STRASSE}}", ind=ADR),
        z("{{PLZORT}}", ind=ADR),
        z(), z(), z(),
        z("Grosswangen, {{DATUM}}", ind=ADR),
        z(), z(),
        absatz("Betreff: {{BETREFF}}", fett=True, sz=24, stil="LindeBetreff", farbe=SCHWARZ),
        z(),
        z("{{BRIEFANREDE}}"),
        z(),
        z("Text"),
        z(), z(),
        z("Freundliche Grüsse"),
        z("Betagtenzentrum Linde", fett=True),
        z(), z(),
        z("{{ABSENDER}}"),
        z("{{EMAIL}}"),
    ])
    setze(teile, "word/document.xml",
          kopf + "<w:body>" + rumpf + sect + "</w:body></w:document>")

    # Formatvorlagen: die Linde-Vorlagen frisch einsetzen, Standard auf
    # Aptos 11, Ueberschrift 1 ohne den alten Einzug.
    st = text(teile, "word/styles.xml")
    st = re.sub(r'<w:style w:type="paragraph" w:customStyle="1" w:styleId="Linde\w+">.*?</w:style>',
                "", st, flags=re.S)
    st = st.replace("</w:styles>", LINDE_STYLES + "</w:styles>")
    st = re.sub(
        r'(<w:style w:type="paragraph" w:default="1" w:styleId="Standard">.*?)<w:rPr>.*?</w:rPr>',
        r'\1<w:rPr>' + APTOS + '<w:sz w:val="22"/><w:szCs w:val="22"/>'
        '<w:lang w:eastAsia="de-DE"/></w:rPr>', st, count=1, flags=re.S)
    st = re.sub(r'(<w:style w:type="paragraph" w:styleId="berschrift1">.*?)<w:ind w:left="540"/>',
                r"\1", st, count=1, flags=re.S)
    setze(teile, "word/styles.xml", st)

    # Kopfzeile mit mittigem Logo, Fusszeile mit Adresszeile.
    setze(teile, "word/header1.xml",
          DECL + hdr_tag + absatz(logo_inline("rIdLogo", LOGO_CY_BRIEF), stil="LindeKopf", jc="center")
          + "</w:hdr>")
    rels_ergaenzen(teile, "word/_rels/header1.xml.rels", "rIdLogo", REL_IMG, "media/image1.jpeg")
    teile["word/media/image1.jpeg"] = LOGO.read_bytes()
    rels_ergaenzen(teile, "word/_rels/document.xml.rels", "rIdHdrLogo", REL_HDR, "header1.xml")
    content_types(teile, ("/word/header1.xml", CT_HDR))

    ftr = text(teile, "word/footer1.xml")
    setze(teile, "word/footer1.xml",
          DECL + wurzel_tag(ftr) +
          absatz(FUSSZEILE_BRIEF, sz=16, stil="LindeFuss", jc="center", farbe=SCHWARZ) + "</w:ftr>")

    if all(teile.get(k) == vorher.get(k) for k in set(teile) | set(vorher)):
        print("brief.docx: bereits angepasst")
        return
    schreiben(pfad, teile, infos)
    print("brief.docx: Logo mittig, Linde-Formatvorlagen, Layout wie Muster, Fusszeile")


# ─── Protokoll ──────────────────────────────────────────────────────────────

def protokoll(hdr_tag):
    pfad = VORLAGEN / "protokoll.docx"
    teile, infos = lesen(pfad)
    doc = text(teile, "word/document.xml")
    if "{{PROTOKOLLFUEHRUNG}}" in doc:
        print("protokoll.docx: bereits angepasst")
        return

    # Kopfzeilen: nur noch das Logo (Seite 1), sonst leer.
    setze(teile, "word/header2.xml",
          DECL + hdr_tag + absatz(logo_inline("rId1", LOGO_CY_GROSS), stil="OutputprofileTitle")
          + "</w:hdr>")
    h4 = text(teile, "word/header4.xml")
    setze(teile, "word/header4.xml",
          DECL + wurzel_tag(h4) + '<w:p><w:pPr><w:pStyle w:val="Kopfzeile"/></w:pPr></w:p></w:hdr>')
    teile["word/media/image1.jpeg"] = LOGO.read_bytes()

    kopf = doc[:doc.find("<w:body>") + len("<w:body>")]
    body = doc[len(kopf):doc.rfind("</w:body>")]
    bl = bloecke(body)

    def idx(pred, ab=0):
        for k in range(ab, len(bl)):
            if pred(bl[k]):
                return k
        raise SystemExit("protokoll: Block nicht gefunden")

    # Abschnittswechsel-Absatz auf Minimalhoehe, leere Tabelle danach weg.
    k = idx(lambda b: "<w:sectPr" in b and b.startswith("<w:p"))
    bl[k] = re.sub(r"<w:pPr>.*?<w:sectPr", '<w:pPr><w:spacing w:before="0" w:after="0" '
                   'w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/>'
                   '</w:rPr><w:sectPr', bl[k], count=1, flags=re.S)
    if bl[k + 1].startswith("<w:tbl>") and not nur_text(bl[k + 1]).strip():
        del bl[k + 1]

    # Kopftabelle: Tippfehler, Protokollfuehrung.
    k = idx(lambda b: "Verteiler" in b)
    t = bl[k]
    t, n = re.subn(r'<w:proofErr w:type="spellStart"\s*/>(<w:r\b[^>]*>(?:(?!</w:r>).)*<w:t>)Teilneh'
                   r'(</w:t></w:r>)<w:r\b[^>]*>(?:(?!</w:r>).)*<w:t>nde</w:t></w:r>'
                   r'<w:proofErr w:type="spellEnd"\s*/>', r"\1Teilnehmende\2", t, flags=re.S)
    if n != 1:
        raise SystemExit("protokoll: 'Teilnehnde' nicht gefunden")
    t = felder_glaetten(t)
    zeilen = re.findall(r"<w:tr\b.*?</w:tr>", t, re.S)
    z = [r for r in zeilen if nur_text(r).strip() == "Protokoll"]
    if len(z) != 1:
        raise SystemExit("protokoll: Zeile 'Protokoll' nicht gefunden")
    zelle = z[0].rfind("</w:pPr></w:p></w:tc></w:tr>")
    neu = (z[0][:zelle] + '</w:pPr><w:r><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/>'
           '<w:szCs w:val="22"/></w:rPr><w:t>{{PROTOKOLLFUEHRUNG}}</w:t></w:r></w:p></w:tc></w:tr>')
    bl[k] = t.replace(z[0], neu)

    # Traktanden: je Zelle ein Absatz, Punkt hinter "Genehmigung Protokoll" weg.
    k = idx(lambda b: b.startswith("<w:tbl>") and "Genehmigung Protokoll" in b)
    t = felder_glaetten(bl[k])
    t = re.sub(r'(<w:t>Genehmigung Protokoll</w:t></w:r>)<w:r\b[^>]*>(?:(?!</w:r>).)*'
               r'<w:t xml:space="preserve">\. </w:t></w:r><w:r\b[^>]*>(?:(?!</w:r>).)*'
               r'<w:t xml:space="preserve"> </w:t></w:r>', r"\1", t, flags=re.S)

    def zelle_kuerzen(m):
        c = m.group(0)
        tcpr = re.match(r"<w:tc>(<w:tcPr>.*?</w:tcPr>)?", c, re.S)
        rest = c[tcpr.end():]
        erster = re.search(r"<w:p\b[^>]*(?<!/)>.*?</w:p>", rest, re.S)
        marken = "".join(re.findall(r"<w:bookmark(?:Start|End)\b[^>]*/>", rest[erster.end():]))
        return "<w:tc>" + (tcpr.group(1) or "") + rest[:erster.end()] + marken + "</w:tc>"
    zeilen = re.findall(r"<w:tr\b.*?</w:tr>", t, re.S)
    for r in zeilen[1:]:
        t = t.replace(r, re.sub(r"<w:tc>.*?</w:tc>", zelle_kuerzen, r, flags=re.S))
    bl[k] = t

    # Ende: senkrechte Hilfslinie weg, Beilagen-Block -> Datum, weniger Leerraum.
    k = idx(lambda b: "Gerade Verbindung" in b)
    bl[k] = re.sub(r"<w:r\b[^>]*>(?:(?!</w:r>).)*<mc:AlternateContent>.*?</mc:AlternateContent>"
                   r"</w:r>", "", bl[k], flags=re.S)
    if "<w:tab" in bl[k + 1] and not nur_text(bl[k + 1]).strip():
        del bl[k + 1]
    a = idx(lambda b: "CustomField.Beilagen" in b)
    e = idx(lambda b: "Organisation.City" in b, a)
    ppr = '<w:pPr><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:szCs w:val="22"/></w:rPr></w:pPr>'
    rpr_ = '<w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:szCs w:val="22"/></w:rPr>'
    def p(txt=""):
        return "<w:p>" + ppr + ('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr_, txt)
                                if txt else "") + "</w:p>"
    n = idx(lambda b: "Vorname/Name Protokollführung" in b, e)
    pend = idx(lambda b: "Pendenzenliste" in b, n)
    bl[a:pend] = [p("Grosswangen, {{DATUM}}"), p(), p("{{PROTOKOLLFUEHRUNG}}"), p()]

    body = felder_glaetten("".join(bl))
    for wort in ("DOCPROPERTY", "OawJumpToField", "Teilnehnde", "Doc.Enclosures"):
        if wort in body:
            raise SystemExit("protokoll: %r steht noch im Rumpf" % wort)
    setze(teile, "word/document.xml", kopf + body + doc[doc.rfind("</w:body>"):])

    schreiben(pfad, teile, infos)
    print("protokoll.docx: Kopfzeilen, Felder, Abstaende, Traktanden, Protokollfuehrung")


# ─── Prozess ────────────────────────────────────────────────────────────────

def prozess(hdr_tag):
    pfad = VORLAGEN / "prozess.docx"
    teile, infos = lesen(pfad)
    if "word/_rels/header1.xml.rels" in teile:
        print("prozess.docx: bereits angepasst")
        return

    def rechts(inhalt):
        return absatz(inhalt, fett=True, stil="Kopfzeile", jc="right").replace(
            '<w:pStyle w:val="Kopfzeile"/>',
            '<w:pStyle w:val="Kopfzeile"/><w:tabs><w:tab w:val="clear" w:pos="4536"/>'
            '<w:tab w:val="clear" w:pos="9072"/></w:tabs>')
    lauf = lambda s: '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr(True), s)
    hdr = (rechts(logo_schwebend("rIdLogo", LOGO_CY_PROZESS) + lauf("{{PROZESSGRUPPE}}")) +
           rechts(lauf("{{PROZESSNR}} {{PROZESSNAME}}")) +
           rechts(lauf("Prozessverantwortung: ") + lauf("{{VERANTWORTUNG}}")))
    setze(teile, "word/header1.xml", DECL + hdr_tag + hdr + "</w:hdr>")
    rels_ergaenzen(teile, "word/_rels/header1.xml.rels", "rIdLogo", REL_IMG, "media/image2.jpeg")
    teile["word/media/image2.jpeg"] = LOGO.read_bytes()
    content_types(teile)

    # Oberer Rand: Kopfzeile (1.25 cm) plus Logo (2 cm) plus Luft.
    doc = text(teile, "word/document.xml")
    doc, n = re.subn(r'<w:pgMar w:top="1417"', '<w:pgMar w:top="1985"', doc)
    if n != 1:
        raise SystemExit("prozess: pgMar nicht wie erwartet")
    setze(teile, "word/document.xml", doc)

    schreiben(pfad, teile, infos)
    print("prozess.docx: Logo in der Kopfzeile")


# ─── Prozess-Steckbrief ─────────────────────────────────────────────────────

def steckbrief():
    pfad = VORLAGEN / "prozess-steckbrief.docx"
    teile, infos = lesen(pfad)
    doc = text(teile, "word/document.xml")

    def balken(m):
        p = m.group(0)
        if 'w:fill="D9D9D9"' not in p.split("</w:pPr>")[0]:
            return p
        def rpr_fix(r):
            s = re.sub(r"<w:sz(?:Cs)?\b[^>]*/>", "", r.group(0))
            sz = '<w:sz w:val="22"/><w:szCs w:val="22"/>'
            # Reihenfolge im Schema: rFonts, b, bCs, ..., sz, szCs, ..., shd
            s, n = re.subn(r"(<w:b\s*/>(?:<w:bCs\s*/>)?)", r"\1" + sz, s, count=1)
            if n:
                return s
            return re.sub(r"(<w:rFonts\b[^>]*/>)", r"\1" + sz, s, count=1)
        return re.sub(r"<w:rPr>.*?</w:rPr>", rpr_fix, p, flags=re.S)
    neu = re.sub(r"<w:p\b[^>]*(?<!/)>.*?</w:p>", balken, doc, flags=re.S)
    neu = logo_descr_bereinigen(neu)
    if neu == doc and teile["word/media/image1.jpeg"] == LOGO.read_bytes():
        print("prozess-steckbrief.docx: bereits angepasst")
        return
    grau = [p for p in re.findall(r"<w:p\b[^>]*(?<!/)>.*?</w:p>", neu, re.S)
            if 'w:fill="D9D9D9"' in p.split("</w:pPr>")[0]]
    if not grau or any('<w:sz w:val="22"/>' not in p for p in grau):
        raise SystemExit("steckbrief: grauer Balken nicht auf 11 Punkt")
    setze(teile, "word/document.xml", neu)
    teile["word/media/image1.jpeg"] = LOGO.read_bytes()
    schreiben(pfad, teile, infos)
    print("prozess-steckbrief.docx: Titel im grauen Balken 11 Punkt, leichtes Logo")


def main():
    if not LOGO.exists():
        raise SystemExit("assets/logo-linde.jpg fehlt")
    tag = hdr_wurzel()
    brief(tag)
    protokoll(tag)
    prozess(tag)
    steckbrief()
    return 0


if __name__ == "__main__":
    sys.exit(main())
