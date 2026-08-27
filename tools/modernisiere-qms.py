#!/usr/bin/env python3
"""Zieht die QMS-Hilfsmittel-Vorlage auf den modernen Linde-Kopf und -Fuss um.

    node tools/dump-kopffuss.js   # legt hdr-body.xml, ftr-body.xml, logo.b64 ab
    python3 tools/modernisiere-qms.py <dump-ordner>

Laeuft NACH tools/clean-vorlagen.py und veraendert nur
vorlagen/qms-hilfsmittel.docx. Die uebrigen vier Vorlagen behalten auf
Wunsch der Linde ihr bisheriges Erscheinungsbild.

Was passiert:
  * der alte, schwarz umrandete Tabellenkopf im Rumpf (Bereich / Nummer /
    Dokumentart / Titel / Logo) wird entfernt
  * stattdessen: moderne Kopfzeile wie in Muster_Dokument_A4_hoch.docx -
    Logo links schwebend, rechts {{BEREICH}} / {{NUMMER}} fett ueber
    {{DOKUMENTART}}, Trennlinie darunter
  * moderne Fusszeile: Adressblock mit Rauten-Trennzeichen links,
    "Seite X von Y" und Dateiname/Speicherdatum rechts, Trennlinie darueber
  * {{TITEL}} kommt als zentrierter Dokumenttitel an den Rumpfanfang
  * Seitenraender auf die Masse der Muster-Vorlage

Kopf und Fuss stammen aus denselben Bausteinen, die auch der Zweig
"Nur Kopf- und Fusszeile" im Add-In verwendet (tools/dump-kopffuss.js
leitet sie aus taskpane.html ab) - eine Quelle, ein Erscheinungsbild.
"""

import json
import pathlib
import re
import sys
import zipfile

ROOT = pathlib.Path(__file__).resolve().parent.parent
ZIEL = ROOT / "vorlagen" / "qms-hilfsmittel.docx"

DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
REL_HDR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
CT_HDR = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
PGMAR = ('<w:pgMar w:top="1701" w:right="1134" w:bottom="1304" w:left="1134" '
         'w:header="680" w:footer="510" w:gutter="0" />')
TITEL_P = ('<w:p><w:pPr><w:spacing w:before="0" w:after="240"/>'
           '<w:jc w:val="center"/>'
           '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
           '<w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:pPr>'
           '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
           '<w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>'
           '<w:t>{{TITEL}}</w:t></w:r></w:p>')


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    dump = pathlib.Path(sys.argv[1])
    hdr_body = (dump / "hdr-body.xml").read_text(encoding="utf-8")
    ftr_body = (dump / "ftr-body.xml").read_text(encoding="utf-8")
    logo_b64 = (dump / "logo.b64").read_text()
    meta = json.loads((dump / "meta.json").read_text())

    zin = zipfile.ZipFile(ZIEL)
    teile = {i.filename: zin.read(i.filename) for i in zin.infolist()}
    infos = {i.filename: i for i in zin.infolist()}

    doc = teile["word/document.xml"].decode("utf-8")

    # Namensraeume der Kopf-/Fusszeile: dieselben wie im Add-In-Paket.
    ns = meta["docNs"]

    # 1) Alten Tabellenkopf entfernen, {{TITEL}} an den Anfang setzen.
    m = re.search(r"<w:body>(\s*)<w:tbl>.*?</w:tbl>", doc, re.S)
    if not m:
        raise SystemExit("Tabellenkopf im Rumpf nicht gefunden")
    doc = doc[:m.start()] + "<w:body>" + TITEL_P + doc[m.end():]
    if "{{TITEL}}" not in doc:
        raise SystemExit("Titel-Absatz fehlt nach dem Umbau")

    # 2) sectPr: Kopfzeilen-Verweis ergaenzen, Raender der Muster-Vorlage.
    doc, n = re.subn(r'<w:footerReference\b',
                     '<w:headerReference w:type="default" r:id="rIdHdrModern" />'
                     '<w:footerReference', doc, count=1)
    if n != 1:
        raise SystemExit("footerReference nicht gefunden")
    doc = re.sub(r"<w:pgMar[^>]*/>", PGMAR, doc, count=1)
    teile["word/document.xml"] = doc.encode("utf-8")

    # 3) Kopfzeile samt Logo und Beziehungen anlegen.
    teile["word/header1.xml"] = (DECL + "<w:hdr " + ns + ">" + hdr_body +
                                 "</w:hdr>").encode("utf-8")
    teile["word/_rels/header1.xml.rels"] = (
        DECL + '<Relationships xmlns="%s">'
        '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/image" Target="media/%s"/>'
        "</Relationships>" % (RELS_NS, meta["logoRid"], meta["logoFn"])
    ).encode("utf-8")
    import base64
    teile["word/media/" + meta["logoFn"]] = base64.b64decode(logo_b64)

    # 4) Fusszeile ersetzen (Teilname und Verweis bleiben).
    teile["word/footer1.xml"] = (DECL + "<w:ftr " + ns + ">" + ftr_body +
                                 "</w:ftr>").encode("utf-8")

    # 5) document.xml.rels: Verweis auf die neue Kopfzeile.
    rels = teile["word/_rels/document.xml.rels"].decode("utf-8")
    if "rIdHdrModern" not in rels:
        rels = rels.replace("</Relationships>",
                            '<Relationship Id="rIdHdrModern" Type="%s" '
                            'Target="header1.xml"/></Relationships>' % REL_HDR)
    teile["word/_rels/document.xml.rels"] = rels.encode("utf-8")

    # 6) Inhaltstypen: Kopfzeile eintragen, png-Standard sicherstellen.
    ct = teile["[Content_Types].xml"].decode("utf-8")
    if "/word/header1.xml" not in ct:
        ct = ct.replace("</Types>",
                        '<Override PartName="/word/header1.xml" ContentType="%s"/>'
                        "</Types>" % CT_HDR)
    if 'Extension="png"' not in ct:
        ct = ct.replace("</Types>",
                        '<Default Extension="png" ContentType="image/png"/></Types>')
    teile["[Content_Types].xml"] = ct.encode("utf-8")

    with zipfile.ZipFile(ZIEL, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in teile.items():
            zout.writestr(infos.get(name, name), data)

    print("vorlagen/qms-hilfsmittel.docx modernisiert "
          "(Kopf- und Fusszeile wie Muster_Dokument_A4_hoch)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
