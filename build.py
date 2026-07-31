#!/usr/bin/env python3
"""Baut taskpane.html aus src/taskpane.template.html.

Das Add-In ist bewusst eine einzige, in sich geschlossene HTML-Datei (wie beim
BWH-Add-In): GitHub Pages liefert sie aus, Word laedt sie direkt. Die grossen
Blobs -- Logo (PNG, base64) sowie styles.xml und numbering.xml aus der
Muster-Vorlage -- werden hier hineingeschrieben, damit die Quelldatei lesbar
bleibt.

    python3 build.py

Platzhalter in src/taskpane.template.html:
    @@LOGO_PNG_B64@@   base64 von assets/logo-dok.png
    @@STYLES_XML@@     assets/styles.xml + assets/styles-extra.xml (einzeilig)
    @@NUMBERING_XML@@  assets/numbering.xml (einzeilig)
    @@BUILD@@          Build-Nummer aus BUILD
"""

import base64
import pathlib
import re
import sys

ROOT = pathlib.Path(__file__).resolve().parent
SRC = ROOT / "src" / "taskpane.template.html"
OUT = ROOT / "taskpane.html"


def js_string(s: str) -> str:
    """XML als einfach-gequoteten JS-String ausgeben."""
    return s.replace("\\", "\\\\").replace("'", "\\'")


def load_xml(path: pathlib.Path) -> str:
    """XML einlesen, XML-Deklaration entfernen, Whitespace zwischen Tags weg."""
    txt = path.read_text(encoding="utf-8")
    txt = re.sub(r"^\s*<\?xml[^>]*\?>\s*", "", txt)
    txt = re.sub(r">\s+<", "><", txt.strip())
    return txt


def main() -> int:
    styles = load_xml(ROOT / "assets" / "styles.xml")
    extra = load_xml(ROOT / "assets" / "styles-extra.xml")
    extra = re.sub(r"<!--.*?-->", "", extra, flags=re.S).strip()
    if "</w:styles>" not in styles:
        print("assets/styles.xml sieht nicht wie styles.xml aus", file=sys.stderr)
        return 1
    styles = styles.replace("</w:styles>", extra + "</w:styles>")

    numbering = load_xml(ROOT / "assets" / "numbering.xml")
    logo_b64 = base64.b64encode((ROOT / "assets" / "logo-dok.png").read_bytes()).decode()
    build_no = (ROOT / "BUILD").read_text(encoding="utf-8").strip()

    html = SRC.read_text(encoding="utf-8")
    for token, value in (
        ("@@LOGO_PNG_B64@@", logo_b64),
        ("@@STYLES_XML@@", js_string(styles)),
        ("@@NUMBERING_XML@@", js_string(numbering)),
        ("@@BUILD@@", build_no),
    ):
        if token not in html:
            print("Platzhalter fehlt in der Vorlage: " + token, file=sys.stderr)
            return 1
        html = html.replace(token, value)

    OUT.write_text(html, encoding="utf-8")
    print("taskpane.html geschrieben (%d KB, Build %s)" % (len(html) // 1024, build_no))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
