#!/usr/bin/env python3
"""Baut taskpane.html aus src/taskpane.template.html.

Das Add-In ist bewusst eine einzige, in sich geschlossene HTML-Datei (wie beim
BWH-Add-In): GitHub Pages liefert sie aus, Word laedt sie direkt. Das Logo
wird hier als base64 hineingeschrieben, damit die Quelldatei lesbar bleibt.

    python3 build.py

Platzhalter in src/taskpane.template.html:
    @@LOGO_PNG_B64@@   base64 von assets/logo-dok.png
    @@BUILD@@          Build-Nummer aus BUILD
"""

import base64
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent
SRC = ROOT / "src" / "taskpane.template.html"
OUT = ROOT / "taskpane.html"


def main() -> int:
    logo_b64 = base64.b64encode((ROOT / "assets" / "logo-dok.png").read_bytes()).decode()
    build_no = (ROOT / "BUILD").read_text(encoding="utf-8").strip()

    html = SRC.read_text(encoding="utf-8")
    for token, value in (
        ("@@LOGO_PNG_B64@@", logo_b64),
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
