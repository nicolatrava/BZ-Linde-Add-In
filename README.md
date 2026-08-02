# Linde Dokumentenvorlagen – Word-Add-In

Word-Add-In für das Betagtenzentrum Linde Grosswangen. Es setzt Kopfzeile,
Fusszeile und Seitenformat eines bestehenden Word-Dokuments auf die Gestaltung
aus `Muster_Dokument_A4_hoch.docx` (in diesem Repository abgelegt).

Aufbau und Bedienung folgen dem BWH-Add-In
([nicolatrava/bwh-addin](https://github.com/nicolatrava/bwh-addin)). Bewusst
**nicht** übernommen sind dessen Musterinhalt, Inhaltsverzeichnis und
Titelblatt: dafür gibt es bei der Linde keine eigene Vorlage, und erfundener
Inhalt wäre schlechter als gar keiner. Sobald ein echtes Musterdokument
vorliegt, lässt es sich als weiterer Baustein ergänzen.

## Was das Add-In macht

| Schritt | Auswahl |
|---|---|
| 1 Dokumenttyp | Internes Dokument · Externes Dokument · Brief |
| 2 Format | A4 Hoch · A4 Quer (beim Brief immer A4 Hoch) |
| 3 Kopfzeile | die beiden Zeilen rechts oben, Kürzel, neue Version |
| 4 Aktualisieren | Fusszeile neu schreiben · Felder und Inhaltsverzeichnis aktualisieren |

**Kopfzeile** – zweispaltige Tabelle mit Trennlinie unten, links das
schwebende Logo, rechts zwei frei erfassbare Zeilen: die erste fett, die
zweite normal – genau wie in der Muster-Vorlage («Sekretariat / C 3.4.2
Kreditoren» über «Ordner Kreditoren»). Bleiben die Felder leer, stehen dort
«Prozessname» und «Ablageort» als Platzhalter. Externe Dokumente bekommen nur Logo
und Linie, ohne interne Metadaten.

Der Anker des Logos ist bewusst 1:1 aus der Vorlage übernommen (waagrecht am
linken Seitenrand, senkrecht −1.62 cm ab dem Trägerabsatz nach der Tabelle).
Ein seitenbezogener Anker setzt Word im Browser nicht um – das Logo landet
dann mitten im Text.

**Fusszeile** – Tabelle mit Trennlinie oben, links der Adressblock mit
Wingdings-Rauten als Trennzeichen, rechts «Seite X von Y» und darunter
`Dateiname / Speicherdatum`, ergänzt um Version und Kürzel.

**Seitenformat** – A4 mit den Rändern der Vorlage (oben 3,0 cm, links und
rechts 2,0 cm, unten 2,3 cm, Kopfzeile 1,2 cm, Fusszeile 0,9 cm) bzw. dieselben
Ränder im Querformat.

Der Text des Dokuments wird nie angefasst – nur Kopfzeile, Fusszeile und
Seitenformat. Gespeichert wird nichts automatisch. Kopf- und Fusszeile setzen
Arial ausdrücklich, sonst erben sie die Schrift des Dokuments (in einem
frischen Word-Dokument wäre das Aptos).

## Installation

Das Add-In liegt auf GitHub Pages unter
`https://nicolatrava.github.io/BZ-Linde-Add-In/`. Verteilt wird es über
`manifest.xml`:

* **Für die ganze Organisation** – Microsoft 365 Admin Center →
  *Einstellungen* → *Integrierte Apps* → *App hochladen* → *Office-Add-In* →
  `manifest.xml` hochladen.
* **Zum Ausprobieren (einzelner Rechner)** – `manifest.xml` in einen
  freigegebenen Netzwerkordner legen, diesen in Word unter *Datei* →
  *Optionen* → *Trust Center* → *Einstellungen für das Trust Center* →
  *Vertrauenswürdige Add-In-Kataloge* eintragen, Word neu starten und das
  Add-In über *Einfügen* → *Meine Add-Ins* → *Freigegebener Ordner* laden.

Danach erscheint in Word auf dem Register *Start* die Gruppe **Linde Vorlagen**.

## Anpassen

Die häufigsten Anpassungen stehen alle zuoberst in
`src/taskpane.template.html`:

* `ORG_DEFAULT` – Adresse für die Fusszeile. Sie lässt sich zusätzlich im
  Taskpane unter *Adresse der Fusszeile* ändern; die Änderung merkt sich der
  jeweilige Rechner.

Nach jeder Änderung neu bauen und die Build-Nummer erhöhen:

```bash
echo 5 > BUILD          # Build-Nummer hochzählen
python3 build.py        # erzeugt taskpane.html
node tools/test.js      # Bedienlogik und OOXML prüfen
```

Ausserdem in `manifest.xml` die `?v=`-Nummer der drei URLs und die `<Version>`
erhöhen, damit Word die neue Fassung lädt und nicht die zwischengespeicherte.

## Entwicklung

```
manifest.xml                    Add-In-Manifest für Word
taskpane.html                   ausgeliefertes Add-In (erzeugt – nicht direkt bearbeiten)
src/taskpane.template.html      Quelle: Bedienoberfläche und OOXML-Erzeugung
build.py                        baut taskpane.html aus src/ und assets/
BUILD                           Build-Nummer, erscheint unten im Taskpane
assets/logo-dok.png             Logo, freigestellt aus der Muster-Vorlage
icon16/32/80.png                Symbole für das Menüband
tools/test.js                   Schnelltest ohne Word
tools/preview.js                erzeugt Beispiel-Dokumente als .docx
Muster_Dokument_A4_hoch.docx    Ausgangsvorlage
```

`tools/preview.js` schreibt fertige Word-Dateien aller Varianten nach
`preview/` – praktisch, um eine Änderung anzusehen, ohne das Add-In erst
ausrollen zu müssen:

```bash
node tools/preview.js
```

Der Test prüft neben der Bedienlogik auch die vom OOXML-Schema vorgegebene
Reihenfolge der XML-Elemente. Word repariert eine falsche Reihenfolge nicht,
sondern meldet das Dokument als beschädigt – deshalb vor jedem Ausrollen
`node tools/test.js` laufen lassen.
