# Linde Dokumentenvorlagen – Word-Add-In

Word-Add-In für das Betagtenzentrum Linde Grosswangen. Es setzt Kopfzeile,
Fusszeile, Seitenformat und – auf Wunsch – Musterinhalt, Inhaltsverzeichnis,
Titelblatt oder eine Briefvorlage in ein bestehendes Word-Dokument.

Aufbau und Bedienung entsprechen dem BWH-Add-In
([nicolatrava/bwh-addin](https://github.com/nicolatrava/bwh-addin)); die
Gestaltung stammt aus `Muster_Dokument_A4_hoch.docx` (in diesem Repository
abgelegt).

## Was das Add-In macht

| Schritt | Auswahl |
|---|---|
| 1 Dokumenttyp | Internes Dokument · Externes Dokument · Brief |
| 2 Format | A4 Hoch · A4 Quer (beim Brief immer A4 Hoch) |
| 3 Inhalt | Musterinhalt · Inhaltsverzeichnis · Titelblatt (nur A4 Hoch) |
| 4 Bereich | Sekretariat, Pflege und Betreuung, Hotellerie … oder eigener Text |
| 5 Metadaten | Dokument-Nr., Prozessname, Ablageort, Kürzel, neue Version; beim Brief das persönliche Profil |
| 6 Aktualisieren | Fusszeile neu schreiben · Felder und Inhaltsverzeichnis aktualisieren |

**Kopfzeile** – zweispaltige Tabelle mit Trennlinie unten, links das Logo,
rechts zwei Zeilen: `Bereich / Dokument-Nr. Prozessname` (fett) und darunter
der Ablageort. Genau wie in der Muster-Vorlage
(«Sekretariat / C 3.4.2 Kreditoren» über «Ordner Kreditoren»). Externe
Dokumente bekommen nur Logo und Linie, ohne interne Metadaten.

**Fusszeile** – Tabelle mit Trennlinie oben, links der Adressblock mit
Wingdings-Rauten als Trennzeichen, rechts «Seite X von Y» und darunter
`Dateiname / Speicherdatum`, ergänzt um Version und Kürzel.

**Seitenformat** – A4 mit den Rändern der Vorlage (oben 3,0 cm, links und
rechts 2,0 cm, unten 2,3 cm, Kopfzeile 1,2 cm, Fusszeile 0,9 cm) bzw. dieselben
Ränder im Querformat.

Vorhandener Text bleibt immer erhalten: Kopf- und Fusszeile werden ersetzt,
neue Inhalte kommen davor. Gespeichert wird nichts automatisch.

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
* `BEREICHE` – Liste der Bereiche für Schritt 4. Der letzte Eintrag
  («Eigener Bereich») blendet ein Freitextfeld ein.
* `AAD_CLIENT_ID` – siehe *Active Directory* unten.

Nach jeder Änderung neu bauen und die Build-Nummer erhöhen:

```bash
echo 2 > BUILD          # Build-Nummer hochzählen
python3 build.py        # erzeugt taskpane.html
node tools/test.js      # Bedienlogik und OOXML prüfen
```

Ausserdem in `manifest.xml` die `?v=`-Nummer der drei URLs und die `<Version>`
erhöhen, damit Word die neue Fassung lädt und nicht die zwischengespeicherte.

### Active Directory (optional)

Beim Brief kann das Profil (Name, Funktion, Telefon, E-Mail) aus dem
Verzeichnis geladen werden. Dafür braucht es eine App-Registrierung im
Microsoft-Entra-Portal der Linde:

1. Neue App-Registrierung, Kontotyp «nur eigenes Verzeichnis».
2. Plattform *Single-Page-Anwendung*, Umleitungs-URI
   `https://nicolatrava.github.io/BZ-Linde-Add-In/taskpane.html`.
3. API-Berechtigung *Microsoft Graph → User.Read* (delegiert).
4. Die Anwendungs-(Client-)ID in `src/taskpane.template.html` bei
   `AAD_CLIENT_ID` eintragen und neu bauen.

Ohne Registrierung bleibt die Schaltfläche wirkungslos; das Profil wird von
Hand erfasst und danach auf dem Rechner gemerkt. Alles andere funktioniert
unabhängig davon.

## Entwicklung

```
manifest.xml                    Add-In-Manifest für Word
taskpane.html                   ausgeliefertes Add-In (erzeugt – nicht direkt bearbeiten)
src/taskpane.template.html      Quelle: Bedienoberfläche und OOXML-Erzeugung
build.py                        baut taskpane.html aus src/ und assets/
BUILD                           Build-Nummer, erscheint unten im Taskpane
assets/logo-dok.png             Logo, freigestellt aus der Muster-Vorlage
assets/styles.xml               styles.xml der Muster-Vorlage
assets/styles-extra.xml         Zusatz-Stile (Verzeichnis 2/3, Hyperlink)
assets/numbering.xml            numbering.xml der Muster-Vorlage
icon16/32/80.png                Symbole für das Menüband
msal-browser.min.js             MSAL für die Anmeldung am Verzeichnis
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
