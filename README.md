# Linde Dokumentenvorlagen – Word-Add-In

Word-Add-In für das Betagtenzentrum Linde Grosswangen. Es legt die
Hausvorlagen in ein Word-Dokument und bringt bestehende Dokumente ins CI.
Aufbau und Bedienung folgen dem BWH-Add-In
([nicolatrava/bwh-addin](https://github.com/nicolatrava/bwh-addin)).

## Was das Add-In macht

**Schritt 1 – Vorlage**

| Vorlage | Inhalt |
|---|---|
| Brief | Logo, Empfängeradresse, Betreff, Grussformel mit Anrede und E-Mail |
| Protokoll | Traktandenliste, Pendenzen, Verteiler, Protokollführung |
| Prozess | Ablauftabelle: Grafik · Tätigkeit · Wer · Hilfsmittel |
| Prozess-Steckbrief | Ziel, Input/Output, Messgrössen, Chancen und Risiken |
| QMS-Hilfsmittel | Richtlinie oder Merkblatt: Ziel, Grundsätze, Vorgehen, Hinweise |
| Nur Kopf- und Fusszeile | bestehendes Dokument ins CI bringen, Text bleibt |

**Schritt 2 – Angaben.** Je nach Vorlage Prozessgruppe, Prozess-Nr.,
Prozessname, Bereich, Dokumentart oder Titel. Was leer bleibt, bleibt als
Platzhalter im Dokument stehen. Das Datum setzt das Add-In immer selbst,
beim Brief auch die Grussanrede («Sehr geehrte Frau Muster» aus Anrede und
Name). Absender, E-Mail und Protokollführung merkt sich das Add-In auf dem
Gerät (`localStorage`, Schlüssel `lindeFelder`) und trägt sie beim nächsten
Mal von selbst ein.

Die fünf Vorlagen werden **nicht nachgebaut**: unter `vorlagen/` liegen die
Original-.docx der Linde, das Add-In holt die gewählte Datei und legt sie mit
`insertFileFromBase64` ins Dokument. Damit stammen Kopfzeile, Fusszeile,
Stile und Logo eins zu eins aus den Originalen. Anschliessend ersetzt das
Add-In die Platzhalter `{{...}}` in Rumpf, Kopf- und Fusszeilen.

> Eine Vorlage einzusetzen **ersetzt das ganze Dokument**. Der Zweig
> «Nur Kopf- und Fusszeile» lässt den Text unangetastet.

Word braucht dafür WordApi 1.5. Ältere Versionen bekommen nur den Rumpf
eingesetzt; das Add-In sagt das im Statustext.

Beim Einsetzen übergibt das Add-In Import-Optionen (`importStyles`,
`importParagraphSpacing`, `importTheme`): Formatvorlagen, Absatzabstände und
Design kommen aus der Vorlage. Ohne diese Optionen legt Word die Vorlagen des
Zieldokuments darüber – ein frisches Word-Dokument bringt «Standard» mit
8 pt Abstand nach und Zeilenabstand 1.16 sowie eine blaue «Überschrift 1»
mit, und genau so sah der erste Brief bei der Linde aus. Schlägt der Aufruf
mit Optionen fehl, versucht es das Add-In ohne, danach nur mit dem Rumpf.
Der Brief ist zusätzlich gegen dieses Verhalten gebaut: Er verwendet eigene
Formatvorlagen «Linde Brief», «Linde Betreff», «Linde Kopfzeile» und
«Linde Fusszeile», die kein Zieldokument kennt und Word darum unverändert
mitnimmt.

### Nur Kopf- und Fusszeile

| Schritt | Auswahl |
|---|---|
| 2 Dokumenttyp | Internes Dokument · Externes Dokument · Brief |
| 3 Format | A4 Hoch · A4 Quer (beim Brief immer A4 Hoch) |
| 4 Kopfzeile | die beiden Zeilen rechts oben, Kürzel, neue Version |
| 5 Aktualisieren | Fusszeile neu schreiben · Felder und Inhaltsverzeichnis aktualisieren |

Gestaltung nach `Muster_Dokument_A4_hoch.docx` (in diesem Repository abgelegt).

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

## Vorlagen pflegen

Die Originale kommen ausgefüllt aus dem Betrieb und enthalten echte
Personendaten – Empfängeradressen, Namen von Mitarbeitenden, Pfade aus dem
P-Laufwerk – sowie Rückstände der Fremdvorlage, aus der die Protokoll-Vorlage
stammt (SSBL Luzern, bis hinein in `settings.xml` und `theme1.xml`).
**Dieses Repository ist öffentlich: die Originale gehören nicht hinein.**

`tools/clean-vorlagen.py` macht daraus leere Vorlagen und protokolliert jede
Ersetzung. Zum Schluss läuft eine Sperrliste über jede erzeugte Datei; schlägt
sie an, bricht das Skript ab.

```bash
python3 tools/clean-vorlagen.py /pfad/zu/den/originalen
node tools/dump-kopffuss.js /tmp/kopffuss          # moderne Kopf-/Fusszeile ausleiten
python3 tools/modernisiere-qms.py /tmp/kopffuss    # nur die QMS-Vorlage umziehen
python3 tools/feedback-anpassungen.py              # Rückmeldungen der Linde einarbeiten
python3 tools/check-vorlagen.py     # Platzhalter gegen die Felder im Add-In
```

Die QMS-Hilfsmittel-Vorlage bekommt dabei – auf Wunsch der Linde als einzige –
den modernen Kopf und Fuss aus `Muster_Dokument_A4_hoch.docx`; die Bausteine
stammen aus denselben Funktionen wie der Zweig «Nur Kopf- und Fusszeile». Die
übrigen vier Vorlagen behalten ihr bisheriges Erscheinungsbild.

`tools/feedback-anpassungen.py` setzt die Rückmeldungen aus der Prüfung im
September 2026 um (Brief: Logo mittig, Adressblock und Leerzeilen wie das
freigegebene Muster, Aptos 11/12 schwarz über eigene Linde-Formatvorlagen,
Grussformel mit E-Mail, Adress-Fusszeile · Protokoll: Abstände, Traktanden-Zellen,
Protokollführung, Felder der Fremdvorlage aufgelöst · Prozess: Logo in der
Kopfzeile · Steckbrief: grauer Balken 11 Punkt). Es arbeitet auf den bereits
bereinigten Dateien, erkennt angepasste Vorlagen und lässt sich darum
beliebig oft laufen lassen. Alle vier nutzen dasselbe Logo
`assets/logo-linde.jpg` – aus dem 1-MB-CMYK-Original der Protokoll-Vorlage
über sein ICC-Profil farbtreu nach sRGB gewandelt, 57 KB.

`tools/check-vorlagen.py` vergleicht beide Richtungen: ein `{{...}}` ohne
Eingabefeld bliebe im fertigen Dokument stehen, ein Eingabefeld ohne
Platzhalter liefe ins Leere.

Neue Vorlage aufnehmen: Datei nach `vorlagen/`, Eintrag im `VORLAGEN`-Array in
`src/taskpane.template.html`, Platzhalter in Grossbuchstaben in die Datei
schreiben, `check-vorlagen.py` laufen lassen.

## Anpassen

Die häufigsten Anpassungen stehen alle zuoberst in
`src/taskpane.template.html`:

* `ORG_DEFAULT` – Adresse für die Fusszeile. Sie lässt sich zusätzlich im
  Taskpane unter *Adresse der Fusszeile* ändern; die Änderung merkt sich der
  jeweilige Rechner.

Nach jeder Änderung neu bauen und die Build-Nummer erhöhen:

```bash
echo 8 > BUILD          # Build-Nummer hochzählen
python3 build.py        # erzeugt taskpane.html
node tools/test.js      # Bedienlogik und OOXML prüfen
python3 tools/check-vorlagen.py
```

**Das `manifest.xml` bleibt dabei unangetastet.** Es liegt nach dem einmaligen
Hochladen im Microsoft-365-Mandanten, nicht auf GitHub – Word holt von hier nur
`taskpane.html`, `vorlagen/` und die Icons. Änderungen an Oberfläche, Feldern
oder Vorlagen erreichen die Nutzenden also allein über einen Push, ohne dass
jemand etwas neu einspielen muss. Beim BWH-Add-In steht das Manifest seit
`?v=22` still, während das Taskpane bei Build 46 ist.

Ein neues Manifest braucht es nur, wenn sich etwas ändert, das *im* Manifest
steht: Name oder Symbol im Menüband, Berechtigungen, die Adresse des Hostings.
Die Build-Nummer unten im Taskpane zeigt, welcher Stand geladen ist.

### Ausrollen

`main` **allein** pushen:

```bash
git push origin main
```

Werden mehrere Branches in einem Befehl gepusht (`git push origin main feature`),
startet GitHub Pages keinen Build – der Push kommt als ein Ereignis an und
Pages sieht `main` nicht. Das Repository liegt dann neu, die ausgelieferte
Seite aber weiterhin auf dem alten Stand, und zwar für alle: Desktop wie Web.
Cache-Löschen hilft dagegen nichts.

Nach dem Push kontrollieren, ob wirklich deployt wurde – unter *Actions* muss
ein neuer Lauf «pages build and deployment» auf dem gepushten Commit stehen.
Bleibt er aus, einen leeren Commit nachschieben:

```bash
git commit --allow-empty -m "Pages-Deployment anstossen" && git push origin main
```

## Entwicklung

```
manifest.xml                    Add-In-Manifest für Word
taskpane.html                   ausgeliefertes Add-In (erzeugt – nicht direkt bearbeiten)
src/taskpane.template.html      Quelle: Bedienoberfläche und OOXML-Erzeugung
build.py                        baut taskpane.html aus src/ und assets/
BUILD                           Build-Nummer, erscheint unten im Taskpane
assets/logo-dok.png             Logo, freigestellt aus der Muster-Vorlage (Kopfzeile des Add-Ins)
assets/logo-linde.jpg           Logo der Hausvorlagen (Brief, Protokoll, Prozess, Steckbrief)
icon16/32/80.png                Symbole für das Menüband
vorlagen/*.docx                 die fünf Hausvorlagen, bereinigt
tools/test.js                   Schnelltest ohne Word
tools/preview.js                erzeugt Beispiel-Dokumente als .docx
tools/clean-vorlagen.py         macht aus ausgefüllten Originalen leere Vorlagen
tools/dump-kopffuss.js          leitet die moderne Kopf-/Fusszeile aus taskpane.html ab
tools/modernisiere-qms.py       zieht die QMS-Vorlage auf den modernen Kopf/Fuss um
tools/feedback-anpassungen.py   arbeitet die Rückmeldungen der Linde in die Vorlagen ein
tools/check-vorlagen.py         Platzhalter gegen die Felder im Add-In prüfen
Muster_Dokument_A4_hoch.docx    Ausgangsvorlage für Kopf-/Fusszeile
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
