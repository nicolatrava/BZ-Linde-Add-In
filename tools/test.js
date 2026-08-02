#!/usr/bin/env node
/*
 * Schnelltest für taskpane.html – prüft die Bedienlogik und die erzeugten
 * OOXML-Bausteine, ohne Word zu starten.
 *
 *   node tools/test.js
 *
 * Der Skriptteil von taskpane.html läuft in einer Sandbox mit minimal
 * nachgebautem DOM, Office und Word.
 */
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.resolve(__dirname, '..');

const fields = {};
function el(id) {
  if (!fields[id]) {
    fields[id] = { value: '', checked: false, className: '', textContent: '',
                   style: {}, disabled: false, innerHTML: '',
                   classList: { add() {}, remove() {} },
                   getAttribute() { return '0'; }, querySelector() { return null; } };
  }
  return fields[id];
}

const sandbox = {
  console, setTimeout, Promise,
  localStorage: { getItem() { return null; }, setItem() {}, removeItem() {} },
  document: { getElementById: el, querySelectorAll() { return []; }, body: {} },
  Office: {
    HostType: { Word: 'Word' },
    onReady() {},
    context: {
      requirements: { isSetSupported() { return false; } },
      document: {
        url: '',
        settings: { get() { return null; }, set() {}, saveAsync() {}, refreshAsync() {} },
        getFilePropertiesAsync(cb) { cb({ value: { url: '' } }); }
      }
    }
  },
  Word: { InsertLocation: { start: 'Start' }, run() {} }
};
sandbox.window = sandbox;

const html = fs.readFileSync(path.join(ROOT, 'taskpane.html'), 'utf8');
const m = html.match(/<script>\n([\s\S]*?)\n<\/script>/);
if (!m) { console.error('Inline-<script> nicht gefunden'); process.exit(1); }
vm.createContext(sandbox);
vm.runInContext(m[1], sandbox, { filename: 'taskpane.html' });

let failed = 0;
function t(name, got, exp) {
  const ok = got === exp;
  if (!ok) failed++;
  console.log((ok ? 'ok  ' : 'FAIL') + '  ' + name + (ok ? '' : ' -> ' + got + ' (erwartet ' + exp + ')'));
}

const S = sandbox;

/* ── Bedienlogik ──────────────────────────────────────────────────────────── */
S.setFmt('hoch'); S.setTyp('intern');
S.toggleInhalt('titel');  t('A4 Hoch: Titelblatt wählbar', !!S.inhaltSet.titel, true);
S.setFmt('quer');         t('A4 Quer: Titelblatt wird abgewählt', !!S.inhaltSet.titel, false);
S.toggleInhalt('titel');  t('A4 Quer: Titelblatt bleibt gesperrt', !!S.inhaltSet.titel, false);
S.toggleInhalt('ivz');    t('A4 Quer: Inhaltsverzeichnis wählbar', !!S.inhaltSet.ivz, true);
S.setTyp('brief');        t('Brief: Inhaltsverzeichnis wird abgewählt', !!S.inhaltSet.ivz, false);
                          t('Brief erzwingt A4 Hoch', S.fmt, 'hoch');
S.toggleInhalt('muster'); t('Brief: Musterinhalt wählbar', !!S.inhaltSet.muster, true);
S.setTyp('intern');

/* ── Kopfzeile ────────────────────────────────────────────────────────────── */
const hdr = S.mkHdrPkg(false, true, 'Sekretariat / C 3.4.2 Kreditoren', 'Ordner Kreditoren');
t('Kopfzeile: Zeile 1', hdr.indexOf('Sekretariat / C 3.4.2 Kreditoren') > 0, true);
t('Kopfzeile: Zeile 2', hdr.indexOf('Ordner Kreditoren') > 0, true);
t('Kopfzeile: Logo eingebettet', hdr.indexOf('image/png') > 0, true);
t('Kopfzeile: Trennlinie', hdr.indexOf('w:bottom w:val="single" w:sz="4" w:space="0" w:color="999999"') > 0, true);
const hdrX = S.mkHdrPkg(false, false, 'darf nicht erscheinen', 'auch nicht');
t('Externes Dokument: keine Kopfzeilen-Metadaten', hdrX.indexOf('darf nicht erscheinen'), -1);

/* ── Fusszeile ────────────────────────────────────────────────────────────── */
const ftr = S.mkFtrPkg(false, S.ORG_DEFAULT, '31.07.26', 'NT', true);
t('Fusszeile: Adresse', ftr.indexOf('Dorfstrasse 6e') > 0, true);
t('Fusszeile: Version', ftr.indexOf('Version: 31.07.26') > 0, true);
t('Fusszeile: Kürzel', ftr.indexOf('Kürzel: NT') > 0, true);
t('Fusszeile: Wingdings-Trennzeichen', ftr.indexOf('w:sym w:font="Wingdings" w:char="F09F"') > 0, true);
t('Fusszeile: Seitenzahlen', ftr.indexOf('NUMPAGES') > 0, true);
const ftrB = S.mkFtrPkg(false, S.ORG_DEFAULT, '', '', false);
t('Brief-Fusszeile: ohne Dateinamen', ftrB.indexOf('FILENAME'), -1);

/* ── Inhalte ──────────────────────────────────────────────────────────────── */
const c = S.mkContentPkg(['titel', 'ivz', 'muster'], false);
t('Inhalt: styles.xml mitgeliefert', c.indexOf('/word/styles.xml') > 0, true);
t('Inhalt: numbering.xml mitgeliefert', c.indexOf('/word/numbering.xml') > 0, true);
t('Inhalt: Inhaltsverzeichnis-Feld', c.indexOf('TOC \\o "1-3"') > 0, true);
t('Inhalt: zwei Seitenumbrüche', (c.match(/w:br w:type="page"/g) || []).length, 2);
const ids = (c.match(/<wp:docPr[^>]*\bid="(\d+)"/g) || []);
t('Inhalt: docPr-IDs eindeutig', new Set(ids).size, ids.length);
t('Nur Titelblatt: Umbruch dahinter', (S.mkContentPkg(['titel'], false).match(/w:br w:type="page"/g) || []).length, 1);
t('Nichts gewählt: kein Paket', S.mkContentPkg([], false), null);

/* ── Seitengeometrie ──────────────────────────────────────────────────────── */
const sample =
  '<w:document><w:body><w:p/><w:sectPr w:rsidR="00A">' +
  '<w:headerReference w:type="default" r:id="rId1"/>' +
  '<w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1417" w:right="1417" w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>' +
  '<w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>';
const geo = S.setPageGeometry(sample, false, false);
t('Geometrie: Ränder der Vorlage gesetzt', geo.indexOf('w:top="1701"') > 0, true);
t('Geometrie: genau ein pgSz', (geo.match(/<w:pgSz\b/g) || []).length, 1);
t('Geometrie: pgMar nach headerReference',
  geo.indexOf('<w:headerReference') < geo.indexOf('<w:pgMar'), true);
t('Geometrie: pgMar vor cols', geo.indexOf('<w:pgMar') < geo.indexOf('<w:cols'), true);
const quer = S.setPageGeometry(sample, true, false);
t('Geometrie: Querformat', quer.indexOf('w:orient="landscape"') > 0, true);
t('Geometrie: ohne sectPr wird eines angelegt',
  S.setPageGeometry('<w:document><w:body><w:p/></w:body></w:document>', false, false).indexOf('<w:sectPr>') > 0, true);

console.log(failed ? '\n' + failed + ' Test(s) fehlgeschlagen' : '\nalle Tests ok');
process.exitCode = failed ? 1 : 0;

/* ── Platzhalter in der Kopfzeile ─────────────────────────────────────────── */
const hdrLeer = S.mkHdrPkg(false, true, '', '');
t('Kopfzeile leer: Platzhalter Prozessname', hdrLeer.indexOf('Prozessname') > 0, true);
t('Kopfzeile leer: Platzhalter Ablageort', hdrLeer.indexOf('Ablageort') > 0, true);
t('Extern: auch ohne Platzhalter', S.mkHdrPkg(false, false, '', '').indexOf('Prozessname'), -1);
console.log(failed ? failed + ' Test(s) fehlgeschlagen' : 'Platzhalter ok');
process.exitCode = failed ? 1 : 0;
