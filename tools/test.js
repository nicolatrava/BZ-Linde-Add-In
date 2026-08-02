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

/* ── Bedienlogik: jeder Umschalter muss fehlerfrei durchlaufen ────────────── */
['hoch', 'quer'].forEach(function (f) {
  ['intern', 'extern', 'brief'].forEach(function (ty) {
    let err = null;
    try { S.setFmt(f); S.setTyp(ty); } catch (e) { err = e.message; }
    t('Umschalten ' + f + '/' + ty, err, null);
  });
});
t('Brief erzwingt A4 Hoch', (S.setFmt('quer'), S.setTyp('brief'), S.fmt), 'hoch');
S.setFmt('hoch'); S.setTyp('intern');

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
t('Brief-Kopfzeile: nur Logo, keine Tabelle', S.mkBriefHdrPkg().indexOf('<w:tbl>'), -1);

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

/* ── Platzhalter in der Kopfzeile ─────────────────────────────────────────── */
const hdrLeer = S.mkHdrPkg(false, true, '', '');
t('Kopfzeile leer: Platzhalter Prozessname', hdrLeer.indexOf('Prozessname') > 0, true);
t('Kopfzeile leer: Platzhalter Ablageort', hdrLeer.indexOf('Ablageort') > 0, true);
t('Extern: auch ohne Platzhalter', S.mkHdrPkg(false, false, '', '').indexOf('Prozessname'), -1);
console.log(failed ? '\n' + failed + ' Test(s) fehlgeschlagen' : '\nalle Tests ok');
process.exitCode = failed ? 1 : 0;
