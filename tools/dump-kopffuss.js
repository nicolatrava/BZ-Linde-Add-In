#!/usr/bin/env node
/*
 * Leitet die moderne Kopf- und Fusszeile aus taskpane.html ab und legt sie
 * als Bausteine fuer tools/modernisiere-qms.py ab.
 *
 *   node tools/dump-kopffuss.js <zielordner>
 *
 * So bleiben Add-In (Zweig "Nur Kopf- und Fusszeile") und die modernisierte
 * QMS-Vorlage dieselbe Quelle - aendert sich das Layout im Add-In, genuegt
 * dump + modernisiere, und beide sehen wieder gleich aus.
 */
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.resolve(__dirname, '..');
const OUT = process.argv[2];
if (!OUT) { console.error('Zielordner fehlt'); process.exit(1); }
fs.mkdirSync(OUT, { recursive: true });

const m = fs.readFileSync(path.join(ROOT, 'taskpane.html'), 'utf8')
  .match(/<script>\n([\s\S]*?)\n<\/script>/);
const f = {};
const el = (i) => f[i] || (f[i] = { value: '', checked: false, className: '',
  textContent: '', style: {}, disabled: false, innerHTML: '',
  classList: { add() {}, remove() {} } });
const sb = {
  console, setTimeout, Promise,
  localStorage: { getItem: () => null, setItem() {}, removeItem() {} },
  document: { getElementById: el, querySelectorAll: () => [], body: {} },
  Office: { HostType: { Word: 'Word' }, onReady() {}, context: {
    requirements: { isSetSupported: () => false },
    document: { url: '', settings: { get: () => null, set() {}, saveAsync() {},
      refreshAsync() {} }, getFilePropertiesAsync(c) { c({ value: { url: '' } }); } } } },
  Word: { InsertLocation: { start: 'Start' }, run() {} }
};
sb.window = sb;
vm.createContext(sb);
vm.runInContext(m[1], sb, { filename: 'taskpane.html' });

sb._applyN = 1;
/* Kopfzeile: rechts oben Bereich/Nummer fett, darunter die Dokumentart -
   die Platzhalter ersetzt das Add-In nach dem Einsetzen. */
const hdr = sb.mkHdrPkg(false, true, '{{BEREICH}} / {{NUMMER}}', '{{DOKUMENTART}}');
const ftr = sb.mkFtrPkg(false, sb.ORG_DEFAULT, '', '', true);

const body = (p) => p.match(/<w:body>([\s\S]*)<\/w:body>/)[1];
const img = hdr.match(/<pkg:part pkg:name="\/word\/media\/[^"]+"[^>]*><pkg:binaryData>([\s\S]*?)<\/pkg:binaryData>/);
const rel = hdr.match(/<Relationship Id="(rIdLogo\d+)"[^>]*Target="media\/([^"]+)"\/>/);

fs.writeFileSync(path.join(OUT, 'hdr-body.xml'), body(hdr));
fs.writeFileSync(path.join(OUT, 'ftr-body.xml'), body(ftr));
fs.writeFileSync(path.join(OUT, 'logo.b64'), img[1]);
fs.writeFileSync(path.join(OUT, 'meta.json'),
  JSON.stringify({ logoRid: rel[1], logoFn: rel[2], docNs: sb.DOC_NS }));
console.log('Kopf- und Fusszeile abgelegt in ' + OUT);
