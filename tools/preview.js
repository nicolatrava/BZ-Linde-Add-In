#!/usr/bin/env node
/*
 * Baut aus taskpane.html echte .docx-Dateien, ohne Word zu starten.
 *
 *   node tools/preview.js [Zielordner]
 *
 * Der Skriptteil von taskpane.html wird in einer Sandbox ausgeführt (Office,
 * Word und DOM sind minimal nachgebaut). Danach werden die OOXML-Bausteine
 * abgerufen, die das Add-In auch in Word einfügen würde, und zu vollständigen
 * Word-Dateien zusammengesetzt. So lässt sich jede Änderung an Kopfzeile,
 * Fusszeile oder Musterinhalt prüfen, bevor sie ausgerollt wird.
 */
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const zlib = require('zlib');

const ROOT = path.resolve(__dirname, '..');
const OUTDIR = path.resolve(process.argv[2] || path.join(ROOT, 'preview'));

/* ── taskpane.html in einer Sandbox ausführen ─────────────────────────────── */
function loadAddin() {
  const html = fs.readFileSync(path.join(ROOT, 'taskpane.html'), 'utf8');
  const m = html.match(/<script>\n([\s\S]*?)\n<\/script>/);
  if (!m) throw new Error('Inline-<script> in taskpane.html nicht gefunden');

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
    console,
    setTimeout,
    Promise,
    fetch: undefined,
    localStorage: { getItem() { return null; }, setItem() {}, removeItem() {} },
    document: {
      getElementById: el,
      querySelectorAll() { return []; },
      body: {}
    },
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
  vm.createContext(sandbox);
  vm.runInContext(m[1], sandbox, { filename: 'taskpane.html' });
  return { api: sandbox, field: el };
}

/* ── Minimaler ZIP-Writer (deflate, keine Abhängigkeiten) ─────────────────── */
function crc32(buf) {
  let c, table = crc32.table;
  if (!table) {
    table = crc32.table = new Int32Array(256);
    for (let n = 0; n < 256; n++) {
      c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      table[n] = c;
    }
  }
  c = -1;
  for (let i = 0; i < buf.length; i++) c = table[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ -1) >>> 0;
}

function zip(entries) {
  const locals = [], central = [];
  let offset = 0;
  for (const e of entries) {
    const name = Buffer.from(e.name, 'utf8');
    const raw = Buffer.isBuffer(e.data) ? e.data : Buffer.from(e.data, 'utf8');
    const data = zlib.deflateRawSync(raw, { level: 9 });
    const crc = crc32(raw);
    const lh = Buffer.alloc(30);
    lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(20, 4); lh.writeUInt16LE(0, 6);
    lh.writeUInt16LE(8, 8); lh.writeUInt16LE(0, 10); lh.writeUInt16LE(0x21, 12);
    lh.writeUInt32LE(crc, 14); lh.writeUInt32LE(data.length, 18);
    lh.writeUInt32LE(raw.length, 22); lh.writeUInt16LE(name.length, 26);
    lh.writeUInt16LE(0, 28);
    locals.push(lh, name, data);
    const ch = Buffer.alloc(46);
    ch.writeUInt32LE(0x02014b50, 0); ch.writeUInt16LE(20, 4); ch.writeUInt16LE(20, 6);
    ch.writeUInt16LE(0, 8); ch.writeUInt16LE(8, 10); ch.writeUInt16LE(0, 12);
    ch.writeUInt16LE(0x21, 14); ch.writeUInt32LE(crc, 16);
    ch.writeUInt32LE(data.length, 20); ch.writeUInt32LE(raw.length, 24);
    ch.writeUInt16LE(name.length, 28); ch.writeUInt32LE(0, 30);
    ch.writeUInt16LE(0, 34); ch.writeUInt32LE(0, 38); ch.writeUInt32LE(offset, 42);
    central.push(ch, name);
    offset += 30 + name.length + data.length;
  }
  const cd = Buffer.concat(central);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10); end.writeUInt32LE(cd.length, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat([Buffer.concat(locals), cd, end]);
}

/* ── Flat-OPC-Paket auseinandernehmen ─────────────────────────────────────── */
function bodyOf(pkg) {
  const m = pkg.match(/<w:body>([\s\S]*)<\/w:body>/);
  if (!m) throw new Error('kein <w:body> im Paket');
  return m[1];
}
function imagesOf(pkg) {
  const out = [];
  const re = /<pkg:part pkg:name="\/word\/media\/([^"]+)"[^>]*><pkg:binaryData>([\s\S]*?)<\/pkg:binaryData><\/pkg:part>/g;
  let m;
  while ((m = re.exec(pkg))) out.push({ name: m[1], data: Buffer.from(m[2], 'base64') });
  return out;
}
function relsOf(pkg) {
  const m = pkg.match(/<pkg:part pkg:name="\/word\/_rels\/document\.xml\.rels"[^>]*><pkg:xmlData>([\s\S]*?)<\/pkg:xmlData><\/pkg:part>/);
  if (!m) return [];
  const out = [];
  const re = /<Relationship Id="([^"]+)"[^>]*Target="media\/([^"]+)"\/>/g;
  let r;
  while ((r = re.exec(m[1]))) out.push({ id: r[1], target: r[2] });
  return out;
}

const NS = 'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:w10="urn:schemas-microsoft-com:office:word" ' +
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
  'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ' +
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ' +
  'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" ' +
  'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" ' +
  'mc:Ignorable="w14 w15 wp14"';
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/* ── .docx aus den Bausteinen zusammensetzen ──────────────────────────────── */
function makeDocx(api, { hdrPkg, ftrPkg, contentPkg, sectPr }) {
  const media = [];
  const rels = [
    '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
    '<Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>',
    '<Relationship Id="rIdHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
    '<Relationship Id="rIdFtr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
  ];
  const hdrRels = [], ftrRels = [];

  for (const img of imagesOf(hdrPkg)) media.push(img);
  for (const r of relsOf(hdrPkg)) hdrRels.push(r);
  for (const img of imagesOf(ftrPkg)) media.push(img);
  for (const r of relsOf(ftrPkg)) ftrRels.push(r);
  if (contentPkg) {
    for (const img of imagesOf(contentPkg)) media.push(img);
    for (const r of relsOf(contentPkg)) {
      rels.push('<Relationship Id="' + r.id + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/' + r.target + '"/>');
    }
  }

  const relXml = (list) => DECL +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    list.map((r) => typeof r === 'string' ? r :
      '<Relationship Id="' + r.id + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/' + r.target + '"/>'
    ).join('') + '</Relationships>';

  const docBody = (contentPkg ? bodyOf(contentPkg) : '<w:p/>') +
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdHdr"/>' +
    '<w:footerReference w:type="default" r:id="rIdFtr"/>' + sectPr +
    '<w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>';

  const entries = [
    { name: '[Content_Types].xml', data: DECL +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Default Extension="png" ContentType="image/png"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
      '</Types>' },
    { name: '_rels/.rels', data: DECL +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>' },
    { name: 'word/document.xml', data: DECL + '<w:document ' + NS + '><w:body>' + docBody + '</w:body></w:document>' },
    { name: 'word/_rels/document.xml.rels', data: relXml(rels) },
    { name: 'word/header1.xml', data: DECL + '<w:hdr ' + NS + '>' + bodyOf(hdrPkg) + '</w:hdr>' },
    { name: 'word/_rels/header1.xml.rels', data: relXml(hdrRels) },
    { name: 'word/footer1.xml', data: DECL + '<w:ftr ' + NS + '>' + bodyOf(ftrPkg) + '</w:ftr>' },
    { name: 'word/_rels/footer1.xml.rels', data: relXml(ftrRels) },
    { name: 'word/styles.xml', data: DECL + api.STYLES_XML },
    { name: 'word/numbering.xml', data: DECL + api.NUMBERING_XML }
  ];
  const seen = {};
  for (const img of media) {
    if (seen[img.name]) continue;
    seen[img.name] = true;
    entries.push({ name: 'word/media/' + img.name, data: img.data });
  }
  return zip(entries);
}

/* ── Schema-Prüfung ───────────────────────────────────────────────────────────
 * Word akzeptiert nur die im OOXML-Schema festgelegte Reihenfolge der
 * Kind-Elemente; bei einem Verstoss meldet es «Das Dokument ist beschädigt».
 * Geprüft werden die Container, die dieses Add-In selbst zusammenbaut.
 */
const ORDER = {
  'w:pPr': ['w:pStyle','w:keepNext','w:keepLines','w:pageBreakBefore','w:framePr',
    'w:widowControl','w:numPr','w:suppressLineNumbers','w:pBdr','w:shd','w:tabs',
    'w:suppressAutoHyphens','w:kinsoku','w:wordWrap','w:overflowPunct','w:topLinePunct',
    'w:autoSpaceDE','w:autoSpaceDN','w:bidi','w:adjustRightInd','w:snapToGrid','w:spacing',
    'w:ind','w:contextualSpacing','w:mirrorIndents','w:suppressOverlap','w:jc',
    'w:textDirection','w:textAlignment','w:textboxTightWrap','w:outlineLvl','w:divId',
    'w:cnfStyle','w:rPr','w:sectPr','w:pPrChange'],
  'w:rPr': ['w:rStyle','w:rFonts','w:b','w:bCs','w:i','w:iCs','w:caps','w:smallCaps',
    'w:strike','w:dstrike','w:outline','w:shadow','w:emboss','w:imprint','w:noProof',
    'w:snapToGrid','w:vanish','w:webHidden','w:color','w:spacing','w:w','w:kern',
    'w:position','w:sz','w:szCs','w:highlight','w:u','w:effect','w:bdr','w:shd',
    'w:fitText','w:vertAlign','w:rtl','w:cs','w:em','w:lang','w:eastAsianLayout',
    'w:specVanish','w:oMath'],
  'w:tblPr': ['w:tblStyle','w:tblpPr','w:tblOverlap','w:bidiVisual','w:tblStyleRowBandSize',
    'w:tblStyleColBandSize','w:tblW','w:jc','w:tblCellSpacing','w:tblInd','w:tblBorders',
    'w:shd','w:tblLayout','w:tblCellMar','w:tblLook','w:tblCaption','w:tblDescription'],
  'w:tcPr': ['w:cnfStyle','w:tcW','w:gridSpan','w:hMerge','w:vMerge','w:tcBorders','w:shd',
    'w:noWrap','w:tcMar','w:textDirection','w:tcFitText','w:vAlign','w:hideMark'],
  'w:trPr': ['w:cnfStyle','w:divId','w:gridBefore','w:gridAfter','w:wBefore','w:wAfter',
    'w:cantSplit','w:trHeight','w:tblHeader','w:tblCellSpacing','w:jc','w:hidden'],
  'w:tblBorders': ['w:top','w:start','w:left','w:bottom','w:end','w:right','w:insideH','w:insideV'],
  'w:sectPr': ['w:headerReference','w:footerReference','w:footnotePr','w:endnotePr','w:type',
    'w:pgSz','w:pgMar','w:paperSrc','w:pgBorders','w:lnNumType','w:pgNumType','w:cols',
    'w:formProt','w:vAlign','w:noEndnote','w:titlePg','w:textDirection','w:bidi',
    'w:rtlGutter','w:docGrid','w:printerSettings','w:sectPrChange'],
  'w:tbl': ['w:tblPr','w:tblGrid','w:tr'],
  'w:tr': ['w:tblPrEx','w:trPr','w:tc'],
  'w:tc': ['w:tcPr'],
  'w:p': ['w:pPr'],
  'w:r': ['w:rPr'],
  'wp:anchor': ['wp:simplePos','wp:positionH','wp:positionV','wp:extent','wp:effectExtent',
    'wp:wrapNone','wp:wrapSquare','wp:wrapTight','wp:wrapThrough','wp:wrapTopAndBottom',
    'wp:docPr','wp:cNvGraphicFramePr','a:graphic','wp14:sizeRelH','wp14:sizeRelV'],
  'wp:inline': ['wp:extent','wp:effectExtent','wp:docPr','wp:cNvGraphicFramePr','a:graphic'],
  'pic:pic': ['pic:nvPicPr','pic:blipFill','pic:spPr'],
  'pic:nvPicPr': ['pic:cNvPr','pic:cNvPicPr'],
  'a:blipFill': ['a:blip','a:srcRect','a:tile','a:stretch'],
  'pic:spPr': ['a:xfrm','a:custGeom','a:prstGeom','a:noFill','a:solidFill','a:gradFill',
    'a:blipFill','a:pattFill','a:grpFill','a:ln']
};

/* Sehr einfacher XML-Walker: liefert für jeden Knoten Tag + Kind-Tags. */
function checkOrder(xml, where, problems) {
  const stack = [];
  const re = /<(\/?)([A-Za-z0-9_]+:[A-Za-z0-9_]+)((?:"[^"]*"|'[^']*'|[^>"'])*?)(\/?)>/g;
  let m;
  while ((m = re.exec(xml))) {
    const closing = m[1] === '/';
    const tag = m[2];
    const selfClose = m[4] === '/';
    if (!closing) {
      if (stack.length) stack[stack.length - 1].kids.push(tag);
      if (!selfClose) stack.push({ tag, kids: [] });
    } else {
      const node = stack.pop();
      if (!node) continue;
      const order = ORDER[node.tag];
      if (order) {
        let last = -1, lastTag = '';
        for (const k of node.kids) {
          const idx = order.indexOf(k);
          if (idx === -1) continue;           /* unbekannte Kinder ignorieren */
          if (idx < last) {
            problems.push(where + ': <' + node.tag + '> – <' + k + '> steht nach <' + lastTag + '>');
          } else { last = idx; lastTag = k; }
        }
      }
    }
  }
}

/* ── Varianten erzeugen ───────────────────────────────────────────────────── */
function main() {
  const { api, field } = loadAddin();
  fs.mkdirSync(OUTDIR, { recursive: true });
  const org = api.ORG_DEFAULT;
  const cases = [];

  api._applyN = 1;
  cases.push({
    file: 'A4-Hoch_intern.docx',
    hdrPkg: api.mkHdrPkg(false, true, 'Sekretariat / C 3.4.2 Kreditoren', 'Ordner Kreditoren'),
    ftrPkg: api.mkFtrPkg(false, org, '31.07.26', 'NT', true),
    contentPkg: api.mkContentPkg(['muster'], false),
    sectPr: api.GEO.hoch.pgSz + api.PGMAR
  });

  api._applyN = 2;
  cases.push({
    file: 'A4-Hoch_intern_Titelblatt_IVZ_Muster.docx',
    hdrPkg: api.mkHdrPkg(false, true, 'Pflege und Betreuung / B 2.1 Eintritt', 'Q-Handbuch, Kapitel B'),
    ftrPkg: api.mkFtrPkg(false, org, '31.07.26', 'JM', true),
    contentPkg: api.mkContentPkg(['titel', 'ivz', 'muster'], false),
    sectPr: api.GEO.hoch.pgSz + api.PGMAR
  });

  api._applyN = 3;
  cases.push({
    file: 'A4-Quer_intern.docx',
    hdrPkg: api.mkHdrPkg(true, true, 'Hotellerie / D 1.2 Menüplanung', 'Ordner Hotellerie'),
    ftrPkg: api.mkFtrPkg(true, org, '', '', true),
    contentPkg: api.mkContentPkg(['muster'], true),
    sectPr: api.GEO.quer.pgSz + api.PGMAR
  });

  api._applyN = 4;
  cases.push({
    file: 'A4-Hoch_extern.docx',
    hdrPkg: api.mkHdrPkg(false, false, '', ''),
    ftrPkg: api.mkFtrPkg(false, org, '', '', true),
    contentPkg: api.mkContentPkg(['muster'], false),
    sectPr: api.GEO.hoch.pgSz + api.PGMAR
  });

  api._applyN = 5;
  const person = { name: 'Maria Muster', funktion: 'Leiterin Pflege',
                   telefon: '041 984 29 30', email: 'm.muster@bz-linde.ch' };
  cases.push({
    file: 'Brief.docx',
    hdrPkg: api.mkBriefHdrPkg(),
    ftrPkg: api.mkFtrPkg(false, org, '', '', false),
    contentPkg: api.mkBriefPkg(person, org, 'Ihre Anfrage vom 12. Juni'),
    sectPr: api.GEO.hoch.pgSz + api.PGMAR_BRIEF
  });

  const problems = [];
  for (const c of cases) {
    checkOrder(bodyOf(c.hdrPkg), c.file + ' Kopfzeile', problems);
    checkOrder(bodyOf(c.ftrPkg), c.file + ' Fusszeile', problems);
    if (c.contentPkg) checkOrder(bodyOf(c.contentPkg), c.file + ' Inhalt', problems);
    const buf = makeDocx(api, c);
    fs.writeFileSync(path.join(OUTDIR, c.file), buf);
    console.log(c.file.padEnd(44) + (buf.length / 1024).toFixed(0) + ' KB');
  }

  /* setPageGeometry gegen ein realistisches Body-OOXML prüfen: so gibt Word
     ein bestehendes Dokument mit Kopf-/Fusszeilenverweisen zurück. */
  const sample =
    '<w:document><w:body><w:p/><w:sectPr w:rsidR="00A">' +
    '<w:headerReference w:type="default" r:id="rId1"/>' +
    '<w:footerReference w:type="default" r:id="rId2"/>' +
    '<w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1417" w:right="1417" w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>' +
    '<w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>';
  const geo = api.setPageGeometry(sample, false, false);
  checkOrder(geo, 'setPageGeometry', problems);
  if (geo.indexOf('w:top="1701"') === -1) problems.push('setPageGeometry: neue Ränder fehlen');
  if ((geo.match(/<w:pgSz\b/g) || []).length !== 1) problems.push('setPageGeometry: pgSz nicht genau einmal');
  if ((geo.match(/<w:pgMar\b/g) || []).length !== 1) problems.push('setPageGeometry: pgMar nicht genau einmal');

  if (problems.length) {
    console.error('\nSchema-Reihenfolge verletzt:');
    problems.forEach((p) => console.error('  ' + p));
    process.exitCode = 1;
  } else {
    console.log('\nSchema-Reihenfolge: ok');
  }
  console.log('-> ' + OUTDIR);
}

main();
