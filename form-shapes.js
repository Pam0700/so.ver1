/*
 * form-shapes.js — Module độc lập: chép lại các ô chữ (text box) của form gốc vào file Excel đã điền.
 *
 * ExcelJS không giữ được hình vẽ/ô chữ (vd. ô chữ ký "Người báo cáo (QC) / Reported by",
 * "Kho & vận chuyển / Inspected by"), nên sau khi ghi file, module này mở gói .xlsx bằng JSZip,
 * lấy các ô chữ trong form gốc, dời xuống đúng vị trí của từng vùng (vùng 1, 2, 3...) rồi ghi vào file mới.
 *
 * CPFormShapes.inject(outBuffer, templateBuffer, sheetName, zoneMap) -> Promise<{ buffer, count }>
 *   zoneMap = [{ base, start, end, parts:[{a,z,d}] }]  (do template-fill.js tạo ra: vùng bắt đầu ở dòng `base` của
 *   file mới, lấy từ dòng start..end của form gốc, phần a..z của form gốc nằm ở base + d của file mới)
 * Cần JSZip (tự nạp từ cdnjs nếu chưa có).
 */
(function (root) {
  'use strict';

  const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const NS_PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const DRAWING_REL = NS_REL + '/drawing';
  const DRAWING_CT = 'application/vnd.openxmlformats-officedocument.drawing+xml';

  function loadJSZip() {
    if (root.JSZip) return Promise.resolve(root.JSZip);
    return new Promise(function (res, rej) {
      const s = document.createElement('script');
      s.src = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
      s.onload = function () { res(root.JSZip); };
      s.onerror = function () { rej(new Error('không tải được thư viện JSZip')); };
      document.head.appendChild(s);
    });
  }
  const unesc = function (s) { return String(s).replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&'); };
  const read = async function (z, path) { const f = z.file(path); return f ? f.async('string') : null; };
  const dirOf = function (p) { return p.slice(0, p.lastIndexOf('/') + 1); };
  function resolve(base, target) {          // đường dẫn tương đối trong gói
    if (target.charAt(0) === '/') return target.slice(1);
    const parts = (dirOf(base) + target).split('/'), out = [];
    parts.forEach(function (p) { if (p === '..') out.pop(); else if (p && p !== '.') out.push(p); });
    return out.join('/');
  }

  // tìm file drawing của trang tính có tên sheetName trong form gốc
  async function findTemplateDrawing(tz, sheetName) {
    const wbXml = await read(tz, 'xl/workbook.xml'), rels = await read(tz, 'xl/_rels/workbook.xml.rels');
    if (!wbXml || !rels) return null;
    let rid = null;
    (wbXml.match(/<sheet\b[^>]*>/g) || []).forEach(function (t) {
      const n = /\bname="([^"]*)"/.exec(t), r = /\br:id="([^"]*)"/.exec(t);
      if (n && r && unesc(n[1]) === sheetName) rid = r[1];
    });
    if (!rid) return null;
    let sheetPath = null;
    (rels.match(/<Relationship\b[^>]*>/g) || []).forEach(function (t) {
      const i = /\bId="([^"]*)"/.exec(t), tg = /\bTarget="([^"]*)"/.exec(t);
      if (i && tg && i[1] === rid) sheetPath = resolve('xl/workbook.xml', tg[1]);
    });
    if (!sheetPath) return null;
    const relPath = dirOf(sheetPath) + '_rels/' + sheetPath.slice(sheetPath.lastIndexOf('/') + 1) + '.rels';
    const sr = await read(tz, relPath);
    if (!sr) return null;
    let drawing = null;
    (sr.match(/<Relationship\b[^>]*>/g) || []).forEach(function (t) {
      if (/\bType="[^"]*\/drawing"/.test(t)) { const tg = /\bTarget="([^"]*)"/.exec(t); if (tg) drawing = resolve(sheetPath, tg[1]); }
    });
    return drawing;
  }

  function parseAnchors(xml) {
    const out = [];
    const re = /<xdr:(oneCellAnchor|twoCellAnchor|absoluteAnchor)\b[\s\S]*?<\/xdr:\1>/g;
    let m;
    while ((m = re.exec(xml))) {
      const a = m[0];
      if (/\br:(embed|id|link)=/.test(a)) continue;                 // hình ảnh/biểu đồ có quan hệ riêng: bỏ qua
      const f = /<xdr:from><xdr:col>(\d+)<\/xdr:col><xdr:colOff>(-?\d+)<\/xdr:colOff><xdr:row>(\d+)<\/xdr:row><xdr:rowOff>(-?\d+)<\/xdr:rowOff><\/xdr:from>/.exec(a);
      if (!f) continue;
      const ext = /<xdr:ext cx="(\d+)" cy="(\d+)"/.exec(a);
      out.push({ xml: a, row0: +f[3], key: [f[1], f[2], f[3], f[4], ext ? ext[1] + 'x' + ext[2] : ''].join('|') });
    }
    return out;
  }

  const shiftRows = function (a, delta) {
    return a.replace(/<xdr:row>(\d+)<\/xdr:row>/g, function (m, r) { return '<xdr:row>' + (+r + delta) + '</xdr:row>'; });
  };

  async function inject(outBuf, tplBuf, sheetName, zoneMap) {
    const JSZip = await loadJSZip();
    const tz = await JSZip.loadAsync(tplBuf);
    const dPath = await findTemplateDrawing(tz, sheetName);
    if (!dPath) return { buffer: outBuf, count: 0 };
    const dXml = await read(tz, dPath);
    if (!dXml) return { buffer: outBuf, count: 0 };
    const anchors = parseAnchors(dXml);

    // dựng drawing mới: mỗi vùng lấy các ô chữ nằm trong vùng của form gốc, dời xuống vị trí mới
    const made = []; let id = 1;
    (zoneMap || []).forEach(function (zone) {
      const seen = {};
      anchors.forEach(function (an) {
        const row1 = an.row0 + 1;
        if (row1 < zone.start || row1 > zone.end || seen[an.key]) return;
        const part = (zone.parts || []).find(function (p) { return row1 >= p.a && row1 <= p.z; });
        if (!part) return;
        seen[an.key] = true;
        const delta = (zone.base + part.d + (row1 - part.a)) - row1;
        id++;
        made.push(shiftRows(an.xml, delta)
          .replace(/<a:extLst>[\s\S]*?<\/a:extLst>/g, '')
          .replace(/<xdr:cNvPr id="\d+" name="[^"]*"/, '<xdr:cNvPr id="' + id + '" name="TextBox ' + (id - 1) + '"'));
      });
    });
    if (!made.length) return { buffer: outBuf, count: 0 };

    const z = await JSZip.loadAsync(outBuf);
    const sheetPath = Object.keys(z.files).filter(function (n) { return /^xl\/worksheets\/sheet\d+\.xml$/.test(n); }).sort()[0];
    if (!sheetPath) throw new Error('không tìm thấy trang tính trong file');
    let n = 1; while (z.file('xl/drawings/drawing' + n + '.xml')) n++;
    const newDrawing = 'xl/drawings/drawing' + n + '.xml';
    z.file(newDrawing, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      made.join('') + '</xdr:wsDr>', { createFolders: false });

    // quan hệ trang tính -> drawing
    const relPath = dirOf(sheetPath) + '_rels/' + sheetPath.slice(sheetPath.lastIndexOf('/') + 1) + '.rels';
    const rid = 'rIdShapes1', relEntry = '<Relationship Id="' + rid + '" Type="' + DRAWING_REL + '" Target="../drawings/drawing' + n + '.xml"/>';
    let rels = await read(z, relPath);
    rels = rels ? rels.replace('</Relationships>', relEntry + '</Relationships>')
                : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="' + NS_PKG + '">' + relEntry + '</Relationships>';
    z.file(relPath, rels, { createFolders: false });

    // kiểu nội dung
    let ct = await read(z, '[Content_Types].xml');
    if (ct.indexOf('/' + newDrawing) < 0) ct = ct.replace('</Types>', '<Override PartName="/' + newDrawing + '" ContentType="' + DRAWING_CT + '"/></Types>');
    z.file('[Content_Types].xml', ct, { createFolders: false });

    // thẻ <drawing> trong trang tính (đúng thứ tự: trước legacyDrawing / tableParts / extLst, sau rowBreaks)
    let sx = await read(z, sheetPath);
    if (!/<worksheet\b[^>]*\sxmlns:r=/.test(sx)) sx = sx.replace(/<worksheet\b/, '<worksheet xmlns:r="' + NS_REL + '"');
    const tag = '<drawing r:id="' + rid + '"/>';
    const m = /<(legacyDrawing|legacyDrawingHF|drawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b/.exec(sx);
    sx = m ? sx.slice(0, m.index) + tag + sx.slice(m.index) : sx.replace('</worksheet>', tag + '</worksheet>');
    z.file(sheetPath, sx, { createFolders: false });

    const buffer = await z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    return { buffer: buffer, count: made.length };
  }

  const api = { inject: inject, parseAnchors: parseAnchors };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.CPFormShapes = api;
})(typeof window !== 'undefined' ? window : globalThis);
