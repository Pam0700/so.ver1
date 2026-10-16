/*
 * dispatch.js — Module độc lập: "Điều xe" -> mỗi dòng 1 file form báo cáo đã điền, gom thành file zip.
 *
 * File điều xe (.xlsx): 2 cột có tiêu đề "STT" và "Ship".
 *   - Ship là 3 hoặc 4 số cuối của mã ship (vd. 158 ứng với 24SPF47158). Ô "458+544" nghĩa là 2 ship.
 *   - Ô Ship trống thì bỏ qua.
 * Mỗi dòng có ship: lọc dữ liệu của (các) ship đó, điền vào form báo cáo, đặt tên file theo STT (vd. 10.xlsx),
 * ghi STT vào ô J4, J43... mỗi vùng, đặt sẵn vùng in A4. Sau đó nén tất cả thành 1 file zip.
 *
 * Cách gắn vào trang: thêm một thẻ div có id="dispatchBox" và nạp file này bằng thẻ script (src).
 * Cần các móc của trang chính (CPGetDgRows, CPBuildGroups, CPHasProducts), template-fill.js, ExcelJS, JSZip.
 */
(function (root) {
  'use strict';

  const T = function (s) { return String(s == null ? '' : s).trim(); };
  const plain = function (s) {   // bỏ dấu, chữ thường: để nhận tiêu đề "STT", "Ship"
    return T(s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase();
  };

  /* ---------- đọc bảng điều xe (rows = mảng các mảng chuỗi) ---------- */
  function parseTable(rows) {
    let hi = -1, cStt = -1, cShip = -1;
    for (let i = 0; i < Math.min(rows.length, 15) && hi < 0; i++) {
      const low = rows[i].map(plain);
      const a = low.indexOf('stt'), b = low.indexOf('ship');
      if (a >= 0 && b >= 0) { hi = i; cStt = a; cShip = b; }
    }
    if (hi < 0) throw new Error('Không thấy tiêu đề "STT" và "Ship" trong file điều xe.');
    const out = [];
    for (let i = hi + 1; i < rows.length; i++) {
      const stt = T(rows[i][cStt]), ship = T(rows[i][cShip]);
      if (!stt && !ship) continue;
      out.push({ line: i + 1, stt: stt, ship: ship });
    }
    return out;
  }

  // "458+544" -> ["458","544"]
  function splitShips(text) {
    return T(text).split(/[+,;\/&]|\s+và\s+/i).map(function (s) { return s.replace(/\.0+$/, '').trim(); }).filter(Boolean);
  }

  /* ---------- khớp ship ---------- */
  // allShips: danh sách mã ship đầy đủ trong dg1. Số 3-4 chữ số khớp với các chữ số cuối của mã ship.
  function matchOne(token, allShips) {
    const t = token.toUpperCase();
    const exact = allShips.filter(function (s) { return s.toUpperCase() === t; });
    if (exact.length) return exact;
    return allShips.filter(function (s) { return s.toUpperCase().slice(-t.length) === t; });
  }

  // status: ok | partial | none | skip | nostt
  function planRow(item, allShips) {
    const tokens = splitShips(item.ship);
    const res = { line: item.line, stt: item.stt, ship: item.ship, tokens: tokens, found: [], missing: [], ambiguous: [], status: 'ok' };
    if (!tokens.length) { res.status = 'skip'; return res; }
    if (!item.stt) { res.status = 'nostt'; return res; }
    tokens.forEach(function (t) {
      const c = matchOne(t, allShips);
      if (c.length === 1) { if (res.found.indexOf(c[0]) < 0) res.found.push(c[0]); }
      else if (c.length === 0) res.missing.push(t);
      else res.ambiguous.push(t + ' (' + c.join(', ') + ')');
    });
    if (!res.found.length) res.status = 'none';
    else if (res.missing.length || res.ambiguous.length) res.status = 'partial';
    return res;
  }

  function planAll(items, dgRows) {
    const seen = {}, allShips = [];
    dgRows.forEach(function (r) { if (r.ship && !seen[r.ship]) { seen[r.ship] = 1; allShips.push(r.ship); } });
    return items.map(function (it) { return planRow(it, allShips); });
  }

  const safeName = function (s) { return T(s).replace(/[\\/:*?"<>|]/g, '').replace(/\s+/g, '_') || 'file'; };

  /* ---------- giao diện ---------- */
  function loadJSZip() {
    if (root.JSZip) return Promise.resolve(root.JSZip);
    return new Promise(function (res, rej) {
      const s = document.createElement('script');
      s.src = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
      s.onload = function () { res(root.JSZip); };
      s.onerror = function () { rej(new Error('không tải được thư viện nén zip')); };
      document.head.appendChild(s);
    });
  }
  const esc = function (s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;'); };

  function initUI() {
    const box = document.getElementById('dispatchBox');
    if (!box) return;
    box.innerHTML =
      '<fieldset><legend>Điều xe → mỗi dòng một form báo cáo, gom thành file zip</legend>' +
      '<label>File điều xe (.xlsx, 2 cột STT và Ship): <input type="file" id="dxFile" accept=".xlsx"></label>' +
      '<label>Form báo cáo (bỏ trống để dùng form lưu sẵn trên hệ thống, nếu có): <input type="file" id="dxForm" accept=".xlsx"></label>' +
      '<label><input type="checkbox" id="dxTitle" checked> Ghi tên file (STT) vào ô J4, J43... mỗi vùng có dữ liệu</label>' +
      '<label><input type="checkbox" id="dxPrint" checked> Đặt sẵn vùng in: A4, mỗi vùng một mặt giấy</label>' +
      '<label>Tên file zip: <input type="text" id="dxZip" value="dieu_xe" style="width:240px"> .zip</label>' +
      '<button id="dxCheck" type="button">Kiểm tra lại</button>' +
      '<button id="dxGo" type="button" class="primary" disabled>Tạo file &amp; tải zip</button>' +
      '<div id="dxMsg" class="muted" style="margin-top:6px"></div>' +
      '<div id="dxOut" style="overflow-x:auto"></div></fieldset>';
    const $q = function (id) { return box.querySelector(id); };
    const msg = function (html) { $q('#dxMsg').innerHTML = html; };
    let items = [], plan = [];

    async function readTable(file) {
      const wb = new root.ExcelJS.Workbook();
      await wb.xlsx.load(await file.arrayBuffer());
      const ws = wb.worksheets.find(function (w) {
        let ok = false;
        w.eachRow({ includeEmpty: false }, function (row, n) {
          if (n > 15 || ok) return;
          const low = []; for (let c = 1; c <= 10; c++) low.push(plain(row.getCell(c).text));
          if (low.indexOf('stt') >= 0 && low.indexOf('ship') >= 0) ok = true;
        });
        return ok;
      }) || wb.worksheets[0];
      const rows = [];
      ws.eachRow({ includeEmpty: true }, function (row, n) {
        const r = []; for (let c = 1; c <= 10; c++) r.push(row.getCell(c).text);
        while (rows.length < n - 1) rows.push([]);
        rows.push(r);
      });
      return parseTable(rows);
    }

    function refresh() {
      if (!items.length) return;
      const dg = root.CPGetDgRows ? root.CPGetDgRows() : [];
      if (!dg.length) { $q('#dxOut').innerHTML = ''; $q('#dxGo').disabled = true; return msg('<span class="err">Chưa tải file giao hàng (dg1) ở bước 1.</span>'); }
      plan = planAll(items, dg);
      const order = (document.getElementById('order') || {}).value || '';
      let html = '<table><tr><th>STT</th><th>Ship</th><th>Ship tìm được</th><th>Số khách</th><th>Trạng thái</th></tr>';
      let nOk = 0, nSkip = 0, nBad = 0;
      plan.forEach(function (p) {
        if (p.status === 'skip') { nSkip++; return; }          // ô Ship trống: bỏ qua, không liệt kê
        let customers = '', st = '', cls = '';
        if (p.found.length) {
          const rows = dg.filter(function (r) { return p.found.indexOf(r.ship) >= 0; });
          const g = root.CPBuildGroups ? root.CPBuildGroups(rows, false, order) : { list: [] };
          p.nCus = g.list.length; customers = String(g.list.length);
          p.rows = rows;
        }
        if (p.status === 'ok') { st = 'Sẽ tạo ' + safeName(p.stt) + '.xlsx'; cls = 'ok'; nOk++; }
        else if (p.status === 'partial') { st = 'Sẽ tạo, thiếu: ' + p.missing.concat(p.ambiguous).join(', '); cls = 'err'; nOk++; nBad++; }
        else if (p.status === 'none') { st = 'Không tìm thấy ship ' + p.missing.concat(p.ambiguous).join(', '); cls = 'err'; nBad++; }
        else if (p.status === 'nostt') { st = 'Thiếu STT'; cls = 'err'; nBad++; }
        html += '<tr><td>' + esc(p.stt) + '</td><td>' + esc(p.ship) + '</td><td>' + esc(p.found.join(' + ')) + '</td><td>' + customers +
                '</td><td class="' + cls + '">' + esc(st) + '</td></tr>';
      });
      $q('#dxOut').innerHTML = html + '</table>';
      $q('#dxGo').disabled = nOk === 0;
      msg('<span class="ok">' + nOk + ' file sẽ được tạo</span>, ' + nSkip + ' dòng Ship trống bỏ qua' +
          (nBad ? ', <span class="err">' + nBad + ' dòng có vấn đề (xem bảng)</span>' : '') + '.');
    }

    $q('#dxFile').addEventListener('change', async function (e) {
      const f = e.target.files[0]; items = []; plan = [];
      $q('#dxOut').innerHTML = ''; $q('#dxGo').disabled = true;
      if (!f) return msg('');
      try { msg('Đang đọc...'); items = await readTable(f); refresh(); }
      catch (err) { msg('<span class="err">' + esc(err.message) + '</span>'); }
    });
    $q('#dxCheck').addEventListener('click', refresh);

    async function formBuffer() {
      const f = $q('#dxForm').files[0];
      if (f) return f.arrayBuffer();
      if (typeof root.CPFormProvider === 'function') { const b = await root.CPFormProvider(); if (b) return b; }
      const tf = document.getElementById('tfFile');
      if (tf && tf.files[0]) return tf.files[0].arrayBuffer();
      throw new Error('Chưa có form báo cáo: chọn file form ở ô trên (hoặc nhờ admin lưu sẵn form trên hệ thống).');
    }

    $q('#dxGo').addEventListener('click', async function () {
      const go = $q('#dxGo'); go.disabled = true;
      try {
        if (!root.CPTemplateFill || !root.CPTemplateFill.renderFile) throw new Error('thiếu module template-fill.js');
        if (root.CPHasProducts && !root.CPHasProducts()) throw new Error('chưa tải file sản phẩm ở bước 1');
        refresh();
        const todo = plan.filter(function (p) { return p.found.length && p.rows; });
        if (!todo.length) throw new Error('không có dòng nào để tạo file');
        const order = (document.getElementById('order') || {}).value || '';
        msg('Đang nạp form...');
        const formBuf = await formBuffer();
        const wb = new root.ExcelJS.Workbook();
        await wb.xlsx.load(formBuf.slice(0));
        const JSZip = await loadJSZip(), zip = new JSZip();
        const used = {}, notes = []; let made = 0, shapesLost = 0;
        for (let i = 0; i < todo.length; i++) {
          const p = todo[i];
          msg('Đang tạo file ' + (i + 1) + ' / ' + todo.length + ' (STT ' + esc(p.stt) + ')...');
          await new Promise(function (r) { setTimeout(r, 0); });   // nhường giao diện cập nhật
          const g = root.CPBuildGroups(p.rows, false, order);
          if (!g.list.length) { notes.push('STT ' + p.stt + ': không có dữ liệu sản phẩm, bỏ qua'); continue; }
          const r = await root.CPTemplateFill.renderFile(wb, formBuf, g.list,
            { title: $q('#dxTitle').checked ? String(p.stt) : '', print: $q('#dxPrint').checked });
          let name = safeName(p.stt); if (used[name]) { used[name]++; name += '_' + used[name]; } else used[name] = 1;
          zip.file(name + '.xlsx', r.buffer, { createFolders: false });
          if (r.shapeError) shapesLost++;
          made++;
        }
        if (!made) throw new Error('không tạo được file nào');
        msg('Đang nén zip...');
        const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
        const zname = safeName($q('#dxZip').value.replace(/\.zip$/i, '')) + '.zip';
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob); a.download = zname; a.click();
        setTimeout(function () { URL.revokeObjectURL(a.href); }, 1500);
        msg('<span class="ok">Đã tạo ' + made + ' file, nén vào ' + esc(zname) + '.</span>' +
            (shapesLost ? ' <span class="err">' + shapesLost + ' file chưa giữ được ô chữ ký.</span>' : '') +
            (notes.length ? '<br><span class="err">' + notes.map(esc).join('<br>') + '</span>' : ''));
      } catch (err) { msg('<span class="err">Lỗi: ' + esc(err.message) + '</span>'); }
      go.disabled = false;
    });
  }

  const api = { parseTable: parseTable, splitShips: splitShips, matchOne: matchOne, planRow: planRow, planAll: planAll, initUI: initUI };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.CPDispatch = api;
  if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', initUI);
    else initUI();
  }
})(typeof window !== 'undefined' ? window : globalThis);
