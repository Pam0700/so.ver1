/*
 * product-short.js — Module độc lập: rút gọn tên sản phẩm theo file quy ước tải lên.
 *
 * File quy ước (.xlsx), 2 cột:  Sản phẩm (gốc) | Sản phẩm rút gọn
 *   - Dòng nào để trống cột rút gọn thì giữ tên gốc.
 *   - Tên gốc được so khớp sau khi đã bỏ chữ VietGap, không phân biệt hoa/thường và khoảng trắng thừa.
 *   - Thứ tự sắp xếp (5kg, 2kg, 1kg, TF, 500g, Khay...) vẫn dựa trên tên GỐC,
 *     còn tên hiển thị trong kết quả là tên rút gọn. Các sản phẩm trùng tên sau khi rút gọn được gộp 1 lần.
 *
 * Cách gắn vào trang: thêm một thẻ div có id="productShortBox" và nạp file này bằng thẻ script (src).
 * Trang chính gọi CPProductShort.shorten(tên) và cung cấp window.CPGetProducts() (danh sách tên gốc).
 * Cần ExcelJS đã nạp sẵn.
 */
(function (root) {
  'use strict';

  const clean = function (s) { return String(s == null ? '' : s).normalize('NFC').replace(/\s+/g, ' ').trim(); };
  const keyOf = function (s) { return clean(s).toLowerCase(); };
  let map = new Map();          // khóa tên gốc -> { orig, short }

  function shorten(name) {
    const m = map.get(keyOf(name));
    return m && m.short ? m.short : name;
  }
  function size() { return map.size; }

  async function loadMap(file) {
    const wb = new root.ExcelJS.Workbook();
    await wb.xlsx.load(await file.arrayBuffer());
    const ws = wb.worksheets[0];
    const next = new Map();
    ws.eachRow(function (row, n) {
      if (n === 1) return;
      const orig = clean(row.getCell(1).text), short = clean(row.getCell(2).text);
      if (orig && short) next.set(keyOf(orig), { orig: orig, short: short });
    });
    map = next;
    return map.size;
  }

  async function buildTemplate(names) {
    // danh sách = tên trong file sản phẩm đã tải + các tên đã có trong quy ước (để không mất khi cập nhật)
    const all = new Map();
    names.forEach(function (n) { all.set(keyOf(n), clean(n)); });
    map.forEach(function (m, k) { if (!all.has(k)) all.set(k, m.orig); });
    const list = Array.from(all.entries()).sort(function (a, b) { return a[1].localeCompare(b[1], 'vi'); });

    const wb = new root.ExcelJS.Workbook();
    const ws = wb.addWorksheet('Quy ước sản phẩm');
    ws.columns = [{ width: 60 }, { width: 40 }];
    const h = ws.addRow(['Sản phẩm (gốc)', 'Sản phẩm rút gọn']);
    h.font = { bold: true };
    const thin = { style: 'thin' }, box = { top: thin, bottom: thin, left: thin, right: thin };
    h.eachCell(function (c) { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE7E6E6' } }; c.border = box; });
    list.forEach(function (e) {
      const m = map.get(e[0]);
      const r = ws.addRow([e[1], m ? m.short : '']);
      r.eachCell({ includeEmpty: true }, function (c, i) {
        c.border = box; c.alignment = { vertical: 'top', wrapText: true };
        if (i === 2) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF9DB' } };
      });
    });
    ws.views = [{ state: 'frozen', ySplit: 1 }];
    return { buffer: await wb.xlsx.writeBuffer(), count: list.length };
  }

  function initUI() {
    const box = document.getElementById('productShortBox');
    if (!box) return;
    box.innerHTML =
      '<fieldset><legend>Quy ước rút gọn tên sản phẩm (tuỳ chọn)</legend>' +
      '<label>File quy ước (2 cột: Sản phẩm gốc, Sản phẩm rút gọn): <input type="file" id="psFile" accept=".xlsx"></label>' +
      '<button id="psTpl" type="button">Tải file mẫu (danh sách sản phẩm từ file đã tải)</button>' +
      '<div id="psMsg" class="muted"></div></fieldset>';
    const msg = function (t, cls) { const m = box.querySelector('#psMsg'); m.textContent = t; m.className = cls || 'muted'; };

    box.querySelector('#psFile').addEventListener('change', async function (e) {
      const f = e.target.files[0];
      if (!f) { map = new Map(); return msg(''); }
      try { msg('Đã nạp ' + (await loadMap(f)) + ' dòng quy ước sản phẩm. Bấm "Xử lý" lại để áp dụng.', 'ok'); }
      catch (err) { msg('Không đọc được file quy ước: ' + err.message, 'err'); }
    });

    box.querySelector('#psTpl').onclick = async function () {
      const names = typeof root.CPGetProducts === 'function' ? root.CPGetProducts() : [];
      if (!names.length && !map.size) return msg('Hãy tải file sản phẩm (50, 54...) trước.', 'err');
      try {
        const r = await buildTemplate(names);
        const a = document.createElement('a');
        a.href = URL.createObjectURL(new Blob([r.buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
        a.download = 'quy_uoc_san_pham.xlsx'; a.click();
        setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
        msg('Đã tạo file mẫu với ' + r.count + ' sản phẩm. Điền cột "Sản phẩm rút gọn" rồi tải lên lại.', 'ok');
      } catch (err) { msg('Lỗi: ' + err.message, 'err'); }
    };
  }

  const api = { shorten: shorten, size: size, loadMap: loadMap, buildTemplate: buildTemplate, initUI: initUI };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.CPProductShort = api;
  if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', initUI);
    else initUI();
  }
})(typeof window !== 'undefined' ? window : globalThis);
