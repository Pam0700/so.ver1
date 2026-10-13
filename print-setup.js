/*
 * print-setup.js — Module độc lập: đặt sẵn vùng in cho file Excel tải xuống (mở ra là in được ngay).
 *
 * Quy tắc: giấy A4, mỗi "vùng" là 1 mặt giấy. Module tự chọn hướng giấy (ngang/dọc) và tỉ lệ in (%)
 * sao cho cả bề rộng lẫn vùng cao nhất đều nằm gọn trong 1 trang, đặt ngắt trang sau từng vùng.
 * (Không dùng "Fit to page" vì Excel bỏ qua ngắt trang thủ công khi dùng chế độ đó.)
 *
 * CPPrint.apply(ws, {
 *   widths:      [độ rộng từng cột in, đơn vị ký tự Excel],
 *   zoneHeights: [chiều cao từng vùng, đơn vị pt],       // gồm cả dòng tiêu đề lặp lại nếu có
 *   margins:     { left, right, top, bottom, header, footer }  (inch),
 *   orientation: 'auto' | 'portrait' | 'landscape',
 *   breakRows:   [số dòng cuối của vùng cần ngắt trang],   // bỏ trống nếu đã tự đặt ngắt trang
 *   printTitlesRow: '1:1', footer: true
 * })  ->  { orientation, scale, fits }
 *
 * Dùng được cho bất kỳ worksheet ExcelJS nào; trang chính và template-fill.js gọi nếu module có mặt.
 */
(function (root) {
  'use strict';

  const A4 = { w: 595.28, h: 841.89 };   // pt
  const SAFETY = 0.98;                   // chừa 2% để không tràn trang do làm tròn

  // độ rộng cột (ký tự Excel, phông mặc định Calibri 11) -> pt
  const colPt = function (w) { return Math.round((w || 8.43) * 7 + 5) * 0.75; };

  function plan(o) {
    const W = o.widths.reduce(function (s, w) { return s + colPt(w); }, 0);
    const H = Math.max.apply(null, o.zoneHeights.concat([1]));
    const m = Object.assign({ left: 0.25, right: 0.25, top: 0.25, bottom: 0.25 }, o.margins || {});
    const mx = (m.left + m.right) * 72, my = (m.top + m.bottom) * 72;
    const list = ['portrait', 'landscape'].filter(function (x) { return !o.orientation || o.orientation === 'auto' || o.orientation === x; });
    let best = null;
    list.forEach(function (ori) {
      const pw = (ori === 'portrait' ? A4.w : A4.h) - mx, ph = (ori === 'portrait' ? A4.h : A4.w) - my;
      const r = Math.min(pw / W, ph / H);   // >= 1 nghĩa là vừa trang ở tỉ lệ 100%
      if (!best || Math.min(r, 1) > Math.min(best.r, 1)) best = { orientation: ori, r: r };
    });
    const minScale = o.minScale || 30;
    const scale = best.r * SAFETY >= 1 ? 100 : Math.max(minScale, Math.floor(best.r * SAFETY * 100));   // còn dư >= 2% thì in 100%
    return { orientation: best.orientation, scale: scale, fits: best.r * SAFETY * 100 >= minScale, width: W, height: H };
  }

  function apply(ws, o) {
    const p = plan(o);
    const ps = ws.pageSetup;
    ps.paperSize = 9;                 // A4
    ps.orientation = p.orientation;
    ps.scale = p.scale;
    ps.fitToPage = false;
    ps.horizontalCentered = true;
    if (o.margins) ps.margins = Object.assign({}, ps.margins || {}, o.margins);
    if (o.printTitlesRow) ps.printTitlesRow = o.printTitlesRow;
    if (o.footer) ws.headerFooter = Object.assign({}, ws.headerFooter || {}, { oddFooter: '&CTrang &P / &N' });
    (o.breakRows || []).forEach(function (r) { ws.getRow(r).addPageBreak(); });
    return p;
  }

  const api = { plan: plan, apply: apply, colPt: colPt };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.CPPrint = api;
})(typeof window !== 'undefined' ? window : globalThis);
