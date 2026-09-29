/* Lucas online catalogues — builds the Request-for-Price Excel file in the browser.
 * Loads ExcelJS (self-hosted) only when needed.
 *   LucasRFQExcel.build(request) -> Promise<Blob>
 *   request = { ref, date (Date), name, company, email, phone, country, type, message, items: [cart items] }
 */
(function () {
  'use strict';
  var BASE = (function () { var s = document.currentScript; return s ? s.src.replace(/[^/]*$/, '') : '/assets/'; })();
  var GREEN = 'FF00954C', GREEN_10 = 'FFE6F4ED', GREY = 'FFEDEDED', INPUT = 'FFFFF8DC', BLACK = 'FF231F20', MUTED = 'FF6D6E71';
  var FONT = 'Arial';

  var libPromise = null;
  function loadLib() {
    if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
    if (!libPromise) {
      libPromise = new Promise(function (resolve, reject) {
        var s = document.createElement('script');
        s.src = BASE + 'vendor-exceljs.min.js';
        s.onload = function () { window.ExcelJS ? resolve(window.ExcelJS) : reject(new Error('Excel library did not load')); };
        s.onerror = function () { libPromise = null; reject(new Error('Excel library did not load')); };
        document.head.appendChild(s);
      });
    }
    return libPromise;
  }
  function loadLogo() {
    return fetch(BASE + 'lucas-logo.png').then(function (r) { if (!r.ok) throw 0; return r.arrayBuffer(); })
      .then(function (buf) {
        var b = new Uint8Array(buf), s = '';
        for (var i = 0; i < b.length; i += 0x8000) s += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000));
        return btoa(s);
      }).catch(function () { return null; });
  }

  var SKIP = /^(lucas\s*no\.?|#)$/i;
  var NUMERIC = /\((mm|n|kg|cm|l|v|w|bar)\)$/i;
  function fallbackDetails(it) { return [['Make', it.make || ''], ['Application', it.app || ''], ['OE No.', it.oe || '']]; }
  function detailsOf(it) { return (it.details && it.details.length ? it.details : fallbackDetails(it)).filter(function (d) { return d && d[0] && !SKIP.test(d[0]); }); }
  function asValue(label, v) {
    v = v == null ? '' : String(v).trim();
    if (NUMERIC.test(label) && /^-?\d+(\.\d+)?$/.test(v)) return Number(v);
    return v;
  }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function fmtDate(d) { return pad(d.getDate()) + '/' + pad(d.getMonth() + 1) + '/' + d.getFullYear() + ' ' + pad(d.getHours()) + ':' + pad(d.getMinutes()); }

  function build(req) {
    return Promise.all([loadLib(), loadLogo()]).then(function (res) {
      var ExcelJS = res[0], logo = res[1];
      var items = req.items || [];
      var date = req.date || new Date();

      // union of detail columns, in first-seen order
      var detailCols = [], seen = {};
      items.forEach(function (it) { detailsOf(it).forEach(function (d) { if (!seen[d[0]]) { seen[d[0]] = true; detailCols.push(d[0]); } }); });

      var cols = [{ h: 'No.', w: 6 }, { h: 'Catalogue', w: 14 }, { h: 'Lucas No.', w: 14 }]
        .concat(detailCols.map(function (h) {
          var w = /applic/i.test(h) ? 42 : /cross/i.test(h) ? 34 : /^oe/i.test(h) ? 16 : /make/i.test(h) ? 16 : /fit/i.test(h) ? 14 : NUMERIC.test(h) ? 11 : 16;
          return { h: h, w: w, detail: true };
        }))
        .concat([{ h: 'Qty', w: 8 }, { h: 'Unit Price', w: 13 }, { h: 'Line Total', w: 14 }, { h: 'Availability / Lead Time', w: 22 }, { h: 'Remarks', w: 26 }]);
      var N = cols.length;
      var cQty = N - 4, cPrice = N - 3, cTotal = N - 2; // 1-based column numbers
      function L(c) { var s = ''; while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = (c - m - 1) / 26; } return s; }

      var wb = new ExcelJS.Workbook();
      wb.creator = 'Lucas online catalogue'; wb.created = date;
      var ws = wb.addWorksheet('Request for Price', {
        views: [{ showGridLines: false }],
        pageSetup: { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 } },
        headerFooter: { oddFooter: '&L' + (req.ref || '') + '&RPage &P of &N' }
      });
      ws.columns = cols.map(function (c) { return { width: c.w }; });

      var thin = { style: 'thin', color: { argb: 'FFD0D0D0' } };
      var box = { top: thin, left: thin, bottom: thin, right: thin };
      function fill(argb) { return { type: 'pattern', pattern: 'solid', fgColor: { argb: argb } }; }

      // Row 1: green brand bar with logo
      ws.getRow(1).height = 42;
      for (var c = 1; c <= N; c++) ws.getCell(1, c).fill = fill(GREEN);
      if (logo) {
        var id = wb.addImage({ base64: logo, extension: 'png' });
        ws.addImage(id, { tl: { col: 0, row: 0 }, ext: { width: 208, height: 42 }, editAs: 'absolute' });
      }
      ws.mergeCells(1, Math.max(4, N - 4), 1, N);
      var t = ws.getCell(1, Math.max(4, N - 4));
      t.value = 'REQUEST FOR PRICE'; t.font = { name: FONT, size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
      t.alignment = { horizontal: 'right', vertical: 'middle' };

      // Customer / request details
      var info = [
        ['Reference', req.ref], ['Date', fmtDate(date)], ['Name', req.name], ['Company', req.company],
        ['Email', req.email], ['Phone / WhatsApp', req.phone], ['Country', req.country], ['Customer type', req.type],
        ['Message', req.message]
      ];
      var r = 3;
      info.forEach(function (row) {
        ws.mergeCells(r, 1, r, 2); ws.mergeCells(r, 3, r, Math.min(N, 8));
        var a = ws.getCell(r, 1), b = ws.getCell(r, 3);
        a.value = row[0]; a.font = { name: FONT, size: 10, bold: true, color: { argb: MUTED } }; a.fill = fill(GREY);
        a.alignment = { vertical: 'top' }; a.border = box;
        b.value = row[1] || '–'; b.font = { name: FONT, size: 10, bold: row[0] === 'Reference', color: { argb: BLACK } };
        b.alignment = { vertical: 'top', wrapText: true }; b.border = box;
        if (row[0] === 'Email' && row[1]) { b.value = { text: row[1], hyperlink: 'mailto:' + row[1] }; b.font = { name: FONT, size: 10, color: { argb: GREEN }, underline: true }; }
        if (row[0] === 'Message' && row[1] && row[1].length > 90) ws.getRow(r).height = Math.min(90, 15 * Math.ceil(row[1].length / 90));
        r++;
      });

      // Parts table
      r += 1;
      ws.mergeCells(r, 1, r, 4);
      var sec = ws.getCell(r, 1);
      sec.value = 'Parts requested (' + items.length + ' line' + (items.length === 1 ? '' : 's') + ')';
      sec.font = { name: FONT, size: 12, bold: true, color: { argb: GREEN } };
      r += 1;
      var headRow = r;
      var hr = ws.getRow(r); hr.height = 30;
      cols.forEach(function (col, i) {
        var cell = hr.getCell(i + 1);
        cell.value = col.h;
        cell.font = { name: FONT, size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = fill(GREEN);
        cell.alignment = { vertical: 'middle', horizontal: (i + 1 >= cQty && i + 1 <= cTotal) || NUMERIC.test(col.h) || i === 0 ? 'center' : 'left', wrapText: true };
        cell.border = box;
      });
      r++;
      var first = r;
      items.forEach(function (it, idx) {
        var map = {}; detailsOf(it).forEach(function (d) { map[d[0]] = d[1]; });
        var vals = [idx + 1, it.catalogue || '', it.no]
          .concat(detailCols.map(function (h) { return asValue(h, map[h]); }))
          .concat([Number(it.qty) || 1, null, null, null, null]);
        var row = ws.getRow(r);
        vals.forEach(function (v, i) {
          var cell = row.getCell(i + 1);
          cell.value = v;
          cell.font = { name: FONT, size: 10, color: { argb: BLACK }, bold: i === 2 };
          cell.alignment = { vertical: 'top', wrapText: true, horizontal: i === 0 || typeof v === 'number' ? 'center' : 'left' };
          cell.border = box;
          if (idx % 2 === 1) cell.fill = fill('FFF7F7F7');
        });
        row.getCell(3).font = { name: FONT, size: 10, bold: true, color: { argb: GREEN } };
        [cPrice, cTotal + 1, cTotal + 2].forEach(function (cc) { row.getCell(cc).fill = fill(INPUT); });
        row.getCell(cPrice).numFmt = '#,##0.00';
        row.getCell(cTotal).value = { formula: 'IF(' + L(cPrice) + r + '="","",' + L(cQty) + r + '*' + L(cPrice) + r + ')' };
        row.getCell(cTotal).numFmt = '#,##0.00';
        row.getCell(cTotal).alignment = { vertical: 'top', horizontal: 'right' };
        r++;
      });
      var last = r - 1;

      // Totals
      var tr = ws.getRow(r); tr.height = 20;
      ws.mergeCells(r, 1, r, cQty - 1);
      var tl = tr.getCell(1); tl.value = 'Total'; tl.alignment = { horizontal: 'right', vertical: 'middle' };
      tl.font = { name: FONT, size: 10, bold: true, color: { argb: BLACK } };
      tr.getCell(cQty).value = items.length ? { formula: 'SUM(' + L(cQty) + first + ':' + L(cQty) + last + ')' } : 0;
      tr.getCell(cTotal).value = items.length ? { formula: 'SUM(' + L(cTotal) + first + ':' + L(cTotal) + last + ')' } : 0;
      tr.getCell(cTotal).numFmt = '#,##0.00';
      for (var k = 1; k <= N; k++) {
        var cell = tr.getCell(k); cell.fill = fill(GREEN_10); cell.border = { top: { style: 'medium', color: { argb: GREEN } }, bottom: thin, left: thin, right: thin };
        if (k === cQty || k === cTotal) { cell.font = { name: FONT, size: 10, bold: true, color: { argb: BLACK } }; cell.alignment = { horizontal: k === cQty ? 'center' : 'right', vertical: 'middle' }; }
      }
      r += 2;
      ws.mergeCells(r, 1, r, N);
      var note = ws.getCell(r, 1);
      note.value = 'Yellow cells are for Lucas sales to complete (unit price, availability / lead time, remarks). Dimensions are for reference only; always confirm against the physical part before ordering.';
      note.font = { name: FONT, size: 9, italic: true, color: { argb: MUTED } }; note.alignment = { wrapText: true };
      ws.getRow(r).height = 26;

      ws.views = [{ state: 'frozen', ySplit: headRow, showGridLines: false }];
      if (items.length) ws.autoFilter = { from: { row: headRow, column: 1 }, to: { row: last, column: N } };
      ws.pageSetup.printTitlesRow = headRow + ':' + headRow;

      return wb.xlsx.writeBuffer().then(function (buf) {
        return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      });
    });
  }

  window.LucasRFQExcel = { build: build, preload: loadLib, fileName: function (ref) { return 'Lucas-RFQ-' + (ref || 'request') + '.xlsx'; } };
})();
