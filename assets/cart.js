/* Lucas online catalogues — shared request-for-price cart.
 *
 * One cart for every catalogue on this site (stored in the visitor's browser).
 *
 * Catalogue pages: add ONE line before </body>:
 *   <script src="/assets/cart.js" data-catalogue="Gas Springs"></script>
 * The script adds a tick box to every row that has a .part-no cell, plus an
 * "Add to cart" bar and a cart button. No other changes to the catalogue are needed.
 *
 * Any page: elements with [data-cart-count] show the number of parts in the cart.
 */
(function () {
  'use strict';
  var KEY = 'lucas-rfq-cart-v1';
  var CART_URL = '/cart/';
  var script = document.currentScript;
  var CATALOGUE = script && script.getAttribute('data-catalogue');

  /* ---------- storage ---------- */
  var memory = [];
  var listeners = [];
  function load() {
    try {
      var v = JSON.parse(localStorage.getItem(KEY) || '[]');
      return Array.isArray(v) ? v : [];
    } catch (e) { return memory.slice(); }
  }
  function save(items) {
    memory = items.slice();
    try { localStorage.setItem(KEY, JSON.stringify(items)); } catch (e) { /* private mode: keep in memory */ }
    emit();
  }
  function emit() { var items = load(); listeners.forEach(function (fn) { try { fn(items); } catch (e) {} }); updateCounts(items); }
  function keyOf(it) { return (it.catalogue || '') + '|' + it.no; }
  function count(items) { return (items || load()).length; }

  var api = {
    items: load,
    count: function () { return count(); },
    has: function (no, catalogue) { return load().some(function (i) { return i.no === no && (!catalogue || i.catalogue === catalogue); }); },
    add: function (list) {
      var items = load(), seen = {};
      items.forEach(function (i) { seen[keyOf(i)] = true; });
      var added = 0;
      list.forEach(function (it) {
        if (!it || !it.no || seen[keyOf(it)]) return;
        items.push({ no: it.no, catalogue: it.catalogue || '', make: it.make || '', app: it.app || '', oe: it.oe || '', details: it.details || null, qty: Math.max(1, parseInt(it.qty, 10) || 1) });
        seen[keyOf(it)] = true; added++;
      });
      save(items); return added;
    },
    remove: function (no, catalogue) { save(load().filter(function (i) { return !(i.no === no && i.catalogue === catalogue); })); },
    setQty: function (no, catalogue, qty) {
      var q = Math.max(1, Math.min(99999, parseInt(qty, 10) || 1));
      save(load().map(function (i) { if (i.no === no && i.catalogue === catalogue) i.qty = q; return i; }));
    },
    clear: function () { save([]); },
    onChange: function (fn) { listeners.push(fn); },
    url: CART_URL
  };
  window.LucasCart = api;
  window.addEventListener('storage', function (e) { if (e.key === KEY) emit(); });

  function updateCounts(items) {
    var n = count(items);
    var els = document.querySelectorAll('[data-cart-count]');
    for (var i = 0; i < els.length; i++) { els[i].textContent = n; els[i].setAttribute('data-empty', n ? 'false' : 'true'); }
  }

  /* ---------- catalogue enhancement ---------- */
  var CSS = [
    '.lc-sel{width:44px;text-align:center;padding-left:10px!important;padding-right:4px!important;vertical-align:middle!important}',
    '.lc-sel input{width:18px;height:18px;accent-color:#00954C;cursor:pointer;margin:0;vertical-align:middle}',
    '.lc-sel input:disabled{cursor:default;opacity:.55}',
    'tr.lc-picked td{background:#e6f4ed!important}',
    'tr.lc-incart td{background:#f3faf6}',
    '.lc-incart-tag{display:block;font-size:10px;line-height:1;margin-top:3px;color:#00954C;font-weight:700;letter-spacing:.02em;white-space:nowrap}',
    '.lc-bar{position:fixed;right:16px;bottom:16px;z-index:9999;display:flex;align-items:center;gap:10px;flex-wrap:wrap;justify-content:flex-end;',
    ' font-family:"League Spartan",Arial,sans-serif;max-width:calc(100vw - 32px)}',
    '.lc-pick{display:none;align-items:center;gap:10px;background:#231F20;color:#fff;padding:8px 8px 8px 16px;border-radius:6px;box-shadow:0 6px 20px rgba(0,0,0,.25)}',
    '.lc-pick.show{display:flex}',
    '.lc-pick span{font-size:15px;white-space:nowrap}',
    '.lc-btn{font:inherit;font-weight:600;font-size:15px;border:0;border-radius:4px;padding:10px 16px 8px;cursor:pointer;text-decoration:none;display:inline-flex;align-items:center;gap:8px;white-space:nowrap}',
    '.lc-btn.add{background:#00954C;color:#fff}.lc-btn.add:hover{background:#007a3e}',
    '.lc-btn.ghost{background:transparent;color:#fff;padding:10px 8px 8px;text-decoration:underline}',
    '.lc-btn.cart{background:#fff;color:#231F20;border:2px solid #00954C;box-shadow:0 6px 20px rgba(0,0,0,.18)}',
    '.lc-btn.cart:hover{background:#e6f4ed}',
    '.lc-btn.cart b{background:#00954C;color:#fff;border-radius:20px;min-width:24px;padding:3px 7px 1px;font-size:13px;text-align:center}',
    '.lc-btn:focus-visible{outline:3px solid #231F20;outline-offset:2px}',
    '.lc-toast{position:fixed;left:50%;bottom:84px;transform:translateX(-50%) translateY(20px);opacity:0;transition:.2s;z-index:10000;',
    ' background:#00954C;color:#fff;padding:12px 18px 10px;border-radius:6px;font-family:"League Spartan",Arial,sans-serif;font-weight:600;font-size:15px;box-shadow:0 6px 20px rgba(0,0,0,.25);pointer-events:none;max-width:calc(100vw - 32px);text-align:center}',
    '.lc-toast.show{opacity:1;transform:translateX(-50%) translateY(0)}',
    '@media (max-width:600px){.lc-bar{right:8px;left:8px;bottom:8px;max-width:none}.lc-pick{flex:1;justify-content:space-between}.lc-btn{font-size:14px}}'
  ].join('\n');

  function text(el) { return el ? (el.textContent || '').trim() : ''; }

  function enhanceCatalogue() {
    var tbody = document.getElementById('resultsBody') || document.querySelector('table tbody');
    if (!tbody) return;
    var table = tbody.closest('table');
    var headRow = table && table.querySelector('thead tr');
    if (!headRow) return;

    var style = document.createElement('style'); style.textContent = CSS; document.head.appendChild(style);

    // map column names before inserting our column
    var heads = Array.prototype.map.call(headRow.children, function (th) { return text(th).toLowerCase(); });
    // Readable column labels for the Excel sheet, e.g. "Length (mm)", "Rod Ø (mm)"
    var labels = Array.prototype.map.call(headRow.children, function (th) {
      var c = th.cloneNode(true);
      Array.prototype.forEach.call(c.querySelectorAll('.key,[aria-hidden="true"]'), function (k) { k.remove(); });
      Array.prototype.forEach.call(c.querySelectorAll('small'), function (sm) { var u = text(sm); sm.textContent = u ? ' (' + u + ')' : ''; });
      return text(c).replace(/\s+/g, ' ');
    });
    function col(re) { for (var i = 0; i < heads.length; i++) if (re.test(heads[i])) return i; return -1; }
    var cMake = col(/^make/), cApp = col(/applic/), cOe = col(/^oe/);

    var selected = {}; // no -> item

    var thSel = document.createElement('th');
    thSel.className = 'lc-sel';
    thSel.innerHTML = '<input type="checkbox" aria-label="Select all parts shown" title="Select all parts shown">';
    headRow.insertBefore(thSel, headRow.firstChild);
    var allBox = thSel.querySelector('input');

    function rowItem(tr) {
      var cells = tr.children, off = tr.__lcOff || 0;
      function c(i) { return i < 0 ? '' : text(cells[i + off]); }
      var details = [];
      for (var i = 0; i < labels.length; i++) { if (labels[i]) details.push([labels[i], c(i)]); }
      return { no: tr.__lcNo, catalogue: CATALOGUE, make: c(cMake), app: c(cApp), oe: c(cOe), details: details };
    }

    function paint(tr) {
      var box = tr.__lcBox; if (!box) return;
      var inCart = api.has(tr.__lcNo, CATALOGUE);
      box.disabled = inCart;
      box.checked = inCart || !!selected[tr.__lcNo];
      box.title = inCart ? 'Already in your cart' : 'Select ' + tr.__lcNo;
      tr.classList.toggle('lc-incart', inCart);
      tr.classList.toggle('lc-picked', !inCart && !!selected[tr.__lcNo]);
      var tag = tr.__lcTag;
      if (inCart && !tag) { tag = document.createElement('span'); tag.className = 'lc-incart-tag'; tag.textContent = 'IN CART'; box.parentNode.appendChild(tag); tr.__lcTag = tag; }
      if (!inCart && tag) { tag.remove(); tr.__lcTag = null; }
    }

    function enhanceRow(tr) {
      if (tr.__lcBox || tr.parentNode !== tbody) return;
      var pn = tr.querySelector('.part-no');
      if (!pn) return;
      tr.__lcNo = text(pn);
      var td = document.createElement('td');
      td.className = 'lc-sel';
      td.innerHTML = '<input type="checkbox">';
      tr.insertBefore(td, tr.firstChild);
      tr.__lcOff = 1;
      var box = td.querySelector('input');
      box.setAttribute('aria-label', 'Select ' + tr.__lcNo);
      tr.__lcBox = box;
      box.addEventListener('change', function () {
        if (box.checked) selected[tr.__lcNo] = rowItem(tr); else delete selected[tr.__lcNo];
        paint(tr); refreshBar();
      });
      paint(tr);
    }

    function rows() { return Array.prototype.filter.call(tbody.children, function (tr) { return tr.__lcBox; }); }
    function enhanceAll() { Array.prototype.forEach.call(tbody.children, enhanceRow); syncAll(); }
    function syncAll() {
      var rs = rows(), selectable = rs.filter(function (tr) { return !tr.__lcBox.disabled; });
      allBox.checked = selectable.length > 0 && selectable.every(function (tr) { return selected[tr.__lcNo]; });
      allBox.indeterminate = !allBox.checked && selectable.some(function (tr) { return selected[tr.__lcNo]; });
    }

    allBox.addEventListener('change', function () {
      rows().forEach(function (tr) {
        if (tr.__lcBox.disabled) return;
        if (allBox.checked) selected[tr.__lcNo] = rowItem(tr); else delete selected[tr.__lcNo];
        paint(tr);
      });
      refreshBar();
    });

    new MutationObserver(enhanceAll).observe(tbody, { childList: true });
    enhanceAll();

    // bar
    var bar = document.createElement('div');
    bar.className = 'lc-bar';
    bar.innerHTML =
      '<div class="lc-pick" role="region" aria-label="Selected parts"><span><b class="lc-n">0</b> selected</span>' +
      '<button type="button" class="lc-btn ghost lc-clr">Clear</button>' +
      '<button type="button" class="lc-btn add lc-add">Add to cart</button></div>' +
      '<a class="lc-btn cart" href="' + CART_URL + '">View cart <b data-cart-count>0</b></a>';
    document.body.appendChild(bar);
    var pick = bar.querySelector('.lc-pick'), nEl = bar.querySelector('.lc-n');

    var toast = document.createElement('div');
    toast.className = 'lc-toast'; toast.setAttribute('role', 'status'); toast.setAttribute('aria-live', 'polite');
    document.body.appendChild(toast);
    var tt;
    function say(msg) { toast.textContent = msg; toast.classList.add('show'); clearTimeout(tt); tt = setTimeout(function () { toast.classList.remove('show'); }, 2600); }

    function refreshBar() {
      var n = Object.keys(selected).length;
      nEl.textContent = n; pick.classList.toggle('show', n > 0);
      syncAll();
    }
    bar.querySelector('.lc-clr').addEventListener('click', function () { selected = {}; rows().forEach(paint); refreshBar(); });
    bar.querySelector('.lc-add').addEventListener('click', function () {
      var list = Object.keys(selected).map(function (k) { return selected[k]; });
      var added = api.add(list);
      selected = {};
      rows().forEach(paint); refreshBar();
      say(added === 1 ? '1 part added to your cart' : added + ' parts added to your cart');
    });

    api.onChange(function () { rows().forEach(function (tr) { if (api.has(tr.__lcNo, CATALOGUE)) delete selected[tr.__lcNo]; paint(tr); }); refreshBar(); });
  }

  // Make the brand logo in a catalogue's header link back to the home page
  function linkLogoHome() {
    var img = document.querySelector('header img');
    if (!img || img.closest('a')) return;
    var a = document.createElement('a');
    a.href = '/';
    a.setAttribute('aria-label', 'Lucas online catalogues home');
    a.style.display = 'inline-block';
    img.parentNode.insertBefore(a, img);
    a.appendChild(img);
  }

  function start() {
    if (CATALOGUE) linkLogoHome();
    updateCounts(load());
    if (CATALOGUE) enhanceCatalogue();
    updateCounts(load());
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();
})();
