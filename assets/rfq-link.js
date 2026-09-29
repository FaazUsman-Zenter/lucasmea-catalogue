/* Lucas online catalogues — packs a request for price into a link, and unpacks it again.
 *
 * The whole request (customer details + parts) is compressed into the part of the
 * link after "#". That part never leaves the browser (it is not sent to any server),
 * and the /rfq/ page rebuilds the Excel file from it. No storage or paid service needed.
 *
 *   LucasRFQLink.make(req)    -> Promise<string>  full https://… link
 *   LucasRFQLink.read(hash)   -> Promise<req>     same shape LucasRFQExcel.build() expects
 */
(function () {
  'use strict';
  var ORIGIN = 'https://lucasmeacatalogue.com';
  var PATH = '/rfq/';

  function b64url(bytes) {
    var s = '';
    for (var i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  }
  function unb64url(str) {
    str = str.replace(/-/g, '+').replace(/_/g, '/');
    while (str.length % 4) str += '=';
    var bin = atob(str), out = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }
  function pipe(bytes, stream) {
    var s = new Blob([bytes]).stream().pipeThrough(stream);
    return new Response(s).arrayBuffer().then(function (b) { return new Uint8Array(b); });
  }
  var canZip = typeof CompressionStream === 'function' && typeof DecompressionStream === 'function';

  // Compact form: detail columns are listed once, each part stores only its values.
  function pack(req) {
    var cols = [], seen = {};
    (req.items || []).forEach(function (it) {
      (it.details || []).forEach(function (d) { if (d && d[0] && !seen[d[0]]) { seen[d[0]] = 1; cols.push(d[0]); } });
    });
    return {
      v: 1, r: req.ref, d: (req.date || new Date()).toISOString(),
      n: req.name, c: req.company, e: req.email, p: req.phone, k: req.country, t: req.type, m: req.message,
      h: cols,
      i: (req.items || []).map(function (it) {
        var map = {}; (it.details || []).forEach(function (d) { map[d[0]] = d[1]; });
        return [it.catalogue || '', it.no, Number(it.qty) || 1, it.make || '', it.app || '', it.oe || '',
                cols.map(function (h) { return map[h] == null ? '' : map[h]; })];
      })
    };
  }
  function unpack(o) {
    var cols = o.h || [];
    return {
      ref: o.r, date: o.d ? new Date(o.d) : new Date(),
      name: o.n || '', company: o.c || '', email: o.e || '', phone: o.p || '', country: o.k || '', type: o.t || '', message: o.m || '',
      items: (o.i || []).map(function (a) {
        var vals = a[6] || [];
        return {
          catalogue: a[0], no: a[1], qty: a[2], make: a[3], app: a[4], oe: a[5],
          details: cols.length ? cols.map(function (h, j) { return [h, vals[j] == null ? '' : vals[j]]; }) : null
        };
      })
    };
  }

  function make(req) {
    var bytes = new TextEncoder().encode(JSON.stringify(pack(req)));
    var p = canZip ? pipe(bytes, new CompressionStream('deflate-raw')).then(function (z) { return 'z' + b64url(z); })
                   : Promise.resolve('j' + b64url(bytes));
    return p.catch(function () { return 'j' + b64url(bytes); }).then(function (code) { return ORIGIN + PATH + '#' + code; });
  }

  function read(hash) {
    var code = String(hash || '').replace(/^#/, '');
    if (!code) return Promise.reject(new Error('This link has no request in it.'));
    var kind = code.charAt(0), bytes;
    try { bytes = unb64url(code.slice(1)); } catch (e) { return Promise.reject(new Error('This link is damaged or incomplete.')); }
    var p;
    if (kind === 'z') {
      if (!canZip) return Promise.reject(new Error('Please open this link in an up-to-date browser (Chrome, Edge, Safari or Firefox).'));
      p = pipe(bytes, new DecompressionStream('deflate-raw'));
    } else if (kind === 'j') p = Promise.resolve(bytes);
    else return Promise.reject(new Error('This link is damaged or incomplete.'));
    return p.then(function (b) { return unpack(JSON.parse(new TextDecoder().decode(b))); })
      .catch(function (e) { throw new Error(e && /browser/.test(e.message) ? e.message : 'This link is damaged or incomplete. Copy the whole link from the email and try again.'); });
  }

  window.LucasRFQLink = { make: make, read: read };
})();
