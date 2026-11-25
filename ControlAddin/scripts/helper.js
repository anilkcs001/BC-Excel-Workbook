/* AK Excel Workbook - helper.js
 * Utilities shared by main.js: AL bridge with chunking, date conversion, Excel <-> Univer
 * converters, DOM helpers and icons. No business logic lives here.
 */
(function (global) {
  'use strict';

  var H = {};

  // ---------------------------------------------------------------- AL bridge
  H.isBC = function () {
    return !!(global.Microsoft && global.Microsoft.Dynamics && global.Microsoft.Dynamics.NAV);
  };

  /** Calls an AL event and resolves once BC has accepted it, so calls stay in order. */
  H.invoke = function (name, args) {
    return new Promise(function (resolve, reject) {
      if (!H.isBC()) {
        console.warn('[AK] Not running in Business Central, skipped', name);
        resolve();
        return;
      }
      try {
        global.Microsoft.Dynamics.NAV.InvokeExtensibilityMethod(name, args || [], false, function () { resolve(); });
      } catch (e) {
        reject(e);
      }
    });
  };

  /** Sends a large text to AL in ordered chunks through the OnPayloadChunk event. */
  H.sendChunked = function (purpose, text, chunkSize) {
    var size = chunkSize > 0 ? chunkSize : 400000;
    var count = Math.max(1, Math.ceil(text.length / size));
    var chain = Promise.resolve();
    for (var i = 0; i < count; i++) {
      (function (index) {
        chain = chain.then(function () {
          return H.invoke('OnPayloadChunk', [purpose, index, count, text.substr(index * size, size)]);
        });
      })(i);
    }
    return chain;
  };

  /** Reassembles chunks received from AL. Returns the full text when the last chunk arrives, else null. */
  H.ChunkAssembler = function () {
    this.parts = {};
  };
  H.ChunkAssembler.prototype.add = function (purpose, index, count, data) {
    if (index === 0 || !this.parts[purpose]) {
      this.parts[purpose] = { count: count, items: new Array(count), received: 0 };
    }
    var p = this.parts[purpose];
    if (p.items[index] === undefined) p.received++;
    p.items[index] = data;
    if (p.received < p.count) return null;
    delete this.parts[purpose];
    return p.items.join('');
  };

  // ---------------------------------------------------------------- general
  H.debounce = function (fn, ms) {
    var t = null;
    var wrapped = function () {
      var args = arguments, self = this;
      clearTimeout(t);
      t = setTimeout(function () { fn.apply(self, args); }, ms);
    };
    wrapped.cancel = function () { clearTimeout(t); };
    return wrapped;
  };

  H.uid = function (prefix) {
    return (prefix || 'id') + '-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 8);
  };

  H.escapeHtml = function (s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  };

  H.formatInt = function (n) {
    return Number(n || 0).toLocaleString();
  };

  H.el = function (tag, className, html) {
    var e = document.createElement(tag);
    if (className) e.className = className;
    if (html != null) e.innerHTML = html;
    return e;
  };

  /** Excel sheet names: max 31 chars, none of []:*?/\ */
  H.safeSheetName = function (name) {
    var n = String(name || 'Sheet').replace(/[\[\]:*?\/\\]/g, ' ').trim();
    return (n || 'Sheet').substring(0, 31);
  };

  H.colLetter = function (index) {
    var s = '', n = index + 1;
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  };

  // ---------------------------------------------------------------- dates (Excel serial numbers)
  var EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);

  /** '2026-10-23' -> 46318 */
  H.isoDateToSerial = function (iso) {
    if (!iso) return '';
    var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso);
    if (!m) return iso;
    return (Date.UTC(+m[1], +m[2] - 1, +m[3]) - EXCEL_EPOCH_UTC) / 86400000;
  };

  /** ISO UTC date-time -> serial in the user's local time zone. */
  H.isoDateTimeToSerial = function (iso) {
    if (!iso) return '';
    var d = new Date(iso);
    if (isNaN(d.getTime())) return iso;
    var localMs = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds());
    return (localMs - EXCEL_EPOCH_UTC) / 86400000;
  };

  /** JS Date (from ExcelJS, which reads dates as UTC) -> serial */
  H.dateToSerial = function (d) {
    return (d.getTime() - EXCEL_EPOCH_UTC) / 86400000;
  };

  // ---------------------------------------------------------------- base64 <-> binary
  H.arrayBufferToBase64 = function (buffer) {
    var bytes = new Uint8Array(buffer), out = [], step = 0x8000;
    for (var i = 0; i < bytes.length; i += step) {
      out.push(String.fromCharCode.apply(null, bytes.subarray(i, i + step)));
    }
    return btoa(out.join(''));
  };

  H.base64ToArrayBuffer = function (b64) {
    var bin = atob(b64), len = bin.length, bytes = new Uint8Array(len);
    for (var i = 0; i < len; i++) bytes[i] = bin.charCodeAt(i);
    return bytes.buffer;
  };

  // ---------------------------------------------------------------- colors
  H.argbToHex = function (argb) {
    if (!argb || typeof argb !== 'string') return null;
    var hex = argb.length === 8 ? argb.substring(2) : argb;
    return /^[0-9a-fA-F]{6}$/.test(hex) ? '#' + hex.toUpperCase() : null;
  };

  H.hexToArgb = function (color) {
    if (!color || typeof color !== 'string') return null;
    var c = color.trim();
    var rgb = /^rgba?\((\d+),\s*(\d+),\s*(\d+)/i.exec(c);
    if (rgb) {
      return 'FF' + [rgb[1], rgb[2], rgb[3]].map(function (x) { return ('0' + (+x).toString(16)).slice(-2); }).join('').toUpperCase();
    }
    if (c[0] === '#') c = c.substring(1);
    if (c.length === 3) c = c[0] + c[0] + c[1] + c[1] + c[2] + c[2];
    return /^[0-9a-fA-F]{6}$/.test(c) ? 'FF' + c.toUpperCase() : null;
  };

  // ---------------------------------------------------------------- Univer style <-> ExcelJS style
  var H_ALIGN_TO_XL = { 1: 'left', 2: 'center', 3: 'right', 4: 'justify' };
  var XL_TO_H_ALIGN = { left: 1, center: 2, right: 3, justify: 4 };
  var V_ALIGN_TO_XL = { 1: 'top', 2: 'middle', 3: 'bottom' };
  var XL_TO_V_ALIGN = { top: 1, middle: 2, bottom: 3 };

  H.univerStyleToExcel = function (s) {
    var out = {};
    if (!s) return out;
    var font = {};
    if (s.bl === 1) font.bold = true;
    if (s.it === 1) font.italic = true;
    if (s.ul && s.ul.s === 1) font.underline = true;
    if (s.st && s.st.s === 1) font.strike = true;
    if (s.fs) font.size = s.fs;
    if (s.ff) font.name = s.ff;
    if (s.cl && s.cl.rgb) { var fc = H.hexToArgb(s.cl.rgb); if (fc) font.color = { argb: fc }; }
    if (Object.keys(font).length) out.font = font;
    if (s.bg && s.bg.rgb) {
      var bg = H.hexToArgb(s.bg.rgb);
      if (bg) out.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
    }
    if (s.n && s.n.pattern) out.numFmt = s.n.pattern;
    var al = {};
    if (H_ALIGN_TO_XL[s.ht]) al.horizontal = H_ALIGN_TO_XL[s.ht];
    if (V_ALIGN_TO_XL[s.vt]) al.vertical = V_ALIGN_TO_XL[s.vt];
    if (s.tb === 3) al.wrapText = true;
    if (Object.keys(al).length) out.alignment = al;
    if (s.bd) {
      var border = {};
      ['t', 'b', 'l', 'r'].forEach(function (k) {
        var b = s.bd[k];
        if (b && b.s) {
          var side = { t: 'top', b: 'bottom', l: 'left', r: 'right' }[k];
          border[side] = { style: b.s >= 8 ? 'medium' : 'thin' };
          var bc = b.cl && H.hexToArgb(b.cl.rgb);
          if (bc) border[side].color = { argb: bc };
        }
      });
      if (Object.keys(border).length) out.border = border;
    }
    return out;
  };

  H.excelStyleToUniver = function (cell) {
    var s = {}, f = cell.font, fill = cell.fill, al = cell.alignment;
    if (f) {
      if (f.bold) s.bl = 1;
      if (f.italic) s.it = 1;
      if (f.underline) s.ul = { s: 1 };
      if (f.strike) s.st = { s: 1 };
      if (f.size) s.fs = f.size;
      if (f.name) s.ff = f.name;
      var fc = f.color && H.argbToHex(f.color.argb);
      if (fc) s.cl = { rgb: fc };
    }
    if (fill && fill.type === 'pattern' && fill.pattern === 'solid' && fill.fgColor) {
      var bg = H.argbToHex(fill.fgColor.argb);
      if (bg) s.bg = { rgb: bg };
    }
    if (cell.numFmt && cell.numFmt !== 'General') s.n = { pattern: cell.numFmt };
    if (al) {
      if (XL_TO_H_ALIGN[al.horizontal]) s.ht = XL_TO_H_ALIGN[al.horizontal];
      if (XL_TO_V_ALIGN[al.vertical]) s.vt = XL_TO_V_ALIGN[al.vertical];
      if (al.wrapText) s.tb = 3;
    }
    return Object.keys(s).length ? s : null;
  };

  /** 'B2:D5' -> {startRow:1,endRow:4,startColumn:1,endColumn:3} */
  H.decodeA1Range = function (ref) {
    var parts = String(ref).split(':');
    var a = decodeA1(parts[0]), b = decodeA1(parts[1] || parts[0]);
    if (!a || !b) return null;
    return { startRow: Math.min(a.r, b.r), endRow: Math.max(a.r, b.r), startColumn: Math.min(a.c, b.c), endColumn: Math.max(a.c, b.c) };
  };
  function decodeA1(s) {
    var m = /^\$?([A-Z]+)\$?(\d+)$/i.exec(String(s).trim());
    if (!m) return null;
    var col = 0, letters = m[1].toUpperCase();
    for (var i = 0; i < letters.length; i++) col = col * 26 + (letters.charCodeAt(i) - 64);
    return { r: +m[2] - 1, c: col - 1 };
  }

  // ---------------------------------------------------------------- toasts
  H.toast = function (level, message) {
    var host = document.getElementById('ak-toasts');
    if (!host) { console.log('[AK]', level, message); return; }
    var lvl = ['info', 'success', 'warning', 'error'].indexOf(level) >= 0 ? level : 'info';
    var t = H.el('div', 'ak-toast ak-toast--' + lvl);
    t.setAttribute('role', lvl === 'error' ? 'alert' : 'status');
    t.innerHTML = '<span class="ak-toast__icon">' + H.icon(lvl === 'success' ? 'check' : lvl === 'error' ? 'alert' : 'info') + '</span>' +
      '<span class="ak-toast__msg">' + H.escapeHtml(message) + '</span>' +
      '<button class="ak-toast__close" aria-label="Dismiss">' + H.icon('x') + '</button>';
    var remove = function () {
      t.classList.add('is-leaving');
      setTimeout(function () { if (t.parentNode) t.parentNode.removeChild(t); }, 200);
    };
    t.querySelector('button').addEventListener('click', remove);
    host.appendChild(t);
    while (host.children.length > 4) host.removeChild(host.firstChild);
    setTimeout(remove, lvl === 'error' ? 8000 : 4000);
  };

  // ---------------------------------------------------------------- icons (Fluent-style outline, 20px)
  var P = {
    database: '<ellipse cx="10" cy="4.5" rx="6" ry="2.5"/><path d="M4 4.5v11c0 1.4 2.7 2.5 6 2.5s6-1.1 6-2.5v-11"/><path d="M4 10c0 1.4 2.7 2.5 6 2.5s6-1.1 6-2.5"/>',
    refresh: '<path d="M16 10a6 6 0 1 1-1.8-4.3"/><path d="M16 3.5V7h-3.5"/>',
    x: '<path d="M5 5l10 10M15 5L5 15"/>',
    search: '<circle cx="8.5" cy="8.5" r="5"/><path d="M12.5 12.5L17 17"/>',
    table: '<rect x="3" y="4" width="14" height="12" rx="2"/><path d="M3 8h14M8 8v8"/>',
    check: '<path d="M4.5 10.5l3.5 3.5 7.5-8"/>',
    alert: '<circle cx="10" cy="10" r="7"/><path d="M10 6.5v4.5M10 13.5v.1"/>',
    info: '<circle cx="10" cy="10" r="7"/><path d="M10 9v4.5M10 6.5v.1"/>',
    panel: '<rect x="3" y="4" width="14" height="12" rx="2"/><path d="M12 4v12"/>',
    link: '<path d="M8.5 11.5l3-3"/><path d="M9.5 6l1.3-1.3a3 3 0 0 1 4.3 4.3L13.8 10.3"/><path d="M10.5 14l-1.3 1.3a3 3 0 0 1-4.3-4.3L6.2 9.7"/>',
    unlink: '<path d="M9.5 6l1.3-1.3a3 3 0 0 1 4.3 4.3L13.8 10.3"/><path d="M10.5 14l-1.3 1.3a3 3 0 0 1-4.3-4.3L6.2 9.7"/><path d="M4 4l12 12"/>',
    key: '<circle cx="7" cy="12" r="3.5"/><path d="M9.5 9.5L16 3M13 6l2 2"/>',
    sigma: '<path d="M14.5 5H6l5 5-5 5h8.5"/>',
    arrowLeft: '<path d="M12 5l-5 5 5 5"/>',
    sheet: '<rect x="3" y="3" width="14" height="14" rx="2"/><path d="M3 8h14M3 12.5h14M8 3v14"/>',
    save: '<path d="M4.5 3h9l3 3v9.5a1.5 1.5 0 0 1-1.5 1.5h-10A1.5 1.5 0 0 1 3.5 15.5v-11A1.5 1.5 0 0 1 4.5 3z"/><path d="M7 3v4h6V3"/><rect x="6.5" y="11" width="7" height="6"/>',
    upload: '<path d="M10 13V3.5M6.5 7L10 3.5 13.5 7"/><path d="M3.5 13v2.5A1.5 1.5 0 0 0 5 17h10a1.5 1.5 0 0 0 1.5-1.5V13"/>',
    download: '<path d="M10 3.5V13M6.5 9.5L10 13l3.5-3.5"/><path d="M3.5 13v2.5A1.5 1.5 0 0 0 5 17h10a1.5 1.5 0 0 0 1.5-1.5V13"/>',
    scale: '<path d="M10 3v14M6 17h8"/><path d="M4 7h12"/><path d="M4 7l-2 5a2.5 2.5 0 0 0 4 0L4 7zM16 7l-2 5a2.5 2.5 0 0 0 4 0l-2-5z"/>',
    bank: '<path d="M3 8l7-4.5L17 8"/><path d="M4.5 8v6.5M8.5 8v6.5M11.5 8v6.5M15.5 8v6.5M3 16.5h14"/>',
    receipt: '<path d="M5 2.5h10v15l-2-1.3-1.5 1.3-1.5-1.3-1.5 1.3L7 16.2l-2 1.3z"/><path d="M7.5 6.5h5M7.5 9.5h5M7.5 12.5h3"/>',
    users: '<circle cx="7.5" cy="7" r="2.8"/><path d="M2.5 16.5c.6-2.8 2.5-4.3 5-4.3s4.4 1.5 5 4.3"/><path d="M13 4.5a2.6 2.6 0 0 1 0 5M14.5 12.4c1.6.5 2.6 1.9 3 4.1"/>',
    sliders: '<path d="M4 5h7M15 5h1M4 10h2M10 10h6M4 15h9M17 15h-1"/><circle cx="13" cy="5" r="1.8"/><circle cx="8" cy="10" r="1.8"/><circle cx="15" cy="15" r="1.8"/>',
    merge: '<circle cx="5" cy="5" r="2"/><circle cx="5" cy="15" r="2"/><circle cx="15" cy="10" r="2"/><path d="M7 5h1.5a3 3 0 0 1 3 3v0a2 2 0 0 0 2 2h0M7 15h1.5a3 3 0 0 0 3-3v0a2 2 0 0 1 2-2"/>'
  };
  H.icon = function (name) {
    return '<svg class="ak-ico" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + (P[name] || '') + '</svg>';
  };

  /**
   * In-add-in choice dialog. Browser confirm() can be blocked inside the BC add-in iframe,
   * so this renders its own modal. buttons: [{ label, value, primary }]. Escape / backdrop -> null.
   */
  H.choose = function (message, buttons, title) {
    return new Promise(function (resolve) {
      var host = document.getElementById('ak-app') || document.body;
      var back = H.el('div', 'ak-modal');
      back.innerHTML = '<div class="ak-modal__box" role="alertdialog" aria-modal="true" aria-label="' + H.escapeHtml(title || 'Confirm') + '">' +
        (title ? '<div class="ak-modal__title">' + H.escapeHtml(title) + '</div>' : '') +
        '<div class="ak-modal__msg">' + H.escapeHtml(message) + '</div>' +
        '<div class="ak-modal__btns"><button class="ak-btn ak-btn--subtle" data-i="-1">Cancel</button>' +
        buttons.map(function (b, i) {
          return '<button class="ak-btn' + (b.primary ? ' ak-btn--primary' : '') + '" data-i="' + i + '">' + H.escapeHtml(b.label) + '</button>';
        }).join('') + '</div></div>';
      var done = function (v) {
        document.removeEventListener('keydown', onKey, true);
        if (back.parentNode) back.parentNode.removeChild(back);
        resolve(v);
      };
      var onKey = function (e) { if (e.key === 'Escape') { e.stopPropagation(); done(null); } };
      back.addEventListener('click', function (e) {
        var b = e.target.closest('[data-i]');
        if (b) { var i = +b.getAttribute('data-i'); done(i < 0 ? null : buttons[i].value); }
        else if (e.target === back) done(null);
      });
      document.addEventListener('keydown', onKey, true);
      host.appendChild(back);
      var primary = back.querySelector('.ak-btn--primary') || back.querySelector('[data-i="0"]');
      if (primary) primary.focus();
    });
  };

  /** Yes/no wrapper around choose(). Resolves true or false. */
  H.confirm = function (message, okLabel) {
    return H.choose(message, [{ label: okLabel || 'OK', value: true, primary: true }]).then(function (v) { return v === true; });
  };

  /** Per-viewer preference in browser storage; never throws (storage can be blocked in iframes). */
  H.pref = function (key, value) {
    try {
      if (arguments.length > 1) { localStorage.setItem('ak.' + key, JSON.stringify(value)); return value; }
      var raw = localStorage.getItem('ak.' + key);
      return raw == null ? undefined : JSON.parse(raw);
    } catch (e) { return undefined; }
  };

  global.AKHelper = H;
})(window);
