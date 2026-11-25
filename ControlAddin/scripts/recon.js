/* AK Excel Workbook - recon.js
 * Reconciliation engine (Phase 3). Pure functions: no DOM, no Univer, so it can be unit-tested.
 *
 *   AKRecon.run(left, right, def) -> { header, types, rows, stats }
 *     left/right: { name, header[], rows[][] }  (raw sheet values)
 *     def: see defaults() below
 *
 * Matching passes (each row is used at most once):
 *   1. Key + amount   - same key and |difference| <= tolerance          -> "Matched"
 *   2. Key            - same key, amounts differ                         -> "Amount difference"
 *   3. Amount + date  - no key match; same amount within tolerance and
 *                       dates within N days (closest date wins)          -> "Matched"
 *   Leftovers         -> "Only in <left label>" / "Only in <right label>"
 */
(function (global) {
  'use strict';

  var R = {};

  R.defaults = function () {
    return {
      leftLabel: 'Books',
      rightLabel: 'Statement',
      keys: [],                    // [{ left: 'Invoice No.', right: 'Invoice number' }]
      normalize: true,             // document numbers: ignore case, symbols, leading zeros
      group: false,                // sum amounts per key on each side before matching
      amount: { left: '', leftMinus: '', right: '', rightMinus: '' },
      flip: false,                 // right side amounts have the opposite sign
      tolerance: 0.01,
      date: { left: '', right: '' },
      useAmountDate: false,        // pass 3
      days: 3,
      extraLeft: [],
      extraRight: []
    };
  };

  // ------------------------------------------------------------------ value parsing
  /** Numbers from cells or text: "1,23,456.50", "(500.00)", "₹ 1,200", "-12.5". Blank/invalid -> 0. */
  R.toNumber = function (v) {
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    if (v == null || v === '') return 0;
    if (typeof v === 'boolean') return v ? 1 : 0;
    var s = String(v).trim();
    var neg = /^\(.*\)$/.test(s) || /^-/.test(s);
    s = s.replace(/[^0-9.]/g, '');
    var n = parseFloat(s);
    if (!isFinite(n)) return 0;
    return neg ? -n : n;
  };

  var MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  var EPOCH = Date.UTC(1899, 11, 30);
  function serial(y, m, d) {
    if (y < 100) y += 2000;
    if (m < 1 || m > 12 || d < 1 || d > 31) return null;
    return (Date.UTC(y, m - 1, d) - EPOCH) / 86400000;
  }

  /** Excel serial from a cell: serial numbers pass through; text dates use Indian order (dd-mm-yyyy). */
  R.toDate = function (v) {
    if (typeof v === 'number') return v > 0 ? Math.floor(v) : null;
    if (v == null || v === '') return null;
    var s = String(v).trim(), m;
    if ((m = /^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/.exec(s))) return serial(+m[1], +m[2], +m[3]);
    if ((m = /^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})/.exec(s))) return serial(+m[3], +m[2], +m[1]);
    if ((m = /^(\d{1,2})[-\s\/.]([A-Za-z]{3})[A-Za-z]*[-\s\/.,]+(\d{2,4})/.exec(s))) {
      var mo = MONTHS[m[2].toLowerCase()];
      return mo ? serial(+m[3], mo, +m[1]) : null;
    }
    return null;
  };

  /** Key normaliser. With normalize: "inv/0012" == "INV-12" == "Inv 12". */
  R.normKey = function (v, normalize) {
    if (v == null) return '';
    if (typeof v === 'number') return String(Math.round(v * 1e6) / 1e6);
    var s = String(v).trim();
    if (!normalize) return s;
    return s.toUpperCase()
      .replace(/[^A-Z0-9]/g, '')
      .replace(/(^|[A-Z])0+(?=\d)/g, '$1');
  };

  // ------------------------------------------------------------------ helpers
  function idx(table, name, what) {
    if (!name) return -1;
    var i = table.header.indexOf(name);
    if (i < 0) throw new Error(what + ' column "' + name + '" was not found in sheet "' + table.name + '".');
    return i;
  }

  function round2(n) { return Math.round(n * 100) / 100; }

  /**
   * Turns a sheet into match items: { key, amount, date, row, lines }.
   * With group = true, rows sharing a key collapse into one item with the summed amount.
   */
  function prepare(table, side, def) {
    var keyIdx = def.keys.map(function (k) { return idx(table, k[side], 'Key'); });
    var amtIdx = idx(table, def.amount[side], 'Amount');
    var minusIdx = idx(table, def.amount[side + 'Minus'], 'Amount');
    var dateIdx = idx(table, def.date[side], 'Date');
    var sign = side === 'right' && def.flip ? -1 : 1;

    var items = table.rows.map(function (row) {
      var amt = amtIdx >= 0 ? R.toNumber(row[amtIdx]) : 0;
      if (minusIdx >= 0) amt -= R.toNumber(row[minusIdx]);
      return {
        key: keyIdx.length ? keyIdx.map(function (i) { return R.normKey(row[i], def.normalize); }).join('\u0001') : '',
        amount: round2(amt * sign),
        date: dateIdx >= 0 ? R.toDate(row[dateIdx]) : null,
        row: row,
        lines: 1
      };
    });

    if (def.group && keyIdx.length) {
      var byKey = new Map(), grouped = [];
      items.forEach(function (it) {
        if (!it.key.replace(/\u0001/g, '')) { grouped.push(it); return; }
        var g = byKey.get(it.key);
        if (!g) { g = { key: it.key, amount: 0, date: it.date, row: it.row, lines: 0 }; byKey.set(it.key, g); grouped.push(g); }
        g.amount = round2(g.amount + it.amount);
        g.lines++;
        if (g.date == null || (it.date != null && it.date < g.date)) g.date = it.date;
      });
      items = grouped;
    }
    return { items: items, keyIdx: keyIdx, dateIdx: dateIdx };
  }

  // ------------------------------------------------------------------ engine
  R.run = function (left, right, inDef) {
    var def = Object.assign(R.defaults(), inDef || {});
    def.amount = Object.assign(R.defaults().amount, def.amount || {});
    def.date = Object.assign(R.defaults().date, def.date || {});
    if (!def.amount.left || !def.amount.right) throw new Error('Choose the amount column for both sides.');
    if (!def.keys.length && !def.useAmountDate) throw new Error('Add a key column, or turn on matching by amount and date.');
    var tol = Math.max(0, R.toNumber(def.tolerance));
    var days = Math.max(0, Math.round(R.toNumber(def.days)));

    var L = prepare(left, 'left', def), Rt = prepare(right, 'right', def);
    var pairs = [];          // { l, r, status, rule }
    var lUsed = new Array(L.items.length), rUsed = new Array(Rt.items.length);

    // Passes 1 + 2: by key.
    if (def.keys.length) {
      var rByKey = new Map();
      Rt.items.forEach(function (it, i) {
        if (!it.key.replace(/\u0001/g, '')) return;
        var list = rByKey.get(it.key);
        if (!list) rByKey.set(it.key, list = []);
        list.push(i);
      });
      // Pass 1: key + amount within tolerance (closest amount wins).
      L.items.forEach(function (it, li) {
        var list = rByKey.get(it.key);
        if (!list) return;
        var best = -1, bestDiff = Infinity;
        list.forEach(function (ri) {
          if (rUsed[ri]) return;
          var d = Math.abs(it.amount - Rt.items[ri].amount);
          if (d <= tol + 1e-9 && d < bestDiff) { best = ri; bestDiff = d; }
        });
        if (best >= 0) { lUsed[li] = rUsed[best] = true; pairs.push({ l: li, r: best, status: 'match', rule: 'Key + amount' }); }
      });
      // Pass 2: same key, different amount.
      L.items.forEach(function (it, li) {
        if (lUsed[li]) return;
        var list = rByKey.get(it.key);
        if (!list) return;
        for (var k = 0; k < list.length; k++) {
          var ri = list[k];
          if (!rUsed[ri]) { lUsed[li] = rUsed[ri] = true; pairs.push({ l: li, r: ri, status: 'diff', rule: 'Key' }); return; }
        }
      });
    }

    // Pass 3: amount + date window on what is left.
    if (def.useAmountDate) {
      var pool = [];
      Rt.items.forEach(function (it, i) { if (!rUsed[i]) pool.push(i); });
      pool.sort(function (a, b) { return Rt.items[a].amount - Rt.items[b].amount; });
      var amounts = pool.map(function (i) { return Rt.items[i].amount; });
      var lowerBound = function (x) {
        var lo = 0, hi = amounts.length;
        while (lo < hi) { var mid = (lo + hi) >> 1; if (amounts[mid] < x) lo = mid + 1; else hi = mid; }
        return lo;
      };
      L.items.forEach(function (it, li) {
        if (lUsed[li]) return;
        var best = -1, bestScore = Infinity;
        for (var p = lowerBound(it.amount - tol - 1e-9); p < pool.length && amounts[p] <= it.amount + tol + 1e-9; p++) {
          var ri = pool[p];
          if (rUsed[ri]) continue;
          var rd = Rt.items[ri].date, dd = 0;
          if (it.date != null && rd != null) {
            dd = Math.abs(it.date - rd);
            if (dd > days) continue;
          }
          var score = dd * 1000 + Math.abs(it.amount - Rt.items[ri].amount);
          if (score < bestScore) { best = ri; bestScore = score; }
        }
        if (best >= 0) { lUsed[li] = rUsed[best] = true; pairs.push({ l: li, r: best, status: 'match', rule: 'Amount + date' }); }
      });
    }

    L.items.forEach(function (it, li) { if (!lUsed[li]) pairs.push({ l: li, r: -1, status: 'left', rule: '' }); });
    Rt.items.forEach(function (it, ri) { if (!rUsed[ri]) pairs.push({ l: -1, r: ri, status: 'right', rule: '' }); });

    // ---------------------------------------------------------------- output layout
    var LL = def.leftLabel || 'Books', RL = def.rightLabel || 'Statement';
    var STATUS = { match: 'Matched', diff: 'Amount difference', left: 'Only in ' + LL, right: 'Only in ' + RL };
    var exLeft = def.extraLeft.filter(function (h) { return left.header.indexOf(h) >= 0; });
    var exRight = def.extraRight.filter(function (h) { return right.header.indexOf(h) >= 0; });

    var header = ['Status', 'Match rule'], types = ['Text', 'Text'];
    def.keys.forEach(function (k) { header.push(LL + ' ' + k.left); types.push('Text'); });
    if (L.dateIdx >= 0) { header.push(LL + ' date'); types.push('Date'); }
    header.push(LL + ' amount'); types.push('Decimal');
    if (def.group && def.keys.length) { header.push(LL + ' lines'); types.push('Integer'); }
    exLeft.forEach(function (h) { header.push(LL + ' ' + h); types.push('Text'); });
    def.keys.forEach(function (k) { header.push(RL + ' ' + k.right); types.push('Text'); });
    if (Rt.dateIdx >= 0) { header.push(RL + ' date'); types.push('Date'); }
    header.push(RL + ' amount'); types.push('Decimal');
    if (def.group && def.keys.length) { header.push(RL + ' lines'); types.push('Integer'); }
    exRight.forEach(function (h) { header.push(RL + ' ' + h); types.push('Text'); });
    header.push('Difference'); types.push('Decimal');
    if (L.dateIdx >= 0 && Rt.dateIdx >= 0) { header.push('Days apart'); types.push('Integer'); }

    var side = function (table, prep, it, keyCols, extra) {
      var out = [];
      keyCols.forEach(function (i) { out.push(it ? val(it.row[i]) : ''); });
      if (prep.dateIdx >= 0) out.push(it && it.date != null ? it.date : '');
      out.push(it ? it.amount : '');
      if (def.group && def.keys.length) out.push(it ? it.lines : '');
      extra.forEach(function (h) { out.push(it ? val(it.row[table.header.indexOf(h)]) : ''); });
      return out;
    };
    function val(v) { return v == null ? '' : v; }

    // Exceptions first, then matches - that is the order a reviewer works in.
    var ORDER = { diff: 0, left: 1, right: 2, match: 3 };
    pairs.sort(function (a, b) { return ORDER[a.status] - ORDER[b.status]; });

    var stats = { match: 0, diff: 0, left: 0, right: 0, leftAmount: 0, rightAmount: 0 };
    var rows = pairs.map(function (p) {
      var li = p.l >= 0 ? L.items[p.l] : null, ri = p.r >= 0 ? Rt.items[p.r] : null;
      stats[p.status]++;
      if (li) stats.leftAmount = round2(stats.leftAmount + li.amount);
      if (ri) stats.rightAmount = round2(stats.rightAmount + ri.amount);
      var row = [STATUS[p.status], p.rule]
        .concat(side(left, L, li, L.keyIdx, exLeft))
        .concat(side(right, Rt, ri, Rt.keyIdx, exRight));
      row.push(round2((li ? li.amount : 0) - (ri ? ri.amount : 0)));
      if (L.dateIdx >= 0 && Rt.dateIdx >= 0) row.push(li && ri && li.date != null && ri.date != null ? Math.abs(li.date - ri.date) : '');
      return row;
    });

    return {
      header: header,
      types: types,
      rows: rows,
      statusLabels: STATUS,
      leftAmountCol: header.indexOf(LL + ' amount'),
      rightAmountCol: header.indexOf(RL + ' amount'),
      diffCol: header.indexOf('Difference'),
      stats: stats
    };
  };

  // ------------------------------------------------------------------ column guessing for templates
  /** First header whose normalised text contains one of the patterns (patterns tried in order). */
  R.findColumn = function (header, patterns, exclude) {
    var norm = function (s) { return String(s).toLowerCase().replace(/[^a-z0-9]/g, ''); };
    var skip = exclude || [];
    for (var p = 0; p < patterns.length; p++) {
      var pat = norm(patterns[p]);
      for (var i = 0; i < header.length; i++) {
        if (skip.indexOf(header[i]) >= 0) continue;
        if (norm(header[i]).indexOf(pat) >= 0) return header[i];
      }
    }
    return '';
  };

  global.AKRecon = R;
})(window);
