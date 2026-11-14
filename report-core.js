/* Report core: data -> neutral "doc model" -> HTML (print / PDF), Word (.docx), PowerPoint (.pptx).
   Works in the browser (globals docx / PptxGenJS) and in node (for tests). */
(function (root) {
'use strict';
var FONT_STACK = '"Noto Sans Khmer","Khmer UI","Leelawadee UI","Khmer OS Battambang",system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif';
var NOS = '(No section)', NOC = '(No category)', NOV = '(No village)';
var COLORS = { yes: '2F9E55', no: 'D9534F', resign: '6B7A8F', left: 'E9A23B', pend: 'A8D5AD', women: 'E0709B', other: '2F6F6A', done: '256B30' };

function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
function up(s) { return String(s == null ? '' : s).trim().toUpperCase(); }
function num(v) { var n = parseFloat(String(v == null ? '' : v).replace(/,/g, '')); return isNaN(n) ? 0 : n; }
function fmt(n) { return Number(n || 0).toLocaleString('en-US', { maximumFractionDigits: 2 }); }
function pct(a, b) { return b ? (Math.round(a * 1000 / b) / 10) + '%' : '0%'; }
function trunc(s, n) { s = String(s == null ? '' : s); return s.length > n ? s.slice(0, Math.max(1, n - 1)) + '…' : s; }
function isCE(sec) { return /^\s*CE\b/i.test(String(sec || '')); }

/* ============================ ICS ============================ */
function latestMap(arr) {
  var m = {};
  (arr || []).forEach(function (r) {
    var k = up(r.code), p = m[k];
    if (!p || String(r.date || '') >= String(p.date || '')) m[k] = r;
  });
  return m;
}
function cmp(r) { var c = String(r && r.compliance || '').trim().toLowerCase(); return (c === 'yes' || c === 'no' || c === 'resign') ? c : ''; }
function farmerName(f) {
  var h = f.husEn || f.husKh || '', w = f.wifeEn || f.wifeKh || '';
  if (/^(n\/?a|widower|none)$/i.test(String(w).trim())) w = '';
  return h + (w ? ' / ' + w : '');
}
function icsCompute(d, irpg) {
  var m1 = latestMap((d.ics1 || []).concat(d.nc || [])), m2 = latestMap(d.ics2), m3 = latestMap(d.ics3);
  var rows = (d.farmers || []).filter(function (f) { return !irpg || f.irpg === irpg; }).map(function (f) {
    var k = up(f.code), r1 = m1[k], r2 = m2[k], r3 = m3[k], c1 = cmp(r1), c2 = cmp(r2), c3 = cmp(r3), y1 = c1 === 'yes';
    return { f: f, r1: r1, r2: r2, r3: r3, c1: c1, c2: c2, c3: c3,
      done2: y1 && !!r2, yes2: y1 && c2 === 'yes', done3: y1 && c2 === 'yes' && !!r3, next: y1 && c2 === 'yes' && c3 === 'yes' };
  });
  function agg(list) {
    var a = { t: list.length, y: 0, n: 0, rs: 0, d2: 0, y2: 0, d3: 0, nx: 0 };
    list.forEach(function (r) {
      if (r.c1 === 'yes') a.y++; else if (r.c1 === 'no') a.n++; else if (r.c1 === 'resign') a.rs++;
      if (r.done2) a.d2++; if (r.yes2) a.y2++; if (r.done3) a.d3++; if (r.next) a.nx++;
    });
    a.left = a.t - a.y - a.n - a.rs; a.left2 = a.y - a.d2; a.left3 = a.y2 - a.d3;
    return a;
  }
  return { rows: rows, agg: agg, total: agg(rows) };
}
function groupRows(rows, keyFn) {
  var m = {}, out = [];
  rows.forEach(function (r) { var k = keyFn(r); if (!m[k]) { m[k] = { key: k, list: [] }; out.push(m[k]); } m[k].list.push(r); });
  out.sort(function (a, b) { return a.key.localeCompare(b.key, undefined, { numeric: true }); });
  return out;
}
function sortFarmers(a, b) {
  return (a.f.irpg + '|' + a.f.village).localeCompare(b.f.irpg + '|' + b.f.village, undefined, { numeric: true }) ||
         String(a.f.code).localeCompare(String(b.f.code), undefined, { numeric: true });
}

function buildIcsDoc(d, o) {
  var C = icsCompute(d, o.irpg), T = C.total, blocks = [];
  var head = ['Total', 'Yes', 'No', 'Resign', 'Not yet', 'ICS-II done', 'ICS-II left', 'ICS-III done', 'ICS-III left', 'Next year'];
  function line(a) { return [fmt(a.t), fmt(a.y), fmt(a.n), fmt(a.rs), fmt(a.left), fmt(a.d2), fmt(a.left2), fmt(a.d3), fmt(a.left3), fmt(a.nx)]; }
  var widths = [2.1, 0.8, 0.8, 0.8, 0.9, 1, 1, 1, 1, 0.9, 1];
  var align = ['l', 'r', 'r', 'r', 'r', 'r', 'r', 'r', 'r', 'r', 'r'];

  if (o.summary) {
    blocks.push({ t: 'h1', text: 'Summary' });
    blocks.push({ t: 'kpi', items: [
      { l: 'Total farmers', v: fmt(T.t) }, { l: 'ICS-I passed (Yes)', v: fmt(T.y) }, { l: 'No', v: fmt(T.n) },
      { l: 'Resign', v: fmt(T.rs) }, { l: 'Not yet inspected', v: fmt(T.left) } ] });
    blocks.push({ t: 'kpi', items: [
      { l: 'ICS-II done / to do', v: fmt(T.d2) + ' / ' + fmt(T.y) }, { l: 'ICS-III done / to do', v: fmt(T.d3) + ' / ' + fmt(T.y2) },
      { l: 'Continue next year (passed ICS-III)', v: fmt(T.nx) } ] });
    blocks.push({ t: 'chart', kind: 'pie', title: 'ICS-I result · Compliance', labels: ['Yes', 'No', 'Resign', 'Not yet inspected'],
      series: [{ name: 'Farmers', values: [T.y, T.n, T.rs, T.left] }], colors: [COLORS.yes, COLORS.no, COLORS.resign, COLORS.left], side: true });
    blocks.push({ t: 'chart', kind: 'hbar', stacked: true, title: 'Progress by stage (done / left)', labels: ['ICS-I', 'ICS-II', 'ICS-III'],
      series: [{ name: 'Done', values: [T.y + T.n + T.rs, T.d2, T.d3] }, { name: 'Left', values: [T.left, T.left2, T.left3] }], colors: [COLORS.done, COLORS.left], side: true });
  }
  if (o.byIrpg) {
    var gi = groupRows(C.rows, function (r) { return r.f.irpg || '-'; });
    var rowsI = gi.map(function (g) { return [g.key].concat(line(C.agg(g.list))); });
    rowsI.push(['Total'].concat(line(T)));
    blocks.push({ t: 'h1', text: 'By IRPG', pb: true });
    blocks.push({ t: 'table', head: ['IRPG'].concat(head), rows: rowsI, widths: widths, align: align, bold: [rowsI.length - 1] });
  }
  if (o.byVillage) {
    var gv = groupRows(C.rows, function (r) { return (r.f.irpg || '-') + ' · ' + (r.f.village || '-'); });
    var rowsV = gv.map(function (g) { return [g.key].concat(line(C.agg(g.list))); });
    rowsV.push(['Total'].concat(line(T)));
    blocks.push({ t: 'h1', text: 'By village', pb: !o.byIrpg });
    blocks.push({ t: 'table', head: ['IRPG · Village'].concat(head), rows: rowsV, widths: widths, align: align, bold: [rowsV.length - 1] });
  }
  function listBlock(title, list, note, withStatus) {
    if (!list.length) { blocks.push({ t: 'h1', text: title, pb: true, pptx: false }); blocks.push({ t: 'p', text: 'No data.', pptx: false }); return; }
    list.sort(sortFarmers);
    var hd = ['No.', 'Family_ID', 'Name (husband / wife)', 'IRPG', 'Village'];
    var w = [0.5, 1, 3, 0.9, 1.4];
    if (withStatus) { hd.push('Status', 'Date', 'Inspector', 'Note'); w.push(0.9, 1.1, 1.5, 2); }
    var rows = list.map(function (r, i) {
      var base = [String(i + 1), r.f.code, farmerName(r.f), r.f.irpg, r.f.village];
      if (withStatus) base.push(r.c1 === 'resign' ? 'Resign' : 'No', r.r1 && r.r1.date || '', r.r1 && r.r1.inspector || '', r.r1 && r.r1.extra || '');
      return base;
    });
    blocks.push({ t: 'h1', text: title + ' (' + fmt(list.length) + ')', pb: true, pptx: false });
    if (note) blocks.push({ t: 'p', text: note, pptx: false });
    blocks.push({ t: 'table', head: hd, rows: rows, widths: w, align: hd.map(function () { return 'l'; }), pptx: false });
  }
  if (o.lists.no) listBlock('Farmers with No / Resign', C.rows.filter(function (r) { return r.c1 === 'no' || r.c1 === 'resign'; }), '', true);
  if (o.lists.left1) listBlock('Farmers not yet inspected (ICS-I)', C.rows.filter(function (r) { return r.c1 === ''; }), '', false);
  if (o.lists.left2) listBlock('Farmers without ICS-II (passed ICS-I)', C.rows.filter(function (r) { return r.c1 === 'yes' && !r.r2; }), '', false);
  if (o.lists.left3) listBlock('Farmers without ICS-III (passed ICS-II)', C.rows.filter(function (r) { return r.yes2 && !r.r3; }), '', false);

  return { kind: 'ics', color: COLORS.done, landscape: true,
    title: 'ICS Monitoring Report ' + o.year, sub: o.irpg ? 'IRPG: ' + o.irpg : 'All IRPGs',
    meta: ['Data last synced: ' + (d.time ? new Date(d.time).toLocaleString('en-GB') : '-'), 'Generated: ' + new Date().toLocaleString('en-GB')],
    footer: 'ICS Monitoring ' + o.year, blocks: blocks };
}

/* ======================== FIELD WORK ======================== */
function splitNote(n) {
  n = String(n || '');
  var p = '', s = '', rest = n;
  var mp = n.match(/⚠\uFE0F?\s*បញ្ហា៖\s*([\s\S]*?)(?=\n\s*✅|$)/), ms = n.match(/✅\uFE0F?\s*ដំណោះស្រាយ៖\s*([\s\S]*)$/);
  if (mp) { p = mp[1].trim(); rest = rest.replace(mp[0], ''); }
  if (ms) { s = ms[1].trim(); rest = rest.replace(ms[0], ''); }
  return { p: p, s: s, note: rest.trim() };
}
function fldFilter(d, o) {
  return (d.rows || []).filter(function (r) {
    if (o.org === 'RP' && r.org && r.org !== 'RP') return false;
    if (o.org === 'IRCC') { if (r.org && r.org !== 'IRCC') return false; if (isCE(r.section)) return false; }
    if (o.sections && o.sections.indexOf(r.section || NOS) < 0) return false;
    return true;
  });
}
function fldTree(rows) {
  var secs = [], sm = {}, all = { times: 0, t: 0, w: 0, vills: {} };
  rows.forEach(function (r) {
    var sn = r.section || NOS, cn = r.category || NOC, an = r.activity, vn = String(r.village || '').trim() || NOV;
    var s = sm[sn]; if (!s) { s = sm[sn] = { name: sn, cats: [], cm: {}, times: 0, t: 0, w: 0, vills: {} }; secs.push(s); }
    var c = s.cm[cn]; if (!c) { c = s.cm[cn] = { name: cn, acts: [], am: {}, times: 0, t: 0, w: 0, vills: {} }; s.cats.push(c); }
    var a = c.am[an]; if (!a) { a = c.am[an] = { name: an, vs: [], vm: {}, times: 0, t: 0, w: 0, vills: {}, items: [] }; c.acts.push(a); }
    var v = a.vm[vn]; if (!v) { v = a.vm[vn] = { name: vn, times: 0, t: 0, w: 0 }; a.vs.push(v); }
    var n = num(r.count), t = num(r.total), w = num(r.women);
    [all, s, c, a, v].forEach(function (x) { x.times += n; x.t += t; x.w += w; });
    [all, s, c, a].forEach(function (x) { if (x.vills) x.vills[vn] = 1; });
    (r.items || []).forEach(function (i) { a.items.push({ d: i.d, s: i.s, t: i.t, w: i.w, n: i.n, v: vn, a: an }); });
  });
  return { secs: secs, all: all };
}
function shortLabel(s, n) { s = String(s || '').replace(/\s+/g, ' '); return trunc(s, n); }

function buildFieldDoc(d, o) {
  var rows = fldFilter(d, o), tree = fldTree(rows), blocks = [], A = tree.all;
  var nv = function (x) { return Object.keys(x.vills).length; };
  var org = o.org === 'RP' ? 'RP' : (o.org === 'IRCC' ? 'IRCC' : 'RP + IRCC');
  blocks.push({ t: 'h1', text: 'Summary' });
  blocks.push({ t: 'kpi', items: [{ l: 'Activity sessions', v: fmt(A.times) }, { l: 'Villages', v: fmt(nv(A)) },
    { l: 'Total participants', v: fmt(A.t) }, { l: 'Women', v: fmt(A.w) }, { l: '% Women', v: pct(A.w, A.t) }] });
  if (tree.secs.length) {
    blocks.push({ t: 'chart', kind: 'hbar', stacked: true, title: 'Participants by Section (women / others)', labels: tree.secs.map(function (s) { return shortLabel(s.name, 30); }),
      series: [{ name: 'Women', values: tree.secs.map(function (s) { return s.w; }) }, { name: 'Others', values: tree.secs.map(function (s) { return Math.max(0, s.t - s.w); }) }],
      colors: [COLORS.women, COLORS.other] });
    var rs = tree.secs.map(function (s) { return [s.name, fmt(s.times), fmt(nv(s)), fmt(s.t), fmt(s.w), pct(s.w, s.t)]; });
    rs.push(['Total', fmt(A.times), fmt(nv(A)), fmt(A.t), fmt(A.w), pct(A.w, A.t)]);
    blocks.push({ t: 'table', head: ['Section', 'Sessions', 'Villages', 'Total', 'Women', '% Women'], rows: rs, widths: [3, 1, 1, 1, 1, 1], align: ['l', 'r', 'r', 'r', 'r', 'r'], bold: [rs.length - 1] });
  } else {
    blocks.push({ t: 'p', text: 'No data for the selected options.' });
  }

  tree.secs.forEach(function (s) {
    blocks.push({ t: 'h1', text: 'Section: ' + s.name, pb: true });
    blocks.push({ t: 'kpi', items: [{ l: 'Sessions', v: fmt(s.times) }, { l: 'Villages', v: fmt(nv(s)) }, { l: 'Total participants', v: fmt(s.t) }, { l: 'Women', v: fmt(s.w) }, { l: '% Women', v: pct(s.w, s.t) }] });
    blocks.push({ t: 'chart', kind: 'hbar', stacked: true, title: 'Participants by Category', labels: s.cats.map(function (c) { return shortLabel(c.name, 34); }),
      series: [{ name: 'Women', values: s.cats.map(function (c) { return c.w; }) }, { name: 'Others', values: s.cats.map(function (c) { return Math.max(0, c.t - c.w); }) }], colors: [COLORS.women, COLORS.other] });
    var rc = s.cats.map(function (c) { return [c.name, fmt(c.acts.length), fmt(c.times), fmt(c.t), fmt(c.w)]; });
    blocks.push({ t: 'table', head: ['Category', 'Activities', 'Sessions', 'Total', 'Women'], rows: rc, widths: [5, 1, 1, 1, 1], align: ['l', 'r', 'r', 'r', 'r'] });

    s.cats.forEach(function (c) {
      blocks.push({ t: 'h2', text: c.name });
      var ra = [], bold = [];
      c.acts.forEach(function (a) {
        bold.push(ra.length);
        ra.push([a.name, fmt(a.times), fmt(a.t), fmt(a.w)]);
        a.vs.forEach(function (v) { ra.push(['   Village ' + v.name, fmt(v.times), fmt(v.t), fmt(v.w)]); });
      });
      blocks.push({ t: 'table', head: ['Activity / Village', 'Sessions', 'Total', 'Women'], rows: ra, widths: [6, 1, 1, 1], align: ['l', 'r', 'r', 'r'], bold: bold, pptMax: 36 });
      if (o.details) {
        var items = [];
        c.acts.forEach(function (a) { a.items.forEach(function (i) { items.push(i); }); });
        if (items.length) {
          items.sort(function (a, b) { return String(a.d).localeCompare(String(b.d)); });
          blocks.push({ t: 'h3', text: 'Detailed log · ' + shortLabel(c.name, 60), pptx: false });
          blocks.push({ t: 'table', head: ['Date', 'Staff', 'Activity', 'Village', 'Total', 'Women', 'Notes'],
            rows: items.map(function (i) { return [i.d || '', i.s || '', shortLabel(i.a, 70), i.v, fmt(i.t), fmt(i.w), i.n || '']; }),
            widths: [1.1, 1.3, 3, 1.3, 0.7, 0.7, 3], align: ['l', 'l', 'l', 'l', 'r', 'r', 'l'], pptx: false });
        }
      }
    });
    if (o.problems) {
      var ps = [];
      s.cats.forEach(function (c) { c.acts.forEach(function (a) { a.items.forEach(function (i) {
        var x = splitNote(i.n); if (x.p || x.s) ps.push([i.d || '', i.s || '', shortLabel(a.name, 50) + ' · ' + i.v, x.p, x.s]);
      }); }); });
      if (ps.length) {
        ps.sort(function (a, b) { return String(a[0]).localeCompare(String(b[0])); });
        blocks.push({ t: 'h2', text: 'Problems encountered and solutions · ' + s.name });
        blocks.push({ t: 'table', head: ['Date', 'Staff', 'Activity · Village', 'Problem', 'Solution'], rows: ps, widths: [1.1, 1.3, 2.6, 3, 3], align: ['l', 'l', 'l', 'l', 'l'], pptMax: 24 });
      }
    }
  });
  return { kind: 'fld', color: '2F6F6A', landscape: true,
    title: 'Field Work Minute Report', sub: (o.month ? o.month : 'All months') + ' · ' + org,
    meta: ['Generated: ' + new Date().toLocaleString('en-GB')], footer: 'Field Work Minute · ' + (o.month || 'All'), blocks: blocks };
}

/* ============================ CHARTS (SVG) ============================ */
function chartSvg(b, W) {
  W = W || 720;
  var font = 'font-family="' + FONT_STACK.replace(/"/g, "'") + '"';
  var cols = b.colors.map(function (c) { return '#' + c; });
  var out = [], H;
  if (b.kind === 'pie') {
    var vals = b.series[0].values, tot = vals.reduce(function (a, c) { return a + c; }, 0), R = 70, CIR = 2 * Math.PI * R, off = 0;
    H = Math.max(190, b.labels.length * 34 + 50);
    out.push('<circle cx="110" cy="' + (H / 2 + 8) + '" r="' + R + '" fill="none" stroke="#e3e8e1" stroke-width="34"/>');
    vals.forEach(function (v, i) {
      if (!v || !tot) return;
      var len = CIR * v / tot;
      out.push('<circle cx="110" cy="' + (H / 2 + 8) + '" r="' + R + '" fill="none" stroke="' + cols[i % cols.length] + '" stroke-width="34" stroke-dasharray="' + len.toFixed(2) + ' ' + (CIR - len).toFixed(2) + '" stroke-dashoffset="' + (-off).toFixed(2) + '" transform="rotate(-90 110 ' + (H / 2 + 8) + ')"/>');
      off += len;
    });
    out.push('<text x="110" y="' + (H / 2 + 14) + '" text-anchor="middle" font-size="20" font-weight="700" fill="#16221a" ' + font + '>' + esc(fmt(tot)) + '</text>');
    b.labels.forEach(function (l, i) {
      var y = 52 + i * 34;
      out.push('<rect x="250" y="' + (y - 13) + '" width="16" height="16" rx="4" fill="' + cols[i % cols.length] + '"/>');
      out.push('<text x="274" y="' + y + '" font-size="16" fill="#16221a" ' + font + '>' + esc(l) + '</text>');
      out.push('<text x="' + (W - 20) + '" y="' + y + '" font-size="16" font-weight="700" text-anchor="end" fill="#16221a" ' + font + '>' + esc(fmt(vals[i]) + ' (' + pct(vals[i], tot) + ')') + '</text>');
    });
  } else {
    var n = b.labels.length, rowH = 34, top = 44, left = 250, right = 70, bw = W - left - right;
    H = top + n * rowH + 14;
    var sums = b.labels.map(function (l, i) { return b.series.reduce(function (a, s) { return a + (s.values[i] || 0); }, 0); });
    var max = Math.max.apply(null, sums.concat([1]));
    b.series.forEach(function (s, k) {
      var x = left + k * 130;
      out.push('<rect x="' + x + '" y="10" width="14" height="14" rx="3" fill="' + cols[k % cols.length] + '"/><text x="' + (x + 20) + '" y="22" font-size="14" fill="#16221a" ' + font + '>' + esc(s.name) + '</text>');
    });
    b.labels.forEach(function (l, i) {
      var y = top + i * rowH, x = left;
      out.push('<text x="' + (left - 10) + '" y="' + (y + 20) + '" font-size="14" text-anchor="end" fill="#16221a" ' + font + '>' + esc(trunc(l, 32)) + '</text>');
      b.series.forEach(function (s, k) {
        var v = s.values[i] || 0, w = bw * v / max;
        if (!v) return;
        out.push('<rect x="' + x.toFixed(1) + '" y="' + y + '" width="' + w.toFixed(1) + '" height="24" fill="' + cols[k % cols.length] + '"/>');
        if (w > 30) out.push('<text x="' + (x + w / 2).toFixed(1) + '" y="' + (y + 17) + '" font-size="13" font-weight="700" text-anchor="middle" fill="#fff" ' + font + '>' + esc(fmt(v)) + '</text>');
        x += w;
      });
      out.push('<text x="' + (x + 6).toFixed(1) + '" y="' + (y + 18) + '" font-size="13" fill="#5b675f" ' + font + '>' + esc(fmt(sums[i])) + '</text>');
    });
  }
  return { svg: '<svg xmlns="http://www.w3.org/2000/svg" width="' + W + '" height="' + H + '" viewBox="0 0 ' + W + ' ' + H + '">' + out.join('') + '</svg>', w: W, h: H };
}

/* ============================ HTML (preview / PDF) ============================ */
function docToHtml(doc, opt) {
  opt = opt || {};
  var c = '#' + doc.color;
  var css = '@page{size:A4 ' + (doc.landscape ? 'landscape' : 'portrait') + ';margin:12mm}' +
    '*{box-sizing:border-box}body{font-family:' + FONT_STACK + ';color:#16221a;font-size:11pt;margin:0;padding:0;-webkit-print-color-adjust:exact;print-color-adjust:exact}' +
    '.cover{border-left:10px solid ' + c + ';padding:6px 0 6px 16px;margin:0 0 14px}.cover h1{margin:0;font-size:24pt;color:' + c + '}.cover .sub{font-size:14pt;font-weight:600;margin-top:4px}.cover .meta{color:#5b675f;font-size:10pt;margin-top:4px}' +
    'h1.s{font-size:17pt;color:' + c + ';border-bottom:2px solid ' + c + ';padding-bottom:3px;margin:18px 0 10px;break-after:avoid}h1.s.pb{break-before:page;margin-top:0}' +
    'h2{font-size:13pt;margin:16px 0 6px;color:#16221a;break-after:avoid}h3{font-size:11.5pt;margin:12px 0 4px;color:#3b4940;break-after:avoid}' +
    'p{margin:6px 0}.note{color:#5b675f;font-size:10pt}' +
    '.kpi{display:flex;gap:10px;margin:10px 0;flex-wrap:wrap}.kpi div{flex:1;min-width:110px;border:1px solid #d9e0d6;background:#f3f6f1;border-radius:8px;padding:8px 10px;break-inside:avoid}.kpi b{display:block;font-size:19pt;color:' + c + '}.kpi span{font-size:9.5pt;color:#5b675f}' +
    'table{border-collapse:collapse;width:100%;margin:8px 0 12px;font-size:9.5pt}thead{display:table-header-group}th{background:' + c + ';color:#fff;text-align:left;padding:5px 7px;border:1px solid ' + c + '}td{border:1px solid #d9e0d6;padding:4px 7px;vertical-align:top}tr{break-inside:avoid}tbody tr:nth-child(even) td{background:#f7f9f5}tr.b td{background:#e4f1e5!important;font-weight:700}td.r,th.r{text-align:right}' +
    '.chart{margin:8px 0;break-inside:avoid}.chart .ct{font-weight:700;margin-bottom:2px}.chart svg{max-width:100%;height:auto}' +
    '.foot{margin-top:18px;color:#7a857d;font-size:9pt;border-top:1px solid #d9e0d6;padding-top:4px}';
  var h = ['<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>' + esc(doc.title) + '</title>',
    '<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Noto+Sans+Khmer:wght@400;600;700&display=swap"><style>' + css + '</style></head><body>',
    '<div class="cover"><h1>' + esc(doc.title) + '</h1><div class="sub">' + esc(doc.sub) + '</div>' + (doc.meta || []).map(function (m) { return '<div class="meta">' + esc(m) + '</div>'; }).join('') + '</div>'];
  doc.blocks.forEach(function (b) {
    if (b.t === 'h1') h.push('<h1 class="s' + (b.pb ? ' pb' : '') + '">' + esc(b.text) + '</h1>');
    else if (b.t === 'h2') h.push('<h2>' + esc(b.text) + '</h2>');
    else if (b.t === 'h3') h.push('<h3>' + esc(b.text) + '</h3>');
    else if (b.t === 'p') h.push('<p class="note">' + esc(b.text) + '</p>');
    else if (b.t === 'kpi') h.push('<div class="kpi">' + b.items.map(function (i) { return '<div><b>' + esc(i.v) + '</b><span>' + esc(i.l) + '</span></div>'; }).join('') + '</div>');
    else if (b.t === 'chart') { var cs = chartSvg(b, 720); h.push('<div class="chart"><div class="ct">' + esc(b.title || '') + '</div>' + cs.svg + '</div>'); }
    else if (b.t === 'table') {
      var tw = b.widths || b.head.map(function () { return 1; }), ts = tw.reduce(function (a, x) { return a + x; }, 0);
      h.push('<table><thead><tr>' + b.head.map(function (x, i) { return '<th class="' + (b.align && b.align[i] === 'r' ? 'r' : '') + '" style="width:' + (tw[i] * 100 / ts).toFixed(1) + '%">' + esc(x) + '</th>'; }).join('') + '</tr></thead><tbody>' +
        b.rows.map(function (r, ri) { return '<tr' + ((b.bold || []).indexOf(ri) > -1 ? ' class="b"' : '') + '>' + r.map(function (x, i) { return '<td class="' + (b.align && b.align[i] === 'r' ? 'r' : '') + '">' + esc(x).replace(/\n/g, '<br>').replace(/^ +/, function (m) { return '&nbsp;'.repeat(m.length); }) + '</td>'; }).join('') + '</tr>'; }).join('') + '</tbody></table>');
    }
  });
  h.push('<div class="foot">' + esc(doc.footer || '') + ' · ' + esc(new Date().toLocaleDateString('en-GB')) + '</div>');
  if (opt.autoPrint) h.push('<script>window.addEventListener("load",function(){setTimeout(function(){(document.fonts&&document.fonts.ready||Promise.resolve()).then(function(){window.print()})},400)})<\/script>');
  h.push('</body></html>');
  return h.join('');
}

/* ============================ WORD (.docx) ============================ */
function chunk(arr, n) { var o = []; for (var i = 0; i < arr.length; i += n) o.push(arr.slice(i, i + n)); return o; }
async function docToDocx(doc, env) {
  var D = root.docx;
  var font = { ascii: 'Calibri', hAnsi: 'Calibri', cs: 'Khmer UI', eastAsia: 'Khmer UI' };
  var CC = doc.color, PAGE_W = 11906, PAGE_H = 16838, MARGIN = 850;
  var contentW = (doc.landscape ? PAGE_H : PAGE_W) - MARGIN * 2;
  function run(text, o) {
    o = o || {};
    return new D.TextRun({ text: String(text == null ? '' : text), font: font, size: o.size || 20, bold: !!o.bold, boldComplexScript: !!o.bold, sizeComplexScript: o.size || 20, color: o.color, italics: !!o.italics });
  }
  function para(text, o) {
    o = o || {};
    return new D.Paragraph({ alignment: o.align, spacing: { before: o.before || 0, after: o.after == null ? 80 : o.after }, keepNext: o.keepNext, children: [run(text, o)] });
  }
  var BORDER = { style: D.BorderStyle.SINGLE, size: 4, color: 'D9E0D6' };
  var borders = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER, insideHorizontal: BORDER, insideVertical: BORDER };
  function widthsDxa(ratios) {
    var sum = ratios.reduce(function (a, x) { return a + x; }, 0), w = ratios.map(function (r) { return Math.floor(contentW * r / sum); });
    w[w.length - 1] += contentW - w.reduce(function (a, x) { return a + x; }, 0);
    return w;
  }
  function cell(text, w, o) {
    o = o || {};
    var lines = String(text == null ? '' : text).split('\n');
    return new D.TableCell({
      width: { size: w, type: D.WidthType.DXA },
      shading: o.fill ? { type: D.ShadingType.CLEAR, fill: o.fill, color: 'auto' } : undefined,
      margins: { top: 50, bottom: 50, left: 80, right: 80 }, verticalAlign: D.VerticalAlign.CENTER,
      children: lines.map(function (l) { return new D.Paragraph({ alignment: o.align === 'r' ? D.AlignmentType.RIGHT : D.AlignmentType.LEFT, spacing: { before: 0, after: 0 }, children: [run(l, { size: o.size || 18, bold: o.bold, color: o.color })] }); })
    });
  }
  var children = [];
  children.push(new D.Paragraph({ spacing: { after: 60 }, children: [run(doc.title, { size: 44, bold: true, color: CC })] }));
  children.push(para(doc.sub, { size: 28, bold: true, after: 40 }));
  (doc.meta || []).forEach(function (m) { children.push(para(m, { size: 18, color: '5B675F', after: 20 })); });
  children.push(new D.Paragraph({ border: { bottom: { style: D.BorderStyle.SINGLE, size: 12, color: CC, space: 1 } }, spacing: { after: 160 }, children: [] }));

  for (var bi = 0; bi < doc.blocks.length; bi++) {
    var b = doc.blocks[bi];
    if (b.t === 'h1') {
      children.push(new D.Paragraph({ heading: D.HeadingLevel.HEADING_1, pageBreakBefore: !!b.pb, keepNext: true, spacing: { before: b.pb ? 0 : 280, after: 120 },
        border: { bottom: { style: D.BorderStyle.SINGLE, size: 8, color: CC, space: 2 } }, children: [run(b.text, { size: 34, bold: true, color: CC })] }));
    } else if (b.t === 'h2') {
      children.push(new D.Paragraph({ heading: D.HeadingLevel.HEADING_2, keepNext: true, spacing: { before: 240, after: 100 }, children: [run(b.text, { size: 26, bold: true })] }));
    } else if (b.t === 'h3') {
      children.push(new D.Paragraph({ heading: D.HeadingLevel.HEADING_3, keepNext: true, spacing: { before: 160, after: 80 }, children: [run(b.text, { size: 22, bold: true, color: '3B4940' })] }));
    } else if (b.t === 'p') {
      children.push(para(b.text, { size: 19, color: '5B675F' }));
    } else if (b.t === 'kpi') {
      var kw = widthsDxa(b.items.map(function () { return 1; }));
      children.push(new D.Table({ width: { size: contentW, type: D.WidthType.DXA }, columnWidths: kw, borders: borders,
        rows: [new D.TableRow({ cantSplit: true, children: b.items.map(function (it, i) {
          return new D.TableCell({ width: { size: kw[i], type: D.WidthType.DXA }, shading: { type: D.ShadingType.CLEAR, fill: 'F3F6F1', color: 'auto' }, margins: { top: 80, bottom: 80, left: 100, right: 100 },
            children: [new D.Paragraph({ spacing: { after: 0 }, children: [run(it.v, { size: 36, bold: true, color: CC })] }), new D.Paragraph({ spacing: { after: 0 }, children: [run(it.l, { size: 17, color: '5B675F' })] })] });
        }) })] }));
      children.push(new D.Paragraph({ spacing: { after: 100 }, children: [] }));
    } else if (b.t === 'chart') {
      if (env && env.svgToPng) {
        var cs = chartSvg(b, 720), png = await env.svgToPng(cs.svg, cs.w, cs.h);
        if (png) {
          if (b.title) children.push(para(b.title, { size: 20, bold: true, keepNext: true, after: 40 }));
          children.push(new D.Paragraph({ spacing: { after: 120 }, children: [new D.ImageRun({ type: 'png', data: png, transformation: { width: cs.w * 0.8, height: cs.h * 0.8 }, altText: { title: b.title || 'chart', description: b.title || 'chart', name: 'chart' } })] }));
        }
      }
    } else if (b.t === 'table') {
      var tw = widthsDxa(b.widths || b.head.map(function () { return 1; })), isB = b.bold || [];
      var rowsX = [new D.TableRow({ tableHeader: true, cantSplit: true, children: b.head.map(function (h, i) { return cell(h, tw[i], { fill: CC, color: 'FFFFFF', bold: true, align: b.align && b.align[i], size: 18 }); }) })];
      b.rows.forEach(function (r, ri) {
        var fill = isB.indexOf(ri) > -1 ? 'E4F1E5' : (ri % 2 ? 'F7F9F5' : undefined);
        rowsX.push(new D.TableRow({ cantSplit: true, children: r.map(function (x, i) { return cell(x, tw[i], { fill: fill, bold: isB.indexOf(ri) > -1, align: b.align && b.align[i], size: 17 }); }) }));
      });
      children.push(new D.Table({ width: { size: contentW, type: D.WidthType.DXA }, columnWidths: tw, borders: borders, rows: rowsX }));
      children.push(new D.Paragraph({ spacing: { after: 120 }, children: [] }));
    }
  }
  var document = new D.Document({
    creator: 'Field Hub', title: doc.title,
    styles: { default: {
      document: { run: { font: font, size: 20, sizeComplexScript: 20 } },
      heading1: { run: { font: font, size: 34, bold: true, color: CC }, paragraph: { spacing: { before: 280, after: 120 }, keepNext: true } },
      heading2: { run: { font: font, size: 26, bold: true }, paragraph: { spacing: { before: 240, after: 100 }, keepNext: true } },
      heading3: { run: { font: font, size: 22, bold: true }, paragraph: { spacing: { before: 160, after: 80 }, keepNext: true } } } },
    sections: [{
      properties: { page: { size: { width: PAGE_W, height: PAGE_H, orientation: doc.landscape ? D.PageOrientation.LANDSCAPE : D.PageOrientation.PORTRAIT }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } },
      footers: { default: new D.Footer({ children: [new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [run((doc.footer || '') + ' · ', { size: 16, color: '7A857D' }), new D.TextRun({ children: [D.PageNumber.CURRENT], size: 16, color: '7A857D', font: font })] })] }) },
      children: children }]
  });
  return D.Packer.toBlob ? D.Packer.toBlob(document) : D.Packer.toBuffer(document);
}

/* ============================ POWERPOINT (.pptx) ============================ */
async function docToPptx(doc) {
  var P = new root.PptxGenJS();
  P.layout = 'LAYOUT_WIDE'; P.title = doc.title; P.author = 'Field Hub';
  var W = 13.333, H = 7.5, M = 0.5, C = doc.color, F = 'Leelawadee UI', slide = null, y = 0, base = '', cont = 0;
  function frame(t) {
    slide = P.addSlide();
    slide.addShape(P.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.95, fill: { color: C }, line: { color: C, width: 0 } });
    slide.addText(t, { x: M, y: 0.1, w: W - 2 * M, h: 0.75, fontFace: F, fontSize: 24, bold: true, color: 'FFFFFF', valign: 'middle', fit: 'shrink', isTextBox: true });
    slide.addText(doc.footer || '', { x: M, y: H - 0.42, w: 8, h: 0.3, fontFace: F, fontSize: 10, color: '7A857D', isTextBox: true });
    slide.slideNumber = { x: W - 1.2, y: H - 0.42, w: 0.7, h: 0.3, fontFace: F, fontSize: 10, color: '7A857D', align: 'right' };
    y = 1.2;
  }
  function start(t) { base = t; cont = 0; frame(t); }
  function more() { cont++; frame(base + ' (cont.)'); }
  function need(h) { if (y + h > H - 0.55) more(); }

  // title slide
  var s0 = P.addSlide(); s0.background = { color: C };
  s0.addText(doc.title, { x: 0.8, y: 2.0, w: W - 1.6, h: 1.3, fontFace: F, fontSize: 38, bold: true, color: 'FFFFFF', valign: 'bottom', fit: 'shrink', isTextBox: true });
  s0.addText(doc.sub, { x: 0.8, y: 3.4, w: W - 1.6, h: 0.7, fontFace: F, fontSize: 22, color: 'FFFFFF', isTextBox: true });
  s0.addText((doc.meta || []).join('\n'), { x: 0.8, y: 4.3, w: W - 1.6, h: 0.9, fontFace: F, fontSize: 13, color: 'DCEADF', valign: 'top', isTextBox: true });

  function chartOpts(b, x, w, h, half) {
    return b.kind === 'pie'
      ? { x: x, y: y, w: w, h: h, holeSize: 55, showLegend: true, legendPos: half ? 'b' : 'r', legendFontSize: 11, showPercent: false, showValue: true,
          dataLabelColor: 'FFFFFF', dataLabelFontSize: 11, chartColors: b.colors, showTitle: true, title: b.title, titleFontSize: 13, titleColor: '16221A' }
      : { x: x, y: y, w: w, h: h, barDir: 'bar', barGrouping: b.stacked ? 'stacked' : 'clustered', chartColors: b.colors,
          showLegend: b.series.length > 1, legendPos: 'b', legendFontSize: 11, showValue: true, dataLabelColor: 'FFFFFF', dataLabelFontSize: 10, dataLabelPosition: b.stacked ? 'ctr' : 'outEnd',
          catAxisLabelFontSize: 11, valAxisLabelFontSize: 10, catAxisOrientation: 'maxMin', valGridLine: { color: 'E3E8E1', size: 0.5 }, catGridLine: { style: 'none' },
          showTitle: true, title: b.title, titleFontSize: 13, titleColor: '16221A' };
  }
  function chartData(b) { return b.series.map(function (s) { return { name: s.name, labels: b.labels.map(function (l) { return trunc(l, 40); }), values: s.values }; }); }
  var list = doc.blocks.filter(function (b) { return b.pptx !== false; });
  for (var bi = 0; bi < list.length; bi++) { (function (b, nextB) {
    if (b.t === 'h1') { start(b.text); return; }
    if (!slide) start(doc.title);
    if (b.t === 'h2') { need(1.7); slide.addText(b.text, { x: M, y: y, w: W - 2 * M, h: 0.45, fontFace: F, fontSize: 16, bold: true, color: C, fit: 'shrink', isTextBox: true }); y += 0.5; return; }
    if (b.t === 'h3') return;
    if (b.t === 'p') { var ph = 0.32 * Math.ceil(String(b.text).length / 120); need(ph); slide.addText(b.text, { x: M, y: y, w: W - 2 * M, h: ph, fontFace: F, fontSize: 12, color: '5B675F', valign: 'top', isTextBox: true }); y += ph + 0.1; return; }
    if (b.t === 'kpi') {
      need(1.4);
      var n = b.items.length, gap = 0.15, cw = (W - 2 * M - gap * (n - 1)) / n;
      b.items.forEach(function (it, i) {
        var x = M + i * (cw + gap);
        slide.addShape(P.ShapeType.roundRect, { x: x, y: y, w: cw, h: 1.15, fill: { color: 'F3F6F1' }, line: { color: 'D9E0D6', width: 1 }, rectRadius: 0.08 });
        slide.addText(String(it.v), { x: x + 0.1, y: y + 0.08, w: cw - 0.2, h: 0.6, fontFace: F, fontSize: 26, bold: true, color: C, fit: 'shrink', isTextBox: true });
        slide.addText(it.l, { x: x + 0.1, y: y + 0.68, w: cw - 0.2, h: 0.42, fontFace: F, fontSize: 11, color: '5B675F', valign: 'top', fit: 'shrink', isTextBox: true });
      });
      y += 1.35; return;
    }
    if (b.t === 'chart') {
      if (b._done) return;
      var pair = b.side && nextB && nextB.t === 'chart' && nextB.side ? nextB : null;
      if (pair) {
        var ph = Math.min(3.4, H - 0.55 - y);
        if (ph < 2.6) { more(); ph = Math.min(3.4, H - 0.55 - y); }
        var half = (W - 2 * M - 0.2) / 2;
        slide.addChart(b.kind === 'pie' ? P.charts.DOUGHNUT : P.charts.BAR, chartData(b), chartOpts(b, M, half, ph, true));
        slide.addChart(pair.kind === 'pie' ? P.charts.DOUGHNUT : P.charts.BAR, chartData(pair), chartOpts(pair, M + half + 0.2, half, ph, true));
        pair._done = true; y += ph + 0.15; return;
      }
      var ch = b.kind === 'pie' ? 3.3 : Math.min(3.5, Math.max(2.2, 0.55 + b.labels.length * 0.42));
      need(ch + 0.1);
      slide.addChart(b.kind === 'pie' ? P.charts.DOUGHNUT : P.charts.BAR, chartData(b), chartOpts(b, b.kind === 'pie' ? M + 1.5 : M, b.kind === 'pie' ? W - 2 * M - 3 : W - 2 * M, ch, false));
      y += ch + 0.15; return;
    }
    if (b.t === 'table') {
      var rowsAll = b.rows.slice(), extra = 0;
      if (b.pptMax && rowsAll.length > b.pptMax) { extra = rowsAll.length - b.pptMax; rowsAll = rowsAll.slice(0, b.pptMax); }
      var tw = b.widths || b.head.map(function () { return 1; }), ts = tw.reduce(function (a, x) { return a + x; }, 0), full = W - 2 * M;
      var colW = tw.map(function (x) { return +(full * x / ts).toFixed(3); });
      var small = rowsAll.length <= 8, fs = b.head.length > 8 ? (small ? 11 : 9) : (small ? 12 : 10), RH = small ? 0.42 : 0.3, idx = 0;
      var cut = function (t, i) { return trunc(String(t).replace(/\n/g, ' '), Math.max(6, Math.floor(colW[i] * (fs >= 10 ? 10 : 11)))); };
      while (idx < rowsAll.length) {
        var avail = H - 0.55 - y;
        var cap = Math.floor((avail - 0.4) / RH);
        if (cap < 3) { more(); continue; }
        var part = rowsAll.slice(idx, idx + cap);
        var tbl = [b.head.map(function (h, i) { return { text: trunc(h, 30), options: { bold: true, color: 'FFFFFF', fill: { color: C }, align: b.align && b.align[i] === 'r' ? 'right' : 'left', fontFace: F, fontSize: fs } }; })];
        part.forEach(function (r, k) {
          var gi = idx + k, isB = (b.bold || []).indexOf(gi) > -1, fill = isB ? 'E4F1E5' : (gi % 2 ? 'F7F9F5' : 'FFFFFF');
          tbl.push(r.map(function (x, i) { return { text: cut(x, i), options: { bold: isB, color: '16221A', fill: { color: fill }, align: b.align && b.align[i] === 'r' ? 'right' : 'left', fontFace: F, fontSize: fs } }; }));
        });
        slide.addTable(tbl, { x: M, y: y, w: full, colW: colW, rowH: RH, border: { type: 'solid', pt: 0.5, color: 'D9E0D6' }, valign: 'middle' });
        y += RH * (part.length + 1) + 0.15; idx += part.length;
        if (idx < rowsAll.length) more();
      }
      if (extra) { need(0.4); slide.addText('… and ' + extra + ' more rows (see the Word / PDF report)', { x: M, y: y, w: W - 2 * M, h: 0.3, fontFace: F, fontSize: 11, italic: true, color: '7A857D', isTextBox: true }); y += 0.4; }
    }
  })(list[bi], list[bi + 1]); }
  return P.write({ outputType: typeof window === 'undefined' ? 'nodebuffer' : 'blob' });
}

root.ReportCore = { buildIcsDoc: buildIcsDoc, buildFieldDoc: buildFieldDoc, docToHtml: docToHtml, docToDocx: docToDocx, docToPptx: docToPptx,
  NOS: NOS, NOC: NOC, NOV: NOV, icsCompute: icsCompute, fldFilter: fldFilter, fldTree: fldTree, splitNote: splitNote, chartSvg: chartSvg, isCE: isCE, latestMap: latestMap, cmp: cmp };
})(typeof window !== 'undefined' ? window : global);
