/* ERP Implementation Dashboard: data loading, charts, tables and forecast */
(function () {
  'use strict';

  /* ▼▼▼ Paste your Apps Script web app link (ends in /exec) between the quotes ▼▼▼ */
  var DATA_URL = 'https://script.google.com/macros/s/AKfycbz_9F9Z6H4kuz8z6UtmVD3FYXNnSCEWEfMzCJm0k7YnuB99synaewrH7geej-i-yBkk/exec';
  /* ▲▲▲ ------------------------------------------------------------------- ▲▲▲ */

  var REFRESH_MS = 5 * 60 * 1000;          // the page re-reads the sheet every 5 minutes
  var S = { tab: 'summary', model: null, sel: 0, picked: false, sort: 'no', q: '', view: 'cards' };
  try { S.tkDates = localStorage.getItem('erp.tkDates') === '1'; } catch (err) {}      // remembered Show Date & Week No choice

  /* Which column feeds which figure. Matched on the header text, so column order can change. */
  var DEFS = {
    ba:    { tab: 'BA',      re: /^ba$/ },
    baPct: { tab: 'BA',      re: /^ba ?%$/ },
    scope: { tab: 'Scope', alt: 'BA', re: /^scope doc$/ },      // read from the Scope tab; falls back to BA if that tab does not exist
    appr:  { tab: 'Scope', alt: 'BA', re: /approval/ },
    avail: { tab: 'AFI',     re: /avail/ },
    func:  { tab: 'AFI',     re: /function/ },
    inUse: { tab: 'AFI',     re: /in ?use/ },
    rExp:  { tab: 'Reports', re: /expect/ },
    rDone: { tab: 'Reports', re: /done|complete/ },
    iAll:  { tab: 'Issues',  re: /issue/ },
    iDone: { tab: 'Issues',  re: /done|resolved|closed/ }
  };

  var $ = function (id) { return document.getElementById(id); };
  var norm = function (s) { return String(s || '').toLowerCase().replace(/\s+/g, ' ').trim(); };
  var esc = function (s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  };
  var isNum = function (v) { return typeof v === 'number' && isFinite(v); };
  var isYes = function (v) { return v === 1 || /^(y|yes|true|done|completed?|approved)$/i.test(String(v === null ? '' : v).trim()); };
  var fmt = function (v, d) { return isNum(v) ? v.toLocaleString('en-US', { maximumFractionDigits: d || 0 }) : '–'; };
  var pct = function (v) { return isNum(v) ? fmt(v) + '%' : '–'; };
  var MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  function niceDate(s, withYear) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || '');
    if (!m) return s || '';
    return Number(m[3]) + ' ' + MONTHS[Number(m[2]) - 1] + (withYear ? ' ' + m[1] : '');
  }

  /* Week period: the week's date to 6 days later, e.g. "31 Oct – 6 Nov 2026". */
  function weekRange(s) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || '');
    if (!m) return s || '';
    var a = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])), b = new Date(a.getFullYear(), a.getMonth(), a.getDate() + 6);
    var d = function (x, y) { return x.getDate() + ' ' + MONTHS[x.getMonth()] + (y ? ' ' + x.getFullYear() : ''); };
    return d(a, a.getFullYear() !== b.getFullYear()) + ' – ' + d(b, true);
  }

  /* ================= model ================= */
  function build(data) {
    var tabs = data.tabs || {}, order = ['BA', 'AFI', 'Reports', 'Issues', 'Scope'];
    
// Main Drive folder that holds one folder per department (copy the ID from its link: .../folders/THIS_PART)
var DOCS_FOLDER_ID = 'https://drive.google.com/drive/folders/1libn_1eQ7GxQw9S7YW3rRmz4Q_pfwouI?usp=sharing';

// Which file in a department folder is which document (matched on the file name)
var DOC_TYPES = { ba: /\bBA\b/i, scope: /scope/i };

/** { 'finance department': { ba: {id, name, url}, scope: {...} }, ... } */
function getDocs_() {
  var out = {};
  if (!DOCS_FOLDER_ID || /PASTE/.test(DOCS_FOLDER_ID)) return out;
  var folders = DriveApp.getFolderById(DOCS_FOLDER_ID).getFolders();
  while (folders.hasNext()) {
    var f = folders.next(), docs = {}, files = f.getFiles();
    while (files.hasNext()) {
      var file = files.next(), name = file.getName();
      Object.keys(DOC_TYPES).forEach(function (t) {
        if (!docs[t] && DOC_TYPES[t].test(name)) docs[t] = { id: file.getId(), name: name, url: file.getUrl() };
      });
    }
    out[clean_(f.getName()).toLowerCase()] = docs;
  }
  return out;
}

/** Same as getDocs_, but remembered for 10 minutes so Drive is not scanned on every refresh. */
function getDocsCached_() {
  var cache = CacheService.getScriptCache(), hit = cache.get('docs');
  if (hit) return JSON.parse(hit);
  var docs = getDocs_();
  try { cache.put('docs', JSON.stringify(docs), 600); } catch (e) {}
  return docs;
}
    var fixed = {};                                 // figures from a tab with no week blocks (Scope): same value every week, not tracked weekly
    var depts = [], index = {}, nWeeks = 0, weeks = [];

    order.forEach(function (t) {
      if (!tabs[t]) return;
      tabs[t].departments.forEach(function (d) {
        var k = norm(d);
        if (!(k in index)) { index[k] = depts.length; depts.push(d); }
      });
      nWeeks = Math.max(nWeeks, tabs[t].weeks.length);
    });
    for (var w = 0; w < nWeeks; w++) {
      var info = { label: 'Week ' + (w + 1), date: '' };
      for (var i = 0; i < order.length; i++) {
        var tw = tabs[order[i]] && tabs[order[i]].weeks[w];
        if (tw) { info.label = tw.label || info.label; info.date = tw.date || ''; break; }
      }
      info.short = 'W' + (w + 1);
      weeks.push(info);
    }

    var M = {};
    Object.keys(DEFS).forEach(function (key) {
      var def = DEFS[key], tab = tabs[def.tab] || (def.alt ? tabs[def.alt] : null);
      M[key] = weeks.map(function () { return depts.map(function () { return null; }); });
      if (!tab) return;
      tab.weeks.forEach(function (tw, w) {
        var metric = null;
        tw.metrics.forEach(function (m) { if (!metric && def.re.test(norm(m.name))) metric = m; });
        if (!metric) return;
        tab.departments.forEach(function (d, j) { M[key][w][index[norm(d)]] = metric.values[j]; });
      });
      if (tab.weekly === false && weeks.length) {
        fixed[key] = true;
        for (var w = 1; w < weeks.length; w++) M[key][w] = M[key][0].slice();
      }
    });

    var model = { depts: depts, weeks: weeks, M: M, docs: data.docs || {}, missing: data.missing || [], generatedAt: data.generatedAt };
    var has = function (key, w) { return M[key][w].some(function (v) { return v !== null; }); };
    var nums = function (key, w) { return M[key][w].filter(isNum); };
    var sum = function (key, w) { return nums(key, w).reduce(function (a, b) { return a + b; }, 0); };
    var avg = function (key, w) { var n = nums(key, w); return n.length ? sum(key, w) / n.length : null; };
    var when = function (key, fn) { return weeks.map(function (_, w) { return has(key, w) ? fn(w) : null; }); };

    model.has = has;
    /* One value per week for every headline figure; null = that week is not filled in yet. */
    model.T = {
      baAvg:  when('baPct', function (w) { return avg('baPct', w); }),
      baDone: when('baPct', function (w) { return nums('baPct', w).filter(function (v) { return v >= 100; }).length; }),
      scope:  when('scope', function (w) { return M.scope[w].filter(isYes).length; }),
      appr:   when('appr',  function (w) { return M.appr[w].filter(isYes).length; }),
      avail:  when('avail', function (w) { return avg('avail', w); }),
      func:   when('func',  function (w) { return avg('func', w); }),
      inUse:  when('inUse', function (w) { return avg('inUse', w); }),
      rExp:   when('rExp',  function (w) { return sum('rExp', w); }),
      rDone:  when('rExp',  function (w) { return sum('rDone', w); }),
      iAll:   when('iAll',  function (w) { return sum('iAll', w); }),
      iDone:  when('iAll',  function (w) { return sum('iDone', w); })
    };
    var T = model.T;
    T.rPct  = T.rExp.map(function (e, w) { return e ? T.rDone[w] / e * 100 : null; });
    T.iOpen = T.iAll.map(function (a, w) { return a === null ? null : a - T.iDone[w]; });
    T.iPct  = T.iAll.map(function (a, w) { return a ? T.iDone[w] / a * 100 : null; });
    var share = function (s) { return s.map(function (v) { return v === null || !depts.length ? null : v / depts.length * 100; }); };
    T.baYes = when('ba', function (w) { return M.ba[w].filter(isYes).length; });
    T.baYesPct = share(T.baYes);
    T.scopePct = share(T.scope);
    /* Overall progress = average of Availability, Functionality and In use (a figure not yet updated for a week keeps its last value). */
    T.overall = weeks.map(function (_, w) {
      if (!(has('avail', w) || has('func', w) || has('inUse', w))) return null;
      var vs = [at(T.avail, w).v, at(T.func, w).v, at(T.inUse, w).v].filter(isNum);
      return vs.length ? vs.reduce(function (a, b) { return a + b; }, 0) / vs.length : null;
    });

    var weeklyKeys = Object.keys(M).filter(function (k) { return !fixed[k]; });
    model.fixed = fixed;
    model.latest = 0;
    weeks.forEach(function (_, w) {
      if (weeklyKeys.some(function (k) { return has(k, w); })) model.latest = w;
    });
    model.weekHasData = weeks.map(function (_, w) { return weeklyKeys.some(function (k) { return has(k, w); }); });
    return model;
  }

  /* Latest filled-in week at or before `sel` for a weekly series -> {v, w}; and the one before it. */
  function at(series, sel) {
    for (var w = Math.min(sel, series.length - 1); w >= 0; w--) if (series[w] !== null) return { v: series[w], w: w };
    return { v: null, w: -1 };
  }
  function before(series, w) { return w > 0 ? at(series, w - 1) : { v: null, w: -1 }; }
  function weekOf(key, sel) {
    var m = S.model;
    for (var w = Math.min(sel, m.weeks.length - 1); w >= 0; w--) if (m.has(key, w)) return w;
    return -1;
  }
  function deptVal(key, i, sel) { var w = weekOf(key, sel); return w < 0 ? null : S.model.M[key][w][i]; }
  /* ================= small pieces ================= */
  /* The week a view is showing: the Executive Summary is always on the latest updated week, the tracker on the selected one. */
  function viewWeek() { return S.tab === 'summary' ? S.model.latest : S.sel; }
  function staleTag(w) {
    return (w >= 0 && w < viewWeek()) ? '<span class="stale" title="This week is not filled in yet">as of ' + esc(S.model.weeks[w].label) + '</span>' : '';
  }
  function deltaTag(series, cur, opts) {
    if (cur.w < 0) return '';
    var prev = before(series, cur.w);
    if (prev.w < 0) return '<span class="dpill"><span class="delta flat">first week of data</span></span>';
    var d = cur.v - prev.v, unit = opts.unit || '', vs = ' vs ' + esc(S.model.weeks[prev.w].label);
    if (opts.rel) d = prev.v ? d / prev.v * 100 : 0;      // change as a % of last week's figure
    if (Math.abs(d) < 0.05) return '<span class="dpill"><span class="delta flat">no change' + vs + '</span></span>';
    var good = opts.downIsGood ? d < 0 : d > 0;
    return '<span class="dpill"><span class="delta ' + (good ? 'good' : 'bad') + '">' + (d > 0 ? '▲ ' : '▼ ') + fmt(Math.abs(d), opts.dec || 0) + unit + '</span><span>' + vs + '</span></span>';
  }
  /* Clean line icons used instead of emojis (stroke follows the card's colour). */
  var LINE_ICONS = {
    '🏢': '<rect x="4" y="3" width="16" height="18" rx="2"/><path d="M9 21v-4h6v4M8 7h2M14 7h2M8 11h2M14 11h2"/>',
    '✅': '<circle cx="12" cy="12" r="9"/><path d="m8 12.5 2.8 2.8L16.5 9"/>',
    '📊': '<path d="M4 20h16"/><rect x="6" y="11" width="3" height="7" rx="1"/><rect x="11" y="6" width="3" height="12" rx="1"/><rect x="16" y="13" width="3" height="5" rx="1"/>',
    '🚨': '<path d="M10.3 4.2 2.6 17.5A2 2 0 0 0 4.3 20.5h15.4a2 2 0 0 0 1.7-3L13.7 4.2a2 2 0 0 0-3.4 0Z"/><path d="M12 9.5v4M12 17h.01"/>',
    '📄': '<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8Z"/><path d="M14 3v5h5M9 13h6M9 17h4"/>',
    '⚙️': '<rect x="3" y="4" width="18" height="12" rx="2"/><path d="M8 20h8M12 16v4"/><path d="m8.5 10 2.2 2.2L15.5 7.5"/>',
    '🧭': '<circle cx="12" cy="12" r="9"/><path d="m15.5 8.5-2 5-5 2 2-5Z"/>',
    '📝': '<path d="M4 20h4L19 9a2.1 2.1 0 0 0-3-3L5 17Z"/><path d="M13.5 6.5l3 3M14 20h6"/>',
    '🐞': '<rect x="7" y="8" width="10" height="12" rx="5"/><path d="M9 8a3 3 0 0 1 6 0M12 12v8M3 13h4M17 13h4M4 7l3 2M20 7l-3 2M4 19l3-2M20 19l-3-2"/>',
    '🗂️': '<path d="m12 3 9 5-9 5-9-5Z"/><path d="m3 13 9 5 9-5"/>'
  };
  /* Line icon per department, picked from words in its name. */
  var DEPT_ICONS = [
    [/hatch/i, '<path d="M12 3c-3.5 0-6 5-6 9a6 6 0 0 0 12 0c0-4-2.5-9-6-9Z"/><path d="m8 13 2 1.5 2-2 2 2 2-1.5"/>'],
    [/live bird|sales|dispatch/i, '<path d="M3 12V5a2 2 0 0 1 2-2h7l9 9-9 9-9-9Z"/><circle cx="8" cy="8" r="1.5"/>'],
    [/breeder|broiler|farm|bird|poultry/i, '<path d="M4 20h16"/><path d="M6 20v-7l6-5 6 5v7"/><path d="M10 20v-4h4v4"/>'],
    [/feed|mill/i, '<path d="M7 4h10l-1 3c2 2 3 5 3 8a5 5 0 0 1-5 5h-4a5 5 0 0 1-5-5c0-3 1-6 3-8Z"/><path d="M9 13h6"/>'],
    [/buy ?back/i, '<path d="M4 9a8 8 0 0 1 14-3l2 2"/><path d="M20 4v4h-4"/><path d="M20 15a8 8 0 0 1-14 3l-2-2"/><path d="M4 20v-4h4"/>'],
    [/audit/i, '<circle cx="11" cy="11" r="6"/><path d="m20 20-4.5-4.5"/><path d="m8.5 11 1.8 1.8 3.2-3.3"/>'],
    [/security/i, '<path d="M12 3 5 6v6c0 4.5 3 7.5 7 9 4-1.5 7-4.5 7-9V6Z"/>'],
    [/procur|purchas/i, '<circle cx="9" cy="20" r="1.4"/><circle cx="17" cy="20" r="1.4"/><path d="M3 4h2l2.5 11h11l2-8H6.5"/>'],
    [/store|warehouse|inventor/i, '<path d="m12 3 8 4.5v9L12 21l-8-4.5v-9Z"/><path d="m4 7.5 8 4.5 8-4.5M12 12v9"/>'],
    [/auto|vehicle|transport|fleet/i, '<path d="M3 16V7h11v9"/><path d="M14 10h4l3 3v3h-7"/><circle cx="7" cy="17" r="2"/><circle cx="17" cy="17" r="2"/>'],
    [/financ|account/i, '<rect x="3" y="6" width="18" height="12" rx="2"/><circle cx="12" cy="12" r="2.5"/><path d="M7 9.5v5M17 9.5v5"/>'],
    [/\bhr\b|human|people/i, '<circle cx="9" cy="8" r="3"/><path d="M3 20a6 6 0 0 1 12 0"/><circle cx="17" cy="9" r="2.3"/><path d="M16 14a5 5 0 0 1 5 5"/>'],
    [/\bit\b|\bmis\b|technolog/i, '<rect x="3" y="4" width="18" height="12" rx="2"/><path d="M8 20h8M12 16v4"/>'],
    [/process|plant|factory|production/i, '<path d="M3 20V10l5 3V10l5 3V6h4l2 14Z"/>'],
    [/quality|\bqa\b|lab/i, '<path d="M9 3h6M10 3v6l-5 9a2 2 0 0 0 2 3h10a2 2 0 0 0 2-3l-5-9V3"/>']
  ];
  function deptIcon(name) {
    var hit = DEPT_ICONS.filter(function (r) { return r[0].test(name); })[0];
    return '<svg viewBox="0 0 24 24" width="19" height="19" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + (hit ? hit[1] : LINE_ICONS['🏢']) + '</svg>';
  }
  function icon(e, size) {
    return LINE_ICONS[e] ? '<svg class="ic" viewBox="0 0 24 24" width="' + (size || 18) + '" height="' + (size || 18) + '" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + LINE_ICONS[e] + '</svg>' : e;
  }
  function meter(v, color) {
    var p = isNum(v) ? Math.max(0, Math.min(100, v)) : 0;
    return '<div class="meter"><i style="width:' + p + '%' + (color ? ';background:' + color : '') + '"></i></div>';
  }
  function afiRow(label, v, color) {
    return '<div class="afi-row"><span>' + label + '</span>' + meter(v, color) + '<b>' + pct(v) + '</b></div>';
  }

  /* ================= KPI tiles ================= */
  function renderKpis() {
    var m = S.model, T = m.T, n = m.depts.length, sel = S.sel;
    var baDone = at(T.baDone, sel), baAvg = at(T.baAvg, sel);
    var scope = at(T.scope, sel), appr = at(T.appr, sel);
    var rPct = at(T.rPct, sel), rDone = at(T.rDone, sel), rExp = at(T.rExp, sel);
    var iOpen = at(T.iOpen, sel), iAll = at(T.iAll, sel), iDone = at(T.iDone, sel);
    var avail = at(T.avail, sel), func = at(T.func, sel), inUse = at(T.inUse, sel);
    var afiWs = [avail.w, func.w, inUse.w].filter(function (w) { return w >= 0; });
    var afiOld = afiWs.length ? Math.min.apply(null, afiWs) : -1, afiNew = afiWs.length ? Math.max.apply(null, afiWs) : -1;
    var afiStale = afiOld >= 0 && afiOld < sel
      ? '<span class="stale" title="Not every AFI column is filled in for this week">' + (afiNew > afiOld ? 'partly ' : '') + 'as of ' + esc(m.weeks[afiOld].label) + '</span>' : '';

    var baYes = at(T.baYes, sel), baYesPct = at(T.baYesPct, sel), scopePct = at(T.scopePct, sel), iPct = at(T.iPct, sel);
    var topIssue = '';
    if (iAll.w >= 0) {
      var best = -1;
      m.M.iAll[iAll.w].forEach(function (v, i) { if (isNum(v) && v > best) { best = v; topIssue = m.depts[i] + ' (' + fmt(v) + ')'; } });
    }
    var head = function (label, emoji) { return '<div class="t-top"><div class="t-label">' + label + '</div><span class="t-emoji" aria-hidden="true">' + icon(emoji) + '</span></div>'; };

    /* Big number = what is completed; the total and everything else stays small. Week-on-week change is in %. */
    var tiles = [
      head('Total departments', '🏢') +
      '<div class="t-value">' + n + '</div>' +
      '<div class="t-sub">in the ERP rollout scope</div>' +
      '<div class="t-foot"><span>' + esc(m.weeks[sel].label) + (m.weeks[sel].date ? ' · ' + esc(niceDate(m.weeks[sel].date, true)) : '') + '</span></div>',

      head('BA completed departments', '✅') +
      '<div class="t-value c-good">' + fmt(baYes.v) + '<small>of ' + n + '</small></div>' +
      meter(baYesPct.v, 'var(--good)') +
      '<div class="t-sub">departments · average BA progress ' + pct(baAvg.v) + '</div>' +
      '<div class="t-foot">' + deltaTag(T.baAvg, baAvg, { unit: '%' }) + staleTag(baAvg.w) + '</div>',

      head('Total reports', '📊') +
      '<div class="t-value">' + fmt(rExp.v) + '<small>reports</small></div>' +
      meter(rPct.v) +
      '<div class="t-sub"><b>' + fmt(rDone.v) + '</b> completed · ' + pct(rPct.v) + '</div>' +
      '<div class="t-foot">' + deltaTag(T.rPct, rPct, { unit: '%' }) + staleTag(rDone.w) + '</div>',

      head('Current issues', '🚨') +
      '<div class="t-value c-crit">' + fmt(iAll.v) + '</div>' +
      '<div class="t-sub"><b>' + fmt(iDone.v) + '</b> resolved' + (topIssue ? ' · Most: ' + esc(topIssue) : '') + '</div>' +
      '<div class="t-foot">' + deltaTag(T.iAll, iAll, { unit: '%', rel: true, downIsGood: true }) + staleTag(iAll.w) + '</div>',

      head('Scope documents', '📄') +
      '<div class="t-value">' + fmt(scope.v) + '<small>/ ' + n + '</small></div>' +
      '<div class="t-sub">prepared · ' + fmt(appr.v) + ' approved</div>' + meter(scopePct.v) +
      '<div class="t-foot">' + (m.fixed.scope ? '<span class="dpill"><span class="delta flat">current status · not tracked weekly</span></span>' : deltaTag(T.scopePct, scopePct, { unit: '%' }) + staleTag(scope.w)) + '</div>',

      head('System readiness (AFI)', '⚙️') +
      '<div class="afi-rows">' + afiRow('Availability', avail.v, 'var(--s1)') + afiRow('Functionality', func.v, 'var(--s2)') + afiRow('In use', inUse.v, 'var(--s3)') + '</div>' +
      '<div class="t-foot">' + deltaTag(T.inUse, inUse, { unit: '% in use' }) + afiStale + '</div>'
    ];
    var pops = kpiDetails();
    return '<section class="kpis" id="kpis">' + tiles.map(function (t, k) {
      return '<div class="card tile" tabindex="0">' + t + '<div class="t-pop"><div class="t-pop-in">' + pops[k] + '</div></div></div>';
    }).join('') + '</section>';
  }

  /* Department breakdown shown when the cursor is on a headline card (same order as the cards). */
  function kpiDetails() {
    var m = S.model, sel = S.sel;
    var rows = m.depts.map(function (name, i) {
      var exp = deptVal('rExp', i, sel), done = deptVal('rDone', i, sel), all = deptVal('iAll', i, sel), res = deptVal('iDone', i, sel);
      return {
        no: i + 1, name: name, baYes: isYes(deptVal('ba', i, sel)), ba: deptVal('baPct', i, sel),
        scope: isYes(deptVal('scope', i, sel)), appr: isYes(deptVal('appr', i, sel)),
        exp: exp, done: isNum(exp) ? (isNum(done) ? done : 0) : null,
        all: all, res: isNum(all) ? (isNum(res) ? res : 0) : null,
        avail: deptVal('avail', i, sel), func: deptVal('func', i, sel), inUse: deptVal('inUse', i, sel)
      };
    });
    var title = function (t, s) { return '<div class="p-title">' + t + (s ? '<span>' + s + '</span>' : '') + '</div>'; };
    var group = function (cls, label, list, right) {
      return '<div class="p-group ' + cls + '"><div class="p-ghead">' + label + ' <b>' + list.length + '</b></div>' +
        (list.length ? list.map(function (r) { return '<div class="p-row"><span>' + esc(r.name) + '</span>' + (right ? '<b>' + right(r) + '</b>' : '') + '</div>'; }).join('')
          : '<div class="p-none">None</div>') + '</div>';
    };
    var table = function (heads, cells) {
      return '<table class="p-table"><thead><tr><th>Department</th>' + heads.map(function (h) { return '<th>' + h + '</th>'; }).join('') + '</tr></thead><tbody>' +
        rows.map(function (r) { return '<tr><td>' + esc(r.name) + '</td>' + cells(r).map(function (c) { return '<td' + (c[1] ? ' class="' + c[1] + '"' : '') + '>' + c[0] + '</td>'; }).join('') + '</tr>'; }).join('') +
        '</tbody></table>';
    };
    var baPct = function (r) { return pct(r.ba); };
    var sum = function (k) { return rows.reduce(function (a, r) { return a + (isNum(r[k]) ? r[k] : 0); }, 0); };
    var by = function (f) { return rows.filter(f); };

    return [
      title('🏢 All departments', rows.length + ' in scope') +
      '<ol class="p-list">' + rows.map(function (r) { return '<li>' + esc(r.name) + '</li>'; }).join('') + '</ol>',

      title('✅ BA status by department') +
      group('g-done', '✅ BA completed', by(function (r) { return r.baYes; }), baPct) +
      group('g-pend', '⏳ BA pending', by(function (r) { return !r.baYes; }), baPct),

      title('📊 Reports by department', fmt(sum('done')) + ' completed · ' + fmt(sum('exp') - sum('done')) + ' pending') +
      table(['Completed', 'Pending', 'Total'], function (r) {
        return [[fmt(r.done), 'c-done'], [isNum(r.exp) ? fmt(r.exp - r.done) : '–', 'c-pend'], [fmt(r.exp), '']];
      }),

      title('🚨 Issues by department', fmt(sum('res')) + ' resolved · ' + fmt(sum('all') - sum('res')) + ' open') +
      table(['Resolved', 'Open', 'Total'], function (r) {
        return [[fmt(r.res), 'c-done'], [isNum(r.all) ? fmt(r.all - r.res) : '–', 'c-pend'], [fmt(r.all), '']];
      }),

      title('📄 Scope documents by department') +
      group('g-done', '✅ Approved', by(function (r) { return r.scope && r.appr; })) +
      group('g-mid', '📄 Prepared, awaiting approval', by(function (r) { return r.scope && !r.appr; })) +
      group('g-pend', '⏳ Not prepared yet', by(function (r) { return !r.scope; })),

      title('⚙️ AFI by department') +
      table(['Availability', 'Functionality', 'In use'], function (r) { return [[pct(r.avail), ''], [pct(r.func), ''], [pct(r.inUse), '']]; })
    ];
  }
  function placePops() {
    var host = $('kpis');
    if (!host) return;
    var vw = document.documentElement.clientWidth;
    Array.prototype.forEach.call(host.children, function (t) {
      t.classList.toggle('pop-r', t.getBoundingClientRect().left + 440 > vw);
    });
  }

  /* ================= trend charts ================= */
  var CHARTS = [
    { id: 'c-ba', title: 'BA progress', sub: 'Average BA completion across all departments', unit: '%', max: 100,
      series: [{ name: 'Average BA %', key: 'baAvg', color: 'var(--s1)' }] },
    { id: 'c-afi', title: 'System readiness (AFI)', sub: 'Average availability, functionality and usage', unit: '%', max: 100,
      series: [{ name: 'Availability', key: 'avail', color: 'var(--s1)' }, { name: 'Functionality', key: 'func', color: 'var(--s2)' }, { name: 'In use', key: 'inUse', color: 'var(--s3)' }] },
    { id: 'c-rep', title: 'Reports completed', sub: 'Share of expected reports that are done', unit: '%', max: 100,
      series: [{ name: 'Reports done %', key: 'rPct', color: 'var(--s1)' }] },
    { id: 'c-iss', title: 'Issues', sub: 'Issues raised against issues resolved', unit: '',
      series: [{ name: 'Raised', key: 'iAll', color: 'var(--s1)' }, { name: 'Resolved', key: 'iDone', color: 'var(--s2)' }] }
  ];

  function renderChartShells() {
    return '<div class="sec-head"><div><h2>Week-by-week trend</h2><p>Only weeks that have been updated in the sheet are shown; a new week appears as soon as it is filled in. The shaded week is the one selected above.</p></div></div>' +
      '<section class="charts">' + CHARTS.map(function (c) {
        var legend = c.series.length > 1 ? '<div class="legend">' + c.series.map(function (s) {
          return '<span><i style="background:' + s.color + '"></i>' + s.name + '</span>';
        }).join('') + '</div>' : '';
        return '<div class="card chart"><h3>' + c.title + '</h3><p>' + c.sub + '</p>' + legend +
          '<div class="plot" id="' + c.id + '"></div></div>';
      }).join('') + '</section>';
  }

  function niceMax(v) {
    if (!(v > 0)) return 10;
    var p = Math.pow(10, Math.floor(Math.log(v) / Math.LN10)), f = v / p;
    return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 5 ? 5 : 10) * p;
  }

  function drawChart(cfg) {
    var host = $(cfg.id);
    if (!host) return;
    var m = S.model;
    var idx = [];                                   // only weeks that have been updated are plotted
    m.weeks.forEach(function (_, w) {
      var ok = cfg.own ? cfg.series.some(function (s) { return isNum(m.T[s.key][w]); }) : m.weekHasData[w];   // own: only weeks this chart has a value for
      if (ok) idx.push(w);
    });
    var weeks = idx.map(function (w) { return m.weeks[w]; }), n = weeks.length, selPos = idx.indexOf(cfg.sel === undefined ? S.sel : cfg.sel);
    if (!n) { host.innerHTML = ''; return; }
    var W = Math.max(280, host.clientWidth), H = cfg.h || 220, mg = { t: 14, r: 46, b: 30, l: 38 };
    var iw = W - mg.l - mg.r, ih = H - mg.t - mg.b;
    var series = cfg.series.map(function (s) { return { name: s.name, color: s.color, vals: idx.map(function (w) { return m.T[s.key][w]; }) }; });
    var top = cfg.max || niceMax(Math.max.apply(null, [0].concat(series.map(function (s) { return Math.max.apply(null, s.vals.filter(isNum).concat([0])); }))));
    var fc = cfg.fc && cfg.fc.b > 0 ? cfg.fc : null, F = fc ? 3 : 0, N = n + F;   // forecast: 3 extra weeks on the right
    var x = function (i) { return N === 1 ? mg.l + iw / 2 : mg.l + i * iw / (N - 1); };
    var y = function (v) { return mg.t + (1 - v / top) * ih; };
    var step = N > 1 ? iw / (N - 1) : iw;
    var out = [];

    // selected-week band, grid, axis labels
    if (selPos >= 0) out.push('<rect x="' + (x(selPos) - Math.min(step, 56) / 2) + '" y="' + mg.t + '" width="' + Math.min(step, 56) + '" height="' + ih + '" rx="6" fill="var(--track)" opacity=".7"/>');
    [0, 0.5, 1].forEach(function (f) {
      var yy = y(top * f);
      out.push('<line x1="' + mg.l + '" x2="' + (W - mg.r) + '" y1="' + yy + '" y2="' + yy + '" stroke="' + (f === 0 ? 'var(--axis)' : 'var(--grid)') + '" stroke-width="1"/>');
      out.push('<text x="' + (mg.l - 8) + '" y="' + (yy + 4) + '" text-anchor="end">' + fmt(top * f) + cfg.unit + '</text>');
    });
    var every = Math.ceil(n / Math.max(1, Math.floor(iw / 44)));
    weeks.forEach(function (wk, i) {
      if (i % every === 0 || i === selPos) out.push('<text x="' + x(i) + '" y="' + (H - 8) + '" text-anchor="middle"' + (i === selPos ? ' class="val"' : '') + '>' + wk.short + '</text>');
    });
    for (var k = 1; k <= F; k++) out.push('<text x="' + x(n - 1 + k) + '" y="' + (H - 8) + '" text-anchor="middle" opacity=".6">W' + (fc.last + 1 + k) + '</text>');
    if (F) out.push('<rect x="' + (x(n - 1) + step / 2) + '" y="' + mg.t + '" width="' + (x(N - 1) - x(n - 1) - step / 2 + 20) + '" height="' + ih + '" fill="var(--track)" opacity=".35" rx="6"/>' +
      '<text x="' + (x(n - 1) + step / 2 + 8) + '" y="' + (mg.t + 14) + '" opacity=".8">Forecast</text>');

    // lines, markers, end labels
    var labels = [];
    series.forEach(function (s) {
      var d = '', pen = false, last = -1, first = -1;
      s.vals.forEach(function (v, i) {
        if (!isNum(v)) { pen = false; return; }
        if (first < 0) first = i;
        if (pen) {                                  // smooth curve between consecutive weeks
          var x0 = x(last), mid = (x0 + x(i)) / 2;
          d += 'C' + mid + ' ' + y(s.vals[last]) + ' ' + mid + ' ' + y(v) + ' ' + x(i) + ' ' + y(v);
        } else d += 'M' + x(i) + ' ' + y(v);
        pen = true; last = i;
      });
      if (cfg.area && last > first) {               // soft gradient under the line
        out.push('<defs><linearGradient id="' + cfg.id + '-g" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="' + s.color + '" stop-opacity=".32"/><stop offset="1" stop-color="' + s.color + '" stop-opacity="0"/></linearGradient></defs>' +
          '<path class="area" d="' + d + 'L' + x(last) + ' ' + y(0) + 'L' + x(first) + ' ' + y(0) + 'Z" fill="url(#' + cfg.id + '-g)" stroke="none"/>');
      }
      if (d) out.push('<path pathLength="1" d="' + d + '" fill="none" stroke="' + s.color + '" stroke-width="' + (cfg.bold ? 3.5 : 2) + '" stroke-linejoin="round" stroke-linecap="round"/>');
      s.vals.forEach(function (v, i) {
        if (isNum(v)) out.push('<circle cx="' + x(i) + '" cy="' + y(v) + '" r="' + (cfg.bold ? 6.5 : 4.5) + '" fill="' + s.color + '" stroke="var(--surface)" stroke-width="2"/>');
      });
      if (cfg.labelAll) {                           // few points: write the value above every one
        s.vals.forEach(function (v, i) { if (isNum(v)) out.push('<text class="val big" text-anchor="middle" x="' + x(i) + '" y="' + (y(v) - 14) + '">' + fmt(v) + cfg.unit + '</text>'); });
      } else if (last >= 0) labels.push({ x: x(last) + 9, y: y(s.vals[last]) + 4, t: fmt(s.vals[last]) + cfg.unit });
    });
    // dashed regression forecast from the latest real week
    if (fc) {
      var col = series[0].color, pts = [[x(n - 1), y(series[0].vals[n - 1])]], marks = '';
      for (var k = 1; k <= F; k++) {
        var fv = Math.min(100, fc.a + fc.b * (fc.last + k)), px = x(n - 1 + k), py = y(fv);
        pts.push([px, py]);
        marks += '<circle cx="' + px + '" cy="' + py + '" r="4.5" fill="var(--surface)" stroke="' + col + '" stroke-width="2"/>' +
          '<text class="val" text-anchor="middle" x="' + px + '" y="' + (py - 12) + '" opacity=".75">' + fmt(fv) + '%</text>';
      }
      out.push('<path class="fc" d="M' + pts.map(function (p) { return p[0] + ' ' + p[1]; }).join('L') + '" fill="none" stroke="' + col + '" stroke-width="2.5" stroke-dasharray="7 6" stroke-linecap="round" opacity=".75"/>' + marks);
    }
    // two series sitting on the same point need the value written only once
    labels = labels.filter(function (l, i) {
      return !labels.some(function (o, j) { return j < i && o.t === l.t && o.x === l.x && Math.abs(o.y - l.y) < 1; });
    });
    labels.sort(function (a, b) { return a.y - b.y; });
    for (var i = 1; i < labels.length; i++) if (labels[i].y - labels[i - 1].y < 14) labels[i].y = labels[i - 1].y + 14;
    labels.forEach(function (l) { out.push('<text class="val" x="' + l.x + '" y="' + l.y + '">' + l.t + '</text>'); });

    out.push('<line class="cross" x1="0" x2="0" y1="' + mg.t + '" y2="' + (mg.t + ih) + '" stroke="var(--axis)" stroke-width="1" visibility="hidden"/>');
    host.innerHTML = '<svg style="height:' + H + 'px" viewBox="0 0 ' + W + ' ' + H + '" role="img" aria-label="' + esc(cfg.title) + ' by week">' + out.join('') + '</svg><div class="tip"></div>';

    // hover: crosshair + tooltip on the nearest week
    var svg = host.firstChild, tip = host.lastChild, cross = svg.querySelector('.cross');
    svg.addEventListener('mousemove', function (e) {
      var box = svg.getBoundingClientRect(), px = (e.clientX - box.left) * W / box.width;
      var i = n === 1 ? 0 : Math.max(0, Math.min(n - 1, Math.round((px - mg.l) / step)));   // tooltip only on real weeks
      cross.setAttribute('x1', x(i)); cross.setAttribute('x2', x(i)); cross.setAttribute('visibility', 'visible');
      tip.innerHTML = '<div class="tw">' + esc(weeks[i].label) + '</div>' + (weeks[i].date ? '<div class="tp">' + esc(weekRange(weeks[i].date)) + '</div>' : '') +
        series.map(function (s) {
          return '<div class="tr"><span><i style="background:' + s.color + '"></i>' + s.name + '</span><b>' + (isNum(s.vals[i]) ? fmt(s.vals[i]) + cfg.unit : 'not updated') + '</b></div>';
        }).join('');
      tip.style.display = 'block';
      var hostW = host.clientWidth, left = x(i) * hostW / W + 14;      // layout pixels, so the page zoom does not shift the tooltip
      if (left + tip.offsetWidth > hostW) left = x(i) * hostW / W - tip.offsetWidth - 14;
      tip.style.left = Math.max(0, left) + 'px';
      tip.style.top = '8px';
    });
    svg.addEventListener('mouseleave', function () { tip.style.display = 'none'; cross.setAttribute('visibility', 'hidden'); });
  }
  function drawCharts() { CHARTS.forEach(drawChart); }

  /* ================= rollout journey ================= */
  var ZONES = [
    { cls: 'z-risk',  name: 'At risk',  from: 0,  to: 49,  mark: '▲' },
    { cls: 'z-adopt', name: 'Adopting', from: 50, to: 89,  mark: '◆' },
    { cls: 'z-live',  name: 'Live',     from: 90, to: 100, mark: '✓' }
  ];
  var ICONS = [[/bre[ae]der/, '🐣'], [/broiler/, '🐔'], [/buy ?back/, '🤝'], [/feed/, '🌾'], [/audit/, '🔍'], [/security/, '🛡️'],
    [/proc/, '🛒'], [/store/, '🏬'], [/auto/, '🚚'], [/live bird/, '🐓'], [/production/, '🏭'], [/sales/, '📈'], [/^qa|quality/, '🧪'],
    [/finished|goods/, '📦'], [/^hr|human/, '👥'], [/financ|account/, '💰'], [/^mis|^it /, '💻']];
  function zoneOf(v) { return v >= 90 ? 2 : v >= 50 ? 1 : 0; }
  function shortName(s) { return s.replace(/\s+(department|operation)\b/i, '').replace(/\s*\(.*?\)/, '').trim() || s; }
  function iconOf(name) {
    var k = norm(name);
    for (var i = 0; i < ICONS.length; i++) if (ICONS[i][0].test(k)) return { t: ICONS[i][1], txt: false };
    return { t: shortName(name).split(/\s+/).map(function (w) { return w.charAt(0); }).join('').slice(0, 2).toUpperCase(), txt: true };
  }
  function journeyItems() {
    var m = S.model;
    return m.depts.map(function (name, i) {
      return { no: i + 1, name: name, v: deptVal('inUse', i, viewWeek()), avail: deptVal('avail', i, viewWeek()), func: deptVal('func', i, viewWeek()) };
    });
  }
  function renderJourneyShell() {
    var items = journeyItems(), counts = [0, 0, 0], none = [];
    items.forEach(function (d) { if (isNum(d.v)) counts[zoneOf(d.v)]++; else none.push(d.name); });
    var w = weekOf('inUse', viewWeek());
    return '<div class="sec-head"><div><h2>Rollout journey</h2><p>Every department placed by how much the ERP is in use. Click a department to jump to its card. ' + staleTag(w) + '</p></div>' +
      '<div class="zone-chips">' + ZONES.slice().reverse().map(function (z, k) {
        return '<span class="zone-chip ' + z.cls + '"><i class="zdot"></i><b>' + counts[2 - k] + '</b>' + z.mark + ' ' + z.name + '</span>';
      }).join('') + '</div></div>' +
      '<div class="card journey"><div class="j-plot" id="journey"></div>' +
      (none.length ? '<p class="note">No “In use” figure yet: ' + esc(none.join(', ')) + '.</p>' : '') + '</div>';
  }
  function drawJourney(animate) {
    var host = $('journey');
    if (!host) return;
    var W = host.clientWidth;
    var PAD = Math.max(24, Math.min(48, W * .035));
    var GAP = Math.max(38, Math.min(60, W / 24));
    var LEVEL = 82, iw = Math.max(1, W - PAD * 2);
    var items = journeyItems().filter(function (d) { return isNum(d.v); });
    items.forEach(function (d) { d.x = PAD + Math.max(0, Math.min(100, d.v)) / 100 * iw; });
    items.sort(function (a, b) { return a.x - b.x || a.no - b.no; });
    var placed = [], maxLevel = 0;
    items.forEach(function (d) {            // stack departments that would overlap
      var lvl = 0;
      while (placed.some(function (p) { return p.level === lvl && Math.abs(p.x - d.x) < GAP; })) lvl++;
      d.level = lvl; placed.push(d); maxLevel = Math.max(maxLevel, lvl);
    });
    var targetHeight = Math.max(300, Math.min(560, W * .32));
    LEVEL = Math.min(82, Math.max(42, (targetHeight - 60) / (maxLevel + 1)));
    var nodeScale = Math.max(.58, Math.min(1, LEVEL / 82));
    host.style.setProperty('--j-dot', (44 * nodeScale) + 'px');
    host.style.setProperty('--j-icon', (20 * nodeScale) + 'px');
    host.style.setProperty('--j-txt', Math.max(8, 13 * nodeScale) + 'px');
    host.style.setProperty('--j-font', Math.max(8, 10 * nodeScale) + 'px');
    host.style.setProperty('--j-name-w', (60 * nodeScale) + 'px');
    host.style.setProperty('--j-name-h', Math.max(18, 25 * nodeScale) + 'px');
    host.style.height = (52 + (maxLevel + 1) * LEVEL + 8) + 'px';

    var html = '<div class="j-track">' + ZONES.map(function (z) { return '<i class="j-seg ' + z.cls + '" style="flex:' + (z.to - z.from + 1) + '"></i>'; }).join('') + '</div>' +
      '<div class="j-zones">' + ZONES.map(function (z) {
        return '<span style="flex:' + (z.to - z.from + 1) + '"><b>' + z.mark + ' ' + z.name + '</b> ' + z.from + (z.to === 100 ? '%+' : '–' + z.to + '%') + '</span>';
      }).join('') + '</div><span class="j-flag" aria-hidden="true">🏁</span>';
    items.forEach(function (d, k) {
      var z = ZONES[zoneOf(d.v)], ic = iconOf(d.name), start = animate ? PAD : d.x;
      var edge = d.x < 120 ? ' edge-l' : d.x > W - 120 ? ' edge-r' : '';
      var delay = animate ? 'transition-delay:' + (k * 45) + 'ms;' : '';
      html += '<i class="j-stem ' + z.cls + '" data-x="' + d.x + '" data-h="' + (12 + d.level * LEVEL) + '" style="left:' + start + 'px;height:' + (animate ? 0 : 12 + d.level * LEVEL) + 'px;' + delay + '"></i>' +
        '<button type="button" class="j-node ' + z.cls + edge + '" data-no="' + d.no + '" data-x="' + d.x + '" style="left:' + start + 'px;bottom:' + (50 + d.level * LEVEL) + 'px;' + (animate ? 'opacity:0;' : '') + delay + '">' +
        '<span class="j-dot' + (ic.txt ? ' txt' : '') + '">' + ic.t + '</span><span class="j-name">' + esc(shortName(d.name)) + '</span>' +
        '<span class="j-tip"><strong>' + esc(d.name) + '</strong>' + z.mark + ' ' + z.name + ' · In use <b>' + pct(d.v) + '</b><br>Availability <b>' + pct(d.avail) + '</b> · Functionality <b>' + pct(d.func) + '</b></span></button>';
    });
    host.innerHTML = html;
    if (animate) {
      host.getBoundingClientRect();           // commit the start positions, then let them travel
      setTimeout(function () {
        Array.prototype.forEach.call(host.querySelectorAll('[data-x]'), function (el) {
          el.style.left = el.getAttribute('data-x') + 'px';
          el.style.opacity = '';
          if (el.hasAttribute('data-h')) el.style.height = el.getAttribute('data-h') + 'px';
        });
      }, 40);
    }
  }
  /* Headline numbers count up from zero. */
  function countUp() {
    Array.prototype.forEach.call(document.querySelectorAll('.t-value, .x-val'), function (el) {
      var node = el.firstChild;
      if (!node || node.nodeType !== 3) return;
      var target = Number(node.nodeValue.replace(/,/g, ''));
      if (!isFinite(target) || !target) return;
      var t0 = Date.now();
      (function step() {
        var p = Math.min(1, (Date.now() - t0) / 800), e = 1 - Math.pow(1 - p, 3);
        node.nodeValue = fmt(Math.round(target * e));
        if (p < 1) setTimeout(step, 16);
      })();
    });
  }

  /* ================= departments ================= */
  function deptRows() {
    var m = S.model, sel = S.sel;
    var rows = m.depts.map(function (name, i) {
      var exp = deptVal('rExp', i, sel), done = deptVal('rDone', i, sel);
      var all = deptVal('iAll', i, sel), res = deptVal('iDone', i, sel);
      var baW = weekOf('baPct', sel), prevW = baW > 0 ? weekOf('baPct', baW - 1) : -1;
      var ba = deptVal('baPct', i, sel);
      return {
        no: i + 1, name: name, ba: ba,
        baPrev: prevW >= 0 ? m.M.baPct[prevW][i] : null,
        scope: isYes(deptVal('scope', i, sel)), appr: isYes(deptVal('appr', i, sel)),
        avail: deptVal('avail', i, sel), func: deptVal('func', i, sel), inUse: deptVal('inUse', i, sel),
        exp: exp, done: isNum(exp) ? (isNum(done) ? done : 0) : null,
        rPct: isNum(exp) && exp > 0 ? (isNum(done) ? done : 0) / exp * 100 : null,
        all: all, res: isNum(all) ? (isNum(res) ? res : 0) : null,
        open: isNum(all) ? all - (isNum(res) ? res : 0) : null
      };
    });
    var q = norm(S.q);
    if (q) rows = rows.filter(function (r) { return norm(r.name).indexOf(q) >= 0; });
    var low = function (k) { return function (a, b) { return (isNum(a[k]) ? a[k] : Infinity) - (isNum(b[k]) ? b[k] : Infinity) || a.no - b.no; }; };
    var high = function (k) { return function (a, b) { return (isNum(b[k]) ? b[k] : -Infinity) - (isNum(a[k]) ? a[k] : -Infinity) || a.no - b.no; }; };
    var sorts = { no: function (a, b) { return a.no - b.no; }, ba: low('ba'), issues: high('open'), reports: low('rPct'), inuse: low('inUse') };
    return rows.sort(sorts[S.sort] || sorts.no);
  }

  function deptCard(r) {
    var status = !isNum(r.ba) ? '<span class="pill">No BA data</span>'
      : r.ba >= 100 ? '<span class="pill done">✓ BA complete</span>'
      : r.ba > 0 ? '<span class="pill go">BA in progress</span>'
      : '<span class="pill">BA not started</span>';
    var d = isNum(r.ba) && isNum(r.baPrev) ? r.ba - r.baPrev : null;
    var move = d === null ? '' : d > 0 ? ' <span class="delta good">▲ ' + fmt(d) + '%</span>' : d < 0 ? ' <span class="delta bad">▼ ' + fmt(-d) + '%</span>' : '';
    var risk = isNum(r.inUse) && zoneOf(r.inUse) === 0;     // ERP in use below 50%
    if (risk) status = '<div class="d-pills">' + status + '<span class="pill risk">▲ At risk</span></div>';
    return '<article class="card dept' + (risk ? ' risk' : '') + '" id="dept-' + r.no + '" data-no="' + r.no + '" tabindex="0" role="button" aria-label="Open details for ' + esc(r.name) + '">' +
      '<div class="d-head"><div><div class="d-no">Department ' + r.no + '</div><div class="d-name">' + esc(r.name) + '</div></div>' + status + '</div>' +
      '<div class="d-ba"><div class="d-row"><span>BA progress' + move + '</span><b>' + pct(r.ba) + '</b></div>' + meter(r.ba) + '</div>' +
      '<div class="chips"><span class="chip' + (r.scope ? ' yes' : '') + '">' + (r.scope ? '✓ ' : '') + 'Scope doc' + (r.scope ? '' : ' pending') + '</span>' +
      '<span class="chip' + (r.appr ? ' yes' : '') + '">' + (r.appr ? '✓ ' : '') + 'Approval' + (r.appr ? '' : ' pending') + '</span></div>' +
      '<div class="d-afi">' + afiRow('Availability', r.avail, 'var(--s1)') + afiRow('Functionality', r.func, 'var(--s2)') + afiRow('In use', r.inUse, 'var(--s3)') + '</div>' +
      '<div class="d-foot">' +
        '<div class="d-stat"><span>Reports completed</span><b>' + fmt(r.done) + '</b><small>of ' + fmt(r.exp) + (isNum(r.rPct) ? ' · ' + pct(r.rPct) : '') + '</small></div>' +
        '<div class="d-stat"><span>Issues resolved</span><b>' + fmt(r.res) + '</b><small>of ' + fmt(r.all) + (isNum(r.all) && r.all > 0 ? ' · ' + pct(r.res / r.all * 100) : '') + '</small></div>' +
      '</div></article>';
  }

  /* Full detail for one department, opened by clicking its card or its marker on the journey. */
  function openDept(no, week) {
    var m = S.model, i = no - 1, sel = week === undefined ? S.sel : week, name = m.depts[i];
    if (!name) return;
    var v = function (key) { return deptVal(key, i, sel); };
    var ba = v('baPct'), exp = v('rExp'), done = isNum(exp) ? (isNum(v('rDone')) ? v('rDone') : 0) : null;
    var all = v('iAll'), res = isNum(all) ? (isNum(v('iDone')) ? v('iDone') : 0) : null;
    var inUse = v('inUse'), zone = isNum(inUse) ? ZONES[zoneOf(inUse)] : null, ic = iconOf(name);
    var chip = function (ok, yes, noText) { return '<span class="chip' + (ok ? ' yes' : '') + '">' + (ok ? '✓ ' + yes : noText) + '</span>'; };
    var sec = function (title, body) { return '<section class="m-sec"><h4>' + title + '</h4>' + body + '</section>'; };
    var big = function (value, small) { return '<div class="m-big">' + value + (small ? '<small>' + small + '</small>' : '') + '</div>'; };

    var history = m.weeks.map(function (wk, w) {
      if (!m.weekHasData[w]) return '';
      var g = function (key) { return m.M[key][w][i]; };
      var pair = function (a, b) { return isNum(g(b)) ? fmt(isNum(g(a)) ? g(a) : 0) + ' / ' + fmt(g(b)) : '–'; };
      return '<tr' + (w === sel ? ' class="now"' : '') + '><td>' + esc(wk.label) + (wk.date ? '<span>' + esc(niceDate(wk.date)) + '</span>' : '') + '</td><td>' + pct(g('baPct')) + '</td><td>' +
        pct(g('avail')) + '</td><td>' + pct(g('func')) + '</td><td>' + pct(g('inUse')) + '</td><td>' + pair('rDone', 'rExp') + '</td><td>' + pair('iDone', 'iAll') + '</td></tr>';
    }).join('');

    $('modal').innerHTML = '<div class="m-box card" role="dialog" aria-modal="true" aria-label="' + esc(name) + '">' +
      '<button type="button" class="m-close" aria-label="Close">✕</button>' +
      '<div class="m-head"><span class="j-dot ' + (zone ? zone.cls : '') + (ic.txt ? ' txt' : '') + '">' + ic.t + '</span>' +
        '<div><div class="d-no">Department ' + no + ' · status as of ' + esc(m.weeks[sel].label) + '</div><h2>' + esc(name) + '</h2></div>' +
        (zone ? '<span class="zone-chip ' + zone.cls + '"><i class="zdot"></i>' + zone.mark + ' ' + zone.name + '</span>' : '') + '</div>' +
      '<div class="m-grid">' +
        sec('✅ Business analysis (BA)', big(pct(ba), 'BA progress') + meter(ba) +
          '<div class="chips">' + chip(isYes(v('ba')), 'BA done', 'BA pending') + chip(isYes(v('scope')), 'Scope doc prepared', 'Scope doc pending') + chip(isYes(v('appr')), 'Scope doc approved', 'Approval pending') + '</div>') +
        sec('⚙️ System readiness (AFI)', '<div class="afi-rows">' + afiRow('Availability', v('avail'), 'var(--s1)') + afiRow('Functionality', v('func'), 'var(--s2)') + afiRow('In use', inUse, 'var(--s3)') + '</div>') +
        sec('📊 Reports', big(fmt(exp), 'total reports') + meter(isNum(exp) && exp ? done / exp * 100 : null) +
          '<div class="m-line"><span><b class="ok">' + fmt(done) + '</b> completed' + (isNum(exp) && exp ? ' · ' + pct(done / exp * 100) : '') + '</span><span><b class="wait">' + (isNum(exp) ? fmt(exp - done) : '–') + '</b> pending</span></div>') +
        sec('🚨 Issues', big(fmt(all), 'issues raised') + meter(isNum(all) && all ? res / all * 100 : null, 'var(--good)') +
          '<div class="m-line"><span><b class="ok">' + fmt(res) + '</b> resolved' + (isNum(all) && all ? ' · ' + pct(res / all * 100) : '') + '</span><span><b class="wait">' + (isNum(all) ? fmt(all - res) : '–') + '</b> still open</span></div>') +
      '</div>' +
      '<h4 class="m-h">📅 Week by week</h4><div class="table-wrap"><table class="m-table"><thead><tr><th>Week</th><th>BA %</th><th>Availability</th><th>Functionality</th><th>In use</th><th>Reports done</th><th>Issues resolved</th></tr></thead><tbody>' +
      history + '</tbody></table></div></div>';
    $('modal').hidden = false;
    document.body.style.overflow = 'hidden';
    $('modal').querySelector('.m-close').focus();
  }
  /* Rollout journey click: scroll to that department's card and highlight it. */
  function goToDept(no) {
    if (S.q || S.view !== 'cards') {                 // make sure the card is visible
      S.q = ''; S.view = 'cards';
      if ($('q')) $('q').value = '';
      Array.prototype.forEach.call($('view').children, function (c) { c.setAttribute('aria-pressed', String(c.getAttribute('data-view') === 'cards')); });
      renderDeptBody();
    }
    var card = $('dept-' + no);
    if (!card) return;
    card.scrollIntoView({ behavior: 'smooth', block: 'center' });
    card.classList.remove('flash'); void card.offsetWidth; card.classList.add('flash');
    setTimeout(function () { card.focus({ preventScroll: true }); }, 600);
  }
  /* ---------- department documents from Google Drive ---------- */
  var DOC_LABEL = { ba: 'BA document', scope: 'Scope document', other: 'Other document' };
  /* The Drive file for a department + type, matched on the department folder name. */
  function docsFor(i, type) {                        // every file of that type, as a list
    var docs = (S.model && S.model.docs) || {}, k = norm(S.model.depts[i]), hit = docs[k];
    if (!hit) Object.keys(docs).some(function (f) {
      if (f && (k.indexOf(f) >= 0 || f.indexOf(k) >= 0)) { hit = docs[f]; return true; }
    });
    var v = hit && hit[type];
    return !v ? [] : Array.isArray(v) ? v : [v];
  }
  function docFor(i, type) { return docsFor(i, type)[0] || null; }
  /* 📄 button (or a labelled button) that opens the document; empty when there is no file. */
  function docButton(i, type, label) {
    if (!docFor(i, type)) return '';
    return '<button type="button" class="doc-btn' + (label ? ' lbl' : '') + '" data-doc-dept="' + i + '" data-doc-type="' + type + '" title="Open ' + DOC_LABEL[type] + '">' + (label ? '📄 ' + label : '📄') + '</button>';
  }
  /* Large popup with the Drive preview of the document. */
  function openDoc(i, type, k) {
    var d = docsFor(i, type)[k || 0];
    if (!d) return;
    $('modal').innerHTML = '<div class="m-box card m-doc" role="dialog" aria-modal="true" aria-label="' + esc(d.name) + '">' +
      '<button type="button" class="m-close" aria-label="Close">✕</button>' +
      '<div class="m-head"><div><div class="d-no">' + esc(S.model.depts[i]) + ' · ' + DOC_LABEL[type] + '</div><h2>' + esc(d.name) + '</h2></div>' +
      (S.tab === 'scope' ? '<button type="button" class="btn sd-back" data-docs-dept="' + i + '">← All documents</button>' : '') +
      '<a class="btn doc-open" href="' + esc(d.url) + '" target="_blank" rel="noopener">Open in Drive ↗</a></div>' +
      '<iframe class="doc-frame" src="https://drive.google.com/file/d/' + encodeURIComponent(d.id) + '/preview" title="' + esc(d.name) + '"></iframe>' +
      '<p class="note">Not showing? Sign in to Google with an account that can open this file, or use “Open in Drive”.</p></div>';
    $('modal').hidden = false;
    document.body.style.overflow = 'hidden';
    $('modal').querySelector('.m-close').focus();
  }

  /* Popup listing every document that is not BA or Scope; click one to preview it. */
  function openOthers(i) {
    var list = docsFor(i, 'other'), name = S.model.depts[i];
    $('modal').innerHTML = '<div class="m-box card sd-others" role="dialog" aria-modal="true" aria-label="Other documents for ' + esc(name) + '">' +
      '<button type="button" class="m-close" aria-label="Close">✕</button>' +
      '<div class="m-head"><span class="sd-others-ic" aria-hidden="true">📁</span><div><div class="d-no">' + esc(name) + '</div>' +
      '<h2>Other documents <em class="sd-count">' + list.length + '</em></h2></div></div>' +
      (list.length > 6 ? '<input class="sd-others-q" type="search" placeholder="Find a document" aria-label="Find a document">' : '') +
      '<div class="sd-others-list">' + (list.length ? list.map(function (d, k) {
        return '<button type="button" class="sd-file" data-doc-dept="' + i + '" data-doc-type="other" data-doc-idx="' + k + '" data-name="' + esc(norm(d.name)) + '">' +
          '<span class="sd-file-ic" aria-hidden="true">📄</span><span class="sd-file-txt"><b>' + esc(d.name) + '</b>' +
          (d.updated ? '<small>Updated ' + fmtDay(d.updated) + '</small>' : '') + '</span><span class="sd-file-go" aria-hidden="true">View →</span></button>';
      }).join('') : '<p class="note">No other documents in this department folder yet.</p>') + '</div></div>';
    $('modal').hidden = false;
    document.body.style.overflow = 'hidden';
    var q = $('modal').querySelector('.sd-others-q');
    if (q) q.addEventListener('input', function () {
      var t = norm(q.value);
      Array.prototype.forEach.call($('modal').querySelectorAll('.sd-others-list .sd-file'), function (b) {
        b.hidden = t && b.getAttribute('data-name').indexOf(t) < 0;
      });
    });
    $('modal').querySelector('.m-close').focus();
  }

  function closeDept() {
    if ($('modal').hidden) return;
    $('modal').hidden = true;
    $('modal').innerHTML = '';
    document.body.style.overflow = '';
  }

  function deptTable(rows) {
    var yn = function (b) { return b ? '✓ Yes' : 'No'; };
    return '<div class="card table-wrap"><table><thead><tr><th>No</th><th>Department</th><th>BA %</th><th>Scope doc</th><th>Approval</th>' +
      '<th>Availability</th><th>Functionality</th><th>In use</th><th>Reports done</th><th>Expected</th><th>Reports %</th><th>Issues raised</th><th>Resolved</th><th>Open</th></tr></thead><tbody>' +
      rows.map(function (r) {
        return '<tr><td>' + r.no + '</td><td>' + esc(r.name) + '</td><td>' + pct(r.ba) + '</td><td>' + yn(r.scope) + '</td><td>' + yn(r.appr) + '</td><td>' +
          pct(r.avail) + '</td><td>' + pct(r.func) + '</td><td>' + pct(r.inUse) + '</td><td>' + fmt(r.done) + '</td><td>' + fmt(r.exp) + '</td><td>' + pct(r.rPct) +
          '</td><td>' + fmt(r.all) + '</td><td>' + fmt(r.res) + '</td><td>' + fmt(r.open) + '</td></tr>';
      }).join('') + '</tbody></table></div>';
  }

  function renderDeptBody() {
    var rows = deptRows(), host = $('dept-body');
    if (!host) return;
    if (!rows.length) { host.innerHTML = '<div class="card state">No department matches “' + esc(S.q) + '”.</div>'; return; }
    host.innerHTML = S.view === 'table' ? deptTable(rows) : '<div class="depts">' + rows.map(deptCard).join('') + '</div>';
  }

  function renderDeptShell() {
    var opt = function (v, t) { return '<option value="' + v + '"' + (S.sort === v ? ' selected' : '') + '>' + t + '</option>'; };
    var tog = function (v, t) { return '<button type="button" data-view="' + v + '" aria-pressed="' + (S.view === v) + '">' + t + '</button>'; };
    return '<div class="sec-head"><div><h2>Departments</h2><p>Status of each department as of ' + esc(S.model.weeks[S.sel].label) + '.</p></div>' +
      '<div class="tools"><input id="q" type="search" placeholder="Find a department" aria-label="Find a department" value="' + esc(S.q) + '">' +
      '<select id="sort" aria-label="Sort departments">' + opt('no', 'Sheet order') + opt('ba', 'BA progress: lowest first') + opt('issues', 'Open issues: most first') +
      opt('reports', 'Reports: lowest first') + opt('inuse', 'In use: lowest first') + '</select>' +
      '<div class="toggle" id="view">' + tog('cards', 'Cards') + tog('table', 'Table') + '</div></div></div><div id="dept-body"></div>';
  }

  /* ================= page ================= */
  function renderWeeks() {
    var m = S.model;
    return '<nav class="weeks" aria-label="Week"><span class="weeks-label">Week</span><div class="seg" id="seg">' + m.weeks.map(function (wk, w) {
      var ok = m.weekHasData[w];
      return '<button type="button" data-w="' + w + '" aria-pressed="' + (w === S.sel) + '"' + (ok ? '' : ' disabled title="Not updated yet"') + '><b>' + esc(wk.label) + '</b><span>' +
        (ok ? esc(niceDate(wk.date)) || '&nbsp;' : 'not updated') + '</span></button>';
    }).join('') + '</div></nav>';
  }

  /* ================= Tab 1: executive summary (always the latest updated week) ================= */
  var OVERALL = { id: 'c-overall', title: 'Overall progress', unit: '%', max: 100, h: 300, area: true, own: true, labelAll: true, bold: true, series: [{ name: 'Overall progress', key: 'overall', color: 'var(--s1)' }] };
  /* The chart takes the height of its card, so it always fills the space beside the department table. */
  function drawOverall() {
    if (!S.model) return;
    var host = $('c-overall');
    OVERALL.h = host && host.clientHeight > 150 ? host.clientHeight : 300;
    OVERALL.sel = at(S.model.T.overall, S.model.latest).w;
    OVERALL.fc = forecast();
    var eta = $('eta');
    if (eta) eta.innerHTML = OVERALL.fc.html;
    drawChart(OVERALL);
  }

  /* Linear regression Y = a + bX over every filled-in week (X = week index, Y = overall %),
     then solve Y = 100 for the estimated completion week. */
  function forecast() {
    var m = S.model, pts = [];
    m.T.overall.forEach(function (v, w) { if (isNum(v)) pts.push([w, Math.round(v)]); });   // whole percentages, as shown on the chart
    var note = function (t) { return { b: 0, html: '<span class="eta-t">🏁 ' + t + '</span>' }; };
    if (pts.length < 3) return note('Forecast after 3 weeks of data');
    var k = pts.length, mx = 0, my = 0, sxy = 0, sxx = 0;
    pts.forEach(function (p) { mx += p[0] / k; my += p[1] / k; });
    pts.forEach(function (p) { sxy += (p[0] - mx) * (p[1] - my); sxx += (p[0] - mx) * (p[0] - mx); });
    var b = sxx ? sxy / sxx : 0, a = my - b * mx, last = pts[k - 1][0];
    if (pts[k - 1][1] >= 100) return note('Completed');
    if (b <= 0) return note('No upward trend yet');
    var tw = (100 - a) / b, left = Math.max(0, Math.ceil(tw - last)), when = '';
    var d0 = /^(\d{4})-(\d{2})-(\d{2})$/.exec(m.weeks[0].date || '');
    if (d0) {
      var d = new Date(Number(d0[1]), Number(d0[2]) - 1, Number(d0[3]) + Math.round(tw * 7));
      when = ' · ' + d.getDate() + ' ' + MONTHS[d.getMonth()] + ' ' + d.getFullYear();
    }
    return { a: a, b: b, last: last,
      html: '<span class="eta-t">🏁 Est. completion: ~Week ' + (Math.ceil(tw) + 1) + when + '</span>' +
        '<span class="eta-s">' + left + ' week' + (left === 1 ? '' : 's') + ' left · +' + fmt(b, 1) + '% / week</span>' };
  }
  function level(v) { return !isNum(v) ? null : ZONES[zoneOf(v)]; }
  function cellMark(z, text) { return z ? '<span class="x-m ' + z.cls + '"><i>' + z.mark + '</i>' + text + '</span>' : '–'; }

  function renderSummary(animate) {
    var m = S.model, T = m.T, sel = m.latest, n = m.depts.length, wk = m.weeks[sel];
    $('subtitle').textContent = 'Current situation as of ' + wk.label + (wk.date ? ' · ' + niceDate(wk.date, true) : '') + ' · ' + n + ' departments';
    var mean = function (a) { return a.length ? a.reduce(function (x, y) { return x + y; }, 0) / a.length : null; };
    var rows = m.depts.map(function (name, i) {
      var v = function (k) { return deptVal(k, i, sel); };
      var exp = v('rExp'), done = isNum(exp) ? (isNum(v('rDone')) ? v('rDone') : 0) : null;
      var all = v('iAll'), res = isNum(all) ? (isNum(v('iDone')) ? v('iDone') : 0) : null;
      return {
        no: i + 1, name: name, overall: mean([v('avail'), v('func'), v('inUse')].filter(isNum)), ba: v('baPct'), inUse: v('inUse'),
        scope: isYes(v('scope')), appr: isYes(v('appr')), rPct: isNum(exp) && exp > 0 ? done / exp * 100 : null,
        all: all, res: res, open: isNum(all) ? all - res : null
      };
    });
    var cur = function (key) { return at(T[key], sel); };
    var overall = cur('overall'), avail = cur('avail'), func = cur('func'), inUse = cur('inUse');
    var z = level(overall.v) || ZONES[0];
    var statusText = ['Needs attention', 'In progress', 'On track'][ZONES.indexOf(z)];
    var counts = [0, 0, 0];
    rows.forEach(function (r) { if (isNum(r.inUse)) counts[zoneOf(r.inUse)]++; });
    var old = [['Availability', avail], ['Functionality', func], ['In use', inUse]].filter(function (p) { return p[1].w >= 0 && p[1].w < sel; });

    /* Half-circle gauge: a thin outer band marks the three zones, the thick arc is the overall figure. */
    var gauge = function (v) {
      var p = isNum(v) ? Math.max(0, Math.min(100, v)) : 0;
      var arc = function (r) { return 'M' + (110 - r) + ' 108 A' + r + ' ' + r + ' 0 0 1 ' + (110 + r) + ' 108'; };
      var band = function (cls, a, b) { return '<path class="' + cls + '" d="' + arc(98) + '" pathLength="100" fill="none" stroke="var(--z)" stroke-width="4" stroke-dasharray="' + (b - a) + ' 100" stroke-dashoffset="' + (-a) + '"/>'; };
      return '<svg viewBox="0 0 220 122" role="img" aria-label="Overall progress ' + pct(v) + '">' +
        band('z-risk', 0, 49.4) + band('z-adopt', 50, 89.4) + band('z-live', 90, 100) +
        '<path d="' + arc(82) + '" fill="none" stroke="var(--track)" stroke-width="16" stroke-linecap="round"/>' +
        '<path class="g-val" d="' + arc(82) + '" pathLength="100" fill="none" stroke="var(--z)" stroke-width="16" stroke-linecap="round" stroke-dasharray="' + p + ' 100"/>' +
        '<text x="12" y="121" text-anchor="middle">0</text><text x="208" y="121" text-anchor="middle">100</text></svg>';
    };
    var totalZ = counts[0] + counts[1] + counts[2];
    var comp = function (label, v) { return '<div class="x-comp"><span>' + label + '</span>' + meter(v) + '<b>' + pct(v) + '</b></div>'; };
    var banner = '<section class="card x-overall">' +
      '<div class="x-lab">Overall ERP progress</div>' +
      '<div class="x-overall-body"><div class="x-overall-left">' +
        '<div class="x-comps">' + comp('Availability', avail.v) + comp('Functionality', func.v) + comp('In use', inUse.v) + '</div>' +
      '</div><div class="x-overall-right">' +
        '<div class="x-gauge ' + z.cls + '">' + gauge(overall.v) + '<div class="x-gnum"><span class="x-big">' + pct(overall.v) + '</span></div></div>' +
        '<div class="x-gmeta"><span class="x-tag ' + z.cls + '"><i class="zdot"></i>' + statusText + '</span><span class="x-d">' + deltaTag(T.overall, overall, { unit: '%' }) + '</span></div>' +
      '</div></div><div class="x-overall-usage">' +
        (old.length ? '<p class="x-note">' + old.map(function (p) { return p[0]; }).join(', ') + ' not yet updated for ' + esc(wk.label) + '; last figures used.</p>' : '') +
      '</div>' +
      '</section>';

    var fig = function (emoji, label, value, small, sub, delta, cls, tone, bar) {
      return '<div class="card x-fig tone-' + (tone || 'blue') + '"><div class="t-top"><div class="x-lab">' + label + '</div><span class="t-emoji" aria-hidden="true">' + icon(emoji) + '</span></div>' +
        '<div class="x-val' + (cls ? ' ' + cls : '') + '">' + value + (small ? '<small>' + small + '</small>' : '') + '</div>' +
        (bar === undefined ? '' : meter(bar, 'var(--tone)')) +
        '<div class="x-sub">' + sub + '</div>' + (delta ? '<div class="x-d">' + delta + '</div>' : '') + '</div>';
    };
    var baYes = cur('baYes'), baAvg = cur('baAvg'), rDone = cur('rDone'), rExp = cur('rExp'), rPct = cur('rPct');
    var iAll = cur('iAll'), iDone = cur('iDone'), scope = cur('scope'), appr = cur('appr');
    var figs = '<div class="x-figs">' +
      fig('🏢', 'Total departments', n, '', 'in the ERP rollout scope', '', '', 'slate') +
      fig('✅', 'BA completed', fmt(baYes.v), 'of ' + n, 'average BA progress ' + pct(baAvg.v), deltaTag(T.baAvg, baAvg, { unit: '%' }), 'c-good', 'green', baAvg.v) +
      fig('📊', 'Reports completed', fmt(rDone.v), 'of ' + fmt(rExp.v), pct(rPct.v) + ' of all reports', deltaTag(T.rPct, rPct, { unit: '%' }), 'c-blue', 'blue', rPct.v) +
      fig('🚨', 'Current issues', fmt(iAll.v), '', fmt(iDone.v) + ' resolved · ' + (isNum(iAll.v) ? fmt(iAll.v - iDone.v) : '–') + ' open', deltaTag(T.iAll, iAll, { unit: '%', rel: true, downIsGood: true }), 'c-crit', 'red', cur('iPct').v) +
      fig('📄', 'Scope documents', fmt(scope.v), 'of ' + n, 'prepared · ' + fmt(appr.v) + ' approved', m.fixed.scope ? '<span class="dpill"><span class="delta flat">current status</span></span>' : deltaTag(T.scopePct, cur('scopePct'), { unit: '%' }), 'c-amber', 'amber', cur('scopePct').v) +
      fig('⚙️', 'ERP in use', pct(inUse.v), '', 'average across departments', deltaTag(T.inUse, inUse, { unit: '%' }), 'c-teal', 'teal', inUse.v) +
      '</div>';

    /* A small status dot followed by the figure in plain text. */
    var stat = function (zz, text) { return zz ? '<span class="x-s ' + zz.cls + '"><i class="zdot"></i>' + text + '</span>' : '<span class="x-s x-na">–</span>'; };
    var matrix = '<div class="table-wrap"><table class="x-table"><thead><tr><th>Department</th><th>Overall progress</th><th>BA</th><th>Scope document</th><th>In use</th><th>Reports</th><th>Open issues</th></tr></thead><tbody>' +
      rows.map(function (r) {
        var zo = level(r.overall), baZ = !isNum(r.ba) ? null : r.ba >= 100 ? ZONES[2] : r.ba > 0 ? ZONES[1] : ZONES[0];
        var scZ = r.appr ? ZONES[2] : r.scope ? ZONES[1] : ZONES[0];
        var isZ = !isNum(r.open) ? null : r.open === 0 ? ZONES[2] : ZONES[1];
        return '<tr tabindex="0" data-no="' + r.no + '"><td><span>' + r.no + '</span>' + esc(r.name) + '</td>' +
          '<td><div class="x-ov">' + meter(r.overall, zo ? 'var(--z)' : '').replace('class="meter"', 'class="meter ' + (zo ? zo.cls : '') + '"') + '<b>' + pct(r.overall) + '</b></div></td>' +
          '<td>' + stat(baZ, pct(r.ba)) + '</td><td>' + stat(scZ, r.appr ? 'Approved' : r.scope ? 'Prepared' : 'Pending') + '</td>' +
          '<td>' + stat(level(r.inUse), pct(r.inUse)) + '</td><td>' + stat(level(r.rPct), pct(r.rPct)) + '</td>' +
          '<td>' + stat(isZ, isNum(r.open) ? fmt(r.open) + ' of ' + fmt(r.all) : '') + '</td></tr>';
      }).join('') + '</tbody></table></div>' +
      '<div class="x-key">' + ZONES.slice().reverse().map(function (zz, k) { return '<span class="x-s ' + zz.cls + '"><i class="zdot"></i>' + ['Good: 90%+ or done', 'In progress: 50–89%', 'Needs attention: below 50% or pending'][k] + '</span>'; }).join('') + '</div>';

    /* Week-on-week movement: each headline figure this week against the week before. */
    var oPrev = before(T.overall, overall.w);
    var mvRow = function (label, key, unit, downGood) {
      var c = at(T[key], sel), p = before(T[key], c.w);
      var d = isNum(c.v) && isNum(p.v) ? c.v - p.v : null;
      if (d !== null) d = Math.round(c.v) - Math.round(p.v);      // change between the whole numbers shown
      var cls = d === null || d === 0 ? 'flat' : (downGood ? d < 0 : d > 0) ? 'good' : 'bad';
      var txt = d === null ? '–' : d === 0 ? '● 0' + unit : (d > 0 ? '▲ ' : '▼ ') + fmt(Math.abs(d)) + unit;
      return '<tr><td>' + label + '</td><td>' + (isNum(p.v) ? fmt(p.v) + unit : '–') + '</td><td><b>' + (isNum(c.v) ? fmt(c.v) + unit : '–') + '</b></td><td class="delta ' + cls + '">' + txt + '</td></tr>';
    };
    var moves = '<section class="card x-card x-move"><div class="x-lab">Week-on-week movement</div>' +
      '<div class="table-wrap"><table class="x-mtable"><thead><tr><th>Measure</th><th>' + (oPrev.w >= 0 ? esc(m.weeks[oPrev.w].label) : 'Previous') + '</th><th>' + esc(wk.label) + '</th><th>Change</th></tr></thead><tbody>' +
      mvRow('Overall progress', 'overall', '%') + mvRow('Availability', 'avail', '%') + mvRow('Functionality', 'func', '%') + mvRow('In use', 'inUse', '%') +
      mvRow('BA average', 'baAvg', '%') + mvRow('Reports completed', 'rPct', '%') + mvRow('Open issues', 'iOpen', '', true) +
      '</tbody></table></div></section>';

    $('app').innerHTML = viewSwitch() + '<div class="x-sum"><div class="x-top">' + banner + figs + '</div>' +
      '<section class="card x-card chart x-wide"><div class="x-chead"><div><div class="x-lab">Overall progress, week by week</div><p>Average of Availability, Functionality and In use across all departments.</p></div><div class="x-eta" id="eta"></div></div><div class="plot" id="c-overall"></div></section>' +
      moves +
      '<div class="x-main">' +
        '<section class="card x-card"><div class="x-lab">Department status</div><p>Click a row for the department\'s full detail.</p>' + matrix + '</section></div></div>' +
      (m.missing.length ? '<p class="note">Tab not found in the sheet: ' + esc(m.missing.join(', ')) + '. Its figures are shown as “–”.</p>' : '');
    drawOverall();
    bindViewSwitch();
    if (animate) countUp();
    var open = function (e) {
      var tr = e.target.closest('tr[data-no]');
      if (!tr) return;
      if (e.type === 'keydown') { if (e.key !== 'Enter' && e.key !== ' ') return; e.preventDefault(); }
      openDept(Number(tr.getAttribute('data-no')), sel);
    };
    $('app').querySelector('.x-table').addEventListener('click', open);
    $('app').querySelector('.x-table').addEventListener('keydown', open);
  }

  /* ================= Executive Summary > Task Breakdown (Dept Tasks + Dept Summary tabs) ================= */
  function viewSwitch() {
    return '<div class="x-viewbar"><div class="toggle" id="x-view" role="group" aria-label="Executive Summary view">' +
      '<button type="button" data-v="overview" aria-pressed="' + (S.sumView !== 'tasks') + '">' + icon('📊', 15) + ' Overview</button>' +
      '<button type="button" class="x-hl" data-v="tasks" aria-pressed="' + (S.sumView === 'tasks') + '">' + icon('🗂️', 15) + ' Task Breakdown</button></div>' +
      (S.sumView !== 'tasks' ? '<button type="button" class="x-hint"><span aria-hidden="true">←</span> View each department&#39;s Main, Sub &amp; Final tasks</button>' : S.tkHide ? allDeptStrip() : '') + '</div>';
  }
  /* Task Breakdown: one overall progress (average of every department) on the right of the view bar. */
  function allDeptStrip() {
    var ds = S.tasks || [];
    var vals = ds.map(function (x) {
      var p = numOf(x.progress);
      if (p === null) { var v = [numOf(x.a), numOf(x.f), numOf(x.u)].filter(function (y) { return y !== null; }); p = v.length ? v.reduce(function (a, b) { return a + b; }, 0) / v.length : null; }
      return p;
    }).filter(isNum);
    if (!vals.length) return '';
    var avg = vals.reduce(function (a, b) { return a + b; }, 0) / vals.length, z = level(avg);
    var col = avg >= 90 ? '#16a34a' : avg >= 50 ? '#f59e0b' : '#ef4444', w = Math.max(0, Math.min(100, avg));
    return '<div class="x-allp ' + (z ? z.cls : '') + '" style="--z:' + col + '" title="Average Department Progress of ' + vals.length + ' departments">' +
      '<span class="x-allp-l">Overall Departments Progress</span>' +
      '<i class="x-allp-bar"><s style="width:' + Math.max(2, w) + '%"></s><u style="left:' + w + '%"></u></i><b>' + Math.round(avg) + '%</b></div>';
  }
  /* Has this viewer opened Task Breakdown before? (hides the pointer hint once they have) */
  function tkSeen(set) {
    try { if (set) localStorage.setItem('erp.tkSeen', '1'); return localStorage.getItem('erp.tkSeen') === '1'; } catch (e) { return !!set; }
  }
  function bindViewSwitch() {
    var hint = $('app').querySelector('.x-hint');
    if (hint) hint.addEventListener('click', function () { S.sumView = 'tasks'; tkSeen(true); render(false); });
    var el = $('x-view');
    if (el) el.addEventListener('click', function (e) {
      var b = e.target.closest('button[data-v]');
      if (!b || b.getAttribute('data-v') === (S.sumView || 'overview')) return;
      S.sumView = b.getAttribute('data-v');
      if (S.sumView === 'tasks') tkSeen(true);
      render(false);
    });
  }
  /* Dept Tasks rows grouped per department into Main > Sub > Final. */
  function taskModel(data) {
    var t = data && data.deptTasks, sm = data && data.deptSummary;
    if (!t || t.length < 2) return null;
    var H = t[0].map(function (h) { return String(h).toLowerCase().trim(); });
    var c = function (name) { return H.indexOf(name); };
    var C = { dept: c('department'), cat: c('task category'), mt: c('m/t number'), st: c('s/t number'), ft: c('f/t number'), task: c('task') };
    ['availability', 'functionality', 'in use'].forEach(function (k, i) {
      var key = ['a', 'f', 'u'][i];
      C[key] = c(k + ' %'); C[key + 'D'] = c(k + ' date'); C[key + 'W'] = c(k + ' week no');
    });
    var g = function (row, k) { return C[k] === undefined || C[k] < 0 ? '' : String(row[C[k]] || '').trim(); };
    var depts = [], byName = {};
    t.slice(1).forEach(function (row) {
      var name = g(row, 'dept'); if (!name) return;
      var d = byName[name];
      if (!d) { d = byName[name] = { name: name, hod: '', progress: '', a: '', f: '', u: '', mains: [] }; depts.push(d); }
      var cat = g(row, 'cat').toLowerCase();
      var item = { task: g(row, 'task'), no: g(row, 'ft') || g(row, 'st') || g(row, 'mt'), vals: ['a', 'f', 'u'].map(function (k) { return { v: g(row, k), date: g(row, k + 'D'), week: g(row, k + 'W') }; }) };
      var blank = function () { return { task: '', no: '', vals: [{}, {}, {}].map(function () { return { v: '', date: '', week: '' }; }) }; };
      var main = d.mains[d.mains.length - 1];
      if (/main/.test(cat)) { item.subs = []; d.mains.push(item); }
      else if (/sub/.test(cat)) { if (!main) { main = blank(); main.subs = []; d.mains.push(main); } item.finals = []; main.subs.push(item); }
      else if (/final/.test(cat)) {
        if (!main) { main = blank(); main.subs = []; d.mains.push(main); }
        var sub = main.subs[main.subs.length - 1];
        if (!sub) { sub = blank(); sub.finals = []; main.subs.push(sub); }
        sub.finals.push(item);
      }
    });
    if (sm && sm.length > 1) {
      var SH = sm[0].map(function (h) { return String(h).toLowerCase().trim(); });
      sm.slice(1).forEach(function (row) {
        var d = byName[String(row[SH.indexOf('department')] || '').trim()];
        if (!d) return;
        var s = function (n) { var i = SH.indexOf(n); return i < 0 ? '' : String(row[i] || '').trim(); };
        d.hod = s('hod name'); d.progress = s('department progress'); d.a = s('availability %'); d.f = s('functionality %'); d.u = s('in use %');
      });
    }
    return depts;
  }
  function numOf(s) { var n = parseFloat(String(s).replace('%', '')); return isNaN(n) ? null : n; }
  function zMeter(v) { var z = level(v); return meter(v, z ? 'var(--z)' : '').replace('class="meter"', 'class="meter ' + (z ? z.cls : '') + '"'); }

  /* Small progress ring (used in the task list and the task map). */
  function ring(v, size) {
    var n = isNum(v) ? Math.max(0, Math.min(100, v)) : 0, z = level(v), r = 8;
    return '<svg class="tk-ring ' + (z ? z.cls : '') + '" viewBox="0 0 20 20" width="' + (size || 20) + '" height="' + (size || 20) + '" aria-hidden="true">' +
      '<circle cx="10" cy="10" r="' + r + '" fill="none" stroke="var(--track)" stroke-width="3"/>' +
      (n > 0 ? '<circle cx="10" cy="10" r="' + r + '" fill="none" stroke="var(--z, var(--s1))" stroke-width="3" stroke-linecap="round" pathLength="100" stroke-dasharray="' + n + ' 100" transform="rotate(-90 10 10)"/>' : '') + '</svg>';
  }

  /* Task map: All departments → department → Main → Sub → Final, like the document map. */
  var TM_COL = [0, 210, 470, 790, 1110], TM_W = [180, 230, 290, 290, 280], TM_H = 44, TM_ROW = 54;
  function renderTaskMap(host) {
    var depts = S.tasks || [];
    var open = S.tmOpen || (S.tmOpen = { root: true });
    var view = S.tmView || (S.tmView = { x: 24, y: 24, k: 0.9 });
    var avg = function (a) { a = a.filter(isNum); return a.length ? a.reduce(function (x, y) { return x + y; }, 0) / a.length : null; };
    var valsPct = function (vals) { return avg(vals.map(function (v) { return numOf(v.v); })); };
    var deptP = function (d) { var p = numOf(d.progress); return p !== null ? p : avg([numOf(d.a), numOf(d.f), numOf(d.u)]); };
    var tree = { id: 'root', kind: 'root', label: 'All departments', sub: depts.length + ' departments', children: depts.map(function (d, di) {
      var nSub = 0, nFin = 0;
      d.mains.forEach(function (m) { nSub += m.subs.length; m.subs.forEach(function (s) { nFin += s.finals.length; }); });
      return { id: 'd' + di, kind: 'dept', label: d.name, p: deptP(d), sub: d.mains.length + ' main · ' + nSub + ' sub · ' + nFin + ' final', children: d.mains.map(function (m, mi) {
        return { id: 'd' + di + 'm' + mi, kind: 'main', label: (m.no ? m.no + '  ' : '') + m.task, p: valsPct(m.vals), sub: m.subs.length + ' sub tasks', children: m.subs.map(function (s, si) {
          return { id: 'd' + di + 'm' + mi + 's' + si, kind: 'sub', label: (s.no ? s.no + '  ' : '') + s.task, p: valsPct(s.vals), sub: s.finals.length ? s.finals.length + ' final tasks' : 'no final tasks', children: s.finals.map(function (f, fi) {
            var ticks = f.vals.map(function (v) { return /^true$/i.test(v.v); });
            return { id: 'd' + di + 'm' + mi + 's' + si + 'f' + fi, kind: 'fin', label: (f.no ? f.no + '  ' : '') + f.task, ticks: ticks, children: [] };
          }) };
        }) };
      }) };
    }) };

    var nodes = [], links = [], y = 0;
    (function place(nd, depth) {
      nd.depth = depth; nd.x = TM_COL[depth];
      var kids = nd.children.length && open[nd.id] ? nd.children : [];
      if (!kids.length) { nd.y = y; y += TM_ROW; }
      else {
        kids.forEach(function (k) { place(k, depth + 1); });
        nd.y = (kids[0].y + kids[kids.length - 1].y) / 2;
        kids.forEach(function (k) { links.push([nd, k]); });
      }
      nodes.push(nd);
    })(tree, 0);
    var maxDepth = nodes.reduce(function (a, n) { return Math.max(a, n.depth); }, 0);
    var W = TM_COL[maxDepth] + TM_W[maxDepth] + 40, H = Math.max(y, TM_ROW);
    var paths = links.map(function (l) {
      var a = l[0], b = l[1], x1 = a.x + TM_W[a.depth], y1 = a.y + TM_H / 2, x2 = b.x, y2 = b.y + TM_H / 2, mx = (x1 + x2) / 2;
      return '<path d="M' + x1 + ' ' + y1 + 'C' + mx + ' ' + y1 + ' ' + mx + ' ' + y2 + ' ' + x2 + ' ' + y2 + '"/>';
    }).join('');
    var LETTERS = ['A', 'F', 'U'], NAMES = ['Availability', 'Functionality', 'In use'];
    var html = nodes.map(function (nd) {
      var has = nd.children.length, isOpen = !!open[nd.id];
      var pos = 'left:' + nd.x + 'px;top:' + nd.y + 'px;width:' + TM_W[nd.depth] + 'px;height:' + TM_H + 'px';
      var lead = nd.kind === 'root' ? '' : nd.kind === 'fin'
        ? '<span class="tm-ticks">' + nd.ticks.map(function (t, k) { return '<i class="' + (t ? 'on' : '') + '" title="' + NAMES[k] + (t ? ': done' : ': not done') + '">' + LETTERS[k] + '</i>'; }).join('') + '</span>'
        : ring(nd.p, 22);
      var tail = nd.kind !== 'root' && nd.kind !== 'fin' ? '<em class="tm-p">' + pct(nd.p === null ? null : Math.round(nd.p)) + '</em>' : '';
      return '<div class="dm-node tm-' + nd.kind + (has ? ' can-open' : '') + '" style="' + pos + '"' + (has ? ' data-tm="' + nd.id + '" role="button" tabindex="0" aria-expanded="' + isOpen + '"' : '') + ' title="' + esc(nd.label) + '">' +
        lead + '<span class="dm-txt"><b>' + esc(nd.label) + '</b>' + (nd.sub ? '<small>' + esc(nd.sub) + '</small>' : '') + '</span>' + tail +
        (has ? '<span class="dm-tog" aria-hidden="true">' + (isOpen ? '‹' : '›') + '</span>' : '') + '</div>';
    }).join('');

    host.innerHTML = '<div class="card dm tm">' +
      '<div class="dm-bar"><span class="dm-hint">Click a box to open or close it · drag to move · scroll to zoom · A / F / U = Availability, Functionality, In use</span>' +
      '<div class="dm-tools"><button type="button" class="btn" data-tm-act="depts">Departments</button><button type="button" class="btn" data-tm-act="mains">Show main tasks</button><button type="button" class="btn" data-tm-act="subs">Show sub tasks</button>' +
      '<button type="button" class="btn dm-z" data-tm-act="out" aria-label="Zoom out">−</button><button type="button" class="btn dm-z" data-tm-act="in" aria-label="Zoom in">+</button>' +
      '<button type="button" class="btn" data-tm-act="reset">Reset view</button></div></div>' +
      '<div class="dm-view" id="tm-view"><div class="dm-canvas" id="tm-canvas" style="width:' + W + 'px;height:' + H + 'px">' +
      '<svg class="dm-links" width="' + W + '" height="' + H + '" aria-hidden="true">' + paths + '</svg>' + html + '</div></div></div>';

    var canvas = $('tm-canvas'), vp = $('tm-view');
    var apply = function () { canvas.style.transform = 'translate(' + view.x + 'px,' + view.y + 'px) scale(' + view.k + ')'; };
    var zoom = function (f, cx, cy) {
      var k = Math.max(0.3, Math.min(2, view.k * f));
      if (cx === undefined) { var b = vp.getBoundingClientRect(); cx = b.width / 2; cy = b.height / 2; }
      view.x = cx - (cx - view.x) * k / view.k; view.y = cy - (cy - view.y) * k / view.k; view.k = k; apply();
    };
    apply();
    var toggle = function (id) { open[id] = !open[id]; renderTaskMap(host); };
    vp.addEventListener('click', function (e) { var nd = e.target.closest('[data-tm]'); if (nd && !vp.dragged) toggle(nd.getAttribute('data-tm')); });
    vp.addEventListener('keydown', function (e) {
      var nd = e.target.closest('[data-tm]');
      if (nd && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); toggle(nd.getAttribute('data-tm')); }
    });
    var drag = null;
    vp.addEventListener('pointerdown', function (e) { if (e.button !== 0) return; drag = { x: e.clientX, y: e.clientY, vx: view.x, vy: view.y }; vp.dragged = false; });
    var move = function (e) {
      if (!drag) return;
      var dx = e.clientX - drag.x, dy = e.clientY - drag.y;
      if (Math.abs(dx) + Math.abs(dy) > 4) { vp.dragged = true; vp.classList.add('grabbing'); }
      if (vp.dragged) { view.x = drag.vx + dx; view.y = drag.vy + dy; apply(); }
    };
    var up = function () { drag = null; vp.classList.remove('grabbing'); setTimeout(function () { vp.dragged = false; }, 0); };
    window.removeEventListener('pointermove', S.tmMove || move); window.removeEventListener('pointerup', S.tmUp || up);
    S.tmMove = move; S.tmUp = up;
    window.addEventListener('pointermove', move); window.addEventListener('pointerup', up);
    vp.addEventListener('click', function (e) { if (vp.dragged) { e.stopPropagation(); e.preventDefault(); } }, true);
    vp.addEventListener('wheel', function (e) {
      e.preventDefault();
      var b = vp.getBoundingClientRect();
      zoom(e.deltaY < 0 ? 1.1 : 1 / 1.1, e.clientX - b.left, e.clientY - b.top);
    }, { passive: false });
    host.querySelector('.dm-tools').addEventListener('click', function (e) {
      var b = e.target.closest('[data-tm-act]');
      if (!b) return;
      var act = b.getAttribute('data-tm-act');
      if (act === 'in') zoom(1.2);
      else if (act === 'out') zoom(1 / 1.2);
      else if (act === 'reset') { S.tmView = null; renderTaskMap(host); }
      else {
        var o = { root: true };
        tree.children.forEach(function (d) {
          if (act !== 'depts') o[d.id] = true;
          if (act === 'subs') d.children.forEach(function (m) { o[m.id] = true; });
        });
        S.tmOpen = o; S.tmView = null; renderTaskMap(host);
      }
    });
  }

  /* List / Map switch for Task Breakdown */
  function tkModeSwitch() {
    return '<div class="toggle tk-mode" role="group" aria-label="Task view"><button type="button" data-tk-mode="list" aria-pressed="' + (S.tkMode !== 'map') + '">☰ List</button><button type="button" data-tk-mode="map" aria-pressed="' + (S.tkMode === 'map') + '">⟡ Map</button></div>';
  }
  function bindTkMode() {
    Array.prototype.forEach.call($('app').querySelectorAll('[data-tk-mode]'), function (b) {
      b.addEventListener('click', function () { if (S.tkMode === b.getAttribute('data-tk-mode') || (!S.tkMode && b.getAttribute('data-tk-mode') === 'list')) return; S.tkMode = b.getAttribute('data-tk-mode'); renderTasks(); });
    });
  }
  function renderTasks() {
    var depts = S.tasks;
    $('subtitle').textContent = 'Task breakdown by department · Main, Sub and Final tasks as entered by each department';
    if (!depts || !depts.length) {
      $('app').innerHTML = viewSwitch() + '<div class="card state"><b>No task data yet</b>The “Dept Tasks” tab was not found. Run syncDepartments in Apps Script and deploy the latest Code.gs.</div>';
      bindViewSwitch();
      return;
    }
    if (!S.tkDept || !depts.some(function (x) { return x.name === S.tkDept; })) S.tkDept = depts[0].name;
    S.tkOpen = S.tkOpen || {};
    var d = depts.filter(function (x) { return x.name === S.tkDept; })[0];
    var deptPct = function (x) {
      var p = numOf(x.progress);
      if (p === null) { var v = [numOf(x.a), numOf(x.f), numOf(x.u)].filter(function (y) { return y !== null; }); p = v.length ? v.reduce(function (a, b) { return a + b; }, 0) / v.length : null; }
      return p;
    };

    var side = '<aside class="card tk-side"><div class="tk-sidehead"><div class="x-lab">Departments</div><button type="button" class="tk-fold" data-tk-side="1" aria-label="' + (S.tkHide ? 'Show departments' : 'Hide departments') + '" title="' + (S.tkHide ? 'Show departments' : 'Hide departments') + '">' + (S.tkHide ? '›' : '‹') + '</button></div><div class="tk-list">' +
      '<button type="button" class="tk-dept tk-mapbtn" data-tk-map="1" aria-pressed="' + (S.tkMode === 'map') + '" title="Map of every department&#39;s tasks">' + icon('🧭', 17) + '<span class="tk-dn">All departments map</span></button>' +
      depts.map(function (x) {
        var p = deptPct(x);
        return '<button type="button" class="tk-dept tk-dept3" data-tk-dept="' + esc(x.name) + '" aria-pressed="' + (S.tkMode !== 'map' && x === d) + '" title="' + esc(x.name) + '">' +
          '<span class="tk-dn">' + esc(x.name) + '</span><span class="tk-dp">' + pct(p) + '</span>' + zMeter(p) + '</button>';
      }).join('') + '</div></aside>';
    /* One gauge on top: average progress of every department. */
    var allP = depts.map(deptPct).filter(isNum);
    var avgP = allP.length ? allP.reduce(function (a, b) { return a + b; }, 0) / allP.length : null;
    var gz = level(avgP), gn = isNum(avgP) ? Math.max(0, Math.min(100, avgP)) : 0;
    var garc = function (r) { return 'M' + (110 - r) + ' 108 A' + r + ' ' + r + ' 0 0 1 ' + (110 + r) + ' 108'; };
    var gband = function (cls, a, b) { return '<path class="' + cls + '" d="' + garc(98) + '" pathLength="100" fill="none" stroke="var(--z)" stroke-width="4" stroke-dasharray="' + (b - a) + ' 100" stroke-dashoffset="' + (-a) + '"/>'; };
    var gauge = '<div class="tk-allg"><div class="x-gauge ' + (gz ? gz.cls : '') + '"><svg viewBox="0 0 220 122" role="img" aria-label="All departments average ' + pct(avgP) + '">' +
      gband('z-risk', 0, 49.4) + gband('z-adopt', 50, 89.4) + gband('z-live', 90, 100) +
      '<path d="' + garc(82) + '" fill="none" stroke="var(--track)" stroke-width="16" stroke-linecap="round"/>' +
      (gn > 0 ? '<path class="g-val" d="' + garc(82) + '" pathLength="100" fill="none" stroke="var(--z)" stroke-width="16" stroke-linecap="round" stroke-dasharray="' + gn + ' 100"/>' : '') +
      '<text x="12" y="121" text-anchor="middle">0</text><text x="208" y="121" text-anchor="middle">100</text></svg>' +
      '<div class="x-gnum"><span class="x-big">' + pct(avgP === null ? null : Math.round(avgP)) + '</span></div></div>' +
      '<div class="tk-allcap">Overall Departments Progress</div></div>';
    side = side.replace('<div class="tk-list">', gauge + '<div class="tk-list">');
    if (S.tkHide) {
      var rz = level(avgP);
      var rail = '<div class="tk-rail"><div class="tk-railavg ' + (rz ? rz.cls : '') + '" data-tip="All departments average ' + pct(avgP === null ? null : Math.round(avgP)) + '" style="--p:' + gn + '"><b>' + (avgP === null ? '–' : Math.round(avgP)) + '</b></div>' +
        depts.map(function (x) {
          var p = deptPct(x), z = level(p);
          return '<button type="button" class="tk-ricon ' + (z ? z.cls : '') + '" data-tk-dept="' + esc(x.name) + '" aria-pressed="' + (S.tkMode !== 'map' && x === d) + '" aria-label="' + esc(x.name) + '" data-tip="' + esc(x.name) + ' · ' + pct(p === null ? null : Math.round(p)) + '">' + deptIcon(x.name) + '<i></i></button>';
        }).join('') + '</div>';
      side = side.replace('<div class="tk-list">', rail + '<div class="tk-list">');
    }

    var pill = function (label, v, tone) { return '<div class="tk-pill tone-' + tone + '"><span>' + label + '</span><b>' + esc(v || '–') + '</b></div>'; };
    var nSub = 0, nFin = 0;
    d.mains.forEach(function (m) { nSub += m.subs.length; m.subs.forEach(function (s) { nFin += s.finals.length; }); });
    /* Header (v2): name block · progress + rank card · three metric cards. */
    var me = deptPct(d), meR = isNum(me) ? Math.round(me) : null;
    var sorted = depts.map(deptPct).filter(isNum).sort(function (a, b) { return b - a; });
    var rank = isNum(me) ? sorted.indexOf(me) + 1 : null, diff = isNum(me) ? Math.round(me - (avgP || 0)) : null;
    var dcls = diff > 0 ? 'up' : diff < 0 ? 'down' : 'eq';
    var dtxt = diff === null ? '' : diff === 0 ? 'Same as average' : Math.abs(diff) + ' pts ' + (diff > 0 ? 'above' : 'below') + ' average';
    var bar = function (v) { var n = numOf(v); return '<i class="hd-bar"><s style="width:' + (n === null ? 0 : Math.max(2, Math.min(100, n))) + '%"></s></i>'; };
    var SVG = function (p) { return '<svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + p + '</svg>'; };
    var metric = function (cls, label, v, ic) {
      return '<div class="hd-m ' + cls + '"><span class="hd-mi">' + SVG(ic) + '</span><div class="hd-mt"><span>' + label + '</span><b>' + esc(v || '–') + '</b></div>' + bar(v) + '</div>';
    };
    var head = '<div class="tk-head hd">' +
      '<div class="hd-id"><div class="hd-crumb">ERP Implementation / <b>Task Breakdown</b></div>' +
        '<div class="hd-name"><span class="hd-ic">' + deptIcon(d.name) + '</span><div><h2>' + esc(d.name) + '</h2><p title="HOD: ' + esc(d.hod || '–') + '">HOD: <b>' + esc(d.hod || '–') + '</b></p></div></div>' +
        '<div class="tk-chips" role="group" aria-label="Show only">' +
        [['', 'All Tasks', ''], ['main', d.mains.length, 'Main'], ['sub', nSub, 'Sub'], ['final', nFin, 'Final']].map(function (c) {
          return '<button type="button" class="tk-chip tk-chip-' + (c[0] || 'all') + '" data-tk-filter="' + c[0] + '" aria-pressed="' + ((S.tkFilter || '') === c[0]) + '">' + (c[0] ? '<b>' + c[1] + '</b> ' + c[2] : c[1]) + '</button>';
        }).join('') + '</div></div>' +
      '<div class="hd-prog"><div class="hd-pl"><span class="hd-lab">Department Progress</span><b class="hd-big">' + pct(meR) + '</b>' + bar(me) +
        '<small>Overall Departments Progress: ' + pct(avgP === null ? null : Math.round(avgP)) + '</small></div>' +
        (rank ? '<div class="hd-pr"><div class="hd-rank"><span>#</span>' + rank + '<em>/ ' + sorted.length + ' Departments</em></div><span class="hd-diff ' + dcls + '">' + (diff > 0 ? '▲ ' : diff < 0 ? '▼ ' : '') + dtxt + '</span></div>' : '') +
      '</div>' +
      '<div class="hd-ms">' +
        metric('m-a', 'Availability', d.a, '<path d="m12 3 9 5-9 5-9-5Z"/><path d="m3 13 9 5 9-5"/>') +
        metric('m-f', 'Functionality', d.f, '<circle cx="12" cy="12" r="3"/><path d="M12 2v3M12 19v3M4.9 4.9l2.1 2.1M17 17l2.1 2.1M2 12h3M19 12h3M4.9 19.1 7 17M17 7l2.1-2.1"/>') +
        metric('m-u', 'In Use', d.u, '<rect x="3" y="4" width="18" height="12" rx="2"/><path d="M8 20h8M12 16v4"/>') +
      '</div></div>';

    var dw = function (v) { return !S.tkDates ? '' : '<td class="tk-d">' + esc(v.date) + '</td><td class="tk-w">' + esc(v.week) + '</td>'; };
    var pctCell = function (v) {
      var n = numOf(v.v);
      return '<td class="tk-v">' + (n === null ? '<span class="tk-na">–</span>' : '<div class="tk-pc">' + zMeter(n) + '<b>' + Math.round(n) + '%</b></div>') + '</td>' + dw(v);
    };
    var tickCell = function (v) {
      var on = /^true$/i.test(v.v);
      return '<td class="tk-v"><span class="tk-tick' + (on ? ' on' : '') + '" role="img" aria-label="' + (on ? 'Done' : 'Not done') + '">' + (on ? '✓' : '') + '</span></td>' + dw(v);
    };
    var rows = '', flt = S.tkFilter || '';
    if (flt) d.mains.forEach(function (m) {                                    // filtered: one flat list of that level only
      if (flt === 'main') rows += '<tr class="tk-main"><td>Main Task</td><td>' + esc(m.no) + '</td><td class="tk-t">' + esc(m.task) + '</td>' + m.vals.map(pctCell).join('') + '</tr>';
      m.subs.forEach(function (s) {
        if (flt === 'sub') rows += '<tr class="tk-sub"><td>Sub Task' + (s.finals.length ? ' <em>' + s.finals.length + '</em>' : '') + '</td><td>' + esc(s.no) + '</td><td class="tk-t">' + esc(s.task) + '</td>' + s.vals.map(pctCell).join('') + '</tr>';
        if (flt === 'final') s.finals.forEach(function (f) {
          rows += '<tr class="tk-fin"><td>Final Task</td><td>' + esc(f.no) + '</td><td class="tk-t">' + esc(f.task) + '</td>' + f.vals.map(tickCell).join('') + '</tr>';
        });
      });
    });
    else d.mains.forEach(function (m, mi) {
      var mk = d.name + '|' + mi, mOpen = S.tkOpen[mk] === true;
      rows += '<tr class="tk-main" data-tk-key="' + esc(mk) + '" tabindex="0" aria-expanded="' + mOpen + '"><td><span class="tk-car">' + (mOpen ? '▾' : '▸') + '</span>Main Task</td><td>' + esc(m.no) + '</td><td class="tk-t">' + esc(m.task) + '</td>' + m.vals.map(pctCell).join('') + '</tr>';
      if (!mOpen) return;
      m.subs.forEach(function (s, si) {
        var sk = mk + '|' + si, sOpen = S.tkOpen[sk] === true, has = s.finals.length;
        rows += '<tr class="tk-sub" data-tk-key="' + esc(sk) + '"' + (has ? ' tabindex="0" aria-expanded="' + sOpen + '"' : '') + '><td><span class="tk-car">' + (has ? (sOpen ? '▾' : '▸') : '') + '</span>Sub Task' + (has ? ' <em>' + has + '</em>' : '') + '</td><td>' + esc(s.no) + '</td><td class="tk-t">' + esc(s.task) + '</td>' + s.vals.map(pctCell).join('') + '</tr>';
        if (sOpen) s.finals.forEach(function (f) {
          rows += '<tr class="tk-fin"><td>Final Task</td><td>' + esc(f.no) + '</td><td class="tk-t">' + esc(f.task) + '</td>' + f.vals.map(tickCell).join('') + '</tr>';
        });
      });
    });
    // One header row: "Availability %" etc. With dates on, each group is "Availability % | Date | Week No".
    var table = '<div class="table-wrap tk-wrap"><table class="tk-table' + (S.tkDates ? ' tk-dated' : '') + '"><thead>' +
      '<tr><th>Task Category</th><th>No</th><th class="tk-t">Task</th>' + ['Availability', 'Functionality', 'In use'].map(function (h) {
        return '<th class="tk-grp tk-vh">' + h + ' %</th>' + (!S.tkDates ? '' : '<th class="tk-dh">Date</th><th class="tk-dh">Week No</th>');
      }).join('') + '</tr></thead><tbody>' +
      (rows || '<tr><td colspan="12" class="tk-na">No tasks entered yet.</td></tr>') + '</tbody></table></div>';

    $('app').innerHTML = viewSwitch() + '<div class="tk' + (S.tkHide ? ' tk-hide' : '') + '">' + side +
      (S.tkMode === 'map' ? '<section class="tk-mapbody"><div id="tm-host"></div></section></div>' :
      '<section class="card tk-body">' + head +
      '<div class="tk-tools"><label class="tk-switch"><input type="checkbox" data-tk-dates="1"' + (!S.tkDates ? '' : ' checked') + '> Show Date &amp; Week No</label><span class="tk-gap"></span><button type="button" class="btn" data-tk-all="1"><svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><path d="M14 4h6v6M10 20H4v-6M20 4l-7 7M4 20l7-7"/></svg> Expand all</button><button type="button" class="btn" data-tk-all="0"><svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><path d="M5 12h14"/></svg> Collapse all</button></div>' +
      table + '</section></div>');
    bindViewSwitch();
    if (S.tkMode === 'map') renderTaskMap($('tm-host'));
    var box = $('app').querySelector('.tk');
    var act = function (e) {
      if (e.type === 'keydown' && e.key !== 'Enter' && e.key !== ' ') return;
      var b = e.target.closest('[data-tk-dept]');
      if (b) { if (e.type === 'keydown') return; S.tkDept = b.getAttribute('data-tk-dept'); S.tkMode = 'list'; renderTasks(); return; }
      if (e.type === 'click' && e.target.closest('[data-tk-map]')) { S.tkMode = 'map'; renderTasks(); return; }
      var fb = e.target.closest('[data-tk-filter]');
      if (fb) { if (e.type === 'keydown') return; var fv = fb.getAttribute('data-tk-filter'); S.tkFilter = S.tkFilter === fv ? '' : fv; renderTasks(); return; }
      if (e.type === 'click' && e.target.closest('[data-tk-side]')) { S.tkHide = !S.tkHide; renderTasks(); return; }
      var all = e.target.closest('[data-tk-all]');
      if (all) {
        if (e.type === 'keydown') return;
        var open = all.getAttribute('data-tk-all') === '1'; S.tkFilter = '';
        d.mains.forEach(function (m, mi) { S.tkOpen[d.name + '|' + mi] = open; m.subs.forEach(function (s, si) { S.tkOpen[d.name + '|' + mi + '|' + si] = open; }); });
        renderTasks(); return;
      }
      var tr = e.target.closest('tr[aria-expanded]');
      if (!tr) return;
      if (e.type === 'keydown') e.preventDefault();
      var k = tr.getAttribute('data-tk-key');
      S.tkOpen[k] = tr.getAttribute('aria-expanded') !== 'true';
      var scroll = $('app').querySelector('.tk-wrap').scrollTop;
      renderTasks();
      $('app').querySelector('.tk-wrap').scrollTop = scroll;
      var again = Array.prototype.filter.call($('app').querySelectorAll('tr[data-tk-key]'), function (r) { return r.getAttribute('data-tk-key') === k; })[0];
      if (again) again.focus({ preventScroll: true });
    };
    box.addEventListener('change', function (e) { if (e.target.matches('[data-tk-dates]')) { S.tkDates = e.target.checked; try { localStorage.setItem('erp.tkDates', S.tkDates ? '1' : '0'); } catch (err) {} renderTasks(); } });
    box.addEventListener('click', act);
    box.addEventListener('keydown', act);
  }

  /* ================= Tab 3: Scope & Documentation Tracker ================= */
  var SD_STATUS = [
    { key: 'pending',  cls: 'z-risk',  label: 'Scope pending' },
    { key: 'awaiting', cls: 'z-adopt', label: 'Awaiting approval' },
    { key: 'approved', cls: 'z-live',  label: 'Approved' }
  ];
  function fmtDay(ms) { if (!ms) return ''; var d = new Date(ms); return d.getDate() + ' ' + MONTHS[d.getMonth()] + ' ' + d.getFullYear(); }

  /* One row per department: BA progress (BA tab) + scope status (Scope tab) + all files (Drive). */
  function scopeRows() {
    var m = S.model, sel = m.latest;
    return m.depts.map(function (name, i) {
      var prep = isYes(deptVal('scope', i, sel)), appr = isYes(deptVal('appr', i, sel));
      var baDone = isYes(deptVal('ba', i, sel)), baPct = deptVal('baPct', i, sel);
      var ba = docsFor(i, 'ba'), sc = docsFor(i, 'scope'), other = docsFor(i, 'other');
      return { i: i, no: i + 1, name: name, prep: prep || appr, appr: appr, baDone: baDone, baPct: baPct, ba: ba, sc: sc, other: other,
        status: SD_STATUS[appr ? 2 : prep ? 1 : 0], missing: (ba.length ? 0 : 1) + (sc.length ? 0 : 1) };
    });
  }

  function renderScope() {
    var rows = scopeRows(), n = rows.length;
    var count = function (f) { return rows.filter(f).length; };
    var total = function (k) { return rows.reduce(function (a, r) { return a + r[k].length; }, 0); };
    var prep = count(function (r) { return r.prep; }), appr = count(function (r) { return r.appr; });
    var baDone = count(function (r) { return r.baDone; });
    var baDepts = count(function (r) { return r.ba.length; }), scDepts = count(function (r) { return r.sc.length; });
    var baFiles = total('ba'), scFiles = total('sc'), otFiles = total('other');
    var byStatus = SD_STATUS.map(function (s) { return count(function (r) { return r.status === s; }); });
    var share = function (v, of) { return of ? v / of * 100 : 0; };
    var plural = function (v, w) { return v + ' ' + w + (v === 1 ? '' : 's'); };
    $('subtitle').textContent = 'BA and scope documents, approvals and department files · ' + n + ' departments';

    var card = function (tone, emoji, label, value, of, sub) {
      var cls = tone === 'green' ? 'good' : tone === 'red' ? 'crit' : tone;
      return '<div class="card x-fig sd-kpi tone-' + tone + '"><div class="t-top"><div class="x-lab">' + label + '</div><span class="t-emoji" aria-hidden="true">' + icon(emoji) + '</span></div>' +
        '<div class="x-val c-' + cls + '">' + value + (of !== null ? '<small>of ' + of + '</small>' : '') + '</div>' +
        (of !== null ? meter(share(value, of), 'var(--tone)') : '') + '<div class="x-sub">' + sub + '</div></div>';
    };
    var kpis = '<section class="sd-kpis">' +
      card('teal', '🧭', 'BA completed', baDone, n, plural(baDepts, 'department') + ' with BA documents') +
      card('amber', '📝', 'Scope documents prepared', prep, n, (n - prep) + ' still to prepare') +
      card('green', '✅', 'Scope documents approved', appr, n, (prep - appr) + ' awaiting approval') +
      '</section>';

    var f = S.sdFilter || 'all';
    var chip = function (key, label, num) { return '<button type="button" data-f="' + key + '" aria-pressed="' + (f === key) + '">' + label + ' <b>' + num + '</b></button>'; };
    var tools = '<div class="sec-head"><div><h2>Departments</h2><p>Click a 📄 file to preview it.</p></div>' +
      '<div class="tools"><input id="sd-q" type="search" placeholder="Find a department" aria-label="Find a department" value="' + esc(S.sdQ || '') + '">' +
      '<div class="toggle sd-filter" id="sd-filter">' + chip('all', 'All', n) + chip('pending', 'Scope pending', byStatus[0]) +
      chip('awaiting', 'Awaiting approval', byStatus[1]) + chip('approved', 'Approved', byStatus[2]) +
      chip('missing', 'Missing files', count(function (r) { return r.missing; })) + '</div>' +
      '<div class="toggle" id="sd-view"><button type="button" data-v="list" aria-pressed="' + (S.sdView !== 'map') + '">☰ List</button>' +
      '<button type="button" data-v="map" aria-pressed="' + (S.sdView === 'map') + '">🗺 Map</button></div></div></div>';

    $('app').innerHTML = '<div class="sd">' + kpis + tools + '<div id="sd-body"></div></div>';
    renderScopeBody();
    $('sd-view').addEventListener('click', function (e) {
      var b = e.target.closest('button[data-v]');
      if (!b) return;
      S.sdView = b.getAttribute('data-v');
      Array.prototype.forEach.call($('sd-view').children, function (c) { c.setAttribute('aria-pressed', String(c === b)); });
      renderScopeBody();
    });
    $('sd-q').addEventListener('input', function (e) { S.sdQ = e.target.value; renderScopeBody(); });
    $('sd-filter').addEventListener('click', function (e) {
      var b = e.target.closest('button[data-f]');
      if (!b) return;
      S.sdFilter = b.getAttribute('data-f');
      Array.prototype.forEach.call($('sd-filter').children, function (c) { c.setAttribute('aria-pressed', String(c === b)); });
      renderScopeBody();
    });
  }

  function renderScopeBody() {
    var host = $('sd-body');
    if (!host) return;
    var f = S.sdFilter || 'all', q = norm(S.sdQ);
    var rows = scopeRows().filter(function (r) {
      if (q && norm(r.name).indexOf(q) < 0) return false;
      return f === 'all' || (f === 'missing' ? r.missing > 0 : r.status.key === f);
    });
    if (!rows.length) { host.innerHTML = '<div class="card state">No department matches this filter.</div>'; return; }
    if (S.sdView === 'map') { renderDocMap(host, rows); return; }

    /* Compact summary: one row per department; click a row to see and open its documents. */
    var dots = function (r) {
      return '<span class="sd-ticks" aria-label="Progress">' + [r.baDone || r.ba.length > 0, r.prep, r.appr].map(function (d, k) {
        return '<span class="sd-tick' + (d ? ' done' : '') + '" title="' + ['BA', 'Scope prepared', 'Scope approved'][k] + (d ? ': done' : ': pending') + '">' +
          '<i>' + (d ? '✓' : '') + '</i>' + ['BA', 'Scope', 'Approved'][k] + '</span>';
      }).join('') + '</span>';
    };
    var count = function (n) { return '<span class="sd-n' + (n ? '' : ' zero') + '">📄 ' + n + '</span>'; };
    host.innerHTML = '<div class="card sd-list"><div class="table-wrap"><table class="sd-table"><thead><tr>' +
      '<th>Department</th><th>Progress</th><th>BA</th><th>BA files</th><th>Scope</th><th>Scope files</th><th>Other files</th><th></th></tr></thead><tbody>' +
      rows.map(function (r) {
        var baZ = r.baDone ? ZONES[2] : r.ba.length || (isNum(r.baPct) && r.baPct > 0) ? ZONES[1] : ZONES[0];
        return '<tr tabindex="0" data-docs-dept="' + r.i + '" title="View documents for ' + esc(r.name) + '">' +
          '<td><span class="sd-no">' + r.no + '</span>' + esc(r.name) + '</td>' +
          '<td>' + dots(r) + '</td>' +
          '<td><span class="x-s ' + baZ.cls + '"><i class="zdot"></i>' + (r.baDone ? 'Complete' : pct(r.baPct)) + '</span></td>' +
          '<td>' + count(r.ba.length) + '</td>' +
          '<td><span class="x-s ' + r.status.cls + '"><i class="zdot"></i>' + r.status.label + '</span></td>' +
          '<td>' + count(r.sc.length) + '</td>' +
          '<td>' + count(r.other.length) + '</td>' +
          '<td class="sd-go">View documents →</td></tr>';
      }).join('') + '</tbody></table></div></div>';
  }

  /* ================= Document map: ERP Documents → departments → BA / Scope / Other → files ================= */
  var DM_COL = [0, 250, 560, 790], DM_W = [210, 270, 190, 340], DM_ROW = 48, DM_H = 36;

  function docTree(rows) {
    var n = function (list) { return list.length; };
    var total = function (r) { return n(r.ba) + n(r.sc) + n(r.other); };
    var all = rows.reduce(function (a, r) { return a + total(r); }, 0);
    return { id: 'root', kind: 'root', label: 'ERP Documents', sub: all + ' files · ' + rows.length + ' departments',
      children: rows.map(function (r) {
        return { id: 'd' + r.i, kind: 'dept', label: r.name, cls: r.status.cls, sub: total(r) + ' file' + (total(r) === 1 ? '' : 's'),
          children: [['ba', 'BA documents', r.ba], ['scope', 'Scope documents', r.sc], ['other', 'Other documents', r.other]].map(function (g) {
            return { id: 'd' + r.i + g[0], kind: 'group', type: g[0], label: g[1], count: g[2].length,
              children: g[2].map(function (f, k) {
                return { id: 'd' + r.i + g[0] + k, kind: 'file', label: f.name, sub: f.updated ? 'Updated ' + fmtDay(f.updated) : '', dept: r.i, type: g[0], idx: k };
              }) };
          }) };
      }) };
  }

  function renderDocMap(host, rows) {
    var open = S.dmOpen || (S.dmOpen = { root: true });
    var view = S.dmView || (S.dmView = { x: 24, y: 24, k: 1 });
    var tree = docTree(rows), nodes = [], links = [], y = 0;

    // Layout: every visible leaf takes one row; a parent sits in the middle of its children.
    (function place(nd, depth) {
      nd.depth = depth; nd.x = DM_COL[depth];
      var kids = nd.children && nd.children.length && open[nd.id] ? nd.children : [];
      if (!kids.length) { nd.y = y; y += DM_ROW; }
      else {
        kids.forEach(function (k) { place(k, depth + 1); });
        nd.y = (kids[0].y + kids[kids.length - 1].y) / 2;
        kids.forEach(function (k) { links.push([nd, k]); });
      }
      nodes.push(nd);
    })(tree, 0);
    var W = DM_COL[3] + DM_W[3] + 40, H = Math.max(y, DM_ROW);

    var paths = links.map(function (l) {
      var a = l[0], b = l[1], x1 = a.x + DM_W[a.depth], y1 = a.y + DM_H / 2, x2 = b.x, y2 = b.y + DM_H / 2, mx = (x1 + x2) / 2;
      return '<path d="M' + x1 + ' ' + y1 + 'C' + mx + ' ' + y1 + ' ' + mx + ' ' + y2 + ' ' + x2 + ' ' + y2 + '"/>';
    }).join('');

    var html = nodes.map(function (nd) {
      var has = nd.children && nd.children.length, isOpen = !!open[nd.id];
      var pos = 'left:' + nd.x + 'px;top:' + nd.y + 'px;width:' + DM_W[nd.depth] + 'px;height:' + DM_H + 'px';
      var toggle = has ? '<span class="dm-tog" aria-hidden="true">' + (isOpen ? '‹' : '›') + '</span>' : '';
      if (nd.kind === 'file') {
        return '<button type="button" class="dm-node dm-file" style="' + pos + '" data-doc-dept="' + nd.dept + '" data-doc-type="' + nd.type + '" data-doc-idx="' + nd.idx + '" title="Preview ' + esc(nd.label) + '">' +
          '<span class="dm-ic" aria-hidden="true">📄</span><span class="dm-txt"><b>' + esc(nd.label) + '</b>' + (nd.sub ? '<small>' + nd.sub + '</small>' : '') + '</span></button>';
      }
      var extra = nd.kind === 'group' ? '<em class="dm-count">' + nd.count + '</em>' : '';
      var cls = 'dm-node dm-' + nd.kind + (nd.cls ? ' ' + nd.cls : '') + (nd.kind === 'group' && !nd.count ? ' empty' : '') + (has ? ' can-open' : '');
      return '<div class="' + cls + '" style="' + pos + '"' + (has ? ' data-dm="' + nd.id + '" role="button" tabindex="0" aria-expanded="' + isOpen + '"' : '') + '>' +
        (nd.kind === 'dept' ? '<i class="zdot"></i>' : '') + '<span class="dm-txt"><b>' + esc(nd.label) + '</b>' +
        (nd.sub ? '<small>' + esc(nd.sub) + '</small>' : nd.kind === 'group' && !nd.count ? '<small>Not uploaded</small>' : '') + '</span>' + extra + toggle + '</div>';
    }).join('');

    host.innerHTML = '<div class="card dm">' +
      '<div class="dm-bar"><span class="dm-hint">Click a box to open or close it · drag to move · scroll to zoom · click a 📄 file to preview</span>' +
      '<div class="dm-tools"><button type="button" class="btn" data-dm-act="expand">Expand all</button><button type="button" class="btn" data-dm-act="collapse">Collapse all</button>' +
      '<button type="button" class="btn dm-z" data-dm-act="out" aria-label="Zoom out">−</button><button type="button" class="btn dm-z" data-dm-act="in" aria-label="Zoom in">+</button>' +
      '<button type="button" class="btn" data-dm-act="reset">Reset view</button></div></div>' +
      '<div class="dm-view" id="dm-view"><div class="dm-canvas" id="dm-canvas" style="width:' + W + 'px;height:' + H + 'px">' +
      '<svg class="dm-links" width="' + W + '" height="' + H + '" aria-hidden="true">' + paths + '</svg>' + html + '</div></div></div>';

    var canvas = $('dm-canvas'), vp = $('dm-view');
    var apply = function () { canvas.style.transform = 'translate(' + view.x + 'px,' + view.y + 'px) scale(' + view.k + ')'; };
    var zoom = function (f, cx, cy) {
      var k = Math.max(0.4, Math.min(2, view.k * f));
      if (cx === undefined) { var b = vp.getBoundingClientRect(); cx = b.width / 2; cy = b.height / 2; }
      view.x = cx - (cx - view.x) * k / view.k; view.y = cy - (cy - view.y) * k / view.k; view.k = k; apply();
    };
    apply();

    // open / close a branch
    var toggleNode = function (id) { open[id] = !open[id]; renderDocMap(host, rows); };
    vp.addEventListener('click', function (e) {
      var nd = e.target.closest('[data-dm]');
      if (nd && !vp.dragged) toggleNode(nd.getAttribute('data-dm'));
    });
    vp.addEventListener('keydown', function (e) {
      var nd = e.target.closest('[data-dm]');
      if (nd && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); toggleNode(nd.getAttribute('data-dm')); }
    });
    // drag to move, wheel to zoom
    var drag = null;
    vp.addEventListener('pointerdown', function (e) {
      if (e.button !== 0) return;
      drag = { x: e.clientX, y: e.clientY, vx: view.x, vy: view.y }; vp.dragged = false;
    });
    window.addEventListener('pointermove', function (e) {
      if (!drag) return;
      var dx = e.clientX - drag.x, dy = e.clientY - drag.y;
      if (Math.abs(dx) + Math.abs(dy) > 4) { vp.dragged = true; vp.classList.add('grabbing'); }
      if (vp.dragged) { view.x = drag.vx + dx; view.y = drag.vy + dy; apply(); }
    });
    window.addEventListener('pointerup', function () { drag = null; vp.classList.remove('grabbing'); setTimeout(function () { vp.dragged = false; }, 0); });
    vp.addEventListener('click', function (e) { if (vp.dragged) { e.stopPropagation(); e.preventDefault(); } }, true);
    vp.addEventListener('wheel', function (e) {
      e.preventDefault();
      var b = vp.getBoundingClientRect();
      zoom(e.deltaY < 0 ? 1.1 : 1 / 1.1, e.clientX - b.left, e.clientY - b.top);
    }, { passive: false });
    // toolbar
    host.querySelector('.dm-tools').addEventListener('click', function (e) {
      var b = e.target.closest('[data-dm-act]');
      if (!b) return;
      var act = b.getAttribute('data-dm-act');
      if (act === 'in') zoom(1.2);
      else if (act === 'out') zoom(1 / 1.2);
      else if (act === 'reset') { S.dmView = { x: 24, y: 24, k: 1 }; renderDocMap(host, rows); }
      else if (act === 'collapse') { S.dmOpen = { root: true }; S.dmView = { x: 24, y: 24, k: 1 }; renderDocMap(host, rows); }
      else if (act === 'expand') {
        var o = { root: true };
        tree.children.forEach(function (d) { o[d.id] = true; d.children.forEach(function (g) { if (g.count) o[g.id] = true; }); });
        S.dmOpen = o; renderDocMap(host, rows);
      }
    });
  }

  /* Popup with every document of one department: BA, Scope and Other; click a file to preview it. */
  function openDeptDocs(i) {
    var r = scopeRows()[i];
    if (!r) return;
    var section = function (title, type, list) {
      return '<section class="sd-doc-col"><h4>' + title + ' <em class="sd-count">' + list.length + '</em></h4>' + (list.length ? list.map(function (d, k) {
        return '<button type="button" class="sd-file" data-doc-dept="' + i + '" data-doc-type="' + type + '" data-doc-idx="' + k + '" data-name="' + esc(norm(d.name)) + '">' +
          '<span class="sd-file-ic" aria-hidden="true">📄</span><span class="sd-file-txt"><b>' + esc(d.name) + '</b>' +
          (d.updated ? '<small>Updated ' + fmtDay(d.updated) + '</small>' : '') + '</span></button>';
      }).join('') : '<span class="sd-file none">Not uploaded</span>') + '</section>';
    };
    var all = r.ba.length + r.sc.length + r.other.length;
    var baZ = r.baDone ? ZONES[2] : r.ba.length || (isNum(r.baPct) && r.baPct > 0) ? ZONES[1] : ZONES[0];
    $('modal').innerHTML = '<div class="m-box card sd-docs" role="dialog" aria-modal="true" aria-label="Documents for ' + esc(r.name) + '">' +
      '<button type="button" class="m-close" aria-label="Close">✕</button>' +
      '<div class="m-head"><span class="sd-others-ic" aria-hidden="true">📁</span><div><div class="d-no">Department ' + r.no + ' · ' + all + ' document' + (all === 1 ? '' : 's') + '</div>' +
      '<h2>' + esc(r.name) + '</h2></div>' +
      '<span class="x-tag ' + baZ.cls + '"><i class="zdot"></i>BA ' + (r.baDone ? 'complete' : pct(r.baPct)) + '</span>' +
      '<span class="x-tag ' + r.status.cls + '"><i class="zdot"></i>' + r.status.label + '</span></div>' +
      (all > 8 ? '<input class="sd-others-q" type="search" placeholder="Find a document" aria-label="Find a document">' : '') +
      '<div class="sd-doc-cols">' + section('BA documents', 'ba', r.ba) + section('Scope documents', 'scope', r.sc) + section('Other documents', 'other', r.other) + '</div></div>';
    $('modal').hidden = false;
    document.body.style.overflow = 'hidden';
    var q = $('modal').querySelector('.sd-others-q');
    if (q) q.addEventListener('input', function () {
      var t = norm(q.value);
      Array.prototype.forEach.call($('modal').querySelectorAll('.sd-doc-col .sd-file[data-name]'), function (b) {
        b.hidden = t && b.getAttribute('data-name').indexOf(t) < 0;
      });
    });
    $('modal').querySelector('.m-close').focus();
  }

  /* Tabs 3 and 4 are placeholders until they are built. */
  function renderSoon(tab) {
    var info = tab === 'scope'
      ? ['📄', 'Scope & Documentation Tracker', 'Resolves foundational gaps regarding missing scope documents and formal baseline agreements to ensure rigorous project governance.']
      : ['🐞', 'Master Issue / Bug Log', 'A centralized, department-categorized repository for all operational bugs and technical discrepancies, streamlining communication and tracking with the external vendor.'];
    $('subtitle').textContent = info[1];
    $('app').innerHTML = '<section class="card soon"><div class="t-emoji" aria-hidden="true">' + icon(info[0], 30) + '</div><span class="x-status">Coming soon</span><h2>' + info[1] + '</h2><p>' + info[2] + '</p></section>';
  }

  var enterTimer;
  function render(animate) {
    var m = S.model;
    $('app').classList.toggle('enter', !!animate);
    clearTimeout(enterTimer);
    if (animate) enterTimer = setTimeout(function () { $('app').classList.remove('enter'); }, 1800);
    if (!m.depts.length || !m.weeks.length) {
      $('app').innerHTML = '<div class="card state"><b>No data found</b>The BA, AFI, Reports and Issues tabs need a “Department” header row with week blocks to its right.</div>';
      return;
    }
    Array.prototype.forEach.call($('tabs').children, function (b) { b.setAttribute('aria-pressed', String(b.getAttribute('data-tab') === S.tab)); });
    if (S.tab === 'summary') { if (S.sumView === 'tasks') renderTasks(); else renderSummary(animate); return; }
    if (S.tab === 'scope') { renderScope(); return; }
    if (S.tab !== 'tracker') { renderSoon(S.tab); return; }
    var wk = m.weeks[S.sel];
    $('subtitle').textContent = 'Status as of ' + wk.label + (wk.date ? ' · ' + niceDate(wk.date, true) : '') + ' · ' + m.depts.length + ' departments';
    $('app').innerHTML = renderWeeks() + renderKpis() + renderJourneyShell() + renderChartShells() + renderDeptShell() +
      (m.missing.length ? '<p class="note">Tab not found in the sheet: ' + esc(m.missing.join(', ')) + '. Its figures are shown as “–”.</p>' : '');
    renderDeptBody();
    drawCharts();
    drawJourney(animate);
    placePops();
    if (animate) countUp();
    $('journey').addEventListener('click', function (e) {
      var b = e.target.closest('.j-node');
      if (b) goToDept(Number(b.getAttribute('data-no')));
    });
    var openFromCard = function (e) {
      var c = e.target.closest('.dept[data-no]');
      if (!c) return;
      if (e.type === 'keydown') { if (e.key !== 'Enter' && e.key !== ' ') return; e.preventDefault(); }
      openDept(Number(c.getAttribute('data-no')));
    };
    $('dept-body').addEventListener('click', openFromCard);
    $('dept-body').addEventListener('keydown', openFromCard);

    $('seg').addEventListener('click', function (e) {
      var b = e.target.closest('button[data-w]');
      if (!b || b.disabled) return;
      S.sel = Number(b.getAttribute('data-w')); S.picked = S.sel !== m.latest;
      render(true);
    });
    $('q').addEventListener('input', function (e) { S.q = e.target.value; renderDeptBody(); });
    $('sort').addEventListener('change', function (e) { S.sort = e.target.value; renderDeptBody(); });
    $('view').addEventListener('click', function (e) {
      var b = e.target.closest('button[data-view]');
      if (!b) return;
      S.view = b.getAttribute('data-view');
      Array.prototype.forEach.call($('view').children, function (c) { c.setAttribute('aria-pressed', String(c === b)); });
      renderDeptBody();
    });
    var pressed = $('seg').querySelector('[aria-pressed="true"]');
    if (pressed) $('seg').scrollLeft = Math.max(0, pressed.offsetLeft - 60);
  }

  function onData(data) {
    $('refresh').disabled = false;
    if (data.brand && !S.brand) {
      S.brand = data.brand;
      if (data.brand.name) $('company').textContent = data.brand.name;      // from COMPANY_NAME in Code.gs
      var src = data.brand.logo || data.brand.logoUrl, why = data.brand.logoError;
      var logoNote = function () {
        $('logo').hidden = true;
        $('logo-note').textContent = 'Logo not loaded' + (why ? ': ' + why : '') + '. In the Apps Script editor run "authorizeLogo" once, then deploy a new version.';
        $('logo-note').hidden = false;
      };
      if (src) {
        $('logo').onerror = logoNote;
        $('logo').alt = (data.brand.name || 'Company') + ' logo';
        $('logo').src = src;
        $('logo').hidden = false;
      } else if (why) logoNote();
    }
    S.model = build(data);
    S.tasks = taskModel(data);
    if (!S.picked || S.sel >= S.model.weeks.length) S.sel = S.model.latest;   // follow the newest week unless the viewer chose one
    $('updated').textContent = 'Updated ' + new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    var sig = JSON.stringify(data.tabs);     // animate only on first load or when the sheet actually changed
    render(sig !== S.sig);
    S.sig = sig;
  }

  /* ---------- theme ---------- */
  function setTheme(t, save) {
    document.documentElement.setAttribute('data-theme', t);
    $('theme').textContent = t === 'dark' ? '◐ Light' : '◐ Dark';
    if (save) { try { localStorage.setItem('erp-theme', t); } catch (e) {} }
  }
  (function () {
    var saved = null;
    try { saved = localStorage.getItem('erp-theme'); } catch (e) {}
    setTheme(saved === 'dark' || saved === 'light' ? saved : 'dark', false);   // dark unless the viewer chose light
  })();
  $('theme').addEventListener('click', function () {
    setTheme(document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark', true);
  });
  function onError(err) {
    $('refresh').disabled = false;
    var msg = (err && err.message) || 'Unknown error';
    if (S.model) { $('updated').textContent = 'Refresh failed: ' + msg; return; }
    $('app').innerHTML = '<div class="card state"><b>Could not read the sheet</b>' + esc(msg) + '</div>';
  }
  function load() {
    $('refresh').disabled = true;
    if (window.__MOCK__) { onData(window.__MOCK__); return; }           // offline sample data (optional)
    if (!/^https:\/\/script\.google(usercontent)?\.com\//.test(DATA_URL)) {
      onError({ message: 'Set DATA_URL at the top of the script in index.html to your Apps Script /exec link.' });
      return;
    }
    // Reads the sheet figures from the Apps Script data feed; asks for the name + logo only once.
    fetch(DATA_URL + (S.brand ? '' : '?brand=1'), { cache: 'no-store' })
      .then(function (r) { if (!r.ok) throw new Error('Data feed returned HTTP ' + r.status); return r.json(); })
      .then(function (d) { if (d && d.error) throw new Error(d.error); onData(d); })
      .catch(function (e) {
        onError({ message: (e && e.message) || 'Could not reach the data feed' });
      });
  }

  /* Loading screen: step through what is being read until the data arrives. */
  (function () {
    var steps = ['Reading the BA tab…', 'Reading the Scope tab…', 'Reading the AFI tab…', 'Reading the Reports tab…', 'Reading the Issues tab…', 'Building the dashboard…'], i = 0;
    var timer = setInterval(function () {
      var el = $('load-text');
      if (!el) { clearInterval(timer); return; }
      i = Math.min(i + 1, steps.length - 1);
      el.textContent = steps[i];
    }, 900);
  })();
  /* 📄 buttons: open the document instead of the row / card they sit in */
  document.addEventListener('click', function (e) {
    var b = e.target.closest && e.target.closest('[data-doc-type]');
    if (!b) return;
    e.preventDefault(); e.stopPropagation();
    openDoc(Number(b.getAttribute('data-doc-dept')), b.getAttribute('data-doc-type'), Number(b.getAttribute('data-doc-idx') || 0));
  }, true);
  document.addEventListener('click', function (e) {
    var b = e.target.closest && e.target.closest('[data-other-dept]');
    if (!b || b.disabled) return;
    e.preventDefault(); e.stopPropagation();
    openOthers(Number(b.getAttribute('data-other-dept')));
  }, true);
  document.addEventListener('click', function (e) {
    var b = e.target.closest && e.target.closest('[data-docs-dept]');
    if (!b) return;
    e.preventDefault(); e.stopPropagation();
    openDeptDocs(Number(b.getAttribute('data-docs-dept')));
  }, true);
  document.addEventListener('keydown', function (e) {
    var b = e.target.closest && e.target.closest('tr[data-docs-dept]');
    if (b && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); openDeptDocs(Number(b.getAttribute('data-docs-dept'))); }
  });
  document.addEventListener('keydown', function (e) {
    if ((e.key === 'Enter' || e.key === ' ') && e.target.closest && e.target.closest('[data-doc-type]')) e.stopPropagation();
  }, true);

  $('modal').addEventListener('click', function (e) { if (e.target === $('modal') || e.target.closest('.m-close')) closeDept(); });
  document.addEventListener('keydown', function (e) { if (e.key === 'Escape') closeDept(); });
  $('refresh').addEventListener('click', load);
  var resizeTimer;
  window.addEventListener('resize', function () { clearTimeout(resizeTimer); resizeTimer = setTimeout(function () { if (S.model) { drawCharts(); drawJourney(false); placePops(); if (S.tab === 'summary' && S.sumView !== 'tasks') drawOverall(); } }, 150); });
  $('tabs').addEventListener('click', function (e) {
    var b = e.target.closest('button[data-tab]');
    if (!b || !S.model || b.getAttribute('data-tab') === S.tab) return;
    S.tab = b.getAttribute('data-tab');
    render(true);
    window.scrollTo(0, 0);
  });
  setInterval(load, REFRESH_MS);
  load();
})();
