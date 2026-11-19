/* ERP Implementation Dashboard: data loading, charts, tables and forecast */
(function () {
  'use strict';

  /* ▼▼▼ Paste your Apps Script web app link (ends in /exec) between the quotes ▼▼▼ */
  var DATA_URL = 'https://script.google.com/macros/s/AKfycbz_9F9Z6H4kuz8z6UtmVD3FYXNnSCEWEfMzCJm0k7YnuB99synaewrH7geej-i-yBkk/exec';
  /* ▲▲▲ ------------------------------------------------------------------- ▲▲▲ */

  var REFRESH_MS = 5 * 60 * 1000;          // the page re-reads the sheet every 5 minutes
  var S = { tab: 'summary', model: null, sel: 0, picked: false, sort: 'no', q: '', view: 'cards' };

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

    var model = { depts: depts, weeks: weeks, M: M, missing: data.missing || [], generatedAt: data.generatedAt };
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
    var head = function (label, emoji) { return '<div class="t-top"><div class="t-label">' + label + '</div><span class="t-emoji" aria-hidden="true">' + emoji + '</span></div>'; };

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
    return '<div class="sec-head"><div><h2>Rollout journey</h2><p>Every department placed by how much the ERP is in use. Click a department for its details. ' + staleTag(w) + '</p></div>' +
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
        '<div class="x-lab">Departments by ERP usage</div>' +
        '<div class="x-stack">' + ZONES.slice().reverse().map(function (zz, k) { return counts[2 - k] ? '<i class="' + zz.cls + '" style="flex:' + counts[2 - k] + '"></i>' : ''; }).join('') + '</div>' +
        '<div class="x-legend">' + ZONES.slice().reverse().map(function (zz, k) { return '<span class="' + zz.cls + '"><i class="zdot"></i><b>' + counts[2 - k] + '</b> ' + zz.name + '</span>'; }).join('') + (totalZ < n ? '<span>' + (n - totalZ) + ' no data</span>' : '') + '</div>' +
        (old.length ? '<p class="x-note">' + old.map(function (p) { return p[0]; }).join(', ') + ' not yet updated for ' + esc(wk.label) + '; last figures used.</p>' : '') +
      '</div>' +
      '</section>';

    var fig = function (emoji, label, value, small, sub, delta, cls, tone, bar) {
      return '<div class="card x-fig tone-' + (tone || 'blue') + '"><div class="t-top"><div class="x-lab">' + label + '</div><span class="t-emoji" aria-hidden="true">' + emoji + '</span></div>' +
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

    $('app').innerHTML = '<div class="x-sum"><div class="x-top">' + banner + figs + '</div>' +
      '<section class="card x-card chart x-wide"><div class="x-chead"><div><div class="x-lab">Overall progress, week by week</div><p>Average of Availability, Functionality and In use across all departments.</p></div><div class="x-eta" id="eta"></div></div><div class="plot" id="c-overall"></div></section>' +
      moves +
      '<div class="x-main">' +
        '<section class="card x-card"><div class="x-lab">Department status</div><p>Click a row for the department\'s full detail.</p>' + matrix + '</section></div></div>' +
      (m.missing.length ? '<p class="note">Tab not found in the sheet: ' + esc(m.missing.join(', ')) + '. Its figures are shown as “–”.</p>' : '');
    drawOverall();
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

  /* Tabs 3 and 4 are placeholders until they are built. */
  function renderSoon(tab) {
    var info = tab === 'scope'
      ? ['📄', 'Scope & Documentation Tracker', 'Resolves foundational gaps regarding missing scope documents and formal baseline agreements to ensure rigorous project governance.']
      : ['🐞', 'Master Issue / Bug Log', 'A centralized, department-categorized repository for all operational bugs and technical discrepancies, streamlining communication and tracking with the external vendor.'];
    $('subtitle').textContent = info[1];
    $('app').innerHTML = '<section class="card soon"><div class="t-emoji" aria-hidden="true">' + info[0] + '</div><span class="x-status">Coming soon</span><h2>' + info[1] + '</h2><p>' + info[2] + '</p></section>';
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
    if (S.tab === 'summary') { renderSummary(animate); return; }
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
      if (b) openDept(Number(b.getAttribute('data-no')));
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
    var system = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
    setTheme(saved === 'dark' || saved === 'light' ? saved : system, false);
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

  $('modal').addEventListener('click', function (e) { if (e.target === $('modal') || e.target.closest('.m-close')) closeDept(); });
  document.addEventListener('keydown', function (e) { if (e.key === 'Escape') closeDept(); });
  $('refresh').addEventListener('click', load);
  var resizeTimer;
  window.addEventListener('resize', function () { clearTimeout(resizeTimer); resizeTimer = setTimeout(function () { if (S.model) { drawCharts(); drawJourney(false); placePops(); if (S.tab === 'summary') drawOverall(); } }, 150); });
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
