/* Shared live-data loader for the GLGC Sunday Attendance dashboards.
   Reads the Google Sheet via gviz JSONP. The sheet is a pre-aggregated matrix:
     col A = Choir, col B = Governor, col C... = one Sunday date per column.
   Rows are organized into Services (section header rows where col A === col B),
   with TOTAL rows after each section and a GRAND TOTAL near the top. */
window.GLGC = (function () {
  // === Paste the Google Sheet ID here once you've uploaded the .xlsx ===
  var SHEET_ID = '1rXQDsGNKxZFZBa7PA_aMmwrgRIwAiJiHcRo-pC2y05o';
  var SHEET_NAME = 'Attendance';

  function load(cb, onErr) {
    window.__glgcCb = function (resp) {
      try { cb(build(processGviz(resp))); }
      catch (e) { if (onErr) onErr(e); else console.error(e); }
    };
    // Skip the merged-title row (sheet row 1) by starting the range at row 2.
    // headers=1 tells gviz to treat the first row of the range (the date row)
    // as column labels — which is how processGviz() finds the weekly columns.
    var url = 'https://docs.google.com/spreadsheets/d/' + SHEET_ID +
              '/gviz/tq?tqx=out:json;responseHandler:__glgcCb' +
              '&sheet=' + encodeURIComponent(SHEET_NAME) +
              '&range=A2:AZ500&headers=1';
    var s = document.createElement('script');
    s.src = url;
    s.onerror = function () { if (onErr) onErr(new Error('Could not reach the sheet')); };
    document.body.appendChild(s);
  }

  // ---- parse the gviz table into per-governor weekly points ----
  function processGviz(resp) {
    var cols = resp.table.cols;
    // Each weekly column's header label is a date string like "08 February 2026"
    // or "5th April 2026". Walk all cols past the first two, keep those whose
    // label parses as a Sunday-bearing date.
    var weekCols = [];
    for (var i = 2; i < cols.length; i++) {
      var label = String(cols[i].label || '').trim();
      var d = parseDateLabel(label);
      if (!d) continue;
      weekCols.push({
        idx: i,
        date: d,
        dateLabel: fmtDate(d),
        week: isoWeek(d)
      });
    }
    // Sort weeks chronologically so the chart x-axis is left-to-right by date.
    weekCols.sort(function (a, b) { return a.date - b.date; });

    var governors = [];
    var currentService = '';
    resp.table.rows.forEach(function (r) {
      var c = r.c || [];
      var a = cellStr(c[0]);
      var b = cellStr(c[1]);
      if (!a && !b) return;                                       // blank row
      if (b === 'GRAND TOTAL' || b === 'TOTAL') return;           // skip totals
      if (a === 'Name of Choir' || b === 'Name of Governor') return;
      if (a && a === b) { currentService = a; return; }           // section header

      var points = weekCols.map(function (wc) {
        var v = c[wc.idx] && c[wc.idx].v;
        var n = toNumOrNull(v);
        return { week: wc.week, dateLabel: wc.dateLabel, date: wc.date, attendance: n };
      });
      governors.push({
        governor: b,
        choir: a,
        service: currentService,
        points: points
      });
    });

    return { governors: governors, weekCols: weekCols };
  }

  // ---- build the higher-level aggregates ----
  function build(parsed) {
    var weekCols = parsed.weekCols;
    var governors = parsed.governors.map(function (G) {
      var total = G.points.reduce(function (s, p) { return s + (p.attendance || 0); }, 0);
      return { governor: G.governor, choir: G.choir, service: G.service, points: G.points, total: total };
    });

    // Group by choir (sum governor attendance per week).
    var choirsMap = {};
    governors.forEach(function (G) {
      if (!choirsMap[G.choir]) {
        choirsMap[G.choir] = {
          choir: G.choir, service: G.service, governors: [],
          points: weekCols.map(function (wc) { return { week: wc.week, dateLabel: wc.dateLabel, date: wc.date, attendance: 0, reporters: 0 }; })
        };
      }
      choirsMap[G.choir].governors.push(G.governor);
      G.points.forEach(function (p, i) {
        if (p.attendance != null) {
          choirsMap[G.choir].points[i].attendance += p.attendance;
          choirsMap[G.choir].points[i].reporters += 1;
        }
      });
    });
    var choirs = Object.keys(choirsMap).map(function (k) {
      var C = choirsMap[k];
      C.total = C.points.reduce(function (s, p) { return s + p.attendance; }, 0);
      return C;
    }).sort(function (a, b) { return a.choir.localeCompare(b.choir); });

    // Group by service (sum choir attendance per week).
    var servicesMap = {};
    governors.forEach(function (G) {
      var key = G.service || '(none)';
      if (!servicesMap[key]) {
        servicesMap[key] = {
          service: key, choirs: {}, governors: [],
          points: weekCols.map(function (wc) { return { week: wc.week, dateLabel: wc.dateLabel, date: wc.date, attendance: 0, reporters: 0 }; })
        };
      }
      servicesMap[key].choirs[G.choir] = true;
      servicesMap[key].governors.push(G.governor);
      G.points.forEach(function (p, i) {
        if (p.attendance != null) {
          servicesMap[key].points[i].attendance += p.attendance;
          servicesMap[key].points[i].reporters += 1;
        }
      });
    });
    var services = Object.keys(servicesMap).map(function (k) {
      var S = servicesMap[k];
      S.choirs = Object.keys(S.choirs);
      S.total = S.points.reduce(function (s, p) { return s + p.attendance; }, 0);
      return S;
    });

    // Grand total per week (all services combined).
    var weeks = weekCols.map(function (wc, i) {
      var total = 0, reporters = 0;
      governors.forEach(function (G) {
        var p = G.points[i];
        if (p.attendance != null) { total += p.attendance; reporters += 1; }
      });
      return { week: wc.week, dateLabel: wc.dateLabel, date: wc.date, attendance: total, reporters: reporters };
    });

    return { governors: governors, choirs: choirs, services: services, weeks: weeks };
  }

  // ---- helpers ----
  function cellStr(c) { return (c && c.v != null) ? String(c.v).trim() : ''; }

  function toNumOrNull(v) {
    if (v == null) return null;
    if (typeof v === 'number') return v;
    var s = String(v).trim();
    if (s === '' || s === '-' || s === '—' || s === 'N/A') return null;
    var n = parseFloat(s.replace(/[^0-9.\-]/g, ''));
    return isNaN(n) ? null : n;
  }

  var MONTHS = { jan:0,january:0, feb:1,february:1, mar:2,march:2, apr:3,april:3,
                 may:4, jun:5,june:5, jul:6,july:6, aug:7,august:7,
                 sep:8,sept:8,september:8, oct:9,october:9, nov:10,november:10,
                 dec:11,december:11 };

  // Parses date strings like "08 February 2026", "5th April 2026", "12th April 2026".
  function parseDateLabel(s) {
    if (!s) return null;
    var cleaned = s.replace(/(\d+)(st|nd|rd|th)/gi, '$1').trim();
    var m = cleaned.match(/(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})/);
    if (m) {
      var day = parseInt(m[1], 10);
      var mon = MONTHS[m[2].toLowerCase()];
      var yr = parseInt(m[3], 10);
      if (mon == null) return null;
      var d = new Date(yr, mon, day);
      return isNaN(d.getTime()) ? null : d;
    }
    var d2 = new Date(cleaned);
    return isNaN(d2.getTime()) ? null : d2;
  }

  function isoWeek(d) {
    var t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    var day = t.getUTCDay() || 7;
    t.setUTCDate(t.getUTCDate() + 4 - day);
    var ys = new Date(Date.UTC(t.getUTCFullYear(), 0, 1));
    return Math.ceil(((t - ys) / 86400000 + 1) / 7);
  }
  function fmtDate(d) {
    var dd = ('0' + d.getDate()).slice(-2),
        mm = ('0' + (d.getMonth() + 1)).slice(-2),
        yy = String(d.getFullYear()).slice(-2);
    return dd + '/' + mm + '/' + yy;
  }

  // Chart.js plugin: draw each value on top of its bar (always visible).
  var valueLabels = {
    id: 'valueLabels',
    afterDatasetsDraw: function (chart) {
      var ctx = chart.ctx;
      ctx.save();
      ctx.fillStyle = '#ffffff';
      ctx.font = 'bold 12px Arial';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'bottom';
      chart.data.datasets.forEach(function (ds, di) {
        var meta = chart.getDatasetMeta(di);
        if (meta.hidden) return;
        meta.data.forEach(function (bar, i) {
          var v = ds.data[i];
          if (v == null) return;
          ctx.fillText(v, bar.x, bar.y - 4);
        });
      });
      ctx.restore();
    }
  };

  return { load: load, valueLabels: valueLabels };
})();
