/* WPX Data Loader — shared across standings / dashboard / projections
   Fetches the XLSM from Dropbox via Cloudflare Worker proxy, parses with SheetJS.
   ────────────────────────────────────────────────────────────────────
   To update data: overwrite the SAME file in Dropbox (keep filename identical
   so the share-link stays the same). */

const WPX_CONFIG = {
  dropboxUrl: 'https://www.dropbox.com/scl/fi/7hzcatrogbtfw6ypn4xmx/WPXFinal4.30.26.xlsm?rlkey=gal6fll38mz4kxrmlntud2ovn&dl=1',
  workerUrl: 'https://yellow-flower-ca92.kaheins32.workers.dev',
  secret: 'Abigail2011!',
  sheet: 'Player Tracking'
};

async function loadWPXData() {
  const proxyUrl = WPX_CONFIG.workerUrl + '?action=proxy&url=' + encodeURIComponent(WPX_CONFIG.dropboxUrl);
  const resp = await fetch(proxyUrl, {
    headers: { 'X-WPX-Secret': WPX_CONFIG.secret }
  });
  if (!resp.ok) throw new Error('Fetch failed: ' + resp.status);
  const buf = await resp.arrayBuffer();
  const wb = XLSX.read(new Uint8Array(buf));
  const ws = wb.Sheets[WPX_CONFIG.sheet];
  if (!ws) throw new Error('Sheet "' + WPX_CONFIG.sheet + '" not found');
  var result = parsePlayerTracking(XLSX.utils.sheet_to_json(ws, { header: 1 }));
  result.dailyData = parseWeeklySheets(wb);
  return result;
}

var MONTH_MAP = {January:0,February:1,March:2,April:3,May:4,June:5,July:6,August:7,September:8,October:9,November:10,December:11};

function parseWeekSheetDate(name) {
  var m = name.match(/^(\w+)\s+(\d+)\s*-/);
  if (!m) return null;
  var mon = MONTH_MAP[m[1]];
  if (mon === undefined) return null;
  return new Date(2026, mon, parseInt(m[2]));
}

function parseWeeklySheets(wb) {
  var DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
  var sheets = [];
  wb.SheetNames.forEach(function (name) {
    var d = parseWeekSheetDate(name);
    if (d) sheets.push({ name: name, date: d });
  });
  sheets.sort(function (a, b) { return a.date - b.date; });

  var weekOrder = sheets.map(function (s) { return s.name; });
  var playerDaily = {};

  sheets.forEach(function (s, wi) {
    var ws = wb.Sheets[s.name];
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
    for (var r = 1; r < rows.length; r++) {
      var d = rows[r];
      var name = d[0];
      if (!name) continue;
      name = String(name);
      if (!playerDaily[name]) {
        playerDaily[name] = {};
        DAYS.forEach(function (day) { playerDaily[name][day] = []; });
      }
      DAYS.forEach(function (day, di) {
        var val = d[di + 2];
        playerDaily[name][day][wi] = (typeof val === 'number') ? val : 0;
      });
    }
  });

  // Fill gaps with 0 for players who weren't in every week
  var totalWeeks = sheets.length;
  Object.keys(playerDaily).forEach(function (name) {
    DAYS.forEach(function (day) {
      for (var i = 0; i < totalWeeks; i++) {
        if (playerDaily[name][day][i] === undefined) playerDaily[name][day][i] = 0;
      }
    });
  });

  return { weekOrder: weekOrder, players: playerDaily };
}

function parsePlayerTracking(rows) {
  const hdr = rows[0];
  // Discover week column pairs: "Weekly Total (...)" + "Weekly Rank (...)"
  const weekCols = [];
  for (let i = 6; i < hdr.length; i += 2) {
    if (hdr[i] && String(hdr[i]).startsWith('Weekly Total')) {
      const label = String(hdr[i]).replace('Weekly Total (', '').replace(')', '');
      weekCols.push({ label: label, tIdx: i, rIdx: i + 1 });
    } else break;
  }

  const players = {};
  const playerList = [];
  for (let r = 1; r < rows.length; r++) {
    const d = rows[r];
    const name = d[0];
    if (!name) continue;
    const weeks = weekCols.map(function (wc) {
      return { label: wc.label, score: d[wc.tIdx] || 0, rank: d[wc.rIdx] || 0 };
    });
    const p = {
      name: String(name),
      totalScore: d[1] || 0,
      overallRank: d[2] || 0,
      weeklyAvg: d[3] || 0,
      missedDaily: d[4] || 0,
      missedWeekly: d[5] || 0,
      weeks: weeks
    };
    players[p.name] = p;
    playerList.push(p.name);
  }
  return { players: players, playerList: playerList, weekLabels: weekCols.map(function (w) { return w.label; }) };
}
