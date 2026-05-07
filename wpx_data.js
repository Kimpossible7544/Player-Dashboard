// wpx_data.js — fetches and parses the WPX Excel file from Dropbox
// Serves wpx_standings.html, wpx_dashboard.html, and wpx_projections.html
// Depends on SheetJS (xlsx.full.min.js) already loaded on the page.

// Returns true only if the week's Friday end date has fully passed.
function isWeekComplete(weekLabel) {
  // Labels are formatted like "May 4 - May 9" or "March 23 - March 28"
  const parts = weekLabel.split(" - ");
  if (parts.length < 2) return true; // can't parse — include it

  const endPart = parts[1].trim(); // e.g. "May 9"
  const now     = new Date();
  const year    = now.getFullYear();

  // Try current year first, fall back to previous year if the date is
  // more than ~6 months in the future (handles year boundaries).
  let endDate = new Date(`${endPart}, ${year}`);
  if (isNaN(endDate.getTime())) return true; // unparseable — include it

  if (endDate - now > 180 * 24 * 60 * 60 * 1000) {
    endDate = new Date(`${endPart}, ${year - 1}`);
  }

  // The week is complete once the day AFTER the end date has begun
  // (i.e. Saturday 00:00 or later).
  const completionDate = new Date(endDate);
  completionDate.setDate(completionDate.getDate() + 1);

  return now >= completionDate;
}

async function loadWPXData() {
  // ── CONFIGURATION ──────────────────────────────────────────────────────────
  // If you move or rename the file in Dropbox, generate a new share link and
  // replace the URL below. Use dl.dropboxusercontent.com to avoid CORS errors.
  const DROPBOX_URL =
    "https://dl.dropboxusercontent.com/scl/fi/5ncxcj5cvsvuy4vbuoe4f/WPXFinal5.3.26.xlsm" +
    "?rlkey=0m3xb1gu6tpkvccdgz0ijsx3k";

  const PLAYER_TRACKING_SHEET = "Player Tracking";
  const DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
  // ───────────────────────────────────────────────────────────────────────────

  if (typeof XLSX === "undefined") {
    throw new Error("SheetJS (XLSX) is not loaded. Make sure xlsx.full.min.js is included before wpx_data.js.");
  }

  // 1. Fetch the file
  let response;
  try {
    response = await fetch(DROPBOX_URL);
  } catch (err) {
    console.error("[WPX] Fetch failed:", err);
    throw new Error("Could not reach Dropbox. Details: " + err.message);
  }

  if (!response.ok) {
    throw new Error(
      `Dropbox returned ${response.status} ${response.statusText}. ` +
      "Make sure the file is shared as 'Anyone with the link'."
    );
  }

  const contentType = response.headers.get("content-type") || "";
  if (contentType.includes("text/html")) {
    throw new Error(
      "Dropbox returned an HTML page instead of the file. " +
      "The share link may have expired — generate a new one and update DROPBOX_URL in wpx_data.js."
    );
  }

  const buffer = await response.arrayBuffer();

  if (buffer.byteLength < 1000) {
    throw new Error(
      `File is only ${buffer.byteLength} bytes — not a valid Excel file. ` +
      "Check your Dropbox link."
    );
  }

  // 2. Parse with SheetJS
  let workbook;
  try {
    workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  } catch (err) {
    throw new Error("SheetJS could not parse the file: " + err.message);
  }

  // 3. Read Player Tracking sheet
  if (!workbook.SheetNames.includes(PLAYER_TRACKING_SHEET)) {
    throw new Error(
      `Sheet "${PLAYER_TRACKING_SHEET}" not found. ` +
      "Available sheets: " + workbook.SheetNames.join(", ")
    );
  }

  const ptRows = XLSX.utils.sheet_to_json(
    workbook.Sheets[PLAYER_TRACKING_SHEET],
    { header: 1, defval: null }
  );

  if (ptRows.length < 2) throw new Error("Player Tracking sheet is empty.");

  const headers = ptRows[0];

  // 4. Parse week labels and column indices — only keep completed weeks
  const weekLabels    = [];
  const weekTotalCols = [];
  const weekRankCols  = [];

  for (let c = 6; c < headers.length; c += 2) {
    const h = headers[c];
    if (!h || typeof h !== "string") break;
    const match = h.match(/Weekly Total \((.+)\)/);
    if (!match) break;

    const label = match[1];
    if (!isWeekComplete(label)) {
      console.log(`[WPX] Skipping incomplete week: ${label}`);
      continue; // hide weeks that haven't finished yet
    }

    weekLabels.push(label);
    weekTotalCols.push(c);
    weekRankCols.push(c + 1);
  }

  if (weekLabels.length === 0) {
    throw new Error("No completed weeks found yet.");
  }

  // 5. Build players object from Player Tracking
  //    Col 0=Name, 1=Overall Total, 2=Rank, 3=Weekly Avg, 4=Missed Daily, 5=Missed Weekly
  const players = {};

  for (let r = 1; r < ptRows.length; r++) {
    const row  = ptRows[r];
    const name = row[0];
    if (!name) continue;

    players[name] = {
      name,
      totalScore:   row[1] || 0,
      overallRank:  row[2] || null,
      weeklyAvg:    row[3] || 0,
      missedDaily:  row[4] || 0,
      missedWeekly: row[5] || 0,
      weeks: weekTotalCols.map((col, i) => ({
        label: weekLabels[i],
        score: row[col]             || 0,
        rank:  row[weekRankCols[i]] || null,
      })),
    };
  }

  if (Object.keys(players).length === 0) {
    throw new Error("No player data found in Player Tracking sheet.");
  }

  // 6. Build dailyData from individual weekly sheets (completed weeks only)
  const dailyData = {
    weekOrder: weekLabels,
    players:   {},
  };

  Object.keys(players).forEach(name => {
    dailyData.players[name] = {
      Monday: [], Tuesday: [], Wednesday: [], Thursday: [], Friday: []
    };
  });

  for (const weekLabel of weekLabels) {
    if (!workbook.SheetNames.includes(weekLabel)) {
      Object.keys(players).forEach(name => {
        DAYS.forEach(day => dailyData.players[name][day].push(0));
      });
      continue;
    }

    const weekRows = XLSX.utils.sheet_to_json(
      workbook.Sheets[weekLabel],
      { header: 1, defval: null }
    );

    // Cols: 0=Name, 2=Mon, 3=Tue, 4=Wed, 5=Thu, 6=Fri
    const byPlayer = {};
    for (let r = 1; r < weekRows.length; r++) {
      const row  = weekRows[r];
      const name = row[0];
      if (!name) continue;
      byPlayer[name] = {
        Monday:    row[2] || 0,
        Tuesday:   row[3] || 0,
        Wednesday: row[4] || 0,
        Thursday:  row[5] || 0,
        Friday:    row[6] || 0,
      };
    }

    Object.keys(players).forEach(name => {
      const pd = byPlayer[name];
      DAYS.forEach(day => {
        dailyData.players[name][day].push(pd ? (pd[day] || 0) : 0);
      });
    });
  }

  const playerList = Object.keys(players);

  console.log(
    `[WPX] Loaded ${playerList.length} players across ${weekLabels.length} completed weeks:`,
    weekLabels
  );

  return { weekLabels, playerList, players, dailyData };
}
