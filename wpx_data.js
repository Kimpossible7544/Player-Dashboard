// wpx_data.js — fetches and parses the WPX Excel file from Dropbox
// Serves wpx_standings.html, wpx_dashboard.html, and wpx_projections.html
// Depends on SheetJS (xlsx.full.min.js) already loaded on the page.

// Returns true if the week has started (start date has arrived).
function hasWeekStarted(weekLabel) {
  const parts = weekLabel.split(" - ");
  if (parts.length < 2) return true;

  const startPart = parts[0].trim();
  const now = new Date();
  const year = now.getFullYear();

  let startDate = new Date(`${startPart}, ${year}`);
  if (isNaN(startDate.getTime())) return true;

  if (startDate - now > 180 * 24 * 60 * 60 * 1000) {
    startDate = new Date(`${startPart}, ${year - 1}`);
  }

  return now >= startDate;
}

// Returns true if the week's Friday end date has fully passed.
function isWeekComplete(weekLabel) {
  const parts = weekLabel.split(" - ");
  if (parts.length < 2) return true;

  const endPart = parts[1].trim();
  const now = new Date();
  const year = now.getFullYear();

  let endDate = new Date(`${endPart}, ${year}`);
  if (isNaN(endDate.getTime())) return true;

  if (endDate - now > 180 * 24 * 60 * 60 * 1000) {
    endDate = new Date(`${endPart}, ${year - 1}`);
  }

  const completionDate = new Date(endDate);
  completionDate.setDate(completionDate.getDate() + 1);

  return now >= completionDate;
}

async function loadWPXData() {

  // =========================================================
  // DROPBOX FILE LOCATION
  // =========================================================
  // File is hosted on Dropbox. Keep the same filename (WPXStatsFinal.xlsm)
  // and overwrite it each week — the share link stays stable.
  // Use dl.dropboxusercontent.com to avoid CORS errors.
  // =========================================================

  const DROPBOX_URL =
    "https://dl.dropboxusercontent.com/scl/fi/dx7xgqmjshf8hciso3uya/WPXStatsFinal.xlsm" +
    "?rlkey=oyw14lm3fod48uxygykdnixar";

  const PLAYER_TRACKING_SHEET = "Player Tracking";
  const WEEK_SETTINGS_SHEET = "Week Settings";

  const DAYS = [
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday"
  ];

  const DAILY_GOAL = 6000000;
  const WEEKLY_GOAL = 20000000;

  // =========================================================
  // VERIFY SHEETJS EXISTS
  // =========================================================

  if (typeof XLSX === "undefined") {
    throw new Error(
      "SheetJS (XLSX) is not loaded. " +
      "Make sure xlsx.full.min.js loads BEFORE wpx_data.js"
    );
  }

  // =========================================================
  // FETCH EXCEL FILE
  // =========================================================

  let response;

  try {
    response = await fetch(DROPBOX_URL);
  } catch (err) {
    console.error("[WPX] Fetch failed:", err);

    throw new Error(
      "Could not reach Dropbox. Details: " + err.message
    );
  }

  if (!response.ok) {
    throw new Error(
      `Dropbox returned ${response.status} ${response.statusText}. ` +
      "Make sure the file is shared as 'Anyone with the link'."
    );
  }

  const buffer = await response.arrayBuffer();

  if (!buffer || buffer.byteLength < 1000) {
    throw new Error(
      "Downloaded file is invalid or too small."
    );
  }

  // =========================================================
  // PARSE EXCEL FILE
  // =========================================================

  let workbook;

  try {
    workbook = XLSX.read(buffer, {
      type: "array",
      cellDates: true
    });
  } catch (err) {
    console.error("[WPX] XLSX parse failed:", err);

    throw new Error(
      "SheetJS could not parse the Excel file: " + err.message
    );
  }

  // =========================================================
  // VERIFY PLAYER TRACKING SHEET EXISTS
  // =========================================================

  if (!workbook.SheetNames.includes(PLAYER_TRACKING_SHEET)) {
    throw new Error(
      `Sheet "${PLAYER_TRACKING_SHEET}" not found.\n` +
      `Available sheets: ${workbook.SheetNames.join(", ")}`
    );
  }

  // =========================================================
  // READ PLAYER TRACKING SHEET
  // =========================================================

  const ptRows = XLSX.utils.sheet_to_json(
    workbook.Sheets[PLAYER_TRACKING_SHEET],
    {
      header: 1,
      defval: null
    }
  );

  if (ptRows.length < 2) {
    throw new Error("Player Tracking sheet is empty.");
  }

  const headers = ptRows[0];

  // =========================================================
  // PARSE WEEK COLUMNS
  // =========================================================

  const weekLabels = [];
  const weekTotalCols = [];
  const weekRankCols = [];

  for (let c = 6; c < headers.length; c += 2) {

    const h = headers[c];

    if (!h || typeof h !== "string") break;

    const match = h.match(/Weekly Total \((.+)\)/);

    if (!match) break;

    const label = match[1];

    if (!hasWeekStarted(label)) {
      console.log(`[WPX] Skipping future week: ${label}`);
      continue;
    }

    weekLabels.push(label);
    weekTotalCols.push(c);
    weekRankCols.push(c + 1);
  }

  if (weekLabels.length === 0) {
    throw new Error("No valid weeks found.");
  }

  // Determine current in-progress week

  const currentWeekLabel =
    weekLabels.find(label => !isWeekComplete(label)) || null;

  // Reverse so oldest week appears first

  weekLabels.reverse();
  weekTotalCols.reverse();
  weekRankCols.reverse();

  // =========================================================
  // BUILD PLAYER OBJECTS
  // =========================================================

  const players = {};

  for (let r = 1; r < ptRows.length; r++) {

    const row = ptRows[r];

    const name = row[0];

    if (!name) continue;

    players[name] = {
      name,

      totalScore: row[1] || 0,

      overallRank: row[2] || null,

      weeklyAvg: row[3] || 0,

      missedDaily: row[4] || 0,

      missedWeekly: row[5] || 0,

      weeks: weekTotalCols.map((col, i) => ({
        label: weekLabels[i],

        score: row[col] || 0,

        rank: row[weekRankCols[i]] || null,

        inProgress:
          weekLabels[i] === currentWeekLabel
      }))
    };
  }

  if (Object.keys(players).length === 0) {
    throw new Error("No player data found.");
  }

  // =========================================================
  // READ WEEK SETTINGS SHEET
  // =========================================================

  const notPushingWeeks = new Set();

  if (workbook.SheetNames.includes(WEEK_SETTINGS_SHEET)) {

    const wsRows = XLSX.utils.sheet_to_json(
      workbook.Sheets[WEEK_SETTINGS_SHEET],
      {
        header: 1,
        defval: null
      }
    );

    for (let r = 1; r < wsRows.length; r++) {

      const weekLabel = wsRows[r][0];

      const pushing =
        String(wsRows[r][1] || "")
          .trim()
          .toUpperCase();

      if (weekLabel && pushing === "N") {
        notPushingWeeks.add(String(weekLabel).trim());
      }
    }

    console.log(
      "[WPX] Not-pushing weeks:",
      Array.from(notPushingWeeks)
    );
  }

  // Add pushing flags

  Object.values(players).forEach(player => {

    player.weeks.forEach(week => {

      week.pushing =
        !notPushingWeeks.has(week.label);

    });
  });

  // =========================================================
  // BUILD DAILY DATA
  // =========================================================

  const dailyData = {
    weekOrder: weekLabels,
    players: {}
  };

  Object.keys(players).forEach(name => {

    dailyData.players[name] = {
      Monday: [],
      Tuesday: [],
      Wednesday: [],
      Thursday: [],
      Friday: []
    };
  });

  for (const weekLabel of weekLabels) {

    if (!workbook.SheetNames.includes(weekLabel)) {

      Object.keys(players).forEach(name => {

        DAYS.forEach(day => {
          dailyData.players[name][day].push(0);
        });

      });

      continue;
    }

    const weekRows = XLSX.utils.sheet_to_json(
      workbook.Sheets[weekLabel],
      {
        header: 1,
        defval: null
      }
    );

    const byPlayer = {};

    for (let r = 1; r < weekRows.length; r++) {

      const row = weekRows[r];

      const name = row[0];

      if (!name) continue;

      byPlayer[name] = {
        Monday: row[2] || 0,
        Tuesday: row[3] || 0,
        Wednesday: row[4] || 0,
        Thursday: row[5] || 0,
        Friday: row[6] || 0
      };
    }

    Object.keys(players).forEach(name => {

      const pd = byPlayer[name];

      DAYS.forEach(day => {

        dailyData.players[name][day].push(
          pd ? (pd[day] || 0) : 0
        );

      });
    });
  }

  // =========================================================
  // RECALCULATE MISSED GOALS (excluding not-pushing weeks)
  // =========================================================

  Object.values(players).forEach(player => {

    let missedDaily = 0;
    let missedWeekly = 0;

    player.weeks.forEach((week, i) => {

      // Skip not-pushing weeks and in-progress weeks
      if (!week.pushing || week.inProgress) return;

      // Skip weeks where the player wasn't active
      if (!week.score || week.score <= 0) return;

      // Missed weekly: active pushing week with score below weekly goal
      if (week.score < WEEKLY_GOAL) {
        missedWeekly++;
      }

      // Missed daily: count days below daily goal in this week
      const weekLabel = week.label;
      const pd = dailyData.players[player.name];

      if (pd) {
        const weekIdx = dailyData.weekOrder.indexOf(weekLabel);

        if (weekIdx >= 0) {
          DAYS.forEach(day => {
            const dayScore = pd[day][weekIdx];

            if (dayScore > 0 && dayScore < DAILY_GOAL) {
              missedDaily++;
            }
          });
        }
      }
    });

    player.missedDaily = missedDaily;
    player.missedWeekly = missedWeekly;
  });

  // =========================================================
  // FINAL LOGGING
  // =========================================================

  const playerList = Object.keys(players);

  console.log(
    `[WPX] Loaded ${playerList.length} players ` +
    `across ${weekLabels.length} weeks`,
    weekLabels
  );

  // =========================================================
  // RETURN FINAL DATA
  // =========================================================

  return {
    weekLabels,
    playerList,
    players,
    dailyData,
    currentWeekLabel,
    notPushingWeeks,
    DAILY_GOAL,
    WEEKLY_GOAL
  };
}
