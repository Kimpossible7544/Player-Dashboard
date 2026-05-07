// wpx_data.js — fetches and parses the WPX Excel file from Dropbox
// Depends on SheetJS (xlsx.full.min.js) already loaded on the page.

async function loadWPXData() {
  // ── CONFIGURATION ──────────────────────────────────────────────────────────
  // Your Dropbox shared link. IMPORTANT: must end with dl=1 for direct download.
  // If you replace the file or move it, update this URL.
  const DROPBOX_URL =
    "https://www.dropbox.com/scl/fi/7hzcatrogbtfw6ypn4xmx/WPXFinal5.3.26.xlsm" +
    "?rlkey=gal6fll38mz4kxrmlntud2ovn&dl=1";

  const PLAYER_TRACKING_SHEET = "Player Tracking";
  // ───────────────────────────────────────────────────────────────────────────

  // Verify SheetJS is available
  if (typeof XLSX === "undefined") {
    throw new Error("SheetJS (XLSX) is not loaded. Make sure xlsx.full.min.js is included before wpx_data.js.");
  }

  // 1. Download the file as an ArrayBuffer
  let response;
  try {
    response = await fetch(DROPBOX_URL);
  } catch (err) {
    console.error("[WPX] Fetch failed:", err);
    throw new Error(
      "Could not reach Dropbox. Check your internet connection or that the file is publicly shared. " +
      "Details: " + err.message
    );
  }

  if (!response.ok) {
    throw new Error(
      `Dropbox returned an error: ${response.status} ${response.statusText}. ` +
      "Make sure the file is shared as 'Anyone with the link' in Dropbox."
    );
  }

  const contentType = response.headers.get("content-type") || "";
  if (contentType.includes("text/html")) {
    throw new Error(
      "Dropbox returned an HTML page instead of the file. " +
      "The share link may have expired or the file may require sign-in. " +
      "Try generating a new share link in Dropbox and updating DROPBOX_URL in wpx_data.js."
    );
  }

  let buffer;
  try {
    buffer = await response.arrayBuffer();
  } catch (err) {
    throw new Error("Failed to read file data from Dropbox: " + err.message);
  }

  if (buffer.byteLength < 1000) {
    throw new Error(
      `Downloaded file is only ${buffer.byteLength} bytes — this is not a valid Excel file. ` +
      "The Dropbox link may be broken or the file may have moved."
    );
  }

  // 2. Parse with SheetJS
  let workbook;
  try {
    workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  } catch (err) {
    throw new Error("SheetJS could not parse the file: " + err.message);
  }

  // 3. Find the Player Tracking sheet
  if (!workbook.SheetNames.includes(PLAYER_TRACKING_SHEET)) {
    throw new Error(
      `Sheet "${PLAYER_TRACKING_SHEET}" not found in workbook. ` +
      "Available sheets: " + workbook.SheetNames.join(", ")
    );
  }
  const ws = workbook.Sheets[PLAYER_TRACKING_SHEET];

  // 4. Convert to array-of-arrays
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  if (rows.length < 2) {
    throw new Error("Player Tracking sheet appears to be empty.");
  }

  const headers = rows[0];

  // 5. Parse week labels and column indices from headers.
  //    Columns come in pairs: "Weekly Total (Week Label)", "Weekly Rank (Week Label)"
  //    starting at column index 6.
  const weekLabels = [];
  const weekTotalCols = [];

  for (let c = 6; c < headers.length; c += 2) {
    const h = headers[c];
    if (!h || typeof h !== "string") break;
    const match = h.match(/Weekly Total \((.+)\)/);
    if (!match) break;
    weekLabels.push(match[1]);
    weekTotalCols.push(c);
  }

  if (weekLabels.length === 0) {
    throw new Error("No weekly score columns found in Player Tracking sheet. Check the sheet format.");
  }

  // 6. Build the players object
  //    Column layout:
  //      0 = Name            3 = Overall Weekly Average
  //      1 = Overall Total   4 = Overall Missed Daily Goals
  //      2 = Overall Rank    5 = Overall Missed Weekly Goals
  //      6+ = pairs of (Weekly Total, Weekly Rank) per week (most recent first)

  const players = {};

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const name = row[0];
    if (!name) continue; // skip blank rows

    const weeks = weekTotalCols.map((col) => ({
      score: row[col] || 0,
    }));

    players[name] = {
      name:         name,
      totalScore:   row[1] || 0,
      overallRank:  row[2] || null,
      weeklyAvg:    row[3] || 0,
      missedDaily:  row[4] || 0,
      missedWeekly: row[5] || 0,
      weeks,
    };
  }

  const playerCount = Object.keys(players).length;
  if (playerCount === 0) {
    throw new Error("No player data found in Player Tracking sheet.");
  }

  console.log(`[WPX] Loaded ${playerCount} players, ${weekLabels.length} weeks:`, weekLabels);
  return { weekLabels, players };
}
