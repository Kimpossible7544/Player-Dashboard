// wpx_data.js — fetches and parses the WPX Excel file from Dropbox
// Depends on SheetJS (xlsx.full.min.js) already loaded on the page.

async function loadWPXData() {
  // ── CONFIGURATION ──────────────────────────────────────────────────────────
  // To update: replace this URL with your new Dropbox share link.
  // Make sure it ends with dl=1 (not dl=0) for direct download.
  const DROPBOX_URL =
    "https://www.dropbox.com/scl/fi/5ncxcj5cvsvuy4vbuoe4f/WPXFinal5.3.26.xlsm" +
    "?rlkey=0m3xb1gu6tpkvccdgz0ijsx3k&dl=1";

  const PLAYER_TRACKING_SHEET = "Player Tracking";
  // ───────────────────────────────────────────────────────────────────────────

  if (typeof XLSX === "undefined") {
    throw new Error("SheetJS (XLSX) is not loaded. Make sure xlsx.full.min.js is included before wpx_data.js.");
  }

  let response;
  try {
    response = await fetch(DROPBOX_URL);
  } catch (err) {
    console.error("[WPX] Fetch failed:", err);
    throw new Error("Could not reach Dropbox. Details: " + err.message);
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
      "The share link may have expired — try generating a new one and updating DROPBOX_URL in wpx_data.js."
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
      `Downloaded file is only ${buffer.byteLength} bytes — not a valid Excel file. ` +
      "The Dropbox link may be broken."
    );
  }

  let workbook;
  try {
    workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  } catch (err) {
    throw new Error("SheetJS could not parse the file: " + err.message);
  }

  if (!workbook.SheetNames.includes(PLAYER_TRACKING_SHEET)) {
    throw new Error(
      `Sheet "${PLAYER_TRACKING_SHEET}" not found. ` +
      "Available sheets: " + workbook.SheetNames.join(", ")
    );
  }

  const ws = workbook.Sheets[PLAYER_TRACKING_SHEET];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  if (rows.length < 2) throw new Error("Player Tracking sheet appears to be empty.");

  const headers = rows[0];

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
    throw new Error("No weekly score columns found in Player Tracking sheet.");
  }

  const players = {};
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const name = row[0];
    if (!name) continue;

    players[name] = {
      name,
      totalScore:   row[1] || 0,
      overallRank:  row[2] || null,
      weeklyAvg:    row[3] || 0,
      missedDaily:  row[4] || 0,
      missedWeekly: row[5] || 0,
      weeks: weekTotalCols.map((col) => ({ score: row[col] || 0 })),
    };
  }

  const playerCount = Object.keys(players).length;
  if (playerCount === 0) throw new Error("No player data found in Player Tracking sheet.");

  console.log(`[WPX] Loaded ${playerCount} players, ${weekLabels.length} weeks:`, weekLabels);
  return { weekLabels, players };
}
