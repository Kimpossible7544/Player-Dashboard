// wpx_data.js — fetches and parses the WPX Excel file from Dropbox
// Depends on SheetJS (xlsx.full.min.js) already loaded on the page.

async function loadWPXData() {
  // ── CONFIGURATION ──────────────────────────────────────────────────────────
  // To update the data source, change this URL to your new Dropbox share link.
  // Make sure the link ends with   dl=1   (not dl=0) for a direct download.
  const DROPBOX_URL =
    "https://www.dropbox.com/scl/fi/7hzcatrogbtfw6ypn4xmx/WPXFinal5.3.26.xlsm" +
    "?rlkey=gal6fll38mz4kxrmlntud2ovn&st=bvozyveo&dl=1";

  const PLAYER_TRACKING_SHEET = "Player Tracking";
  // ───────────────────────────────────────────────────────────────────────────

  // 1. Download the file as an ArrayBuffer
  const response = await fetch(DROPBOX_URL);
  if (!response.ok) {
    throw new Error(`Failed to fetch file: ${response.status} ${response.statusText}`);
  }
  const buffer = await response.arrayBuffer();

  // 2. Parse with SheetJS
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });

  // 3. Find the Player Tracking sheet
  if (!workbook.SheetNames.includes(PLAYER_TRACKING_SHEET)) {
    throw new Error(`Sheet "${PLAYER_TRACKING_SHEET}" not found. Available: ${workbook.SheetNames.join(", ")}`);
  }
  const ws = workbook.Sheets[PLAYER_TRACKING_SHEET];

  // 4. Convert to array-of-arrays (raw values, no header coercion)
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  if (rows.length < 2) throw new Error("Player Tracking sheet appears empty.");

  const headers = rows[0]; // row 0 = header row

  // 5. Parse week labels from headers.
  //    Headers are pairs:  "Weekly Total (Week Label)"  /  "Weekly Rank (Week Label)"
  //    They start at column index 6 and repeat every 2 columns.
  const weekLabels = [];
  const weekTotalCols = []; // column indices for weekly totals

  for (let c = 6; c < headers.length; c += 2) {
    const h = headers[c];
    if (!h || typeof h !== "string") break;
    const match = h.match(/Weekly Total \((.+)\)/);
    if (!match) break;
    weekLabels.push(match[1]);
    weekTotalCols.push(c);
  }

  if (weekLabels.length === 0) {
    throw new Error("Could not find any weekly score columns in Player Tracking sheet.");
  }

  // 6. Build the players object
  //    Column layout:
  //      0  = Name
  //      1  = Overall Total
  //      2  = Overall Rank
  //      3  = Overall Weekly Average
  //      4  = Overall Missed Daily Goals
  //      5  = Overall Missed Weekly Goals
  //      6+ = pairs of (Weekly Total, Weekly Rank) per week

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

  return { weekLabels, players };
}
