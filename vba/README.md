# VBA Macros for WPXStatsFinal.xlsm

## Setup (one-time, applies to all macros)

1. Open `WPXStatsFinal.xlsm` in Excel
2. Press **Alt + F11** to open the VBA editor
3. Click **Insert → Module**
4. Copy the contents of each `.bas` file into the module (you can put them all in one module)
5. Close the VBA editor

---

## SyncArchivedHeaders

Keeps the "Archived Players" header row in sync with "Player Tracking".

When a new week is added to "Player Tracking", run this macro to automatically add the matching `Weekly Total (...)` and `Weekly Rank (...)` headers to "Archived Players".

**Run:** Alt + F8 → `SyncArchivedHeaders` → Run

### Auto-run on file open (optional)

To sync headers automatically every time the file opens:

1. In the VBA editor, double-click **ThisWorkbook** in the Project Explorer
2. Add:
   ```vb
   Private Sub Workbook_Open()
       SyncArchivedHeaders
       LookupPlayerIDs  ' optional: also fill IDs on the active sheet
   End Sub
   ```
3. Save and close

---

## LookupPlayerIDs

Populates player IDs in **Column V** of the active sheet by looking up each name in Column A.

**Search order:**
1. **Roster** sheet — searches all 4 name columns (B, F, J, N) and AKA columns (D, H, L, P)
2. **Archived Players** sheet — searches Column A and returns the ID from Column CK

Rows that already have a value in Column V are skipped (won't overwrite existing IDs or formulas).

### How to use

1. Navigate to the weekly sheet you want to populate (e.g. "March 23 - March 28")
2. Press **Alt + F8**
3. Select `LookupPlayerIDs`
4. Click **Run**
5. A message box will show how many IDs were filled vs. skipped

### Important note

For the Archived Players fallback to work, you need player IDs in **Column CK** of the "Archived Players" sheet. If that column is empty, the macro will only find players who are still on the Roster.
