# VBA Macros for WPXStatsFinal.xlsm

## SyncArchivedHeaders

Keeps the "Archived Players" header row in sync with "Player Tracking".

When a new week is added to "Player Tracking", run this macro to automatically add the matching `Weekly Total (...)` and `Weekly Rank (...)` headers to "Archived Players".

### Setup (one-time)

1. Open `WPXStatsFinal.xlsm` in Excel
2. Press **Alt + F11** to open the VBA editor
3. Click **Insert → Module**
4. Copy the contents of `SyncArchivedHeaders.bas` into the new module
5. Close the VBA editor

### Run manually

1. Press **Alt + F8**
2. Select `SyncArchivedHeaders`
3. Click **Run**

### Auto-run on file open (optional)

To sync headers automatically every time the file opens:

1. In the VBA editor, double-click **ThisWorkbook** in the Project Explorer
2. Add:
   ```vb
   Private Sub Workbook_Open()
       SyncArchivedHeaders
   End Sub
   ```
3. Save and close — headers will sync on every open
