' SyncArchivedHeaders.bas
' Copies the weekly column headers from "Player Tracking" row 1
' into "Archived Players" row 1 so the two sheets stay in sync.
'
' USAGE
'   1. Open the VBA editor (Alt+F11) in WPXStatsFinal.xlsm
'   2. Insert > Module, then paste this file's contents
'   3. Run SyncArchivedHeaders manually, or wire it to Workbook_Open
'      (see SetupAutoSync at the bottom)
'
' WHAT IT DOES
'   - Reads every header in "Player Tracking" row 1 from column G onward
'   - Writes them into the same columns of "Archived Players" row 1
'   - Clears any leftover headers in Archived Players that no longer
'     exist in Player Tracking (handles weeks being removed)
'   - Leaves columns A-F (Name, Overall Total, etc.) untouched

Option Explicit

Public Sub SyncArchivedHeaders()

    Const PT_SHEET  As String = "Player Tracking"
    Const AP_SHEET  As String = "Archived Players"
    Const START_COL As Long = 7  ' Column G (weekly columns begin here)

    Dim wsPT As Worksheet
    Dim wsAP As Worksheet

    On Error Resume Next
    Set wsPT = ThisWorkbook.Sheets(PT_SHEET)
    Set wsAP = ThisWorkbook.Sheets(AP_SHEET)
    On Error GoTo 0

    If wsPT Is Nothing Then
        MsgBox "Sheet """ & PT_SHEET & """ not found.", vbExclamation
        Exit Sub
    End If
    If wsAP Is Nothing Then
        MsgBox "Sheet """ & AP_SHEET & """ not found.", vbExclamation
        Exit Sub
    End If

    ' Find the last used column in Player Tracking row 1
    Dim lastCol As Long
    lastCol = wsPT.Cells(1, wsPT.Columns.Count).End(xlToLeft).Column

    If lastCol < START_COL Then
        MsgBox "No weekly headers found in " & PT_SHEET & ".", vbInformation
        Exit Sub
    End If

    ' Copy each header from Player Tracking to Archived Players
    Dim col As Long
    For col = START_COL To lastCol
        wsAP.Cells(1, col).Value = wsPT.Cells(1, col).Value
    Next col

    ' Clear any extra headers in Archived Players beyond what PT has
    Dim apLastCol As Long
    apLastCol = wsAP.Cells(1, wsAP.Columns.Count).End(xlToLeft).Column

    If apLastCol > lastCol Then
        wsAP.Range(wsAP.Cells(1, lastCol + 1), wsAP.Cells(1, apLastCol)).ClearContents
    End If

    MsgBox "Archived Players headers synced (" & (lastCol - START_COL + 1) & " columns).", vbInformation

End Sub


' Optional: call from Workbook_Open so headers sync every time the file opens.
' Paste the Sub below into ThisWorkbook (not a standard module).
'
' Private Sub Workbook_Open()
'     SyncArchivedHeaders
' End Sub
