' LookupPlayerID.bas
' Looks up a player name in Column A of the active sheet and writes
' the matching ID to Column V by searching the Roster sheet first,
' then falling back to the Archived Players sheet (Column CK).
'
' USAGE
'   1. Open the VBA editor (Alt+F11) in WPXStatsFinal.xlsm
'   2. Insert > Module (or reuse the existing one), then paste this code
'   3. Navigate to the sheet you want to populate (e.g. a weekly sheet)
'   4. Run LookupPlayerIDs (Alt+F8)
'
' HOW IT WORKS
'   For each non-empty name in Column A (starting at row 2):
'     1. Searches all Roster name columns (B, F, J, N) for an exact match
'     2. If not found, searches Roster AKA columns (D, H, L, P)
'     3. If still not found, searches Archived Players Column A for the name
'        and returns the ID from Archived Players Column CK
'     4. Writes the ID to Column V of the current row
'     5. Skips rows that already have a value or formula in Column V

Option Explicit

Public Sub LookupPlayerIDs()

    Dim wsActive As Worksheet
    Set wsActive = ActiveSheet

    Dim wsRoster As Worksheet
    Dim wsArchived As Worksheet

    On Error Resume Next
    Set wsRoster = ThisWorkbook.Sheets("Roster")
    Set wsArchived = ThisWorkbook.Sheets("Archived Players")
    On Error GoTo 0

    If wsRoster Is Nothing Then
        MsgBox "Sheet ""Roster"" not found.", vbExclamation
        Exit Sub
    End If
    If wsArchived Is Nothing Then
        MsgBox "Sheet ""Archived Players"" not found.", vbExclamation
        Exit Sub
    End If

    ' Target column for IDs on weekly sheets
    Const ID_COL As Long = 22  ' Column V

    ' Find last row with data in Column A
    Dim lastRow As Long
    lastRow = wsActive.Cells(wsActive.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No player names found in Column A.", vbInformation
        Exit Sub
    End If

    Dim r As Long
    Dim playerName As String
    Dim foundID As Variant
    Dim filled As Long
    Dim skipped As Long

    For r = 2 To lastRow
        ' Skip if Column V already has a value or formula
        If wsActive.Cells(r, ID_COL).Value <> "" Then
            skipped = skipped + 1
            GoTo NextRow
        End If

        playerName = Trim(CStr(wsActive.Cells(r, 1).Value))
        If playerName = "" Then GoTo NextRow

        ' Search Roster
        foundID = FindInRoster(wsRoster, playerName)

        ' Fallback: search Archived Players
        If IsEmpty(foundID) Then
            foundID = FindInArchived(wsArchived, playerName)
        End If

        If Not IsEmpty(foundID) Then
            wsActive.Cells(r, ID_COL).Value = foundID
            filled = filled + 1
        End If

NextRow:
    Next r

    MsgBox "Done!  Filled: " & filled & "  |  Skipped (already had ID): " & skipped, vbInformation

End Sub


' ---------------------------------------------------------------------------
'  Search all 4 Roster name+AKA column pairs and return the ID if found.
'  Roster layout:  A/B/C/D  E/F/G/H  I/J/K/L  M/N/O/P
'    ID cols:      A        E        I        M
'    Name cols:    B        F        J        N
'    AKA cols:     D        H        L        P
' ---------------------------------------------------------------------------
Private Function FindInRoster(wsRoster As Worksheet, playerName As String) As Variant

    Dim idCols As Variant
    Dim nameCols As Variant
    Dim akaCols As Variant
    idCols = Array(1, 5, 9, 13)    ' A, E, I, M
    nameCols = Array(2, 6, 10, 14) ' B, F, J, N
    akaCols = Array(4, 8, 12, 16)  ' D, H, L, P

    Dim i As Long
    Dim rng As Range
    Dim found As Range
    Dim lastRow As Long

    ' Search name columns first
    For i = LBound(nameCols) To UBound(nameCols)
        lastRow = wsRoster.Cells(wsRoster.Rows.Count, nameCols(i)).End(xlUp).Row
        If lastRow >= 3 Then
            Set rng = wsRoster.Range(wsRoster.Cells(3, nameCols(i)), wsRoster.Cells(lastRow, nameCols(i)))
            Set found = rng.Find(What:=playerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
            If Not found Is Nothing Then
                FindInRoster = wsRoster.Cells(found.Row, idCols(i)).Value
                Exit Function
            End If
        End If
    Next i

    ' Search AKA columns
    For i = LBound(akaCols) To UBound(akaCols)
        lastRow = wsRoster.Cells(wsRoster.Rows.Count, akaCols(i)).End(xlUp).Row
        If lastRow >= 3 Then
            Set rng = wsRoster.Range(wsRoster.Cells(3, akaCols(i)), wsRoster.Cells(lastRow, akaCols(i)))
            Set found = rng.Find(What:=playerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
            If Not found Is Nothing Then
                FindInRoster = wsRoster.Cells(found.Row, idCols(i)).Value
                Exit Function
            End If
        End If
    Next i

    FindInRoster = Empty

End Function


' ---------------------------------------------------------------------------
'  Search Archived Players Column A for the name and return Column CK value.
' ---------------------------------------------------------------------------
Private Function FindInArchived(wsArchived As Worksheet, playerName As String) As Variant

    Const NAME_COL As Long = 1   ' Column A
    Const ID_COL   As Long = 89  ' Column CK

    Dim lastRow As Long
    lastRow = wsArchived.Cells(wsArchived.Rows.Count, NAME_COL).End(xlUp).Row

    If lastRow < 2 Then
        FindInArchived = Empty
        Exit Function
    End If

    Dim rng As Range
    Set rng = wsArchived.Range(wsArchived.Cells(2, NAME_COL), wsArchived.Cells(lastRow, NAME_COL))

    Dim found As Range
    Set found = rng.Find(What:=playerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

    If Not found Is Nothing Then
        Dim idVal As Variant
        idVal = wsArchived.Cells(found.Row, ID_COL).Value
        If idVal <> "" Then
            FindInArchived = idVal
        Else
            FindInArchived = Empty
        End If
    Else
        FindInArchived = Empty
    End If

End Function
