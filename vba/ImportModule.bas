Attribute VB_Name = "ImportModule"

' ============================================================
' Refresh
' Checks every uncategorized row in ImportTable against
' RulesTable (cols J-M: Match, Type, Kategorie, Subkategorie).
' Skips rows that already have a Type filled in.
' Called automatically after ImportPending, or manually via button.
' ============================================================
Sub Refresh()
    ApplyRules showResult:=True
End Sub

Private Sub ApplyRules(showResult As Boolean)
    Application.ScreenUpdating = False
    Application.StatusBar = "Regeln werden angewendet..."

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Import")

    Dim importTbl As ListObject
    Set importTbl = ws.ListObjects("ImportTable")

    Dim rulesTbl As ListObject
    Set rulesTbl = ws.ListObjects("RulesTable")

    If importTbl.ListRows.Count = 0 Then
        Application.ScreenUpdating = True
        Application.StatusBar = False
        Exit Sub
    End If

    ' ImportTable column positions within ListRow.Range:
    '   1=Datum(B), 2=Betrag(C), 3=Empfaenger(D), 4=Type(E),
    '   5=Kategorie(F), 6=Subkategorie(G), 7=Details(H)
    ' RulesTable column positions:
    '   1=Match(J), 2=Type(K), 3=Kategorie(L), 4=Subkategorie(M)

    Dim matched As Long
    matched = 0

    Dim importRow As ListRow
    For Each importRow In importTbl.ListRows
        ' Only process rows that have an Empfaenger but no Type yet
        Dim empfaenger As String
        Dim typeVal As String
        empfaenger = importRow.Range.Cells(1, 3).Value
        typeVal = importRow.Range.Cells(1, 4).Value

        If empfaenger <> "" And typeVal = "" Then
            Dim amount As Double
            amount = importRow.Range.Cells(1, 2).Value

            Dim ruleRow As ListRow
            For Each ruleRow In rulesTbl.ListRows
                Dim matchVal As String
                matchVal = ruleRow.Range.Cells(1, 1).Value

                If matchVal <> "" And InStr(1, UCase(empfaenger), UCase(matchVal)) > 0 Then
                    Dim ruleType As String
                    ruleType = ruleRow.Range.Cells(1, 2).Value

                    Dim isIncome As Boolean
                    isIncome = (ruleType = "Einkommen")

                    If (isIncome And amount > 0) Or (Not isIncome And amount < 0) Then
                        importRow.Range.Cells(1, 4).Value = ruleType
                        importRow.Range.Cells(1, 5).Value = ruleRow.Range.Cells(1, 3).Value
                        importRow.Range.Cells(1, 6).Value = ruleRow.Range.Cells(1, 4).Value
                        matched = matched + 1
                        Exit For
                    End If
                End If
            Next ruleRow
        End If
    Next importRow

    Application.ScreenUpdating = True
    Application.StatusBar = False

    If showResult Then
        MsgBox matched & " Zeile(n) automatisch kategorisiert.", vbInformation
    End If
End Sub


' ============================================================
' Transfer
' Moves fully categorised rows (Datum + Betrag + Type + Kategorie
' all filled) from ImportTable to the Tracking table on the
' Budget Tracking sheet, then sorts Tracking by Datum ascending.
'
' ImportTable columns:  1=Datum(B), 2=Betrag(C), 3=Empfaenger(D),
'                       4=Type(E),  5=Kategorie(F), 6=Subkategorie(G),
'                       7=Details(H)
' Tracking columns:     1=Datum(C), 2=Type(D), 3=Kategorie(E),
'                       4=Subkategorie(F), 5=Betrag(G), 6=Empfaenger(H),
'                       7=Details(I)
' ============================================================
Sub Transfer()
    Dim wsImport As Worksheet
    Dim wsTracking As Worksheet
    Set wsImport = ThisWorkbook.Sheets("Import")
    Set wsTracking = ThisWorkbook.Sheets("Budget Tracking")

    Dim importTbl As ListObject
    Set importTbl = wsImport.ListObjects("ImportTable")

    Dim trkTbl As ListObject
    Set trkTbl = wsTracking.ListObjects("Tracking")

    If importTbl.ListRows.Count = 0 Then
        MsgBox "Keine Transaktionen zum " & Chr(220) & "bertragen.", vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Collect row indices to transfer first (delete after loop to avoid index shifting)
    Dim rowIndices() As Long
    ReDim rowIndices(1 To importTbl.ListRows.Count)
    Dim idxCount As Long
    idxCount = 0
    Dim transferred As Long
    transferred = 0

    Dim i As Long
    For i = 1 To importTbl.ListRows.Count
        Dim r As ListRow
        Set r = importTbl.ListRows(i)

        Dim hasDatum As Boolean
        Dim hasBetrag As Boolean
        Dim hasType As Boolean
        Dim hasKategorie As Boolean
        hasDatum = r.Range.Cells(1, 1).Value <> ""
        hasBetrag = r.Range.Cells(1, 2).Value <> ""
        hasType = r.Range.Cells(1, 4).Value <> ""
        hasKategorie = r.Range.Cells(1, 5).Value <> ""

        If hasDatum And hasBetrag And hasType And hasKategorie Then
            Dim newRow As ListRow
            Set newRow = trkTbl.ListRows.Add

            newRow.Range.Cells(1, 1).Value = r.Range.Cells(1, 1).Value  ' Datum
            newRow.Range.Cells(1, 1).NumberFormat = "DD.MM.YYYY"
            newRow.Range.Cells(1, 2).Value = r.Range.Cells(1, 4).Value  ' Type
            newRow.Range.Cells(1, 3).Value = r.Range.Cells(1, 5).Value  ' Kategorie
            newRow.Range.Cells(1, 4).Value = r.Range.Cells(1, 6).Value  ' Subkategorie
            newRow.Range.Cells(1, 5).Value = r.Range.Cells(1, 2).Value  ' Betrag
            newRow.Range.Cells(1, 6).Value = r.Range.Cells(1, 3).Value  ' Empfaenger
            newRow.Range.Cells(1, 7).Value = r.Range.Cells(1, 7).Value  ' Details

            idxCount = idxCount + 1
            rowIndices(idxCount) = i
            transferred = transferred + 1
        End If
    Next i

    ' Delete transferred rows in reverse order to keep indices valid
    Dim j As Long
    For j = idxCount To 1 Step -1
        importTbl.ListRows(rowIndices(j)).Delete
    Next j

    ' Sort Tracking by Datum (col 1) ascending — oldest on top
    If transferred > 0 Then
        With trkTbl.Sort
            .SortFields.Clear
            .SortFields.Add Key:=trkTbl.ListColumns(1).DataBodyRange, _
                SortOn:=xlSortOnValues, Order:=xlAscending
            .Header = xlYes
            .Apply
        End With
    End If

    Application.ScreenUpdating = True

    If transferred = 0 Then
        MsgBox "Keine vollst" & Chr(228) & "ndig kategorisierten Transaktionen gefunden.", vbInformation
    Else
        MsgBox transferred & " Transaktion(en) " & Chr(252) & "bertragen.", vbInformation
    End If
End Sub


' ============================================================
' ImportPending
' Reads data/pending_import.tsv (written by main.py),
' appends new rows to ImportTable (skipping duplicates),
' deletes the staging file, then auto-runs RefreshRules.
' Called automatically by Workbook_Open after Python exits.
' ============================================================
Sub ImportPending()
    Dim stagingPath As String
    stagingPath = ThisWorkbook.Path & "\data\pending_import.tsv"

    If Dir(stagingPath) = "" Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Import")

    Dim tbl As ListObject
    Set tbl = ws.ListObjects("ImportTable")

    Dim dataStartRow As Long
    dataStartRow = tbl.HeaderRowRange.Row + 1

    ' Build deduplication set from existing table rows
    Dim existing As Object
    Set existing = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = dataStartRow To dataStartRow + tbl.ListRows.Count - 1
        If ws.Cells(i, 2).Value <> "" Then
            Dim k As String
            k = Format(ws.Cells(i, 2).Value, "yyyy-mm-dd") & "|" & _
                CLng(ws.Cells(i, 3).Value * 100) & "|" & ws.Cells(i, 4).Value
            existing(k) = True
        End If
    Next i

    ' Find first empty row in col B from data start
    Dim nextRow As Long
    nextRow = dataStartRow
    Do While ws.Cells(nextRow, 2).Value <> ""
        nextRow = nextRow + 1
    Loop

    ' Read TSV and import
    Dim fileNum As Integer
    fileNum = FreeFile
    Open stagingPath For Input As #fileNum

    Dim lineText As String
    Dim isHeader As Boolean
    isHeader = True
    Dim imported As Long
    imported = 0

    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        If isHeader Then
            isHeader = False
        ElseIf Len(Trim(lineText)) > 0 Then
            Dim parts() As String
            parts = Split(lineText, Chr(9))
            If UBound(parts) >= 3 Then
                Dim rowKey As String
                rowKey = parts(0) & "|" & CLng(Val(parts(1)) * 100) & "|" & parts(2)

                If Not existing.Exists(rowKey) Then
                    ws.Cells(nextRow, 2).Value = DateSerial( _
                        CLng(Left(parts(0), 4)), _
                        CLng(Mid(parts(0), 6, 2)), _
                        CLng(Right(parts(0), 2)))
                    ws.Cells(nextRow, 3).Value = Val(parts(1))
                    ws.Cells(nextRow, 4).Value = parts(2)
                    ws.Cells(nextRow, 8).Value = parts(3)
                    existing(rowKey) = True
                    nextRow = nextRow + 1
                    imported = imported + 1
                End If
            End If
        End If
    Loop
    Close #fileNum

    Kill stagingPath

    If imported > 0 Then
        ' Auto-apply rules silently, then notify
        ApplyRules showResult:=False
        MsgBox imported & " Transaktion(en) importiert.", vbInformation
    End If
End Sub
