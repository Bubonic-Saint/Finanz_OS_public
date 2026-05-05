Attribute VB_Name = "ThisWorkbook"

' ============================================================
' Workbook_Open
' Paste this code into the ThisWorkbook module in the VBA
' editor (Alt+F11 → double-click "ThisWorkbook" in the
' Project pane). Do NOT import it as a module.
'
' Flow:
'   1. Python script runs and writes data/pending_import.tsv
'   2. ImportPending reads the TSV and writes rows to Import
'   3. Staging file is deleted
' ============================================================
Private Sub Workbook_Open()
    Dim pythonExe As String
    Dim scriptPath As String

    ' Use the venv Python if present, otherwise fall back to system Python
    pythonExe = ThisWorkbook.Path & "\venv\Scripts\python.exe"
    If Dir(pythonExe) = "" Then pythonExe = "python"

    scriptPath = ThisWorkbook.Path & "\scripts\main.py"

    Application.StatusBar = "Transaktionen werden abgerufen..."

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    Dim exitCode As Long
    exitCode = wsh.Run("""" & pythonExe & """ """ & scriptPath & """", 0, True)

    Application.StatusBar = False

    If exitCode <> 0 Then
        MsgBox "Fehler beim Abrufen der Transaktionen (Exit Code: " & exitCode & ").", vbExclamation
        Exit Sub
    End If

    ImportPending
End Sub
