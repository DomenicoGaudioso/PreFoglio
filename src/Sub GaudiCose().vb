Sub GaudiCose()
    Dim objShell As Object
    Dim PythonExe, PythonScript As String
    Dim sWorkbookPath As String
    Dim FileName As String
    Dim KillFile As String
    Dim i As Integer
    Dim FileName2 As String
    Dim targetWorkbook As Workbook
    Dim targetSheet As Worksheet

    ' Imposta il workbook e la sheet attuali
    Set targetWorkbook = ActiveWorkbook
    Set targetSheet = ActiveSheet

    ' Cancella il vecchio file se presente
    sWorkbookPath = targetSheet.Range("A3").Value
    FileName = VBA.FileSystem.Dir(sWorkbookPath)
    If FileName <> VBA.Constants.vbNullString Then
        KillFile = sWorkbookPath
        Kill KillFile
    End If

    Set objShell = VBA.CreateObject("Wscript.Shell")

    ' Lancio Python
    PythonExe = """Z:\tools\Portable Python-3.9.13 x64\App\Python\python.exe"""
    PythonScript = "Z:\tools\GaudiCose\src\Run_PreFoglioPy.py"
    objShell.Run PythonExe & " " & PythonScript, 1, True

    ' Aspetta che il file Excel generato sia presente nella cartella
    For i = 1 To 500
        Application.Wait Now + TimeValue("00:00:01")
        FileName2 = VBA.FileSystem.Dir(sWorkbookPath)
        
        If FileName2 <> VBA.Constants.vbNullString Then
            Set wb = Workbooks.Open(sWorkbookPath)
            wb.Worksheets(1).Range("A4:F18000").Copy
            targetSheet.Range("A4").PasteSpecial Paste:=xlPasteValues
            wb.Close False
            Exit For
        End If
    Next i

End Sub




