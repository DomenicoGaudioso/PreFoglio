Sub GaudiCose()
    Dim objShell As Object
    Dim PythonExe As String, PythonScript As String
    Dim wb As Workbook
    Dim sWorkbookPath As String
    Dim startTime As Double
    Dim timeOut As Double

    ' Legge il percorso dal foglio excel
    sWorkbookPath = ActiveSheet.Range("A3").Value
    
    ' Controlla se il file esiste e lo cancella se presente
    If Len(Dir(sWorkbookPath)) > 0 Then
        Kill sWorkbookPath
    End If

    ' Crea un nuovo oggetto Shell
    Set objShell = CreateObject("Wscript.Shell")

    ' Percorso dell'eseguibile Python e dello script
    PythonExe = "Z:\tools\Portable Python-3.9.13 x64\App\Python\python.exe"
    PythonScript = "Z:\tools\GaudiCose\src\Run_PreFoglioPy.py"

    ' Esegue lo script Python e attende il suo completamento
    objShell.Run """" & PythonExe & """ """ & PythonScript & """", 1, True

    ' Disattiva gli avvisi
    Application.DisplayAlerts = False

    ' Aspetta che il file Excel generato sia presente nella cartella
    startTime = Timer
    timeOut = 100 ' Tempo massimo di attesa in secondi
    Do While Len(Dir(sWorkbookPath)) = 0 And Timer - startTime < timeOut
        DoEvents ' Mantiene attiva l'interfaccia utente
        Application.Wait Now + TimeValue("00:00:01")
    Loop

    If Len(Dir(sWorkbookPath)) > 0 Then
        Set wb = Workbooks.Open(sWorkbookPath)
        wb.Worksheets(1).Range("A4:F18000").Copy
        ThisWorkbook.Worksheets("SFoglio_nuovo").Range("A4").PasteSpecial xlPasteValues
        wb.Close False
    End If

    ' Riattiva gli avvisi
    Application.DisplayAlerts = True
End Sub
