Option Explicit

Sub Btn_Clear_Synthese()
    Dim baseDir As String, archivePath As String, timestamp As String
    Dim ws As Worksheet, wsLC As Worksheet
    Dim newWb As Workbook, newWs As Worksheet, newWsLC As Worksheet
    Dim lastRow As Long, hasData As Boolean, i As Long
    Dim sheetNamesToDelete As Collection
    Dim sheetName As Variant, sht As Worksheet

    If MsgBox("Voulez-vous lancer l'archivage de la feuille SYNTHESE ?" & vbCrLf & _
              "Un nouveau fichier d'archive sera cree.", _
              vbYesNo + vbQuestion, "Confirmation d'archivage") = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_SYNTHESE)
    If Err.Number <> 0 Then MsgBox "La feuille SYNTHESE est introuvable.", vbCritical, "Erreur": Exit Sub
    Set wsLC = ThisWorkbook.Sheets(SHEET_LC)
    If Err.Number <> 0 Then MsgBox "La feuille LC est introuvable.", vbCritical, "Erreur": Exit Sub
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    hasData = (lastRow >= 3)
    timestamp = Format(Now, "ddmmyyyy_HHMMSS")
    archivePath = baseDir & "\Archived\Archive_SYNTHESE_" & timestamp & ".xlsx"

    Application.StatusBar = "Creation du fichier d'archive avec les feuilles SYNTHESE et LC..."
    Set newWb = Workbooks.Add

    ws.Copy Before:=newWb.Sheets(1)
    Set newWs = newWb.Sheets(SHEET_SYNTHESE)
    wsLC.Copy After:=newWb.Sheets(SHEET_SYNTHESE)
    Set newWsLC = newWb.Sheets(SHEET_LC)

    Application.DisplayAlerts = False

    ' Remove shapes and OLEObjects from both sheets
    For i = newWs.Shapes.Count To 1 Step -1: newWs.Shapes(i).Delete: Next i
    For i = newWs.OLEObjects.Count To 1 Step -1: newWs.OLEObjects(i).Delete: Next i
    For i = newWsLC.Shapes.Count To 1 Step -1: newWsLC.Shapes(i).Delete: Next i
    For i = newWsLC.OLEObjects.Count To 1 Step -1: newWsLC.OLEObjects(i).Delete: Next i

    ' Delete default sheets
    Set sheetNamesToDelete = New Collection
    For Each sht In newWb.Sheets
        If sht.Name <> SHEET_SYNTHESE And sht.Name <> SHEET_LC Then sheetNamesToDelete.Add sht.Name
    Next sht
    For Each sheetName In sheetNamesToDelete: newWb.Sheets(sheetName).Delete: Next sheetName
    Application.DisplayAlerts = True

    newWs.Move Before:=newWb.Sheets(1)

    Application.DisplayAlerts = False
    On Error Resume Next
    newWb.SaveAs archivePath, xlOpenXMLWorkbook
    If Err.Number <> 0 Then
        Application.DisplayAlerts = True
        MsgBox "Erreur lors de l'enregistrement de l'archive : " & Err.Description, vbCritical, "Erreur"
        newWb.Close SaveChanges:=False
        GoTo ErrorHandler
    End If
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler

    newWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.StatusBar = False

    If hasData Then
        ws.Rows("3:" & lastRow).Delete Shift:=xlUp
        MsgBox "SYNTHESE archivee et videe." & vbCrLf & "Enregistre dans : " & archivePath, vbInformation, "Archivage termine"
    Else
        MsgBox "Archive creee (SYNTHESE etait deja vide)." & vbCrLf & "Enregistre dans : " & archivePath, vbInformation, "Archivage termine"
    End If
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not newWb Is Nothing Then newWb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Sub Btn_Collect_RM_Data()
    Dim baseDir As String, xmlPath As String
    Dim ws As Worksheet, wsLC As Worksheet
    Dim result As Collection
    Dim exitCode As Long, rowsImported As Long, startRow As Long

    If MsgBox("Voulez-vous lancer l'import des donnees de pointage ?" & vbCrLf & _
              "Cette action importera les donnees de RM_Collaborateurs dans la feuille SYNTHESE.", _
              vbYesNo + vbQuestion, "Confirmation d'import") = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_SYNTHESE)
    If Err.Number <> 0 Then MsgBox "La feuille SYNTHESE est introuvable.", vbCritical, "Erreur": Exit Sub
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    Application.StatusBar = "Export des donnees de pointage depuis les fichiers collaborateurs..."
    exitCode = RunCommand(PYTHONEXE & "--basedir " & """" & baseDir & """" & " pointage")
    Application.StatusBar = False

    If exitCode <> 0 Then
        MsgBox "Erreur lors de l'export des donnees de pointage. Code de sortie : " & exitCode, vbCritical, "Erreur"
        GoTo ErrorHandler
    End If

    xmlPath = baseDir & "\pointage_output.xml"
    If Dir(xmlPath) = "" Then
        MsgBox "Erreur : pointage_output.xml n'a pas ete cree.", vbCritical, "Erreur"
        GoTo ErrorHandler
    End If

    Set result = LoadXMLTable(xmlPath)
    If result Is Nothing Then
        MsgBox "Erreur lors du chargement des donnees XML.", vbCritical, "Erreur"
        GoTo ErrorHandler
    End If

    startRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    If startRow < 3 Then startRow = 3
    rowsImported = 0

    ImportPointageRows ws, result, startRow, rowsImported, 11, 53

    If rowsImported > 0 Then
        Set wsLC = ThisWorkbook.Sheets(SHEET_LC)
        UpdateSyntheseFromLC ws, wsLC, startRow, startRow + rowsImported - 1
    End If

    ApplySyntheseRowColoring ws, startRow, 11, 53, 35
    If Dir(xmlPath) <> "" Then Kill xmlPath

    MsgBox IIf(rowsImported > 0, _
           rowsImported & " ligne(s) importee(s) dans SYNTHESE.", _
           "Aucune donnee a importer."), vbInformation, "Import termine"

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Sub Btn_Collect_RM_Data_Reset()
    Dim baseDir As String, xmlPath As String
    Dim ws As Worksheet, wsLC As Worksheet
    Dim result As Collection
    Dim exitCode As Long, rowsImported As Long, startRow As Long

    If MsgBox("Voulez-vous lancer l'import des donnees de pointage ?" & vbCrLf & _
              "Cette action importera les donnees dans SYNTHESE, archivera RM_Collaborateurs et creera de nouvelles interfaces.", _
              vbYesNo + vbQuestion, "Confirmation d'import") = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_SYNTHESE)
    If Err.Number <> 0 Then MsgBox "La feuille SYNTHESE est introuvable.", vbCritical, "Erreur": Exit Sub
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    ' Step 1: Collect pointage
    Application.StatusBar = "Export des donnees de pointage depuis les fichiers collaborateurs..."
    exitCode = RunCommand(PYTHONEXE & "--basedir " & """" & baseDir & """" & " pointage")
    Application.StatusBar = False

    If exitCode <> 0 Then
        MsgBox "Erreur lors de l'export des donnees de pointage. Code de sortie : " & exitCode, vbCritical, "Erreur"
        GoTo ErrorHandler
    End If

    xmlPath = baseDir & "\pointage_output.xml"
    If Dir(xmlPath) = "" Then
        MsgBox "Erreur : pointage_output.xml n'a pas ete cree.", vbCritical, "Erreur"
        GoTo ErrorHandler
    End If

    Set result = LoadXMLTable(xmlPath)
    If result Is Nothing Then
        MsgBox "Erreur lors du chargement des donnees XML.", vbCritical, "Erreur"
        GoTo ErrorHandler
    End If

    startRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    If startRow < 3 Then startRow = 3
    rowsImported = 0

    ImportPointageRows ws, result, startRow, rowsImported, 11, 53

    If rowsImported > 0 Then
        Set wsLC = ThisWorkbook.Sheets(SHEET_LC)
        UpdateSyntheseFromLC ws, wsLC, startRow, startRow + rowsImported - 1
    End If

    ApplySyntheseRowColoring ws, startRow, 11, 53, 35
    If Dir(xmlPath) <> "" Then Kill xmlPath

    MsgBox IIf(rowsImported > 0, _
           rowsImported & " ligne(s) importee(s) dans SYNTHESE.", _
           "Aucune donnee a importer."), vbInformation, "Import termine"

    ' Step 2: Cleanup + recreate interfaces
    CleanupGestionInterfaces

    If Not CreateCollabsXML(baseDir) Then
        MsgBox "Erreur lors de la creation de collabs.xml. Operation annulee.", vbCritical, "Erreur"
        GoTo ErrorHandler
    End If

    ' Step 3: Delete existing interfaces
    Application.StatusBar = "Suppression des interfaces..."
    exitCode = RunCommand(PYTHONEXE & "--basedir " & """" & baseDir & """" & " delete --force")
    Application.StatusBar = False
    If exitCode <> 0 Then
        MsgBox "Erreur lors de la suppression des interfaces. Code de sortie : " & exitCode, vbCritical, "Erreur"
        GoTo ErrorHandler
    End If
    MsgBox "Interfaces supprimees avec succes.", vbInformation, "Suppression terminee"

    ' Step 4: Create new interfaces
    Application.StatusBar = "Creation des interfaces collaborateurs..."
    exitCode = RunCommand(PYTHONEXE & "--basedir " & """" & baseDir & """" & " create")
    Application.StatusBar = False
    If exitCode <> 0 Then
        MsgBox "Erreur lors de la creation des interfaces. Code de sortie : " & exitCode, vbCritical, "Erreur"
        GoTo ErrorHandler
    End If
    MsgBox "Interfaces collaborateurs creees avec succes.", vbInformation, "Creation terminee"

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub
