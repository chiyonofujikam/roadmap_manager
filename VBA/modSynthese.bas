Option Explicit

Sub Btn_Clear_Synthese()
    Dim baseDir As String, archivePath As String, timestamp As String
    Dim ws As Worksheet, wsLC As Worksheet
    Dim newWb As Workbook, newWs As Worksheet, newWsLC As Worksheet
    Dim lastRow As Long, hasData As Boolean, i As Long
    Dim sheetNamesToDelete As Collection
    Dim sheetName As Variant, sht As Worksheet

    If MsgBox("Do you want to proceed with archiving the SYNTHESE sheet?" & vbCrLf & _
              "A new archive file will be created.", _
              vbYesNo + vbQuestion, "Confirm Archiving") = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_SYNTHESE)
    If Err.Number <> 0 Then MsgBox "SYNTHESE sheet not found.", vbCritical, "Error": Exit Sub
    Set wsLC = ThisWorkbook.Sheets(SHEET_LC)
    If Err.Number <> 0 Then MsgBox "LC sheet not found.", vbCritical, "Error": Exit Sub
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    hasData = (lastRow >= 3)
    timestamp = Format(Now, "ddmmyyyy_HHMMSS")
    archivePath = baseDir & "\Archived\Archive_SYNTHESE_" & timestamp & ".xlsx"

    Application.StatusBar = "Creating archive file with SYNTHESE and LC sheets..."
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
        MsgBox "Error saving archive file: " & Err.Description, vbCritical, "Error"
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
        MsgBox "SYNTHESE archived and cleared." & vbCrLf & "Saved to: " & archivePath, vbInformation, "Archive Complete"
    Else
        MsgBox "Archive created (SYNTHESE was already empty)." & vbCrLf & "Saved to: " & archivePath, vbInformation, "Archive Complete"
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

    If MsgBox("Do you want to proceed with importing the pointage data?" & vbCrLf & _
              "This will import data from RM_Collaborateurs into the SYNTHESE sheet.", _
              vbYesNo + vbQuestion, "Confirm Import") = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_SYNTHESE)
    If Err.Number <> 0 Then MsgBox "SYNTHESE sheet not found.", vbCritical, "Error": Exit Sub
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    Application.StatusBar = "Exporting pointage data from collaborator files..."
    exitCode = RunCommand(PYTHONEXE & "--basedir " & """" & baseDir & """" & " pointage")
    Application.StatusBar = False

    If exitCode <> 0 Then
        MsgBox "Error exporting pointage data. Exit code: " & exitCode, vbCritical, "Error"
        GoTo ErrorHandler
    End If

    xmlPath = baseDir & "\pointage_output.xml"
    If Dir(xmlPath) = "" Then
        MsgBox "Error: pointage_output.xml was not created.", vbCritical, "Error"
        GoTo ErrorHandler
    End If

    Set result = LoadXMLTable(xmlPath)
    If result Is Nothing Then
        MsgBox "Error loading XML data.", vbCritical, "Error"
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
           rowsImported & " row(s) imported into SYNTHESE.", _
           "No data to import."), vbInformation, "Import Complete"

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

    If MsgBox("Do you want to proceed with importing the pointage data?" & vbCrLf & _
              "This will import data into the SYNTHESE sheet, archive RM_Collaborateurs and create new interfaces.", _
              vbYesNo + vbQuestion, "Confirm Import") = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_SYNTHESE)
    If Err.Number <> 0 Then MsgBox "SYNTHESE sheet not found.", vbCritical, "Error": Exit Sub
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    ' Step 1: Collect pointage
    Application.StatusBar = "Exporting pointage data from collaborator files..."
    exitCode = RunCommand(PYTHONEXE & "--basedir " & """" & baseDir & """" & " pointage")
    Application.StatusBar = False

    If exitCode <> 0 Then
        MsgBox "Error exporting pointage data. Exit code: " & exitCode, vbCritical, "Error"
        GoTo ErrorHandler
    End If

    xmlPath = baseDir & "\pointage_output.xml"
    If Dir(xmlPath) = "" Then
        MsgBox "Error: pointage_output.xml was not created.", vbCritical, "Error"
        GoTo ErrorHandler
    End If

    Set result = LoadXMLTable(xmlPath)
    If result Is Nothing Then
        MsgBox "Error loading XML data.", vbCritical, "Error"
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
           rowsImported & " row(s) imported into SYNTHESE.", _
           "No data to import."), vbInformation, "Import Complete"

    ' Step 2: Cleanup + recreate interfaces
    CleanupGestionInterfaces

    If Not CreateCollabsXML(baseDir) Then
        MsgBox "Error creating collabs.xml. Operation aborted.", vbCritical, "Error"
        GoTo ErrorHandler
    End If

    ' Step 3: Delete existing interfaces
    Application.StatusBar = "Deleting interfaces..."
    exitCode = RunCommand(PYTHONEXE & "--basedir " & """" & baseDir & """" & " delete --force")
    Application.StatusBar = False
    If exitCode <> 0 Then
        MsgBox "Error deleting interfaces. Exit code: " & exitCode, vbCritical, "Error"
        GoTo ErrorHandler
    End If
    MsgBox "Interfaces successfully deleted.", vbInformation, "Deletion Complete"

    ' Step 4: Create new interfaces
    Application.StatusBar = "Creating collaborator interfaces..."
    exitCode = RunCommand(PYTHONEXE & "--basedir " & """" & baseDir & """" & " create")
    Application.StatusBar = False
    If exitCode <> 0 Then
        MsgBox "Error creating interfaces. Exit code: " & exitCode, vbCritical, "Error"
        GoTo ErrorHandler
    End If
    MsgBox "Collaborator interfaces successfully created.", vbInformation, "Creation Complete"

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub
