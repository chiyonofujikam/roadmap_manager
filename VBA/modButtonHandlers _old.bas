Option Explicit

' =============================================================================
' BUTTON HANDLERS
' All button click event handlers for the Excel interface
' =============================================================================

Sub Btn_Create_RM()
    Dim createCommand As String
    Dim baseDir As String
    Dim exitCode As Long

    ' Clean up empty rows in Gestion_Interfaces sheet
    CleanupGestionInterfaces

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    If Not CreateCollabsXML(baseDir) Then
        MsgBox "Error creating collabs.xml file. Operation aborted.", vbCritical, "Error"
        GoTo ErrorHandler
    End If

    createCommand = PYTHONEXE & "--basedir " & """" & baseDir & """" & " create --way para"
    Application.StatusBar = "Creating collaborator interfaces..."
    exitCode = RunCommand(createCommand)
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

Sub Btn_Delete_RM()
    Dim deleteCommand As String
    Dim baseDir As String
    Dim forceDelete As VbMsgBoxResult
    Dim archiveChoice As VbMsgBoxResult
    Dim exitCode As Long

    forceDelete = MsgBox("Do you want to FORCE deletion of RM Interfaces?" & vbCrLf & _
                         "(This will delete all generated interfaces)", _
                         vbYesNo + vbQuestion, "Confirm Force Deletion")
    If forceDelete = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    archiveChoice = MsgBox("Do you want to ARCHIVE deleted interfaces?", _
                           vbYesNo + vbQuestion, "Archive Confirmation")

    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    deleteCommand = PYTHONEXE & "--basedir " & """" & baseDir & """" & " delete --force"

    If archiveChoice = vbYes Then
        deleteCommand = deleteCommand & " --archive"
        Application.StatusBar = "Archiving and deleting interfaces..."
    Else
        Application.StatusBar = "Deleting interfaces..."
    End If

    exitCode = RunCommand(deleteCommand)
    Application.StatusBar = False

    If exitCode <> 0 Then
        MsgBox "Error deleting interfaces. Exit code: " & exitCode, vbCritical, "Error"
        GoTo ErrorHandler
    End If

    If archiveChoice = vbYes Then
        MsgBox "Interfaces successfully archived and deleted.", vbInformation, "Deletion Complete"
    Else
        MsgBox "Interfaces successfully deleted.", vbInformation, "Deletion Complete"
    End If

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Sub Btn_Clear_Synthese()
    Dim baseDir As String
    Dim archiveConfirm As VbMsgBoxResult
    Dim ws As Worksheet
    Dim wsLC As Worksheet
    Dim lastRow As Long
    Dim archivePath As String
    Dim newWb As Workbook
    Dim newWs As Worksheet
    Dim newWsLC As Worksheet
    Dim timestamp As String
    Dim archivedFolder As String
    Dim hasData As Boolean
    Dim sht As Worksheet
    Dim sheetNamesToDelete As Collection
    Dim sheetName As Variant
    Dim i As Long

    archiveConfirm = MsgBox("Do you want to proceed with archiving the SYNTHESE sheet?" & vbCrLf & _
                            "A new archive file will be created.", _
                            vbYesNo + vbQuestion, "Confirm Archiving")
    If archiveConfirm = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    ' Check if SYNTHESE sheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_SYNTHESE)
    If Err.Number <> 0 Then
        MsgBox "SYNTHESE sheet not found.", vbCritical, "Error"
        Exit Sub
    End If
    On Error GoTo 0

    ' Check if LC sheet exists
    On Error Resume Next
    Set wsLC = ThisWorkbook.Sheets(SHEET_LC)
    If Err.Number <> 0 Then
        MsgBox "LC sheet not found.", vbCritical, "Error"
        Exit Sub
    End If
    On Error GoTo 0

    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    ' Determine if there's data to archive
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).row
    hasData = (lastRow >= 3)

    ' Set Archived folder path (folder always exists)
    archivedFolder = baseDir & "\Archived"

    ' Generate timestamp in format: ddmmyyyy_HHMMSS
    timestamp = Format(Now, "ddmmyyyy_HHMMSS")
    archivePath = archivedFolder & "\Archive_SYNTHESE_" & timestamp & ".xlsx"

    Application.StatusBar = "Creating archive file with SYNTHESE and LC sheets..."

    ' Create new workbook
    Set newWb = Workbooks.Add

    ' Copy SYNTHESE sheet with all formatting to new workbook
    ws.Copy Before:=newWb.Sheets(1)
    Set newWs = newWb.Sheets(SHEET_SYNTHESE)

    ' Copy LC sheet with all formatting to new workbook
    wsLC.Copy After:=newWb.Sheets(SHEET_SYNTHESE)
    Set newWsLC = newWb.Sheets(SHEET_LC)

    ' Remove all shapes, buttons, rectangles, and form controls from copied sheets
    ' This removes macro assignments and interactive elements
    Application.DisplayAlerts = False
    For i = newWs.Shapes.Count To 1 Step -1
        newWs.Shapes(i).Delete
    Next i

    ' Remove OLEObjects (ActiveX controls) from SYNTHESE sheet
    For i = newWs.OLEObjects.Count To 1 Step -1
        newWs.OLEObjects(i).Delete
    Next i

    ' Remove shapes from LC sheet
    For i = newWsLC.Shapes.Count To 1 Step -1
        newWsLC.Shapes(i).Delete
    Next i

    ' Remove OLEObjects (ActiveX controls) from LC sheet
    For i = newWsLC.OLEObjects.Count To 1 Step -1
        newWsLC.OLEObjects(i).Delete
    Next i

    ' Delete all default sheets (they have generic names like "Sheet1")
    Set sheetNamesToDelete = New Collection

    ' Collect sheet names to delete (avoid deleting during iteration)
    For Each sht In newWb.Sheets
        If sht.Name <> SHEET_SYNTHESE And sht.Name <> SHEET_LC Then
            sheetNamesToDelete.Add sht.Name
        End If
    Next sht

    ' Delete collected sheets
    For Each sheetName In sheetNamesToDelete
        newWb.Sheets(sheetName).Delete
    Next sheetName
    Application.DisplayAlerts = True

    ' Ensure SYNTHESE is first sheet and LC is second
    newWs.Move Before:=newWb.Sheets(1)

    ' Save the archive file
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

    ' Close the archive workbook
    newWb.Close SaveChanges:=False

    Application.ScreenUpdating = True
    Application.StatusBar = False

    ' Delete SYNTHESE sheet rows if it had data (this removes both data and formatting/colors)
    If hasData Then
        ' Delete rows from bottom to top to avoid shifting issues
        ' Delete entire rows to remove both content and formatting (colors)
        ws.rows("3:" & lastRow).Delete Shift:=xlUp
        MsgBox "SYNTHESE sheet successfully archived and cleared." & vbCrLf & _
               "Archive file saved to: " & archivePath, vbInformation, "Archive Complete"
    Else
        MsgBox "Archive file created successfully." & vbCrLf & _
               "Archive file saved to: " & archivePath & vbCrLf & _
               "SYNTHESE sheet was already empty.", vbInformation, "Archive Complete"
    End If

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not newWb Is Nothing Then
        newWb.Close SaveChanges:=False
    End If
    On Error GoTo 0
End Sub

Sub Btn_Collect_RM_Data_Reset()
    Dim pointageCommand As String
    Dim deleteCommand As String
    Dim createCommand As String
    Dim baseDir As String
    Dim xmlPath As String
    Dim ws As Worksheet
    Dim wsLC As Worksheet
    Dim result As Collection
    Dim rowData As Collection
    Dim value As Variant
    Dim r As Long, c As Long
    Dim confirmation As VbMsgBoxResult
    Dim exitCode As Long
    Dim rowsImported As Long
    Dim startRow As Long

    confirmation = MsgBox("Do you want to proceed with importing the pointage data?" & vbCrLf & _
                          "This will import data into the SYNTHESE sheet, archive RM_Collaborateurs and create new interfaces.", _
                          vbYesNo + vbQuestion, "Confirm Import")
    If confirmation = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    ' Check if SYNTHESE sheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_SYNTHESE)
    If Err.Number <> 0 Then
        MsgBox "SYNTHESE sheet not found.", vbCritical, "Error"
        Exit Sub
    End If
    On Error GoTo 0

    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    ' Step 1: Collect pointage data from existing interfaces
    pointageCommand = PYTHONEXE & "--basedir " & """" & baseDir & """" & " pointage"
    Application.StatusBar = "Exporting pointage data from collaborator files..."
    exitCode = RunCommand(pointageCommand)
    Application.StatusBar = False

    If exitCode <> 0 Then
        MsgBox "Error exporting pointage data. Exit code: " & exitCode, vbCritical, "Error"
        GoTo ErrorHandler
    End If

    ' Verify XML file was created
    xmlPath = baseDir & "\pointage_output.xml"
    If Dir(xmlPath) = "" Then
        MsgBox "Error: pointage_output.xml file was not created.", vbCritical, "Error"
        GoTo ErrorHandler
    End If

    ' Load and import XML data
    Set result = LoadXMLTable(xmlPath)
    If result Is Nothing Then
        MsgBox "Error loading XML data. The file may be corrupted or empty.", vbCritical, "Error"
        GoTo ErrorHandler
    End If

    ' Find starting row for data import
    startRow = ws.Cells(ws.rows.Count, "A").End(xlUp).row + 1
    If startRow < 3 Then startRow = 3
    r = startRow
    rowsImported = 0

    ' Import data into SYNTHESE sheet (maps helper to BA) and color rows
    ImportPointageRows ws, result, startRow, rowsImported, 11, 53

    ' Update SYNTHESE columns H and I from LC table matching
    If rowsImported > 0 Then
        Set wsLC = ThisWorkbook.Sheets(SHEET_LC)
        UpdateSyntheseFromLC ws, wsLC, startRow, startRow + rowsImported - 1
    End If

    ' Apply coloring based on K1 totals (helper column after import, using BA)
    ApplySyntheseRowColoring ws, startRow, 11, 53, 35

    ' Clean up temporary XML file
    If Dir(xmlPath) <> "" Then Kill xmlPath

    ' Show import summary
    If rowsImported > 0 Then
        MsgBox "Pointage successfully imported from 'RM_Collaborateurs'." & vbCrLf & _
               rowsImported & " row(s) imported into SYNTHESE sheet.", vbInformation, "Import Complete"
    Else
        MsgBox "No data to import. The pointage file was empty.", vbInformation, "Import Complete"
    End If

    ' Step 2: Clean up empty rows in Gestion_Interfaces sheet
    CleanupGestionInterfaces

    ' Step 3: Create collabs.xml file (needed for delete and create commands)
    If Not CreateCollabsXML(baseDir) Then
        MsgBox "Error creating collabs.xml file. Operation aborted.", vbCritical, "Error"
        GoTo ErrorHandler
    End If

    ' Step 4: Delete existing interfaces
    deleteCommand = PYTHONEXE & "--basedir " & """" & baseDir & """" & " delete --force"
    Application.StatusBar = "Deleting interfaces..."
    exitCode = RunCommand(deleteCommand)
    Application.StatusBar = False

    If exitCode <> 0 Then
        MsgBox "Error deleting interfaces. Exit code: " & exitCode, vbCritical, "Error"
        GoTo ErrorHandler
    End If

    MsgBox "Interfaces successfully deleted.", vbInformation, "Deletion Complete"

    ' Step 5: Create new collaborator interfaces
    createCommand = PYTHONEXE & "--basedir " & """" & baseDir & """" & " create"
    Application.StatusBar = "Creating collaborator interfaces..."
    exitCode = RunCommand(createCommand)
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

Sub Btn_Update_LC()
    Dim baseDir As String
    Dim confirmation As VbMsgBoxResult
    Dim templatePath As String
    Dim rmFolder As String
    Dim wsLCSource As Worksheet
    Dim fileName As String
    Dim filePath As String
    Dim fileCount As Long
    Dim processedCount As Long
    Dim fileList As Collection
    Dim startTime As Double
    Dim endTime As Double
    Dim elapsedTime As Double
    Dim timeMessage As String

    confirmation = MsgBox("Do you want to proceed with updating the conditional lists (LC)?" & vbCrLf & _
                          "This will update LC in the template and all collaborator files.", _
                          vbYesNo + vbQuestion, "Confirm Update")
    If confirmation = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    On Error Resume Next
    Set wsLCSource = ThisWorkbook.Sheets(SHEET_LC)
    On Error GoTo ErrorHandler
    If wsLCSource Is Nothing Then
        MsgBox "LC sheet not found in the current workbook.", vbCritical, "Error"
        Exit Sub
    End If

    ' Start timing
    startTime = Timer

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrorHandler

    templatePath = baseDir & "\RM_template.xlsx"
    rmFolder = baseDir & "\RM_Collaborateurs"

    Application.StatusBar = "Updating LC in template and collaborator files..."

    ' Collect all files first
    Set fileList = New Collection
    fileList.Add templatePath
    
    fileName = Dir(rmFolder & "\RM_*.xlsx")
    Do While fileName <> ""
        If Left$(fileName, 2) <> "~$" Then
            fileList.Add rmFolder & "\" & fileName
        End If
        fileName = Dir()
    Loop

    ' Process all files
    processedCount = 0
    For fileCount = 1 To fileList.Count
        filePath = fileList(fileCount)
        Application.StatusBar = "Updating LC: " & fileCount & " of " & fileList.Count & " files..."
        If UpdateLCInWorkbook(filePath, wsLCSource) Then
            processedCount = processedCount + 1
        End If
        ' Allow Excel to process events and stay responsive
        DoEvents
    Next fileCount

    ' End timing
    endTime = Timer
    elapsedTime = endTime - startTime
    
    ' Handle case where timer crossed midnight
    If elapsedTime < 0 Then elapsedTime = elapsedTime + 86400

    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    ' Format time message
    If elapsedTime < 60 Then
        timeMessage = Format(elapsedTime, "0.00") & " seconds"
    Else
        timeMessage = Format(Int(elapsedTime / 60), "0") & " minute(s) " & Format(elapsedTime Mod 60, "0.00") & " seconds"
    End If

    MsgBox "LC successfully updated in template and " & (processedCount - 1) & " collaborator file(s)." & vbCrLf & _
           "Time taken: " & timeMessage, vbInformation, "Update Complete"
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub


Sub Btn_Cleanup_RM()
    Dim cleanupCommand As String
    Dim baseDir As String
    Dim confirmation As VbMsgBoxResult
    Dim exitCode As Long

    ' Clean up empty rows in Gestion_Interfaces sheet
    CleanupGestionInterfaces

    confirmation = MsgBox("Do you want to proceed with cleaning up missing collaborators?" & vbCrLf & _
                          "This will delete interface files for collaborators not in the current list.", _
                          vbYesNo + vbQuestion, "Confirm Cleanup")
    If confirmation = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    If Not CreateCollabsXML(baseDir) Then
        MsgBox "Error creating collabs.xml file. Operation aborted.", vbCritical, "Error"
        GoTo ErrorHandler
    End If

    cleanupCommand = PYTHONEXE & "--basedir " & """" & baseDir & """" & " cleanup"
    Application.StatusBar = "Cleaning up missing collaborator interfaces..."
    exitCode = RunCommand(cleanupCommand)
    Application.StatusBar = False

    If exitCode <> 0 Then
        MsgBox "Error during cleanup. Exit code: " & exitCode, vbCritical, "Error"
        GoTo ErrorHandler
    End If

    MsgBox "Cleanup complete. Missing collaborator interfaces have been deleted.", vbInformation, "Cleanup Complete"

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Sub Btn_Collect_RM_Data()
    Dim pointageCommand As String
    Dim baseDir As String
    Dim xmlPath As String
    Dim ws As Worksheet
    Dim wsLC As Worksheet
    Dim result As Collection
    Dim rowData As Collection
    Dim value As Variant
    Dim r As Long, c As Long
    Dim confirmation As VbMsgBoxResult
    Dim exitCode As Long
    Dim rowsImported As Long
    Dim startRow As Long

    confirmation = MsgBox("Do you want to proceed with importing the pointage data?" & vbCrLf & _
                          "This will import data from RM_Collaborateurs into the SYNTHESE sheet.", _
                          vbYesNo + vbQuestion, "Confirm Import")
    If confirmation = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    ' Check if SYNTHESE sheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_SYNTHESE)
    If Err.Number <> 0 Then
        MsgBox "SYNTHESE sheet not found.", vbCritical, "Error"
        Exit Sub
    End If
    On Error GoTo 0

    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    ' Collect pointage data from existing interfaces
    pointageCommand = PYTHONEXE & "--basedir " & """" & baseDir & """" & " pointage"
    Application.StatusBar = "Exporting pointage data from collaborator files..."
    exitCode = RunCommand(pointageCommand)
    Application.StatusBar = False

    If exitCode <> 0 Then
        MsgBox "Error exporting pointage data. Exit code: " & exitCode, vbCritical, "Error"
        GoTo ErrorHandler
    End If

    ' Verify XML file was created
    xmlPath = baseDir & "\pointage_output.xml"
    If Dir(xmlPath) = "" Then
        MsgBox "Error: pointage_output.xml file was not created.", vbCritical, "Error"
        GoTo ErrorHandler
    End If

    ' Load and import XML data
    Set result = LoadXMLTable(xmlPath)
    If result Is Nothing Then
        MsgBox "Error loading XML data. The file may be corrupted or empty.", vbCritical, "Error"
        GoTo ErrorHandler
    End If

    ' Find starting row for data import
    startRow = ws.Cells(ws.rows.Count, "A").End(xlUp).row + 1
    If startRow < 3 Then startRow = 3
    r = startRow
    rowsImported = 0

    ' Import data into SYNTHESE sheet (maps helper to BA)
    ImportPointageRows ws, result, startRow, rowsImported, 11, 53

    ' Update SYNTHESE columns H and I from LC table matching
    If rowsImported > 0 Then
        Set wsLC = ThisWorkbook.Sheets(SHEET_LC)
        UpdateSyntheseFromLC ws, wsLC, startRow, startRow + rowsImported - 1
    End If

    ' Apply coloring based on K1 totals (helper column after import, using BA)
    ApplySyntheseRowColoring ws, startRow, 11, 53, 35

    ' Clean up temporary XML file
    If Dir(xmlPath) <> "" Then Kill xmlPath

    ' Show import summary
    If rowsImported > 0 Then
        MsgBox "Pointage successfully imported from 'RM_Collaborateurs'." & vbCrLf & _
               rowsImported & " row(s) imported into SYNTHESE sheet.", vbInformation, "Import Complete"
    Else
        MsgBox "No data to import. The pointage file was empty.", vbInformation, "Import Complete"
    End If

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Sub Btn_Collect_Collab_nb_h()
    Dim wsSynth As Worksheet
    Dim wsGI As Worksheet
    Dim wsVerif As Worksheet
    Dim baseDir As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim r As Long
    Dim c As Long
    Dim collabName As String
    Dim weekCode As Variant
    Dim hoursVal As Variant
    Dim collabList As Collection
    Dim collabDict As Object
    Dim weekDict As Object
    Dim sumDict As Object
    Dim key As String
    Dim i As Long
    Dim arrWeeks() As String
    Dim tmp As String
    
    On Error GoTo ErrorHandler
    


    ' Safely get required worksheets
    On Error Resume Next
    Set wsSynth = ThisWorkbook.Sheets(SHEET_SYNTHESE)
    Set wsGI = ThisWorkbook.Sheets(SHEET_GESTION_INTERFACES)
    Set wsVerif = ThisWorkbook.Sheets(SHEET_VERIF_COLLABORATEUR)
    On Error GoTo ErrorHandler

    If wsSynth Is Nothing Then
        MsgBox "SYNTHESE sheet not found.", vbCritical, "Error"
        Exit Sub
    End If
    If wsGI Is Nothing Then
        MsgBox "Gestion_Interfaces sheet not found.", vbCritical, "Error"
        Exit Sub
    End If
    If wsVerif Is Nothing Then
        MsgBox "Vérif_Collaborateur sheet not found.", vbCritical, "Error"
        Exit Sub
    End If

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    Application.ScreenUpdating = False

    ' Collect collaborators:
    '  1) From Gestion_Interfaces (column B, row 3+)
    '  2) From SYNTHESE (column B, data rows), then merge uniquely
    Set collabList = New Collection
    Set collabDict = CreateObject("Scripting.Dictionary")

    ' 1) From Gestion_Interfaces
    lastRow = wsGI.Cells(wsGI.Rows.Count, 2).End(xlUp).Row
    For r = 3 To lastRow
        collabName = Trim(CStr(wsGI.Cells(r, 2).Value))
        If collabName <> "" Then
            If Not collabDict.Exists(collabName) Then
                collabDict.Add collabName, True
                collabList.Add collabName
            End If
        End If
    Next r

    ' 2) From SYNTHESE (column B), merged uniquely with the above
    lastRow = wsSynth.Cells(wsSynth.Rows.Count, SYN_COL_COLLAB).End(xlUp).Row
    If lastRow < SYN_FIRST_DATA_ROW Then
        lastRow = SYN_FIRST_DATA_ROW - 1
    End If
    For r = SYN_FIRST_DATA_ROW To lastRow
        collabName = Trim(CStr(wsSynth.Cells(r, SYN_COL_COLLAB).Value))
        If collabName <> "" Then
            If Not collabDict.Exists(collabName) Then
                collabDict.Add collabName, True
                collabList.Add collabName
            End If
        End If
    Next r

    If collabList.Count = 0 Then
        Application.ScreenUpdating = True
        MsgBox "No collaborators found in Gestion_Interfaces (column B).", vbInformation, "Nothing to Do"
        Exit Sub
    End If

    ' Aggregate hours per (collaborator, week) from SYNTHESE
    Set weekDict = CreateObject("Scripting.Dictionary")
    Set sumDict = CreateObject("Scripting.Dictionary")

    lastRow = wsSynth.Cells(wsSynth.Rows.Count, SYN_COL_HOURS).End(xlUp).Row
    If lastRow < SYN_FIRST_DATA_ROW Then
        lastRow = SYN_FIRST_DATA_ROW - 1
    End If

    For r = SYN_FIRST_DATA_ROW To lastRow
        collabName = Trim(CStr(wsSynth.Cells(r, SYN_COL_COLLAB).Value))
        weekCode = Trim(CStr(wsSynth.Cells(r, SYN_COL_WEEK).Value))
        hoursVal = wsSynth.Cells(r, SYN_COL_HOURS).Value
        
        If collabName <> "" And weekCode <> "" Then
            If collabDict.Exists(collabName) Then
                If IsNumeric(hoursVal) Then
                    If Not weekDict.Exists(weekCode) Then
                        weekDict.Add weekCode, weekCode
                    End If

                    key = collabName & "||" & weekCode
                    If Not sumDict.Exists(key) Then
                        sumDict.Add key, CDbl(hoursVal)
                    Else
                        sumDict(key) = sumDict(key) + CDbl(hoursVal)
                    End If
                End If
            End If
        End If
    Next r

    If weekDict.Count = 0 Then
        ' No SXXYY weeks found; just clear table and write collaborators
        ' Clear existing data in Vérif_Collaborateur (names and values)
        lastRow = wsVerif.Cells(wsVerif.Rows.Count, VERIF_COL_COLLAB).End(xlUp).Row
        lastCol = wsVerif.Cells(VERIF_HEADER_ROW, wsVerif.Columns.Count).End(xlToLeft).Column
        If lastRow >= VERIF_FIRST_COLLAB_ROW Then
            wsVerif.Range(wsVerif.Cells(VERIF_FIRST_COLLAB_ROW, VERIF_COL_COLLAB), _
                          wsVerif.Cells(lastRow, lastCol)).ClearContents
        End If

        ' Write collaborator list only
        For i = 1 To collabList.Count
            wsVerif.Cells(VERIF_FIRST_COLLAB_ROW + i - 1, VERIF_COL_COLLAB).Value = collabList(i)
        Next i

        Application.ScreenUpdating = True
        MsgBox "No SXXYY entries found in SYNTHESE. Collaborator list has been refreshed.", vbInformation, "No Data"
        Exit Sub
    End If

    ' Copy weeks into array for sorting
    ReDim arrWeeks(1 To weekDict.Count)
    i = 1
    For Each weekCode In weekDict.Keys
        arrWeeks(i) = CStr(weekCode)
        i = i + 1
    Next weekCode

    ' Simple bubble sort of week codes (string comparison)
    For i = LBound(arrWeeks) To UBound(arrWeeks) - 1
        For c = i + 1 To UBound(arrWeeks)
            If arrWeeks(c) < arrWeeks(i) Then
                tmp = arrWeeks(i)
                arrWeeks(i) = arrWeeks(c)
                arrWeeks(c) = tmp
            End If
        Next c
    Next i

    ' Clear existing data in Vérif_Collaborateur (names, week headers, and values)
    lastRow = wsVerif.Cells(wsVerif.Rows.Count, VERIF_COL_COLLAB).End(xlUp).Row
    lastCol = wsVerif.Cells(VERIF_HEADER_ROW, wsVerif.Columns.Count).End(xlToLeft).Column
    If lastCol < VERIF_FIRST_WEEK_COL Then lastCol = VERIF_FIRST_WEEK_COL

    ' Clear collaborator names and matrix
    If lastRow >= VERIF_FIRST_COLLAB_ROW Then
        wsVerif.Range(wsVerif.Cells(VERIF_FIRST_COLLAB_ROW, VERIF_COL_COLLAB), _
                      wsVerif.Cells(lastRow, lastCol)).ClearContents
    End If

    ' Clear old week headers (keep "Liste Complète" in column B)
    wsVerif.Range(wsVerif.Cells(VERIF_HEADER_ROW, VERIF_FIRST_WEEK_COL), _
                  wsVerif.Cells(VERIF_HEADER_ROW, lastCol)).ClearContents

    ' Write sorted week headers, copying design from the first header cell
    c = VERIF_FIRST_WEEK_COL
    For i = LBound(arrWeeks) To UBound(arrWeeks)
        wsVerif.Cells(VERIF_HEADER_ROW, c).Value = arrWeeks(i)
        ' Apply same formatting as the template header cell (first SXXYY header)
        wsVerif.Cells(VERIF_HEADER_ROW, VERIF_FIRST_WEEK_COL).Copy
        wsVerif.Cells(VERIF_HEADER_ROW, c).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        c = c + 1
    Next i

    ' Write collaborator names, copying design from the first collaborator cell
    For i = 1 To collabList.Count
        r = VERIF_FIRST_COLLAB_ROW + i - 1
        wsVerif.Cells(r, VERIF_COL_COLLAB).Value = collabList(i)
        ' Apply same formatting as the template collaborator cell (first data row in column B)
        wsVerif.Cells(VERIF_FIRST_COLLAB_ROW, VERIF_COL_COLLAB).Copy
        wsVerif.Cells(r, VERIF_COL_COLLAB).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
    Next i

    ' Build a lookup from week code to column in Vérif_Collaborateur
    Set weekDict = CreateObject("Scripting.Dictionary")
    c = VERIF_FIRST_WEEK_COL
    For i = LBound(arrWeeks) To UBound(arrWeeks)
        weekDict.Add arrWeeks(i), c
        c = c + 1
    Next i

    ' Fill intersection matrix with summed hours
    For i = 1 To collabList.Count
        collabName = collabList(i)
        r = VERIF_FIRST_COLLAB_ROW + i - 1
        
        For Each weekCode In weekDict.Keys
            key = collabName & "||" & CStr(weekCode)
            ' Apply numeric cell formatting based on template data cell (e.g., C5)
            wsVerif.Cells(VERIF_FIRST_COLLAB_ROW, VERIF_FIRST_WEEK_COL).Copy
            With wsVerif.Cells(r, weekDict(weekCode))
                .PasteSpecial xlPasteFormats
                If sumDict.Exists(key) Then
                    .Value = sumDict(key)
                Else
                    ' If no hours for this (collab, week), explicitly set 0
                    .Value = 0
                End If
            End With
            Application.CutCopyMode = False
        Next weekCode
    Next i

    ' Compute and write percentage of non-zero values per week column
    ' Percentage = non_zero_count / (non_zero_count + zero_count)
    lastRow = wsVerif.Cells(wsVerif.Rows.Count, VERIF_COL_COLLAB).End(xlUp).Row
    If lastRow < VERIF_FIRST_COLLAB_ROW Then
        lastRow = VERIF_FIRST_COLLAB_ROW - 1
    End If

    For Each weekCode In weekDict.Keys
        Dim colWeek As Long
        Dim nonZeroCount As Long
        Dim zeroCount As Long
        Dim denom As Long
        Dim pct As Double

        colWeek = CLng(weekDict(weekCode))
        nonZeroCount = 0
        zeroCount = 0

        For r = VERIF_FIRST_COLLAB_ROW To lastRow
            hoursVal = wsVerif.Cells(r, colWeek).Value
            If IsNumeric(hoursVal) Then
                If CDbl(hoursVal) <> 0 Then
                    nonZeroCount = nonZeroCount + 1
                Else
                    zeroCount = zeroCount + 1
                End If
            End If
        Next r

        denom = nonZeroCount + zeroCount
        wsVerif.Cells(PERCENTAGE_NONZEROS_ROW, PERCENTAGE_NONZEROS_COL).Copy
        With wsVerif.Cells(PERCENTAGE_NONZEROS_ROW, colWeek)
            .PasteSpecial xlPasteFormats
            If denom > 0 Then
                pct = nonZeroCount / denom
                .Value = pct
            Else
                .Value = 0
            End If
        End With
        Application.CutCopyMode = False
    Next weekCode

    Application.ScreenUpdating = True
    
    MsgBox "Vérif_Collaborateur table updated with hours by collaborator and SXXYY.", _
           vbInformation, "Update Complete"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error in Collect_Collab_nb_h: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Unexpected Error"
End Sub

Sub Btn_Reset_Verif_Collaborateur()
    Dim wsVerif As Worksheet
    Dim baseDir As String
    Dim archiveConfirm As VbMsgBoxResult
    Dim lastRow As Long
    Dim lastCol As Long
    Dim archivePath As String
    Dim timestamp As String
    Dim archivedFolder As String


    ' Confirm with user: this action will delete the records
    archiveConfirm = MsgBox("This action will delete all collaborators, weeks, and hours data from Vérif_Collaborateur." & vbCrLf & vbCrLf & _
                           "Do you want to archive the current data before resetting?" & vbCrLf & _
                           "(A copy will be saved in the Archived folder)", _
                           vbYesNoCancel + vbQuestion, "Confirm Reset")
    If archiveConfirm = vbCancel Then Exit Sub

    On Error Resume Next
    Set wsVerif = ThisWorkbook.Sheets(SHEET_VERIF_COLLABORATEUR)
    On Error GoTo 0

    If wsVerif Is Nothing Then
        MsgBox "Vérif_Collaborateur sheet not found.", vbCritical, "Error"
        Exit Sub
    End If

    ' If user chose to archive, get base dir and create archive file
    If archiveConfirm = vbYes Then
        baseDir = GetBaseDir()
        If baseDir = "" Then Exit Sub

        archivedFolder = baseDir & "\Archived"
        timestamp = Format(Now, "ddmmyyyy_HHMMSS")
        archivePath = archivedFolder & "\Vérif_Collaborateur_" & timestamp & ".xlsx"

        Application.ScreenUpdating = False
        Application.StatusBar = "Creating archive file..."

        If Not ArchiveSingleSheet(wsVerif, archivePath, True, SHEET_VERIF_COLLABORATEUR) Then
            Application.StatusBar = False
            Application.ScreenUpdating = True
            Exit Sub
        End If

        Application.StatusBar = False
    End If

    Application.ScreenUpdating = False

    lastRow = wsVerif.Cells(wsVerif.Rows.Count, VERIF_COL_COLLAB).End(xlUp).Row
    lastCol = wsVerif.Cells(VERIF_HEADER_ROW, wsVerif.Columns.Count).End(xlToLeft).Column

    ' Clear percentage row (row 4) across all week columns
    If lastCol >= VERIF_FIRST_WEEK_COL Then
        wsVerif.Range(wsVerif.Cells(PERCENTAGE_NONZEROS_ROW, VERIF_FIRST_WEEK_COL), _
                      wsVerif.Cells(PERCENTAGE_NONZEROS_ROW, lastCol)).ClearContents
    End If

    ' Delete extra rows (from row 7 onwards)
    If lastRow >= VERIF_FIRST_COLLAB_ROW + 1 Then
        wsVerif.Rows((VERIF_FIRST_COLLAB_ROW + 1) & ":" & lastRow).Delete Shift:=xlUp
    End If

    ' Delete extra columns (from column D onwards)
    If lastCol >= VERIF_FIRST_WEEK_COL + 1 Then
        wsVerif.Columns(VERIF_FIRST_WEEK_COL + 1).Resize(, lastCol - VERIF_FIRST_WEEK_COL).Delete Shift:=xlToLeft
    End If

    ' Clear data cells and reset to base template
    wsVerif.Cells(VERIF_FIRST_COLLAB_ROW, VERIF_COL_COLLAB).ClearContents
    wsVerif.Cells(VERIF_FIRST_COLLAB_ROW, VERIF_FIRST_WEEK_COL).ClearContents

    ' Reset first week header to placeholder "SXXYY"
    wsVerif.Cells(VERIF_HEADER_ROW, VERIF_FIRST_WEEK_COL).Value = "SXXYY"

    Application.ScreenUpdating = True

    If archiveConfirm = vbYes Then
        MsgBox "Vérif_Collaborateur archived and reset to base template." & vbCrLf & _
               "Archive saved to: " & archivePath, vbInformation, "Reset Complete"
    Else
        MsgBox "Vérif_Collaborateur reset to base template.", vbInformation, "Reset Complete"
    End If
    Exit Sub
End Sub

Sub Btn_Extract_LC_MSP()
    Dim wsLC As Worksheet
    Dim wsSrc As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim lookupRow As Long
    Dim valN As Variant
    Dim confirmation As VbMsgBoxResult
    Dim baseDir As String
    Dim firstLookupDataRow As Long
    Dim lastLookupRow As Long


    confirmation = MsgBox("Do you want to regenerate the LC lookup table from Extract_MSP?" & vbCrLf & _
                          "This will overwrite existing values in LC (columns F to K starting from row 3).", _
                          vbYesNo + vbQuestion, "Confirm LC Generation")
    If confirmation = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    ' Get LC and Extract_MSP sheets
    On Error Resume Next
    Set wsLC = ThisWorkbook.Sheets(SHEET_LC)
    Set wsSrc = ThisWorkbook.Sheets(SHEET_EXTRACT_MSP)
    On Error GoTo ErrorHandler

    If wsLC Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "LC sheet not found.", vbCritical, "Error"
        Exit Sub
    End If

    If wsSrc Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "Extract_MSP sheet not found.", vbCritical, "Error"
        Exit Sub
    End If

    ' Determine last used row in Extract_MSP (based on column B)
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "B").End(xlUp).Row
    If lastRow < 3 Then
        Application.ScreenUpdating = True
        MsgBox "No data found in Extract_MSP to build LC lookup table.", vbInformation, "Nothing to Do"
        Exit Sub
    End If

    ' Clear existing lookup area F:K from row 3 down (keep row 2 as template/header)
    firstLookupDataRow = LC_LOOKUP_FIRST_ROW + 1     ' typically row 3
    lastLookupRow = wsLC.Cells(wsLC.Rows.Count, LC_LOOKUP_COL_F).End(xlUp).Row
    If lastLookupRow < firstLookupDataRow Then lastLookupRow = firstLookupDataRow

    wsLC.Range(wsLC.Cells(firstLookupDataRow, LC_LOOKUP_COL_F), _
               wsLC.Cells(lastLookupRow, LC_LOOKUP_COL_K)).ClearContents

    ' Reshape Extract_MSP into LC lookup table (no uniqueness, one-to-one row copy)
    lookupRow = firstLookupDataRow
    For r = 2 To lastRow
        ' Skip rows where key source column (B) is empty
        If Trim$(CStr(wsSrc.Cells(r, "B").Value)) <> "" Then
            ' Column F in LC (lookup): from Extract_MSP column B
            wsLC.Cells(lookupRow, LC_LOOKUP_COL_F).Value = wsSrc.Cells(r, "B").Value

            ' Column G in LC (lookup): from Extract_MSP column F
            wsLC.Cells(lookupRow, LC_LOOKUP_COL_G).Value = wsSrc.Cells(r, "F").Value

            ' Column H in LC (lookup): from Extract_MSP column N, blank if 0
            valN = wsSrc.Cells(r, "N").Value
            If IsNumeric(valN) And CDbl(valN) = 0 Then
                wsLC.Cells(lookupRow, LC_LOOKUP_COL_H).ClearContents
            Else
                wsLC.Cells(lookupRow, LC_LOOKUP_COL_H).Value = valN
            End If

            ' Column I in LC (lookup): from Extract_MSP column O
            wsLC.Cells(lookupRow, LC_LOOKUP_COL_I).Value = wsSrc.Cells(r, "O").Value

            ' Column J in LC (lookup): from Extract_MSP column C
            wsLC.Cells(lookupRow, LC_LOOKUP_COL_J).Value = wsSrc.Cells(r, "C").Value

            ' Column K in LC (lookup): from Extract_MSP column U
            wsLC.Cells(lookupRow, LC_LOOKUP_COL_K).Value = wsSrc.Cells(r, "U").Value

            lookupRow = lookupRow + 1
        End If
    Next r

    Application.ScreenUpdating = True
    MsgBox "LC table successfully generated from Extract_MSP.", vbInformation, "Update Complete"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error in Extract_LC_MSP: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Unexpected Error"
End Sub

Sub Btn_Reset_LC()
    Dim wsLC As Worksheet
    Dim archiveConfirm As VbMsgBoxResult
    Dim baseDir As String
    Dim archivePath As String
    Dim timestamp As String
    Dim archivedFolder As String
    Dim firstLookupDataRow As Long
    Dim lastLookupRow As Long


    archiveConfirm = MsgBox("This will clear the LC lookup table (columns F to K starting from row 3)." & vbCrLf & vbCrLf & _
                            "Do you want to ARCHIVE the current LC table before clearing it?" & vbCrLf & _
                            "(A copy will be saved in the Archived folder)", _
                            vbYesNoCancel + vbQuestion, "Confirm LC Reset")
    If archiveConfirm = vbCancel Then Exit Sub
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    ' Get LC sheet safely
    On Error Resume Next
    Set wsLC = ThisWorkbook.Sheets(SHEET_LC)
    On Error GoTo ErrorHandler
    
    If wsLC Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "LC sheet not found.", vbCritical, "Error"
        Exit Sub
    End If

    ' Optionally archive current LC table before clearing
    If archiveConfirm = vbYes Then
        baseDir = GetBaseDir()
        If baseDir = "" Then
            Application.ScreenUpdating = True
            Exit Sub
        End If

        archivedFolder = baseDir & "\Archived"
        timestamp = Format(Now, "ddmmyyyy_HHMMSS")
        archivePath = archivedFolder & "\LC_" & timestamp & ".xlsx"

        Application.StatusBar = "Creating LC archive..."
        
        If Not ArchiveSingleSheet(wsLC, archivePath, True, SHEET_LC) Then
            Application.ScreenUpdating = True
            Application.StatusBar = False
            Exit Sub
        End If

        Application.StatusBar = False
    End If

    ' Clear LC lookup table F:K from row 3 down to last used row (do NOT touch B:D)
    firstLookupDataRow = LC_LOOKUP_FIRST_ROW + 1      ' typically row 3
    lastLookupRow = wsLC.Cells(wsLC.Rows.Count, LC_LOOKUP_COL_F).End(xlUp).Row
    If lastLookupRow < firstLookupDataRow Then lastLookupRow = firstLookupDataRow

    wsLC.Range(wsLC.Cells(firstLookupDataRow, LC_LOOKUP_COL_F), _
               wsLC.Cells(lastLookupRow, LC_LOOKUP_COL_K)).ClearContents

    Application.ScreenUpdating = True
    MsgBox "LC table has been cleared.", vbInformation, "Reset Complete"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error in Btn_Reset_LC: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Unexpected Error"
End Sub

Sub Btn_Collect_FS_Data()
    Dim wsFS As Worksheet
    Dim wsLC As Worksheet
    Dim wsSynth As Worksheet
    Dim lastLcRow As Long
    Dim lastSynthRow As Long
    Dim r As Long
    Dim outRow As Long
    Dim baseDir As String
    Dim key As String
    Dim sumDict As Object
    Dim lcSumDict As Object
    Dim sumDict2 As Object
    Dim lcSumDict2 As Object
    Dim arrKeys As Variant
    Dim valG_LC As Variant
    Dim parts As Variant
    Dim i As Long
    Dim valF As Variant
    Dim valJ As Variant
    Dim valI As Variant
    Dim valE As String
    Dim part1 As String
    Dim posSprint As Long
    Dim lastDataRow As Long

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    ' Safely get required worksheets
    On Error Resume Next
    Set wsFS = ThisWorkbook.Sheets(SHEET_FICHIER_SYNTHESE)
    Set wsLC = ThisWorkbook.Sheets(SHEET_LC)
    Set wsSynth = ThisWorkbook.Sheets(SHEET_SYNTHESE)
    On Error GoTo ErrorHandler

    If wsFS Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "Fichier de synthèse sheet not found.", vbCritical, "Error"
        Exit Sub
    End If
    If wsLC Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "LC sheet not found.", vbCritical, "Error"
        Exit Sub
    End If
    If wsSynth Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "SYNTHESE sheet not found.", vbCritical, "Error"
        Exit Sub
    End If

    ' Build sum of SYNTHÈSE column I by (LC F, LC J) combination.
    ' LC F = first part of SYNTHESE column E (split by "Sprint")
    ' LC J = SYNTHESE column G
    Set sumDict = CreateObject("Scripting.Dictionary")
    sumDict.CompareMode = 1 ' vbTextCompare

    lastSynthRow = wsSynth.Cells(wsSynth.Rows.Count, SYN_COL_I).End(xlUp).Row
    If lastSynthRow < SYN_FIRST_DATA_ROW Then lastSynthRow = SYN_FIRST_DATA_ROW - 1

    For r = SYN_FIRST_DATA_ROW To lastSynthRow
        valI = wsSynth.Cells(r, SYN_COL_I).Value
        
        ' Split SYNTHESE column E by "Sprint" separator and use first part
        valE = Trim$(CStr(wsSynth.Cells(r, SYN_COL_E).Value))
        posSprint = InStr(1, valE, "Sprint", vbTextCompare)
        
        If posSprint > 0 Then
            part1 = Trim$(Left$(valE, posSprint - 1))  ' First part = LC F
            valJ = Trim$(CStr(wsSynth.Cells(r, SYN_COL_G).Value))  ' Column G = LC J
            
            If part1 <> "" Or valJ <> "" Then
                If IsNumeric(valI) Then
                    key = part1 & "||" & valJ  ' Key = (LC F, LC J)
                    If sumDict.Exists(key) Then
                        sumDict(key) = sumDict(key) + CDbl(valI)
                    Else
                        sumDict.Add key, CDbl(valI)
                    End If
                End If
            End If
        End If
    Next r

    ' Now build the first FS table (E6:J...) from LC combinations (F,J),
    ' using:
    '  - G = sum of LC column I for each (F,J) combination
    '  - H = sum of SYNTHESE column I for the same combination
    Set lcSumDict = CreateObject("Scripting.Dictionary")
    lcSumDict.CompareMode = 1 ' vbTextCompare

    lastLcRow = wsLC.Cells(wsLC.Rows.Count, LC_LOOKUP_COL_F).End(xlUp).Row
    If lastLcRow < LC_FIRST_ROW Then lastLcRow = LC_FIRST_ROW

    ' Clear existing data for first table E6:J (keep row 5 as header)
    If wsFS.Rows.Count >= 6 Then
        wsFS.Range("E6:J" & wsFS.Rows.Count).ClearContents
    End If

    outRow = 6

    ' First, aggregate LC column I by (F,J) combination
    ' LC lookup table starts from row 3
    For r = LC_FIRST_ROW To lastLcRow
        valF = Trim$(CStr(wsLC.Cells(r, LC_LOOKUP_COL_F).Value))
        valJ = Trim$(CStr(wsLC.Cells(r, LC_LOOKUP_COL_J).Value))
        valI = wsLC.Cells(r, LC_LOOKUP_COL_I).Value

        If valF <> "" Or valJ <> "" Then
            key = valF & "||" & valJ
            If IsNumeric(valI) Then
                If lcSumDict.Exists(key) Then
                    lcSumDict(key) = lcSumDict(key) + CDbl(valI)
                Else
                    lcSumDict.Add key, CDbl(valI)
                End If
            End If
        End If
    Next r

    ' Then, write one row per distinct (F,J) combination into FS
    Dim valG As Double
    Dim valH As Double
    Dim pctOverrun As Double
    Dim diffHours As Double

    If lcSumDict.Count > 0 Then
        arrKeys = lcSumDict.Keys
        For i = LBound(arrKeys) To UBound(arrKeys)
            key = CStr(arrKeys(i))
            parts = Split(key, "||")
            If UBound(parts) >= 1 Then
                valF = parts(0)
                valJ = parts(1)
            Else
                valF = ""
                valJ = ""
            End If

            wsFS.Cells(outRow, "E").Value = valF               ' LC F
            wsFS.Cells(outRow, "F").Value = valJ               ' LC J
            valG = lcSumDict(key)                               ' sum of LC I
            wsFS.Cells(outRow, "G").Value = valG

            If sumDict.Exists(key) Then
                valH = sumDict(key)                            ' summed SYNTHESE I
            Else
                valH = 0
            End If
            wsFS.Cells(outRow, "H").Value = valH

            ' Calculate column I: % de dépassement = (H - G) / G * 100, rounded to integer
            If valG <> 0 Then
                pctOverrun = ((valH - valG) / valG) * 100
                wsFS.Cells(outRow, "I").Value = Round(pctOverrun, 0)
            Else
                wsFS.Cells(outRow, "I").ClearContents
            End If

            ' Calculate column J: Ecart en h = G - H (inverse)
            diffHours = valG - valH
            wsFS.Cells(outRow, "J").Value = diffHours

            ' Apply formatting from first data row (row 6) to current row
            If outRow > 6 Then
                wsFS.Range("E6:J6").Copy
                wsFS.Range("E" & outRow & ":J" & outRow).PasteSpecial xlPasteFormats
                Application.CutCopyMode = False
            End If

            outRow = outRow + 1
        Next i
    End If

    ' -------------------------------------------------------------------------
    ' Second table (L6:Q): same logic but key = (LC G, SYNTHESE F)
    ' L=StrS (LC G), M=Livrable (SYNTHESE F), N=Somme temps prévu (sum LC I),
    ' O=Somme temps consommé (sum SYNTHESE I), P=% de dépassement, Q=Ecart en h
    ' -------------------------------------------------------------------------
    
    ' Build sum of SYNTHESE column I by (LC G, SYNTHESE F) = SYNTHESE F (LC G = SYNTHESE F in lookup)
    Set sumDict2 = CreateObject("Scripting.Dictionary")
    sumDict2.CompareMode = 1
    
    For r = SYN_FIRST_DATA_ROW To lastSynthRow
        valI = wsSynth.Cells(r, SYN_COL_I).Value
        valG_LC = Trim$(CStr(wsSynth.Cells(r, SYN_COL_F).Value))  ' SYNTHESE F = LC G
        
        If valG_LC <> "" Then
            If IsNumeric(valI) Then
                key = valG_LC
                If sumDict2.Exists(key) Then
                    sumDict2(key) = sumDict2(key) + CDbl(valI)
                Else
                    sumDict2.Add key, CDbl(valI)
                End If
            End If
        End If
    Next r
    
    ' Aggregate LC column I by LC G
    Set lcSumDict2 = CreateObject("Scripting.Dictionary")
    lcSumDict2.CompareMode = 1
    
    For r = LC_FIRST_ROW To lastLcRow
        valG_LC = Trim$(CStr(wsLC.Cells(r, LC_LOOKUP_COL_G).Value))
        valI = wsLC.Cells(r, LC_LOOKUP_COL_I).Value
        
        If valG_LC <> "" Then
            If IsNumeric(valI) Then
                If lcSumDict2.Exists(valG_LC) Then
                    lcSumDict2(valG_LC) = lcSumDict2(valG_LC) + CDbl(valI)
                Else
                    lcSumDict2.Add valG_LC, CDbl(valI)
                End If
            End If
        End If
    Next r
    
    ' Clear and write second table L6:Q
    If wsFS.Rows.Count >= 6 Then
        wsFS.Range("L6:Q" & wsFS.Rows.Count).ClearContents
    End If
    
    outRow = 6
    If lcSumDict2.Count > 0 Then
        arrKeys = lcSumDict2.Keys
        For i = LBound(arrKeys) To UBound(arrKeys)
            key = CStr(arrKeys(i))
            valG = lcSumDict2(key)
            If sumDict2.Exists(key) Then
                valH = sumDict2(key)
            Else
                valH = 0
            End If
            
            wsFS.Cells(outRow, "L").Value = key               ' LC G (StrS)
            wsFS.Cells(outRow, "M").Value = key               ' SYNTHESE F (Livrable)
            wsFS.Cells(outRow, "N").Value = valG              ' sum of LC I
            wsFS.Cells(outRow, "O").Value = valH              ' sum of SYNTHESE I
            
            If valG <> 0 Then
                pctOverrun = ((valH - valG) / valG) * 100
                wsFS.Cells(outRow, "P").Value = Round(pctOverrun, 0)
            Else
                wsFS.Cells(outRow, "P").ClearContents
            End If
            
            diffHours = valG - valH
            wsFS.Cells(outRow, "Q").Value = diffHours
            
            If outRow > 6 Then
                wsFS.Range("L6:Q6").Copy
                wsFS.Range("L" & outRow & ":Q" & outRow).PasteSpecial xlPasteFormats
                Application.CutCopyMode = False
            End If
            
            outRow = outRow + 1
        Next i
    End If

    Application.ScreenUpdating = True
    MsgBox "Fichier de synthèse tables (E6:J and L6:Q) updated from LC and SYNTHESE.", vbInformation, "Update Complete"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error in Btn_Collect_FS_Data: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Unexpected Error"
End Sub

Sub Btn_Reset_FS()
    Dim wsFS As Worksheet
    Dim archiveConfirm As VbMsgBoxResult
    Dim baseDir As String
    Dim archivedFolder As String
    Dim archivePath As String
    Dim timestamp As String
    Dim lastDataRow As Long

    ' Confirm with user and optionally archive the Fichier de synthèse sheet
    archiveConfirm = MsgBox("This will clear the generated tables in 'Fichier de synthèse' (E6:J and L6:Q)." & vbCrLf & vbCrLf & _
                           "Do you want to SAVE/ARCHIVE the current sheet before resetting?" & vbCrLf & _
                           "(A copy will be saved in the Archived folder).", _
                           vbYesNoCancel + vbQuestion, "Confirm FS Reset")
    If archiveConfirm = vbCancel Then Exit Sub

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    On Error Resume Next
    Set wsFS = ThisWorkbook.Sheets(SHEET_FICHIER_SYNTHESE)
    On Error GoTo ErrorHandler

    If wsFS Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "Fichier de synthèse sheet not found.", vbCritical, "Error"
        Exit Sub
    End If

    If archiveConfirm = vbYes Then
        baseDir = GetBaseDir()
        If baseDir = "" Then
            Application.ScreenUpdating = True
            Exit Sub
        End If

        archivedFolder = baseDir & "\Archived"
        timestamp = Format(Now, "ddmmyyyy_HHMMSS")
        archivePath = archivedFolder & "\Fichier_de_synthese_" & timestamp & ".xlsx"

        Application.StatusBar = "Archiving Fichier de synthèse..."

        If Not ArchiveSingleSheet(wsFS, archivePath, True, SHEET_FICHIER_SYNTHESE) Then
            Application.ScreenUpdating = True
            Application.StatusBar = False
            Exit Sub
        End If

        Application.StatusBar = False
    End If

    ' Clear both FS tables:
    ' - Row 6: clear contents only (keep format)
    ' - Rows 7+: clear contents and formats (remove format of all added rows)

    ' First table E6:J...
    wsFS.Range("E6:J6").ClearContents
    lastDataRow = wsFS.Cells(wsFS.Rows.Count, "E").End(xlUp).Row
    If lastDataRow < 7 Then lastDataRow = 7
    wsFS.Range("E7:J" & lastDataRow).Clear

    ' Second table L6:Q...
    wsFS.Range("L6:Q6").ClearContents
    lastDataRow = wsFS.Cells(wsFS.Rows.Count, "L").End(xlUp).Row
    If lastDataRow < 7 Then lastDataRow = 7
    wsFS.Range("L7:Q" & lastDataRow).Clear

    Application.ScreenUpdating = True
    MsgBox "Fichier de synthèse tables have been cleared.", vbInformation, "Reset Complete"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error in Btn_Reset_FS: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Unexpected Error"
End Sub
