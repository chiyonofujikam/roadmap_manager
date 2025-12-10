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
    Set ws = ThisWorkbook.Sheets("SYNTHESE")
    If Err.Number <> 0 Then
        MsgBox "SYNTHESE sheet not found.", vbCritical, "Error"
        Exit Sub
    End If
    On Error GoTo 0

    ' Check if LC sheet exists
    On Error Resume Next
    Set wsLC = ThisWorkbook.Sheets("LC")
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
    Set newWs = newWb.Sheets("SYNTHESE")

    ' Copy LC sheet with all formatting to new workbook
    wsLC.Copy After:=newWb.Sheets("SYNTHESE")
    Set newWsLC = newWb.Sheets("LC")

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
        If sht.Name <> "SYNTHESE" And sht.Name <> "LC" Then
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

    ' Clear SYNTHESE sheet if it had data
    If hasData Then
        ws.rows("3:" & lastRow).ClearContents
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

Sub Btn_Collect_RM_Data()
    Dim pointageCommand As String
    Dim deleteCommand As String
    Dim createCommand As String
    Dim baseDir As String
    Dim xmlPath As String
    Dim ws As Worksheet
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
    Set ws = ThisWorkbook.Sheets("SYNTHESE")
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

    ' Import data into SYNTHESE sheet
    For Each rowData In result
        c = 1
        For Each value In rowData
            ws.Cells(r, c).value = value
            c = c + 1
        Next value
        r = r + 1
        rowsImported = rowsImported + 1
    Next rowData

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
    Dim updateCommand As String
    Dim baseDir As String
    Dim confirmation As VbMsgBoxResult
    Dim exitCode As Long

    confirmation = MsgBox("Do you want to proceed with updating the conditional lists (LC)?" & vbCrLf & _
                          "This will update LC in the template and all collaborator files.", _
                          vbYesNo + vbQuestion, "Confirm Update")
    If confirmation = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    ' Create LC.xlsx file from LC sheet
    If Not CreateLCExcel(baseDir) Then
        MsgBox "Error creating LC.xlsx file. Operation aborted.", vbCritical, "Error"
        GoTo ErrorHandler
    End If

    updateCommand = PYTHONEXE & "--basedir " & """" & baseDir & """" & " update"
    Application.StatusBar = "Updating conditional lists (LC) in all files..."
    exitCode = RunCommand(updateCommand)
    Application.StatusBar = False

    If exitCode <> 0 Then
        MsgBox "Error updating LC. Exit code: " & exitCode, vbCritical, "Error"
        GoTo ErrorHandler
    End If

    MsgBox "LC successfully updated in template and all collaborator files.", vbInformation, "Update Complete"

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
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
