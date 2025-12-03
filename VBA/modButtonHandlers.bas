Option Explicit

' =============================================================================
' BUTTON HANDLERS
' All button click event handlers for the Excel interface
' =============================================================================

Sub Btn_Create_RM()
    Dim command As String
    Dim baseDir As String

    ' Clean up empty rows in Gestion_Interfaces sheet
    CleanupGestionInterfaces

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    ' Create collabs.xml file
    If Not CreateCollabsXML(baseDir) Then Exit Sub

    command = PYTHONEXE & "--basedir " & """" & baseDir & """" & " create --archive"
    Application.StatusBar = "Creating collaborator interfaces..."
    RunCommand command
    Application.StatusBar = False

    MsgBox "Collaborator interfaces successfully created and archived.", vbInformation, "Creation Complete"
End Sub

Sub Btn_Delete_RM()
    Dim command As String
    Dim baseDir As String
    Dim forceDelete As VbMsgBoxResult
    Dim archiveChoice As VbMsgBoxResult

    forceDelete = MsgBox("Do you want to FORCE deletion of RM Interfaces?" & vbCrLf & _
                         "(This will delete all generated interfaces)", _
                         vbYesNo + vbQuestion, "Confirm Force Deletion")
    If forceDelete = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    archiveChoice = MsgBox("Do you want to ARCHIVE deleted interfaces?", _
                           vbYesNo + vbQuestion, "Archive Confirmation")

    command = PYTHONEXE & "--basedir " & """" & baseDir & """" & " delete --force"

    If archiveChoice = vbYes Then
        command = command & " --archive"
        Application.StatusBar = "Archiving and deleting interfaces..."
    Else
        Application.StatusBar = "Deleting interfaces..."
    End If

    RunCommand command
    Application.StatusBar = False

    If archiveChoice = vbYes Then
        MsgBox "Interfaces successfully archived and deleted.", vbInformation, "Deletion Complete"
    Else
        MsgBox "Interfaces successfully deleted.", vbInformation, "Deletion Complete"
    End If
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

    ' Determine if there's data to archive
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).Row
    hasData = (lastRow >= 3)

    ' Create Archived folder path
    archivedFolder = baseDir & "\Archived"
    If Dir(archivedFolder, vbDirectory) = "" Then
        MkDir archivedFolder
    End If

    ' Generate timestamp in format: ddmmyyyy_HHMMSS
    timestamp = Format(Now, "ddmmyyyy_HHMMSS")
    archivePath = archivedFolder & "\Archive_SYNTHESE_" & timestamp & ".xlsx"

    Application.StatusBar = "Creating archive file with SYNTHESE and LC sheets..."
    Application.ScreenUpdating = False

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
    Dim i As Long

    ' Remove shapes from SYNTHESE sheet
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
    newWb.SaveAs archivePath, xlOpenXMLWorkbook
    Application.DisplayAlerts = True

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
End Sub

Sub Btn_Collect_RM_Data()
    Dim command As String
    Dim baseDir As String
    Dim xmlPath As String
    Dim ws As Worksheet
    Dim result As Collection
    Dim rowData As Collection
    Dim value As Variant
    Dim r As Long, c As Long
    Dim confirmation As VbMsgBoxResult

    confirmation = MsgBox("Do you want to proceed with importing the pointage data?" & vbCrLf & _
                          "This will import data from RM_Collaborateurs into the SYNTHESE sheet.", _
                          vbYesNo + vbQuestion, "Confirm Import")
    If confirmation = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    command = PYTHONEXE & "--basedir " & """" & baseDir & """" & " pointage"
    Application.StatusBar = "Exporting pointage data from collaborator files..."
    RunCommand command
    Application.StatusBar = False

    Set ws = ThisWorkbook.Sheets("SYNTHESE")
    r = ws.Cells(ws.rows.Count, "A").End(xlUp).Row + 1
    If r < 3 Then r = 3

    xmlPath = baseDir & "\pointage_output.xml"
    Set result = LoadXMLTable(xmlPath)

    For Each rowData In result
        c = 1
        For Each value In rowData
            ws.Cells(r, c).value = value
            c = c + 1
        Next value
        r = r + 1
    Next rowData

    If Dir(xmlPath) <> "" Then Kill xmlPath

    MsgBox "Pointage successfully imported from 'RM_Collaborateurs'", vbInformation, "Import Complete"
End Sub

Sub Btn_Update_LC()
    Dim command As String
    Dim baseDir As String
    Dim confirmation As VbMsgBoxResult

    confirmation = MsgBox("Do you want to proceed with updating the conditional lists (LC)?" & vbCrLf & _
                          "This will update LC in the template and all collaborator files.", _
                          vbYesNo + vbQuestion, "Confirm Update")
    If confirmation = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    ' Create LC.xlsx file from LC sheet
    If Not CreateLCExcel(baseDir) Then Exit Sub

    command = PYTHONEXE & "--basedir " & """" & baseDir & """" & " update"
    Application.StatusBar = "Updating conditional lists (LC) in all files..."
    RunCommand command
    Application.StatusBar = False

    MsgBox "LC successfully updated in template and all collaborator files.", vbInformation, "Update Complete"
End Sub

Sub Btn_Cleanup_RM()
    Dim command As String
    Dim baseDir As String
    Dim confirmation As VbMsgBoxResult

    ' Clean up empty rows in Gestion_Interfaces sheet
    CleanupGestionInterfaces

    confirmation = MsgBox("Do you want to proceed with cleaning up missing collaborators?" & vbCrLf & _
                          "This will delete interface files for collaborators not in the current list.", _
                          vbYesNo + vbQuestion, "Confirm Cleanup")
    If confirmation = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    ' Create collabs.xml file
    If Not CreateCollabsXML(baseDir) Then Exit Sub

    command = PYTHONEXE & "--basedir " & """" & baseDir & """" & " cleanup"
    Application.StatusBar = "Cleaning up missing collaborator interfaces..."
    RunCommand command
    Application.StatusBar = False

    MsgBox "Cleanup complete. Missing collaborator interfaces have been deleted.", vbInformation, "Cleanup Complete"
End Sub
