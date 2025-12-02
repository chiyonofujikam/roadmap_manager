Option Explicit

' Global variables
Public GLOBAL_BASEDIR As String
Public Const PYTHONEXE1 As String = """" & "C:\Users\MustaphaELKAMILI\AppData\Local\Programs\Python\Python314\Scripts\roadmap.exe" & """" & " "
Public PYTHONEXE As String

' =============================================================================
' UTILITY FUNCTIONS
' =============================================================================

Function RunCommand(cmd As String) As Long
    Dim shellObj As Object
    Set shellObj = CreateObject("WScript.Shell")
    RunCommand = shellObj.Run(cmd, 1, True)
End Function

Function GetBaseDir() As String
    Dim f As FileDialog

    ' Return cached value if already set
    If GLOBAL_BASEDIR <> "" Then
        If PYTHONEXE = "" Then
            PYTHONEXE = """" & GLOBAL_BASEDIR & "\Scripts\roadmap.exe" & """" & " "
        End If
        GetBaseDir = GLOBAL_BASEDIR
        Exit Function
    End If

    ' Prompt user to select folder
    MsgBox "Please select the base directory"
    Set f = Application.FileDialog(msoFileDialogFolderPicker)

    If f.Show <> -1 Then
        MsgBox "No folder selected.", vbExclamation
        Exit Function
    End If

    GLOBAL_BASEDIR = f.SelectedItems(1)
    PYTHONEXE = """" & GLOBAL_BASEDIR & "\Scripts\roadmap.exe" & """" & " "
    GetBaseDir = GLOBAL_BASEDIR
End Function

Function LoadXMLTable(filePath As String) As Collection
    Dim xml As Object
    Dim rows As Object, rowNode As Object, childNode As Object
    Dim table As New Collection
    Dim oneRow As Collection

    Set xml = CreateObject("MSXML2.DOMDocument")
    xml.async = False
    xml.Load filePath

    If xml.parseError.ErrorCode <> 0 Then
        MsgBox "XML parse error: " & xml.parseError.reason, vbCritical
        Exit Function
    End If

    Set rows = xml.SelectNodes("//row")
    For Each rowNode In rows
        Set oneRow = New Collection
        For Each childNode In rowNode.ChildNodes
            oneRow.Add childNode.text
        Next childNode
        table.Add oneRow
    Next rowNode

    Set LoadXMLTable = table
End Function

Function EscapeXML(text As String) As String
    Dim result As String
    result = text
    result = Replace(result, "&", "&amp;")
    result = Replace(result, "<", "&lt;")
    result = Replace(result, ">", "&gt;")
    result = Replace(result, """", "&quot;")
    EscapeXML = result
End Function

' =============================================================================
' BUTTON HANDLERS
' =============================================================================

Sub Btn_Create_RM()
    Dim command As String
    Dim baseDir As String

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

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
    Dim command As String
    Dim baseDir As String
    Dim archiveConfirm As VbMsgBoxResult
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim xmlPath As String
    Dim fileNum As Integer
    Dim r As Long, c As Long
    Dim cellValue As String

    archiveConfirm = MsgBox("Do you want to proceed with archiving the SYNTHESE sheet?" & vbCrLf & _
                            "A new archive file will be created.", _
                            vbYesNo + vbQuestion, "Confirm Archiving")
    If archiveConfirm = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    Set ws = ThisWorkbook.Sheets("SYNTHESE")
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column

    ' Save XML backup before clearing
    If lastRow >= 3 Then
        xmlPath = baseDir & "\Archived\SYNTHESE_Archive_backup_" & Format(Now, "yyyymmdd_hhmmss") & ".xml"
        fileNum = FreeFile
        Open xmlPath For Output As #fileNum

        Print #fileNum, "<?xml version=""1.0"" encoding=""UTF-8""?>"
        Print #fileNum, "<table>"
        For r = 3 To lastRow
            Print #fileNum, "  <row>"
            For c = 1 To lastCol
                cellValue = EscapeXML(CStr(ws.Cells(r, c).value))
                Print #fileNum, "    <col" & c & ">" & cellValue & "</col" & c & ">"
            Next c
            Print #fileNum, "  </row>"
        Next r
        Print #fileNum, "</table>"

        Close #fileNum
    End If

    command = PYTHONEXE & "--basedir " & """" & baseDir & """" & " pointage --delete"
    Application.StatusBar = "Archiving SYNTHESE sheet..."
    RunCommand command
    Application.StatusBar = False

    If lastRow >= 3 Then
        ws.rows("3:" & lastRow).ClearContents
        MsgBox "SYNTHESE sheet successfully archived and cleared." & vbCrLf & _
               "XML backup saved to: " & xmlPath, vbInformation, "Archive Complete"
    Else
        MsgBox "SYNTHESE sheet was already empty. Nothing to archive.", vbInformation, "Archive Complete"
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

    command = PYTHONEXE & "--basedir " & """" & baseDir & """" & " update"
    Application.StatusBar = "Updating conditional lists (LC) in all files..."
    RunCommand command
    Application.StatusBar = False

    MsgBox "LC successfully updated in template and all collaborator files.", vbInformation, "Update Complete"
End Sub
