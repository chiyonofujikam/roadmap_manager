Option Explicit

Public GLOBAL_BASEDIR As String
Public Const PYTHONEXE As String = """" & "C:\Users\MustaphaELKAMILI\AppData\Local\Programs\Python\Python314\Scripts\roadmap.exe" & """" & " "

Function RunCommand(cmd As String)
    Dim shellObj As Object
    Set shellObj = CreateObject("WScript.Shell")
    RunCommand = shellObj.Run(cmd, 1, True)
End Function

Function GetBaseDir() As String
    Dim f As FileDialog

    ' If already stored, just return it
    If GLOBAL_BASEDIR <> "" Then
        GetBaseDir = GLOBAL_BASEDIR
        Exit Function
    End If

    ' Otherwise ask the user
    MsgBox "Please select the base directory"
    Set f = Application.FileDialog(msoFileDialogFolderPicker)

    If f.Show <> -1 Then
        MsgBox "No folder selected.", vbExclamation
        Exit Function
    End If

    GLOBAL_BASEDIR = """" & f.SelectedItems(1) & """"
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

        ' Loop through ALL child nodes (col1, col2, col3, etc.)
        For Each childNode In rowNode.ChildNodes
            oneRow.Add childNode.text
        Next childNode

        table.Add oneRow
    Next rowNode

    Set LoadXMLTable = table
End Function

Sub Btn_Create_RM()
    Dim command As String
    Dim baseDir As String

    baseDir = GetBaseDir()
    If baseDir = "" Then
        Exit Sub   ' user cancelled
    End If

    command = PYTHONEXE & "--basedir " & baseDir & " create --archive"
    Application.StatusBar = "Creating collaborator interfaces..."
    RunCommand command
    Application.StatusBar = False

    MsgBox "Collaborator interfaces successfully created and archived.", vbInformation, "Creation Complete"
End Sub

Sub Btn_Delete_RM()
    Dim command As String
    Dim forceDelete As VbMsgBoxResult
    Dim archiveChoice As VbMsgBoxResult
    Dim baseDir As String

    baseDir = GetBaseDir()
    If baseDir = "" Then
        Exit Sub   ' user cancelled
    End If

    forceDelete = MsgBox("Do you want to FORCE deletion of RM Interfaces?" & vbCrLf & _
                         "(This will delete all generated interfaces)", _
                         vbYesNo + vbQuestion, "Confirm Force Deletion")
    If forceDelete = vbNo Then
        Exit Sub
    End If

    archiveChoice = MsgBox("Do you want to ARCHIVE deleted interfaces?", _
                           vbYesNo + vbQuestion, "Archive Confirmation")

    command = PYTHONEXE & "--basedir " & baseDir & " delete --force"

    If archiveChoice = vbYes Then
        command = command & " --archive"
        Application.StatusBar = "Archiving and deleting interfaces..."
    Else
        Application.StatusBar = "Deleting interfaces..."
    End If

    RunCommand (command)
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
    Dim lastRow As Long

    archiveConfirm = MsgBox("Do you want to proceed with archiving the SYNTHESE sheet?" & vbCrLf & _
                            "A new archive file will be created.", _
                            vbYesNo + vbQuestion, "Confirm Archiving")
    If archiveConfirm = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then
        Exit Sub   ' user cancelled
    End If

    command = PYTHONEXE & "--basedir " & baseDir & " pointage --delete"
    Application.StatusBar = "Archiving SYNTHESE sheet..."
    RunCommand command
    Application.StatusBar = False

    Set ws = ThisWorkbook.Sheets("SYNTHESE")
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).Row

    If lastRow >= 3 Then
        ws.rows("3:" & lastRow).ClearContents
    End If

    MsgBox "SYNTHESE sheet successfully archived and cleared.", vbInformation, "Archive Complete"
End Sub

Sub Btn_Collect_RM_Data()
    Dim xmlPath As String
    Dim result As Collection
    Dim rowData As Collection
    Dim r As Long, c As Long
    Dim command As String
    Dim baseDir As String
    Dim ws As Worksheet
    Dim value As Variant

    baseDir = GetBaseDir()
    If baseDir = "" Then
        Exit Sub   ' user cancelled
    End If

    command = PYTHONEXE & "--basedir " & baseDir & " pointage"
    Application.StatusBar = "Exporting pointage data from collaborator files..."
    RunCommand command
    Application.StatusBar = False

    Set ws = ThisWorkbook.Sheets("SYNTHESE")

    r = ws.Cells(ws.rows.Count, "A").End(xlUp).Row + 1
    If r < 3 Then r = 3

    ' Load XML file
    xmlPath = Replace(baseDir, """", "") & "\pointage_output.xml"
    Set result = LoadXMLTable(xmlPath)

    ' Write rows from XML
    For Each rowData In result
        c = 1
        For Each value In rowData
            ws.Cells(r, c).value = value
            c = c + 1
        Next value
        r = r + 1
    Next rowData

    If Dir(xmlPath) <> "" Then
        Kill xmlPath
    End If

    MsgBox "Pointage successfully imported from 'RM_Collaborateurs'"
End Sub

Sub Btn_Update_LC()
    Dim command As String
    Dim baseDir As String

    baseDir = GetBaseDir()
    If baseDir = "" Then
        Exit Sub   ' user cancelled
    End If

    command = PYTHONEXE & "--basedir " & baseDir & " update"

    Application.StatusBar = "Updating conditional lists (LC) in all files..."
    RunCommand command
    Application.StatusBar = False

    MsgBox "LC successfully updated in template and all collaborator files.", vbInformation, "Update Complete"
End Sub
