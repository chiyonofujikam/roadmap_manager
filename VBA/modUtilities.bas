Option Explicit

' =============================================================================
' UTILITY FUNCTIONS
' Helper functions used by button handlers
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

