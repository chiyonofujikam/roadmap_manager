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
            PYTHONEXE = """" & GLOBAL_BASEDIR & "\script\roadmap.exe" & """" & " "
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
    PYTHONEXE = """" & GLOBAL_BASEDIR & "\script\roadmap.exe" & """" & " "
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
    Dim i As Long
    Dim charCode As Integer
    Dim char As String

    result = ""

    ' Escape XML special characters and remove control characters
    For i = 1 To Len(text)
        charCode = Asc(Mid(text, i, 1))
        char = Mid(text, i, 1)

        ' Skip control characters (except tab, line feed, carriage return)
        If charCode < 32 And charCode <> 9 And charCode <> 10 And charCode <> 13 Then
            ' Skip invalid control characters
        Else
            ' Escape XML special characters
            Select Case char
                Case "&"
                    result = result & "&amp;"
                Case "<"
                    result = result & "&lt;"
                Case ">"
                    result = result & "&gt;"
                Case """"
                    result = result & "&quot;"
                Case "'"
                    result = result & "&apos;"
                Case Else
                    result = result & char
            End Select
        End If
    Next i

    EscapeXML = result
End Function

Function CreateCollabsXML(baseDir As String) As Boolean
    Dim xmlPath As String
    Dim ws As Worksheet
    Dim row As Long
    Dim collabName As String
    Dim xmlContent As String
    Dim xmlStream As Object

    ' Check if Gestion_Interfaces sheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Gestion_Interfaces")
    If Err.Number <> 0 Then
        MsgBox "Gestion_Interfaces sheet not found.", vbCritical, "Error"
        CreateCollabsXML = False
        Exit Function
    End If
    On Error GoTo 0

    ' Create collabs.xml file path
    xmlPath = baseDir & "\collabs.xml"
    xmlContent = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    xmlContent = xmlContent & "<collaborators>" & vbCrLf

    ' Read collaborator names from column B, starting at row 3
    row = 3
    Do While True
        collabName = Trim(ws.Cells(row, 2).value)
        If collabName = "" Then Exit Do

        xmlContent = xmlContent & "  <collaborator>" & EscapeXML(collabName) & "</collaborator>" & vbCrLf
        row = row + 1
    Loop

    xmlContent = xmlContent & "</collaborators>"

    ' Write to file using ADODB.Stream for proper UTF-8 encoding
    On Error Resume Next
    Set xmlStream = CreateObject("ADODB.Stream")
    If Err.Number <> 0 Then
        MsgBox "Error creating file stream: " & Err.Description, vbCritical, "Error"
        CreateCollabsXML = False
        Exit Function
    End If
    On Error GoTo 0

    xmlStream.Type = 2 ' Text stream
    xmlStream.Charset = "UTF-8"
    xmlStream.Open
    xmlStream.WriteText xmlContent
    xmlStream.SaveToFile xmlPath, 2 ' Overwrite mode
    xmlStream.Close
    Set xmlStream = Nothing

    CreateCollabsXML = True
End Function

Function CreateLCExcel(baseDir As String) As Boolean
    Dim excelPath As String
    Dim ws As Worksheet
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim usedRange As Range

    ' Check if LC sheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("LC")
    If Err.Number <> 0 Then
        MsgBox "LC sheet not found.", vbCritical, "Error"
        CreateLCExcel = False
        Exit Function
    End If
    On Error GoTo 0

    ' Create LC.xlsx file path
    excelPath = baseDir & "\LC.xlsx"

    ' Delete existing file if it exists
    On Error Resume Next
    Kill excelPath
    On Error GoTo 0

    ' Create new workbook
    On Error Resume Next
    Set wbNew = Workbooks.Add
    If Err.Number <> 0 Then
        MsgBox "Error creating new workbook: " & Err.Description, vbCritical, "Error"
        CreateLCExcel = False
        Exit Function
    End If
    On Error GoTo 0

    ' Copy the entire LC sheet to the new workbook first (before deleting default sheets)
    On Error Resume Next
    ws.Copy Before:=wbNew.Sheets(1)
    If Err.Number <> 0 Then
        MsgBox "Error copying LC sheet: " & Err.Description, vbCritical, "Error"
        wbNew.Close SaveChanges:=False
        CreateLCExcel = False
        Exit Function
    End If
    On Error GoTo 0

    ' The copied sheet will be the active sheet, rename it if needed
    Set wsNew = wbNew.ActiveSheet
    If wsNew.Name <> "LC" Then
        wsNew.Name = "LC"
    End If

    ' Convert all formulas to their displayed values (preserves exact displayed values)
    ' This ensures cells with formulas like "ASSEM.H" show their calculated values
    ' Read original text values from source sheet to preserve exact format and prevent date conversion
    On Error Resume Next
    Dim usedRngSource As Range
    Dim usedRngDest As Range
    Dim cellSource As Range
    Dim cellDest As Range
    Dim cellText As String
    Dim sourceRow As Long, sourceCol As Long
    
    ' Get used range from source sheet (original LC sheet)
    Set usedRngSource = ws.usedRange
    Set usedRngDest = wsNew.usedRange
    
    If Not usedRngSource Is Nothing And Not usedRngDest Is Nothing Then
        ' Format all destination cells as text FIRST
        usedRngDest.NumberFormat = "@"
        
        ' Read text values from SOURCE sheet and write directly to DESTINATION sheet
        ' This preserves the exact displayed text, preventing any date conversion
        For Each cellSource In usedRngSource
            If Not isEmpty(cellSource) Then
                ' Get the displayed text from the original cell (preserves exact format)
                ' .Text property gives us exactly what the user sees, regardless of internal storage
                cellText = CStr(cellSource.text)
                
                ' Calculate corresponding cell in destination sheet
                sourceRow = cellSource.row
                sourceCol = cellSource.Column
                Set cellDest = wsNew.Cells(sourceRow, sourceCol)
                
                ' Ensure format is text and write the text value
                cellDest.NumberFormat = "@"
                If Len(cellText) > 0 Then
                    ' Write as text value - the '@' format should prevent date interpretation
                    cellDest.value = cellText
                Else
                    cellDest.value = ""
                End If
            End If
        Next cellSource
    End If
    On Error GoTo 0

    ' Remove all shapes from the LC sheet (charts, images, buttons, etc.)
    On Error Resume Next
    While wsNew.Shapes.Count > 0
        wsNew.Shapes(wsNew.Shapes.Count).Delete
    Wend
    On Error GoTo 0

    ' Now delete the default sheets (keeping at least the LC sheet we just copied)
    Application.DisplayAlerts = False
    Dim i As Long
    For i = wbNew.Sheets.Count To 1 Step -1
        If wbNew.Sheets(i).Name <> "LC" Then
            wbNew.Sheets(i).Delete
        End If
    Next i
    Application.DisplayAlerts = True

    ' Save and close the workbook
    On Error Resume Next
    wbNew.SaveAs excelPath, FileFormat:=xlOpenXMLWorkbook
    If Err.Number <> 0 Then
        MsgBox "Error saving LC.xlsx file: " & Err.Description, vbCritical, "Error"
        wbNew.Close SaveChanges:=False
        CreateLCExcel = False
        Exit Function
    End If
    On Error GoTo 0

    wbNew.Close SaveChanges:=False

    CreateLCExcel = True
End Function

Sub CleanupGestionInterfaces()
    Dim ws As Worksheet
    Dim row As Long
    Dim lastRow As Long
    Dim collabNames As Collection
    Dim collabName As String
    Dim i As Long
    Dim startRow As Long

    ' Check if Gestion_Interfaces sheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Gestion_Interfaces")
    If Err.Number <> 0 Then
        ' Sheet doesn't exist, exit silently
        Exit Sub
    End If
    On Error GoTo 0

    ' Find the last row with data in column B
    lastRow = ws.Cells(ws.rows.Count, 2).End(xlUp).row

    ' If no data found or last row is before row 3, exit
    If lastRow < 3 Then Exit Sub

    ' Collect all non-empty collaborator names from column B (starting at row 3)
    Set collabNames = New Collection
    startRow = 3

    For row = startRow To lastRow
        collabName = Trim(ws.Cells(row, 2).value)
        If collabName <> "" Then
            collabNames.Add collabName
        End If
    Next row

    ' Clear all rows from row 3 to lastRow
    If lastRow >= startRow Then
        ws.rows(startRow & ":" & lastRow).ClearContents
    End If

    ' Write back the collected names in a compacted format (no empty rows)
    i = 1
    For row = startRow To startRow + collabNames.Count - 1
        ws.Cells(row, 2).value = collabNames(i)
        i = i + 1
    Next row
End Sub
