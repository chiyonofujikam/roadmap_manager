Option Explicit

Function RunCommand(cmd As String) As Long
    Dim shellObj As Object
    Set shellObj = CreateObject("WScript.Shell")
    RunCommand = shellObj.Run(cmd, 1, True)
End Function

Function GetBaseDir() As String
    Dim f As FileDialog

    If GLOBAL_BASEDIR <> "" Then
        If PYTHONEXE = "" Then PYTHONEXE = """" & GLOBAL_BASEDIR & "\script\roadmap.exe" & """" & " "
        GetBaseDir = GLOBAL_BASEDIR
        Exit Function
    End If

    MsgBox "Veuillez selectionner le dossier racine"
    Set f = Application.FileDialog(msoFileDialogFolderPicker)
    If f.Show <> -1 Then MsgBox "Aucun dossier selectionne.", vbExclamation: Exit Function

    GLOBAL_BASEDIR = f.SelectedItems(1)
    PYTHONEXE = """" & GLOBAL_BASEDIR & "\script\roadmap.exe" & """" & " "
    GetBaseDir = GLOBAL_BASEDIR
End Function

Function LoadXMLTable(filePath As String) As Collection
    Dim xml As Object
    Dim rows As Object, rowNode As Object, childNode As Object
    Dim table As New Collection, oneRow As Collection

    Set xml = CreateObject("MSXML2.DOMDocument")
    xml.async = False
    xml.Load filePath

    If xml.parseError.ErrorCode <> 0 Then
        MsgBox "Erreur d'analyse XML : " & xml.parseError.reason, vbCritical
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
    Dim result As String, i As Long
    Dim charCode As Integer, char As String

    result = ""
    For i = 1 To Len(text)
        charCode = Asc(Mid(text, i, 1))
        char = Mid(text, i, 1)

        If charCode < 32 And charCode <> 9 And charCode <> 10 And charCode <> 13 Then
            ' Skip invalid control characters
        Else
            Select Case char
                Case "&":  result = result & "&amp;"
                Case "<":  result = result & "&lt;"
                Case ">":  result = result & "&gt;"
                Case """": result = result & "&quot;"
                Case "'":  result = result & "&apos;"
                Case Else: result = result & char
            End Select
        End If
    Next i
    EscapeXML = result
End Function

Function CreateCollabsXML(baseDir As String) As Boolean
    Dim xmlPath As String, xmlContent As String
    Dim ws As Worksheet
    Dim row As Long, collabName As String
    Dim xmlStream As Object

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_GESTION_INTERFACES)
    If Err.Number <> 0 Then
        MsgBox "La feuille Gestion_Interfaces est introuvable.", vbCritical, "Erreur"
        CreateCollabsXML = False: Exit Function
    End If
    On Error GoTo 0

    xmlPath = baseDir & "\collabs.xml"
    xmlContent = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & "<collaborators>" & vbCrLf

    row = 3
    Do While True
        collabName = Trim(ws.Cells(row, 2).Value)
        If collabName = "" Then Exit Do
        xmlContent = xmlContent & "  <collaborator>" & EscapeXML(collabName) & "</collaborator>" & vbCrLf
        row = row + 1
    Loop
    xmlContent = xmlContent & "</collaborators>"

    On Error Resume Next
    Set xmlStream = CreateObject("ADODB.Stream")
    If Err.Number <> 0 Then
        MsgBox "Erreur lors de la creation du flux de fichier : " & Err.Description, vbCritical, "Erreur"
        CreateCollabsXML = False: Exit Function
    End If
    On Error GoTo 0

    xmlStream.Type = 2
    xmlStream.Charset = "UTF-8"
    xmlStream.Open
    xmlStream.WriteText xmlContent
    xmlStream.SaveToFile xmlPath, 2
    xmlStream.Close
    Set xmlStream = Nothing

    CreateCollabsXML = True
End Function

Function CreateLCExcel(baseDir As String) As Boolean
    Dim ws As Worksheet, wsNew As Worksheet
    Dim wbNew As Workbook
    Dim excelPath As String
    Dim i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_LC)
    If Err.Number <> 0 Then
        MsgBox "La feuille LC est introuvable.", vbCritical, "Erreur"
        CreateLCExcel = False: Exit Function
    End If
    On Error GoTo 0

    excelPath = baseDir & "\LC.xlsx"
    On Error Resume Next: Kill excelPath: On Error GoTo 0

    On Error Resume Next
    Set wbNew = Workbooks.Add
    If Err.Number <> 0 Then
        MsgBox "Erreur lors de la creation du nouveau classeur : " & Err.Description, vbCritical, "Erreur"
        CreateLCExcel = False: Exit Function
    End If

    ws.Copy Before:=wbNew.Sheets(1)
    If Err.Number <> 0 Then
        MsgBox "Erreur lors de la copie de la feuille LC : " & Err.Description, vbCritical, "Erreur"
        wbNew.Close SaveChanges:=False
        CreateLCExcel = False: Exit Function
    End If
    On Error GoTo 0

    Set wsNew = wbNew.ActiveSheet
    If wsNew.Name <> SHEET_LC Then wsNew.Name = SHEET_LC

    ' Preserve exact displayed values (prevent date conversion)
    Dim usedRngSource As Range, cellSource As Range
    Dim cellText As String

    Set usedRngSource = ws.UsedRange
    On Error Resume Next
    If Not usedRngSource Is Nothing Then
        wsNew.UsedRange.NumberFormat = "@"
        For Each cellSource In usedRngSource
            If Not IsEmpty(cellSource) Then
                cellText = CStr(cellSource.text)
                With wsNew.Cells(cellSource.Row, cellSource.Column)
                    .NumberFormat = "@"
                    .Value = IIf(Len(cellText) > 0, cellText, "")
                End With
            End If
        Next cellSource
    End If
    On Error GoTo 0

    ' Remove shapes
    On Error Resume Next
    While wsNew.Shapes.Count > 0: wsNew.Shapes(wsNew.Shapes.Count).Delete: Wend
    On Error GoTo 0

    ' Delete default sheets
    Application.DisplayAlerts = False
    For i = wbNew.Sheets.Count To 1 Step -1
        If wbNew.Sheets(i).Name <> SHEET_LC Then wbNew.Sheets(i).Delete
    Next i
    Application.DisplayAlerts = True

    On Error Resume Next
    wbNew.SaveAs excelPath, FileFormat:=xlOpenXMLWorkbook
    If Err.Number <> 0 Then
        MsgBox "Erreur lors de l'enregistrement de LC.xlsx : " & Err.Description, vbCritical, "Erreur"
        wbNew.Close SaveChanges:=False
        CreateLCExcel = False: Exit Function
    End If
    On Error GoTo 0

    wbNew.Close SaveChanges:=False
    CreateLCExcel = True
End Function

Sub ApplySyntheseRowColoring(ws As Worksheet, Optional startRow As Long = 3, _
                             Optional dataLastCol As Long = 11, Optional helperCol As Long = 53, _
                             Optional thresholdVal As Double = 35)
    Dim lastRow As Long, r As Long
    Dim totalVal As Variant
    Dim redColor As Long, greenColor As Long

    redColor = RGB(255, 0, 0)
    greenColor = RGB(0, 176, 80)

    lastRow = ws.Cells(ws.Rows.Count, helperCol).End(xlUp).Row
    If lastRow < startRow Then lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < startRow Then Exit Sub

    For r = startRow To lastRow
        totalVal = ws.Cells(r, helperCol).Value
        If IsNumeric(totalVal) Then
            If Val(totalVal) < thresholdVal Then
                ws.Range(ws.Cells(r, 1), ws.Cells(r, dataLastCol)).Interior.Color = redColor
            Else
                ws.Range(ws.Cells(r, 1), ws.Cells(r, dataLastCol)).Interior.Color = greenColor
            End If
        End If
        ws.Cells(r, helperCol).ClearContents
    Next r
End Sub

Sub ImportPointageRows(ws As Worksheet, result As Collection, startRow As Long, _
                       ByRef rowsImported As Long, _
                       Optional dataLastCol As Long = 11, Optional helperCol As Long = 53)
    Dim rowData As Collection
    Dim c As Long, r As Long
    Dim Value As Variant

    rowsImported = 0
    r = startRow

    For Each rowData In result
        For c = 1 To rowData.Count
            Value = rowData(c)
            If c <= dataLastCol Then
                ws.Cells(r, c).Value = Value
            ElseIf c = 12 Then
                ws.Cells(r, helperCol).Value = Value
            End If
        Next c
        r = r + 1
        rowsImported = rowsImported + 1
    Next rowData
End Sub

Sub UpdateSyntheseFromLC(wsSynth As Worksheet, wsLC As Worksheet, startRow As Long, endRow As Long)
    Dim lcDict As Object
    Dim lcArr As Variant
    Dim lcLastRow As Long, i As Long, r As Long, matchRow As Long
    Dim key As String, valE As String, valF As String, valG As String
    Dim part1 As String, part2 As String
    Dim posSprint As Long

    lcLastRow = wsLC.Cells(wsLC.Rows.Count, LC_LOOKUP_COL_F).End(xlUp).Row
    If lcLastRow < LC_LOOKUP_FIRST_ROW Then lcLastRow = LC_LOOKUP_FIRST_ROW - 1

    Set lcDict = CreateObject("Scripting.Dictionary")
    lcDict.CompareMode = 1

    If lcLastRow >= LC_LOOKUP_FIRST_ROW Then
        lcArr = wsLC.Range(wsLC.Cells(LC_LOOKUP_FIRST_ROW, LC_LOOKUP_COL_F), _
                           wsLC.Cells(lcLastRow, LC_LOOKUP_COL_K)).Value
        For i = 1 To UBound(lcArr, 1)
            key = Trim(CStr(lcArr(i, 5))) & LC_LOOKUP_KEY_DELIM & Trim(CStr(lcArr(i, 2))) & LC_LOOKUP_KEY_DELIM & _
                  Trim(CStr(lcArr(i, 1))) & LC_LOOKUP_KEY_DELIM & Trim(CStr(lcArr(i, 6)))
            If lcDict.Exists(key) Then lcDict(key) = -1 Else lcDict.Add key, LC_LOOKUP_FIRST_ROW + i - 1
        Next i
    End If

    For r = startRow To endRow
        valE = Trim(CStr(wsSynth.Cells(r, SYN_COL_E).Value))
        valF = Trim(CStr(wsSynth.Cells(r, SYN_COL_F).Value))
        valG = Trim(CStr(wsSynth.Cells(r, SYN_COL_G).Value))

        posSprint = InStr(1, valE, "Sprint", vbTextCompare)
        If posSprint = 0 Then
            wsSynth.Cells(r, SYN_COL_H).ClearContents
            wsSynth.Cells(r, SYN_COL_I).ClearContents
        Else
            part1 = Trim(Left(valE, posSprint - 1))
            part2 = Trim(Mid(valE, posSprint + 6))
            key = valG & LC_LOOKUP_KEY_DELIM & valF & LC_LOOKUP_KEY_DELIM & part1 & LC_LOOKUP_KEY_DELIM & part2

            If lcDict.Exists(key) Then
                matchRow = lcDict(key)
                If matchRow >= 0 Then
                    wsSynth.Cells(r, SYN_COL_H).Value = wsLC.Cells(matchRow, LC_LOOKUP_COL_H).Value
                    wsSynth.Cells(r, SYN_COL_I).Value = wsLC.Cells(matchRow, LC_LOOKUP_COL_I).Value
                Else
                    wsSynth.Cells(r, SYN_COL_H).ClearContents
                    wsSynth.Cells(r, SYN_COL_I).ClearContents
                End If
            Else
                wsSynth.Cells(r, SYN_COL_H).ClearContents
                wsSynth.Cells(r, SYN_COL_I).ClearContents
            End If
        End If
    Next r
End Sub

Sub CleanupGestionInterfaces()
    Dim ws As Worksheet
    Dim row As Long, lastRow As Long, i As Long
    Dim collabNames As Collection, collabName As String

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_GESTION_INTERFACES)
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0

    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 3 Then Exit Sub

    Set collabNames = New Collection
    For row = 3 To lastRow
        collabName = Trim(ws.Cells(row, 2).Value)
        If collabName <> "" Then collabNames.Add collabName
    Next row

    If lastRow >= 3 Then ws.Rows("3:" & lastRow).ClearContents

    i = 1
    For row = 3 To 3 + collabNames.Count - 1
        ws.Cells(row, 2).Value = collabNames(i)
        i = i + 1
    Next row
End Sub

Function UpdateLCInWorkbook(targetPath As String, wsSource As Worksheet, _
                           ByVal srcLastRow As Long, ByRef srcValues As Variant, _
                           Optional ByRef failureReason As String = "", _
                           Optional ByVal fillPointage As Boolean = False, _
                           Optional ByRef lcPointageMap As Object = Nothing) As Boolean
    Dim wb As Workbook, wsDest As Worksheet
    Dim lastRow As Long
    Dim fso As Object
    Dim pointageErr As String

    UpdateLCInWorkbook = False
    failureReason = ""
    On Error GoTo CleanExit

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(targetPath) Then
        failureReason = "Fichier introuvable."
        Exit Function
    End If
    Set fso = Nothing

    Set wb = Workbooks.Open(targetPath, UpdateLinks:=False, ReadOnly:=False, _
                            Notify:=False, AddToMru:=False)
    If wb Is Nothing Then
        failureReason = "Impossible d'ouvrir le classeur."
        Exit Function
    End If
    If wb.Windows.Count > 0 Then wb.Windows(1).Visible = False

    Set wsDest = wb.Sheets(SHEET_LC)
    If wsDest Is Nothing Then
        failureReason = "Feuille LC absente dans le classeur cible."
        GoTo CleanExit
    End If

    lastRow = srcLastRow
    If lastRow < 2 Then
        failureReason = "Aucune donnee LC source a copier."
        GoTo CleanExit
    End If

    ' Enforce destination formats before writing values:
    ' - Text: B,C,D,E,F,G,J
    ' - Date: H
    ' - Numeric: I,K
    wsDest.Range("B2:G" & lastRow).NumberFormat = "@"
    wsDest.Range("J2:J" & lastRow).NumberFormat = "@"
    wsDest.Range("H2:H" & lastRow).NumberFormat = "dd/mm/yyyy"
    wsDest.Range("I2:I" & lastRow).NumberFormat = "General"
    wsDest.Range("K2:K" & lastRow).NumberFormat = "General"

    ' Ecriture en bloc (beaucoup plus rapide que Copy/Paste + boucle cellule par cellule)
    wsDest.Range("B2:K" & lastRow).ClearContents
    wsDest.Range("B2:K" & lastRow).Value2 = srcValues

    If fillPointage Then
        If lcPointageMap Is Nothing Then
            failureReason = "Table de correspondance LC->POINTAGE indisponible."
            GoTo CleanExit
        End If
        If Not FillPointageFromLCMap(wb, lcPointageMap, pointageErr) Then
            failureReason = pointageErr
            GoTo CleanExit
        End If
    End If

    If wb.Windows.Count > 0 Then wb.Windows(1).Visible = True
    wb.Close SaveChanges:=True
    UpdateLCInWorkbook = True
    Exit Function

CleanExit:
    On Error Resume Next
    If Err.Number <> 0 And failureReason = "" Then
        failureReason = "Erreur VBA " & Err.Number & " : " & Err.Description
    End If
    If Not wb Is Nothing Then
        If wb.Windows.Count > 0 Then wb.Windows(1).Visible = True
        wb.Close SaveChanges:=False
    End If
End Function

Function FillPointageFromLCMap(targetWb As Workbook, lcPointageMap As Object, _
                               Optional ByRef failureReason As String = "") As Boolean
    Dim wsPointage As Worksheet
    Dim lastRow As Long, r As Long
    Dim key As String
    Dim valE As String, valF As String, valG As String
    Dim part1 As String, part2 As String
    Dim posSprint As Long
    Dim mappedVals As Variant

    FillPointageFromLCMap = False
    failureReason = ""
    On Error GoTo ErrorHandler

    Set wsPointage = targetWb.Sheets(SHEET_POINTAGE)
    If wsPointage Is Nothing Then
        failureReason = "Feuille POINTAGE absente dans le classeur cible."
        Exit Function
    End If

    lastRow = wsPointage.Cells(wsPointage.Rows.Count, "E").End(xlUp).Row
    If lastRow < 4 Then
        FillPointageFromLCMap = True
        Exit Function
    End If

    wsPointage.Range("H4:H" & lastRow).NumberFormat = "dd/mm/yyyy"

    For r = 4 To lastRow
        valE = Trim$(CStr(wsPointage.Cells(r, "E").Value))
        valF = Trim$(CStr(wsPointage.Cells(r, "F").Value))
        valG = Trim$(CStr(wsPointage.Cells(r, "G").Value))

        posSprint = InStr(1, valE, "Sprint", vbTextCompare)
        If posSprint = 0 Then
            wsPointage.Cells(r, "H").ClearContents
            wsPointage.Cells(r, "I").ClearContents
            GoTo NextPointageRow
        End If

        part1 = Trim$(Left$(valE, posSprint - 1))      ' -> LC G
        part2 = Trim$(Mid$(valE, posSprint + 6))       ' -> LC K
        If part2 = "" Then
            wsPointage.Cells(r, "H").ClearContents
            wsPointage.Cells(r, "I").ClearContents
            GoTo NextPointageRow
        End If

        ' Mapping final (cas reel RM):
        ' POINTAGE E (partie avant "Sprint") -> LC F
        ' POINTAGE F -> LC G
        ' POINTAGE G -> LC J
        ' POINTAGE E (partie apres "Sprint") -> LC K
        key = part1 & LC_LOOKUP_KEY_DELIM & valF & LC_LOOKUP_KEY_DELIM & valG & LC_LOOKUP_KEY_DELIM & part2

        If lcPointageMap.Exists(key) Then
            mappedVals = lcPointageMap(key)
            wsPointage.Cells(r, "H").Value = mappedVals(0) ' LC H -> POINTAGE H
            wsPointage.Cells(r, "I").Value = mappedVals(1) ' LC I -> POINTAGE I
        Else
            wsPointage.Cells(r, "H").ClearContents
            wsPointage.Cells(r, "I").ClearContents
        End If
NextPointageRow:
    Next r

    FillPointageFromLCMap = True
    Exit Function

ErrorHandler:
    failureReason = "Erreur lors du remplissage de POINTAGE : " & Err.Number & " - " & Err.Description
End Function

Sub FixHiddenWindows()
    Dim baseDir As String, templatePath As String, rmFolder As String
    Dim fileName As String, filePath As String
    Dim wb As Workbook, win As Window
    Dim fixedCount As Long
    Dim fileList As Collection
    Dim fso As Object
    Dim fileCount As Long

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    templatePath = baseDir & "\RM_template.xlsx"
    rmFolder = baseDir & "\RM_Collaborateurs"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo ErrorHandler

    fixedCount = 0

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(templatePath) Then
        Set wb = Workbooks.Open(templatePath, UpdateLinks:=False, ReadOnly:=False)
        If Not wb Is Nothing Then
            For Each win In wb.Windows: win.Visible = True: Next win
            wb.Close SaveChanges:=True
            fixedCount = fixedCount + 1
        End If
    End If
    Set fso = Nothing
    On Error GoTo ErrorHandler

    Set fileList = New Collection
    fileName = Dir(rmFolder & "\RM_*.xlsx")
    Do While fileName <> ""
        If Left$(fileName, 2) <> "~$" Then fileList.Add rmFolder & "\" & fileName
        fileName = Dir()
    Loop

    For fileCount = 1 To fileList.Count
        filePath = fileList(fileCount)
        On Error Resume Next
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(filePath) Then
            Set wb = Workbooks.Open(filePath, UpdateLinks:=False, ReadOnly:=False)
            If Not wb Is Nothing Then
                For Each win In wb.Windows: win.Visible = True: Next win
                wb.Close SaveChanges:=True
                fixedCount = fixedCount + 1
            End If
        End If
        Set fso = Nothing
        On Error GoTo ErrorHandler
    Next fileCount

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Visibilite des fenetres corrigee pour " & fixedCount & " fichier(s).", vbInformation, "Correction terminee"
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Function ArchiveSingleSheet(wsSource As Worksheet, archivePath As String, _
                            Optional removeShapesAndOle As Boolean = True, _
                            Optional forceSheetName As String = "") As Boolean
    Dim newWb As Workbook, newWs As Worksheet
    Dim i As Long

    ArchiveSingleSheet = False
    On Error GoTo ErrHandler
    Application.DisplayAlerts = False

    Set newWb = Workbooks.Add
    wsSource.Copy Before:=newWb.Sheets(1)
    Set newWs = newWb.Sheets(1)

    If forceSheetName <> "" Then
        On Error Resume Next: newWs.Name = forceSheetName: On Error GoTo ErrHandler
    End If

    If removeShapesAndOle Then
        On Error Resume Next
        For i = newWs.Shapes.Count To 1 Step -1: newWs.Shapes(i).Delete: Next i
        For i = newWs.OLEObjects.Count To 1 Step -1: newWs.OLEObjects(i).Delete: Next i
        On Error GoTo ErrHandler
    End If

    For i = newWb.Sheets.Count To 1 Step -1
        If newWb.Sheets(i).Name <> newWs.Name Then newWb.Sheets(i).Delete
    Next i

    newWb.SaveAs archivePath, xlOpenXMLWorkbook
    newWb.Close SaveChanges:=False

    Application.DisplayAlerts = True
    ArchiveSingleSheet = True
    Exit Function

ErrHandler:
    On Error Resume Next
    If Not newWb Is Nothing Then newWb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    MsgBox "Erreur lors de l'archivage de la feuille '" & wsSource.Name & "' : " & Err.Description, vbCritical, "Erreur d'archivage"
End Function
