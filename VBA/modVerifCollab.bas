Option Explicit

Sub Btn_Collect_Collab_nb_h()
    Dim wsSynth As Worksheet, wsGI As Worksheet, wsVerif As Worksheet
    Dim lastRow As Long, lastCol As Long, r As Long, c As Long, i As Long
    Dim collabName As String, key As String, tmp As String
    Dim weekCode As Variant, hoursVal As Variant
    Dim collabList As Collection
    Dim collabDict As Object, weekDict As Object, sumDict As Object
    Dim arrWeeks() As String

    On Error GoTo ErrorHandler

    On Error Resume Next
    Set wsSynth = ThisWorkbook.Sheets(SHEET_SYNTHESE)
    Set wsGI = ThisWorkbook.Sheets(SHEET_GESTION_INTERFACES)
    Set wsVerif = ThisWorkbook.Sheets(SHEET_VERIF_COLLABORATEUR)
    On Error GoTo ErrorHandler

    If wsSynth Is Nothing Then MsgBox "SYNTHESE sheet not found.", vbCritical, "Error": Exit Sub
    If wsGI Is Nothing Then MsgBox "Gestion_Interfaces sheet not found.", vbCritical, "Error": Exit Sub
    If wsVerif Is Nothing Then MsgBox "Vérif_Collaborateur sheet not found.", vbCritical, "Error": Exit Sub

    If GetBaseDir() = "" Then Exit Sub
    Application.ScreenUpdating = False

    ' Collect unique collaborators from Gestion_Interfaces + SYNTHESE
    Set collabList = New Collection
    Set collabDict = CreateObject("Scripting.Dictionary")

    lastRow = wsGI.Cells(wsGI.Rows.Count, 2).End(xlUp).Row
    For r = 3 To lastRow
        collabName = Trim(CStr(wsGI.Cells(r, 2).Value))
        If collabName <> "" And Not collabDict.Exists(collabName) Then
            collabDict.Add collabName, True: collabList.Add collabName
        End If
    Next r

    lastRow = wsSynth.Cells(wsSynth.Rows.Count, SYN_COL_COLLAB).End(xlUp).Row
    If lastRow < SYN_FIRST_DATA_ROW Then lastRow = SYN_FIRST_DATA_ROW - 1
    For r = SYN_FIRST_DATA_ROW To lastRow
        collabName = Trim(CStr(wsSynth.Cells(r, SYN_COL_COLLAB).Value))
        If collabName <> "" And Not collabDict.Exists(collabName) Then
            collabDict.Add collabName, True: collabList.Add collabName
        End If
    Next r

    If collabList.Count = 0 Then
        Application.ScreenUpdating = True
        MsgBox "No collaborators found.", vbInformation, "Nothing to Do": Exit Sub
    End If

    ' Aggregate hours per (collaborator, week)
    Set weekDict = CreateObject("Scripting.Dictionary")
    Set sumDict = CreateObject("Scripting.Dictionary")

    lastRow = wsSynth.Cells(wsSynth.Rows.Count, SYN_COL_HOURS).End(xlUp).Row
    If lastRow < SYN_FIRST_DATA_ROW Then lastRow = SYN_FIRST_DATA_ROW - 1

    For r = SYN_FIRST_DATA_ROW To lastRow
        collabName = Trim(CStr(wsSynth.Cells(r, SYN_COL_COLLAB).Value))
        weekCode = Trim(CStr(wsSynth.Cells(r, SYN_COL_WEEK).Value))
        hoursVal = wsSynth.Cells(r, SYN_COL_HOURS).Value

        If collabName <> "" And weekCode <> "" And collabDict.Exists(collabName) And IsNumeric(hoursVal) Then
            If Not weekDict.Exists(weekCode) Then weekDict.Add weekCode, weekCode
            key = collabName & "||" & weekCode
            If sumDict.Exists(key) Then sumDict(key) = sumDict(key) + CDbl(hoursVal) Else sumDict.Add key, CDbl(hoursVal)
        End If
    Next r

    If weekDict.Count = 0 Then
        ' No weeks found: just refresh collaborator list
        lastRow = wsVerif.Cells(wsVerif.Rows.Count, VERIF_COL_COLLAB).End(xlUp).Row
        lastCol = wsVerif.Cells(VERIF_HEADER_ROW, wsVerif.Columns.Count).End(xlToLeft).Column
        If lastRow >= VERIF_FIRST_COLLAB_ROW Then
            wsVerif.Range(wsVerif.Cells(VERIF_FIRST_COLLAB_ROW, VERIF_COL_COLLAB), _
                          wsVerif.Cells(lastRow, lastCol)).ClearContents
        End If
        For i = 1 To collabList.Count
            wsVerif.Cells(VERIF_FIRST_COLLAB_ROW + i - 1, VERIF_COL_COLLAB).Value = collabList(i)
        Next i
        Application.ScreenUpdating = True
        MsgBox "No SXXYY entries found. Collaborator list refreshed.", vbInformation, "No Data": Exit Sub
    End If

    ' Sort week codes
    ReDim arrWeeks(1 To weekDict.Count)
    i = 1
    For Each weekCode In weekDict.Keys: arrWeeks(i) = CStr(weekCode): i = i + 1: Next weekCode
    For i = LBound(arrWeeks) To UBound(arrWeeks) - 1
        For c = i + 1 To UBound(arrWeeks)
            If arrWeeks(c) < arrWeeks(i) Then tmp = arrWeeks(i): arrWeeks(i) = arrWeeks(c): arrWeeks(c) = tmp
        Next c
    Next i

    ' Clear existing data
    lastRow = wsVerif.Cells(wsVerif.Rows.Count, VERIF_COL_COLLAB).End(xlUp).Row
    lastCol = wsVerif.Cells(VERIF_HEADER_ROW, wsVerif.Columns.Count).End(xlToLeft).Column
    If lastCol < VERIF_FIRST_WEEK_COL Then lastCol = VERIF_FIRST_WEEK_COL
    If lastRow >= VERIF_FIRST_COLLAB_ROW Then
        wsVerif.Range(wsVerif.Cells(VERIF_FIRST_COLLAB_ROW, VERIF_COL_COLLAB), _
                      wsVerif.Cells(lastRow, lastCol)).ClearContents
    End If
    wsVerif.Range(wsVerif.Cells(VERIF_HEADER_ROW, VERIF_FIRST_WEEK_COL), _
                  wsVerif.Cells(VERIF_HEADER_ROW, lastCol)).ClearContents

    ' Write week headers with template formatting
    c = VERIF_FIRST_WEEK_COL
    For i = LBound(arrWeeks) To UBound(arrWeeks)
        wsVerif.Cells(VERIF_HEADER_ROW, c).Value = arrWeeks(i)
        wsVerif.Cells(VERIF_HEADER_ROW, VERIF_FIRST_WEEK_COL).Copy
        wsVerif.Cells(VERIF_HEADER_ROW, c).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        c = c + 1
    Next i

    ' Write collaborator names with template formatting
    For i = 1 To collabList.Count
        r = VERIF_FIRST_COLLAB_ROW + i - 1
        wsVerif.Cells(r, VERIF_COL_COLLAB).Value = collabList(i)
        wsVerif.Cells(VERIF_FIRST_COLLAB_ROW, VERIF_COL_COLLAB).Copy
        wsVerif.Cells(r, VERIF_COL_COLLAB).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
    Next i

    ' Rebuild week-to-column lookup
    Set weekDict = CreateObject("Scripting.Dictionary")
    c = VERIF_FIRST_WEEK_COL
    For i = LBound(arrWeeks) To UBound(arrWeeks): weekDict.Add arrWeeks(i), c: c = c + 1: Next i

    ' Fill intersection matrix
    For i = 1 To collabList.Count
        collabName = collabList(i)
        r = VERIF_FIRST_COLLAB_ROW + i - 1
        For Each weekCode In weekDict.Keys
            key = collabName & "||" & CStr(weekCode)
            wsVerif.Cells(VERIF_FIRST_COLLAB_ROW, VERIF_FIRST_WEEK_COL).Copy
            With wsVerif.Cells(r, weekDict(weekCode))
                .PasteSpecial xlPasteFormats
                .Value = IIf(sumDict.Exists(key), sumDict(key), 0)
            End With
            Application.CutCopyMode = False
        Next weekCode
    Next i

    ' Compute percentage of non-zero values per week
    Dim colWeek As Long, nonZeroCount As Long, zeroCount As Long, denom As Long
    lastRow = wsVerif.Cells(wsVerif.Rows.Count, VERIF_COL_COLLAB).End(xlUp).Row
    If lastRow < VERIF_FIRST_COLLAB_ROW Then lastRow = VERIF_FIRST_COLLAB_ROW - 1

    For Each weekCode In weekDict.Keys
        colWeek = CLng(weekDict(weekCode))
        nonZeroCount = 0: zeroCount = 0
        For r = VERIF_FIRST_COLLAB_ROW To lastRow
            hoursVal = wsVerif.Cells(r, colWeek).Value
            If IsNumeric(hoursVal) Then
                If CDbl(hoursVal) <> 0 Then nonZeroCount = nonZeroCount + 1 Else zeroCount = zeroCount + 1
            End If
        Next r
        denom = nonZeroCount + zeroCount
        wsVerif.Cells(PERCENTAGE_NONZEROS_ROW, PERCENTAGE_NONZEROS_COL).Copy
        With wsVerif.Cells(PERCENTAGE_NONZEROS_ROW, colWeek)
            .PasteSpecial xlPasteFormats
            .Value = IIf(denom > 0, nonZeroCount / denom, 0)
        End With
        Application.CutCopyMode = False
    Next weekCode

    Application.ScreenUpdating = True
    MsgBox "Vérif_Collaborateur table updated.", vbInformation, "Update Complete"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error in Collect_Collab_nb_h: " & Err.Number & " - " & Err.Description, vbCritical, "Unexpected Error"
End Sub

Sub Btn_Reset_Verif_Collaborateur()
    Dim wsVerif As Worksheet
    Dim archiveConfirm As VbMsgBoxResult
    Dim baseDir As String, archivePath As String, timestamp As String
    Dim lastRow As Long, lastCol As Long

    archiveConfirm = MsgBox("This will delete all data from Vérif_Collaborateur." & vbCrLf & vbCrLf & _
                           "Do you want to archive the current data before resetting?" & vbCrLf & _
                           "(A copy will be saved in the Archived folder)", _
                           vbYesNoCancel + vbQuestion, "Confirm Reset")
    If archiveConfirm = vbCancel Then Exit Sub

    On Error Resume Next
    Set wsVerif = ThisWorkbook.Sheets(SHEET_VERIF_COLLABORATEUR)
    On Error GoTo 0
    If wsVerif Is Nothing Then MsgBox "Vérif_Collaborateur sheet not found.", vbCritical, "Error": Exit Sub

    If archiveConfirm = vbYes Then
        baseDir = GetBaseDir()
        If baseDir = "" Then Exit Sub
        timestamp = Format(Now, "ddmmyyyy_HHMMSS")
        archivePath = baseDir & "\Archived\Vérif_Collaborateur_" & timestamp & ".xlsx"
        Application.ScreenUpdating = False
        Application.StatusBar = "Creating archive file..."
        If Not ArchiveSingleSheet(wsVerif, archivePath, True, SHEET_VERIF_COLLABORATEUR) Then
            Application.StatusBar = False: Application.ScreenUpdating = True: Exit Sub
        End If
        Application.StatusBar = False
    End If

    Application.ScreenUpdating = False

    lastRow = wsVerif.Cells(wsVerif.Rows.Count, VERIF_COL_COLLAB).End(xlUp).Row
    lastCol = wsVerif.Cells(VERIF_HEADER_ROW, wsVerif.Columns.Count).End(xlToLeft).Column

    ' Clear percentage row
    If lastCol >= VERIF_FIRST_WEEK_COL Then
        wsVerif.Range(wsVerif.Cells(PERCENTAGE_NONZEROS_ROW, VERIF_FIRST_WEEK_COL), _
                      wsVerif.Cells(PERCENTAGE_NONZEROS_ROW, lastCol)).ClearContents
    End If

    ' Delete extra rows and columns
    If lastRow >= VERIF_FIRST_COLLAB_ROW + 1 Then
        wsVerif.Rows((VERIF_FIRST_COLLAB_ROW + 1) & ":" & lastRow).Delete Shift:=xlUp
    End If
    If lastCol >= VERIF_FIRST_WEEK_COL + 1 Then
        wsVerif.Columns(VERIF_FIRST_WEEK_COL + 1).Resize(, lastCol - VERIF_FIRST_WEEK_COL).Delete Shift:=xlToLeft
    End If

    ' Reset to base template
    wsVerif.Cells(VERIF_FIRST_COLLAB_ROW, VERIF_COL_COLLAB).ClearContents
    wsVerif.Cells(VERIF_FIRST_COLLAB_ROW, VERIF_FIRST_WEEK_COL).ClearContents
    wsVerif.Cells(VERIF_HEADER_ROW, VERIF_FIRST_WEEK_COL).Value = "SXXYY"

    Application.ScreenUpdating = True

    If archiveConfirm = vbYes Then
        MsgBox "Vérif_Collaborateur archived and reset." & vbCrLf & "Saved to: " & archivePath, vbInformation, "Reset Complete"
    Else
        MsgBox "Vérif_Collaborateur reset to base template.", vbInformation, "Reset Complete"
    End If
End Sub
