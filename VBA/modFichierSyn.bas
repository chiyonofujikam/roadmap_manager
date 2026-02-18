Option Explicit

Sub Btn_Collect_FS_Data()
    Dim wsFS As Worksheet, wsLC As Worksheet, wsSynth As Worksheet
    Dim lastLcRow As Long, lastSynthRow As Long, r As Long, i As Long, j As Long
    Dim key As String, keyCombo As String, keyCollab As String
    Dim sumDict As Object, lcSumDict As Object
    Dim sumDict2 As Object, lcSumDict2 As Object
    Dim sumDict2Combo As Object, lcSumDict2Combo As Object
    Dim perCollabDict As Object, collabDict As Object
    Dim arrKeys As Variant, parts As Variant
    Dim valE As String, part1 As String, funcLabel As String, strSCode As String
    Dim posSprint As Long, collabName As String
    Dim valG As Double, valH As Double
    Dim collabList As Collection
    Dim headerRow3 As Long, firstCol3 As Long, firstCollabCol As Long
    Dim lastCol3 As Long, lastRow3 As Long, nCombos As Long, nCollabs As Long
    ' Bulk-read arrays
    Dim synArr As Variant, lcArr As Variant
    ' Bulk-write arrays
    Dim outT1() As Variant, outT2() As Variant, outT3() As Variant
    Dim nT1 As Long, nT2 As Long, nT3 As Long
    Dim dblVal As Double

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error Resume Next
    Set wsFS = ThisWorkbook.Sheets(SHEET_FICHIER_SYNTHESE)
    Set wsLC = ThisWorkbook.Sheets(SHEET_LC)
    Set wsSynth = ThisWorkbook.Sheets(SHEET_SYNTHESE)
    On Error GoTo ErrorHandler

    If wsFS Is Nothing Then
        Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: Application.EnableEvents = True
        MsgBox "Fichier de synthèse sheet not found.", vbCritical, "Error": Exit Sub
    End If
    If wsLC Is Nothing Then
        Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: Application.EnableEvents = True
        MsgBox "LC sheet not found.", vbCritical, "Error": Exit Sub
    End If
    If wsSynth Is Nothing Then
        Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: Application.EnableEvents = True
        MsgBox "SYNTHESE sheet not found.", vbCritical, "Error": Exit Sub
    End If

    ' === BULK-READ source data into arrays (single COM call each) ===
    lastSynthRow = wsSynth.Cells(wsSynth.Rows.Count, SYN_COL_HOURS).End(xlUp).Row
    If lastSynthRow < SYN_FIRST_DATA_ROW Then lastSynthRow = SYN_FIRST_DATA_ROW - 1

    ' Read SYNTHESE cols B(2),E(5),F(6),G(7),J(10) => read A:J, use indices 2,5,6,7,10
    If lastSynthRow >= SYN_FIRST_DATA_ROW Then
        synArr = wsSynth.Range("A" & SYN_FIRST_DATA_ROW & ":J" & lastSynthRow).Value
    End If

    lastLcRow = wsLC.Cells(wsLC.Rows.Count, LC_LOOKUP_COL_F).End(xlUp).Row
    If lastLcRow < LC_FIRST_ROW Then lastLcRow = LC_FIRST_ROW

    ' Read LC cols F(6),G(7),I(9),J(10) => read F:J, use indices 1,2,4,5
    If lastLcRow >= LC_FIRST_ROW Then
        lcArr = wsLC.Range(wsLC.Cells(LC_FIRST_ROW, LC_LOOKUP_COL_F), _
                           wsLC.Cells(lastLcRow, LC_LOOKUP_COL_J)).Value
    End If

    ' === Build sorted collaborator list in memory ===
    Set collabList = New Collection
    Set collabDict = CreateObject("Scripting.Dictionary")
    collabDict.CompareMode = 1

    If IsArray(synArr) Then
        For r = 1 To UBound(synArr, 1)
            collabName = Trim$(CStr(synArr(r, 2)))  ' col B
            If collabName <> "" Then
                If Not collabDict.Exists(collabName) Then
                    collabDict.Add collabName, True
                    collabList.Add collabName
                End If
            End If
        Next r
    End If

    If collabList.Count > 1 Then
        Dim arrNames() As String, tmpName As String
        Dim ii As Long, jj As Long
        ReDim arrNames(1 To collabList.Count)
        For i = 1 To collabList.Count: arrNames(i) = collabList(i): Next i
        For ii = LBound(arrNames) To UBound(arrNames) - 1
            For jj = ii + 1 To UBound(arrNames)
                If arrNames(jj) < arrNames(ii) Then
                    tmpName = arrNames(ii): arrNames(ii) = arrNames(jj): arrNames(jj) = tmpName
                End If
            Next jj
        Next ii
        Set collabList = New Collection
        For i = LBound(arrNames) To UBound(arrNames): collabList.Add arrNames(i): Next i
    End If

    ' === ALL AGGREGATION from arrays (zero COM calls) ===

    ' Table 1: SYNTHESE!J by (split-E-part1, G)
    Set sumDict = CreateObject("Scripting.Dictionary"):       sumDict.CompareMode = 1
    ' Tables 2+3: SYNTHESE!J by StrS, by (Fonction,StrS), by (Fonction,StrS,Collab)
    Set sumDict2 = CreateObject("Scripting.Dictionary"):      sumDict2.CompareMode = 1
    Set sumDict2Combo = CreateObject("Scripting.Dictionary"):  sumDict2Combo.CompareMode = 1
    Set perCollabDict = CreateObject("Scripting.Dictionary"):  perCollabDict.CompareMode = 1

    If IsArray(synArr) Then
        For r = 1 To UBound(synArr, 1)
            If Not IsNumeric(synArr(r, 10)) Then GoTo NextSynRow  ' col J
            dblVal = CDbl(synArr(r, 10))
            valE = Trim$(CStr(synArr(r, 5)))           ' col E
            posSprint = InStr(1, valE, "Sprint", vbTextCompare)
            If posSprint > 0 Then
                part1 = Trim$(Left$(valE, posSprint - 1))
            Else
                part1 = ""
            End If

            ' Table 1 aggregation: key = (part1, G)
            If posSprint > 0 Then
                Dim sG As String
                sG = Trim$(CStr(synArr(r, 7)))         ' col G
                If part1 <> "" Or sG <> "" Then
                    key = part1 & "||" & sG
                    If sumDict.Exists(key) Then sumDict(key) = sumDict(key) + dblVal Else sumDict.Add key, dblVal
                End If
            End If

            ' Tables 2+3 aggregation: key = (part1, F)
            Dim sF As String
            sF = Trim$(CStr(synArr(r, 6)))             ' col F
            collabName = Trim$(CStr(synArr(r, 2)))     ' col B
            If sF <> "" Then
                If sumDict2.Exists(sF) Then sumDict2(sF) = sumDict2(sF) + dblVal Else sumDict2.Add sF, dblVal
                If part1 <> "" Then
                    keyCombo = part1 & "||" & sF
                    If sumDict2Combo.Exists(keyCombo) Then sumDict2Combo(keyCombo) = sumDict2Combo(keyCombo) + dblVal Else sumDict2Combo.Add keyCombo, dblVal
                    If collabName <> "" Then
                        keyCollab = keyCombo & "||" & collabName
                        If perCollabDict.Exists(keyCollab) Then perCollabDict(keyCollab) = perCollabDict(keyCollab) + dblVal Else perCollabDict.Add keyCollab, dblVal
                    End If
                End If
            End If
NextSynRow:
        Next r
    End If

    ' LC aggregation: by (F,J) for table 1, by G and (F,G) for tables 2+3
    Set lcSumDict = CreateObject("Scripting.Dictionary"):      lcSumDict.CompareMode = 1
    Set lcSumDict2 = CreateObject("Scripting.Dictionary"):     lcSumDict2.CompareMode = 1
    Set lcSumDict2Combo = CreateObject("Scripting.Dictionary"): lcSumDict2Combo.CompareMode = 1

    If IsArray(lcArr) Then
        For r = 1 To UBound(lcArr, 1)
            ' lcArr columns: 1=F, 2=G, 3=H, 4=I, 5=J
            If Not IsNumeric(lcArr(r, 4)) Then GoTo NextLcRow  ' col I
            dblVal = CDbl(lcArr(r, 4))
            Dim lcF As String, lcG As String, lcJ As String
            lcF = Trim$(CStr(lcArr(r, 1)))
            lcG = Trim$(CStr(lcArr(r, 2)))
            lcJ = Trim$(CStr(lcArr(r, 5)))

            ' Table 1: by (F, J)
            If lcF <> "" Or lcJ <> "" Then
                key = lcF & "||" & lcJ
                If lcSumDict.Exists(key) Then lcSumDict(key) = lcSumDict(key) + dblVal Else lcSumDict.Add key, dblVal
            End If

            ' Tables 2+3: by G and (F, G)
            If lcG <> "" Then
                If lcSumDict2.Exists(lcG) Then lcSumDict2(lcG) = lcSumDict2(lcG) + dblVal Else lcSumDict2.Add lcG, dblVal
                If lcF <> "" Then
                    keyCombo = lcF & "||" & lcG
                    If lcSumDict2Combo.Exists(keyCombo) Then lcSumDict2Combo(keyCombo) = lcSumDict2Combo(keyCombo) + dblVal Else lcSumDict2Combo.Add keyCombo, dblVal
                End If
            End If
NextLcRow:
        Next r
    End If

    ' === BUILD OUTPUT ARRAYS in memory, then bulk-write ===

    ' --- Table 1 (E6:J) ---
    nT1 = lcSumDict.Count
    If wsFS.Rows.Count >= 6 Then wsFS.Range("E6:J" & wsFS.Rows.Count).ClearContents

    If nT1 > 0 Then
        ReDim outT1(1 To nT1, 1 To 6)  ' cols E,F,G,H,I,J
        arrKeys = lcSumDict.Keys
        For i = 0 To nT1 - 1
            key = CStr(arrKeys(i))
            parts = Split(key, "||")
            outT1(i + 1, 1) = IIf(UBound(parts) >= 0, parts(0), "")   ' E
            outT1(i + 1, 2) = IIf(UBound(parts) >= 1, parts(1), "")   ' F
            valG = lcSumDict(key)
            outT1(i + 1, 3) = valG                                      ' G
            valH = 0: If sumDict.Exists(key) Then valH = sumDict(key)
            outT1(i + 1, 4) = valH                                      ' H
            If valG <> 0 Then
                outT1(i + 1, 5) = Round(((valH - valG) / valG) * 100, 0)  ' I
            Else
                outT1(i + 1, 5) = Empty
            End If
            outT1(i + 1, 6) = valG - valH                               ' J
        Next i
        wsFS.Range("E6:J" & (5 + nT1)).Value = outT1
        ' Apply formatting in one shot from row 6 template
        If nT1 > 1 Then
            wsFS.Range("E6:J6").Copy
            wsFS.Range("E7:J" & (5 + nT1)).PasteSpecial xlPasteFormats
            Application.CutCopyMode = False
        End If
    End If

    ' --- Table 2 (L6:Q) ---
    nT2 = lcSumDict2Combo.Count
    If wsFS.Rows.Count >= 6 Then wsFS.Range("L6:Q" & wsFS.Rows.Count).ClearContents

    If nT2 > 0 Then
        ReDim outT2(1 To nT2, 1 To 6)  ' cols L,M,N,O,P,Q
        arrKeys = lcSumDict2Combo.Keys
        For i = 0 To nT2 - 1
            keyCombo = CStr(arrKeys(i))
            parts = Split(keyCombo, "||")
            outT2(i + 1, 1) = IIf(UBound(parts) >= 0, parts(0), "")   ' L = funcLabel
            outT2(i + 1, 2) = IIf(UBound(parts) >= 1, parts(1), "")   ' M = strSCode
            valG = lcSumDict2Combo(keyCombo)
            outT2(i + 1, 3) = valG                                      ' N
            valH = 0: If sumDict2Combo.Exists(keyCombo) Then valH = sumDict2Combo(keyCombo)
            outT2(i + 1, 4) = valH                                      ' O
            If valG <> 0 Then
                outT2(i + 1, 5) = Round(((valH - valG) / valG) * 100, 0)  ' P
            Else
                outT2(i + 1, 5) = Empty
            End If
            outT2(i + 1, 6) = valG - valH                               ' Q
        Next i
        wsFS.Range("L6:Q" & (5 + nT2)).Value = outT2
        If nT2 > 1 Then
            wsFS.Range("L6:Q6").Copy
            wsFS.Range("L7:Q" & (5 + nT2)).PasteSpecial xlPasteFormats
            Application.CutCopyMode = False
        End If
    End If

    ' --- Table 3 (S onwards): combo keys + per-collaborator breakdown ---
    nCombos = lcSumDict2Combo.Count
    nCollabs = collabList.Count
    If collabList Is Nothing Or nCollabs = 0 Or nCombos = 0 Then GoTo Done

    headerRow3 = 5
    firstCol3 = wsFS.Range("S5").Column       ' S = 19
    firstCollabCol = firstCol3 + 4            ' W = 23
    lastCol3 = firstCollabCol + nCollabs - 1

    ' Clear previous data (row 7+)
    lastRow3 = wsFS.Cells(wsFS.Rows.Count, firstCol3).End(xlUp).Row
    If lastRow3 >= 7 Then
        wsFS.Range(wsFS.Cells(7, firstCol3), wsFS.Cells(lastRow3, lastCol3)).ClearContents
    End If

    ' Write collaborator headers in row 5 + format row 6 headers
    ' Copy template format in bulk first, then write names in bulk
    If nCollabs > 1 Then
        wsFS.Cells(headerRow3, firstCollabCol).Copy
        wsFS.Range(wsFS.Cells(headerRow3, firstCollabCol), wsFS.Cells(headerRow3, lastCol3)).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        wsFS.Cells(6, firstCollabCol).Copy
        wsFS.Range(wsFS.Cells(6, firstCollabCol), wsFS.Cells(6, lastCol3)).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
    End If
    ' Write header names via array
    Dim hdrArr() As Variant
    ReDim hdrArr(1 To 1, 1 To nCollabs)
    For i = 1 To nCollabs: hdrArr(1, i) = collabList(i): Next i
    wsFS.Range(wsFS.Cells(headerRow3, firstCollabCol), wsFS.Cells(headerRow3, lastCol3)).Value = hdrArr

    ' Build Table 3 data array: nCombos rows x (4 base + nCollabs) cols
    Dim nColsT3 As Long
    nColsT3 = 4 + nCollabs
    ReDim outT3(1 To nCombos, 1 To nColsT3)
    arrKeys = lcSumDict2Combo.Keys

    For i = 0 To nCombos - 1
        keyCombo = CStr(arrKeys(i))
        parts = Split(keyCombo, "||")
        funcLabel = IIf(UBound(parts) >= 0, parts(0), "")
        strSCode = IIf(UBound(parts) >= 1, parts(1), "")

        outT3(i + 1, 1) = strSCode                                              ' S = StrS
        outT3(i + 1, 2) = funcLabel                                             ' T = Livrable
        outT3(i + 1, 3) = lcSumDict2Combo(keyCombo)                             ' U = temps prévu
        valH = 0: If sumDict2Combo.Exists(keyCombo) Then valH = sumDict2Combo(keyCombo)
        outT3(i + 1, 4) = valH                                                  ' V = temps consommé

        For j = 1 To nCollabs
            keyCollab = keyCombo & "||" & collabList(j)
            If perCollabDict.Exists(keyCollab) Then
                outT3(i + 1, 4 + j) = perCollabDict(keyCollab)
            Else
                outT3(i + 1, 4 + j) = 0
            End If
        Next j
    Next i

    ' Bulk-write Table 3 data starting at row 6
    wsFS.Range(wsFS.Cells(6, firstCol3), wsFS.Cells(5 + nCombos, lastCol3)).Value = outT3

    ' Apply formatting in bulk: copy row 6 template to all data rows
    If nCombos > 1 Then
        wsFS.Range(wsFS.Cells(6, firstCol3), wsFS.Cells(6, lastCol3)).Copy
        wsFS.Range(wsFS.Cells(7, firstCol3), wsFS.Cells(5 + nCombos, lastCol3)).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
    End If

Done:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Fichier de synthèse tables updated.", vbInformation, "Update Complete"
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error in Btn_Collect_FS_Data: " & Err.Number & " - " & Err.Description, vbCritical, "Unexpected Error"
End Sub

Sub Btn_Reset_FS()
    Dim wsFS As Worksheet
    Dim archiveConfirm As VbMsgBoxResult
    Dim baseDir As String, archivePath As String, timestamp As String
    Dim lastDataRow As Long, headerRow3 As Long, firstCol3 As Long
    Dim lastCol3 As Long, lastCol3_row6 As Long, lastRow3 As Long

    archiveConfirm = MsgBox("This will clear the generated tables in 'Fichier de synthèse'." & vbCrLf & vbCrLf & _
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
        If baseDir = "" Then Application.ScreenUpdating = True: Exit Sub
        timestamp = Format(Now, "ddmmyyyy_HHMMSS")
        archivePath = baseDir & "\Archived\Fichier_de_synthese_" & timestamp & ".xlsx"
        Application.StatusBar = "Archiving Fichier de synthèse..."
        If Not ArchiveSingleSheet(wsFS, archivePath, True, SHEET_FICHIER_SYNTHESE) Then
            Application.ScreenUpdating = True: Application.StatusBar = False: Exit Sub
        End If
        Application.StatusBar = False
    End If

    ' Table 1 (E:J) - row 6 contents only, rows 7+ full clear
    wsFS.Range("E6:J6").ClearContents
    lastDataRow = wsFS.Cells(wsFS.Rows.Count, "E").End(xlUp).Row
    If lastDataRow < 7 Then lastDataRow = 7
    wsFS.Range("E7:J" & lastDataRow).Clear

    ' Table 2 (L:Q) - row 6 contents only, rows 7+ full clear
    wsFS.Range("L6:Q6").ClearContents
    lastDataRow = wsFS.Cells(wsFS.Rows.Count, "L").End(xlUp).Row
    If lastDataRow < 7 Then lastDataRow = 7
    wsFS.Range("L7:Q" & lastDataRow).Clear

    ' Table 3 (S onwards)
    headerRow3 = 5
    firstCol3 = wsFS.Range("S" & headerRow3).Column
    lastCol3 = wsFS.Cells(headerRow3, wsFS.Columns.Count).End(xlToLeft).Column
    lastCol3_row6 = wsFS.Cells(headerRow3 + 1, wsFS.Columns.Count).End(xlToLeft).Column
    If lastCol3_row6 > lastCol3 Then lastCol3 = lastCol3_row6
    If lastCol3 < firstCol3 Then lastCol3 = firstCol3

    ' Row 5: clear first collab header contents, fully clear extras
    If lastCol3 >= firstCol3 + 4 Then
        wsFS.Cells(headerRow3, firstCol3 + 4).ClearContents
        If lastCol3 >= firstCol3 + 5 Then
            wsFS.Range(wsFS.Cells(headerRow3, firstCol3 + 5), wsFS.Cells(headerRow3, lastCol3)).Clear
        End If
    End If

    ' Row 6: clear contents (keep template formatting), fully clear extras
    wsFS.Range(wsFS.Cells(headerRow3 + 1, firstCol3), wsFS.Cells(headerRow3 + 1, firstCol3 + 4)).ClearContents
    If lastCol3 >= firstCol3 + 5 Then
        wsFS.Range(wsFS.Cells(headerRow3 + 1, firstCol3 + 5), wsFS.Cells(headerRow3 + 1, lastCol3)).Clear
    End If

    ' Rows 7+: full clear
    lastRow3 = wsFS.Cells(wsFS.Rows.Count, firstCol3).End(xlUp).Row
    If lastRow3 >= headerRow3 + 2 Then
        wsFS.Range(wsFS.Cells(headerRow3 + 2, firstCol3), wsFS.Cells(lastRow3, lastCol3)).Clear
    End If

    Application.ScreenUpdating = True
    MsgBox "Fichier de synthèse tables have been cleared.", vbInformation, "Reset Complete"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error in Btn_Reset_FS: " & Err.Number & " - " & Err.Description, vbCritical, "Unexpected Error"
End Sub
