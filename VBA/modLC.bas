Option Explicit

Sub Btn_Update_LC()
    Dim baseDir As String, templatePath As String, rmFolder As String
    Dim wsLCSource As Worksheet
    Dim fileName As String, filePath As String
    Dim fileCount As Long, processedCount As Long, failedCount As Long
    Dim startTime As Double, elapsedTime As Double
    Dim fileList As Collection, failedDetails As Collection
    Dim failureReason As String, shortName As String
    Dim updateOk As Boolean
    Dim finalMsg As String, failMsg As String
    Dim i As Long
    Dim srcLastRow As Long
    Dim srcValues As Variant
    Dim lcPointageMap As Object
    Dim isRmCollabFile As Boolean
    Dim lcF As String, lcG As String, lcJ As String, lcK As String, mapKey As String
    Dim mapVal(1 To 2) As Variant

    If MsgBox("Voulez-vous lancer la mise a jour des listes LC ?" & vbCrLf & _
              "Cette action mettra a jour LC dans le template et tous les fichiers collaborateurs.", _
              vbYesNo + vbQuestion, "Confirmation de mise a jour") = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    On Error Resume Next
    Set wsLCSource = ThisWorkbook.Sheets(SHEET_LC)
    On Error GoTo ErrorHandler
    If wsLCSource Is Nothing Then
        MsgBox "La feuille LC est introuvable dans le classeur courant.", vbCritical, "Erreur"
        Exit Sub
    End If

    If Application.WorksheetFunction.CountA(wsLCSource.Cells) = 0 Then
        MsgBox "La feuille LC source est vide.", vbExclamation, "Aucune action"
        Exit Sub
    End If

    srcLastRow = wsLCSource.UsedRange.Rows(wsLCSource.UsedRange.Rows.Count).Row
    If srcLastRow < 2 Then
        MsgBox "Aucune donnee LC source a copier.", vbExclamation, "Aucune action"
        Exit Sub
    End If
    srcValues = wsLCSource.Range("B2:K" & srcLastRow).Value2
    Set lcPointageMap = CreateObject("Scripting.Dictionary")
    lcPointageMap.CompareMode = 1
    For i = 3 To UBound(srcValues, 1)
        lcF = Trim$(CStr(srcValues(i, 5))) ' LC col F
        lcG = Trim$(CStr(srcValues(i, 6))) ' LC col G
        lcJ = Trim$(CStr(srcValues(i, 9))) ' LC col J
        lcK = Trim$(CStr(srcValues(i, 10))) ' LC col K
        mapKey = lcF & LC_LOOKUP_KEY_DELIM & lcG & LC_LOOKUP_KEY_DELIM & lcJ & LC_LOOKUP_KEY_DELIM & lcK
        If Not lcPointageMap.Exists(mapKey) Then
            mapVal(1) = srcValues(i, 7)    ' LC col H -> POINTAGE H
            mapVal(2) = srcValues(i, 8)    ' LC col I -> POINTAGE I
            lcPointageMap.Add mapKey, Array(mapVal(1), mapVal(2))
        End If
    Next i

    startTime = Timer
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrorHandler

    templatePath = baseDir & "\RM_template.xlsx"
    rmFolder = baseDir & "\RM_Collaborateurs"
    Application.StatusBar = "Mise a jour de LC dans le template et les fichiers collaborateurs..."

    Set fileList = New Collection
    Set failedDetails = New Collection
    fileList.Add templatePath
    fileName = Dir(rmFolder & "\RM_*.xlsx")
    Do While fileName <> ""
        If Left$(fileName, 2) <> "~$" Then fileList.Add rmFolder & "\" & fileName
        fileName = Dir()
    Loop

    processedCount = 0
    failedCount = 0
    For fileCount = 1 To fileList.Count
        filePath = CStr(fileList(fileCount))
        shortName = Mid$(filePath, InStrRev(filePath, "\") + 1)
        isRmCollabFile = (InStr(1, filePath, rmFolder & "\", vbTextCompare) = 1)
        Application.StatusBar = "Mise a jour LC : " & fileCount & " sur " & fileList.Count & " fichiers... (" & shortName & ")"

        failureReason = ""
        On Error Resume Next
        updateOk = UpdateLCInWorkbook(filePath, wsLCSource, srcLastRow, srcValues, failureReason, isRmCollabFile, lcPointageMap)
        If updateOk Then
            processedCount = processedCount + 1
        Else
            failedCount = failedCount + 1
            If failureReason = "" Then failureReason = "Echec inconnu."
            failedDetails.Add shortName & " : " & failureReason
        End If
        If Err.Number <> 0 Then
            If updateOk Then
                failedCount = failedCount + 1
                processedCount = processedCount - 1
                failedDetails.Add shortName & " : Erreur VBA " & Err.Number & " - " & Err.Description
            ElseIf failureReason = "" Then
                failedDetails.Add shortName & " : Erreur VBA " & Err.Number & " - " & Err.Description
            End If
            Err.Clear
        End If
        On Error GoTo ErrorHandler
        DoEvents
    Next fileCount

    elapsedTime = Timer - startTime
    If elapsedTime < 0 Then elapsedTime = elapsedTime + 86400

    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Dim timeMsg As String
    If elapsedTime < 60 Then
        timeMsg = Format(elapsedTime, "0.00") & " seconds"
    Else
        timeMsg = Format(Int(elapsedTime / 60), "0") & " min " & Format(elapsedTime Mod 60, "0.00") & " s"
    End If

    finalMsg = "Mise a jour LC terminee." & vbCrLf & _
               "- Fichiers traites : " & fileList.Count & vbCrLf & _
               "- Succes : " & processedCount & vbCrLf & _
               "- Echecs : " & failedCount & vbCrLf & _
               "Duree : " & timeMsg

    If failedCount > 0 Then
        failMsg = ""
        For i = 1 To failedDetails.Count
            failMsg = failMsg & vbCrLf & "  * " & failedDetails(i)
        Next i
        finalMsg = finalMsg & vbCrLf & vbCrLf & "Fichiers en echec :" & failMsg
        MsgBox finalMsg, vbExclamation, "Mise a jour terminee avec erreurs"
    Else
        MsgBox finalMsg, vbInformation, "Mise a jour terminee"
    End If
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Sub Btn_Reset_LC()
    Dim wsLC As Worksheet
    Dim archiveConfirm As VbMsgBoxResult
    Dim baseDir As String, archivePath As String, timestamp As String
    Dim firstLookupDataRow As Long, lastLookupRow As Long

    archiveConfirm = MsgBox("Cette action va vider la table de correspondance LC (colonnes F a K a partir de la ligne 3)." & vbCrLf & vbCrLf & _
                            "Voulez-vous ARCHIVER la table LC actuelle avant de la vider ?" & vbCrLf & _
                            "(Une copie sera enregistree dans le dossier Archived)", _
                            vbYesNoCancel + vbQuestion, "Confirmation de reinitialisation LC")
    If archiveConfirm = vbCancel Then Exit Sub

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    On Error Resume Next
    Set wsLC = ThisWorkbook.Sheets(SHEET_LC)
    On Error GoTo ErrorHandler

    If wsLC Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "La feuille LC est introuvable.", vbCritical, "Erreur"
        Exit Sub
    End If

    If archiveConfirm = vbYes Then
        baseDir = GetBaseDir()
        If baseDir = "" Then Application.ScreenUpdating = True: Exit Sub
        timestamp = Format(Now, "ddmmyyyy_HHMMSS")
        archivePath = baseDir & "\Archived\LC_" & timestamp & ".xlsx"
        Application.StatusBar = "Creation de l'archive LC..."
        If Not ArchiveSingleSheet(wsLC, archivePath, True, SHEET_LC) Then
            Application.ScreenUpdating = True: Application.StatusBar = False: Exit Sub
        End If
        Application.StatusBar = False
    End If

    firstLookupDataRow = LC_LOOKUP_FIRST_ROW + 1
    lastLookupRow = wsLC.Cells(wsLC.Rows.Count, LC_LOOKUP_COL_F).End(xlUp).Row
    If lastLookupRow < firstLookupDataRow Then lastLookupRow = firstLookupDataRow

    wsLC.Range(wsLC.Cells(firstLookupDataRow, LC_LOOKUP_COL_F), _
               wsLC.Cells(lastLookupRow, LC_LOOKUP_COL_K)).ClearContents

    Application.ScreenUpdating = True
    MsgBox "La table LC a ete videe.", vbInformation, "Reinitialisation terminee"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Erreur dans Btn_Reset_LC : " & Err.Number & " - " & Err.Description, vbCritical, "Erreur inattendue"
End Sub

Sub Btn_Extract_LC_MSP()
    Dim wsLC As Worksheet, wsSrc As Worksheet
    Dim lastRow As Long, r As Long, outIdx As Long
    Dim valN As Variant
    Dim firstLookupDataRow As Long, lastLookupRow As Long
    Dim srcData As Variant, outArr() As Variant
    Dim dict As Object, keyFK As String
    Dim vB As String, vF As String, vN As Variant, vO As Variant, vC As Variant, vU As Variant

    If MsgBox("Voulez-vous regenerer la table de correspondance LC depuis Extract_MSP ?" & vbCrLf & _
              "Cette action ecrasera les valeurs existantes de LC (colonnes F a K a partir de la ligne 3).", _
              vbYesNo + vbQuestion, "Confirmation de generation LC") = vbNo Then Exit Sub

    If GetBaseDir() = "" Then Exit Sub

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error Resume Next
    Set wsLC = ThisWorkbook.Sheets(SHEET_LC)
    Set wsSrc = ThisWorkbook.Sheets(SHEET_EXTRACT_MSP)
    On Error GoTo ErrorHandler

    If wsLC Is Nothing Then
        Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
        MsgBox "La feuille LC est introuvable.", vbCritical, "Erreur": Exit Sub
    End If
    If wsSrc Is Nothing Then
        Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
        MsgBox "La feuille Extract_MSP est introuvable.", vbCritical, "Erreur": Exit Sub
    End If

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "B").End(xlUp).Row
    If lastRow < 3 Then
        Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
        MsgBox "Aucune donnee trouvee dans Extract_MSP.", vbInformation, "Aucune action"
    End If

    ' Bulk-read all needed source columns (B,C,F,N,O,U) into one array (cols A-U = 1-21)
    srcData = wsSrc.Range("A2:U" & lastRow).Value

    ' Deduplicate in memory and build output array
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1
    ReDim outArr(1 To UBound(srcData, 1), 1 To 6)
    outIdx = 0

    For r = 1 To UBound(srcData, 1)
        vB = Trim$(CStr(srcData(r, 2)))          ' col B = source col 2
        If vB <> "" Then
            vF = CStr(srcData(r, 6))              ' col F = source col 6
            vN = srcData(r, 14)                    ' col N = source col 14
            vO = srcData(r, 15)                    ' col O = source col 15
            vC = srcData(r, 3)                     ' col C = source col 3
            vU = srcData(r, 21)                    ' col U = source col 21

            ' Blank out N if it equals 0
            If IsNumeric(vN) Then
                If CDbl(vN) = 0 Then vN = Empty
            End If

            keyFK = vB & "||" & vF & "||" & CStr(vN) & "||" & CStr(vO) & "||" & CStr(vC) & "||" & CStr(vU)
            If Not dict.Exists(keyFK) Then
                dict.Add keyFK, True
                outIdx = outIdx + 1
                outArr(outIdx, 1) = vB             ' -> LC col F
                outArr(outIdx, 2) = vF             ' -> LC col G
                outArr(outIdx, 3) = vN             ' -> LC col H
                outArr(outIdx, 4) = vO             ' -> LC col I
                outArr(outIdx, 5) = vC             ' -> LC col J
                outArr(outIdx, 6) = vU             ' -> LC col K
            End If
        End If
    Next r

    ' Clear existing LC lookup area
    firstLookupDataRow = LC_LOOKUP_FIRST_ROW + 1
    lastLookupRow = wsLC.Cells(wsLC.Rows.Count, LC_LOOKUP_COL_F).End(xlUp).Row
    If lastLookupRow < firstLookupDataRow Then lastLookupRow = firstLookupDataRow
    wsLC.Range(wsLC.Cells(firstLookupDataRow, LC_LOOKUP_COL_F), _
               wsLC.Cells(lastLookupRow, LC_LOOKUP_COL_K)).ClearContents

    ' Bulk-write unique rows in one shot
    If outIdx > 0 Then
        Dim writeArr() As Variant
        ReDim writeArr(1 To outIdx, 1 To 6)
        For r = 1 To outIdx
            writeArr(r, 1) = outArr(r, 1)
            writeArr(r, 2) = outArr(r, 2)
            writeArr(r, 3) = outArr(r, 3)
            writeArr(r, 4) = outArr(r, 4)
            writeArr(r, 5) = outArr(r, 5)
            writeArr(r, 6) = outArr(r, 6)
        Next r
        wsLC.Range(wsLC.Cells(firstLookupDataRow, LC_LOOKUP_COL_F), _
                   wsLC.Cells(firstLookupDataRow + outIdx - 1, LC_LOOKUP_COL_K)).Value = writeArr
    End If

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Table LC generee : " & outIdx & " ligne(s) unique(s) depuis " & (lastRow - 1) & " ligne(s) source.", _
           vbInformation, "Mise a jour terminee"
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Erreur dans Extract_LC_MSP : " & Err.Number & " - " & Err.Description, vbCritical, "Erreur inattendue"
End Sub
