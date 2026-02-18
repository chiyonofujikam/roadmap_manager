Option Explicit

Sub Btn_Update_LC()
    Dim baseDir As String, templatePath As String, rmFolder As String
    Dim wsLCSource As Worksheet
    Dim fileName As String, filePath As String
    Dim fileCount As Long, processedCount As Long
    Dim startTime As Double, elapsedTime As Double
    Dim fileList As Collection

    If MsgBox("Do you want to proceed with updating the conditional lists (LC)?" & vbCrLf & _
              "This will update LC in the template and all collaborator files.", _
              vbYesNo + vbQuestion, "Confirm Update") = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    On Error Resume Next
    Set wsLCSource = ThisWorkbook.Sheets(SHEET_LC)
    On Error GoTo ErrorHandler
    If wsLCSource Is Nothing Then
        MsgBox "LC sheet not found in the current workbook.", vbCritical, "Error"
        Exit Sub
    End If

    startTime = Timer
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrorHandler

    templatePath = baseDir & "\RM_template.xlsx"
    rmFolder = baseDir & "\RM_Collaborateurs"
    Application.StatusBar = "Updating LC in template and collaborator files..."

    Set fileList = New Collection
    fileList.Add templatePath
    fileName = Dir(rmFolder & "\RM_*.xlsx")
    Do While fileName <> ""
        If Left$(fileName, 2) <> "~$" Then fileList.Add rmFolder & "\" & fileName
        fileName = Dir()
    Loop

    processedCount = 0
    For fileCount = 1 To fileList.Count
        Application.StatusBar = "Updating LC: " & fileCount & " of " & fileList.Count & " files..."
        If UpdateLCInWorkbook(fileList(fileCount), wsLCSource) Then processedCount = processedCount + 1
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

    MsgBox "LC updated in template and " & (processedCount - 1) & " collaborator file(s)." & vbCrLf & _
           "Time: " & timeMsg, vbInformation, "Update Complete"
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

    archiveConfirm = MsgBox("This will clear the LC lookup table (columns F to K starting from row 3)." & vbCrLf & vbCrLf & _
                            "Do you want to ARCHIVE the current LC table before clearing it?" & vbCrLf & _
                            "(A copy will be saved in the Archived folder)", _
                            vbYesNoCancel + vbQuestion, "Confirm LC Reset")
    If archiveConfirm = vbCancel Then Exit Sub

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    On Error Resume Next
    Set wsLC = ThisWorkbook.Sheets(SHEET_LC)
    On Error GoTo ErrorHandler

    If wsLC Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "LC sheet not found.", vbCritical, "Error"
        Exit Sub
    End If

    If archiveConfirm = vbYes Then
        baseDir = GetBaseDir()
        If baseDir = "" Then Application.ScreenUpdating = True: Exit Sub
        timestamp = Format(Now, "ddmmyyyy_HHMMSS")
        archivePath = baseDir & "\Archived\LC_" & timestamp & ".xlsx"
        Application.StatusBar = "Creating LC archive..."
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
    MsgBox "LC table has been cleared.", vbInformation, "Reset Complete"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error in Btn_Reset_LC: " & Err.Number & " - " & Err.Description, vbCritical, "Unexpected Error"
End Sub

Sub Btn_Extract_LC_MSP()
    Dim wsLC As Worksheet, wsSrc As Worksheet
    Dim lastRow As Long, r As Long, outIdx As Long
    Dim valN As Variant
    Dim firstLookupDataRow As Long, lastLookupRow As Long
    Dim srcData As Variant, outArr() As Variant
    Dim dict As Object, keyFK As String
    Dim vB As String, vF As String, vN As Variant, vO As Variant, vC As Variant, vU As Variant

    If MsgBox("Do you want to regenerate the LC lookup table from Extract_MSP?" & vbCrLf & _
              "This will overwrite existing values in LC (columns F to K starting from row 3).", _
              vbYesNo + vbQuestion, "Confirm LC Generation") = vbNo Then Exit Sub

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
        MsgBox "LC sheet not found.", vbCritical, "Error": Exit Sub
    End If
    If wsSrc Is Nothing Then
        Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
        MsgBox "Extract_MSP sheet not found.", vbCritical, "Error": Exit Sub
    End If

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "B").End(xlUp).Row
    If lastRow < 3 Then
        Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
        MsgBox "No data found in Extract_MSP.", vbInformation, "Nothing to Do": Exit Sub
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
    MsgBox "LC table generated: " & outIdx & " unique rows from " & (lastRow - 1) & " source rows.", _
           vbInformation, "Update Complete"
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error in Extract_LC_MSP: " & Err.Number & " - " & Err.Description, vbCritical, "Unexpected Error"
End Sub
