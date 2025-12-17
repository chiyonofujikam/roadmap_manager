Sub Btn_Update_LC_old()
    Dim updateCommand As String
    Dim baseDir As String
    Dim confirmation As VbMsgBoxResult
    Dim exitCode As Long

    confirmation = MsgBox("Do you want to proceed with updating the conditional lists (LC)?" & vbCrLf & _
                          "This will update LC in the template and all collaborator files.", _
                          vbYesNo + vbQuestion, "Confirm Update")
    If confirmation = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    ' Create LC.xlsx file from LC sheet
    If Not CreateLCExcel(baseDir) Then
        MsgBox "Error creating LC.xlsx file. Operation aborted.", vbCritical, "Error"
        GoTo ErrorHandler
    End If

    updateCommand = PYTHONEXE & "--basedir " & """" & baseDir & """" & " update"
    Application.StatusBar = "Updating conditional lists (LC) in all files..."
    exitCode = RunCommand(updateCommand)
    Application.StatusBar = False

    If exitCode <> 0 Then
        MsgBox "Error updating LC. Exit code: " & exitCode, vbCritical, "Error"
        GoTo ErrorHandler
    End If

    MsgBox "LC successfully updated in template and all collaborator files.", vbInformation, "Update Complete"

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub