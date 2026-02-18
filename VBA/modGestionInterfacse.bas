Option Explicit

Sub Btn_Create_RM()
    Dim baseDir As String
    Dim exitCode As Long

    CleanupGestionInterfaces

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    If Not CreateCollabsXML(baseDir) Then
        MsgBox "Error creating collabs.xml file. Operation aborted.", vbCritical, "Error"
        GoTo ErrorHandler
    End If

    Application.StatusBar = "Creating collaborator interfaces..."
    exitCode = RunCommand(PYTHONEXE & "--basedir " & """" & baseDir & """" & " create --way para")
    Application.StatusBar = False

    If exitCode <> 0 Then
        MsgBox "Error creating interfaces. Exit code: " & exitCode, vbCritical, "Error"
        GoTo ErrorHandler
    End If

    MsgBox "Collaborator interfaces successfully created.", vbInformation, "Creation Complete"
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Sub Btn_Delete_RM()
    Dim baseDir As String
    Dim deleteCommand As String
    Dim archiveChoice As VbMsgBoxResult
    Dim exitCode As Long

    If MsgBox("Do you want to FORCE deletion of RM Interfaces?" & vbCrLf & _
              "(This will delete all generated interfaces)", _
              vbYesNo + vbQuestion, "Confirm Force Deletion") = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    archiveChoice = MsgBox("Do you want to ARCHIVE deleted interfaces?", _
                           vbYesNo + vbQuestion, "Archive Confirmation")

    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    deleteCommand = PYTHONEXE & "--basedir " & """" & baseDir & """" & " delete --force"
    If archiveChoice = vbYes Then
        deleteCommand = deleteCommand & " --archive"
        Application.StatusBar = "Archiving and deleting interfaces..."
    Else
        Application.StatusBar = "Deleting interfaces..."
    End If

    exitCode = RunCommand(deleteCommand)
    Application.StatusBar = False

    If exitCode <> 0 Then
        MsgBox "Error deleting interfaces. Exit code: " & exitCode, vbCritical, "Error"
        GoTo ErrorHandler
    End If

    MsgBox IIf(archiveChoice = vbYes, _
               "Interfaces successfully archived and deleted.", _
               "Interfaces successfully deleted."), vbInformation, "Deletion Complete"

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Sub Btn_Cleanup_RM()
    Dim baseDir As String
    Dim exitCode As Long

    CleanupGestionInterfaces

    If MsgBox("Do you want to proceed with cleaning up missing collaborators?" & vbCrLf & _
              "This will delete interface files for collaborators not in the current list.", _
              vbYesNo + vbQuestion, "Confirm Cleanup") = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    If Not CreateCollabsXML(baseDir) Then
        MsgBox "Error creating collabs.xml file. Operation aborted.", vbCritical, "Error"
        GoTo ErrorHandler
    End If

    Application.StatusBar = "Cleaning up missing collaborator interfaces..."
    exitCode = RunCommand(PYTHONEXE & "--basedir " & """" & baseDir & """" & " cleanup")
    Application.StatusBar = False

    If exitCode <> 0 Then
        MsgBox "Error during cleanup. Exit code: " & exitCode, vbCritical, "Error"
        GoTo ErrorHandler
    End If

    MsgBox "Cleanup complete. Missing collaborator interfaces have been deleted.", vbInformation, "Cleanup Complete"
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub
