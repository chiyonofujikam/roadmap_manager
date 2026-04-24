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
        MsgBox "Erreur lors de la creation du fichier collabs.xml. Operation annulee.", vbCritical, "Erreur"
        GoTo ErrorHandler
    End If

    Application.StatusBar = "Creation des interfaces collaborateurs..."
    exitCode = RunCommand(PYTHONEXE & "--basedir " & """" & baseDir & """" & " create --way para")
    Application.StatusBar = False

    If exitCode <> 0 Then
        MsgBox "Erreur lors de la creation des interfaces. Code de sortie : " & exitCode, vbCritical, "Erreur"
        GoTo ErrorHandler
    End If

    MsgBox "Interfaces collaborateurs creees avec succes.", vbInformation, "Creation terminee"
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

    If MsgBox("Voulez-vous FORCER la suppression des interfaces RM ?" & vbCrLf & _
              "(Cela supprimera toutes les interfaces generees)", _
              vbYesNo + vbQuestion, "Confirmation de suppression forcee") = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    archiveChoice = MsgBox("Voulez-vous ARCHIVER les interfaces supprimees ?", _
                           vbYesNo + vbQuestion, "Confirmation d'archivage")

    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    deleteCommand = PYTHONEXE & "--basedir " & """" & baseDir & """" & " delete --force"
    If archiveChoice = vbYes Then
        deleteCommand = deleteCommand & " --archive"
        Application.StatusBar = "Archivage et suppression des interfaces..."
    Else
        Application.StatusBar = "Suppression des interfaces..."
    End If

    exitCode = RunCommand(deleteCommand)
    Application.StatusBar = False

    If exitCode <> 0 Then
        MsgBox "Erreur lors de la suppression des interfaces. Code de sortie : " & exitCode, vbCritical, "Erreur"
        GoTo ErrorHandler
    End If

    MsgBox IIf(archiveChoice = vbYes, _
               "Interfaces archivees et supprimees avec succes.", _
               "Interfaces supprimees avec succes."), vbInformation, "Suppression terminee"

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

    If MsgBox("Voulez-vous lancer le nettoyage des collaborateurs manquants ?" & vbCrLf & _
              "Cette action supprimera les fichiers d'interface des collaborateurs absents de la liste actuelle.", _
              vbYesNo + vbQuestion, "Confirmation du nettoyage") = vbNo Then Exit Sub

    baseDir = GetBaseDir()
    If baseDir = "" Then Exit Sub

    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    If Not CreateCollabsXML(baseDir) Then
        MsgBox "Erreur lors de la creation du fichier collabs.xml. Operation annulee.", vbCritical, "Erreur"
        GoTo ErrorHandler
    End If

    Application.StatusBar = "Nettoyage des interfaces collaborateurs manquantes..."
    exitCode = RunCommand(PYTHONEXE & "--basedir " & """" & baseDir & """" & " cleanup")
    Application.StatusBar = False

    If exitCode <> 0 Then
        MsgBox "Erreur pendant le nettoyage. Code de sortie : " & exitCode, vbCritical, "Erreur"
        GoTo ErrorHandler
    End If

    MsgBox "Nettoyage termine. Les interfaces des collaborateurs manquants ont ete supprimees.", vbInformation, "Nettoyage termine"
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub
