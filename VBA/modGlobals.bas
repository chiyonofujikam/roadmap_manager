Option Explicit

' Global variables
Public GLOBAL_BASEDIR As String
Public PYTHONEXE As String

' Sheet names
Public Const SHEET_SYNTHESE As String = "SYNTHESE"
Public Const SHEET_LC As String = "LC"
Public Const SHEET_GESTION_INTERFACES As String = "Gestion_Interfaces"
Public Const SHEET_VERIF_COLLABORATEUR As String = "Vérif_Collaborateur"
Public Const SHEET_EXTRACT_MSP As String = "Extract_MSP"
Public Const SHEET_FICHIER_SYNTHESE As String = "Fichier de synthèse"

' SYNTHESE layout (columns A-J)
Public Const SYN_FIRST_DATA_ROW As Long = 3
Public Const SYN_COL_COLLAB As Long = 2     ' B: collaborator name
Public Const SYN_COL_WEEK As Long = 3       ' C: week code (SXXYY)
Public Const SYN_COL_E As Long = 5          ' E: LC lookup (split by "Sprint")
Public Const SYN_COL_F As Long = 6          ' F: LC lookup
Public Const SYN_COL_G As Long = 7          ' G: LC lookup
Public Const SYN_COL_H As Long = 8          ' H: filled from LC
Public Const SYN_COL_I As Long = 9          ' I: filled from LC
Public Const SYN_COL_HOURS As Long = 10     ' J: Heures passées

' Vérif_Collaborateur layout
Public Const VERIF_HEADER_ROW As Long = 5
Public Const VERIF_FIRST_COLLAB_ROW As Long = 6
Public Const VERIF_COL_COLLAB As Long = 4           ' D: collaborator names
Public Const VERIF_FIRST_WEEK_COL As Long = 5       ' E: first week header
Public Const PERCENTAGE_NONZEROS_ROW As Long = 4
Public Const PERCENTAGE_NONZEROS_COL As Long = 5

' LC layout
Public Const LC_FIRST_ROW As Long = 3
Public Const LC_COL_KEY As Long = 2          ' B
Public Const LC_COL_LIBELLE As Long = 3      ' C
Public Const LC_COL_FUNCTION As Long = 4     ' D

' LC lookup table (F:K)
Public Const LC_LOOKUP_FIRST_ROW As Long = 2
Public Const LC_LOOKUP_LAST_ROW As Long = 10000
Public Const LC_LOOKUP_COL_F As Long = 6
Public Const LC_LOOKUP_COL_G As Long = 7
Public Const LC_LOOKUP_COL_H As Long = 8
Public Const LC_LOOKUP_COL_I As Long = 9
Public Const LC_LOOKUP_COL_J As Long = 10
Public Const LC_LOOKUP_COL_K As Long = 11
Public Const LC_LOOKUP_KEY_DELIM As String = "|"

' Extract_MSP columns
Public Const SRC_COL_KEY As Long = 2         ' B
Public Const SRC_COL_LIBELLE As Long = 6     ' F
Public Const SRC_COL_FUNCTION As Long = 3    ' C
