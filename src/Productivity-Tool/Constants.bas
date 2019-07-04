Attribute VB_Name = "Constants"
Option Explicit

Public Const DEBUG_MODE As Boolean = True

Public Const DB_NAME As String = "ProductivityToolDatabase.accdb"

Public Const DEBUG_LOGGING As Boolean = False

Public Const DB_PASSWORD As String = "PT@WORK"

Public Const EMPTY_DATE As Date = "12:00:00 AM"

Public RefTableMng As ReferenceTableManager

Public ControlGroupManager As ControlGroupManagerClass

'Location functions
Public Function DEV_EXPORT_LOCATION() As String
    If Not DEBUG_MODE Then
        'Do nothing test only
    Else:
        DEV_EXPORT_LOCATION = Left(ThisWorkbook.Path, InStrRev(ThisWorkbook.Path, "\") - 1) & "\src\Productivity-Tool\"
    End If
End Function

Public Function DB_LOC() As String
    If Not DEBUG_MODE Then
        DB_LOC = ""
    Else:
        DB_LOC = ThisWorkbook.Path & "\"
    End If
End Function

Public Function DB_BACKUP_LOC() As String
    If Not DEBUG_MODE Then
        DB_BACKUP_LOC = ""
    Else:
        DB_BACKUP_LOC = ThisWorkbook.Path & "\Backup\"
    End If
End Function

Public Function TEXT_LOG_LOC() As String
    If Not DEBUG_MODE Then
        TEXT_LOG_LOC = ""
    Else:
        TEXT_LOG_LOC = ThisWorkbook.Path & "\Debug Log\"
    End If
End Function
