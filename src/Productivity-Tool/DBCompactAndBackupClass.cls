VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBCompactAndBackupClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Backup and Compact/Repair database Class object for use in other projects.
'By Derek Pierce (10/3/2012)
'Ver. 1.2 6/12/2014
'
'Takes a Backup Location, Database Directory location, and Database Name as arguments to the BackupDatabase function,
'and automatically creates a backup folder if not currently created, and proceeds to backup then compact and repair the database
'using DAO.DBEngine Object.
'
'Also maintains backup folder removing backups that are older than xDaysold or folder count > XNumToKeep.
'
'A back up is run at the place in the code where you make the call. Each time it is done, it will create a new file in the
'backup folder. Running this multiple times in succession can cause your older backups to be removed depending on the settings
'used for xDaysOld or xNumToKeep. Any Backup calls run within the same minute will be overwritten until the date time rolls over to
'the next minute.


'Requires reference to "Microsoft Office Access 14.0 Database Engine Object Library"
Private dbEng As New DAO.DBEngine

'Requires reference to "Windows Scripting Host Object Model"
Private srcFolder As Folder
Private srcFile As file
Private p_FSO As New FileSystemObject

'class variables
Private xDaysOld As Long
Private xNumToKeep As Long

Private Sub Class_Initialize()
'xDaysOld = How many days old should the oldest backup be.
'xNumToKeep = How many total backup database files to keep.
'
' -1 value assigned to either will cause the respective folder cleanup to not be done.
'IE. To disable removal of old backups based on days, xDaysOld = -1
'To disable the restriction on the number of files to keep xNumToKeep = -1
' -1 to both will keep all backups.
'
'Be advised, if both are disabled, and then turned on at a later date, it will automatically apply those settings and clean up the folder
'to match the new values. Removing Files counts > xNumToKeep, or anything older than xDaysOld.
'
'All removals are targeted at the oldest file available in the folder.

    xDaysOld = 7
    xNumToKeep = -1
    
End Sub

Public Function CompactAndRepair(ByVal TargetDatabaseLocation As String, TargetDatabase As String)
'Compacts and repairs with out using the backup functionality
    Dim compactStatus As Long
    compactStatus = 0
    
    compactStatus = compactDB(TargetDatabaseLocation, TargetDatabase)
    If compactStatus <> 0 Then
        If compactStatus = "3704" Then
            'Error for database file currently open directly.
            MsgBox ("Database file appears to be open directly by a user. Compact & Repair will not be performed.")
            Exit Function
        End If
        
        'If unknown error, Display and exit function
        MsgBox ("Database compact failed: " & compactStatus & Chr(13) & Error(compactStatus) & Chr(13) & "Compact & Repair will not be performed.")
        Exit Function
    End If

    
End Function


Public Function BackupDatabase(ByVal BackupFolderLocation As String, TargetDatabaseLocation As String, TargetDatabase As String) As Long
    'Returns a Long that is either 0 or the error number if an error occured. Use to check if backup failed.
    Dim compactStatus As Long

    'Checks for backup folder
    If Not p_FSO.FolderExists(BackupFolderLocation) Then
        p_FSO.CreateFolder (BackupFolderLocation)
    End If
    
    'checkes to make sure backup folder exists again, incause first check failed and folder was just created.
    If p_FSO.FolderExists(BackupFolderLocation) Then
        compactStatus = 0
        compactStatus = backupAndCR(BackupFolderLocation, TargetDatabaseLocation, TargetDatabase)
        If compactStatus <> 0 Then
            If compactStatus = "3704" Then
                'Error for database file currently open directly.
                MsgBox ("Database file appears to be open directly by a user. Compact function will be skipped.")
                BackupDatabase = compactStatus
                Exit Function
            End If
            
            'If unknown error, Display and exit function
            MsgBox ("Database compact failed: " & compactStatus & Chr(13) & Error(compactStatus) & Chr(13) & "Backup functions will be skipped.")
            BackupDatabase = compactStatus
            Exit Function
        Else:
            'Clean up backup folder
            cleanUpDailyBackupFolder BackupFolderLocation
        End If
    Else:
        MsgBox ("Unable to Locate/Create backup folder. Backup functions will be skipped.")
    End If
    
End Function

Private Function backupAndCR(ByVal TargetBackupLoc As String, TargetLoc As String, TargetDB As String) As Long
    Dim fld As Folder
    Dim compactStatus As Long
    
    'Create backup file in backup folder
    p_FSO.CopyFile TargetLoc & TargetDB, TargetBackupLoc & dbBackupName(TargetDB)
    
    'Compact database
    'compactStatus = compactDB(TargetLoc, TargetDB)
    
    'return error code if any occur
    backupAndCR = compactStatus

End Function

Private Function compactDB(ByVal TargetLoc As String, TargetDB As String) As Long

On Error GoTo ErrFailed
        'Remove any previously generated .tmp files incase compact failed last run
        If p_FSO.FileExists(TargetLoc & TargetDB & ".tmp") Then
            p_FSO.DeleteFile (TargetLoc & TargetDB & ".tmp")
        End If
        
        'Create new Compacted database file with .tmp at the end of the name
        dbEng.CompactDatabase TargetLoc & TargetDB, TargetLoc & TargetDB & ".tmp"
        
    On Error GoTo 0 ' turn off error handling
        ' Delete old Database file (Uncompacted version)
        p_FSO.DeleteFile (TargetLoc & TargetDB)
        'Change filename of tmp/compacted database to original name
        Name TargetLoc & TargetDB & ".tmp" As TargetLoc & TargetDB
        
ErrFailed:
    'Return compact error if any
    compactDB = Err.Number 'handle error number
End Function

Private Function dbBackupName(ByVal dbName As String) As String
    'Find the base name and the extention name to modify and add date the backup
    'was done on the file.
    
    Dim dbStripName As String
    Dim dbExt As String
    
    
    dbStripName = p_FSO.GetBaseName(dbName)
    dbExt = p_FSO.GetExtensionName(dbName)
    
    dbBackupName = dbStripName & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & "." & dbExt
End Function

Private Function cleanUpDailyBackupFolder(strDBLoc As String)
    'Based on xDaysOld and xNumToKeep remove the oldest database backup file in folder.
    'This compare is done on file creation date. Any files in there will be deleted.
    'So only store database backups in the backup folder.
    'This code skips if oldest backup returns "Nothing" meaning that the folder is empty.
    
    If Not oldestBackUp(strDBLoc) Is Nothing Then
        If xDaysOld <> -1 Then
            Do While DateDiff("d", oldestBackUp(strDBLoc).DateCreated, Date) > xDaysOld
                p_FSO.DeleteFile (strDBLoc & oldestBackUp(strDBLoc).Name)
            Loop
        ElseIf xNumToKeep <> -1 Then
            Do While p_FSO.GetFolder(strDBLoc).Files.Count >= xNumToKeep
                p_FSO.DeleteFile (strDBLoc & oldestBackUp(strDBLoc).Name)
            Loop
        End If
    End If
    
End Function

Private Function oldestBackUp(strDBLoc As String) As file
    'Looks for the oldest file in the folder and returns it as type File
    'Returns "Nothing" if no files are present in folder.
    Dim oldestFile As file
    
    For Each srcFile In p_FSO.GetFolder(strDBLoc).Files
        If oldestFile Is Nothing Then
            Set oldestFile = srcFile
        Else:
            If oldestFile.DateCreated > srcFile.DateCreated Then
                Set oldestFile = srcFile
            End If
        End If
    Next srcFile
    Set oldestBackUp = oldestFile
End Function
