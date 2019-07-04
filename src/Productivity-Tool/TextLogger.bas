Attribute VB_Name = "TextLogger"
Option Explicit

Public Function WriteLogToFile(ByVal StrLogText As String)
    Dim objFSO As IWshRuntimeLibrary.FileSystemObject
    Dim objText As IWshRuntimeLibrary.TextStream
    
    If DEBUG_LOGGING Then
    
        Set objFSO = New IWshRuntimeLibrary.FileSystemObject
    
        If LogExists Then
            Set objText = objFSO.OpenTextFile(TEXT_LOG_LOC & GetUserName, ForAppending)
            objText.WriteLine (StrLogText)
            objText.Close
        End If
    End If

End Function

Private Function LogExists() As Boolean
    Dim objFSO As IWshRuntimeLibrary.FileSystemObject
    Dim rtnVal As Boolean
    
    rtnVal = False
    
    Set objFSO = New IWshRuntimeLibrary.FileSystemObject
    
    If objFSO.FileExists(TEXT_LOG_LOC & GetUserName) Then
        rtnVal = True
    Else:
        CreateLog
        rtnVal = True
    End If
        
    LogExists = rtnVal

End Function

Private Function CreateLog()
    Dim objFSO As IWshRuntimeLibrary.FileSystemObject
    
    Set objFSO = New IWshRuntimeLibrary.FileSystemObject
    
    If Not objFSO.FolderExists(TEXT_LOG_LOC) Then
        objFSO.CreateFolder (TEXT_LOG_LOC)
    End If
    
    objFSO.CreateTextFile TEXT_LOG_LOC & GetUserName, False
    
    
End Function
