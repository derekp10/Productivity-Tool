Attribute VB_Name = "Utilities"
Option Explicit

Public Function GetUserName() As String
    Dim objNet As IWshRuntimeLibrary.WshNetwork
    Dim rtnString As String
    
    Set objNet = New IWshRuntimeLibrary.WshNetwork
    
    rtnString = objNet.UserName
    
    GetUserName = rtnString
End Function

'Populates ComboBox with Reftable data
Public Function PopulateRefDataToComboBox(ByRef TargetCombobox As MSForms.ComboBox, TargetRefTableCollection As Collection, Optional BlankStarter As Boolean = False, Optional DisplayAll As Boolean = False)
    Dim arrRefTableData()
    Dim lCntr As Long
    Dim RefTDataCls As RefTableDataClass
    Dim usableEntryCount As Long
    
    TargetCombobox.Clear
    
    If DisplayAll Then
        usableEntryCount = TargetRefTableCollection.Count
    Else:
        For Each RefTDataCls In TargetRefTableCollection
            If RefTDataCls.RefTypeExtraData.GetValueForFieldName("Disabled") = False Then
                usableEntryCount = usableEntryCount + 1
            End If
        Next RefTDataCls
    End If
    
    If BlankStarter Then
        ReDim arrRefTableData(usableEntryCount, 1)
    
        lCntr = 0
        
        arrRefTableData(lCntr, 0) = 0
        arrRefTableData(lCntr, 1) = ""
        lCntr = lCntr + 1
    Else:
        ReDim arrRefTableData(usableEntryCount - 1, 1)
        lCntr = 0
    End If
    
    For Each RefTDataCls In TargetRefTableCollection
        If DisplayAll Then
            arrRefTableData(lCntr, 0) = RefTDataCls.RefTypeID
            arrRefTableData(lCntr, 1) = RefTDataCls.RefTypeName
            lCntr = lCntr + 1
        Else:
            If RefTDataCls.RefTypeExtraData.GetValueForFieldName("Disabled") = False Then
                arrRefTableData(lCntr, 0) = RefTDataCls.RefTypeID
                arrRefTableData(lCntr, 1) = RefTDataCls.RefTypeName
                lCntr = lCntr + 1
            End If
        End If
    Next RefTDataCls
    
    TargetCombobox.List = arrRefTableData
    
End Function

Public Function WriteJobDataToDB(ByRef JobData As JobsDataClass)
    'On Error GoTo ADO_ACCESS_ERROR
    Dim DBCon As ADODB.Connection
    Dim dbRS As ADODB.Recordset
    Dim strSQL As String
    Dim lngLockErrCnt As Long
    Dim tryCnt As Long
    
'    Set DBCon = GetDBCon
'
'    DBCon.Open (DB_LOC & DB_NAME)
    
'    Set dbRS = GetDBRS(DBCon)

    
'    If JobData.JobID <> 0 Then
'        dbRS.Source = "Select * from Jobs Where JobID = " & JobData.JobID
        strSQL = "Select * from Jobs Where JobID = " & JobData.JobID
'    Else:
'        dbRS.Source = "Select * from Jobs"
'    End If
    
'    dbRS.Open

    CombinedConRecordSetPrep dbRS, strSQL, adOpenKeyset, adLockOptimistic
    
    If dbRS.RecordCount = 0 Then
        dbRS.AddNew
        dbRS.Fields("JobCreationDateTime").Value = JobData.JobCreationDateTime
        dbRS.Fields("StartDateTime").Value = JobData.StartDateTime
        dbRS.Fields("LastCountUpdate").Value = JobData.LastCountUpdate
        dbRS.Fields("JFID").Value = JobData.JFID
        dbRS.Fields("UserID").Value = JobData.UserID
        dbRS.Fields("StatusID").Value = JobData.StatusID
        dbRS.Fields("JobCount").Value = JobData.JobCount
        If JobData.ExternalMod <> vbNullString Then
            dbRS.Fields("ExternalMod").Value = JobData.ExternalMod
        End If
        dbRS.Update
        JobData.JobID = dbRS.Fields("JobID").Value
    Else:
        dbRS.Fields("StartDateTime").Value = JobData.StartDateTime
        dbRS.Fields("LastCountUpdate").Value = JobData.LastCountUpdate
        dbRS.Fields("JFID").Value = JobData.JFID
        dbRS.Fields("UserID").Value = JobData.UserID
        dbRS.Fields("StatusID").Value = JobData.StatusID
        If JobData.JobCountDelta <> 0 Then
            Debug.Print (DateTime.Now & " WriteJobToDB: JobID: " & JobData.JobID & " JFID: " & JobData.JFID & " DBVal: " & dbRS.Fields("JobCount").Value & ", JobDelta: " & JobData.JobCountDelta)
            WriteLogToFile (DateTime.Now & " WriteJobToDB: JobID: " & JobData.JobID & " JFID: " & JobData.JFID & " DBVal: " & dbRS.Fields("JobCount").Value & ", JobDelta: " & JobData.JobCountDelta)
            dbRS.Fields("JobCount").Value = dbRS.Fields("JobCount").Value - JobData.JobCountDelta
            JobData.JobCountDelta = 0
        Else:
            'Not sure this makes since to do this.
            'dbRS.Fields("JobCount").Value = JobData.JobCount
        End If
        dbRS.Update
        
    End If
    
CLEAN_EXIT:
    If dbRS.State = adStateOpen Then
        dbRS.Close
    End If
    
'    DBCon.Close
    
    Set dbRS = Nothing
'    Set DBCon = Nothing
    
    Exit Function

ADO_ACCESS_ERROR:
    Select Case Err.Number
        Case -2147467259
            lngLockErrCnt = lngLockErrCnt + 1
            'For tryCnt = 1 To 100: Next tryCnt
            Debug.Print "RecordLock Attempt: " & lngLockErrCnt
            WriteLogToFile (DateTime.Now & " RecordLock Attempt: " & lngLockErrCnt)
            Resume
        Case Else:
             MsgBox "Error: " & Err.Number & ": " & Err.Description
             GoTo CLEAN_EXIT
    End Select
    
End Function

Public Function GetJobDataFromDB(ByRef JobID As Long) As JobsDataClass
    Dim DBCon As ADODB.Connection
    Dim dbRS As ADODB.Recordset
    Dim rtnJob As JobsDataClass
    
    Set DBCon = GetDBCon
    
    DBCon.Open (DB_LOC & DB_NAME)
    
    Set dbRS = GetDBRS(DBCon)
    
    dbRS.Source = "Select * from Jobs Where JobID = " & JobID

    
    dbRS.Open
    
    Set rtnJob = New JobsDataClass
    
    If dbRS.RecordCount <> 0 Then
        rtnJob.JobID = dbRS.Fields("JobID").Value
        
        rtnJob.JobCreationDateTime = dbRS.Fields("JobCreationDateTime").Value
        rtnJob.StartDateTime = dbRS.Fields("StartDateTime").Value
        rtnJob.LastCountUpdate = dbRS.Fields("LastCountUpdate").Value
        rtnJob.JFID = dbRS.Fields("JFID").Value
        rtnJob.UserID = dbRS.Fields("UserID").Value
        rtnJob.StatusID = dbRS.Fields("StatusID").Value
        rtnJob.JobCount = dbRS.Fields("JobCount").Value
        
    End If
    
    dbRS.Close
    
    DBCon.Close
    
    Set dbRS = Nothing
    Set DBCon = Nothing
    
    Set GetJobDataFromDB = rtnJob
    
End Function

Public Function GetJobColFromDB(ByRef UserID As Long) As Collection
    Dim DBCon As ADODB.Connection
    Dim dbRS As ADODB.Recordset
    Dim rtnJob As JobsDataClass
    Dim rtnCol As Collection
    
    Set rtnCol = New Collection
    
    Set DBCon = GetDBCon
    
    DBCon.Open (DB_LOC & DB_NAME)
    
    Set dbRS = GetDBRS(DBCon)
    
'    dbRS.Source = "Select * from Jobs Where UserID = " & UserID & " and StartDateTime >= #" & DateAdd("d", -5, DateTime.Now) & "# And StartDateTime <= #" & DateAdd("h", -12, DateTime.Now) & "# Order by StartDateTime Desc"
    dbRS.Source = "Select * from Jobs Where UserID = " & UserID & " and StartDateTime >= #" & DateAdd("d", -5, DateTime.Now) & "# Order by StartDateTime Desc"

    dbRS.Open
    
    If dbRS.RecordCount <> 0 Then
        Do Until dbRS.EOF
            Set rtnJob = New JobsDataClass
            
            rtnJob.JobID = dbRS.Fields("JobID").Value
            
            rtnJob.JobCreationDateTime = dbRS.Fields("JobCreationDateTime").Value
            rtnJob.StartDateTime = dbRS.Fields("StartDateTime").Value
            rtnJob.LastCountUpdate = dbRS.Fields("LastCountUpdate").Value
            rtnJob.JFID = dbRS.Fields("JFID").Value
            rtnJob.UserID = dbRS.Fields("UserID").Value
            rtnJob.StatusID = dbRS.Fields("StatusID").Value
            rtnJob.JobCount = dbRS.Fields("JobCount").Value
            
            rtnCol.Add rtnJob, CStr(rtnJob.JobID)
            
            dbRS.MoveNext
        Loop
        
    End If
    
    dbRS.Close
    
    DBCon.Close
    
    Set dbRS = Nothing
    Set DBCon = Nothing
    
    Set GetJobColFromDB = rtnCol
    
End Function

Public Function ClearExternalModStatus(ByVal ModJobID As JobsDataClass)
    Dim DBCon As ADODB.Connection
    Dim dbRS As ADODB.Recordset
    
    Set DBCon = GetDBCon
    
    DBCon.Open (DB_LOC & DB_NAME)
    
    Set dbRS = GetDBRS(DBCon)
    
    dbRS.Source = "Select * from Jobs Where JobID = " & ModJobID.JobID

    dbRS.Open
    
    If dbRS.Fields("ExternalMod").Value = ModJobID.ExternalMod Then
        dbRS.Fields("ExternalMod").Value = Null
        dbRS.Update
    End If
    
    dbRS.Close
    
    DBCon.Close
    
    Set dbRS = Nothing
    Set DBCon = Nothing
End Function

Public Function GetTodaysJobsFromDB(ByRef UserID As Long, Optional ExtModOnly As Boolean = False) As Collection
    Dim DBCon As ADODB.Connection
    Dim dbRS As ADODB.Recordset
    Dim rtnJob As JobsDataClass
    Dim rtnCol As Collection
    
    Set rtnCol = New Collection
    
    Set DBCon = GetDBCon
    
    DBCon.Open (DB_LOC & DB_NAME)
    
    Set dbRS = GetDBRS(DBCon)
    
    If ExtModOnly Then
        dbRS.Source = "Select * from Jobs Where UserID = " & UserID & " and ExternalMod <> Null and StartDateTime >= #" & DateAdd("h", -12, DateTime.Now) & "# Order By JobID"
    Else:
        dbRS.Source = "Select * from Jobs Where UserID = " & UserID & " and StartDateTime >= #" & DateAdd("h", -12, DateTime.Now) & "# Order By JobID"
    End If

    dbRS.Open
    
    If dbRS.RecordCount <> 0 Then
        Do Until dbRS.EOF
            Set rtnJob = New JobsDataClass
            
            rtnJob.JobID = dbRS.Fields("JobID").Value
            
            rtnJob.JobCreationDateTime = dbRS.Fields("JobCreationDateTime").Value
            rtnJob.StartDateTime = dbRS.Fields("StartDateTime").Value
            rtnJob.LastCountUpdate = dbRS.Fields("LastCountUpdate").Value
            rtnJob.JFID = dbRS.Fields("JFID").Value
            rtnJob.UserID = dbRS.Fields("UserID").Value
            rtnJob.StatusID = dbRS.Fields("StatusID").Value
            rtnJob.JobCount = dbRS.Fields("JobCount").Value
            
            If ExtModOnly Then
                rtnJob.ExternalMod = dbRS.Fields("ExternalMod").Value
            End If
            
            rtnCol.Add rtnJob, CStr(rtnJob.JobID)
            
            dbRS.MoveNext
        Loop
        
    End If
    
    dbRS.Close
    
    DBCon.Close
    
    Set dbRS = Nothing
    Set DBCon = Nothing
    
    Set GetTodaysJobsFromDB = rtnCol
End Function

Public Function WriteModificationDataToDB(ByRef ModData As ModificationDataClass)
    Dim DBCon As ADODB.Connection
    Dim dbRS As ADODB.Recordset
    
    Set DBCon = GetDBCon
    
    DBCon.Open (DB_LOC & DB_NAME)
    
    Set dbRS = GetDBRS(DBCon)

    dbRS.Source = "Select * from Modifications"

    dbRS.Open
    
    dbRS.AddNew
    dbRS.Fields("JobID").Value = ModData.JobID
    dbRS.Fields("UserID").Value = ModData.UserID
    dbRS.Fields("ModDateTime").Value = ModData.ModDateTime
    dbRS.Fields("ModElmID").Value = ModData.ModElmID
    dbRS.Fields("NewValue").Value = ModData.NewValue
    dbRS.Fields("OldValue").Value = ModData.OldValue
    dbRS.Update

    
    dbRS.Close
    
    DBCon.Close
    
    Set dbRS = Nothing
    Set DBCon = Nothing
    
    
End Function

Public Function WriteInterruptionDataToDB(ByRef IntrData As InterruptionDataClass)
    Dim DBCon As ADODB.Connection
    Dim dbRS As ADODB.Recordset
    
    Set DBCon = GetDBCon
    
    DBCon.Open (DB_LOC & DB_NAME)
    
    Set dbRS = GetDBRS(DBCon)

    dbRS.Source = "Select * from Interruptions"

    dbRS.Open
    
    dbRS.AddNew
    dbRS.Fields("UserID").Value = IntrData.UserID
    dbRS.Fields("InterruptDateTime").Value = IntrData.InterruptDateTime
    dbRS.Fields("InterruptTypeID").Value = IntrData.InterruptTypeID
    dbRS.Fields("InterruptionLength").Value = IntrData.InterruptionLength
    dbRS.Update

    
    dbRS.Close
    
    DBCon.Close
    
    Set dbRS = Nothing
    Set DBCon = Nothing
    
    
End Function

Public Function GetExternalUpdates(ByVal UserID As Long) As Collection
    Dim dbRS As ADODB.Recordset
    Dim strSQL As String
    Dim rtnCol As Collection
    Dim ExternalUpdateData As ExternalUpdateDataClass
    
    Set rtnCol = New Collection
    
    strSQL = "Select * from ExternalUpdates where UserID = " & UserID & " Order By CreationDateTime Asc"
    
    CombinedConRecordSetPrep dbRS, strSQL, adOpenKeyset, adLockReadOnly
    
    If dbRS.RecordCount <> 0 Then
        Do Until dbRS.EOF
            Set ExternalUpdateData = New ExternalUpdateDataClass
            
            ExternalUpdateData.UpdateID = dbRS.Fields("UpdateID").Value
            ExternalUpdateData.JobID = dbRS.Fields("JobID").Value
            ExternalUpdateData.CreationDateTime = dbRS.Fields("CreationDateTime").Value
            ExternalUpdateData.JFID = dbRS.Fields("JFID").Value
            ExternalUpdateData.UserID = dbRS.Fields("UserID").Value
            ExternalUpdateData.UpdateAmount = dbRS.Fields("UpdateAmount").Value
            
            rtnCol.Add ExternalUpdateData, CStr(ExternalUpdateData.UpdateID)
            
            dbRS.MoveNext
        Loop
    End If
    
    Set GetExternalUpdates = rtnCol
End Function

Public Function InCollection(colCollection As Collection, ByVal strItemToCheck As String) As Boolean
    'Compares value provided to collection provided key's refrence. On error (no item found) returns false, else true
    ' a.k.a Key found.
    On Error GoTo HandleError
    'On Error Resume Next
    Dim var As Variant
    var = colCollection(strItemToCheck)
    InCollection = True
    Exit Function
HandleError:
    If Err.Number <> 438 Then
        InCollection = False
    Else:
        InCollection = True
    End If
End Function

Public Function InControls(ControlsGoup As Controls, ByVal strItemToCheck As String) As Boolean
    'This is a rif on InCollection, only designed to work with the Controls object in MSForm.Controls
    'Compares value provided to collection provided key's refrence. On error (no item found) returns false, else true
    ' a.k.a Key found.
    On Error GoTo HandleError
    'On Error Resume Next
    Dim var As Variant
    var = ControlsGoup(strItemToCheck)
    InControls = True
    Exit Function
HandleError:
    If Err.Number <> 438 Then
        InControls = False
    Else:
        InControls = True
    End If
End Function

Public Function InFields(colCollection As ADODB.Fields, ByVal strItemToCheck As String) As Boolean
    'This is a rif on InCollection, only designed to work with the Fields object in ADODB Recordsets
    'Need to watch the 438 Error check as it may cause issues with the "Fields" Collection
    'Compares value provided to collection provided key's refrence. On error (no item found) returns false, else true
    ' a.k.a Key found.
    On Error GoTo HandleError
    'On Error Resume Next
    Dim var As Variant
    var = colCollection(strItemToCheck).Name
    InFields = True
    Exit Function
HandleError:
    If Err.Number <> 438 Then
        InFields = False
    Else:
        InFields = True
    End If
End Function

Public Function LastRefTableUpdate(Optional ByVal UpdateDate As Boolean = False) As Date
    Dim DBCon As ADODB.Connection
    Dim dbRS As ADODB.Recordset
    Dim rtnDateTime As Date
    
    Set DBCon = GetDBCon
    DBCon.Open DB_LOC & DB_NAME
    
    Set dbRS = GetDBRS(DBCon)
    
    dbRS.Source = "Select * from ApplicationDataStore where ApplicationElement = 'LastRefTableUpdate'"
    
    dbRS.Open
    
    If UpdateDate = False Then
        If Not IsNull(dbRS.Fields("ElementDataStore").Value) Then
            rtnDateTime = dbRS.Fields("ElementDataStore").Value
        Else:
            rtnDateTime = DateTime.Now
        End If
    Else:
        dbRS.Fields("ElementDataStore").Value = DateTime.Now
        dbRS.Update
        rtnDateTime = dbRS.Fields("ElementDataStore").Value
    End If
        
    dbRS.Close
    DBCon.Close
    
    Set dbRS = Nothing
    Set DBCon = Nothing
    
    LastRefTableUpdate = rtnDateTime

End Function

Public Function LastDatabaseBackup(Optional ByVal UpdateDate As Boolean = False) As Date
    Dim DBCon As ADODB.Connection
    Dim dbRS As ADODB.Recordset
    Dim rtnDateTime As Date
    
    Set DBCon = GetDBCon
    DBCon.Open DB_LOC & DB_NAME
    
    Set dbRS = GetDBRS(DBCon)
    
    dbRS.Source = "Select * from ApplicationDataStore where ApplicationElement = 'LastDBBackup'"
    
    dbRS.Open
    
    If UpdateDate = False Then
        If Not IsNull(dbRS.Fields("ElementDataStore").Value) Then
            rtnDateTime = dbRS.Fields("ElementDataStore").Value
        Else:
            rtnDateTime = DateTime.Now
        End If
    Else:
        dbRS.Fields("ElementDataStore").Value = DateTime.Now
        dbRS.Update
        rtnDateTime = dbRS.Fields("ElementDataStore").Value
    End If
        
    dbRS.Close
    DBCon.Close
    
    Set dbRS = Nothing
    Set DBCon = Nothing
    
    LastDatabaseBackup = rtnDateTime

End Function

Public Function BackupDatabase()
    Dim DBBackup As DBCompactAndBackupClass
    Dim lastBackup As Date
    
    Set DBBackup = New DBCompactAndBackupClass
    
    lastBackup = LastDatabaseBackup
    
    If DateDiff("h", lastBackup, DateTime.Now) > 3 Then
        DBBackup.BackupDatabase DB_BACKUP_LOC, DB_LOC, DB_NAME
        LastDatabaseBackup True
    End If
    
End Function

Private Sub UpdateTimeAmount()
    Debug.Print ("UpdateTimeAmmount Call")
    InterruptionUserForm.UpdateTimeAmount
End Sub

Private Sub CheckForJobUpdates()
    Debug.Print ("CheckForJobUpdates Call")
    WriteLogToFile (DateTime.Now & " CheckForJobUpdates Call")
    'TrackUserForm.ScheduledCheckForDataUpdates
    ExternalUpdateManager.ScheduledCheckForDataUpdates
End Sub

Private Sub ControlGroupTimerCall(ByVal JobID As String)
    Debug.Print ("ControlGroupTimerCall Call JobID: " & JobID)
    WriteLogToFile (DateTime.Now & " ControlGroupTimerCall Call JobID: " & JobID)
    Dim tmpJD As JobsDataClass
    Set tmpJD = ControlGroupManager.JobDataCollection(JobID)
    tmpJD.ControlGroup.ClearUserModifyFlag
End Sub

Public Function GetNonDisabledRefTableDataForSearchFunction(ByRef TargetRefCol As Collection) As Collection
    'Only use with ComboBox generation code. Defaults to rtkt_RefTypeRefName
    Dim tmpCol As Collection
    Dim tmpRTDC As RefTableDataClass
    
    Set tmpCol = New Collection
    
    For Each tmpRTDC In TargetRefCol
        If tmpRTDC.RefTypeExtraData.GetValueForFieldName("Disabled") = False Then
            tmpCol.Add tmpRTDC, CStr(tmpRTDC.RefTypeName)
        End If
    Next tmpRTDC
    
    Set GetNonDisabledRefTableDataForSearchFunction = tmpCol
    
End Function

Public Function JobFunctionExistsInTodaysList(ByVal JobFunctionID As Long) As Boolean
    Dim colTodaysJobs As Collection
    Dim JobData As JobsDataClass
    Dim blnFound As Boolean
    
    Set colTodaysJobs = GetTodaysJobsFromDB(RefTableMng.GetIDFromString(rte_Users, GetUserName))
    
    For Each JobData In colTodaysJobs
        If JobData.JFID = JobFunctionID Then
            blnFound = True
            Exit For
        End If
    Next JobData
    
    JobFunctionExistsInTodaysList = blnFound
    
End Function

Public Function GetJobDataForJobFunctionID(ByRef JobCollection As Collection, JobFunctionID As Long) As JobsDataClass
    Dim tmpJobData As JobsDataClass
    Dim rtnJobData As JobsDataClass
    
'    Set rtnJobData = New JobsDataClass
'
'    rtnJobData.JobID = -1
    
    For Each tmpJobData In JobCollection
        If tmpJobData.JFID = JobFunctionID Then
            Set rtnJobData = tmpJobData
            Exit For
        End If
    Next tmpJobData
    

    Set GetJobDataForJobFunctionID = rtnJobData
End Function

Public Function GetPossiblePastJobsForUpdateData(ByRef UpdateData As ExternalUpdateDataClass) As Collection
    Dim dbRS As ADODB.Recordset
    Dim strSQL As String
    Dim rtnCol As Collection
    Dim tmpJobData As JobsDataClass
    
    strSQL = "Select * from Jobs Where UserID = " & UpdateData.UserID & " and JFID = " & UpdateData.JFID & " And StartDateTime Between #" & UpdateData.DatePast & "# And #" & UpdateData.DateFuture & "#"
    
    CombinedConRecordSetPrep dbRS, strSQL, adOpenKeyset, adLockReadOnly
    
    Set rtnCol = New Collection
    
    Do Until dbRS.EOF
        Set tmpJobData = GetJobDataFromDB(dbRS.Fields("JobID").Value)
        
        rtnCol.Add tmpJobData, CStr(tmpJobData.JobID)
        
        dbRS.MoveNext
    Loop
    
    
    dbRS.Close
    
    Set dbRS = Nothing
    
    Set GetPossiblePastJobsForUpdateData = rtnCol
    
End Function

Public Function RemoveProcessedUpdate(ByRef ExternalUpdateData As ExternalUpdateDataClass)
    'On Error GoTo ADO_ACCESS_ERROR
    Dim dbRS As ADODB.Recordset
    Dim strSQL As String
    Dim lngLockErrCnt As Long
    Dim tryCnt As Long
    
    strSQL = "Select * from ExternalUpdates where UpdateID = " & ExternalUpdateData.UpdateID
    
    CombinedConRecordSetPrep dbRS, strSQL, adOpenKeyset, adLockPessimistic
    
    If dbRS.RecordCount = 1 Then
        dbRS.Delete
        dbRS.Update
    End If
    
CLEAN_EXIT:
    If dbRS.State = adStateOpen Then
        dbRS.Close
    End If
    
    Set dbRS = Nothing
    
    Exit Function
    
ADO_ACCESS_ERROR:
    Select Case Err.Number
        Case -2147467259
            lngLockErrCnt = lngLockErrCnt + 1
            'For tryCnt = 1 To 100: Next tryCnt
            Debug.Print "RecordLock Attempt: " & lngLockErrCnt
            WriteLogToFile (DateTime.Now & "RecordLock Attempt: " & lngLockErrCnt)
            Resume
        Case Else:
             MsgBox "Error: " & Err.Number & ": " & Err.Description
             GoTo CLEAN_EXIT
    End Select
    
End Function
