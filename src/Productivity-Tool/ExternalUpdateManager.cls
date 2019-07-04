VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExternalUpdateManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Single Instance class

Public Event ExternalUpdateStart()
Public Event ExternalUpdateEnd()

Private TimeInterval As Variant
Public UFActive As Boolean

Public Function ScheduledCheckForDataUpdates()
    RaiseEvent ExternalUpdateStart

    'CheckForDataUpdates
    ProcessExternalUpdates
    
    If UFActive Then
        ScheduleNextTrigger
    End If
    
    RaiseEvent ExternalUpdateEnd
End Function

Public Sub ScheduleNextTrigger(Optional ByVal KeepTracking As Boolean = True)
    If KeepTracking Then
        TimeInterval = Now + TimeValue("00:00:10")
        Application.OnTime TimeInterval, "Utilities.CheckForJobUpdates"
    Else:
        'To help check for different erros that are not the expected one.
        On Error Resume Next
        Application.OnTime TimeInterval, "Utilities.CheckForJobUpdates", , False
        If Err.Number <> 0 Then
            Debug.Print (Err.Number & ": " & Err.Description & " 'ExternalUpdateManager'")
        End If
        On Error GoTo 0
        
        Set TimeInterval = Nothing
    End If
End Sub

Public Function ProcessExternalUpdates()
    Dim colUpdates As Collection
    Dim ExternalUpdateData As ExternalUpdateDataClass
    Dim InMemoryJobData As JobsDataClass
    Dim tmpJobData As JobsDataClass
    Dim lastJFID As Long
    
    Set colUpdates = GetExternalUpdates(RefTableMng.GetIDFromString(rte_Users, GetUserName))
    
    For Each ExternalUpdateData In colUpdates
        DoEvents
        
        If ExternalUpdateData.JobID <= 0 Then
            If JobFunctionExistsInTodaysList(ExternalUpdateData.JFID) Then
                Set InMemoryJobData = GetJobDataForJobFunctionID(ControlGroupManager.JobDataCollection, ExternalUpdateData.JFID)
                If ExternalUpdateData.IsJobWithinValidDateWindow(InMemoryJobData) Then
                    If Not InMemoryJobData.ControlGroup.UserModifyingData Then
                        UpdateInMemoryData ExternalUpdateData, InMemoryJobData
                    Else:
                        Debug.Print ("EXCEPT")
                        WriteLogToFile (DateTime.Now & " EXCEPT")
                    End If
                Else:
                    'Job isn't within time frame find another.
                    HandleNoMatchUpdate ExternalUpdateData
                    CheckForDataUpdates
                End If
            Else:
                'Possibly not for today's date or New entry
                HandleNoMatchUpdate ExternalUpdateData
                CheckForDataUpdates
            End If
                
        Else:
            If JobFunctionExistsInTodaysList(ExternalUpdateData.JFID) Then
                'Update job directly
                Set InMemoryJobData = GetJobDataForJobFunctionID(ControlGroupManager.JobDataCollection, ExternalUpdateData.JFID)
            End If
            
            If Not InMemoryJobData Is Nothing Then
                If InMemoryJobData.JobID = ExternalUpdateData.JobID Then
                    'Update
                    UpdateInMemoryData ExternalUpdateData, InMemoryJobData
                Else:
                    HandleNoMatchUpdate ExternalUpdateData
                    CheckForDataUpdates
                End If
            Else:
                'Update job from database directly
                Set tmpJobData = GetJobDataFromDB(ExternalUpdateData.JobID)
                
                Debug.Print (DateTime.Now & " ProcessExternalUpdatesPreUpdatePre: JobID: " & tmpJobData.JobID & ", JFID: " & tmpJobData.JFID & ", JobCount: " & tmpJobData.JobCount & ", JobCountDelta: " & tmpJobData.JobCountDelta)
                
                WriteLogToFile (DateTime.Now & " ProcessExternalUpdatesPreUpdatePre: JobID: " & tmpJobData.JobID & ", JFID: " & tmpJobData.JFID & ", JobCount: " & tmpJobData.JobCount & ", JobCountDelta: " & tmpJobData.JobCountDelta)
                
                tmpJobData.JobCount = tmpJobData.JobCount + ExternalUpdateData.UpdateAmount
                tmpJobData.JobCountDelta = tmpJobData.JobCountDelta - ExternalUpdateData.UpdateAmount
                
                Debug.Print (DateTime.Now & " ProcessExternalUpdatesPreUpdatePost: JobCount: " & tmpJobData.JobCount & ", JFID: " & tmpJobData.JFID & ", JobCountDelta: " & tmpJobData.JobCountDelta)
                
                WriteLogToFile (DateTime.Now & " ProcessExternalUpdatesPreUpdatePost: JobCount: " & tmpJobData.JobCount & ", JFID: " & tmpJobData.JFID & ", JobCountDelta: " & tmpJobData.JobCountDelta)
                
                If tmpJobData.LastCountUpdate < ExternalUpdateData.CreationDateTime Then
                    tmpJobData.LastCountUpdate = ExternalUpdateData.CreationDateTime
                End If
                
                WriteJobDataToDB tmpJobData
                
                RemoveProcessedUpdate ExternalUpdateData
                
             End If
            
        End If
        
    Next ExternalUpdateData
    
End Function

Private Function UpdateInMemoryData(ByRef ExternalUpdateData As ExternalUpdateDataClass, InMemoryJobData As JobsDataClass)
    InMemoryJobData.UpdateStart
    
    If InMemoryJobData.LastCountUpdate < ExternalUpdateData.CreationDateTime Then
        InMemoryJobData.LastCountUpdate = ExternalUpdateData.CreationDateTime
    End If
    
    InMemoryJobData.JobCount = InMemoryJobData.JobCount + ExternalUpdateData.UpdateAmount
    
    Debug.Print (DateTime.Now & " UpdateInMemoryData: JobID: " & InMemoryJobData.JobID & ", JFID: " & InMemoryJobData.JFID & ", InMemJobDelta: " & InMemoryJobData.JobCountDelta & ", ExtUpdAmt: " & ExternalUpdateData.UpdateAmount)
    
    WriteLogToFile (DateTime.Now & " UpdateInMemoryData: JobID: " & InMemoryJobData.JobID & ", JFID: " & InMemoryJobData.JFID & ", InMemJobDelta: " & InMemoryJobData.JobCountDelta & ", ExtUpdAmt: " & ExternalUpdateData.UpdateAmount)
    
    InMemoryJobData.JobCountDelta = InMemoryJobData.JobCountDelta - ExternalUpdateData.UpdateAmount
    
    WriteJobDataToDB InMemoryJobData
    
    RemoveProcessedUpdate ExternalUpdateData
    
    InMemoryJobData.UpdateComplete
End Function

Private Function HandleNoMatchUpdate(ByRef ExternalUpdateData As ExternalUpdateDataClass)
    Dim colPastJobs As Collection
    Dim PastJobData As JobsDataClass
    Dim NewJobData As JobsDataClass
        
    Set colPastJobs = GetPossiblePastJobsForUpdateData(ExternalUpdateData)
    
    If colPastJobs.Count = 0 Then
        'Add New
        Set NewJobData = New JobsDataClass
        
        NewJobData.JobCreationDateTime = ExternalUpdateData.CreationDateTime
        NewJobData.StartDateTime = ExternalUpdateData.CreationDateTime
        NewJobData.LastCountUpdate = ExternalUpdateData.CreationDateTime
        NewJobData.JFID = ExternalUpdateData.JFID
        NewJobData.UserID = ExternalUpdateData.UserID
        NewJobData.StatusID = 2 'Inactive
        NewJobData.JobCountDelta = 0 - ExternalUpdateData.UpdateAmount
        NewJobData.JobCount = ExternalUpdateData.UpdateAmount
        
        If NewJobData.StartDateTime >= DateAdd("h", -12, DateTime.Now) Then
            NewJobData.ExternalMod = DateTime.Now
        End If
        
        WriteJobDataToDB NewJobData
        
        RemoveProcessedUpdate ExternalUpdateData
        
    ElseIf colPastJobs.Count = 1 Then
        'How to deal with possible more than one match...
        For Each PastJobData In colPastJobs
            
            If PastJobData.LastCountUpdate < ExternalUpdateData.CreationDateTime Then
                PastJobData.LastCountUpdate = ExternalUpdateData.CreationDateTime
            End If
            
            Debug.Print (DateTime.Now & " HandleNoMatchUpdate-Count = 1PreCalc: JobID: " & PastJobData.JobID & ", JFID: " & PastJobData.JFID & ", CurJobDelta: " & PastJobData.JobCountDelta & "JobCount: " & PastJobData.JobCount)
            
            WriteLogToFile (DateTime.Now & " HandleNoMatchUpdate-Count = 1PreCalc: JobID: " & PastJobData.JobID & ", JFID: " & PastJobData.JFID & ", CurJobDelta: " & PastJobData.JobCountDelta & "JobCount: " & PastJobData.JobCount)
            
            PastJobData.JobCountDelta = (0 - ExternalUpdateData.UpdateAmount)
            PastJobData.JobCount = PastJobData.JobCount + ExternalUpdateData.UpdateAmount
            
            Debug.Print (DateTime.Now & " HandleNoMatchUpdate-Count = 1 PostCalc: NewJobDelta: " & PastJobData.JobCountDelta & ", JobCount: " & PastJobData.JobCount)
            
            WriteLogToFile (DateTime.Now & " HandleNoMatchUpdate-Count = 1 PostCalc: NewJobDelta: " & PastJobData.JobCountDelta & ", JobCount: " & PastJobData.JobCount)
            
            WriteJobDataToDB PastJobData
        Next PastJobData
        
        RemoveProcessedUpdate ExternalUpdateData
        
    ElseIf colPastJobs.Count > 1 Then
        'This hopefully won't happen.
        Debug.Print ("Multiple jobs for update date range search")
    End If

End Function

Public Function CheckForDataUpdates()
    Dim colModdedJobs As Collection
    Dim ModdedJobData As JobsDataClass
    Dim InMemoryJobData As JobsDataClass
    
    'UpdatingLabel.Visible = True
    
    Set colModdedJobs = GetTodaysJobsFromDB(RefTableMng.GetIDFromString(rte_Users, GetUserName), True)
    
    If colModdedJobs.Count <> 0 Then
        For Each ModdedJobData In colModdedJobs
            If InCollection(ControlGroupManager.JobDataCollection, CStr(ModdedJobData.JobID)) Then
                
                'This shouldn't be needed, was meant for when the fields were being updated directly before I
                'changed over to staging the updates in a separate table.
                
'                'update
'                'Get reference to job in memory
'                Set InMemoryJobData = ControlGroupManager.JobDataCollection.Item(CStr(ModdedJobData.JobID))
'
'                'Check if user is modifying data, postpone if being modified
'                If Not InMemoryJobData.ControlGroup.UserModifyingData Then
'                    'raise updatestart event
'                    InMemoryJobData.UpdateStart
'
'                    'update data to current db state
'                    InMemoryJobData.LastCountUpdate = ModdedJobData.LastCountUpdate
'                    InMemoryJobData.JobCountDelta = InMemoryJobData.JobCount - ModdedJobData.JobCount
'                    InMemoryJobData.JobCount = InMemoryJobData.JobCount - InMemoryJobData.JobCountDelta
'                    InMemoryJobData.JobCountDelta = 0
'
'                    InMemoryJobData.UpdateComplete
'
'                    ClearExternalModStatus ModdedJobData
'                End If
                
            Else:
                'Add New
                ControlGroupManager.AddJob ModdedJobData
                
                ClearExternalModStatus ModdedJobData
                
            End If
        Next ModdedJobData
    End If
    
    'UpdatingLabel.Visible = False

End Function

Public Function StopNextTrigger()
    If Now < TimeInterval Then
        ScheduleNextTrigger False
    End If
End Function

Public Function GetInstance() As ExternalUpdateManager
    Set GetInstance = Me
End Function
