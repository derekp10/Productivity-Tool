Attribute VB_Name = "TestCode"
Option Explicit

Private Sub TestBot()
    Dim RefTableMng As ReferenceTableManager
    Dim CurrentJob As JobsDataClass
    Dim jobCnt As Long
    
    Set RefTableMng = New ReferenceTableManager
    
    Do
        Set CurrentJob = New JobsDataClass
        
        CurrentJob.StartDateTime = DateTime.Now
        CurrentJob.JobCreationDateTime = CurrentJob.StartDateTime
        CurrentJob.JFID = CLng(((RefTableMng.GetRefCol(rte_JobFunction).Count - 1) * Rnd) + 1)
        'CurrentJob.JobCount = CLng((100 * Rnd))
        CurrentJob.UserID = CLng(((RefTableMng.GetRefCol(rte_Users).Count - 2) * Rnd) + 2)
        'CurrentJob.UserID = CLng((7 * Rnd) + 2)
        jobCnt = CLng((100 * Rnd))
        Do Until CurrentJob.JobCount = jobCnt
            DoEvents
            CurrentJob.JobCount = CurrentJob.JobCount + 1
            CurrentJob.StatusID = 1
            WriteJobDataToDB CurrentJob
        Loop
        
        'CurrentJob.StatusID = CLng((RefTableMng.GetRefCol(rte_Status).Count * Rnd) + 1)
        CurrentJob.JobCreationDateTime = DateTime.Now
        CurrentJob.StatusID = 2
        
        WriteJobDataToDB CurrentJob
        
        DoEvents
    Loop
    
End Sub

Public Function CreateTestJob() As JobsDataClass
    Dim RefTableMng As ReferenceTableManager
    Dim returnJob As JobsDataClass
    
    Set RefTableMng = New ReferenceTableManager
    
    Set returnJob = New JobsDataClass
        
    returnJob.StartDateTime = DateTime.Now
    returnJob.JobCreationDateTime = returnJob.StartDateTime
    'returnJob.JFID = CLng(((RefTableMng.GetRefCol(rte_JobFunction).Count - 1) * Rnd) + 1)
    returnJob.JFID = 54 'LabelControlGroup Item
    'returnJob.JobCount = CLng((100 * Rnd))
    returnJob.UserID = CLng(((RefTableMng.GetRefCol(rte_Users).Count - 2) * Rnd) + 2)
    'returnJob.UserID = CLng((7 * Rnd) + 2)
    returnJob.JobCount = CLng((100 * Rnd))
    
    'returnJob.StatusID = CLng((RefTableMng.GetRefCol(rte_Status).Count * Rnd) + 1)
    returnJob.LastCountUpdate = DateTime.Now
    returnJob.StatusID = 2
    
    Set CreateTestJob = returnJob
End Function
