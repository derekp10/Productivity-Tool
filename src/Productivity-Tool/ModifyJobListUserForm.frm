VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModifyJobListUserForm 
   Caption         =   "Previous Tasks"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8475
   OleObjectBlob   =   "ModifyJobListUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ModifyJobListUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim colUserJobs As Collection

Private Sub JobsListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ModifyCommandButton_Click
End Sub

Private Sub ModifyCommandButton_Click()
    Dim CurrentSelect As Long
    Dim lIndex As Long
    
    If JobsListBox.ListCount <> 0 Then
        If JobsListBox.ListIndex <> -1 Then
            CurrentSelect = JobsListBox.Value
            lIndex = JobsListBox.ListIndex
            Load ModifyElementsUserForm
'            Me.Hide
            ModifyElementsUserForm.LoadJob colUserJobs.Item(CStr(CurrentSelect))
            ModifyElementsUserForm.Show
            
            BuildJobsListBox
            
            JobsListBox.ListIndex = lIndex
            
            JobsListBox.SetFocus
            'ModifyCommandButton.SetFocus
            
'            Me.Show

        End If
        
    End If
    
End Sub

Private Function BuildJobsListBox()
    Set colUserJobs = GetJobColFromDB(RefTableMng.GetIDFromString(rte_Users, GetUserName))
    PopulateJobsListBox colUserJobs
End Function

Private Sub UserForm_Initialize()
    'Initialize ReferenceTableManager
    If RefTableMng Is Nothing Then
        Set RefTableMng = New ReferenceTableManager
    End If
    
    BuildJobsListBox
    
End Sub

Private Sub CancelCommandButton_Click()
    UserForm_Terminate
End Sub

Private Sub UserForm_Terminate()
    Unload Me
End Sub

'Iterates through the colUserJobs collection and creates the listbox data (an array).
Private Function PopulateJobsListBox(ByRef colUserJobs As Collection)
    Dim arrJobsQueue()
    Dim lCounter As Long
    Dim JobData As JobsDataClass
    
    JobsListBox.Clear
    
    lCounter = 0
    
    If colUserJobs.Count <> 0 Then
    
        ReDim arrJobsQueue(colUserJobs.Count - 1, 3)
        
        For Each JobData In colUserJobs
            
            arrJobsQueue(lCounter, 0) = JobData.JobID
            arrJobsQueue(lCounter, 1) = Format(JobData.StartDateTime, "mm/dd/yyyy")
            
            arrJobsQueue(lCounter, 2) = RefTableMng.GetStringFromID(rte_JobFunction, JobData.JFID)
            arrJobsQueue(lCounter, 3) = JobData.JobCount
            
            lCounter = lCounter + 1
        Next JobData
        
        JobsListBox.List = arrJobsQueue
    End If
    
    
    
    '13 to 14 = 1 character (Aproximate)
'    WorkQueueListBox.ColumnWidths = "0;85;175;60;60;125;125;14"
    JobsListBox.ColumnWidths = "0;75;255;50"

    If lCounter <> 0 Then
        JobsListBox.ListIndex = 0
    End If
    
End Function

