VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OLD_ModifyElementsUserForm 
   Caption         =   "Modify Job"
   ClientHeight    =   1918
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2970
   OleObjectBlob   =   "OLD_ModifyElementsUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OLD_ModifyElementsUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CurrentJob As JobsDataClass
Dim OriginalJobData As JobsDataClass
Dim ModData As ModificationDataClass
Dim textVal As String
Dim numVal As String

Private Sub JobCountTextBox_Change()
    textVal = JobCountTextBox.Value
    
    If IsNumeric(textVal) Then
    
        If textVal > 2147483647 Then
            MsgBox ("Number entered is larger than humanly possible.")
            Exit Sub
        End If
        
        numVal = Abs(textVal)
        JobCountTextBox.Text = CStr(numVal)
    Else
        If textVal = vbNullString Then
            JobCountTextBox.Text = 0
        Else:
            'CountTextBox.Text = CStr(Abs(numVal))
            JobCountTextBox.Text = CStr(numVal)
        End If
    End If
    
End Sub

Private Sub SaveCommandButton_Click()
    
    If textVal > 2147483647 Then
        MsgBox ("Number entered is larger than humanly possible.")
        Exit Sub
    End If
    
    ProcessModifications
    UserForm_Terminate
End Sub
Private Sub CancelCommandButton_Click()
    UserForm_Terminate
End Sub

Private Sub UserForm_Initialize()
    'Initialize ReferenceTableManager
    If RefTableMng Is Nothing Then
        Set RefTableMng = New ReferenceTableManager
    End If
    
    LoadJob CreateTestJob
    
End Sub

Private Sub UserForm_Terminate()
    Unload Me
End Sub

Public Function LoadJob(ByVal TargetJob As JobsDataClass)
    Set OriginalJobData = New JobsDataClass
    Set CurrentJob = New JobsDataClass
    Set ModData = New ModificationDataClass
    
    OriginalJobData.JFID = TargetJob.JFID
    OriginalJobData.LastCountUpdate = TargetJob.LastCountUpdate
    OriginalJobData.JobCount = TargetJob.JobCount
    OriginalJobData.JobID = TargetJob.JobID
    OriginalJobData.StartDateTime = TargetJob.StartDateTime
    OriginalJobData.JobCreationDateTime = TargetJob.JobCreationDateTime
    OriginalJobData.StatusID = TargetJob.StatusID
    OriginalJobData.UserID = TargetJob.UserID
    
    Set CurrentJob = TargetJob
    
    FillFormWithJobData CurrentJob
    
End Function

Private Function FillFormWithJobData(ByRef TargetJob As JobsDataClass)
    
    JobNameLabel.Caption = "Job Name: " & RefTableMng.GetStringFromID(rte_JobFunction, TargetJob.JFID)
    
    JobCountTextBox = TargetJob.JobCount
    
End Function

Private Function ProcessModifications()
    Dim ModDateTime As Date
    Dim DataChanged As Boolean
    
    ModDateTime = DateTime.Now
    
    If JobCountTextBox.Text <> OriginalJobData.JobCount Then
        
        Set ModData = New ModificationDataClass
        ModData.JobID = OriginalJobData.JobID
        ModData.UserID = OriginalJobData.UserID
        ModData.ModDateTime = ModDateTime
        ModData.ModElmID = RefTableMng.GetIDFromString(rte_ModElm, "Job Count")
        ModData.NewValue = JobCountTextBox.Text
        ModData.OldValue = OriginalJobData.JobCount
        CurrentJob.JobCount = JobCountTextBox.Text
        
        WriteModificationDataToDB ModData
        
        DataChanged = True
    End If
    
    If DataChanged Then
        WriteJobDataToDB CurrentJob
    End If
End Function
