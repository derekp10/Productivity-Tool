VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModifyElementsUserForm 
   Caption         =   "Modify Task"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6345
   OleObjectBlob   =   "ModifyElementsUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ModifyElementsUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim RefTableMng As ReferenceTableManager
Dim RefTblDC As RefTableDataClass
Dim CurrentJob As JobsDataClass
Dim OriginalJobData As JobsDataClass
Dim JobFunctSearch As SearchableComboboxClass
Dim StartDTPicker As GenericDateTimePicker
Dim EndDTPicker As GenericDateTimePicker
Dim ModData As ModificationDataClass
Dim textVal As String
Dim numVal As String

'EndOptionButton and EndDay comboboxes are now being used for LastCountUpdate

Private Sub EndDayComboBox_Change()
    EndDTPicker.DayCB_Change
    CheckForValidDates
End Sub

Private Sub EndHourComboBox_Change()
    EndDTPicker.HourCB_Change
    CheckForValidDates
End Sub

Private Sub EndMinComboBox_Change()
    EndDTPicker.MinCB_Change
    CheckForValidDates
End Sub

Private Sub EndMonthComboBox_Change()
    EndDTPicker.MonthCB_Change
    CheckForValidDates
End Sub

Private Sub EndOptionButton_Change()
    LockDataFields
End Sub

Private Sub EndOptionButton_Click()
    LockDataFields
    EndDTPicker.SetEnabled True
End Sub

Private Sub EndYearComboBox_Change()
    EndDTPicker.YearCB_Change
    CheckForValidDates
End Sub

Private Sub InterruptionCommandButton_Click()
    Load InterruptionUserForm
    InterruptionUserForm.LoadJob CurrentJob
    'TODO: Fix Model issue
    InterruptionUserForm.Show

End Sub

Private Sub JobCountOptionButton_Change()
    LockDataFields
End Sub

Private Sub JobCountOptionButton_Click()
    LockDataFields
    JobCountTextBox.Enabled = True
End Sub

Private Sub JobCountTextBox_Change()
'    textVal = JobCountTextBox.Value
'
'    If IsNumeric(textVal) Then
'        numVal = Abs(textVal)
'        JobCountTextBox.Text = CStr(numVal)
'    Else
'        If textVal = vbNullString Then
'            JobCountTextBox.Text = 0
'        Else:
'            'CountTextBox.Text = CStr(Abs(numVal))
'            JobCountTextBox.Text = CStr(numVal)
'        End If
'    End If

    numVal = OriginalJobData.JobCount
    
    textVal = JobCountTextBox.Value
    
    If numVal <> textVal Then
        If IsNumeric(textVal) Then
        
            If textVal > 2147483647 Then
                MsgBox ("Number entered is larger than humanly possible.")
                SaveCommandButton.Enabled = False
                Exit Sub
            Else:
                SaveCommandButton.Enabled = True
            End If
            
            numVal = Abs(textVal)
            JobCountTextBox.Text = CStr(numVal)
        Else
            If textVal = vbNullString Then
                JobCountTextBox.Text = 0
                numVal = 0
            Else:
                'CountTextBox.Text = CStr(Abs(numVal))
                'JobCountTextBox.Text = CStr(numVal)
                JobCountTextBox.Text = CStr(CurrentJob.JobCount)
            End If
        End If

    End If
    
End Sub

Private Sub JobFunctionComboBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    JobFunctSearch.KeyDownEvent KeyCode, Shift
End Sub

Private Sub JobFunctionComboBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    JobFunctSearch.KeyPressEvent
End Sub

Private Sub JobFunctionComboBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    JobFunctSearch.KeyUpEvent KeyCode, Shift
End Sub

Private Sub JobFunOptionButton_Change()
    LockDataFields
End Sub

Private Sub JobFunOptionButton_Click()
    LockDataFields
    JobFunctionComboBox.Enabled = True
End Sub

Private Sub SaveCommandButton_Click()
    ProcessModifications
    UserForm_Terminate
End Sub

Private Sub StartDayComboBox_Change()
    StartDTPicker.DayCB_Change
    CheckForValidDates
End Sub

Private Sub StartHourComboBox_Change()
    StartDTPicker.HourCB_Change
    CheckForValidDates
End Sub

Private Sub StartMinComboBox_Change()
    StartDTPicker.MinCB_Change
    CheckForValidDates
End Sub

Private Sub StartMonthComboBox_Change()
    StartDTPicker.MonthCB_Change
    CheckForValidDates
End Sub

Private Sub StartOptionButton_Change()
    LockDataFields
End Sub

Private Sub StartOptionButton_Click()
    LockDataFields
    StartDTPicker.SetEnabled True
End Sub

Private Sub StartYearComboBox_Change()
    StartDTPicker.YearCB_Change
    CheckForValidDates
End Sub

Private Sub StatusOptionButton_Change()
    LockDataFields
End Sub

Private Sub StatusOptionButton_Click()
    LockDataFields
    StatusComboBox.Enabled = True
End Sub

Private Sub UserForm_Initialize()
    'Initialize ReferenceTableManager
    If RefTableMng Is Nothing Then
        Set RefTableMng = New ReferenceTableManager
    End If
    
    'Set up Searchable Combox Functionality
    Set JobFunctSearch = New SearchableComboboxClass
    JobFunctSearch.SetClassParameters GetNonDisabledRefTableDataForSearchFunction(RefTableMng.GetRefCol(rte_JobFunction)), RefTblDC, JobFunctionComboBox
    
    'Build ComboBox Lists
    PopulateRefDataToComboBox JobFunctionComboBox, RefTableMng.GetRefCol(rte_JobFunction), , True
    PopulateRefDataToComboBox StatusComboBox, RefTableMng.GetRefCol(rte_Status), , True
    
    'Set up DateTimePickers
    Set StartDTPicker = New GenericDateTimePicker
    Set EndDTPicker = New GenericDateTimePicker
    
    StartDTPicker.AddUFElement StartMonthComboBox
    StartDTPicker.AddUFElement StartDayComboBox
    StartDTPicker.AddUFElement StartYearComboBox
    StartDTPicker.AddUFElement StartHourComboBox
    StartDTPicker.AddUFElement StartMinComboBox
    
    EndDTPicker.AddUFElement EndMonthComboBox
    EndDTPicker.AddUFElement EndDayComboBox
    EndDTPicker.AddUFElement EndYearComboBox
    EndDTPicker.AddUFElement EndHourComboBox
    EndDTPicker.AddUFElement EndMinComboBox
    
    
    LockDataFields
    
    LoadJob CreateTestJob
    
End Sub

Private Sub CancelCommandButton_Click()
    UserForm_Terminate
End Sub

Private Sub UserForm_Terminate()
    Unload Me
End Sub

Private Function LockDataFields()
    StartDTPicker.SetEnabled False
    EndDTPicker.SetEnabled False
    JobFunctionComboBox.Enabled = False
    StatusComboBox.Enabled = False
    JobCountTextBox.Enabled = False
End Function

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
    
    ResetDataFields
    
    JobIDLabel.Caption = "JobID: " & TargetJob.JobID
    StartDTPicker.DateTimeValue = TargetJob.StartDateTime
    EndDTPicker.DateTimeValue = TargetJob.LastCountUpdate
    JobFunctionComboBox.Text = RefTableMng.GetStringFromID(rte_JobFunction, TargetJob.JFID)
    StatusComboBox.Text = RefTableMng.GetStringFromID(rte_Status, TargetJob.StatusID)
    JobCountTextBox = TargetJob.JobCount
    
End Function

Private Function ResetDataFields()
    StartDTPicker.DateTimeValue = EMPTY_DATE
    EndDTPicker.DateTimeValue = EMPTY_DATE
    JobFunctionComboBox.Text = ""
    StatusComboBox.Text = ""
    JobCountTextBox.Text = ""
End Function

Private Function ProcessModifications()
    Dim ModDateTime As Date
    Dim DataChanged As Boolean
    
    ModDateTime = DateTime.Now

    If StartDTPicker.DateTimeValue <> Format(OriginalJobData.StartDateTime, "mm/dd/yyyy hh:nn") Then
        Set ModData = New ModificationDataClass
        ModData.JobID = OriginalJobData.JobID
        ModData.UserID = OriginalJobData.UserID
        ModData.ModDateTime = ModDateTime
        ModData.ModElmID = RefTableMng.GetIDFromString(rte_ModElm, "StartDateTime")
        ModData.NewValue = StartDTPicker.DateTimeValue
        ModData.OldValue = OriginalJobData.StartDateTime
        CurrentJob.StartDateTime = StartDTPicker.DateTimeValue
        
        WriteModificationDataToDB ModData
        
        DataChanged = True
    End If
    
    If EndDTPicker.DateTimeValue <> Format(OriginalJobData.LastCountUpdate, "mm/dd/yyyy hh:nn") Then
        Set ModData = New ModificationDataClass
        ModData.JobID = OriginalJobData.JobID
        ModData.UserID = OriginalJobData.UserID
        ModData.ModDateTime = ModDateTime
        ModData.ModElmID = RefTableMng.GetIDFromString(rte_ModElm, "LastCountUpdate")
        ModData.NewValue = EndDTPicker.DateTimeValue
        ModData.OldValue = OriginalJobData.LastCountUpdate
        CurrentJob.LastCountUpdate = EndDTPicker.DateTimeValue
        
        WriteModificationDataToDB ModData
        
        DataChanged = True
    End If
    
    If JobFunctionComboBox.Value <> OriginalJobData.JFID Then
        Set ModData = New ModificationDataClass
        ModData.JobID = OriginalJobData.JobID
        ModData.UserID = OriginalJobData.UserID
        ModData.ModDateTime = ModDateTime
        ModData.ModElmID = RefTableMng.GetIDFromString(rte_ModElm, "JobFunction")
        ModData.NewValue = JobFunctionComboBox.Value
        ModData.OldValue = OriginalJobData.JFID
        CurrentJob.JFID = JobFunctionComboBox.Value
        
        WriteModificationDataToDB ModData
        
        DataChanged = True
    End If
    
    If StatusComboBox.Value <> OriginalJobData.StatusID Then
        Set ModData = New ModificationDataClass
        ModData.JobID = OriginalJobData.JobID
        ModData.UserID = OriginalJobData.UserID
        ModData.ModDateTime = ModDateTime
        ModData.ModElmID = RefTableMng.GetIDFromString(rte_ModElm, "Status")
        ModData.NewValue = StatusComboBox.Value
        ModData.OldValue = OriginalJobData.StatusID
        CurrentJob.StatusID = StatusComboBox.Value
        
        WriteModificationDataToDB ModData
        
        DataChanged = True
    End If
    
    If JobCountTextBox.Text <> OriginalJobData.JobCount Then
        Set ModData = New ModificationDataClass
        ModData.JobID = OriginalJobData.JobID
        ModData.UserID = OriginalJobData.UserID
        ModData.ModDateTime = ModDateTime
        ModData.ModElmID = RefTableMng.GetIDFromString(rte_ModElm, "Job Count")
        ModData.NewValue = JobCountTextBox.Text
        ModData.OldValue = OriginalJobData.JobCount
        
        CurrentJob.JobCountDelta = OriginalJobData.JobCount - JobCountTextBox.Value
        
        CurrentJob.JobCount = CurrentJob.JobCount - CurrentJob.JobCountDelta
        
        WriteModificationDataToDB ModData
        
        DataChanged = True
    End If
    
    If DataChanged Then
        WriteJobDataToDB CurrentJob
    End If
End Function

Private Function CheckForValidDates()
    If Not StartDTPicker.IsDataValid Or Not EndDTPicker.IsDataValid Then
        SaveCommandButton.Enabled = False
    Else:
        SaveCommandButton.Enabled = True
    End If
End Function


