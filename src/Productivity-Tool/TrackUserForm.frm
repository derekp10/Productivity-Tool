VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TrackUserForm 
   Caption         =   "Track"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5310
   OleObjectBlob   =   "TrackUserForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TrackUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim textVal As String
Dim numVal As String
Dim JFSearch As SearchableComboboxClass
Dim RefTblDC As RefTableDataClass
Dim ResettingToStart As Boolean

Dim WithEvents CurrentJob As JobsDataClass
Attribute CurrentJob.VB_VarHelpID = -1
Dim WithEvents EUM As ExternalUpdateManager
Attribute EUM.VB_VarHelpID = -1

Private Function CheckActiveJobsScroll()
    If ActiveJobsFrame.Height < ControlGroupManager.GetScrollHeight Then
        'This will create a vertical scrollbar
        If ActiveJobsFrame.ScrollBars <> fmScrollBarsVertical Then
            ActiveJobsFrame.ScrollBars = fmScrollBarsVertical
        End If

        'Change the values of 2 as Per your requirements
        'ActiveJobsFrame.ScrollHeight = .InsideHeight * 1.1
        ActiveJobsFrame.ScrollHeight = ControlGroupManager.GetScrollHeight
        'ActiveJobsFrame.ScrollWidth = .InsideWidth * 9
    Else:
        ActiveJobsFrame.ScrollBars = fmScrollBarsNone
    End If
End Function

Private Sub AddJobCommandButton_Click()
    Dim NewJob As New JobsDataClass
    
    ExternalUpdateManager.CheckForDataUpdates
    
    NewJob.JobCreationDateTime = DateTime.Now
    NewJob.StartDateTime = NewJob.JobCreationDateTime
    NewJob.JFID = JFComboBox.Value
    NewJob.StatusID = RefTableMng.GetIDFromString(rte_Status, "Inactive")
    NewJob.UserID = RefTableMng.GetIDFromString(rte_Users, GetUserName)
        
    If Not ControlGroupManager.IsJobInGroup(NewJob) Then
        WriteJobDataToDB NewJob
        ControlGroupManager.AddJob NewJob
        CheckActiveJobsScroll
        ControlGroupManager.SetFocusTarget NewJob.JFID
    Else:
        'MsgBox ("Job is already in Today's Active Jobs.")
        ControlGroupManager.SetFocusTarget NewJob.JFID
    End If
    
    JFComboBox.Text = ""
    JFSearch.KeyPressEvent
    
End Sub

Private Sub InterruptionCommandButton_Click()
    Load InterruptionUserForm
    'InterruptionUserForm.LoadJob CurrentJob
    InterruptionUserForm.Show
End Sub

Private Sub JFComboBox_Change()
    If JFComboBox.Text <> vbNullString Then
        'JFComboBox.Enabled = False
        'CurrentJob.JFID = RefTableMng.GetIDFromString(rte_JobFunction, JFComboBox.Value)
        If IsNumeric(JFComboBox.Value) Then
            AddJobCommandButton.Enabled = True
            'StartCommandButton.Enabled = True
            'CurrentJob.JFID = JFComboBox.Value
        Else:
            'StartCommandButton.Enabled = False
            AddJobCommandButton.Enabled = False
        End If
    Else:
        'StartCommandButton.Enabled = False
        AddJobCommandButton.Enabled = False
    End If
    
End Sub

Private Sub ModifyCommandButton_Click()
    ExternalUpdateManager.StopNextTrigger
    
    PushPendingChanges
    
    Load ModifyJobListUserForm
    'Me.Hide
    ModifyJobListUserForm.Show
    
    ControlGroupManager.JobDataCollection = GetTodaysJobsFromDB(CurrentJob.UserID)
    
    'Me.Show
    
    ExternalUpdateManager.ScheduledCheckForDataUpdates
    
End Sub

Private Sub UserForm_Initialize()
    Dim tmpRefTableData As RefTableDataClass
    Dim enabledJFCol
    'Initialize ReferenceTableManager
    If RefTableMng Is Nothing Then
        Set RefTableMng = New ReferenceTableManager
    End If
    
    ExternalUpdateManager.UFActive = True
    
    'Set up Searchable Combox Functionality
    Set JFSearch = New SearchableComboboxClass
    
    'This implementation is not effected by RefTableUpdates. (Static Copy)
    JFSearch.SetClassParameters GetNonDisabledRefTableDataForSearchFunction(RefTableMng.GetRefCol(rte_JobFunction)), RefTblDC, JFComboBox
    
    'Build JobFunction List
    PopulateRefDataToComboBox JFComboBox, RefTableMng.GetRefCol(rte_JobFunction)
    
    'Initialize JobsDataClass
    Set CurrentJob = New JobsDataClass
    
    If Not RefTableMng.CheckIfTextExists(rte_Users, GetUserName) Then
        Set tmpRefTableData = RefTableMng.GetConfiguredRefTableDataClass(rte_Users)
        tmpRefTableData.RefTypeName = GetUserName
        RefTableMng.AddNewRefData rte_Users, tmpRefTableData
    End If
    
    CurrentJob.UserID = RefTableMng.GetIDFromString(rte_Users, GetUserName)
    
    'Create control groups for form
    Set ControlGroupManager = New ControlGroupManagerClass
    ControlGroupManager.FrameForControlGroups = Me.ActiveJobsFrame
    
    ControlGroupManager.JobDataCollection = GetTodaysJobsFromDB(CurrentJob.UserID)
    
    ExternalUpdateManager.ScheduledCheckForDataUpdates
    
    'Enable scrolling and set height for the ActiveJobsFrame
    CheckActiveJobsScroll
    
    ResetToStart
    
    'ScheduleNextTrigger
    
    Set EUM = ExternalUpdateManager.GetInstance
    
End Sub

Private Sub JFComboBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    JFSearch.KeyDownEvent KeyCode, Shift
End Sub

Private Sub JFComboBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    JFSearch.KeyPressEvent
End Sub

Private Sub JFComboBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    JFSearch.KeyUpEvent KeyCode, Shift
    
    If KeyCode = vbKeyReturn Then
        If AddJobCommandButton.Enabled Then
            AddJobCommandButton_Click
        End If
    End If
End Sub

Private Sub CancelCommandButton_Click()
    'ScheduleNextTrigger False
    UserForm_Terminate
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Debug.Print ("Escape")
    End If
End Sub

Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyCode = vbKeyEscape Then
        Debug.Print ("Escape")
    End If
End Sub

Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Debug.Print ("Escape")
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
'        If StopCommandButton.Enabled = True Then
'            MsgBox ("Please use stop button if you are done with this job function.")
'            Cancel = True
'        Else:
'            Cancel = False
'        End If
    End If
    
    ExternalUpdateManager.UFActive = False
    
    ExternalUpdateManager.StopNextTrigger
End Sub

Private Sub UserForm_Terminate()
    'ScheduleNextTrigger False
    'ExternalUpdateManager.StopUpdateChecks
    PushPendingChanges
    Unload Me
    'Me.Hide
End Sub

Private Function ResetToStart()
    ResettingToStart = True
    'StartCommandButton.Enabled = False
    'StopCommandButton.Enabled = False
    'CountTextBox.Enabled = False
    'CountSpinButton.Enabled = False
    'InterruptionCommandButton.Enabled = False
    'StartDateTimeLabel.Caption = ""
    'EndDateTimeLabel.Caption = ""
    JFComboBox.Enabled = True
    PopulateRefDataToComboBox JFComboBox, RefTableMng.GetRefCol(rte_JobFunction)
    CancelCommandButton.Enabled = True
    'CountTextBox.Text = 0
    Me.Repaint
    JFComboBox.SetFocus
    ResettingToStart = False
    
End Function

Private Sub EUM_ExternalUpdateEnd()
    UpdatingLabel.Visible = False
    ModifyCommandButton.Enabled = True
    CheckActiveJobsScroll
End Sub

Private Sub EUM_ExternalUpdateStart()
    UpdatingLabel.Visible = True
    ModifyCommandButton.Enabled = False
End Sub

Private Function PushPendingChanges()
    'To be called on cancel/close events.
    Dim tmpJobData As JobsDataClass
    
    For Each tmpJobData In ControlGroupManager.JobDataCollection
        'If pending usermodifications
        If tmpJobData.ControlGroup.UserModifyingData Then
            'Push modifications and clear Application.Ontime
            tmpJobData.ControlGroup.ClearUserModifyFlag
        End If
    Next tmpJobData
    
End Function
