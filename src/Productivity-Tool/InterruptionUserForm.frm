VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InterruptionUserForm 
   Caption         =   "Add Interruption"
   ClientHeight    =   2667
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3855
   OleObjectBlob   =   "InterruptionUserForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InterruptionUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IntrFTDC As RefTableDataClass
'Dim RefTableMng As ReferenceTableManager
Dim IntrSearch As SearchableComboboxClass
Dim CurrentJob As JobsDataClass
Dim CurrentIntr As InterruptionDataClass
Dim textVal As String
Dim numVal As Long
Dim MinuteCount As Long
Dim TimeInterval As Variant

Private Sub AmountTextBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    TimerEnabledCheckBox = False
    AmountTextBox.SetFocus
End Sub

Private Sub TimerEnabledCheckBox_Click()
    If TimerEnabledCheckBox.Value = True Then
        If TimeInterval Is Nothing Then
            'ScheduleNextTrigger
            TimerEnabledCheckBox.Caption = "Timer is On"
        End If
    Else:
        'ScheduleNextTrigger False
        TimerEnabledCheckBox.Caption = "Timer is Off"
    End If
End Sub

Private Sub UserForm_Initialize()
    'Initialize ReferenceTableManager
    If RefTableMng Is Nothing Then
        Set RefTableMng = New ReferenceTableManager
    End If
    
    AmountTextBox.Value = 0
    textVal = 0
    
    'ScheduleNextTrigger
    
    'Set up Searchable Combox Functionality
    Set IntrSearch = New SearchableComboboxClass
    IntrSearch.SetClassParameters GetNonDisabledRefTableDataForSearchFunction(RefTableMng.GetRefCol(rte_InterruptionType)), IntrFTDC, InterruptionsComboBox
    
    'Build JobFunction List
    PopulateRefDataToComboBox InterruptionsComboBox, RefTableMng.GetRefCol(rte_InterruptionType)
    
    Set CurrentIntr = New InterruptionDataClass
    
    CurrentIntr.UserID = RefTableMng.GetIDFromString(rte_Users, GetUserName)
    
    ResetToStart
    
End Sub

Private Sub AddCommandButton_Click()

    If textVal > 2147483647 Then
        MsgBox ("Number entered is larger than humanly possible.")
        Exit Sub
    End If

    TimerEnabledCheckBox = False
    ProcessInterruption
    UserForm_Terminate
End Sub

Private Sub AmountTextBox_Change()
    textVal = AmountTextBox.Value
    
    If IsNumeric(textVal) Then
        
        If textVal > 2147483647 Then
            MsgBox ("Number entered is larger than humanly possible.")
            Exit Sub
        End If
        
        numVal = Abs(textVal)
        AmountTextBox.Text = CStr(numVal)
    Else
        If textVal = vbNullString Then
            AmountTextBox.Text = 0
        Else:
            'CountTextBox.Text = CStr(Abs(numVal))
            AmountTextBox.Text = CStr(numVal)
        End If
    End If
    
    CurrentIntr.InterruptionLength = numVal

End Sub

Private Sub AmountTextBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    If KeyCode = vbKeyUp Then
'        LengthSpinButton_SpinUp
'        LengthSpinButton.SetFocus
'    ElseIf KeyCode = vbKeyDown Then
'        LengthSpinButton_SpinDown
'        LengthSpinButton.SetFocus
'    End If
'    TimerEnabledCheckBox = False
End Sub

Private Sub CancelCommandButton_Click()
    'ScheduleNextTrigger False
    UserForm_Terminate
End Sub

Private Sub InterruptionsComboBox_Change()
    If InterruptionsComboBox.Text <> vbNullString Then
        'JFComboBox.Enabled = False
        'CurrentJob.JFID = RefTableMng.GetIDFromString(rte_JobFunction, JFComboBox.Value)
        If IsNumeric(InterruptionsComboBox.Value) Then
            IntrRefAddCommandButton.Enabled = False
            AddCommandButton.Enabled = True
            'TimerEnabledCheckBox = False
            EnableButtons
        Else:
            IntrRefAddCommandButton.Enabled = True
            AddCommandButton.Enabled = False
            ResetToStart
        End If
    Else:
        IntrRefAddCommandButton.Enabled = False
        AddCommandButton = False
        ResetToStart
    End If
End Sub

Private Sub InterruptionsComboBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    IntrSearch.KeyDownEvent KeyCode, Shift
End Sub

Private Sub InterruptionsComboBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    IntrSearch.KeyPressEvent
End Sub

Private Sub InterruptionsComboBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    IntrSearch.KeyUpEvent KeyCode, Shift
    
    If KeyCode = vbKeyReturn Then
'        If IntrRefAddCommandButton.Enabled Then
'            IntrRefAddCommandButton_Click
'        End If
'    Else:
'        If AmountTextBox.Enabled = True Then
'            AmountTextBox.SetFocus
'        End If
    End If
End Sub

Private Sub IntrRefAddCommandButton_Click()
    Dim tmpTextVal As String
    
    If Not IsNumeric(InterruptionsComboBox.Value) Then
        If Not RefTableMng.CheckIfTextExists(rte_InterruptionType, InterruptionsComboBox.Text) Then
            RefTableMng.AddNewRefData rte_InterruptionType, InterruptionsComboBox.Text
            
            tmpTextVal = InterruptionsComboBox.Text
            
            IntrSearch.UpdateTargetCollection RefTableMng.GetRefCol(rte_InterruptionType)
            
            PopulateRefDataToComboBox InterruptionsComboBox, RefTableMng.GetRefCol(rte_InterruptionType)
            
            InterruptionsComboBox.Text = tmpTextVal
            
        End If
    End If
End Sub

Private Sub LengthSpinButton_SpinDown()
    TimerEnabledCheckBox = False
    If Not AmountTextBox.Text - 1 < 0 Then
        AmountTextBox.Text = AmountTextBox.Text - 1
    End If
End Sub

Private Sub LengthSpinButton_SpinUp()
    TimerEnabledCheckBox = False
    AmountTextBox.Text = AmountTextBox.Text + 1
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        'ScheduleNextTrigger False
    End If
End Sub

Private Sub UserForm_Terminate()
    AmountTextBox = 0
    Set CurrentIntr = Nothing
    Unload Me
End Sub

''Deprecated
'Public Function LoadJob(ByVal TargetJob As JobsDataClass)
'    Set CurrentJob = New JobsDataClass
'    Set CurrentIntr = New InterruptionDataClass
'
'    Set CurrentJob = TargetJob
'
'    'CurrentIntr.JobID = CurrentJob.JobID
'    CurrentIntr.UserID = CurrentJob.UserID
'    CurrentIntr.InterruptDateTime = CurrentJob.StartDateTime
'
'    ResetToStart
'
'End Function

Private Function ProcessInterruption()
    If CurrentIntr.InterruptDateTime = EMPTY_DATE Then
        CurrentIntr.InterruptDateTime = DateTime.Now
    End If
    CurrentIntr.InterruptTypeID = InterruptionsComboBox.Value
    CurrentIntr.InterruptionLength = AmountTextBox.Text
    
    WriteInterruptionDataToDB CurrentIntr
    
End Function

Private Function ResetToStart()
    AddCommandButton.Enabled = False
    AmountTextBox.Enabled = False
    LengthSpinButton.Enabled = False
End Function
Private Function EnableButtons()
    AddCommandButton.Enabled = True
    AmountTextBox.Enabled = True
    LengthSpinButton.Enabled = True
End Function

Public Sub UpdateTimeAmount()
        'MinuteCount = MinuteCount + 1
        AmountTextBox.Text = AmountTextBox.Text + 1
        'ScheduleNextTrigger
End Sub

Private Sub ScheduleNextTrigger(Optional ByVal KeepTracking As Boolean = True)
    If KeepTracking Then
        TimeInterval = Now + TimeValue("00:01:00")
        Application.OnTime TimeInterval, "Utilities.UpdateTimeAmount"
    Else:
        'To help check for different erros that are not the expected one.
        On Error Resume Next
        Application.OnTime TimeInterval, "Utilities.UpdateTimeAmount", , False
        If Err.Number <> 0 Then
            Debug.Print (Err.Number & ": " & Err.Description)
        End If
        On Error GoTo 0
        
        Set TimeInterval = Nothing
    End If
End Sub
