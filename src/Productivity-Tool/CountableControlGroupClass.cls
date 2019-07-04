VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CountableControlGroupClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IControlGroupClass

Private WithEvents m_Label As MSForms.Label
Attribute m_Label.VB_VarHelpID = -1
Private WithEvents m_TextBox As MSForms.TextBox
Attribute m_TextBox.VB_VarHelpID = -1
Private WithEvents m_SpinButton As MSForms.SpinButton
Attribute m_SpinButton.VB_VarHelpID = -1
Private WithEvents m_Frame As MSForms.Frame
Attribute m_Frame.VB_VarHelpID = -1

Private m_GroupName As String

Private WithEvents m_BoundJob As JobsDataClass
Attribute m_BoundJob.VB_VarHelpID = -1

Private m_LabelTopModifier As Double

Private m_FrameTop As Double
Private m_ControlTop As Double
Private m_FrameLeft As Double

Private m_PreviousGFHeight As Double
Private m_PreviousGFWidth As Double
Private m_PreviousGFTop As Double

Private textVal As String
Private numVal As String
Private boolLoadingJob As Boolean
Private boolUserModifying As Boolean
Private TimeInterval As Variant

Private Sub Class_Initialize()
    m_LabelTopModifier = 4
    
    m_FrameTop = 5
    m_ControlTop = 5
    m_FrameLeft = 5
    
    m_PreviousGFHeight = 0
    m_PreviousGFWidth = 0
    m_PreviousGFTop = 0
    
    numVal = 0
    
End Sub

'Class Functions
Public Function CreateControlFrame(ByRef TargetFrame As MSForms.Frame)
    Set m_Frame = TargetFrame.Controls.Add("Forms.Frame.1")
    
    boolLoadingJob = True
    
    m_Frame.Caption = ""
    m_Frame.Name = m_GroupName & "_Frame"
    
    m_Frame.Height = 33
    
    AddControls m_Frame
    
    m_Frame.Top = m_PreviousGFTop + m_PreviousGFHeight + m_FrameTop
    m_Frame.Left = m_FrameLeft
    m_Frame.Width = m_SpinButton.Left + m_SpinButton.Width + 8
    
    
    
    boolLoadingJob = False
    'Set GetControlFrame = m_Frame
    
    
End Function

Private Function AddControls(ByRef TargetFrame As MSForms.Frame)
    Set m_Label = TargetFrame.Controls.Add("Forms.Label.1")
    Set m_TextBox = TargetFrame.Controls.Add("Forms.TextBox.1")
    Set m_SpinButton = TargetFrame.Controls.Add("Forms.SpinButton.1")
    
    m_TextBox.Name = m_GroupName & "_TextBox"
    m_Label.Name = m_GroupName & "_Label"
    m_SpinButton.Name = m_GroupName & "_SpinButton"
    
    m_Label.Caption = RefTableMng.GetStringFromID(rte_JobFunction, BoundJob.JFID)
    
    m_TextBox.Text = BoundJob.JobCount
    
    If m_ControlTop + m_LabelTopModifier + m_Label.Height > TargetFrame.Height Then
        m_Label.Top = m_ControlTop
    Else:
        m_Label.Top = m_ControlTop + m_LabelTopModifier
    End If

    m_Label.Width = 165
    m_Label.WordWrap = False
    
    m_TextBox.Top = m_ControlTop
    m_TextBox.Width = 30
    m_TextBox.Height = 18
    
    m_SpinButton.Top = m_ControlTop
    m_SpinButton.Width = 12
    m_SpinButton.Height = 18

    m_Label.Left = 5
    m_TextBox.Left = m_Label.Left + m_Label.Width + 5
    m_SpinButton.Left = m_TextBox.Left + m_TextBox.Width
    
    
End Function

Public Function SetFramePlacement(ByRef LastGroupFrame As IControlGroupClass)
    m_PreviousGFTop = LastGroupFrame.GetGroupFrame.Top
    m_PreviousGFWidth = LastGroupFrame.GetGroupFrame.Width
    m_PreviousGFHeight = LastGroupFrame.GetGroupFrame.Height
    
End Function

Private Function IControlGroupClass_ClearUserModifyFlag() As Variant
    ClearUserModifyFlag
End Function

Private Sub IControlGroupClass_MakeGroupFocusTarget()
    'm_TextBox.SetFocus
    m_SpinButton.SetFocus
    'm_TextBox.CurLine = 0
End Sub

Private Function LockEditing(ByVal LockVal As Boolean)
    m_TextBox.Enabled = Not LockVal
    m_SpinButton.Enabled = Not LockVal
End Function

Public Function UserModifyingData() As Boolean
    UserModifyingData = boolUserModifying
End Function

Public Function ClearUserModifyFlag()
    
    
    If Now < TimeInterval Then
        ScheduleNextTrigger False
    End If
    
    TimeInterval = Empty
    Debug.Print ("FlagClear JobID: " & BoundJob.JobID)
    WriteLogToFile (DateTime.Now & " FlagClear JobID: " & BoundJob.JobID)
    ProcessChanges
    
    boolUserModifying = False
End Function

Private Sub ScheduleNextTrigger(Optional ByVal KeepTracking As Boolean = True)
    If KeepTracking Then
        TimeInterval = Now + TimeValue("00:00:05")
        Application.OnTime TimeInterval, "'Utilities.ControlGroupTimerCall " & BoundJob.JobID & "'"
    Else:
        'To help check for different erros that are not the expected one.
        On Error Resume Next
        Application.OnTime TimeInterval, "'Utilities.ControlGroupTimerCall " & BoundJob.JobID & "'", , False
        If Err.Number <> 0 Then
            Debug.Print (Err.Number & ": " & Err.Description & " 'CountableControlGroup'")
        End If
        On Error GoTo 0
        
        Set TimeInterval = Nothing
    End If
End Sub

Private Function IControlGroupClass_UserModifyingData() As Boolean
    IControlGroupClass_UserModifyingData = UserModifyingData
End Function

'Events

Private Sub m_BoundJob_DatabaseUpdateStart()
    LockEditing True
End Sub

Private Sub m_BoundJob_DatabaseUpdateComplete()
    
    boolLoadingJob = True
    
    m_TextBox.Value = BoundJob.JobCount
    numVal = BoundJob.JobCount
    
    boolLoadingJob = False
    LockEditing False
End Sub

Private Sub m_TextBox_Change()
    Dim BoolProcessChange As Boolean
    
    If Not boolLoadingJob Then
'        If Not boolUserModifying Then
            numVal = BoundJob.JobCount
            
            If InStr(1, m_TextBox.Value, ".", vbTextCompare) > 0 Then
                m_TextBox.Value = Replace(m_TextBox.Value, ".", "", , , vbTextCompare)
            End If
            
            textVal = m_TextBox.Value
            
            Debug.Print (DateTime.Now & " numVal: " & numVal & " textVal: " & textVal)
            
            WriteLogToFile (DateTime.Now & " numVal: " & numVal & " textVal: " & textVal)
            
            If numVal <> textVal Then
                If IsNumeric(textVal) Then
                
                    If textVal > 2147483647 Then
                        MsgBox ("Number entered is larger than humanly possible.")
                        m_TextBox.Value = numVal
                        Exit Sub
                    End If
                    
                    numVal = Abs(textVal)
                    m_TextBox.Text = CStr(numVal)
                    BoolProcessChange = True
                Else
                    If textVal = vbNullString Then
                        m_TextBox.Text = 0
                        numVal = 0
                    Else:
                        'CountTextBox.Text = CStr(Abs(numVal))
                        'm_TextBox.Text = CStr(numVal)
                        m_TextBox.Text = CStr(BoundJob.JobCount)
                        numVal = BoundJob.JobCount
                    End If
                End If
                
                If BoolProcessChange Then

                    If Not boolUserModifying Then
                        ProcessChanges
                    End If
                End If
            

            End If
'        End If
    End If
End Sub

Private Function ProcessChanges()
    'If Not boolUserModifying Then
        'if numVal <> BoundJob.JobCount
        
        Debug.Print (DateTime.Now & " ProcessChanges1: JobID: " & BoundJob.JobID & ", JFID: " & BoundJob.JFID & ", CurBoundJobCnt: " & BoundJob.JobCount & ", numVal: " & numVal & ", JobCntDeltaOld: " & BoundJob.JobCountDelta)
        
        WriteLogToFile (DateTime.Now & " ProcessChanges1: JobID: " & BoundJob.JobID & ", JFID: " & BoundJob.JFID & ", CurBoundJobCnt: " & BoundJob.JobCount & ", numVal: " & numVal & ", JobCntDeltaOld: " & BoundJob.JobCountDelta)
        
        BoundJob.JobCountDelta = BoundJob.JobCount - numVal
        'BoundJob.JobCount = numVal
        BoundJob.JobCount = BoundJob.JobCount - BoundJob.JobCountDelta
        
        Debug.Print (DateTime.Now & " ProcessChanges2: JobCntDeltaNew: " & BoundJob.JobCountDelta & ", NewBoundJobCnt: " & BoundJob.JobCount)
        
        WriteLogToFile (DateTime.Now & " ProcessChanges2: JobCntDeltaNew: " & BoundJob.JobCountDelta & ", NewBoundJobCnt: " & BoundJob.JobCount)
        
        BoundJob.LastCountUpdate = DateTime.Now
        WriteJobDataToDB BoundJob
    'End If
End Function

Private Sub m_SpinButton_SpinDown()
    'boolUserModifying = True
    
    UserModFlag
    
    Debug.Print ("SpinBtn Down Event: JobID: " & BoundJob.JobID & " JFID: " & BoundJob.JFID)
    WriteLogToFile (DateTime.Now & " SpinBtn Down Event: JobID: " & BoundJob.JobID & " JFID: " & BoundJob.JFID)
    
    If Not m_TextBox.Text - 1 < 0 Then
        m_TextBox.Text = m_TextBox.Text - 1
    End If
    'boolUserModifying = False
End Sub

Private Sub m_SpinButton_SpinUp()
    'boolUserModifying = True
    
    UserModFlag
    
    Debug.Print ("SpinBtn Up Event: JobID: " & BoundJob.JobID & " JFID: " & BoundJob.JFID)
    WriteLogToFile (DateTime.Now & " SpinBtn Up Event: JobID: " & BoundJob.JobID & " JFID: " & BoundJob.JFID)
    
    If m_TextBox.Text + 1 > 2147483647 Then
        MsgBox ("Number entered is larger than humanly possible.")
    Else:
        m_TextBox.Text = m_TextBox.Text + 1
    End If
    'boolUserModifying = False
End Sub

'Get/Let/Set
Public Property Get BoundJob() As JobsDataClass
    Set BoundJob = m_BoundJob
End Property

Public Property Set BoundJob(ByVal JobData As JobsDataClass)
    JobData.ControlGroup = Me
    Set m_BoundJob = JobData
    numVal = JobData.JobCount
End Property

Public Function GetGroupFrame() As MSForms.Frame
    Set GetGroupFrame = m_Frame
End Function

Public Property Get GroupName() As String
    GroupName = m_GroupName
End Property

Public Property Let GroupName(ByVal strNewValue As String)
    m_GroupName = strNewValue
End Property

'Interface Implementations
Private Function IControlGroupClass_GetGroupFrame() As MSForms.IOptionFrame
    Set IControlGroupClass_GetGroupFrame = GetGroupFrame
End Function

Private Property Get IControlGroupClass_BoundJob() As JobsDataClass
    IControlGroupClass_BoundJob = BoundJob
End Property

Private Property Let IControlGroupClass_BoundJob(RHS As JobsDataClass)
    Set BoundJob = RHS
End Property

Private Sub IControlGroupClass_CreateControlFrame(TargetFrame As MSForms.IOptionFrame)
    CreateControlFrame TargetFrame
End Sub

Private Property Let IControlGroupClass_GroupName(ByVal RHS As String)
    GroupName = RHS
End Property

Private Property Get IControlGroupClass_GroupName() As String
    IControlGroupClass_GroupName = GroupName
End Property

Private Sub IControlGroupClass_SetFramePlacement(LastGroupFrame As IControlGroupClass)
    SetFramePlacement LastGroupFrame
End Sub

Private Sub m_TextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    UserModFlag
    
    Debug.Print ("TxtBx KeyDown Event: JobID: " & BoundJob.JobID & " JFID: " & BoundJob.JFID & " KeyVal: " & KeyCode & " ShiftVal: " & Shift)
    WriteLogToFile (DateTime.Now & " TxtBx KeyDown Event: JobID: " & BoundJob.JobID & " JFID: " & BoundJob.JFID & " KeyVal: " & KeyCode & " ShiftVal: " & Shift)
End Sub

Private Function UserModFlag()
    boolUserModifying = True
    Debug.Print ("FlagSet JobID: " & BoundJob.JobID)
    
    WriteLogToFile (DateTime.Now & " FlagSet JobID: " & BoundJob.JobID)
    
    If Not IsEmpty(TimeInterval) Then
        ScheduleNextTrigger False
        ScheduleNextTrigger
    Else:
        ScheduleNextTrigger
    End If
End Function

Private Sub m_TextBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    If KeyCode = vbKeyUp Then
'        m_SpinButton_SpinUp
'        m_SpinButton.SetFocus
'    ElseIf KeyCode = vbKeyDown Then
'        m_SpinButton_SpinDown
'        m_SpinButton.SetFocus
'    End If
End Sub

Private Sub m_TextBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Debug.Print ("TxtBx MouseDwn Event: JobID: " & BoundJob.JobID & " JFID: " & BoundJob.JFID & " ButtonVal: " & Button & " ShiftVal: " & Shift & " CordsX/Y: " & X & "," & Y)
    WriteLogToFile (DateTime.Now & " TxtBx MouseDwn Event: JobID: " & BoundJob.JobID & " JFID: " & BoundJob.JFID & " ButtonVal: " & Button & " ShiftVal: " & Shift & " CordsX/Y: " & X & "," & Y)
    UserModFlag
End Sub
