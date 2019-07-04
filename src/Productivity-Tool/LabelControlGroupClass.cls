VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LabelControlGroupClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IControlGroupClass

Private WithEvents m_Frame As MSForms.Frame
Attribute m_Frame.VB_VarHelpID = -1
Private WithEvents m_Label As MSForms.Label
Attribute m_Label.VB_VarHelpID = -1

Private m_GroupName As String

Private m_BoundJob As JobsDataClass

Private m_LabelTopModifier As Double

Private m_FrameTop As Double
Private m_ControlTop As Double
Private m_FrameLeft As Double

Private m_PreviousGFHeight As Double
Private m_PreviousGFWidth As Double
Private m_PreviousGFTop As Double

Private Sub Class_Initialize()
    m_LabelTopModifier = 4
    
    m_FrameTop = 5
    m_ControlTop = 5
    m_FrameLeft = 5
    
    m_PreviousGFHeight = 0
    m_PreviousGFWidth = 0
    m_PreviousGFTop = 0
    
End Sub

Private Function IControlGroupClass_ClearUserModifyFlag() As Variant
    'Does Nothing
End Function

Private Sub IControlGroupClass_CreateControlFrame(TargetFrame As MSForms.IOptionFrame)
    Set m_Frame = TargetFrame.Controls.Add("Forms.Frame.1")
    m_Frame.Caption = ""
    m_Frame.Name = m_GroupName & "_Frame"
    
    m_Frame.Height = 33
    
    AddControls m_Frame
    
    m_Frame.Top = m_PreviousGFTop + m_PreviousGFHeight + m_FrameTop
    m_Frame.Left = m_FrameLeft
    m_Frame.Width = m_Label.Left + m_Label.Width + 8
    
    If m_BoundJob.JobCount = 0 Then
        ProcessNoCounterJob
    End If
    
End Sub

Private Function AddControls(ByRef TargetFrame As MSForms.Frame)
    Set m_Label = TargetFrame.Controls.Add("Forms.Label.1")
    
    m_Label.Name = m_GroupName & "_Label"
    
    m_Label.Caption = RefTableMng.GetStringFromID(rte_JobFunction, BoundJob.JFID)
    
    
    m_Label.Top = m_ControlTop + m_LabelTopModifier

    m_Label.Width = 212
    m_Label.WordWrap = False

    m_Label.Left = 5
    
    
End Function

Private Function ProcessNoCounterJob()
    m_BoundJob.JobCount = 1
    m_BoundJob.LastCountUpdate = m_BoundJob.StartDateTime
    
    WriteJobDataToDB m_BoundJob
End Function

Public Property Get BoundJob() As JobsDataClass
    Set BoundJob = m_BoundJob
End Property

Public Property Set BoundJob(ByVal JobData As JobsDataClass)
    JobData.ControlGroup = Me
    Set m_BoundJob = JobData
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

Private Property Let IControlGroupClass_BoundJob(RHS As JobsDataClass)
    Set BoundJob = RHS
End Property

Private Property Get IControlGroupClass_BoundJob() As JobsDataClass
    Set IControlGroupClass_BoundJob = BoundJob
End Property

Private Function IControlGroupClass_GetGroupFrame() As MSForms.IOptionFrame
    Set IControlGroupClass_GetGroupFrame = GetGroupFrame
End Function

Private Property Let IControlGroupClass_GroupName(ByVal RHS As String)
    GroupName = RHS
End Property

Private Property Get IControlGroupClass_GroupName() As String
    IControlGroupClass_GroupName = GroupName
End Property

Private Sub IControlGroupClass_MakeGroupFocusTarget()
    m_Frame.SetFocus
End Sub

Private Sub IControlGroupClass_SetFramePlacement(LastGroupFrame As IControlGroupClass)
    m_PreviousGFTop = LastGroupFrame.GetGroupFrame.Top
    m_PreviousGFWidth = LastGroupFrame.GetGroupFrame.Width
    m_PreviousGFHeight = LastGroupFrame.GetGroupFrame.Height
End Sub

Private Function IControlGroupClass_UserModifyingData() As Boolean
    'Always false as no data or controls to modify
    IControlGroupClass_UserModifyingData = False
End Function
