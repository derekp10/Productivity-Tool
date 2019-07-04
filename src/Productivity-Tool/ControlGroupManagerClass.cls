VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlGroupManagerClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_JobDataCol As Collection
Private m_ControlGroupCol As Collection
Private m_LastGroupFrame As IControlGroupClass
Private m_TempCountableControlGroup As CountableControlGroupClass
Private m_FrameForControlGroups As MSForms.Frame

Private Sub Class_Initialize()
    Set m_ControlGroupCol = New Collection
    Set m_LastGroupFrame = Nothing
End Sub

Public Property Get JobDataCollection() As Collection
    Set JobDataCollection = m_JobDataCol
End Property

Public Property Let JobDataCollection(ByVal NewCollection As Collection)
    Set m_JobDataCol = NewCollection
    ResetFrameForControlGroups
    Set m_ControlGroupCol = New Collection
    BindJobsToControlGroups
End Property

Public Property Get FrameForControlGroups() As MSForms.Frame
    Set FrameForControlGroups = m_FrameForControlGroups
End Property

Public Property Let FrameForControlGroups(ByRef TargetFrame As MSForms.Frame)
    Set m_FrameForControlGroups = TargetFrame
End Property

Private Function BindJobsToControlGroups()
    Dim m_TempJobData As JobsDataClass
    Dim m_TempGroupFrame As IControlGroupClass
    Dim m_TempRefTableExtraDataCol As RefTableExtraDataCollection
    
    If m_JobDataCol Is Nothing Then
        MsgBox ("JobsDataCollection is empty")
        Exit Function
    End If
    
    If m_FrameForControlGroups Is Nothing Then
        MsgBox ("FrameForControlGroups not set.")
        Exit Function
    End If
    
    For Each m_TempJobData In m_JobDataCol
        Set m_TempRefTableExtraDataCol = RefTableMng.GetExtraDataFromID(rte_JobFunction, m_TempJobData.JFID)
        
        Select Case m_TempRefTableExtraDataCol.GetValueForFieldName("ControlGroupTypeID")
        Case 1:
            Set m_TempGroupFrame = New CountableControlGroupClass
        Case 2:
            Set m_TempGroupFrame = New LabelControlGroupClass
        Case Else:
            MsgBox ("Unimplemented ControlGroup select on job: " & Chr(34) & RefTableMng.GetStringFromID(rte_JobFunction, m_TempJobData.JFID) & Chr(34) & "." & _
                        Chr(13) & "Please contact the developer.")
            Set m_TempGroupFrame = New UnImplementedControlGroupClass
        End Select
        
        m_TempGroupFrame.BoundJob = m_TempJobData
        
        m_TempGroupFrame.GroupName = RefTableMng.GetStringFromID(rte_JobFunction, m_TempJobData.JFID)
        
        If Not m_LastGroupFrame Is Nothing Then
            m_TempGroupFrame.SetFramePlacement m_LastGroupFrame
        End If
        
        m_TempGroupFrame.CreateControlFrame m_FrameForControlGroups
        
        Set m_LastGroupFrame = m_TempGroupFrame
        
        m_ControlGroupCol.Add m_TempGroupFrame, CStr(m_TempGroupFrame.GroupName)
    Next m_TempJobData
    
End Function

Public Function AddJob(ByRef JobData As JobsDataClass)
    Dim m_TempGroupFrame As IControlGroupClass
    Dim m_TempRefTableExtraDataCol As RefTableExtraDataCollection
    
    Set m_TempRefTableExtraDataCol = RefTableMng.GetExtraDataFromID(rte_JobFunction, JobData.JFID)
    
    Select Case m_TempRefTableExtraDataCol.GetValueForFieldName("ControlGroupTypeID")
    Case 1:
        Set m_TempGroupFrame = New CountableControlGroupClass
    Case 2:
        Set m_TempGroupFrame = New LabelControlGroupClass
    Case Else:
        MsgBox ("Unimplemented ControlGroup select on job: " & Chr(34) & RefTableMng.GetStringFromID(rte_JobFunction, JobData.JFID) & Chr(34) & "." & _
                    Chr(13) & "Please contact the developer.")
        Set m_TempGroupFrame = New UnImplementedControlGroupClass
    End Select
    
    m_JobDataCol.Add JobData, CStr(JobData.JobID)
    
    m_TempGroupFrame.BoundJob = JobData
    
    m_TempGroupFrame.GroupName = RefTableMng.GetStringFromID(rte_JobFunction, JobData.JFID)
    
    If Not m_LastGroupFrame Is Nothing Then
        m_TempGroupFrame.SetFramePlacement m_LastGroupFrame
    End If
    
    m_TempGroupFrame.CreateControlFrame m_FrameForControlGroups
    
    Set m_LastGroupFrame = m_TempGroupFrame
    
    m_ControlGroupCol.Add m_TempGroupFrame, CStr(m_TempGroupFrame.GroupName)
    
End Function

Public Function IsJobInGroup(ByRef JobData As JobsDataClass) As Boolean
    Dim rtnBln As Boolean
    
    If InCollection(m_ControlGroupCol, CStr(RefTableMng.GetStringFromID(rte_JobFunction, JobData.JFID))) Then
        rtnBln = True
    End If
    
    IsJobInGroup = rtnBln
    
End Function

Public Function GetScrollHeight()
    'Should probably get the 5 from the m_lastGroupFrame
    If Not m_LastGroupFrame Is Nothing Then
        GetScrollHeight = m_LastGroupFrame.GetGroupFrame.Top + m_LastGroupFrame.GetGroupFrame.Height + 5
    Else
        GetScrollHeight = 0
    End If
End Function

Private Function ResetFrameForControlGroups()
    Set m_LastGroupFrame = Nothing
    'Doesn't work on controls bound to data or created in designer(errors out)
    'per Excel Developer Reference (help)
    m_FrameForControlGroups.Controls.Clear
End Function

Public Sub SetFocusTarget(ByVal JobFunctionID As Long)
    Dim tmpControlGroup As IControlGroupClass
    Dim tmpJobName As String
    
    tmpJobName = RefTableMng.GetStringFromID(rte_JobFunction, JobFunctionID)
    
    If InCollection(m_ControlGroupCol, tmpJobName) Then
        Set tmpControlGroup = m_ControlGroupCol.Item(tmpJobName)
        
        tmpControlGroup.MakeGroupFocusTarget
    End If

End Sub
