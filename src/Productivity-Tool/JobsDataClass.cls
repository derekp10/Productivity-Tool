VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JobsDataClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Dataclass for holding Jobs Data

Private m_JobID As Long
Private m_JobCreationDateTime As Date
Private m_StartDateTime As Date
Private m_LastCountUpdate As Date
Private m_JFID As Long
Private m_UserID As Long
Private m_StatusID As Long
Private m_JobCount As Long
Private m_JobCountDelta As Long
Private m_ExternalMod As String

Private m_ControlGroup As IControlGroupClass

Public Event DatabaseUpdateStart()
Public Event DatabaseUpdateComplete()


Public Property Get JobID() As Long
    JobID = m_JobID
End Property

Public Property Let JobID(ByVal lNewValue As Long)
    m_JobID = lNewValue
End Property

Public Property Get StartDateTime() As Date
    StartDateTime = m_StartDateTime
End Property

Public Property Let StartDateTime(ByVal dtNewValue As Date)
    m_StartDateTime = dtNewValue
End Property

Public Property Get LastCountUpdate() As Date
    LastCountUpdate = m_LastCountUpdate
End Property

Public Property Let LastCountUpdate(ByVal dtNewValue As Date)
    m_LastCountUpdate = dtNewValue
End Property

Public Property Get JFID() As Long
    JFID = m_JFID
End Property

Public Property Let JFID(ByVal lNewValue As Long)
    m_JFID = lNewValue
End Property

Public Property Get UserID() As Long
    UserID = m_UserID
End Property

Public Property Let UserID(ByVal lNewValue As Long)
    m_UserID = lNewValue
End Property

Public Property Get StatusID() As Long
    StatusID = m_StatusID
End Property

Public Property Let StatusID(ByVal lNewValue As Long)
    m_StatusID = lNewValue
End Property

Public Property Get JobCount() As Long
    JobCount = m_JobCount
End Property

Public Property Let JobCount(ByVal lNewValue As Long)
    m_JobCount = lNewValue
End Property

Public Property Get JobCreationDateTime() As Date
    JobCreationDateTime = m_JobCreationDateTime
End Property

Public Property Let JobCreationDateTime(ByVal dtNewValue As Date)
    m_JobCreationDateTime = dtNewValue
End Property

Public Property Get JobCountDelta() As Long
    JobCountDelta = m_JobCountDelta
End Property

Public Property Let JobCountDelta(ByVal lNewValue As Long)
    m_JobCountDelta = lNewValue
End Property

Public Property Get ExternalMod() As String
    ExternalMod = m_ExternalMod
End Property

Public Property Let ExternalMod(ByVal sNewValue As String)
    m_ExternalMod = sNewValue
End Property

Public Property Get ControlGroup() As IControlGroupClass
    'Debug.Print (TypeName(m_ControlGroup) & " Store")
    Set ControlGroup = m_ControlGroup
End Property

Public Property Let ControlGroup(ByRef objControlGroup As IControlGroupClass)
    'Debug.Print (TypeName(objControlGroup))
    Set m_ControlGroup = objControlGroup
End Property

Public Function UpdateStart()
    RaiseEvent DatabaseUpdateStart
End Function

Public Function UpdateComplete()
    RaiseEvent DatabaseUpdateComplete
End Function

