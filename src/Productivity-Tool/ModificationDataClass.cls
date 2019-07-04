VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModificationDataClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ModificationID As Long
Private m_JobID As Long
Private m_UserID As Long
Private m_ModDateTime As Date
Private m_ModElmID As Long
Private m_NewValue As String
Private m_OldValue As String

Public Property Get ModificationID() As Long
    ModificationID = m_ModificationID
End Property

Public Property Let ModificationID(ByVal lNewValue As Long)
    m_ModificationID = lNewValue
End Property

Public Property Get JobID() As Long
    JobID = m_JobID
End Property

Public Property Let JobID(ByVal lNewValue As Long)
    m_JobID = lNewValue
End Property


Public Property Get ModDateTime() As Date
    ModDateTime = m_ModDateTime
End Property

Public Property Let ModDateTime(ByVal dNewValue As Date)
    m_ModDateTime = dNewValue
End Property

Public Property Get ModElmID() As Long
    ModElmID = m_ModElmID
End Property

Public Property Let ModElmID(ByVal lNewValue As Long)
    m_ModElmID = lNewValue
End Property

Public Property Get NewValue() As String
    NewValue = m_NewValue
End Property

Public Property Let NewValue(ByVal sNewValue As String)
    m_NewValue = sNewValue
End Property

Public Property Get OldValue() As String
    OldValue = m_OldValue
End Property

Public Property Let OldValue(ByVal sNewValue As String)
    m_OldValue = sNewValue
End Property

Public Property Get UserID() As Long
    UserID = m_UserID
End Property

Public Property Let UserID(ByVal lNewValue As Long)
    m_UserID = lNewValue
End Property
