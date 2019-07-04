VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InterruptionDataClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_InterruptID As Long
Private m_UserID As Long
Private m_InterruptDateTime As Date
Private m_InterruptTypeID As Long
Private m_InterruptionLength As Long

Public Property Get InterruptID() As Long
    InterruptID = m_InterruptID
End Property

Public Property Let InterruptID(ByVal lNewValue As Long)
    m_InterruptID = lNewValue
End Property

Public Property Get InterruptDateTime() As Date
    InterruptDateTime = m_InterruptDateTime
End Property

Public Property Let InterruptDateTime(ByVal dtNewValue As Date)
    m_InterruptDateTime = dtNewValue
End Property

Public Property Get InterruptTypeID() As Long
    InterruptTypeID = m_InterruptTypeID
End Property

Public Property Let InterruptTypeID(ByVal lNewValue As Long)
    m_InterruptTypeID = lNewValue
End Property

Public Property Get InterruptionLength() As Long
    InterruptionLength = m_InterruptionLength
End Property

Public Property Let InterruptionLength(ByVal lNewValue As Long)
    m_InterruptionLength = lNewValue
End Property

Public Property Get UserID() As Long
    UserID = m_UserID
End Property

Public Property Let UserID(ByVal lNewValue As Long)
    m_UserID = lNewValue
End Property
