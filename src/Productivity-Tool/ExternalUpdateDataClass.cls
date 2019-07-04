VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExternalUpdateDataClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_UpdateID As Long
Private p_JobID As Long
Private p_CreationDateTime As Date
Private p_JFID As Long
Private p_UserID As Long
Private p_UpdateAmount As Long


Public Property Get UpdateID() As Long
    UpdateID = p_UpdateID
End Property

Public Property Let UpdateID(ByVal vNewValue As Long)
    p_UpdateID = vNewValue
End Property

Public Property Get JobID() As Long
    JobID = p_JobID
End Property

Public Property Let JobID(ByVal vNewValue As Long)
    p_JobID = vNewValue
End Property

Public Property Get CreationDateTime() As Date
    CreationDateTime = p_CreationDateTime
End Property

Public Property Let CreationDateTime(ByVal vNewValue As Date)
    p_CreationDateTime = vNewValue
End Property

Public Property Get JFID() As Long
    JFID = p_JFID
End Property

Public Property Let JFID(ByVal vNewValue As Long)
    p_JFID = vNewValue
End Property

Public Property Get UserID() As Long
    UserID = p_UserID
End Property

Public Property Let UserID(ByVal vNewValue As Long)
    p_UserID = vNewValue
End Property

Public Property Get UpdateAmount() As Long
    UpdateAmount = p_UpdateAmount
End Property

Public Property Let UpdateAmount(ByVal vNewValue As Long)
    p_UpdateAmount = vNewValue
End Property

Public Function DatePast() As Date
    DatePast = CreationDateTime - TimeValue("09:00:00")
End Function

Public Function DateFuture() As Date
    DateFuture = CreationDateTime + TimeValue("9:00:00")
End Function

Public Function IsJobWithinValidDateWindow(ByRef JobData As JobsDataClass) As Boolean
    Dim rtnBool As Boolean
    
    If JobData.StartDateTime >= DatePast And JobData.StartDateTime <= DateFuture Then
        rtnBool = True
    End If
    
    IsJobWithinValidDateWindow = rtnBool

End Function
