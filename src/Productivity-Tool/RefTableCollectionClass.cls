VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RefTableCollectionClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'RefTableCollectionClass
'By Derek Pierce(8/11/2015)
'Ver. 1.0 8/11/2015 First Reusable Build
'Ver. 1.1 3/28/2015

Private m_ColKeyTypeBy_RefTypeName As Collection
Private m_ColKeyTypeBy_RefTypeID As Collection

Public Enum RefTableKeyType
    rtkt_RefTypeRefName = 1
    rtkt_RefTypeRefID
End Enum

Private Sub Class_Initialize()
    Set m_ColKeyTypeBy_RefTypeID = New Collection
    Set m_ColKeyTypeBy_RefTypeName = New Collection
End Sub

Public Sub Add(ByRef RefTableData As RefTableDataClass)
    m_ColKeyTypeBy_RefTypeID.Add RefTableData, CStr(RefTableData.RefTypeID)
    m_ColKeyTypeBy_RefTypeName.Add RefTableData, CStr(RefTableData.RefTypeName)
End Sub

Public Function Item(ByVal Index As Variant, Optional KeyType As RefTableKeyType = rtkt_RefTypeRefName) As RefTableDataClass
    Dim rtnRefTableData As RefTableDataClass
    If KeyType = rtkt_RefTypeRefName Then
        Set rtnRefTableData = m_ColKeyTypeBy_RefTypeName.Item(CStr(Index))
    ElseIf KeyType = rtkt_RefTypeRefID Then
        Set rtnRefTableData = m_ColKeyTypeBy_RefTypeID.Item(CStr(Index))
    End If
    
    Set Item = rtnRefTableData
End Function

Public Function GetCollectionByRefType(ByVal KeyType As RefTableKeyType) As Collection
    Dim rtnCol As Collection
    If KeyType = rtkt_RefTypeRefName Then
        Set rtnCol = m_ColKeyTypeBy_RefTypeName
    ElseIf KeyType = rtkt_RefTypeRefID Then
        Set rtnCol = m_ColKeyTypeBy_RefTypeID
    End If
    
    Set GetCollectionByRefType = rtnCol
End Function

