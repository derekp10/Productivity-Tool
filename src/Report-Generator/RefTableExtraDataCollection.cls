VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RefTableExtraDataCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'RefTableExtraDataCollection
'By Derek Pierce (8/11/2015)
'Ver. 1.0 8/11/2015 First Reusable Build
'Ver. 1.1 3/28/2015
'Used to store and access fields other than RefID and RefName that are not known
'at run time, or very between tables.

Private m_FieldNameCol As Collection
Private m_FieldValueCol As Collection

Public Enum RefTableExtraDataType
    rtedt_FieldValue = 1
    rtedt_FieldNames
End Enum

Private Sub Class_Initialize()
    Set m_FieldValueCol = New Collection
    Set m_FieldNameCol = New Collection
End Sub

Public Sub Add(ByVal Value As Variant, Key As String)
    m_FieldNameCol.Add Key, CStr(Key)
    m_FieldValueCol.Add Value, CStr(Key)
End Sub

Public Function oldItem(ByVal Index As Variant, Optional ExtraDataType As RefTableExtraDataType = rtedt_FieldValue) As Variant
    Dim rtnValue As Variant

    If ExtraDataType = rtedt_FieldValue Then
        rtnValue = m_FieldValueCol.Item(CStr(Index))
    ElseIf ExtraDataType = rtedt_FieldNames Then
        rtnValue = m_FieldNameCol.Item(CStr(Index))
    End If

    oldItem = rtnValue
End Function

Public Property Get Item(ByVal Index As Variant, ExtraDataType As RefTableExtraDataType) As Variant
    Dim rtnValue As Variant
    'TODO: Write this Better
    If InCollection(m_FieldNameCol, CStr(Index)) Then
        If ExtraDataType = rtedt_FieldValue Then
            rtnValue = m_FieldValueCol.Item(CStr(Index))
        ElseIf ExtraDataType = rtedt_FieldNames Then
            rtnValue = m_FieldNameCol.Item(CStr(Index))
        End If
    Else:
        Debug.Print ("Invalid Index value used when attempting to add new value to RefTableExtraDataCollection class" & _
                        Chr(13) & "IndexUsed: " & Index)
    End If

    Item = rtnValue
End Property

Public Property Let Item(ByVal Index As Variant, ExtraDataType As RefTableExtraDataType, vNewValue As Variant)
    'TODO: write this better
    If InCollection(m_FieldNameCol, CStr(Index)) Then
        If ExtraDataType = rtedt_FieldValue Then
            m_FieldValueCol.Remove (CStr(Index))
            m_FieldValueCol.Add vNewValue, CStr(Index)
        ElseIf ExtraDataType = rtedt_FieldNames Then
            m_FieldNameCol.Remove (CStr(Index))
            m_FieldNameCol.Add vNewValue, CStr(Index)
        End If
    Else:
        Debug.Print ("Invalid Index value used when attempting to add new value to RefTableExtraDataCollection class" & _
                        Chr(13) & "IndexUsed: " & Index)
    End If

End Property

Public Function GetCollectionByExtraDataType(ByVal ExtraDataType As RefTableExtraDataType) As Collection
    Dim rtnCol As Collection
    If ExtraDataType = rtedt_FieldValue Then
        Set rtnCol = m_FieldValueCol
    ElseIf ExtraDataType = rtedt_FieldNames Then
        Set rtnCol = m_FieldNameCol
    End If
    
    Set GetCollectionByExtraDataType = rtnCol
End Function

Public Function GetValueForFieldName(ByVal ExtraDataFieldName As String) As Variant
    Dim rtnVar As Variant
    
    rtnVar = Item(Item(CStr(ExtraDataFieldName), rtedt_FieldNames), rtedt_FieldValue)
    
    GetValueForFieldName = rtnVar
End Function

Public Function SetValueForFieldName(ByVal ExtraDataFieldName As String, ByRef NewValue As Variant)
    'Really hacky, might not work.
    Item(Item(CStr(ExtraDataFieldName), rtedt_FieldNames), rtedt_FieldValue) = NewValue
End Function
