VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RefTableDataClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'RefTableDataClass
'By Derek Pierce (8/11/2015)
'Ver. 1.0 8/11/2015 First Reusable Build
'Ver. 1.1 3/28/2015

'Generic data class that can represent any form of RefTableData (2 or 3+ fields of data found in a reftable in the database)
'RefTables follow this template.
'Field(0) = ID/PK (Generic Field Name:(RefID))
'Field(1) = String representation/value assigned to ID/PK(Field(1). (Generic Field Name:(RefName))
'Field(3+) = Misc data pertaining to the entry. (Discriptions, Canned Comments, etc)

'Implements ISearchable so that it's data can be used by
'SearchableComboBox class

Implements ISearchable

Private p_RefTypeID As Long '-1 = Assign new PK.
Private p_RefTypeName As String
'Private p_RefTypeOrder As Long
Private p_RefTypeExtraData As RefTableExtraDataCollection

Private Sub Class_Initialize()
    Set p_RefTypeExtraData = New RefTableExtraDataCollection
    p_RefTypeID = -1
End Sub

'Standard getters and setters (Get/Let) for the various data elements in the table.(fields)

Public Property Get RefTypeID() As Long
    RefTypeID = p_RefTypeID
End Property

Public Property Let RefTypeID(ByVal vNewValue As Long)
    p_RefTypeID = vNewValue
End Property

Public Property Get RefTypeName() As String
    RefTypeName = p_RefTypeName
End Property

Public Property Let RefTypeName(ByVal vNewValue As String)
    p_RefTypeName = vNewValue
End Property

'Public Property Get RefTypeOrder() As Long
'    RefTypeOrder = p_RefTypeOrder
'End Property
'
'Public Property Let RefTypeOrder(ByVal vNewValue As Long)
'    p_RefTypeOrder = vNewValue
'End Property


Public Property Get RefTypeExtraData() As RefTableExtraDataCollection
    Set RefTypeExtraData = p_RefTypeExtraData
End Property

Public Property Let RefTypeExtraData(ByVal colNewValue As RefTableExtraDataCollection)
    Set p_RefTypeExtraData = colNewValue
End Property

Private Function ISearchable_ObjectIdentifier() As Variant
    ISearchable_ObjectIdentifier = RefTypeID
End Function

Private Function ISearchable_ObjectDisplayedData() As Variant
    ISearchable_ObjectDisplayedData = RefTypeName
End Function
