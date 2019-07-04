VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISearchable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'This is the Abstract Interface used for the SearchableCombobox Function class
'By Derek Pierce (10/4/2012)
'v1.2 (3/6/2014)
'
'Used to distinguish data classes as Searchable for the SearchableCombobox class

Public Function ObjectIdentifier() As Variant
'returns the Identifier for the data

'Public Function ISearchable_ObjectIdentifier() As Variant
'    ISearchable_ObjectIdentifier = 'Item in data class to be returned
'End Function

End Function

Public Function ObjectDisplayedData() As Variant
'Returns human readable name, or discription text.
'This part is what is searched on.

'Public Function Isearchable_ObjectDisplayedData() As Variant
'    Isearchable_ObjectDisplayedData = 'Item in data class to be returned
'End Function

End Function
