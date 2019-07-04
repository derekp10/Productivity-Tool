Attribute VB_Name = "Utilities"
Option Explicit

'Populates ComboBox with Reftable data
Public Function PopulateRefDataToComboBox(ByRef TargetComboBox As MSForms.ComboBox, TargetRefTableCollection As Collection, Optional BlankStarter As Boolean = False)
    Dim arrRefTableData()
    Dim lCntr As Long
    Dim RefTDataCls As RefTableDataClass
    
    TargetComboBox.Clear
    
    If BlankStarter Then
        ReDim arrRefTableData(TargetRefTableCollection.Count, 1)
    
        lCntr = 0
        
        arrRefTableData(lCntr, 0) = 0
        arrRefTableData(lCntr, 1) = "All"
        lCntr = lCntr + 1
    Else:
        ReDim arrRefTableData(TargetRefTableCollection.Count - 1, 1)
        lCntr = 0
    End If
    
    For Each RefTDataCls In TargetRefTableCollection
        arrRefTableData(lCntr, 0) = RefTDataCls.RefTypeID
        arrRefTableData(lCntr, 1) = RefTDataCls.RefTypeName
        lCntr = lCntr + 1
    Next RefTDataCls
    
    TargetComboBox.List = arrRefTableData
    
End Function

'Populates ListBox with Reftable data
Public Function PopulateRefDataToListBox(ByRef TargetListBox As MSForms.ListBox, TargetRefTableCollection As Collection, Optional BlankStarter As Boolean = False)
    Dim arrRefTableData()
    Dim lCntr As Long
    Dim RefTDataCls As RefTableDataClass
    
    TargetListBox.Clear
    
    If BlankStarter Then
        ReDim arrRefTableData(TargetRefTableCollection.Count, 1)
    
        lCntr = 0
        
        arrRefTableData(lCntr, 0) = 0
        arrRefTableData(lCntr, 1) = "All"
        lCntr = lCntr + 1
    Else:
        ReDim arrRefTableData(TargetRefTableCollection.Count - 1, 1)
        lCntr = 0
    End If
    
    For Each RefTDataCls In TargetRefTableCollection
        arrRefTableData(lCntr, 0) = RefTDataCls.RefTypeID
        arrRefTableData(lCntr, 1) = RefTDataCls.RefTypeName
        lCntr = lCntr + 1
    Next RefTDataCls
    
    TargetListBox.List = arrRefTableData
    
End Function

Public Function InCollection(colCollection As Collection, ByVal strItemToCheck As String) As Boolean
    'Compares value provided to collection provided key's refrence. On error (no item found) returns false, else true
    ' a.k.a Key found.
    On Error GoTo HandleError
    'On Error Resume Next
    Dim var As Variant
    var = colCollection(strItemToCheck)
    InCollection = True
    Exit Function
HandleError:
    If Err.Number <> 438 Then
        InCollection = False
    Else:
        InCollection = True
    End If
End Function

Public Function InControls(ControlsGoup As Controls, ByVal strItemToCheck As String) As Boolean
    'This is a rif on InCollection, only designed to work with the Controls object in MSForm.Controls
    'Compares value provided to collection provided key's refrence. On error (no item found) returns false, else true
    ' a.k.a Key found.
    On Error GoTo HandleError
    'On Error Resume Next
    Dim var As Variant
    var = ControlsGoup(strItemToCheck)
    InControls = True
    Exit Function
HandleError:
    If Err.Number <> 438 Then
        InControls = False
    Else:
        InControls = True
    End If
End Function

Public Function InFields(colCollection As ADODB.Fields, ByVal strItemToCheck As String) As Boolean
    'This is a rif on InCollection, only designed to work with the Fields object in ADODB Recordsets
    'Need to watch the 438 Error check as it may cause issues with the "Fields" Collection
    'Compares value provided to collection provided key's refrence. On error (no item found) returns false, else true
    ' a.k.a Key found.
    On Error GoTo HandleError
    'On Error Resume Next
    Dim var As Variant
    var = colCollection(strItemToCheck).Name
    InFields = True
    Exit Function
HandleError:
    If Err.Number <> 438 Then
        InFields = False
    Else:
        InFields = True
    End If
End Function
