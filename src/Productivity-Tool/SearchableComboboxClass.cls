VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SearchableComboboxClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'SearchableCombobox class
'By Derek Pierce (10/3/2012)
'Ver. 1.3 5/8/2015
'
'Created to add a dynamic search function to MSForm.Combobox types.
'This class is to be used with the interface ISearchable to allow for any custom data types to be used.
'Creates and assigns an array based off of a collection passed to it. (Presumably the same collection you want the combo
'box to be created out of.) This is passed to the combo box's .list function.
'
'This class must be set up in order to work. Using SetClassParameters,
'this method takes a collection (made up of Objects that implement ISearchable),
'a data type that implements ISearchable, and the Combo box that it is to be bound too.
'
'Each Combo box needs its own instance of the SearchableCombobox if it is to use this feature.
'
'When setting up events, have the event call the corresponding event in this class. (KeyUp = KeyUpEvent, etc...)
'Currently the only events supported are KeyUp, KeyDown, and KeyPress.
'
'This implementation has only been tested on 2 element combo boxes. (A BoundColumn, usualy the Unique ID, and the TextColumn.)
'In order to get around the issue of not being able to assign data based off of the BoundColumn, the BoundColumn is changed
'to the TextColumn, the assignment is done based on the text, and then the combo box's BoundColumn is set back to what it was.
'
'ISearchable makes available 2 required methods to implement for this to work. ObjectIdentifier which is what BoundColumn is based on
'(And is usualy a unique identifier for that particular data.) and ObjectDisplayData which is the human readable text,
'that best describes the data.(Say you had a PersonName, and PersonID, ObjectIdentifier = PersonID (Returns rather),
'and ObjectDisplayData = PersonName.) ObjectDisplayData is the data in which the search is performed on. So in the case of names,
'the full name of a person would be best suited for this field.



'TODO: Convert Spelt Numbers to Digits, and vise versa.
'TODO: Support Multiple colums of data over 2 (ObjectIdentifier and ObjectDisplayedData)

Private objSearchableData As ISearchable
Private colValidData As Collection
Private p_TargetCollection As Collection
Private p_CboBox As MSForms.ComboBox
Private boolCtrlKey As Boolean

Public Function SetClassParameters(ByRef TargetCollection As Collection, ByVal TargetDataType As ISearchable, ByRef ComboBox As MSForms.ComboBox)
'Used to set this instance of the class to work with one specific combo box.
    Set objSearchableData = TargetDataType
    Set p_TargetCollection = TargetCollection
    Set p_CboBox = ComboBox
    
    'Set the ComboBox Properties to be compatible with this function
    ComboBox.AutoWordSelect = False
    ComboBox.BoundColumn = 1
    ComboBox.ColumnCount = 2
    ComboBox.ColumnWidths = "0pt"
    ComboBox.MatchRequired = False
    ComboBox.MatchEntry = fmMatchEntryNone
    ComboBox.TextColumn = 2
    
    
End Function

Public Function UpdateTargetCollection(ByRef TargetCollection As Collection)
'Used to update the localy stored version of the collection.
'This is usefull for handling collection refreshes/rebuilds, as apparently they are not passed and stored
'by reference in vba.

    Set p_TargetCollection = TargetCollection

End Function

Private Function ReturnTargetedArray() As Variant
'Returns a built targeted array based on the current search data.
    ReturnTargetedArray = BuildTargetedArray
End Function

Public Function SetSearchData(ByVal SearchData As Variant)
'Used to set the Search data that is used to narrow the list.
'(SearchData would usually be what the user is typing into the combo box.)
    QnDSearch p_TargetCollection, SearchData
End Function

Private Function QnDSearch(ByVal colOfData As Collection, SearchData As Variant)
'This is a quick and dirty implementation of the collection search. It uses an InStr() to check all elements in the
'specified collection (built of data types that implement ISearchable for the use of ObjectDisplayedData.)
'and adds them to an internal collection which is used to build the array used to add to the ComboBox.List function.

    
    If Not colOfData Is Nothing Or Not objSearchableData Is Nothing Then
        Set colValidData = New Collection
        
        For Each objSearchableData In colOfData
            If InStr(1, objSearchableData.ObjectDisplayedData, SearchData, vbTextCompare) <> 0 Then
                colValidData.Add objSearchableData
            End If
        Next objSearchableData
        
        'Debug.Print (colValidData.Count)
    Else:
        Debug.Print ("Class Paramaters do not appear to be set. Please make sure they have been set before calling methods from this class.")
        'MsgBox ("Class Paramaters do not appear to be set. Please make sure they have been set before calling methods from this class.")
    End If
    

End Function

Private Function BuildTargetedArray() As Variant
'Builds and returns an array based off of the internal collection, for use with ComboBox.List method.
'Returns the limited list if matching text was found, otherwise, returns the full list.
    Dim tmpArray()
    Dim lngCounter As Long
    
    lngCounter = 0
    
    If colValidData.Count <> 0 Then
        ReDim tmpArray(colValidData.Count - 1, 1)
        For Each objSearchableData In colValidData
            tmpArray(lngCounter, 0) = objSearchableData.ObjectIdentifier
            tmpArray(lngCounter, 1) = objSearchableData.ObjectDisplayedData
            lngCounter = lngCounter + 1
        Next objSearchableData
        BuildTargetedArray = tmpArray()
    Else:
        If p_TargetCollection.Count <> 0 Then
            ReDim tmpArray(p_TargetCollection.Count - 1, 1)
            For Each objSearchableData In p_TargetCollection
                tmpArray(lngCounter, 0) = objSearchableData.ObjectIdentifier
                tmpArray(lngCounter, 1) = objSearchableData.ObjectDisplayedData
                lngCounter = lngCounter + 1
            Next objSearchableData
            BuildTargetedArray = tmpArray()
        Else:
            ReDim tmpArray(0, 1)
            BuildTargetedArray = tmpArray()
        End If
    End If
    
    
End Function

Public Function KeyUpEvent(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'Used to handle the various keys and key combination that could be used while the target has focus.
'This is done to exclude certain keys from triggering the list reload (Escape, Tab, etc...)
'while allowing other keys (ctrl+V) to cause a list refresh.


'    Debug.Print ("KeyCode: " & KeyCode)
'    Debug.Print ("Shift State: " & Shift)
    Dim BoundColumnStore As String
    

    BoundColumnStore = p_CboBox.BoundColumn
    p_CboBox.BoundColumn = p_CboBox.TextColumn
    
    p_CboBox.DropDown
    
    SetSearchData (p_CboBox.Text)

    If KeyCode = 8 Or KeyCode = 46 Then 'Backspace/Delete
        p_CboBox.List = ReturnTargetedArray
'        Exit Sub
    End If

    If KeyCode = 13 Then 'Enter
    'TODO: Make this usefull?
        'p_CboBox.Value = Trim(p_CboBox.Text)
'        Exit Sub
    End If

    If Shift = 0 Then ' no control/alt/shift modifier
        If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Or _
           KeyCode = 37 Or KeyCode = 9 Or KeyCode = 27 Or _
           KeyCode = 18 Or KeyCode = 34 Or KeyCode = 33 Or _
           KeyCode = 36 Or KeyCode = 35 Or KeyCode = 45 Then
           'UP/DOWN/RIGHT
           'LEFT/TAB/ESC
           'ALT/PGDwn/PGUp
           'HOME/END/INSERT
    '        Debug.Print ("Keycheck True")
        ElseIf boolCtrlKey Then
            If KeyCode = 67 Then '"C"
                boolCtrlKey = False
            End If
            If KeyCode = 90 Then '"Z"
                boolCtrlKey = False
                p_CboBox.List = ReturnTargetedArray
            End If
            If KeyCode = 88 Then '"X"
                boolCtrlKey = False
                p_CboBox.List = ReturnTargetedArray
            End If
            If KeyCode = 86 Then '"V"
                boolCtrlKey = False
                p_CboBox.List = ReturnTargetedArray
            End If
            boolCtrlKey = False
        Else:
            p_CboBox.List = ReturnTargetedArray
        End If
    ElseIf Shift = 1 Then
        If KeyCode = 45 Then '"-"
            p_CboBox.List = ReturnTargetedArray
        Else:
            p_CboBox.List = ReturnTargetedArray
        End If
    ElseIf Shift = 2 Then
        If KeyCode = 86 Then '"V"
            p_CboBox.List = ReturnTargetedArray
        End If
        If KeyCode = 67 Then '"C"
        End If
        If KeyCode = 88 Then '"X"
            boolCtrlKey = False
            p_CboBox.List = ReturnTargetedArray
        End If
    End If
    
    p_CboBox.BoundColumn = BoundColumnStore
    
End Function

Public Function KeyDownEvent(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'Used to store the control key status. If it is pressed and one of the monitored keys in
'KeyUpEvent are used it allows the command (like paste) to execute, and then updating the list.
'The status of the bool is reset by the KeyUpEvent Under instance where a valid shortcut key was used.

    If Shift = 2 Then
        boolCtrlKey = True
    End If
End Function

Public Function KeyPressEvent()
'Used to trigger a general update to the most recent text typed in to the combobox.
'This includes setting the SearchData and then setting the targeted Combo Box's .List
'value with the new search data

    Dim BoundColumnStore As String
    
    BoundColumnStore = p_CboBox.BoundColumn
    p_CboBox.BoundColumn = p_CboBox.TextColumn
    
    SetSearchData (p_CboBox.Text)
    p_CboBox.List = ReturnTargetedArray
    
    p_CboBox.BoundColumn = BoundColumnStore
    
End Function
