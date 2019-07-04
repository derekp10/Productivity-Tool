VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DateSelectUF 
   Caption         =   "Date Selection"
   ClientHeight    =   10230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5430
   OleObjectBlob   =   "DateSelectUF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DateSelectUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private DTPStart As GenericDateTimePicker
Private DTPEnd As GenericDateTimePicker
Private UsersSCB As SearchableComboboxClass
Private JobsSCB As SearchableComboboxClass
Private RefTableDC As RefTableDataClass

Private Sub AllDatesCheckBox_Click()
    CheckAllDates
End Sub

Private Sub DayCommandButton_Click()
    DTPStart.DateValue = DateAdd("h", -24, DTPEnd.DateValue)
End Sub

Private Sub JFRemoveCommandButton_Click()
    ListBoxRemove JobFunctionListBox
End Sub

Private Function ListBoxRemove(ByRef TargetListBox As MSForms.ListBox)
    Dim lCnt As Long
    
    For lCnt = TargetListBox.ListCount - 1 To 0 Step -1
        If TargetListBox.Selected(lCnt) = True Then
            TargetListBox.RemoveItem (lCnt)
        End If
    Next lCnt
End Function

Private Sub JobFunctionComboBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    JobsSCB.KeyDownEvent KeyCode, Shift
End Sub

Private Sub JobFunctionComboBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    JobsSCB.KeyPressEvent
End Sub

Private Sub JobFunctionComboBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    JobsSCB.KeyUpEvent KeyCode, Shift
    
    AddToTargetListbox KeyCode, Shift, JobFunctionComboBox, JFSelectCommandButton, JobFunctionListBox
    
End Sub

Private Sub MonthCommandButton_Click()
    DTPStart.DateValue = DateAdd("m", -1, DTPEnd.DateValue)
End Sub

Private Sub SearchCommandButton_Click()
    Dim SDate As Date
    Dim EDate As Date
    
    Debug.Print (GetSelectedItems(JobFunctionListBox))
    Debug.Print (GetSelectedItems(UserListBox))
        
    SelectListBoxData UserListBox
    SelectListBoxData JobFunctionListBox
        
    If AllDatesCheckBox = True Then
        SDate = EMPTY_DATE
        EDate = EMPTY_DATE
    Else:
        SDate = DTPStart.DateTimeValue
        EDate = DTPEnd.DateTimeValue
    End If
    
    If IncludeInterruptionsCheckBox <> True Then
        BuildDataList SDate, EDate, GetSelectedItems(UserListBox), GetSelectedItems(JobFunctionListBox), StatusComboBox.Value
    Else:
        BuildDataListWithInterruptions SDate, EDate, GetSelectedItems(UserListBox), GetSelectedItems(JobFunctionListBox), StatusComboBox.Value
    End If
    
    UserForm_Terminate
End Sub

Private Function SelectListBoxData(ByRef TargetListBox As MSForms.ListBox)
    Dim lCnt As Long
    For lCnt = 0 To TargetListBox.ListCount - 1
        TargetListBox.Selected(lCnt) = True
    Next lCnt
End Function

Private Sub JFSelectCommandButton_Click()
    ListBoxSelect JobFunctionListBox, JobFunctionComboBox, rte_JobFunctionRefTable
End Sub

Private Function ListBoxSelect(ByRef TargetListBox As MSForms.ListBox, ByRef TargetComboBox As MSForms.ComboBox, ByVal TargetRefTable As RefTableEnu)
    TargetListBox.AddItem
    TargetListBox.List(UBound(TargetListBox.List), 0) = TargetComboBox.Value
    TargetListBox.List(UBound(TargetListBox.List), 1) = RefTableMng.GetStringFromID(TargetRefTable, TargetComboBox.Value)
End Function

Private Sub TwoDayCommandButton_Click()
    DTPStart.DateValue = DateAdd("h", -48, DTPEnd.DateValue)
End Sub

Private Sub UserForm_Initialize()

    If RefTableMng Is Nothing Then
        Set RefTableMng = New ReferenceTableManager
    End If
    
    Set DTPStart = New GenericDateTimePicker
    Set DTPEnd = New GenericDateTimePicker
    
    DTPStart.AddUFElement StartMonthComboBox
    DTPStart.AddUFElement StartDayComboBox
    DTPStart.AddUFElement StartYearComboBox
    DTPStart.AddUFElement StartHourComboBox
    DTPStart.AddUFElement StartMinComboBox
    
    DTPEnd.AddUFElement EndMonthComboBox
    DTPEnd.AddUFElement EndDayComboBox
    DTPEnd.AddUFElement EndYearComboBox
    DTPEnd.AddUFElement EndHourComboBox
    DTPEnd.AddUFElement EndMinComboBox
    
    DTPStart.DateTimeValue = DateTime.Now
    DTPStart.TimeValue = "00:00:00"
    DTPEnd.DateTimeValue = DateTime.Now
    DTPEnd.TimeValue = "23:59:59"
    
    Set UsersSCB = New SearchableComboboxClass
    Set JobsSCB = New SearchableComboboxClass
    
    UsersSCB.SetClassParameters RefTableMng.GetRefCol(rte_UsersRefTable), RefTableDC, UserNameComboBox
    JobsSCB.SetClassParameters RefTableMng.GetRefCol(rte_JobFunctionRefTable), RefTableDC, JobFunctionComboBox
    
    
    
    PopulateRefDataToComboBox UserNameComboBox, RefTableMng.GetRefCol(rte_UsersRefTable, rtkt_RefTypeRefID), True
    UserNameComboBox.ListIndex = 0
    
'    PopulateRefDataToListBox JobFunctionListBox, RefTableMng.GetRefCol(rte_JobFunctionRefTable, rtkt_RefTypeRefID), True
'    JobFunctionListBox.Selected(0) = True
    
    PopulateRefDataToComboBox JobFunctionComboBox, RefTableMng.GetRefCol(rte_JobFunctionRefTable, rtkt_RefTypeRefID), True
    JobFunctionComboBox.ListIndex = 0
    
    PopulateRefDataToComboBox StatusComboBox, RefTableMng.GetRefCol(rte_StatusRefTable, rtkt_RefTypeRefID), True
    StatusComboBox.ListIndex = 0
    
    CheckAllDates
    
End Sub

Private Sub CancelCommandButton_Click()
    UserForm_Terminate
End Sub

Private Sub UserForm_Terminate()
    Me.Hide
End Sub

Public Function GetSelectedItems(ByRef lBox As MSForms.ListBox) As String
'returns an array of selected items in a ListBox
'http://stackoverflow.com/questions/19551754/how-do-i-return-multi-select-listbox-values-into-a-sentence-using-word-vba
Dim tmpArray() As Variant
Dim i As Integer
Dim selCount As Integer
    selCount = -1
    '## Iterate over each item in the ListBox control:
    For i = 0 To lBox.ListCount - 1
        '## Check to see if this item is selected:
        If lBox.Selected(i) = True Then
            '## If this item is selected, then add it to the array
            selCount = selCount + 1
            ReDim Preserve tmpArray(selCount)
            tmpArray(selCount) = lBox.List(i)
        End If
    Next

    If selCount = -1 Then
        '## If no items were selected, return an empty string
        GetSelectedItems = "" ' or "No items selected", etc.
    Else:
        '## Otherwise, return the array of items as a string,
        '   delimited by commas
        GetSelectedItems = Join(tmpArray, ", ")
    End If
End Function

Public Function GetSelectedItemsRefTableLookup(ByRef lBox As MSForms.ListBox, RefTable As RefTableEnu) As String
'returns an array of selected items in a ListBox
'http://stackoverflow.com/questions/19551754/how-do-i-return-multi-select-listbox-values-into-a-sentence-using-word-vba
Dim tmpArray() As Variant
Dim i As Integer
Dim selCount As Integer
    selCount = -1
    '## Iterate over each item in the ListBox control:
    For i = 0 To lBox.ListCount - 1
        '## Check to see if this item is selected:
        If lBox.Selected(i) = True Then
            '## If this item is selected, then add it to the array
            selCount = selCount + 1
            ReDim Preserve tmpArray(selCount)
            tmpArray(selCount) = "'" & RefTableMng.GetStringFromID(RefTable, lBox.List(i)) & "'"
        End If
    Next

    If selCount = -1 Then
        '## If no items were selected, return an empty string
        GetSelectedItemsRefTableLookup = "" ' or "No items selected", etc.
    Else:
        '## Otherwise, return the array of items as a string,
        '   delimited by commas
        GetSelectedItemsRefTableLookup = Join(tmpArray, ", ")
    End If
End Function

Private Function CheckAllDates()
    DTPStart.SetEnabled Not AllDatesCheckBox
    DTPEnd.SetEnabled Not AllDatesCheckBox
    DayCommandButton.Enabled = Not AllDatesCheckBox
    TwoDayCommandButton.Enabled = Not AllDatesCheckBox
    WeekCommandButton.Enabled = Not AllDatesCheckBox
    MonthCommandButton.Enabled = Not AllDatesCheckBox
End Function

Private Sub UserNameComboBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    UsersSCB.KeyDownEvent KeyCode, Shift
End Sub

Private Sub UserNameComboBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    UsersSCB.KeyPressEvent
End Sub

Private Sub UserNameComboBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    UsersSCB.KeyUpEvent KeyCode, Shift
    
    AddToTargetListbox KeyCode, Shift, UserNameComboBox, UserSelectCommandButton, UserListBox
    
'    If KeyCode = vbKeyReturn Then
'        If Not IsNull(UserNameComboBox.Value) Then
'            UserSelectCommandButton_Click
'            'UserNameComboBox.Value = ""
'            'PopulateRefDataToComboBox UserNameComboBox, RefTableMng.GetRefCol(rte_UsersRefTable, rtkt_RefTypeRefID), True
'            UserNameComboBox.ListIndex = 0
'        End If
'    End If
End Sub

Private Function AddToTargetListbox(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer, ByRef TargetComboBox As MSForms.ComboBox, ByRef TargetSelectButton As MSForms.CommandButton, ByRef TargetListBox As MSForms.ListBox)
    If KeyCode = vbKeyReturn Then
        If Not IsNull(TargetComboBox.Value) Then
            TargetSelectButton.Value = True
            'UserNameComboBox.Value = ""
            'PopulateRefDataToComboBox UserNameComboBox, RefTableMng.GetRefCol(rte_UsersRefTable, rtkt_RefTypeRefID), True
            TargetComboBox.ListIndex = 0
        End If
    End If
End Function


Private Sub UserRemoveCommandButton_Click()
    ListBoxRemove UserListBox
End Sub

Private Sub UserSelectCommandButton_Click()
    If Not IsNull(UserNameComboBox.Value) Then
        ListBoxSelect UserListBox, UserNameComboBox, rte_UsersRefTable
    End If
End Sub

Private Sub WeekCommandButton_Click()
    DTPStart.DateValue = DateAdd("d", -7, DTPEnd.DateValue)
End Sub
