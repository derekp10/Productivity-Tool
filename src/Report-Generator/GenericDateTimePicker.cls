VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericDateTimePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'********************************
' DateTime ComboBox code
' To replace DateTimePicker on 64bit
' OS/Office versions
' Derek Pierce 4/1/16 V 0.05
'********************************

Private m_ControlCollection As Collection
Private m_ComboBoxCollection As Collection
Private m_EventTracker As Collection

Private Const M_EMPTY_DATE = "12:00:00 AM"

Private WithEvents GMonthCB As MSForms.ComboBox
Attribute GMonthCB.VB_VarHelpID = -1
Private WithEvents GDayCB As MSForms.ComboBox
Attribute GDayCB.VB_VarHelpID = -1
Private WithEvents GYearCB As MSForms.ComboBox
Attribute GYearCB.VB_VarHelpID = -1
Private WithEvents GHourCB As MSForms.ComboBox
Attribute GHourCB.VB_VarHelpID = -1
Private WithEvents GMinCB As MSForms.ComboBox
Attribute GMinCB.VB_VarHelpID = -1


Private LastValidMonth As Integer
Private LastValidDay As Integer
Private LastValidYear As Integer
Private LastValidHour As Integer
Private LastValidMin As Integer

Private Sub Class_Initialize()
    Set m_ControlCollection = New Collection
    Set m_ComboBoxCollection = New Collection
    
    LastValidMonth = 12
    LastValidDay = 31
    LastValidYear = 1899
    LastValidHour = 0
    LastValidMin = 0
    
End Sub

Public Sub MonthCB_Change()
    
    ValidateComboBoxData
    
    If (GMonthCB.Value > 0 And GMonthCB.Value < 13) Or GMonthCB.Value = "" Then
        UpdateDay
    Else:
        'GMonthCB = LastValidMonth
    End If
    
End Sub

Public Sub DayCB_Change()
    Dim ValidMonth As Integer
    Dim ValidYear As Integer
    
'    If GMonthCB.Value <> "" And IsNumeric(GMonthCB.Value) Then
'        ValidMonth = GMonthCB.Value
'    Else:
'        ValidMonth = LastValidMonth
'    End If
'
'    If GYearCB.Value <> "" And IsNumeric(GYearCB.Value) Then
'        ValidYear = GYearCB.Value
'    Else:
'        ValidYear = LastValidYear
'    End If
    
    ValidateComboBoxData
    
'    If (GDayCB.Value > 0 And (GDayCB.Value <= FindNumberOfDays(ValidMonth, ValidYear) Or GDayCB.Value < 32)) Or GDayCB.Value = "" Then
'        ValidateComboBoxData
'    Else:
'        'GDayCB.Value = LastValidDay
'    End If
    
End Sub

Public Sub YearCB_Change()
    'Need to figure out some sort of validation for this.
    'DateSerial doesn't not appear to support years over 9999
    
    ValidateComboBoxData
    
    If (GYearCB.Value < 0 And GYearCB.Value < 10000) Or GYearCB.Value = "" Then
        UpdateDay
    Else:
        'Else stuff
    End If

End Sub

Public Sub HourCB_Change()

    ValidateComboBoxData
    
    If (GHourCB.Value > -1 And GHourCB.Value < 24) Or GHourCB.Value = "" Then
        ValidateComboBoxData
    Else:
        'GHourCB.Value = FormatTimeValue(LastValidHour)
    End If
    
End Sub

Public Sub MinCB_Change()

    ValidateComboBoxData
    
    If (GMinCB.Value > -1 And GMinCB.Value < 60) Or GMinCB.Value = "" Then
        ValidateComboBoxData
    Else:
        'GMinCB.Value = FormatTimeValue(LastValidMin)
    End If
    
End Sub

Private Function BuildComboBoxData(ByVal NumElements As Integer, ByRef TargetComboBox As MSForms.ComboBox, Optional bTime As Boolean = False)
    'Fills target combobox with sequence of numbers.
    'If bTime = True (Combobox is for time)
    'Number sequence begins at 0
    Dim iCnt As Integer
    
    TargetComboBox.Clear
    
    If bTime = False Then
        For iCnt = 1 To NumElements
            TargetComboBox.AddItem iCnt
        Next iCnt
    Else:
        For iCnt = 0 To NumElements
            If iCnt < 10 Then
                TargetComboBox.AddItem "0" & iCnt
            Else:
                TargetComboBox.AddItem iCnt
            End If
        Next iCnt
    End If
    
End Function

Private Function BuildYearComboBoxData()
    'Used to auto populate the Year CB.
    'May not use this on Wednesday process. But useful elsewhere.
    Dim iCnt As Integer
    Dim iYear As Integer
    
    GYearCB.Clear
    
    iYear = DateAdd("y", -3, DateTime.Year(Date))
    
    For iCnt = 0 To 6
        GYearCB.AddItem iYear + (1 * iCnt)
    Next iCnt
    
End Function

Private Function FindNumberOfDays(ByVal TargetMonth As Integer, TargetYear As Integer) As Integer
    'DateSerial doesn't appear to support years over 9999
    If TargetYear < 10000 Then
        FindNumberOfDays = Day(DateSerial(TargetYear, TargetMonth + 1, 0))
    Else:
        FindNumberOfDays = 0
    End If
End Function

Public Function UpdateDay()
    'Updates days listed using Month/Year
    Dim iDayData As Integer
    Dim iDayCount As Integer
    
    If GDayCB.Value <> "" And IsNumeric(GDayCB.Value) Then
        iDayData = GDayCB.Value
    End If
    
    If IsNumeric(GMonthCB.Value) And IsNumeric(GYearCB.Value) Then
        iDayCount = FindNumberOfDays(GMonthCB.Value, GYearCB.Value)
    End If
    
    BuildComboBoxData iDayCount, GDayCB
    
    GDayCB.Value = iDayData
    
    ValidateComboBoxData
    
'    If Not IsDate(DateValue) Then
'        MsgBox ("Please input a valid date")
'    End If
End Function

Private Function ValidateComboBoxData()
    'Validates that data in comboboxes and notifies if not
    Dim iDayData As Integer
    Dim iDayCount As Integer

    If Not GDayCB Is Nothing Then
        If GDayCB.Value <> "" And IsNumeric(GDayCB.Value) Then
            iDayData = GDayCB.Value
        End If
        
        If IsNumeric(GMonthCB.Value) And IsNumeric(GYearCB.Value) Then
            iDayCount = FindNumberOfDays(GMonthCB.Value, GYearCB.Value)
        End If
        
        'TODO: Add check for zero on IDayData
        If iDayData <> 0 Then
            If iDayCount < iDayData Then
                If iDayCount <> 0 Then
                    'GDayCB = iDayCount
    '                GDayCB = iDayData
                    'GDayCB.BackColor = vbYellow
                    NoteCBIssue GDayCB, True
                Else:
    '                GDayCB = iDayData
                    'GDayCB.BackColor = vbWhite
                    NoteCBIssue GDayCB, False
                End If
            Else:
    '            GDayCB = iDayData
                'GDayCB.BackColor = vbWhite
                NoteCBIssue GDayCB, False
            End If
        Else:
            'GDayCB.BackColor = vbYellow
            NoteCBIssue GDayCB, True
        End If
    End If
    
    If Not GMonthCB Is Nothing Then
        If GMonthCB.Value > 12 Or GMonthCB.Value < 1 Or GMonthCB.Value = "" Then
            'GMonthCB.BackColor = vbYellow
            NoteCBIssue GMonthCB, True
        Else:
            'GMonthCB.BackColor = vbWhite
            NoteCBIssue GMonthCB, False
        End If
    End If
    
    If Not GYearCB Is Nothing Then
        If GYearCB.Value < 0 Or GYearCB.Value > 9999 Or GYearCB.Value = "" Then
            'GYearCB.BackColor = vbYellow
            NoteCBIssue GYearCB, True
        Else:
            'GYearCB.BackColor = vbWhite
            NoteCBIssue GYearCB, False
        End If
    End If
    
    If Not GHourCB Is Nothing Then
        If GHourCB.Value > 23 Or GHourCB.Value < 0 Or GHourCB.Value = "" Then
            'GHourCB.BackColor = vbYellow
            NoteCBIssue GHourCB, True
        Else:
            'GHourCB.BackColor = vbWhite
            NoteCBIssue GHourCB, False
        End If
    End If
    
    If GMinCB Is Nothing = False Then
        If GMinCB.Value > 59 Or GMinCB.Value < 0 Or GMinCB.Value = "" Then
            'GMinCB.BackColor = vbYellow
            NoteCBIssue GMinCB, True
        Else:
            'GMinCB.BackColor = vbWhite
            NoteCBIssue GMinCB, False
        End If
    End If
    
    
End Function

Private Function NoteCBIssue(ByRef TargetCB As MSForms.ComboBox, bFlag As Boolean)
    'Changes Target Combobox background color to signify an issue with entered data.
    
    If bFlag Then
        TargetCB.BackColor = vbYellow
    Else:
        TargetCB.BackColor = vbWhite
    End If
End Function


Public Function AddUFElement(ByRef UFObj As Object)
    'Add user form elements/controls to be handled by code.
    
    If Not InCollection(m_ControlCollection, UFObj.Name) Then
        m_ControlCollection.Add UFObj, UFObj.Name
        If TypeName(UFObj) = "ComboBox" Then
            CheckForKeyCombobox UFObj
        End If
    Else:
        Debug.Print ("Object " & Chr(34) & UFObj.Name & Chr(34) & " is already in the collection.")
    End If
End Function

Public Function SetVisible(ByVal bVisible As Boolean)
    'Changes the visibility flag of controls associated with this instance
    Dim obj As Object
    
    For Each obj In m_ControlCollection
        obj.Visible = bVisible
    Next obj
    
End Function

Public Function SetEnabled(ByVal bEnabled As Boolean)
    'Changes the visibility flag of controls associated with this instance
    Dim obj As Object
    
    For Each obj In m_ControlCollection
        obj.Enabled = bEnabled
    Next obj
    
End Function

Public Function SetTabStop(ByVal bTabStop As Boolean)
    'Changes the tabstop flag of controls associated with this instance
    Dim obj As Object
    Dim tmpObj As Object
    
    For Each obj In m_ControlCollection
        obj.TabStop = bTabStop
    Next obj
End Function

Private Function CheckForKeyCombobox(ByRef CBobj As Object)
    'Checks the name of the combobox to determin what type it is. (Day, Month, Hour, etc..)
    'Only works if the name contains one of those elements
    'Adds it to a collection of Important date/time comboboxes with name reference. (day, hour, etc...)
    Dim tmpCB As MSForms.ComboBox
    
    'Cast the object to type combobox
    Set tmpCB = CBobj
    
    If InStr(1, tmpCB.Name, "Month", vbTextCompare) <> 0 Then
        m_ComboBoxCollection.Add CBobj, "Month"
        Set GMonthCB = tmpCB
        BuildComboBoxData 12, tmpCB
        Exit Function
    End If
    
    If InStr(1, tmpCB.Name, "Day", vbTextCompare) <> 0 Then
        m_ComboBoxCollection.Add CBobj, "Day"
        Set GDayCB = tmpCB
        Exit Function
    End If
    
    If InStr(1, tmpCB.Name, "Year", vbTextCompare) <> 0 Then
        m_ComboBoxCollection.Add CBobj, "Year"
        Set GYearCB = tmpCB
        BuildYearComboBoxData
        Exit Function
    End If
    
    If InStr(1, tmpCB.Name, "Hour", vbTextCompare) <> 0 Then
        m_ComboBoxCollection.Add CBobj, "Hour"
        Set GHourCB = tmpCB
        BuildComboBoxData 23, tmpCB, True
        Exit Function
    End If
    
    If InStr(1, tmpCB.Name, "Min", vbTextCompare) <> 0 Then
        'This will bug out if name contains Min
        If Not InCollection(m_ComboBoxCollection, "Min") Then
            m_ComboBoxCollection.Add CBobj, "Min"
            Set GMinCB = tmpCB
            BuildComboBoxData 59, tmpCB, True
            Exit Function
        Else:
            Debug.Print ("Something with Min in the name already assigned to ComboboxCollection")
            Exit Function
        End If
    End If
    
End Function

Public Property Get DateValue() As Date
    'Gets the datevalue
    
    DateValue = CDate(GMonthCB.Value & "/" & GDayCB.Value & "/" & GYearCB.Value)
    
End Property

Public Property Let DateValue(ByVal NewDate As Date)
    'Sets the datevalue

    If IsDate(NewDate) Then
        SetValidDate (NewDate)
    End If
    
    
    GMonthCB.Value = Month(NewDate)
    GYearCB.Value = Year(NewDate)
    
    GDayCB.Value = Day(NewDate)
    
    UpdateDay
    
    
End Property

Public Property Get TimeValue() As Date
    'Gets the time value
    
    TimeValue = CDate(GHourCB.Value & ":" & GMinCB.Value)
    
End Property

Public Property Let TimeValue(ByVal NewTime As Date)
    'Sets the time value in 00:00 format
    
    If IsDate(NewTime) Then
        SetValidTime (NewTime)
    End If
    
    
    If Hour(NewTime) < 10 Then
        GHourCB.Value = "0" & Hour(NewTime)
    Else:
        GHourCB.Value = Hour(NewTime)
    End If
    
    If Minute(NewTime) < 10 Then
        GMinCB.Value = "0" & Minute(NewTime)
    Else:
        GMinCB.Value = Minute(NewTime)
    End If
    
End Property

Public Property Get DateTimeValue() As Date
    'gets date time value
    If DateValue = M_EMPTY_DATE And TimeValue = M_EMPTY_DATE Then
        DateTimeValue = M_EMPTY_DATE
    Else:
        DateTimeValue = CDate(DateValue & " " & TimeValue)
    End If
End Property

Public Property Let DateTimeValue(ByVal NewDateTime As Date)
    'sets date time value
    DateValue = NewDateTime
    TimeValue = NewDateTime
End Property

Public Function IsDataValid(Optional ByVal ShowPrompt As Boolean = False) As Boolean
    'Returns true if valid date is currently in the comboboxs.
    'Allows for a prompt to be shown when this is not true.
    Dim strDate As String
    
    If GHourCB Is Nothing And GMinCB Is Nothing Then
        strDate = GMonthCB.Value & "/" & GDayCB.Value & "/" & GYearCB.Value
    Else:
        strDate = GMonthCB.Value & "/" & GDayCB.Value & "/" & GYearCB.Value & " " & GHourCB.Value & ":" & GMinCB.Value
    End If
    
    If IsDate(strDate) Then
        IsDataValid = True
    Else:
        If ShowPrompt Then
            MsgBox ("Please use a valid date and/or time.")
        End If
        IsDataValid = False
    End If
    
End Function

Private Function SetValidDate(ByVal ValidDate As Date)

    LastValidMonth = Month(ValidDate)
    LastValidDay = Day(ValidDate)
    LastValidYear = Year(ValidDate)

End Function
Private Function SetValidTime(ByVal ValidTime As Date)

    LastValidHour = Hour(ValidTime)
    LastValidMin = Minute(ValidTime)
    
End Function

Private Function FormatTimeValue(ByVal TimeData As Integer) As String
    'Adds zeros to make number "HH" Format on single digit times
    If TimeData > -1 And TimeData < 10 Then
        FormatTimeValue = "0" & TimeData
    Else:
        FormatTimeValue = TimeData
    End If
End Function

Private Function InCollection(colCollection As Collection, ByVal strItemToCheck As String) As Boolean
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

'ComboBox Change Event handlers
Private Sub GDayCB_Change()
    DayCB_Change
End Sub

Private Sub GHourCB_Change()
    HourCB_Change
End Sub

Private Sub GMinCB_Change()
    MinCB_Change
End Sub

Private Sub GMonthCB_Change()
    MonthCB_Change
End Sub

Private Sub GYearCB_Change()
    YearCB_Change
End Sub
'/ComboBox Change Event handlers
