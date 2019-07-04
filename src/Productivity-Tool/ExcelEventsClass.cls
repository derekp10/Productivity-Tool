VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelEventsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents thisApp As Application
Attribute thisApp.VB_VarHelpID = -1
Private newApp As Application

Private Sub Class_Initialize()
    Set thisApp = Application
    Set newApp = New Excel.Application
    
    If RefTableMng Is Nothing Then
        Set RefTableMng = New ReferenceTableManager
    End If
End Sub

Private Sub thisApp_WorkbookOpen(ByVal Wb As Workbook)
    Dim wbToOpen As String
    If InCollection(RefTableMng.GetRefCol(rte_OpenSpecial), CStr(Wb.Name)) Then
        wbToOpen = Wb.FullName
        Wb.Close
        'newApp.Wait (Now + TimeValue("0:00:05"))
        newApp.Workbooks.Open wbToOpen
        newApp.Visible = True
    End If
End Sub


