Attribute VB_Name = "Main"
Option Explicit

Public Sub Start()
    If RefTableMng Is Nothing Then
        Set RefTableMng = New ReferenceTableManager
    End If
    
    DateSelectUF.Show
    
End Sub


Public Function BuildDataList(ByVal StartDate As Date, ByVal EndDate As Date, ByVal UserName As String, ByVal JobFunction As String, ByVal Status As Long)
    Dim dbCon As ADODB.Connection
    Dim dbRS As ADODB.Recordset
    Dim strSQL As String
    Dim strWClause As String
    Dim lCnt As Long
    Dim lCnt2 As Long
    
    Set dbCon = GetDBCon
    
    dbCon.Open (DB_LOC & DB_NAME)
    
    Set dbRS = GetDBRS(dbCon)
    
    strSQL = "SELECT JHR.* from Jobs AS JB "
    strSQL = strSQL & "Left Join JobsHR as JHR ON JHR.JobID = JB.JobID "
    'strSQL = strSQL & "WHERE "
    
    'Generate Where Clause
    If StartDate <> EMPTY_DATE And EndDate <> EMPTY_DATE Then
        strWClause = strWClause & "JB.StartDateTime Between #" & StartDate & "# AND #" & EndDate & "# AND "
    End If
    
    If Left(UserName, 1) <> 0 And UserName <> vbNullString Then
        strWClause = strWClause & "JB.UserID in (" & UserName & ") AND "
    End If
    
    If Left(JobFunction, 1) <> 0 And JobFunction <> vbNullString Then
        strWClause = strWClause & "JB.JFID IN (" & JobFunction & ") AND "
    End If
    
    If Status <> 0 Then
        strWClause = strWClause & "JB.Status = " & Status & " AND "
    End If
    
    'Chech for where clause string data and format
    If Len(strWClause) > 0 Then
        'Remove any extra " AND " at end
        If Right(strWClause, 5) = " AND " Then
            strWClause = Left(strWClause, Len(strWClause) - 5)
        End If
        
        'Add Where
        strWClause = "WHERE " & strWClause
    End If
    
    strSQL = strSQL & strWClause
    
    dbRS.Source = strSQL
    
    dbRS.Open
    
    AddDataToDataSheet dbRS
    
'    ThisWorkbook.Worksheets("Data").Cells.Clear
'
'    ThisWorkbook.Worksheets("Data").Range("A1").Select
'
'    For lCnt = 1 To dbRS.Fields.Count
'        ThisWorkbook.Worksheets("Data").Cells(1, lCnt).Value = dbRS.Fields(lCnt - 1).Name
'    Next lCnt
'
'
'    For lCnt = 1 To dbRS.RecordCount
'        For lCnt2 = 1 To dbRS.Fields.Count
'            ThisWorkbook.Worksheets("Data").Cells(lCnt + 1, lCnt2).Value = dbRS.Fields(lCnt2 - 1).Value
'        Next lCnt2
'        dbRS.MoveNext
'    Next lCnt
'
'    ThisWorkbook.Worksheets("Data").Columns.AutoFit
        
    dbRS.Close
    dbCon.Close
    
    Set dbRS = Nothing
    Set dbCon = Nothing
    
    
End Function

Public Function BuildDataListWithInterruptions(ByVal StartDate As Date, ByVal EndDate As Date, ByVal UserName As String, ByVal JobFunction As String, ByVal Status As Long)
    Dim dbCon As ADODB.Connection
    Dim dbRS As ADODB.Recordset
    Dim strSQL As String
    Dim strWClause As String
    
    Set dbCon = GetDBCon
    
    dbCon.Open (DB_LOC & DB_NAME)
    
    Set dbRS = GetDBRS(dbCon)
    
    strSQL = "SELECT JWIHR.* from JobsWithInterruptionsUnion AS JWI "
    strSQL = strSQL & "Left Join JobsHRWithInterruptionHRData AS JWIHR ON (JWIHR.Source = JWI.Source AND JWIHR.SourceID = JWI.SourceID) "
    'strSQL = strSQL & "WHERE "
    
    'Generate Where Clause
    If StartDate <> EMPTY_DATE And EndDate <> EMPTY_DATE Then
        strWClause = strWClause & "CDate(JWI.StartDateTime) Between #" & StartDate & "# AND #" & EndDate & "# AND "
    End If
    
    If Left(UserName, 1) <> 0 And UserName <> vbNullString Then
        strWClause = strWClause & "JWI.UserID in (" & UserName & ") AND "
    End If
    
    If Left(JobFunction, 1) <> 0 And JobFunction <> vbNullString Then
        strWClause = strWClause & "((JWI.Source = 'JobsHR' AND JWI.EntryID IN (" & JobFunction & ")) OR JWI.Source = 'InterruptionsHR') AND "
    End If
    
    If Status <> 0 Then
        strWClause = strWClause & "((JWI.Source = 'JobsHR' AND JWI.StatusID = " & Status & ") OR JWI.Source = 'InterruptionsHR') AND "
    End If
    
    'Old Version
'    'Generate Where Clause
'    If StartDate <> EMPTY_DATE And EndDate <> EMPTY_DATE Then
'        strWClause = strWClause & "CDate(StartDateTimeSTR) Between #" & StartDate & "# AND #" & EndDate & "# AND "
'    End If
'
'    If Left(UserName, 1) <> 0 And UserName <> vbNullString Then
'        strWClause = strWClause & "UserName in (" & UserName & ") AND "
'    End If
'
'    If Left(JobFunction, 1) <> 0 And JobFunction <> vbNullString Then
'        strWClause = strWClause & "(Source = 'JobsHR' AND EntryIdentifier IN (" & JobFunction & ")) AND "
'    End If
'
'    If Status <> 0 Then
'        strWClause = strWClause & "(Source = 'JobsHR' AND Status = " & Status & ") AND "
'    End If
    
    'Chech for where clause string data and format
    If Len(strWClause) > 0 Then
        'Remove any extra " AND " at end
        If Right(strWClause, 5) = " AND " Then
            strWClause = Left(strWClause, Len(strWClause) - 5)
        End If
        
        'Add Where
        strWClause = "WHERE " & strWClause
    End If
    
    strSQL = strSQL & strWClause & "Order By JWIHR.StartDateTime"
    
    dbRS.Source = strSQL
    
    dbRS.Open
    
    AddDataToDataSheet dbRS
        
    dbRS.Close
    dbCon.Close
    
    Set dbRS = Nothing
    Set dbCon = Nothing
    
    
End Function

Private Function AddDataToDataSheet(ByRef dbRS As ADODB.Recordset)
    Dim lCnt As Long
    Dim lCnt2 As Long
    Dim strThisWorkbookName As String
    
    strThisWorkbookName = "Report Generator.xlsm"
    
    Workbooks(strThisWorkbookName).Worksheets("Data").Cells.Clear
    
    
    
    For lCnt = 1 To dbRS.Fields.Count
        Workbooks(strThisWorkbookName).Worksheets("Data").Cells(1, lCnt).Value = dbRS.Fields(lCnt - 1).Name
    Next lCnt
    
    
    For lCnt = 1 To dbRS.RecordCount
        For lCnt2 = 1 To dbRS.Fields.Count
            Workbooks(strThisWorkbookName).Worksheets("Data").Cells(lCnt + 1, lCnt2).Value = dbRS.Fields(lCnt2 - 1).Value
        Next lCnt2
        dbRS.MoveNext
    Next lCnt
    
    Workbooks(strThisWorkbookName).Worksheets("Data").Columns.AutoFit
    
    Workbooks(strThisWorkbookName).Activate
    
    Workbooks(strThisWorkbookName).Worksheets("Data").Cells(1, 1).Select
End Function

