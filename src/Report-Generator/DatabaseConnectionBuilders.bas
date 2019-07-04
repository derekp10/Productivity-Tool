Attribute VB_Name = "DatabaseConnectionBuilders"
Option Explicit

'Public Const DB_LOC As String = "\\hemingway\h_pds$\trans\Productivity Tool\"
'Public Const DB_NAME As String = "Discharge Tracking Database.accdb"
'Public Const DB_CON As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source ="
Public Const DB_PROV As String = "Microsoft.ACE.OLEDB.12.0"

Public PUBDBCon As ADODB.Connection

Public Function GetDBCon() As ADODB.Connection
    Dim con As ADODB.Connection
    
    Set con = New ADODB.Connection
    con.Provider = DB_PROV
    
    If Not DEBUG_MODE Then
        con.Properties("Jet OLEDB:Database Password") = "PT@HMC"
    End If
    
    'con.Open (DB_LOC)
    
    Set GetDBCon = con
End Function

Public Function GetExCon() As ADODB.Connection
    Dim con As ADODB.Connection
    
    Set con = New ADODB.Connection
    con.Provider = DB_PROV
    con.Properties("Extended Properties").Value = "Excel 8.0;HDR=Yes;"
    'con.Properties("Extended Properties").Value = "text;HDR=Yes;FMT=Delimited"
    'con.Open (DB_LOC)
    
    Set GetExCon = con
End Function

Public Function GetTextCon() As ADODB.Connection
    Dim con As ADODB.Connection
    
    Set con = New ADODB.Connection
    con.Provider = DB_PROV
    con.Properties("Extended Properties").Value = "text;HDR=Yes;FMT=Delimited(,)"
    'con.Open (DB_LOC)
    
    Set GetTextCon = con
End Function

Public Function GetDBRS(ByRef DBConnection As ADODB.Connection) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = DBConnection
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    
    Set GetDBRS = rs
End Function

Public Function GetExRS(ByRef DBConnection As ADODB.Connection) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = DBConnection
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenStatic
    
    Set GetExRS = rs
End Function

Public Function GetTextRS(ByRef DBConnection As ADODB.Connection) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = DBConnection
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenStatic
    
    Set GetTextRS = rs
End Function

Public Function CombinedConRecordSetPrep(ByRef rs As ADODB.Recordset, ByVal strSQL As String, Optional enuCursor As ADODB.CursorTypeEnum, Optional enuLock As ADODB.LockTypeEnum, Optional bolClientSide As Boolean) As ADODB.Recordset
    'Used to open the supplied recordset object using only one instance of a connection object.
    'This is to fix the Too Many Users error that occured when to many connections were opened or if they didn't
    'close properly.
    
    If PUBDBCon Is Nothing Then
        Set PUBDBCon = GetDBCon
        PUBDBCon.Open DB_LOC & DB_NAME
    End If
    
    If PUBDBCon.State = adStateClosed Then
        Set PUBDBCon = GetDBCon
        PUBDBCon.Open DB_LOC & DB_NAME
    End If
    
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = PUBDBCon
    If bolClientSide Then
        rs.CursorLocation = adUseClient
    Else:
        rs.CursorLocation = adUseServer
    End If
    
    rs.CursorType = IIf((enuCursor = 0), adOpenStatic, enuCursor)
    rs.LockType = IIf((enuLock = 0), adLockReadOnly, enuLock)
    rs.Source = strSQL
    rs.Open
    
End Function

Public Function ClosePublicADODBConnection()
    'This function is to be used with the backupDatabase function in Utils to close any active connection
    'so the DB can be compacted.
    'DO NOT USE THIS IN ANY OF THE IMPORT CODE OR IN THE MIDDLE OF ANY DATABASE ACTION
    'IT WILL CAUSE THE PROCESS USING THE CONNECTION TO FAIL.
    If PUBDBCon.State = adStateOpen Then
        PUBDBCon.Close
    End If
End Function

