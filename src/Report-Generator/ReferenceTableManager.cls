VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReferenceTableManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ReferenceTableManager
'By Derek Pierce (8/11/2015)
'Ver. 1.0 8/11/2015 First Reusable Build
'Ver. 1.1 3/28/2015

'Used to manage and look up the data for different RefTables in the Database
'!IMPORTANT!: All sections with the "!IMPORTANT!: This section must be set up prior to use with a new project."
'line must be configured inorder for this system to work, as well as the variables below.
'The rest of the code should work without modification, as long as all required classes are included.
'Those are:
'ReferenceTableManager (This Class)
'RefTableCollectionClass
'RefTableDataClass (Requires ISearchable interface from the SearchableComboBox Code folder)
'RefTableExtraDataCollection
'ISearchable (Unless the Interface settings and code is removed from RefTableDataClass)


'These RefTableCollectionClasses represent the various tables they are named after
Private colRefTableCollections As Collection


'Table rows are stored as RefTableDataClass
'All of the collections are built out of/contain this data class
Private RefTableData As RefTableDataClass

'Enumerator used to hard define the tables used/managed by this code.
'Makes referencing tables easier in the "GetXFromY" function lookups
'If enum has no entries, then references will break and not work properly.
'!IMPORTANT!: This section must be set up prior to use with a new project.
Public Enum RefTableEnu
    [_First] = 1
        rte_PLACEHOLDERDONOTUSE = 0 'DO NOT USE THIS. CODE WILL BREAK. INTERNAL CLASS USE ONLY
        'rte_DebugRefTable = 1 'Example Entry
        rte_UsersRefTable = 1
        rte_JobFunctionRefTable = 2
        rte_StatusRefTable = 3
    [_Last] = 3 'Update to number of last entry in Enumerator
End Enum

Private RefTableBuildDate As Collection
Private LastDBCheckTime As Date
Private blnIgnoreRefUpdateChecks As Boolean

'Builds the various RefTable collections base on the named table.
Private Sub Class_Initialize()
    Set RefTableBuildDate = New Collection
    Set colRefTableCollections = New Collection
    UpdateRefTableCollections
End Sub

Private Function GetTableNameForRefTableEnu(ByVal TargetRefTable As RefTableEnu) As String
    '!IMPORTANT!: This section must be set up prior to use with a new project.
    'NOTICE: Add RefTable Names here.
    'If table(s) are not added to code it returns null string or error
    Dim strTableName As String
    
    'Example Entry
'    If TargetRefTable = rte_DebugRefTable Then 'Specify the Enumerator
'        strTableName = "DebugRefTable" 'Specify the data returned for the enumerator
'        GoTo TableFound:
'    End If


    If TargetRefTable = rte_UsersRefTable Then 'Specify the Enumerator
        strTableName = "UsersRefTable" 'Specify the data returned for the enumerator
        GoTo TableFound:
    End If
    
    If TargetRefTable = rte_JobFunctionRefTable Then 'Specify the Enumerator
        strTableName = "JobFunctionRefTable" 'Specify the data returned for the enumerator
        GoTo TableFound:
    End If
    
    If TargetRefTable = rte_StatusRefTable Then 'Specify the Enumerator
        strTableName = "StatusRefTable" 'Specify the data returned for the enumerator
        GoTo TableFound:
    End If
    
TableFound:
    GetTableNameForRefTableEnu = strTableName
End Function

Public Function GetRefCol(ByVal TargetRefTable As RefTableEnu, Optional KeyType As RefTableKeyType = rtkt_RefTypeRefName) As Collection
    'Returns the local stored version of a RefTable as a collection for use with an optional RefTableKeyType Paramater if
    'the perefered collection key should be the RefID, or the RefName (Defaults to RefName)
    Dim rtnRefTblCol As RefTableCollectionClass
    
    If blnIgnoreRefUpdateChecks = False Then
        UpdateRefTableCollections , TargetRefTable
    End If

    Set rtnRefTblCol = colRefTableCollections.Item(GetTableNameForRefTableEnu(TargetRefTable))
    
    Set GetRefCol = rtnRefTblCol.GetCollectionByRefType(KeyType)
End Function

Private Function SetRefTableCollection(ByVal TargetRefTable As RefTableEnu, ByRef NewRefTableCollectionData As RefTableCollectionClass)
    'Used by the UpdateRefTableCollections to set the specific RefTableCollectionClass data to a new
    '(if updated or forced) set of data from the database.
    Dim tmpRefTableCollection As RefTableCollectionClass

    If Not InCollection(colRefTableCollections, GetTableNameForRefTableEnu(TargetRefTable)) Then
        colRefTableCollections.Add NewRefTableCollectionData, GetTableNameForRefTableEnu(TargetRefTable)
    Else:
        colRefTableCollections.Remove (GetTableNameForRefTableEnu(TargetRefTable))
        colRefTableCollections.Add NewRefTableCollectionData, GetTableNameForRefTableEnu(TargetRefTable)
    End If
    
End Function

Private Function UpdateRefTableCollections(Optional ByVal ForceUpdate As Boolean = False, Optional ByVal TargetRefTable As RefTableEnu = rte_PLACEHOLDERDONOTUSE)
    Dim tmpLastTime As Date
    Dim RefEnuNum As Long

    
'    If DateDiff("s", LastDBCheckTime, DateTime.Now) > 5 Then
'        tmpLastTime = LastRefTableUpdate
'        LastDBCheckTime = DateTime.Now
'    End If
    
    If TargetRefTable = rte_PLACEHOLDERDONOTUSE Then
        For RefEnuNum = RefTableEnu.[_First] To RefTableEnu.[_Last]
            If Not InCollection(RefTableBuildDate, GetTableNameForRefTableEnu(RefEnuNum)) Then
                tmpLastTime = LastRefTableUpdate(RefEnuNum)
                RefTableBuildDate.Add tmpLastTime, GetTableNameForRefTableEnu(RefEnuNum)
                SetRefTableCollection RefEnuNum, BuildRefTableCollectionClass(RefEnuNum)
            ElseIf RefTableBuildDate.Item(GetTableNameForRefTableEnu(RefEnuNum)) < LastRefTableUpdate(RefEnuNum) Or ForceUpdate = True Then
                tmpLastTime = LastRefTableUpdate(RefEnuNum)
                SetRefTableCollection RefEnuNum, BuildRefTableCollectionClass(RefEnuNum)
                RefTableBuildDate.Remove (GetTableNameForRefTableEnu(RefEnuNum))
                RefTableBuildDate.Add tmpLastTime, GetTableNameForRefTableEnu(RefEnuNum)
            End If
        Next RefEnuNum
    Else:
        If ForceUpdate = True Then
            tmpLastTime = LastRefTableUpdate(TargetRefTable)
            SetRefTableCollection TargetRefTable, BuildRefTableCollectionClass(TargetRefTable)
            RefTableBuildDate.Remove (GetTableNameForRefTableEnu(TargetRefTable))
            RefTableBuildDate.Add tmpLastTime, GetTableNameForRefTableEnu(TargetRefTable)
        Else:
            If RefTableBuildDate.Item(GetTableNameForRefTableEnu(TargetRefTable)) < LastRefTableUpdate(TargetRefTable) Then
                tmpLastTime = LastRefTableUpdate(TargetRefTable)
                SetRefTableCollection TargetRefTable, BuildRefTableCollectionClass(TargetRefTable)
                RefTableBuildDate.Remove (GetTableNameForRefTableEnu(TargetRefTable))
                RefTableBuildDate.Add tmpLastTime, GetTableNameForRefTableEnu(TargetRefTable)
            End If
        End If
    End If

    
End Function

'Using a string value and a reference table to look up, returns the ID for that string
Public Function GetIDFromString(ByVal TargetTypeRefTable As RefTableEnu, strTarget As String) As Long
    'Key lookups are CASE SENSITIVE!!
    GetIDFromString = GetRelevantDataByStr(TargetTypeRefTable, strTarget).RefTypeID

End Function

'Using a long(variant) value and a reference table to look up, returns the String value for that ID
Public Function GetStringFromID(ByVal TargetTypeRefTable As RefTableEnu, lngTarget As Variant) As String

    GetStringFromID = GetRelevantDataByID(TargetTypeRefTable, lngTarget).RefTypeName

End Function

'Gets any extra data that may be present for the Reftable. (Forward comments etc..)
Public Function GetExtraDataFromID(ByVal TargetTypeRefTable As RefTableEnu, lngTarget As Variant) As Collection
    
    Set GetExtraDataFromID = GetRelevantDataByID(TargetTypeRefTable, lngTarget).RefTypeExtraData
    
End Function

'Returns a RefTableDataClass for use with the GetXDatafromY functions based off a string search.
Private Function GetRelevantDataByStr(ByVal TargetTypeRefTable As RefTableEnu, varTarget As String) As RefTableDataClass
    Dim tmpRefTableData As RefTableDataClass
    
    If varTarget <> "No Data" Then
        Set tmpRefTableData = GetRefCol(TargetTypeRefTable, rtkt_RefTypeRefName).Item(CStr(varTarget))
    End If
    
    If tmpRefTableData Is Nothing Then
        Set tmpRefTableData = New RefTableDataClass
        tmpRefTableData.RefTypeID = 0
        tmpRefTableData.RefTypeName = ""
        tmpRefTableData.RefTypeExtraData = Nothing
    End If
    
    Set GetRelevantDataByStr = tmpRefTableData

End Function

'Returns a RefTableDataClass for use with the GetXDataFromY functions based off a long(variant) search.
Private Function GetRelevantDataByID(ByVal TargetTypeRefTable As RefTableEnu, varTarget As Variant) As RefTableDataClass
    Dim tmpRefTableData As RefTableDataClass
    
    If InCollection(GetRefCol(TargetTypeRefTable, rtkt_RefTypeRefID), CStr(varTarget)) Then
        Set tmpRefTableData = GetRefCol(TargetTypeRefTable, rtkt_RefTypeRefID).Item(CStr(varTarget))
    End If
    
'Old Blind to 0 Key code
'    If varTarget <> 0 Then
'        Set tmpRefTableData = GetRefCol(TargetTypeRefTable, rtkt_RefTypeRefID).Item(CStr(varTarget))
'    End If
    
    If tmpRefTableData Is Nothing Then
        Set tmpRefTableData = New RefTableDataClass
        tmpRefTableData.RefTypeID = 0
        tmpRefTableData.RefTypeName = ""
        tmpRefTableData.RefTypeExtraData = Nothing
    End If
    
    Set GetRelevantDataByID = tmpRefTableData

End Function

'Class function used to create the various reference collections from the database data.
'Select TargetRefTable using the RefTableEnu enumerator, (If not set up it won't work.) get collection representing that tables data.
'Reqired Fields are RefID and RefName, and are used by the system to get the basic information required for this system to work.
'(A Query/View where the Primary Unique Identifier is aliased as RefID and a easy readable name as RefName can be used instead)
'NOTE: Above line is incorrect. ApplicationDataStore has no way of knowing about linkage of query to table and updates so
'       updates to primary table will not cause a change in ApplicationDataStore for the query entry.
'       Only use a Query or View if you know what you are doing. TODO: Fix this issue.
'Any fields not named RefID and RefName are added to a RefTableExtraDataCollection object inside of the RefTableDataClass.
'The Field.Name is used as the Key (Cast as String) for the collection, and the Field.Value is used as the Value.
'Not tested where the Field.Value is an object.
Private Function BuildRefTableCollectionClass(ByVal TargetRefTable As RefTableEnu, Optional KeyType As RefTableKeyType = rtkt_RefTypeRefName) As RefTableCollectionClass
    'Generic function to build a collection based of a 2 field table. (IE: RefTable Type Tables)
    Dim strSQL As String
    Dim rtm_dbConnection As New ADODB.Connection
    Dim rtm_rsRecordset As New ADODB.Recordset
    Dim tmpCol As New RefTableCollectionClass
    Dim lCount As Long
    Dim extDataHold As Variant
    Dim extDataHoldCol As Collection
    
    strSQL = "SELECT * FROM " & GetTableNameForRefTableEnu(TargetRefTable)
    

'    Set rtm_dbConnection = GetDBCon
'    rtm_dbConnection.Open (DB_LOC & DB_NAME)
'
'    Set rtm_rsRecordset = GetDBRS(rtm_dbConnection)
'
'    rtm_rsRecordset.LockType = adLockReadOnly
'    rtm_rsRecordset.CursorType = adOpenStatic
'
'    rtm_rsRecordset.Source = strSQL
'
'    rtm_rsRecordset.Open
    
    CombinedConRecordSetPrep rtm_rsRecordset, strSQL
    
    Do Until rtm_rsRecordset.EOF
        Set RefTableData = New RefTableDataClass

        Set extDataHoldCol = New Collection
        For lCount = 0 To rtm_rsRecordset.Fields.Count - 1
            
            Select Case rtm_rsRecordset.Fields(lCount).Name
                Case "RefID" 'Should be a Unique identifier or PK Duplicated values will cause failure when adding to collection.
                    'TODO: Maybe create a way to handle if this field in DB is not Unique.
                    RefTableData.RefTypeID = rtm_rsRecordset.Fields(lCount).Value
                Case "RefName"
                    If Not IsNull(rtm_rsRecordset.Fields(lCount).Value) Then
                        RefTableData.RefTypeName = rtm_rsRecordset.Fields(lCount).Value
                    Else:
                        RefTableData.RefTypeName = ""
                    End If
                Case Else:
                    'Field name is the Key, Field value is the Item
                    If Not IsNull(rtm_rsRecordset.Fields(lCount)) Then
                        RefTableData.RefTypeExtraData.Add rtm_rsRecordset.Fields(lCount).Value, CStr(rtm_rsRecordset.Fields(lCount).Name)
                    Else
                        RefTableData.RefTypeExtraData.Add "", CStr(rtm_rsRecordset.Fields(lCount).Name)
                    End If
            End Select
            
        Next lCount
        
        tmpCol.Add RefTableData
        
        rtm_rsRecordset.MoveNext
    Loop
    
    rtm_rsRecordset.Close
'    rtm_dbConnection.Close
    Set rtm_rsRecordset = Nothing
    Set rtm_dbConnection = Nothing
    
    Set BuildRefTableCollectionClass = tmpCol
    
    
End Function

Public Function CheckIfTextExists(ByVal TargetRefTable As RefTableEnu, ByVal TargetText As String) As Boolean
    Dim colToCheck As Collection
    
    Set colToCheck = GetRefCol(TargetRefTable)
    
    CheckIfTextExists = InCollection(colToCheck, CStr(TargetText))
            
End Function

Public Function CheckIfIDExists(ByVal TargetRefTable As RefTableEnu, ByVal TargetID As String) As Boolean
    Dim colToCheck As Collection
    
    Set colToCheck = GetRefCol(TargetRefTable, rtkt_RefTypeRefID)
    
    CheckIfIDExists = InCollection(colToCheck, CStr(TargetID))
    
End Function

Public Function AddNewRefData(ByVal TargetRefTable As RefTableEnu, ByRef NewRefTableData As RefTableDataClass)
    'RefTables with Extra Data (Field count over 2)
    Dim strSQL As String
    Dim rtm_dbConnection As New ADODB.Connection
    Dim rtm_rsRecordset As New ADODB.Recordset
    Dim tmpCol As New Collection
    Dim lCount As Long
    Dim extDataFieldName As Variant
    Dim extDataHold As Variant
    Dim tmpRefTableColData As RefTableCollectionClass
    
    If LastRefTableUpdate(TargetRefTable) > CDate(RefTableBuildDate.Item(TargetRefTable)) Then
        UpdateRefTableCollections , TargetRefTable
    End If
    'TODO: Is this still a proper check?
    If Not CheckIfTextExists(TargetRefTable, CStr(NewRefTableData.RefTypeName)) Then
        
        strSQL = "SELECT * FROM " & GetTableNameForRefTableEnu(TargetRefTable)
        
    
'        Set rtm_dbConnection = GetDBCon
'        rtm_dbConnection.Open (DB_LOC & DB_NAME)
'
'        Set rtm_rsRecordset = GetDBRS(rtm_dbConnection)
'
'        rtm_rsRecordset.LockType = adLockOptimistic
'        rtm_rsRecordset.CursorType = adOpenKeyset
'
'        rtm_rsRecordset.Source = strSQL
'
'        rtm_rsRecordset.Open
        
        CombinedConRecordSetPrep rtm_rsRecordset, strSQL, adOpenKeyset, adLockOptimistic
        
        rtm_rsRecordset.AddNew
        If NewRefTableData.RefTypeID <> -1 Then
            rtm_rsRecordset.Fields("RefID").Value = NewRefTableData.RefTypeID
        End If
        rtm_rsRecordset.Fields("RefName").Value = NewRefTableData.RefTypeName
        
        For Each extDataFieldName In NewRefTableData.RefTypeExtraData.GetCollectionByExtraDataType(rtedt_FieldNames)
            'TODO: May need to add error handling code for Field changes mid process... though this should be midigated
            'by using a system that creates an empty ReftableDataClass to put the new data into. (GetConfiguredRefTableDataClass)
            rtm_rsRecordset.Fields(extDataFieldName).Value = NewRefTableData.RefTypeExtraData.Item(extDataFieldName, rtedt_FieldValue)
        Next extDataFieldName
        
        
        rtm_rsRecordset.Update
        
        
        LastRefTableUpdate TargetRefTable, True
        
        rtm_rsRecordset.Close
'        rtm_dbConnection.Close
        
        Set rtm_rsRecordset = Nothing
        Set rtm_dbConnection = Nothing
        
'        If blnIgnoreRefUpdateChecks = False Then
        UpdateRefTableCollections True, TargetRefTable
'        Else:
'            Set tmpRefTableColData = colRefTableCollections.Item(GetTableNameForRefTableEnu(TargetRefTable))
'            tmpRefTableColData.Add NewRefTableData
'        End If
        
    End If
    
    
End Function

Public Function UpdateExistingRefData(ByVal TargetRefTable As RefTableEnu, ByRef NewRefTableData As RefTableDataClass)
    Dim strSQL As String
    Dim rtm_dbConnection As New ADODB.Connection
    Dim rtm_rsRecordset As New ADODB.Recordset
    Dim tmpCol As New Collection
    Dim lCount As Long
    Dim extDataFieldName As Variant
    Dim extDataHold As Variant
    Dim tmpRefTableColData As RefTableCollectionClass
    
    If LastRefTableUpdate(TargetRefTable) > CDate(RefTableBuildDate.Item(TargetRefTable)) Then
        UpdateRefTableCollections , TargetRefTable
    End If
    
    'issues if refID is a string
    strSQL = "SELECT * FROM " & GetTableNameForRefTableEnu(TargetRefTable) & " Where RefID = " & NewRefTableData.RefTypeID

'    Set rtm_dbConnection = GetDBCon
'    rtm_dbConnection.Open (DB_LOC & DB_NAME)
'
'    Set rtm_rsRecordset = GetDBRS(rtm_dbConnection)
'
'    rtm_rsRecordset.LockType = adLockOptimistic
'    rtm_rsRecordset.CursorType = adOpenKeyset
'
'    rtm_rsRecordset.Source = strSQL
'
'    rtm_rsRecordset.Open

    CombinedConRecordSetPrep rtm_rsRecordset, strSQL, adOpenKeyset, adLockOptimistic
    
    If rtm_rsRecordset.RecordCount = 0 Then
        AddNewRefData TargetRefTable, NewRefTableData
    ElseIf rtm_rsRecordset.RecordCount = 1 Then
    
        If rtm_rsRecordset.Fields("RefID").Value <> NewRefTableData.RefTypeID Then
            rtm_rsRecordset.Fields("RefID").Value = NewRefTableData.RefTypeID
        End If
        
        
        If rtm_rsRecordset.Fields("RefName").Value <> NewRefTableData.RefTypeName Then
            rtm_rsRecordset.Fields("RefName").Value = NewRefTableData.RefTypeName
        End If
        
        For Each extDataFieldName In NewRefTableData.RefTypeExtraData.GetCollectionByExtraDataType(rtedt_FieldNames)
            'TODO: May need to add error handling code for Field changes mid process... though this should be midigated
            'by using a system that creates an empty ReftableDataClass to put the new data into. (GetConfiguredRefTableDataClass)
            If rtm_rsRecordset.Fields(extDataFieldName).Value <> NewRefTableData.RefTypeExtraData.Item(extDataFieldName, rtedt_FieldValue) Then
                rtm_rsRecordset.Fields(extDataFieldName).Value = NewRefTableData.RefTypeExtraData.Item(extDataFieldName, rtedt_FieldValue)
            End If
        Next extDataFieldName
        
        rtm_rsRecordset.Update
        
        LastRefTableUpdate TargetRefTable, True
    ElseIf rtm_rsRecordset.RecordCount >= 2 Then
        'Shouldn't happen but could depending on config or improper use
    End If
    
    rtm_rsRecordset.Close
'    rtm_dbConnection.Close
    
    Set rtm_rsRecordset = Nothing
    Set rtm_dbConnection = Nothing
    
'   If blnIgnoreRefUpdateChecks = False Then
        UpdateRefTableCollections True, TargetRefTable
'   Else:
'       Set tmpRefTableColData = colRefTableCollections.Item(GetTableNameForRefTableEnu(TargetRefTable))
'       tmpRefTableColData.Add NewRefTableData
'   End If
    
End Function

Private Function LastRefTableUpdate(ByVal TargetRefTable As RefTableEnu, Optional ByVal UpdateDate As Boolean = False) As Date
    'Returns the Date in the ApplicationDataStore pertaining to the TargetRefTable supplied, or if UpdateDate = True
    'updates the date to the current date and time and returns that. Also returns current date if ApplicationDataStore's
    'ElementDataStore field is Null.
    '
    'Handles creation of new entries in the ApplicationDataStore table for initial load of reftables, or as new ones
    'get created in the DB and added to the code at a later date. (This does not create new RefTables. Only adds there
    'ApplicationElement and ElementDataStore data to the fields in ApplicationDataStore.)
    
    Dim dbCon As ADODB.Connection
    Dim dbRS As ADODB.Recordset
    Dim rtnDateTime As Date
    Dim strSQL As String
    
    strSQL = "Select * from ApplicationDataStore where ApplicationElement = '" & GetTableNameForRefTableEnu(TargetRefTable) & "Update'"
    
'    Set DBCon = GetDBCon
'    DBCon.Open DB_LOC & DB_NAME
'
'    Set dbRS = GetDBRS(DBCon)
'
'    dbRS.Source = strSQL
'
'    dbRS.Open

    CombinedConRecordSetPrep dbRS, strSQL, adOpenKeyset, adLockOptimistic
    
    'Adds new Reftable entry to ApplicationDataStore if not found
    If dbRS.RecordCount = 0 Then
        dbRS.AddNew
        dbRS.Fields("ApplicationElement").Value = GetTableNameForRefTableEnu(TargetRefTable) & "Update"
        dbRS.Update
    End If
    
    'Checks for date and/or updates date
    If UpdateDate = False Then
        If Not IsNull(dbRS.Fields("ElementDataStore").Value) Then
            rtnDateTime = dbRS.Fields("ElementDataStore").Value
        Else:
            rtnDateTime = EMPTY_DATE
        End If
    Else:
        dbRS.Fields("ElementDataStore").Value = DateTime.Now
        dbRS.Update
        rtnDateTime = dbRS.Fields("ElementDataStore").Value
    End If
        
    dbRS.Close
    'DBCon.Close
    
    Set dbRS = Nothing
    Set dbCon = Nothing
    
    LastRefTableUpdate = rtnDateTime

End Function

Public Function GetConfiguredRefTableDataClass(ByVal TargetRefTable As RefTableEnu) As RefTableDataClass
    'Returns an empty RefTableDataClass that has been configured with the fields from the TargetRefTable
    'for use with adding new RefTableData
    Dim dbCon As ADODB.Connection
    Dim dbRS As ADODB.Recordset
    Dim lCount As Long
    Dim rtnRefTabDataCls As RefTableDataClass
    Dim strSQL As String
    
    strSQL = "SELECT * FROM " & GetTableNameForRefTableEnu(TargetRefTable)
    
'    Set DBCon = GetDBCon
'    DBCon.Open DB_LOC & DB_NAME
'
'    Set dbRS = GetDBRS(DBCon)
'
'    dbRS.Source = strSQL
'
'    dbRS.Open

    CombinedConRecordSetPrep dbRS, strSQL, adOpenKeyset, adLockOptimistic
    
    Set rtnRefTabDataCls = New RefTableDataClass
    
    For lCount = 0 To dbRS.Fields.Count - 1
        If dbRS.Fields(lCount).Name <> "RefID" And dbRS.Fields(lCount).Name <> "RefName" Then
            rtnRefTabDataCls.RefTypeExtraData.Add "", CStr(dbRS.Fields(lCount).Name)
        End If
    Next lCount
    
    dbRS.Close
'    DBCon.Close
    
    Set dbRS = Nothing
    Set dbCon = Nothing
    
    Set GetConfiguredRefTableDataClass = rtnRefTabDataCls
    
End Function

Public Function IgnoreRefUpdateChecks(ByVal Ignore As Boolean)
    blnIgnoreRefUpdateChecks = Ignore
End Function

Public Function IsIgnoringUpdates() As Boolean
    IsIgnoringUpdates = blnIgnoreRefUpdateChecks
End Function
