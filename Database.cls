VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Const OPENDATABASE_E_NOTFOUND As Long = vbObjectError + 1003
Const CREATESTRUCTURE_E_ERROR As Long = vbObjectError + 1004
Const CREATEINDEX_E_ERROR As Long = vbObjectError + 1005

Dim stdR As New Registry

Sub CreateDatabase(sFile As Variant)
    Dim oCatalog As Object

    Set oCatalog = CreateObject("ADOX.Catalog")
    oCatalog.Create ("Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Jet OLEDB:Engine Type=5;Data Source=" & sFile)

    Set oCatalog = Nothing
End Sub

Function OpenDatabase(theFile As Variant) As Object ' ADODB.Connection
    Dim oConnection As Object 'ADODB.Connection
    
    On Error Resume Next
    Set oConnection = CreateObject("ADODB.Connection")
    oConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & theFile
    If Err.Number <> 0 Then
        Err.Raise OPENDATABASE_E_NOTFOUND, , Err.Description
    End If
    Err.Clear
    On Error GoTo 0
    Set OpenDatabase = oConnection
End Function

Function OpenRecordSet(ByRef oDb As Variant, sQuery As Variant) As Object ''ADODB.Connection ADODB.Recordset
    Dim oRecordSet As Object 'ADODB.Recordset
    Set oRecordSet = CreateObject("ADODB.RecordSet")
    oRecordSet.Open sQuery, oDb, adOpenKeyset, adLockOptimistic
    Set OpenRecordSet = oRecordSet
End Function

Sub CreateStructure(oDb As Variant, sSpec As Variant) 'ADODB.Connection
    On Error Resume Next
    oDb.Execute sSpec
    If Err.Number <> 0 Then
        Err.Raise CREATESTRUCTURE_E_ERROR, , Err.Description
    End If
End Sub

Sub CreateIndex(oDb As Variant, sSpec As Variant)    'ADODB.Connection
    On Error Resume Next
    oDb.Execute sSpec
    If Err.Number <> 0 Then
        Err.Raise CREATEINDEX_E_ERROR, , Err.Description
    End If
End Sub



Private Sub Class_Initialize()
    stdR.Needs "ADODB.Connection"
    If Err.Number = vbObjectError + 999 Then
        MsgBox "Class 'Database' problem: " & Err.Description
        End
    End If
    stdR.Needs "ADOX.Catalog"
    If Err.Number = vbObjectError + 999 Then
        MsgBox "Class 'Database' problem: " & Err.Description
        End
    End If
    stdR.Needs "ADODB.RecordSet"
    If Err.Number = vbObjectError + 999 Then
        MsgBox "Class 'Database' problem: " & Err.Description
        End
    End If
End Sub
