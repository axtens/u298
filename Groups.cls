VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Groups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim REG As New Registry

'http://www.windowsitpro.com/WindowsScripting/Article/ArticleID/4734/4734.html
'---------------------------------------------------------------
' Function:     GetComputerGroups
' Description:  This function will return the names of
'                 the NT Groups for a specified computer.
' Parameters:   ByVal strComputerName - Name of an NT
'                 Computer or Domain to retrieve the list of
'           Groups from.
' Returns:      Variant array of NT Groups for the specified
'           computer
'----------------------------------------------------------------
'Public Function GetComputerGroups(ByVal strComputerName As String) As Variant()
'   Dim objIADsContainer      As ActiveDs.IADsContainer
'   Dim objIADsGroup          As ActiveDs.IADsGroup
'   Dim vReturn()               As Variant
'   ReDim vReturn(0)
'
'   ' connect to the computer or Domain
'   Set objIADsContainer = GetObject("WinNT://" & strComputerName)
'
'   ' Set the filter on the container to retrieve only
'   ' objects of Class Group
'   objIADsContainer.Filter = Array("Group")
'
'   ' iterate through each group object in the container
'   For Each objIADsGroup In objIADsContainer
'      If Trim(vReturn(0)) <> vbNullString Then
'         ReDim Preserve vReturn(UBound(vReturn) + 1)
'      End If
'      vReturn(UBound(vReturn)) = objIADsGroup.Name
'   Next
'
'   GetComputerGroups = vReturn
'End Function

Function GetGroups() As Variant()
    Dim oRecordSet As Object
    Dim oCommand As Object
    Dim oConnection As Object
    Dim i As Integer
    Dim aList() As Variant
    ReDim aList(10000)
    On Error Resume Next

    Const ADS_SCOPE_SUBTREE = 2

    Set oConnection = CreateObject("ADODB.Connection")
    Set oCommand = CreateObject("ADODB.Command")
    oConnection.Provider = "ADsDSOObject"
    oConnection.Open "Active Directory Provider"
    Set oCommand.ActiveConnection = oConnection

    oCommand.Properties("Page Size") = 1000
    oCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE

    oCommand.CommandText = _
        "SELECT cn, name, ADsPath, description FROM 'LDAP://" & GetDefaultNamingContext() & "' WHERE " _
            & "objectCategory='group'"
    Set oRecordSet = oCommand.Execute
    oRecordSet.MoveFirst
    i = 0
    Do Until oRecordSet.EOF
        aList(i) = oRecordSet.Fields("cn").Value & vbTab & oRecordSet.Fields("ADsPath").Value
        oRecordSet.MoveNext
        i = i + 1
    Loop

    ReDim Preserve aList(i - 1)
    GetGroups = aList
End Function

Function GetDefaultNamingContext() As String
    GetDefaultNamingContext = GetObject("LDAP://rootDSE").Get("defaultNamingContext")
End Function

Private Sub Class_Initialize()
    On Error Resume Next
    REG.Needs "ADODB.Connection"
    If Err.Number = vbObjectError + 999 Then
        MsgBox "Class 'Groups' problem: " & Err.Description
        End
    End If
    
    REG.Needs "ADODB.Command"
    If Err.Number = vbObjectError + 999 Then
        MsgBox "Class 'Groups' problem: " & Err.Description
        End
    End If
End Sub
