VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Users"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim stdA As New Arrays
Dim stdC As New Computers
Dim stdS As New Strings
Dim stdR As New Registry

Function GetOUofUser(sDomain As String, sUsername As String) As String
    Dim sResult As String
    Dim oConnection As Object
    Dim oCommand As Object
    Dim oRecordSet As Object
    
    sResult = vbNullString

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
        "SELECT distinguishedName FROM 'LDAP://" & sDomain & "' WHERE objectCategory='user' " & _
            "AND sAMAccountName='" & sUsername & "'"
    Set oRecordSet = oCommand.Execute

    oRecordSet.MoveFirst
    Do Until oRecordSet.EOF
        sResult = oRecordSet.Fields("distinguishedName").Value
        oRecordSet.MoveNext
    Loop

    GetOUofUser = sResult
End Function

'Function GetLoginProfiles(sComputer) As Variant()
'    Dim cItems As SWbemObjectSet
'    Dim oItem As SWbemObject
'    Dim aResult() As Variant
'    Dim i As Integer
'
'    On Error Resume Next
'    Set cItems = GetObject("winmgmts:\\" & sComputer & "\root\cimv2").InstancesOf("Win32_NetworkLoginProfile")
'    If Err.Number = 0 Then
'        ReDim aResult(cItems.Count)
'        For i = 0 To cItems.Count
'            aResult(i) = cItems.Item(i).Name
'        Next
'    End If
'    On Error GoTo 0
'
'    GetLoginProfiles = aResult
'End Function

'Function GetMac(sDevice) As String
'    Dim cItems As SWbemObjectSet
'    Dim oItem As SWbemObject
'    Dim sResult As String
'    Dim i As Integer
'    sResult = vbNullString
'
'    On Error Resume Next
'    Set cItems = GetObject("winmgmts:\\" & sDevice & "\root\cimv2").InstancesOf("Win32_NetworkAdapterConfiguration")
'    If Err.Number = 0 Then
'
'        For i = 0 To cItems.Count
'            sResult = sResult & "[" & cItems.Item(i).Caption & "]" & cItems.Item(i).MACAddress & vbTab
'        Next
'
'        sResult = Mid(sResult, 1, Len(sResult) - 1)
'    End If
'    On Error GoTo 0
'
'    GetMac = sResult
'End Function

'http://www.windowsitpro.com/WindowsScripting/Article/ArticleID/4734/4734.html
'----------------------------------------------------------------
' Function:     GetComputerUsers
' Description:  This function will return the names of the users
'                 for a specified computer or domain
' Parameters:   ByVal strComputerName - Name of an NT Computer or
'                 domain to retrieve the list of Users from.
' Returns:      Variant array of user account for the specified
'           computer
'----------------------------------------------------------------
'Function GetComputerUsers(ByVal strComputerName As String) As Variant()
'   Dim objIADsContainer      As ActiveDs.IADsContainer
'   Dim objIADsUser           As ActiveDs.IADsUser
'   Dim vReturn()               As Variant
'   ReDim vReturn(0)
'
'   ' Connect to the computer or domain
'   Set objIADsContainer = GetObject("WinNT://" & strComputerName)
'
'   ' apply a filter to retrieve only objects of class User
'   objIADsContainer.Filter = Array("User")
'
'   ' iterate all user objects within the container
'   For Each objIADsUser In objIADsContainer
'      If Trim(vReturn(0)) <> vbNullString Then
'         ReDim Preserve vReturn(UBound(vReturn) + 1)
'      End If
'      vReturn(UBound(vReturn)) = objIADsUser.Name
'      '~ WScript.Echo objIADsUser.Fullname
'   Next
'
'   GetComputerUsers = vReturn
'End Function
'
'Function GetComputerUsers2(ByVal strComputerName) As Variant()
'   Dim objIADsContainer      As ActiveDs.IADsContainer
'   Dim objIADsUser           As ActiveDs.IADsUser
'   Dim vReturn()               As Variant
'   ReDim vReturn(0)
'
'   ' Connect to the computer or domain
'   Set objIADsContainer = GetObject("WinNT://" & strComputerName)
'
'   ' apply a filter to retrieve only objects of class User
'   objIADsContainer.Filter = Array("User")
'
'   ' iterate all user objects within the container
'   For Each objIADsUser In objIADsContainer
'      If Trim(vReturn(0)) <> vbNullString Then
'         ReDim Preserve vReturn(UBound(vReturn) + 1)
'      End If
'      vReturn(UBound(vReturn)) = objIADsUser.Name & "^" & objIADsUser.FullName
'      '~ WScript.Echo objIADsUser.Fullname
'   Next
'
'   GetComputerUsers2 = vReturn
'End Function
'
'Function GetDomainUser(sComputer As String) As String
'    Dim oWMI As Object
'    Dim cProc As SWbemObjectSet
'    Dim oProcess As SWbemObject
'    Dim sUser As String
'    Dim sDomain As String
'    Dim sResult As String
'
'    sResult = "<not switched on>"
'    If stdC.ResolveIP(sComputer) <> Empty Then
'
'        On Error Resume Next
'        Set oWMI = GetObject("winmgmts:\\" & sComputer & "\root\CIMV2")
'        If Err.Number = 0 Then
'            'Get currently logged on user and update CurrentUser element.
'            'This bit of code is "cheating" slightly in that all it's doing is getting the owner of the explorer.exe process.  Seeing as explorer.exe
'            'only runs if a user is logged on it is a reliable (but sneaky) way of getting the logged-on user.  It does, however, mean that even if a
'            'user is logged on locally or the account is a generic account, it will retrieve the correct values (WMI cannot do this).
'
'            Set cProc = oWMI.ExecQuery("Select Name from Win32_Process Where Name='explorer.exe' and  SessionID=0")
'
'            If cProc.Count > 0 Then
'                For Each oProcess In cProc
'                    oProcess.GetOwner sUser, sDomain
'                    sResult = sDomain & "\" & sUser
'                Next
'             Else
'                sResult = "<not logged on>"
'            End If
'            Err.Clear
'            On Error GoTo 0
'        End If
'    End If
'    GetDomainUser = sResult
'End Function
'
'Function GetNameAndDescription(spDomain As String, sUsername As String) As String
'    Dim sDomain As String
'    Dim oIADSUser As ActiveDs.IADsUser
'    Dim sResult As String
'
'    sResult = vbNullString
'
'    If IsEmpty(spDomain) Then
'        sDomain = CreateObject("WScript.Network").UserDomain
'    Else
'        sDomain = spDomain
'    End If
'
'    If Len(sUsername) > 0 Then
'        On Error Resume Next
'        Set oIADSUser = GetObject(stdS.Subst("WinNT://[1]/[2],user", Array(sDomain, sUsername)))
'        If Err.Number = 0 Then
'            sResult = oIADSUser.FullName & vbTab & oIADSUser.Description
'        End If
'        On Error GoTo 0
'
'    End If
'
'    GetNameAndDescription = sResult
'End Function
'
Private Sub Class_Initialize()
    stdR.Needs "ADODB.Connection"
    If Err.Number = vbObjectError + 999 Then
        MsgBox "Class 'Users' problem: " & Err.Description
        End
    End If
    stdR.Needs "ADODB.Command"
    If Err.Number = vbObjectError + 999 Then
        MsgBox "Class 'Users' problem: " & Err.Description
        End
    End If

End Sub
