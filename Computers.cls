VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Computers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim stdS As New Strings
Dim stdy As New System
Dim stdR As New Registry

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'Function GetRemoteEnvironment(sComputer) As Variant()
'    Dim aResult() As Variant
'    Dim cEnvs As SWbemObjectSet
'    Dim oVar As SWbemObject
'    Dim i As Integer
'
'    Set cEnvs = GetObject("winmgmts:\\" & sComputer & "\root\cimv2").InstancesOf("Win32_Environment")
'    ReDim aResult(cEnvs.Count)
'    For i = 0 To cEnvs.Count
'        aResult(i) = oVar.SystemVariable & vbTab & oVar.Name & vbTab & oVar.VariableValue & vbTab & oVar.Description & vbTab & oVar.UserName
'    Next
'    GetRemoteEnvironment = aResult
'End Function

Function GetOnlineComputers(sDomain As String) As Variant()
    Dim sTemp As String
    Dim aTemp() As String
    Dim aResult() As Variant
    Dim i As Integer
    sTemp = GetNetView(sDomain)
    aTemp = Split(sTemp, vbNewLine)
    ReDim aResult(UBound(aTemp) - 6)

    For i = 3 To UBound(aTemp) - 3
        aResult(i - 3) = stdS.NthField(aTemp(i), " ", 1)
    Next

    GetOnlineComputers = aResult
End Function

Function GetNetView(sDomain As String) As String
    GetNetView = stdy.CaptureDOS("%COMSPEC% /C NET VIEW /DOMAIN:" & sDomain, True)
End Function

Function GetComputers(sDomain As String) As Variant()
    Dim oController As Object
    Dim nCount As Integer
    Dim oComputer As Object
    Dim aResult() As Variant
    ReDim aResult(5000)

    Set oController = GetObject("WinNT://" & sDomain)
    oController.Filter = Array("Computer")

    nCount = 0
    For Each oComputer In oController
        aResult(nCount) = oComputer.Name
        nCount = nCount + 1
    Next
    ReDim Preserve aResult(nCount - 1)
    GetComputers = aResult
End Function

Function GetNBT(sComputer As String) As String
    Dim sText As String
    Dim aText() As String
    Dim i As Integer
    Dim sTemp As String
    Dim n As Integer
    Dim sResult As String

    sResult = ""

    sText = stdy.CaptureDOS("%COMSPEC% /c nbtstat -a " & sComputer, True)
    aText = Split(sText, vbNewLine)

    If stdS.Contains(aText(4), "Host not found.") = False Then
        n = 0
        For i = 0 To UBound(aText)
            sTemp = Replace(RTrim(LTrim(aText(i))), vbCr, "")
            aText(i) = sTemp
        Next

        Dim sIP As String
        Dim sLine As String
        Dim sMac As String

        sIP = stdS.Between(aText(2), "[", "]")

        n = 7
        Do While True
            If Len(aText(n)) = 0 Then
                Exit Do
            End If
            sLine = aText(n)
            n = n + 1
        Loop

        sLine = Replace(sLine, "  ", " ")
        sLine = Replace(sLine, "  ", " ")
        sLine = Replace(sLine, "  ", " ")

        sMac = Split(aText(n + 1), " = ")(1)
        sResult = sIP & vbTab & sMac & vbTab & Split(sLine, " ")(0)
    End If
    GetNBT = sResult
End Function

Function GetNBTA(sIPAddr As String) As String
    Dim sText As String
    Dim aText() As String
    Dim i As Integer
    Dim sTemp As String
    Dim n As Integer
    Dim sResult As String

    sResult = ""

    sText = stdy.CaptureDOS("%COMSPEC% /c nbtstat -A " & sIPAddr, True)
    aText = Split(sText, vbNewLine)

    If stdS.Contains(aText(4), "Host not found.") = False Then
        n = 0
        For i = 0 To UBound(aText)
            sTemp = Replace(RTrim(LTrim(aText(i))), vbCr, "")
            aText(i) = sTemp
        Next

        Dim sIP As String
        Dim sLine As String
        Dim sMac As String

        sIP = stdS.Between(aText(2), "[", "]")

        n = 7
        Do While True
            If Len(aText(n)) = 0 Then
                Exit Do
            End If
            sLine = aText(n)
            n = n + 1
        Loop

        sLine = Replace(sLine, "  ", " ")
        sLine = Replace(sLine, "  ", " ")
        sLine = Replace(sLine, "  ", " ")

        sMac = Split(aText(n + 1), " = ")(1)
        sResult = sIP & vbTab & sMac & vbTab & Split(sLine, " ")(0)
    End If
    GetNBTA = sResult
End Function

'http://www.windowsitpro.com/WindowsScripting/Article/ArticleID/4734/4734.html
'----------------------------------------------------------------
' Function:     GetDomainComputers
' Description:  Returns a listing of NT Computer Accounts for a
'           given domain
' Parameters:   ByVal strDomain - Name of an NT Domain to retrieve
'           the list of Computer from.
' Returns:      Variant array of NT Computer names for the
'           specified domain.
'----------------------------------------------------------------
'Public Function GetDomainComputers(ByVal strDomain As String) As Variant()
'   Dim objIADsContainer As ActiveDs.IADsDomain ' Container
'   Dim objIADsComputer  As ActiveDs.IADsComputer
'   Dim vReturn() As Variant
'
'   ' connect to the computer.
'   Set objIADsContainer = GetObject("WinNT://" & strDomain)
'
'   ' set the filter to retrieve only objects of class Computer
'   objIADsContainer.Filter = Array("Computer")
'
'   ReDim vReturn(0)
'   For Each objIADsComputer In objIADsContainer
'      If Trim(vReturn(0)) <> "" Then
'         ReDim Preserve vReturn(UBound(vReturn) + 1)
'      End If
'      vReturn(UBound(vReturn)) = objIADsComputer.Name
'   Next
'
'   GetDomainComputers = vReturn
'End Function

'Function GetSpecs(sComputer As String) As String
'    Dim oWMI As Object
'    Dim cItems As SWbemObjectSet
'    Dim sQuery As String
'    Dim oItem As SWbemObject
'    Dim sResult As String
'
'    sResult = ""
'
'    On Error Resume Next
'    Set oWMI = GetObject("winmgmts:\\" & sComputer & "\root\CIMV2")
'    If Err.Number = 0 Then
'        sQuery = "SELECT Manufacturer, Model, Name, SystemType, TotalPhysicalMemory, DomainRole FROM Win32_ComputerSystem"
'        Set cItems = oWMI.ExecQuery(sQuery)
'        For Each oItem In cItems
'            With oItem
'                sResult = Trim(.Manfacturer)
'                sResult = sResult & vbTab & Trim(.Model)
'                sResult = sResult & vbTab & Trim(.Name)
'                sResult = sResult & vbTab & Trim(.SystemType)
'                sResult = sResult & vbTab & Trim(Int(((.TotalPhysicalMemory) / 1024) / 1024)) & " MB"
'
'                Select Case .DomainRole
'                    Case 0
'                        sResult = sResult & vbTab & "Standalone Workstation"
'                    Case 1
'                        sResult = sResult & vbTab & "Member Workstation"
'                    Case 2
'                        sResult = sResult & vbTab & "Standalone Server"
'                    Case 3
'                        sResult = sResult & vbTab & "Member Server"
'                    Case 4
'                        sResult = sResult & vbTab & "Backup Domain Controller"
'                    Case 5
'                        sResult = sResult & vbTab & "Primary Domain Controller"
'                End Select
'            End With
'        Next
'        Err.Clear
'        On Error GoTo 0
'    End If
'    GetSpecs = sResult
'End Function

'Function ResolveIP(sName As String) As String
'    Dim sHost As String
'    Dim sTarget As String
'    Dim cPingResults As SWbemObjectSet
'    Dim oPingResult As SWbemObject
'    Dim sResult As String
'
'    sResult = vbNullString
'    sHost = CreateObject("WScript.Network").ComputerName
'    sTarget = sName
'
'    Set cPingResults = GetObject("winmgmts:{impersonationLevel=impersonate}//" & _
'        sHost & "/root/cimv2").ExecQuery("SELECT * FROM Win32_PingStatus " & _
'        "WHERE Address = '" + sTarget + "'")
'
'    For Each oPingResult In cPingResults
'        If oPingResult.StatusCode = 0 Then
'            sResult = oPingResult.ProtocolAddress
'        End If
'    Next
'    ResolveIP = sResult
'End Function

'Function DetectTempFolder(sDir As String) As String
'    Dim sTemp1 As String
'    Dim sTemp2 As String
'    Dim sResult As String
'
'    sTemp1 = "c:\Windows\Temp"
'    sTemp2 = "c:\TEMP"
'    sResult = vbNullString
'
'    If ExistsFolder(V.Recall("COMPUTERNAME"), sTemp1) Then
'        sResult = "\\" & V.Recall("COMPUTERNAME") & Replace(sTemp1, "c:\", "\C$\")
'    ElseIf ExistsFolder(V.Recall("COMPUTERNAME"), sTemp2) Then
'        sResult = "\\" & V.Recall("COMPUTERNAME") & Replace(sTemp2, "c:\", "\C$\")
'    Else
'        '
'    End If
'    DetectTempFolder = sResult
'
'End Function

'Function ExistsFolder(sComputer As String, sDir As String) As Boolean
'    Dim oWMI As Object
'    Dim cFiles As SWbemObjectSet
'    Dim oFile As SWbemObject
'    Dim sUNC As String
'
'    sUNC = Replace(sDir, "\", "\\")
'
'    Set oWMI = GetObject("winmgmts:" _
'        & "{impersonationLevel=impersonate}!\\" & sComputer & "\root\cimv2")
'
'    Set cFiles = oWMI.ExecQuery _
'        ("Select * from CIM_Directory Where name = '" & sUNC & "'")
'
'    ExistsFolder = cFiles.Count > 0
'End Function

Public Function GetTheComputerName() As String

    Dim strComputerName As String ' Variable to return the path of computer name
    
    strComputerName = Space(250) ' Initilize the buffer to receive the string
    GetComputerName strComputerName, Len(strComputerName)
    strComputerName = Mid(Trim$(strComputerName), 1, Len(Trim$(strComputerName)) - 1)
    GetTheComputerName = strComputerName

End Function

Public Function ComputerName() As String
'    ComputerName = stdR.ValueEx(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Control\ComputerName\ComputerName", "ComputerName", REG_SZ, GetTheComputerName())
    ComputerName = stdR.ValueEx(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Control\ComputerName\ComputerName", "ComputerName", REG_SZ, vbNullString)
End Function

