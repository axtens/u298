VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IEDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim stdS As New Strings
Dim stdF As New Files
Dim stdy As New System

Dim oFSO As Scripting.FileSystemObject

Const TEMPORARY_FOLDER As Integer = 2

Dim oShell As Object
Dim oIE As Object
Dim oFolder As Scripting.Folder
Public sName As String
Dim oStream As Object
Dim oDIV As Object

Dim nHeight As Integer
Dim nWidth As Integer
Dim nTop As Integer
Dim nLeft As Integer

Private Sub Class_Initialize()
    Dim sComputer As String
    Dim colItems As SWbemObjectSet
    Dim oItem As SWbemObject
    Dim arrMonitors(10, 1) As Integer
    Dim nMonitorCount As Integer
    
    Set oShell = CreateObject("WScript.Shell")
    Set oIE = CreateObject("InternetExplorer.Application")
    Set oFSO = New FileSystemObject
    
    sComputer = "."
    Set colItems = GetObject("winmgmts:\\" & sComputer & "\root\cimv2").InstancesOf("Win32_DesktopMonitor")
    
    On Error Resume Next
    nMonitorCount = 0
    For Each oItem In colItems
        arrMonitors(nMonitorCount, 0) = oItem.ScreenHeight
        arrMonitors(nMonitorCount, 1) = oItem.ScreenWidth
        nMonitorCount = nMonitorCount + 1
    Next
    On Error GoTo 0
    
    nHeight = arrMonitors(0, 0)
    nWidth = arrMonitors(0, 1)
    
    Set oFolder = oFSO.GetSpecialFolder(TEMPORARY_FOLDER)
    sName = oFSO.BuildPath(oFolder, oFSO.GetTempName) & ".html"
    stdF.WriteFileU sName, Join(Array("<HTML><HEAD><TITLE>Information</TITLE></HEAD>", _
                                     "<BODY SCROLL='NO'><CENTER><FONT FACE='arial black'> <HR COLOR='BLACK'>", _
                                     "<DIV id='MakeMeAnObject'></DIV>", _
                                     "<HR COLOR='BLACK'></FONT></CENTER></BODY></HTML>"), vbCrLf), 2 'WF_CREATE
    nTop = 0
    nLeft = 0
End Sub

Sub Init(sPosition As Variant)
    'NW, N, NE, W, CENTRE, E, SW, S, SE
    Select Case CStr(sPosition)
    Case "NW"
        nTop = 0
        nLeft = 0
    Case "N"
        nTop = 0
        nLeft = (nWidth / 2) - 250
    Case "NE"
        nTop = 0
        nLeft = nWidth - 500
    Case "W"
        nTop = (nHeight / 2) - 55
        nLeft = 0
    Case "CENTRE"
        nTop = (nHeight / 2) - 55
        nLeft = (nWidth / 2) - 250
    Case "E"
        nTop = (nHeight / 2) - 55
        nLeft = nWidth - 500
    Case "SW"
        nTop = nHeight - 110
        nLeft = 0
    Case "S"
        nTop = nHeight - 110
        nLeft = (nWidth / 2) - 250
    Case "SE"
        nTop = nHeight - 110
        nLeft = nWidth - 500
    Case Else
        nTop = 0
        nLeft = 0
    End Select
        
    SetupIE (sName)
    Set oDIV = oIE.document.All("MakeMeAnObject")
End Sub

Private Sub Class_Terminate()
    'Close IE and delete the file
    oIE.Quit
    '~ optionally you may want to get rid of the temp file
End Sub

Public Sub Display(sMsg As String, nMillisec As Integer)
    oDIV.InnerHTML = sMsg
    stdy.Sleep nMillisec
End Sub

Private Sub SetupIE(File2Load As String)
     oIE.Navigate File2Load
     oIE.ToolBar = False
     oIE.StatusBar = False
     oIE.Resizable = False

     Do
     Loop While oIE.Busy

     oIE.WIDTH = 500
     oIE.Height = 110
     oIE.left = nLeft
     oIE.top = nTop
     oIE.Visible = True
     oShell.AppActivate ("Microsoft Internet Explorer")
End Sub

