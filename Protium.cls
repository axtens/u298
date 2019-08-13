VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Protium"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim Files As New Files
Dim traces As New traces
Dim Strings As New Strings
Const PROTIUM As String = "C:\protium\bin\pdb.exe"

Public Function AskProtium(ByVal Instructions As Variant, ByVal CommandLine As Variant) As Variant
    Dim sTempFile As String
    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")
    sTempFile = Files.RandomInputFile(".p")
    traces.trace sTempFile
    Files.WriteFileU sTempFile, Instructions, US2.WF_CREATE
    oShell.Run Strings.DQ(PROTIUM) & " " & Strings.DQ(sTempFile) & " " & CommandLine, 0, True
    AskProtium = Files.ReadFileU(Files.RandomOutputFile(sTempFile, ".out"))
End Function

