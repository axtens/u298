VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Environments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim REG As New Registry

Function GetUserEnv(strEnvVar As String) As String
    Dim sResult As String
    With CreateObject("WScript.Shell")
        GetUserEnv = .Environment("USER")(UCase(strEnvVar))
    End With
End Function

Function GetSystemEnv(strEnvVar As String) As String
    Dim sResult As String
    With CreateObject("WScript.Shell")
        GetSystemEnv = .Environment("SYSTEM")(UCase(strEnvVar))
    End With
End Function

Function GetVolatileEnv(strEnvVar As String) As String
    Dim sResult As String
    With CreateObject("WScript.Shell")
        GetVolatileEnv = .Environment("VOLATILE")(UCase(strEnvVar))
    End With
End Function

Function GetProcessEnv(strEnvVar As String) As String
    Dim sResult As String
    With CreateObject("WScript.Shell")
        GetProcessEnv = .Environment("PROCESS")(UCase(strEnvVar))
    End With
End Function

Private Sub Class_Initialize()
    REG.Needs "WScript.Shell"
    If Err.Number = vbObjectError + 999 Then
        MsgBox "Class 'Environments' problem: " & Err.Description
        End
    End If

End Sub
