VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Traces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Trace"
Option Explicit
Public Tracing As Boolean
Public TraceFile As String
Dim oFSO As Scripting.FileSystemObject
Dim stdF As New Files
Dim stdS As New Strings
Public Sub Trace(Optional ByVal sTag As Variant = "!@#$%^&*()_")
Attribute Trace.VB_UserMemId = 0
    Dim oFSO As New Scripting.FileSystemObject
    
    If Tracing Then
        If sTag = "!@#$%^&*()_" Then
            stdF.WriteFileU TraceFile, CStr(Now) & vbNewLine, IIF(oFSO.FileExists(TraceFile), 1, 2)
        Else
            stdF.WriteFileU TraceFile, CStr(sTag) & vbNewLine, IIF(oFSO.FileExists(TraceFile), 1, 2) 'append , create
        End If
    End If
End Sub
Sub ClearTrace()
    Set oFSO = New FileSystemObject
    If Tracing Then
        If oFSO.FileExists(TraceFile) Then
            Kill TraceFile
        End If
    End If
End Sub

Private Sub Class_Initialize()
    TraceFile = App.Path & "\" & App.ExeName & ".log"
End Sub
