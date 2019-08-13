VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Emails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim CLP As New Clipboards
Dim ARR As New Arrays
Dim str As New Strings
Dim SYS As New System
Dim REG As New Registry

Public Sub Mailto(ByVal vTo As Variant, ByVal vSubj As Variant, ByVal vtext As Variant)
    Dim sCC As Variant
    Dim sTo As Variant
    Dim sCmd As String
    
    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")
    
    Dim Scr As Object

    
    If IsArray(vTo) Then
        sTo = str.URLEncode(vTo(0))
        sCC = str.URLEncode(vTo(1))
    Else
        sTo = str.URLEncode(vTo)
    End If
    
    vSubj = str.URLEncode(vSubj)
    vtext = str.URLEncode(vtext)
    
    If IsArray(vTo) Then
        sCmd = str.Subst("mailto:[1]?subject=[3]&cc=[2]&body=[4]", sTo, sCC, vSubj, vtext)
    Else
        sCmd = str.Subst("mailto:[1]?subject=[2]&body=[3]", sTo, vSubj, vtext)
    End If
    
    On Error Resume Next
    oShell.Run sCmd, 0, True
    If Err.Number <> 0 Then
        CLP.ClearClip
        CLP.TextToClip str.URLDecode(vtext)
        MsgBox "Windows didn't like the size of message. I've put it on the clipboard"
        sCmd = str.Subst("mailto:[1]?subject=[2]", sTo, vSubj)
        oShell.Run sCmd, 0, True
        SYS.Sleep 5000
    End If
End Sub

Private Sub Class_Initialize()
    REG.Needs "WScript.Shell"
    If Err.Number = vbObjectError + 999 Then
        MsgBox "Class 'Emails' problem: " & Err.Description
        End
    End If

End Sub
