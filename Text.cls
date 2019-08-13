VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Text"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Init"
Dim PRN As Object
Dim lastPRN As Object

'Dim trace As New Traces

Public Sub WriteText(Optional vtext As Variant = "")
    'trace.trace "WriteText: vtext=" & vtext
    PRN.Print vtext
End Sub

Public Sub WriteLine(nX1 As Variant, nY1 As Variant, nX2 As Variant, nY2 As Variant)
    'trace.trace "WriteLine: " & nX1 & "," & nY1 & "," & nX2 & "," & nY2
    PRN.Line (nX1, nY1)-(nX2, nY2), , B
End Sub

Public Sub CentreText(Optional nLine As Variant = -1, Optional sText As Variant = "")
    'trace.trace "CentreText: " & nLine & "," & sText
    PRN.CurrentX = (PRN.ScaleWidth - PRN.TextWidth(sText)) / 2
    PRN.CurrentY = IIF(nLine = 1, PRN.CurrentY, nLine)
    PRN.Print sText
End Sub

Public Sub RightAlignText(Optional nLine As Variant = -1, Optional sText As Variant = " ")
    'trace.trace "RightAlignText: " & nLine & "," & sText
    PRN.CurrentX = PRN.ScaleWidth - PRN.TextWidth(sText)
    PRN.CurrentY = IIF(nLine = 1, PRN.CurrentY, nLine)
    PRN.Print sText
End Sub

Public Sub PointToPrinter()
    Set PRN = lastPRN
End Sub

Public Sub PointToPanel()
    Set PRN = Form1
End Sub

Public Sub Init(p As Object)
    'trace.trace "Text Init"
    Set lastPRN = p
    Set PRN = lastPRN
End Sub


