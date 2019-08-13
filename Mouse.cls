VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, _
    lpPoint As POINTAPI) As Long

' Get mouse X coordinates in pixels
'
' If a window handle is passed, the result is relative to the client area
' of that window, otherwise the result is relative to the screen

Public Function MouseX(Optional ByVal hWnd As Long) As Long
    Dim lpPoint As POINTAPI
    GetCursorPos lpPoint
    If hWnd Then ScreenToClient hWnd, lpPoint
    MouseX = lpPoint.X
End Function

' Get mouse Y coordinates in pixels
'
' If a window handle is passed, the result is relative to the client area
' of that window, otherwise the result is relative to the screen

Public Function MouseY(Optional ByVal hWnd As Long) As Long
    Dim lpPoint As POINTAPI
    GetCursorPos lpPoint
    If hWnd Then ScreenToClient hWnd, lpPoint
    MouseY = lpPoint.Y
End Function

'Francesco Balena
'http://www.devx.com/vb2themax/Tip/19124


Public Function Between(ByVal nNumber As Integer, ByVal nLowerBound As Integer, ByVal nUpperBound As Integer) As Boolean
    Between = ((nLowerBound <= nNumber) And (nNumber <= nUpperBound))
End Function


