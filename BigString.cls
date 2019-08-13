VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BigString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim vBigStr As Variant
Dim ARR As New Arrays
Public Function GetStrArray() As Variant
    GetStrArray = vBigStr
End Function

Public Sub ClearStr()
    ReDim vBigStr(0)
End Sub

Public Function GetStr(Optional ByVal vJoiner As Variant = "") As String
    GetStr = Join(vBigStr, vJoiner)
End Function

Public Sub PutStr(ByVal vString As Variant)
    Set ARR = New Arrays
    ARR.AAdd_PostIncrement vBigStr, vString
End Sub

Private Sub Class_Initialize()
    ReDim vBigStr(0)
End Sub

Private Sub Class_Terminate()
    ReDim vBigStr(0)
    vBigStr = Empty
End Sub
