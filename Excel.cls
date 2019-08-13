VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Excel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim oExcel As Object
Dim oWorkbook As Object
Dim oSheet As Object
Dim REG As New Registry

Sub ExcelStart(Optional bHidden As Boolean = False)
    Set oExcel = CreateObject("Excel.Application")
    If bHidden = True Then
        oExcel.Visible = False
        oExcel.Interactive = False
    Else
        oExcel.Visible = True
        oExcel.Interactive = True
    End If
End Sub

Sub ExcelNewSheet()
    Set oWorkbook = oExcel.Workbooks.add()
    Set oSheet = oWorkbook.ActiveSheet
End Sub

Sub ExcelColumnNames(aHeadings() As Variant, Optional bBold As Boolean = True)
    Dim i As Integer
    For i = 0 To UBound(aHeadings)
        oSheet.Cells(1, i + 1).Value = aHeadings(i)
        oSheet.Cells(1, i + 1).Font.Bold = True
    Next
End Sub

Private Sub Class_Initialize()
    REG.Needs "Excel.Application"
    If Err.Number = vbObjectError + 999 Then
        MsgBox "Class 'Excel' problem: " & Err.Description
        End
    End If

End Sub
