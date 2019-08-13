VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Init"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Dim PRN As Printer
Dim pic As StdPicture ' New StdPicture

'local variable(s) to hold property value(s)
Private mvarWideness As Integer 'local copy
Private mvarHighness As Integer 'local copy

Public SHeight As Integer
Public SWidth As Integer
Public PHeight As Integer
Public PWidth As Integer

Public Property Let Highness(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Highness = 5
    mvarHighness = vData
End Property


Public Property Get Highness() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Highness
    Highness = mvarHighness
End Property



Public Property Let Wideness(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Wideness = 5
    mvarWideness = vData
End Property


Public Property Get Wideness() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Wideness
    Wideness = mvarWideness
End Property




'Dim trace As New Traces
Public Sub LoadFile(Filename As Variant)
    Set pic = LoadPicture(Filename)
    Highness = pic.Height
    Wideness = pic.WIDTH
    SHeight = PRN.ScaleHeight
    SWidth = PRN.ScaleWidth
    PHeight = PRN.Height
    PWidth = PRN.WIDTH
End Sub


Public Sub Show(sFilename As Variant, x1 As Variant, y1 As Variant)
    'trace.trace "Show: " & sFilename & "," & x1 & "," & y1
    'On Error Resume Next
    LoadFile sFilename
    'Set pic = LoadPicture(sFilename)
    PRN.PaintPicture pic, x1, y1
End Sub

Public Sub Constrained(sFilename As Variant, x1 As Variant, y1 As Variant, Optional X2 As Variant, Optional y2 As Variant)
    'trace.trace "Constrained: " & sFilename & "," & x1 & "," & y1 & "," & x2 & "," & y2
    'Set pic = LoadPicture(sFilename)
    LoadFile sFilename
    If IsMissing(y2) Then
        PRN.PaintPicture pic, x1, y1, X2
    ElseIf IsMissing(X2) Then
        PRN.PaintPicture pic, x1, y1
    Else
        PRN.PaintPicture pic, x1, y1, X2, y2
    End If
End Sub
'http://www.vbforums.com/showthread.php?t=404792
Public Sub Scaled(sFilename As Variant, x1 As Variant, y1 As Variant, W1 As Variant, H1 As Variant)
    Dim imgResizeW As Integer
    Dim imgResizeH As Integer
    LoadFile sFilename
    With pic
        imgResizeW = W1 ' this is the desired width
        imgResizeH = H1 ' this is the desired height
 
        If .WIDTH / imgResizeW > .Height / imgResizeH Then
            imgResizeH = .Height / (.WIDTH / imgResizeW)
        Else
            imgResizeW = .WIDTH / (.Height / imgResizeH)
        End If
    ' determines which one needs to shrink the most to retain aspect ratio
        PRN.PaintPicture pic, x1, y1, imgResizeW, imgResizeH
    ' now put the picture to the printer (prtitems(i,2) and (i,3) are the x/y coord on the page
    End With
End Sub

Public Sub Init(p As Object)
    'trace.trace "Pict Init"
    Set PRN = p
End Sub
