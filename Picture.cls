VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Picture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public iType As Integer
Dim sFontName As String
Dim iFontSize As Integer
Dim iBG As Integer
Dim iFG As Integer
Public Function TextHeight(ByVal sText As String) As Integer
    TextHeight = Form1.Picture1.TextHeight(sText)
End Function

Public Function TextWidth(ByVal sText As String) As Integer
    TextWidth = Form1.Picture1.TextWidth(sText)
End Function

Public Sub Attributes(ByVal sAttribList As String)
    Dim i As Integer
    Dim c As String * 1
    For i = 1 To Len(sAttribList)
        c = Mid$(sAttribList, i, 1)
        Select Case c
        Case "B", "b"
            Form1.Picture1.FontBold = (c = "B")
        Case "I", "i"
            Form1.Picture1.FontItalic = (c = "I")
        Case "S", "s"
            Form1.Picture1.FontStrikethru = (c = "S")
        Case "U", "u"
            Form1.Picture1.FontUnderline = (c = "U")
        Case "T", "t"
            Form1.Picture1.FontTransparent = (c = "T")
        End Select
    Next
End Sub


Public Sub ScaleMode(ByVal iMode As Integer)
    Form1.Picture1.ScaleMode = iMode
End Sub

Public Sub Colour(ByVal iFGColour As Long, ByVal iBGColour As Long)
    Form1.Picture1.BackColor = iBGColour
    Form1.Picture1.ForeColor = iFGColour
End Sub

Public Sub Font(ByVal sName As String, ByVal iSize As Integer)
    Form1.Picture1.Font.Name = sName
    Form1.Picture1.Font.Size = iSize
End Sub


Public Sub AtSay(ByVal nX As Integer, ByVal nY As Integer, ByVal sText As String)
    At nX, nY
    Say sText
End Sub


Public Sub At(ByVal nX As Integer, ByVal nY As Integer)
    Form1.Picture1.CurrentX = nX
    Form1.Picture1.CurrentY = nY
End Sub

Public Sub Say(ByVal sText As String)
    Form1.Picture1.Print sText
End Sub


Public Sub SavePic(sFilename As String, Optional iType As Integer = vbPicTypeBitmap)
    With Form1.Picture1
        Select Case iType
        Case vbPicTypeIcon, vbPicTypeNone
            SavePicture Form1.Picture1.Image, sFilename & ".ico"
        Case vbPicTypeBitmap
            SavePicture Form1.Picture1.Image, sFilename & ".BMP"
        Case vbPicTypeMetafile
            SavePicture Form1.Picture1.Image, sFilename & ".WMF"
        Case vbPicTypeEMetafile
            SavePicture Form1.Picture1.Image, sFilename & ".EMF"
        End Select
    End With
End Sub

Public Sub LoadPic(sFilename As String, Optional nDestX As Integer = 0, Optional nDestY As Integer = 0)
    Form1.Picture1.PaintPicture LoadPicture(sFilename), nDestX, nDestY
    iType = Form1.Picture1.Picture.Type
End Sub


Public Sub Dimensions(ByVal nHeight As Integer, ByVal nWidth As Integer)
    With Form1
         .Picture1.Width = nWidth '.Picture1.ScaleX(nWidth, vbPixels, vbTwips)
         .Picture1.Height = nHeight ' .Picture1.ScaleY(nHeight, vbPixels, vbTwips)
    End With
End Sub


Private Sub Class_Initialize()
    Load Form1
    Form1.Picture1.AutoRedraw = True
End Sub

Private Sub Class_Terminate()
    Unload Form1
End Sub
