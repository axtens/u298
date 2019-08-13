VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Wallpaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Dim reg As New StdRegistry
'Public Enum WALLPAPER_STYLES
'    WS_CENTRED = 0
'    WS_TILED = 1
'    WS_STRETCHED = 2
'End Enum

Private Declare Function IIDFromString Lib "ole32" ( _
  ByVal lpszIID As Long, _
  iid As Any) As Long
 
Private Declare Function CoCreateInstance Lib "ole32" ( _
  rclsid As Any, _
  ByVal pUnkOuter As Long, _
  ByVal dwClsContext As Long, _
  riid As Any, _
  ByVal ppv As Long) As Long
 
Private Declare Function CallWindowProcA Lib "user32" ( _
  ByVal addr As Long, _
  ByVal p1 As Long, _
  ByVal p2 As Long, _
  ByVal p3 As Long, _
  ByVal p4 As Long) As Long
 
Private Declare Sub RtlMoveMemory Lib "kernel32" ( _
  pDst As Any, _
  pSrc As Any, _
  ByVal dlen As Long)
 
Private Const CLSCTX_INPROC_SERVER  As Long = 1&
 
Private Const CLSID_ActiveDesktop   As String = "{75048700-EF1F-11D0-9888-006097DEACF9}"
Private Const IID_ActiveDesktop     As String = "{F490EB00-1240-11D1-9888-006097DEACF9}"
 
Private Type GUID
  Data1                   As Long
  Data2                   As Integer
  Data3                   As Integer
  Data4(7)                As Byte
End Type
 
Private Type IActiveDesktop
  ' IUnknown
  QueryInterface          As Long
  AddRef                  As Long
  Release                 As Long
  ' IActiveDesktop
  ApplyChanges            As Long
  GetWallpaper            As Long
  SetWallpaper            As Long
  GetWallpaperOptions     As Long
  SetWallpaperOptions     As Long
  GetPattern              As Long
  SetPattern              As Long
  GetDesktopItemOptions   As Long
  SetDesktopItemOptions   As Long
  AddDesktopItem          As Long
  AddDesktopItemWithUI    As Long
  ModifyDesktopItem       As Long
  RemoveDesktopItem       As Long
  GetDesktopItemCount     As Long
  GetDesktopItem          As Long
  GetDesktopItemByID      As Long
  GenerateDesktopItemHtml As Long
  AddUrl                  As Long
  GetDesktopItemBySource  As Long
End Type
 
Private Enum AD_APPLY
  AD_APPLY_SAVE = &H1
  AD_APPLY_HTMLGEN = &H2
  AD_APPLY_REFRESH = &H4
  AD_APPLY_ALL = &H7
  AD_APPLY_FORCE = &H8
  AD_APPLY_BUFFERED_REFRESH = &H10
  AD_APPLY_DYNAMICREFRESH = &H20
End Enum

Public Picture As Object
'local variable(s) to hold property value(s)
'Private mvarWallpaper As Variant 'local copy


'Public Property Let Wallpaper(ByVal vData As Variant)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.Wallpaper = 5
'
''if filename ends with HTM or HTML handle one way, else the other
'    reg.ClassKey = HKEY_CURRENT_USER
'    reg.SectionKey = "Software\Microsoft\Internet Explorer\Desktop\General"
'    reg.ValueKey = "Wallpaper"
'    reg.ValueType = REG_SZ
'    reg.Default = ""
'    reg.Value = vData
'    mvarWallpaper = vData
'End Property
'
'Public Property Get Wallpaper() As Variant
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.Wallpaper
'
''if mvarWallpaper <> vbnullstring then look at its filetype and handle accordingly
'
'    Dim sResult As String
'    sResult = ""
'    reg.ClassKey = HKEY_CURRENT_USER
'    reg.SectionKey = "Software\Microsoft\Internet Explorer\Desktop\General"
'    reg.ValueKey = "Wallpaper"
'    reg.ValueType = REG_SZ
'    reg.Default = ""
'    sResult = reg.Value
'    Wallpaper = sResult
'End Property
'
'Public Property Let WallpaperStyle(ByVal vData As WALLPAPER_STYLES)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.WallpaperStyle = 5
'
''if mvarWallpaper <> vbnullstring then look at its filetype and handle accordingly
'
'    reg.ClassKey = HKEY_CURRENT_USER
'    reg.SectionKey = "Software\Microsoft\Internet Explorer\Desktop\General"
'    reg.ValueKey = "WallpaperStyle"
'    reg.ValueType = REG_SZ
'    reg.Default = ""
'    reg.Value = CStr(vData)
'    If vData = WS_TILED Then
'        TileWallpaper 1
'    Else
'        TileWallpaper 0
'    End If
'End Property
'
'Public Property Get WallpaperStyle() As WALLPAPER_STYLES
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.WallpaperStyle
'
''if mvarWallpaper <> vbnullstring then look at its filetype and handle accordingly
'    reg.ClassKey = HKEY_CURRENT_USER
'    reg.SectionKey = "Software\Microsoft\Internet Explorer\Desktop\General"
'    reg.ValueKey = "WallpaperStyle"
'    reg.ValueType = REG_SZ
'    reg.Default = ""
'    WallpaperStyle = CInt(reg.Value)
'End Property

'Public Sub SetWallpaperStyle(ws As WALLPAPER_STYLES)
'    'User Key: [HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System]
'    'Value Name: Wallpaper , WallpaperStyle
'    'Data Type: REG_SZ (String Value)
'    reg.ClassKey = HKEY_CURRENT_USER
'    'reg.SectionKey = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
'    reg.SectionKey = "Control Panel\Desktop"
'    reg.ValueKey = "WallpaperStyle"
'    reg.ValueType = REG_SZ
'    reg.Default = ""
'    reg.Value = CStr(ws)
'End Sub


'Public Sub ForceActiveDesktopOff()
'    reg.ClassKey = HKEY_CURRENT_USER
'    reg.SectionKey = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
'    reg.ValueKey = "ForceActiveDesktopOn"
'    reg.ValueType = REG_DWORD
'    reg.Default = 0
'    reg.Value = 0
'End Sub
'
'Public Sub ForceActiveDesktopOn()
'    reg.ClassKey = HKEY_CURRENT_USER
'    reg.SectionKey = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
'    reg.ValueKey = "ForceActiveDesktopOn"
'    reg.ValueType = REG_DWORD
'    reg.Default = 0
'    reg.Value = 1
'End Sub

'Public Function CurrentWallpaper() As String
'    CurrentWallpaper = reg.ValueEx(HKEY_CURRENT_USER, "Control Panel\Desktop", "Wallpaper", REG_SZ, "")
'End Function
'
'Public Sub SetWallpaper(cFilename As Variant)
'    reg.ClassKey = HKEY_CURRENT_USER
'    reg.SectionKey = "Control Panel\Desktop"
'    reg.ValueKey = "Wallpaper"
'    reg.ValueType = REG_SZ
'    reg.Default = ""
'    reg.Value = cFilename
'End Sub

'Public Sub RefreshDesktop()
'    Dim oShell As Object
'    Set oShell = CreateObject("WScript.Shell")
'    oShell.Run "%windir%\System32\RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters", 1, True
'End Sub
''
'Private Sub TileWallpaper(n As Integer)
'    reg.ClassKey = HKEY_CURRENT_USER
'    reg.SectionKey = "Software\Microsoft\Internet Explorer\Desktop\General"
'    reg.ValueKey = "TileWallpaper"
'    reg.ValueType = REG_SZ
'    reg.Default = ""
'    reg.Value = CStr(n)
'End Sub

'http://www.vbforums.com/showthread.php?t=367342
Public Function ActiveDesktopSetWallpaper(ByVal strFile As String) As Boolean
 
    Dim vtbl            As IActiveDesktop
    Dim vtblptr         As Long
 
    Dim classid         As GUID
    Dim iid             As GUID
 
    Dim OBJ             As Long
    Dim hRes            As Long
 
    ' CLSID (BSTR) to CLSID (GUID)
    hRes = IIDFromString(StrPtr(CLSID_ActiveDesktop), classid)
    If hRes <> 0 Then
        Exit Function
    End If
 
    ' IID (BSTR) to IID (GUID)
    hRes = IIDFromString(StrPtr(IID_ActiveDesktop), iid)
    If hRes <> 0 Then
        Exit Function
    End If
 
    ' create an instance of IActiveDesktop
    ' (Set IActiveDesktop = New IActiveDesktop)
    hRes = CoCreateInstance(classid, 0, CLSCTX_INPROC_SERVER, iid, VarPtr(OBJ))
    If hRes <> 0 Then
        Exit Function
    End If
 
    ' obj points now to a pointer to the VTable
    ' of IActiveDesktop
    '
    ' dereference the VTable pointer
    RtlMoveMemory vtblptr, ByVal OBJ, 4
    ' copy the VTable to our IActiveDesktop structure
    RtlMoveMemory vtbl, ByVal vtblptr, Len(vtbl)
 
    ' call IActiveDesktop::SetWallpaper
    '
    ' the first parameter is always the object pointer
    ' the return value should always be a HRESULT (0 = S_OK)
    hRes = CallPointer(vtbl.SetWallpaper, OBJ, StrPtr(strFile), 0)
    If hRes = 0 Then
        ActiveDesktopSetWallpaper = True
    End If
 
    ' call IActiveDesktop::ApplyChanges
    hRes = CallPointer(vtbl.ApplyChanges, OBJ, AD_APPLY_ALL Or AD_APPLY_FORCE)
 
    ' release IActiveDesktop to free memory
    ' (Set IActiveDesktop = Nothing)
    CallPointer vtbl.Release, OBJ
End Function
 
Private Function CallPointer( _
    ByVal fnc As Long, _
    ParamArray params() _
) As Long
 
  Dim btASM(&HEC00& - 1)  As Byte
  Dim pASM                As Long
  Dim i                   As Integer
 
  pASM = VarPtr(btASM(0))
 
  AddByte pASM, &H58                  ' POP EAX
  AddByte pASM, &H59                  ' POP ECX
  AddByte pASM, &H59                  ' POP ECX
  AddByte pASM, &H59                  ' POP ECX
  AddByte pASM, &H59                  ' POP ECX
  AddByte pASM, &H50                  ' PUSH EAX
 
  For i = UBound(params) To 0 Step -1
    AddPush pASM, CLng(params(i))     ' PUSH dword
  Next
 
  AddCall pASM, fnc                   ' CALL rel addr
  AddByte pASM, &HC3                  ' RET
 
  CallPointer = CallWindowProcA(VarPtr(btASM(0)), 0, 0, 0, 0)
End Function
 
Private Sub AddPush(pASM As Long, lng As Long)
  AddByte pASM, &H68
  AddLong pASM, lng
End Sub
 
Private Sub AddCall(pASM As Long, addr As Long)
  AddByte pASM, &HE8
  AddLong pASM, addr - pASM - 4
End Sub
 
Private Sub AddLong(pASM As Long, lng As Long)
  RtlMoveMemory ByVal pASM, lng, 4
  pASM = pASM + 4
End Sub
 
Private Sub AddByte(pASM As Long, BT As Byte)
  RtlMoveMemory ByVal pASM, BT, 1
  pASM = pASM + 1
End Sub

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

Public Function LoadPic(sFilename As String, Optional nDestX As Integer = 0, Optional nDestY As Integer = 0)
    Form1.Picture1.PaintPicture LoadPicture(sFilename), nDestX, nDestY
    LoadPic = Form1.Picture1.Picture.Type
End Function


Public Sub Dimensions(ByVal nHeight As Integer, ByVal nWidth As Integer)
    With Form1
         .Picture1.WIDTH = nWidth '.Picture1.ScaleX(nWidth, vbPixels, vbTwips)
         .Picture1.Height = nHeight ' .Picture1.ScaleY(nHeight, vbPixels, vbTwips)
    End With
End Sub


Private Sub Class_Initialize()
    Load Form1
    Form1.Picture1.AutoRedraw = True
    Set Picture = Form1.Picture1.Image
End Sub

Private Sub Class_Terminate()
    Set Picture = Nothing
    Unload Form1
End Sub

