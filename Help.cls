VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'**********************************************************************
'**modWinHelp
'**(c) 1997-1999, Shadow Mountain Tech. All rights reserved.
'**
'**HTMLHelp API function declarations and constant definitions.
'**
'**HTML constants and declarations were extracted from
'**clsHTMLHelp.cls by David Liske.
'**
'**********************************************************************
Option Explicit

Public sHelpFile As String
Private HHCookie As Long

' HTML Help Constants
Private Const HH_DISPLAY_TOPIC = &H0            '  WinHelp equivalent
Private Const HH_DISPLAY_TOC = &H1              '  WinHelp equivalent
Private Const HH_DISPLAY_INDEX = &H2            '  WinHelp equivalent
Private Const HH_DISPLAY_SEARCH = &H3           '  WinHelp equivalent
Private Const HH_SET_WIN_TYPE = &H4
Private Const HH_GET_WIN_TYPE = &H5
Private Const HH_GET_WIN_HANDLE = &H6
Private Const HH_SYNC = &H9
Private Const HH_ADD_NAV_UI = &HA               ' not currently implemented
Private Const HH_ADD_BUTTON = &HB               ' not currently implemented
Private Const HH_GETBROWSER_APP = &HC           ' not currently implemented
Private Const HH_KEYWORD_LOOKUP = &HD           '  WinHelp equivalent
Private Const HH_DISPLAY_TEXT_POPUP = &HE       ' display string resource id
                                                ' or text in a popup window
                                                ' value in dwData
Private Const HH_HELP_CONTEXT = &HF             '  display mapped numeric
Private Const HH_CLOSE_ALL = &H12               '  WinHelp equivalent
Private Const HH_ALINK_LOOKUP = &H13            '  ALink version of
                                                '  HH_KEYWORD_LOOKUP
Private Const HH_SET_GUID = &H1A                ' For Microsoft Installer -- dwData is a pointer to the GUID string

Private Const HH_INITIALIZE = &H1C
Private Const HH_UNINITIALIZE = &H1D
Private Const HH_FTS_DEFAULT_PROXIMITY = -1

' HTML Help window constants. These are also used
' in the window definitions in HHP files
Private Const HHWIN_PROP_ONTOP = &H2              ' Top-most window (not currently implemented)
Private Const HHWIN_PROP_NOTITLEBAR = &H4         ' no title bar
Private Const HHWIN_PROP_NODEF_STYLES = &H8       ' no default window styles (only HH_WINTYPE.dwStyles)
Private Const HHWIN_PROP_NODEF_EXSTYLES = &H10    ' no default extended window styles (only HH_WINTYPE.dwExStyles)
Private Const HHWIN_PROP_TRI_PANE = &H20          ' use a tri-pane window
Private Const HHWIN_PROP_NOTB_TEXT = &H40         ' no text on toolbar buttons
Private Const HHWIN_PROP_POST_QUIT = &H80         ' post WM_QUIT message when window closes
Private Const HHWIN_PROP_AUTO_SYNC = &H100        ' automatically ssync contents and index
Private Const HHWIN_PROP_TRACKING = &H200         ' send tracking notification messages
Private Const HHWIN_PROP_TAB_SEARCH = &H400       ' include search tab in navigation pane
Private Const HHWIN_PROP_TAB_HISTORY = &H800      ' include history tab in navigation pane
Private Const HHWIN_PROP_TAB_BOOKMARKS = &H1000   ' include bookmark tab in navigation pane
Private Const HHWIN_PROP_CHANGE_TITLE = &H2000    ' Put current HTML title in title bar
Private Const HHWIN_PROP_NAV_ONLY_WIN = &H4000    ' Only display the navigation window
Private Const HHWIN_PROP_NO_TOOLBAR = &H8000      ' Don't display a toolbar
Private Const HHWIN_PROP_MENU = &H10000           ' Menu
Private Const HHWIN_PROP_TAB_ADVSEARCH = &H20000  ' Advanced FTS UI.
Private Const HHWIN_PROP_USER_POS = &H40000       ' After initial creation, user controls window size/position

Private Const HHWIN_PARAM_PROPERTIES = &H2        ' valid fsWinProperties
Private Const HHWIN_PARAM_STYLES = &H4            ' valid dwStyles
Private Const HHWIN_PARAM_EXSTYLES = &H8          ' valid dwExStyles
Private Const HHWIN_PARAM_RECT = &H10             ' valid rcWindowPos
Private Const HHWIN_PARAM_NAV_WIDTH = &H20        ' valid iNavWidth
Private Const HHWIN_PARAM_SHOWSTATE = &H40        ' valid nShowState
Private Const HHWIN_PARAM_INFOTYPES = &H80        ' valid apInfoTypes
Private Const HHWIN_PARAM_TB_FLAGS = &H100        ' valid fsToolBarFlags
Private Const HHWIN_PARAM_EXPANSION = &H200       ' valid fNotExpanded
Private Const HHWIN_PARAM_TABPOS = &H400          ' valid tabpos
Private Const HHWIN_PARAM_TABORDER = &H800        ' valid taborder
Private Const HHWIN_PARAM_HISTORY_COUNT = &H1000  ' valid cHistory
Private Const HHWIN_PARAM_CUR_TAB = &H2000        ' valid curNavType

Private Const HHWIN_BUTTON_EXPAND = &H2           ' Expand/contract button
Private Const HHWIN_BUTTON_BACK = &H4             ' Back button
Private Const HHWIN_BUTTON_FORWARD = &H8          ' Forward button
Private Const HHWIN_BUTTON_STOP = &H10            ' Stop button
Private Const HHWIN_BUTTON_REFRESH = &H20         ' Refresh button
Private Const HHWIN_BUTTON_HOME = &H40            ' Home button
Private Const HHWIN_BUTTON_BROWSE_FWD = &H80      ' not implemented
Private Const HHWIN_BUTTON_BROWSE_BCK = &H100     ' not implemented
Private Const HHWIN_BUTTON_NOTES = &H200          ' not implemented
Private Const HHWIN_BUTTON_CONTENTS = &H400       ' not implemented
Private Const HHWIN_BUTTON_SYNC = &H800           ' Locate button
Private Const HHWIN_BUTTON_OPTIONS = &H1000       ' Options button
Private Const HHWIN_BUTTON_PRINT = &H2000         ' Print button
Private Const HHWIN_BUTTON_INDEX = &H4000         ' not implemented
Private Const HHWIN_BUTTON_SEARCH = &H8000        ' not implemented
Private Const HHWIN_BUTTON_HISTORY = &H10000      ' not implemented
Private Const HHWIN_BUTTON_BOOKMARKS = &H20000    ' not implemented
Private Const HHWIN_BUTTON_JUMP1 = &H40000        ' Jump1 button
Private Const HHWIN_BUTTON_JUMP2 = &H80000        ' Jump2 button
Private Const HHWIN_BUTTON_ZOOM = &H100000        ' Font sizing button
Private Const HHWIN_BUTTON_TOC_NEXT = &H200000    ' Browse next TOC topic button
Private Const HHWIN_BUTTON_TOC_PREV = &H400000    ' Browse previous TOC topic button

' Default button set
Private Const HHWIN_DEF_BUTTONS = (HHWIN_BUTTON_EXPAND Or HHWIN_BUTTON_BACK Or HHWIN_BUTTON_OPTIONS Or HHWIN_BUTTON_PRINT)

' Button IDs
Private Const IDTB_EXPAND = 200
Private Const IDTB_CONTRACT = 201
Private Const IDTB_STOP = 202
Private Const IDTB_REFRESH = 203
Private Const IDTB_BACK = 204
Private Const IDTB_HOME = 205
Private Const IDTB_SYNC = 206
Private Const IDTB_PRINT = 207
Private Const IDTB_OPTIONS = 208
Private Const IDTB_FORWARD = 209
Private Const IDTB_NOTES = 210             ' not implemented
Private Const IDTB_BROWSE_FWD = 211
Private Const IDTB_BROWSE_BACK = 212
Private Const IDTB_CONTENTS = 213          ' not implemented
Private Const IDTB_INDEX = 214             ' not implemented
Private Const IDTB_SEARCH = 215            ' not implemented
Private Const IDTB_HISTORY = 216           ' not implemented
Private Const IDTB_BOOKMARKS = 217         ' not implemented
Private Const IDTB_JUMP1 = 218
Private Const IDTB_JUMP2 = 219
Private Const IDTB_CUSTOMIZE = 221
Private Const IDTB_ZOOM = 222
Private Const IDTB_TOC_NEXT = 223
Private Const IDTB_TOC_PREV = 224

Private Enum HHACT_
  HHACT_TAB_CONTENTS
  HHACT_TAB_INDEX
  HHACT_TAB_SEARCH
  HHACT_TAB_HISTORY
  HHACT_TAB_FAVORITES
    
  HHACT_EXPAND
  HHACT_CONTRACT
  HHACT_BACK
  HHACT_FORWARD
  HHACT_STOP
  HHACT_REFRESH
  HHACT_HOME
  HHACT_SYNC
  HHACT_OPTIONS
  HHACT_PRINT
  HHACT_HIGHLIGHT
  HHACT_CUSTOMIZE
  HHACT_JUMP1
  HHACT_JUMP2
  HHACT_ZOOM
  HHACT_TOC_NEXT
  HHACT_TOC_PREV
  HHACT_NOTES

  HHACT_LAST_ENUM
End Enum

Private Enum HHWIN_NAVTYPE_
  HHWIN_NAVTYPE_TOC
  HHWIN_NAVTYPE_INDEX
  HHWIN_NAVTYPE_SEARCH
  HHWIN_NAVTYPE_HISTORY       ' not implemented
  HHWIN_NAVTYPE_FAVORITES     ' not implemented
End Enum

Enum HHWIN_NAVTAB_
  HHWIN_NAVTAB_TOP
  HHWIN_NAVTAB_LEFT
  HHWIN_NAVTAB_BOTTOM
End Enum

Private Const HH_MAX_TABS = 19               ' maximum number of tabs

Private Enum HH_TAB_
  HH_TAB_CONTENTS
  HH_TAB_INDEX
  HH_TAB_SEARCH
  HH_TAB_HISTORY
  HH_TAB_FAVORITES
End Enum

Private Type RECT
    left    As Long
    top     As Long
    right   As Long
    bottom  As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type HH_WINTYPE
  cbStruct As Long            ' IN: size of this structure including all Information Types
  fUniCodeStrings As Long     ' IN/OUT: TRUE if all strings are in UNICODE
  pszType  As String          ' IN/OUT: Name of a type of window
  fsValidMembers As Long      ' IN: Bit flag of valid members (HHWIN_PARAM_)
  fsWinProperties As Long     ' IN/OUT: Properties/attributes of the window (HHWIN_)
  pszCaption As String        ' IN/OUT: Window title
  dwStyles  As Long           ' IN/OUT: Window styles
  dwExStyles As Long          ' IN/OUT: Extended Window styles
  rcWindowPos As RECT         ' IN: Starting position, OUT: current position
  nShowState As Long          ' IN: show state (e.g., SW_SHOW)
  hwndHelp As Long            ' OUT: window handle
  hwndCaller As Long          ' OUT: who called this window
  paInfoTypes As Long         ' IN: Pointer to an array of Information Types

  ' The following members are only valid if HHWIN_PROP_TRI_PANE is set

  hwndToolBar As Long         ' OUT: toolbar window in tri-pane window
  hwndNavigation As Long      ' OUT: navigation window in tri-pane window
  hwndHTML As Long            ' OUT: window displaying HTML in tri-pane window
  iNavWidth As Long           ' IN/OUT: width of navigation window
  rcHTML As RECT              ' OUT: HTML window coordinates

  pszToc As String            ' IN: Location of the table of contents file
  pszIndex As String          ' IN: Location of the index file
  pszFile As String           ' IN: Default location of the html file
  pszHome As String           ' IN/OUT: html file to display when Home button is clicked
  fsToolBarFlags As Long      ' IN: flags controling the appearance of the toolbar
  fNotExpanded As Long        ' IN: TRUE/FALSE to contract or expand, OUT: current state
  curNavType As Long          ' IN/OUT: UI to display in the navigational pane
  tabpos As HHWIN_NAVTAB_     ' IN/OUT: HHWIN_NAVTAB_TOP, HHWIN_NAVTAB_LEFT, or HHWIN_NAVTAB_BOTTOM
  idNotify As Long            ' IN: ID to use for WM_NOTIFY messages
  tabOrder(HH_MAX_TABS) As Byte ' IN/OUT: tab order: Contents, Index, Search, History, Favorites, Reserved 1-5, Custom tabs
  cHistory As Long            ' IN/OUT: number of history items to keep (default is 30)
  pszJump1 As String          ' Text for HHWIN_BUTTON_JUMP1
  pszJump2 As String          ' Text for HHWIN_BUTTON_JUMP2
  pszUrlJump1 As String       ' URL for HHWIN_BUTTON_JUMP1
  pszUrlJump2 As String       ' URL for HHWIN_BUTTON_JUMP2
  rcMinSize As RECT           ' Minimum size for window (ignored in version 1)
  cbInfoTypes As Long         ' size of paInfoTypes;
End Type

' UDT for text popups
Private Type HH_POPUP
  cbStruct As Long                         ' sizeof this structure
  hinst As Long                               ' instance handle for string resource
  idString As Long                            ' string resource id, or text id if pszFile
                                              ' is specified in HtmlHelp call
  pszText As String                           ' used if idString is zero
  pt As POINTAPI                              ' top center of popup window
  clrForeground As ColorConstants             ' either use VB constant or &HBBGGRR
  clrBackground As ColorConstants             ' either use VB constant or &HBBGGRR
  rcMargins As RECT                           ' amount of space between edges of window and
                                              ' text, -1 for each member to ignore
  pszFont As String                           ' facename, point size, char set, BOLD ITALIC
                                              ' UNDERLINE
End Type

' UDT for keyword and ALink searches
Private Type HH_AKLINK
  cbStruct          As Long
  fReserved         As Boolean
  pszKeywords       As String
  pszUrl            As String
  pszMsgText        As String
  pszMsgTitle       As String
  pszWindow         As String
  fIndexOnFail      As Boolean
End Type

' UDT for accessing the Search tab
Private Type HH_FTS_QUERY
  cbStruct          As Long
  fUniCodeStrings   As Long
  pszSearchQuery    As String
  iProximity        As Long
  fStemmedSearch    As Long
  fTitleOnly        As Long
  fExecute          As Long
  pszWindow         As String
End Type

'~ Const VK_H = 72
'~ Const VK_E = 69
'~ Const VK_L = 76
'~ Const VK_O = 79
Const INPUT_MOUSE As Long = 0
Const INPUT_KEYBOARD As Long = 1
Const INPUT_HARDWARE As Long = 2
Const KEYEVENTF_EXTENDEDKEY As Long = &H1
Const KEYEVENTF_KEYUP As Long = &H2
Const KEYEVENTF_UNICODE As Long = &H4
Const KEYEVENTF_SCANCODE As Long = &H8
Const XBUTTON1 As Long = &H1
Const XBUTTON2 As Long = &H2
Const MOUSEEVENTF_MOVE As Long = &H1
Const MOUSEEVENTF_LEFTDOWN As Long = &H2
Const MOUSEEVENTF_LEFTUP As Long = &H4
Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Const MOUSEEVENTF_RIGHTUP As Long = &H10
Const MOUSEEVENTF_MIDDLEDOWN As Long = &H20
Const MOUSEEVENTF_MIDDLEUP As Long = &H40
Const MOUSEEVENTF_XDOWN As Long = &H80
Const MOUSEEVENTF_XUP As Long = &H100
Const MOUSEEVENTF_WHEEL As Long = &H800
Const MOUSEEVENTF_VIRTUALDESK As Long = &H4000
Const MOUSEEVENTF_ABSOLUTE As Long = &H8000

Const VK_SHIFT        As Long = &H10
Const VK_CONTROL          As Long = &H11
Const VK_MENU         As Long = &H12
Const VK_ESCAPE           As Long = &H1B
Const VK_BACK         As Long = &H8
Const VK_TAB          As Long = &H9
Const VK_RETURN           As Long = &HD
Const VK_PRIOR        As Long = &H21
Const VK_NEXT         As Long = &H22
Const VK_END          As Long = &H23
Const VK_HOME         As Long = &H24
Const VK_LEFT         As Long = &H25
Const VK_UP           As Long = &H26
Const VK_RIGHT        As Long = &H27
Const VK_DOWN         As Long = &H28
Const VK_SELECT           As Long = &H29
Const VK_PRINT        As Long = &H2A
Const VK_EXECUTE          As Long = &H2B
Const VK_SNAPSHOT         As Long = &H2C
Const VK_INSERT           As Long = &H2D
Const VK_DELETE           As Long = &H2E
Const VK_HELP         As Long = &H2F
Const VK_NUMPAD0          As Long = &H60
Const VK_NUMPAD1          As Long = &H61
Const VK_NUMPAD2          As Long = &H62
Const VK_NUMPAD3          As Long = &H63
Const VK_NUMPAD4          As Long = &H64
Const VK_NUMPAD5          As Long = &H65
Const VK_NUMPAD6          As Long = &H66
Const VK_NUMPAD7          As Long = &H67
Const VK_NUMPAD8          As Long = &H68
Const VK_NUMPAD9          As Long = &H69
Const VK_MULTIPLY         As Long = &H6A
Const VK_ADD          As Long = &H6B
Const VK_SEPARATOR        As Long = &H6C
Const VK_SUBTRACT         As Long = &H6D
Const VK_DECIMAL          As Long = &H6E
Const VK_DIVIDE           As Long = &H6F
Const VK_F1           As Long = &H70
Const VK_F2           As Long = &H71
Const VK_F3           As Long = &H72
Const VK_F4           As Long = &H73
Const VK_F5           As Long = &H74
Const VK_F6           As Long = &H75
Const VK_F7           As Long = &H76
Const VK_F8           As Long = &H77
Const VK_F9           As Long = &H78
Const VK_F10          As Long = &H79
Const VK_F11          As Long = &H7A
Const VK_F12          As Long = &H7B
Const VK_OEM_1        As Long = &HBA   ' ',:' for US
Const VK_OEM_PLUS         As Long = &HBB   ' '+' any country
Const VK_OEM_COMMA        As Long = &HBC   ' ',' any country
Const VK_OEM_MINUS        As Long = &HBD   ' '-' any country
Const VK_OEM_PERIOD       As Long = &HBE   ' '.' any country
Const VK_OEM_2        As Long = &HBF   ' '/?' for US
Const VK_OEM_3        As Long = &HC0   ' '`~' for US
Const VK_MEDIA_NEXT_TRACK     As Long = &HB0
Const VK_MEDIA_PREV_TRACK     As Long = &HB1
Const VK_MEDIA_STOP           As Long = &HB2
Const VK_MEDIA_PLAY_PAUSE     As Long = &HB3
Const VK_LWIN      As Long = &H5B
Const VK_RWIN      As Long = &H5C

Private Type MOUSEINPUT
  dx As Long
  dy As Long
  mouseData As Long
  dwFlags As Long
  time As Long
  dwExtraInfo As Long
End Type

Private Type KEYBDINPUT
  wVk As Integer
  wScan As Integer
  dwFlags As Long
  time As Long
  dwExtraInfo As Long
End Type

Private Type HARDWAREINPUT
  uMsg As Long
  wParamL As Integer
  wParamH As Integer
End Type

Private Type GENERALINPUT
  dwType As Long
  xi(0 To 23) As Byte
End Type

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function HTMLHelp Lib "hhctrl.ocx" Alias "HtmlHelpW" (ByVal hWnd As Long, ByVal lpHelpFile As Long, ByVal wCommand As Long, ByVal dwData As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public Sub HHInitialize()
    Call HTMLHelp(0, 0, HH_INITIALIZE, VarPtr(HHCookie))
End Sub

Public Sub HHUninitialize()
    Call HTMLHelp(0, 0, HH_UNINITIALIZE, HHCookie)
End Sub

Public Function HHDisplayHeadTopic(ByVal lHwnd As Long) As Long
'    Dim lDW As Long
'    lDW = GetDesktopWindow()
    HHDisplayHeadTopic = HTMLHelp(lHwnd, StrPtr(sHelpFile), HH_DISPLAY_TOPIC, 0)
End Function

Public Function HHKeywordLookup(ByVal lHwnd As Long, Optional sKeyword As String = "") As Long
'    Dim lDW As Long
    Dim link As HH_AKLINK
'    lDW = GetDesktopWindow()
    If sKeyword <> vbNullString Then
        
        link.cbStruct = Len(link)
        link.fReserved = False
        link.pszKeywords = sKeyword
        link.fIndexOnFail = True
        HHKeywordLookup = HTMLHelp(lHwnd, StrPtr(sHelpFile), HH_KEYWORD_LOOKUP, VarPtr(link))
    Else
        HHKeywordLookup = HTMLHelp(lHwnd, StrPtr(sHelpFile), HH_KEYWORD_LOOKUP, 0)
    End If
End Function

Public Function HHHelpContents(ByVal lHwnd As Long) As Long
  ' Force the Help window to display
  ' the Contents file (*.hhc) in the left pane
'    Dim lDW As Long
'    lDW = GetDesktopWindow()
    HHHelpContents = HTMLHelp(lHwnd, StrPtr(sHelpFile & ">Main"), HH_DISPLAY_TOC, 0)
End Function

Public Function HHHelpIndex(ByVal lHwnd As Long, Optional toSearch As String = "") As Long
  ' Force the Help window to display
  ' the Contents file (*.hhc) in the left pane
'    Dim lDW As Long
'    lDW = GetDesktopWindow()
    Dim lDI As Long
    If toSearch = vbNullString Then
        lDI = 0
    Else
        lDI = StrPtr(toSearch)
    End If
    HHHelpIndex = HTMLHelp(lHwnd, StrPtr(sHelpFile), HH_DISPLAY_INDEX, lDI)

    If toSearch <> vbNullString Then
        Sleep 100 'milliseconds
        
        SendK 0, VK_MENU
        SendK 0, AscB("D")
        SendK KEYEVENTF_KEYUP, AscB("D")
        SendK KEYEVENTF_KEYUP, VK_MENU
    End If
End Function

Public Function HHDisplaySearch(ByVal lHwnd As Long, Optional toSearch As String = "") As Long
'    Dim lDW As Long
'    lDW = GetDesktopWindow()
    Dim oQuery As HH_FTS_QUERY
    oQuery.cbStruct = Len(oQuery)
    oQuery.fStemmedSearch = 0
    oQuery.fTitleOnly = 0
    oQuery.fUniCodeStrings = True
    oQuery.iProximity = HH_FTS_DEFAULT_PROXIMITY
    oQuery.pszSearchQuery = toSearch
    oQuery.pszWindow = ""
    oQuery.fExecute = True
    HHDisplaySearch = HTMLHelp(lHwnd, StrPtr(sHelpFile), HH_DISPLAY_SEARCH, VarPtr(oQuery))
    
    'Dim i As Integer
    'Dim b As Byte
    'toSearch = StrConv(toSearch, vbFromUnicode)
'    For i = 1 To Len(toSearch)
'        b = Asc(Mid(toSearch, i, 1))
'        Debug.Print b;
'        SendKey b
'    Next
'    SendKey 13
    If toSearch <> vbNullString Then
        StringSendKeyU toSearch '& vbCr
        'send Alt-L
'        SendK KEYEVENTF_SCANCODE, 0, 38
'        SendK KEYEVENTF_SCANCODE + KEYEVENTF_KEYUP, 0, 38
'
'        'send Alt-D
'        SendK KEYEVENTF_SCANCODE, 0, 32
'        SendK KEYEVENTF_SCANCODE + KEYEVENTF_KEYUP, 0, 32
        
        Sleep 100 'milliseconds
        
        SendK 0, VK_MENU
        SendK 0, AscB("L")
        SendK KEYEVENTF_KEYUP, AscB("L")
        SendK KEYEVENTF_KEYUP, VK_MENU
        
        Sleep 100 'milliseconds
        
        SendK 0, VK_MENU
        SendK 0, AscB("D")
        SendK KEYEVENTF_KEYUP, AscB("D")
        SendK KEYEVENTF_KEYUP, VK_MENU
    End If
End Function


'~ Private Sub Form_KeyPress(KeyAscii As Integer)
    '~ 'Print the key on the form
    '~ Me.Print Chr$(KeyAscii);
'~ End Sub

'~ Private Sub Form_Paint()
    '~ 'KPD-Team 2000
    '~ 'URL: http://www.allapi.net/
    '~ 'E-Mail: KPDTeam@Allapi.net
    '~ 'Clear the form
    '~ Me.Cls
    '~ 'call the SendKey-function
    '~ SendKey VK_H
    '~ SendKey VK_E
    '~ SendKey VK_L
    '~ SendKey VK_L
    '~ SendKey VK_O
'~ End Sub

Private Sub SendKey(bKey As Byte)
    Dim GInput(0 To 1) As GENERALINPUT
    Dim KInput As KEYBDINPUT
    KInput.wVk = bKey  'the key we're going to press
    KInput.dwFlags = 0 'press the key
    'copy the structure into the input array's buffer.
    GInput(0).dwType = INPUT_KEYBOARD   ' keyboard input
    CopyMemory GInput(0).xi(0), KInput, Len(KInput)
    'do the same as above, but for releasing the key
    KInput.wVk = bKey  ' the key we're going to realease
    KInput.dwFlags = KEYEVENTF_KEYUP  ' release the key
    GInput(1).dwType = INPUT_KEYBOARD  ' keyboard input
    CopyMemory GInput(1).xi(0), KInput, Len(KInput)
    'send the input now
    Call SendInput(2, GInput(0), Len(GInput(0)))
End Sub


Private Sub StringSendKey(sString As String)
    Dim i As Integer
    For i = 1 To Len(sString)
        SendKey Asc(Mid(sString, i, 1))
    Next
End Sub

Private Sub SendKeyU(wKeyU As Integer)
    Dim GInput(0 To 1) As GENERALINPUT
    Dim KInput As KEYBDINPUT
    KInput.wVk = 0  'the key we're going to press
    KInput.wScan = wKeyU 'press the key
    KInput.dwFlags = KEYEVENTF_UNICODE
    'copy the structure into the input array's buffer.
    GInput(0).dwType = INPUT_KEYBOARD   ' keyboard input
    CopyMemory GInput(0).xi(0), KInput, Len(KInput)
    'do the same as above, but for releasing the key
    KInput.wVk = 0  ' the key we're going to realease
    KInput.dwFlags = KEYEVENTF_UNICODE + KEYEVENTF_KEYUP  ' release the key
    GInput(1).dwType = INPUT_KEYBOARD  ' keyboard input
    CopyMemory GInput(1).xi(0), KInput, Len(KInput)
    'send the input now
    Call SendInput(2, GInput(0), Len(GInput(0)))
End Sub


Private Sub StringSendKeyU(sString As String)
    Dim i As Integer
    For i = 1 To Len(sString)
        SendKeyU AscW(Mid$(sString, i, 1))
    Next
End Sub

Private Sub KeyB(bKey As Byte)
    keybd_event bKey, 0, 0, 0   ' press
    keybd_event bKey, 0, KEYEVENTF_KEYUP, 0   ' release
End Sub

Private Sub StringKeyB(sString As String)
    Dim s As String
    Dim i As Integer
    s = ""
    For i = 1 To Len(sString)
        s = s & Asc(Mid(sString, i, 1))
        KeyB Asc(UCase(Mid(sString, i, 1)))
    Next
    
End Sub

'Private Sub Form_Paint()
'    'KPD-Team 2000
'    'URL: http://www.allapi.net/
'    'E-Mail: KPDTeam@Allapi.net
'    'Clear the form
'    Me.Cls
'    keybd_event VK_H, 0, 0, 0   ' press H
'    keybd_event VK_H, 0, KEYEVENTF_KEYUP, 0   ' release H
'    keybd_event VK_E, 0, 0, 0  ' press E
'    keybd_event VK_E, 0, KEYEVENTF_KEYUP, 0  ' release E
'    keybd_event VK_L, 0, 0, 0  ' press L
'    keybd_event VK_L, 0, KEYEVENTF_KEYUP, 0  ' release L
'    keybd_event VK_L, 0, 0, 0  ' press L
'    keybd_event VK_L, 0, KEYEVENTF_KEYUP, 0  ' release L
'    keybd_event VK_O, 0, 0, 0  ' press O
'    keybd_event VK_O, 0, KEYEVENTF_KEYUP, 0  ' release O
'End Sub

'http://www.pinvoke.net/default.aspx/user32.SendInput

Private Function SendK(ByVal lFlags As Long, ByVal bKey As Byte, Optional ByVal wScancode As Integer = 0) As Long
    Dim GInput(0 To 1) As GENERALINPUT
    Dim KInput As KEYBDINPUT
    KInput.wVk = bKey  'the key we're going to press
    KInput.wScan = wScancode
    KInput.dwFlags = lFlags 'press the key
    'copy the structure into the input array's buffer.
    GInput(0).dwType = INPUT_KEYBOARD   ' keyboard input
    CopyMemory GInput(0).xi(0), KInput, Len(KInput)
    'send the input now
    SendK = SendInput(1, GInput(0), Len(GInput(0)))
End Function

Private Sub Class_Initialize()
    sHelpFile = App.HelpFile
End Sub
