VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Temp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetTempFileNameA Lib "kernel32" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempFileNameW Lib "kernel32" (lpszPath As Byte, lpPrefixString As Byte, ByVal wUnique As Long, lpName As Byte) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempPathWide Lib "kernel32" Alias "GetTempPathW" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Const MAX_FILENAME_LEN = 256
Private Const UNIQUE_NAME = &H0

Public Function TempDirWide() As String
    Dim lngRet As Long
    Dim strTempDir As String
    Dim lngBuf As Long
    strTempDir = String$(255, 0)
    lngBuf = Len(strTempDir)
    lngRet = GetTempPathWide(lngBuf, strTempDir)
    If lngRet > lngBuf Then
        strTempDir = String$(lngRet, 0)
        lngBuf = Len(strTempDir)
        lngRet = GetTempPath(lngBuf, strTempDir)
    End If
    TempDirWide = left$(strTempDir, lngRet)
End Function

Public Function TempDir() As String
    Dim lngRet As Long
    Dim strTempDir As String
    Dim lngBuf As Long
    strTempDir = String$(255, 0)
    lngBuf = Len(strTempDir)
    lngRet = GetTempPath(lngBuf, strTempDir)
    If lngRet > lngBuf Then
        strTempDir = String$(lngRet, 0)
        lngBuf = Len(strTempDir)
        lngRet = GetTempPath(lngBuf, strTempDir)
    End If
    TempDir = left$(strTempDir, lngRet)
End Function

Public Function TempFile(Create As Boolean, Optional lpPrefixString As Variant, Optional lpszPath As Variant) As String
    Dim lpTempFileName As String * 255
    Dim strTemp As String
    Dim lngRet As Long
    If IsMissing(lpszPath) Then
        lpszPath = TempDir
    End If
    If IsMissing(lpPrefixString) Then
        lpPrefixString = "tmp"
    End If
    lngRet = GetTempFileNameA(lpszPath, lpPrefixString, 0, lpTempFileName)
    strTemp = lpTempFileName
    lngRet = InStr(lpTempFileName, vbNullChar)
    strTemp = left$(lpTempFileName, lngRet - 1)
    If Create = False Then
        Kill strTemp
        Do Until Dir(strTemp) = vbNullString: DoEvents: Loop
    End If
    TempFile = strTemp
End Function

Public Function GetTempFileName(Optional sSeed As String) As String
    Dim s As String, s2 As String, s3 As String
    
    If sSeed = vbNullString Then
        s2 = TempDir 'GetTempPath
        s3 = App.ExeName
        s = Space(Len(s2) + MAX_FILENAME_LEN)
        Call GetTempFileNameA(s2, s3, UNIQUE_NAME, s)
        GetTempFileName = left$(s, InStr(s, vbNullChar) - 1)
    Else
        s2 = TempDir 'GetTempPath
        s3 = App.ExeName
        s = Space(Len(s2) + MAX_FILENAME_LEN)
        Call GetTempFileNameA(s2, "XXX", UNIQUE_NAME, s)
        s = left$(s, InStr(s, vbNullChar) - 1)
        Mid(s, Len(s2) + 1, 3) = left(sSeed, 3)
        GetTempFileName = UnicodeTempFile(s)
    End If
End Function

Public Function UnicodeTempFile(ByVal iFileName As String) As String
    Dim bFileName() As Byte
    Dim bPath()     As Byte
    Dim bTempFile() As Byte
    Dim bPrefix()   As Byte
    Dim nRet        As Long
    Dim sTempFile   As String
    bPath = LocGetFilePath(iFileName) & vbNullChar
    bPrefix = left(Mid(iFileName, Len(LocGetFilePath(iFileName)) + 2), 3) & vbNullChar
    ReDim bTempFile((Len(iFileName) + 12) * 2)
    nRet = GetTempFileNameW(bPath(0), bPrefix(0), 0, bTempFile(0))
    sTempFile = bTempFile
    UnicodeTempFile = left(sTempFile, InStr(sTempFile & vbNullChar, vbNullChar) - 1)
End Function

Public Function LocGetFilePath(ByVal iCFName As String) As String
    Dim nPos     As Long
    nPos = InStrRev(iCFName, "\")
    If nPos > 0 Then
        LocGetFilePath = left(iCFName, nPos - 1)
    Else
        LocGetFilePath = ""
    End If
End Function
