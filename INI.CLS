VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INI1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'INI File Functions...
Private Declare Function WritePrivateProfileStringA Lib "kernel32" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileStringA Lib "kernel32" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'~ Public Declare Function WritePrivateProfileStringW Lib "KERNEL32" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
'~ Public Declare Function GetPrivateProfileStringW Lib "KERNEL32" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileStringW Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpKeyName As Long, ByVal lpString As Long, ByVal lpFileName As Long) As Long
Private Declare Function GetPrivateProfileStringW Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpKeyName As Long, ByVal lpDefault As Long, ByVal lpReturnedString As Long, ByVal nSize As Long, ByVal lpFileName As Long) As Long

Function WriteINIA(sSection As Variant, sKeyName As Variant, sNewString As Variant, sINIFileName As Variant) As Boolean

    On Local Error Resume Next

    Call WritePrivateProfileStringA(CStr(sSection), CStr(sKeyName), CStr(sNewString), CStr(sINIFileName))

    WriteINIA = (Err.Number = 0)

End Function
Function ReadINIA(sSection As Variant, sKeyName As Variant, sINIFileName As Variant) As String

    On Local Error Resume Next

    Dim sRet As String

    sRet = String(255, Chr(0))

    'Note: INI Filename can point to a local ini file or a remote ini file...
    ReadINIA = left(sRet, GetPrivateProfileStringA(CStr(sSection), ByVal CStr(sKeyName), "", sRet, Len(sRet), CStr(sINIFileName)))

End Function

Function WriteINIU(sSection As String, sKeyName As String, sNewString As String, sINIFileName As String) As Boolean
    Dim oFSO As New Scripting.FileSystemObject
    Dim oStream As Scripting.TextStream
    Set oStream = oFSO.OpenTextFile(sINIFileName, ForAppending, True, TristateTrue)
    oStream.Close

    'On Local Error Resume Next

    Call WritePrivateProfileStringW(StrPtr(sSection), StrPtr(sKeyName), StrPtr(sNewString), StrPtr(sINIFileName))

    WriteINIU = (Err.Number = 0)

End Function
Function ReadINIU(sSection As String, sKeyName As String, sINIFileName As String) As String

'    On Local Error Resume Next

    Dim sRet As String
    Dim sEmpty As String
    
    sRet = String(255, Chr(0))
    sEmpty = vbNullString
    
    'Note: INI Filename can point to a local ini file or a remote ini file...
    ReadINIU = left(sRet, GetPrivateProfileStringW(StrPtr(sSection), StrPtr(sKeyName), StrPtr(sEmpty), StrPtr(sRet), Len(sRet), StrPtr(sINIFileName)))

End Function

