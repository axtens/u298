VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INI2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Const STDINI_E_FILENOTFOUND As Long = vbObjectError + 1006
Const STDINI_E_INIFILENOTLOADED As Long = vbObjectError + 1007

Dim sBuffer As String
Dim sFilename As String
Dim stdF As New Files

Private Sub Class_Initialize()
    sBuffer = ""
    sFilename = ""
End Sub

Private Function getfile(strFile As Variant) As String
    Dim oFSO As New Scripting.FileSystemObject
    Dim oStream As Scripting.TextStream
    Dim sData As String

    sData = ""
    sFilename = strFile

    If Not stdF.Exists(sFilename) Then
        Err.Raise STDINI_E_FILENOTFOUND, , "STDINI_E_FILENOTFOUND, " & sFilename & " not found."
    End If
    Set oStream = oFSO.OpenTextFile(sFilename, ForReading, False, TristateMixed)
    sData = oStream.ReadAll

    oStream.Close
    getfile = sData
End Function

Public Function Load(sIniName As Variant) As Boolean
    sBuffer = getfile(sIniName)
    Load = True
End Function

Public Sub Clear()
    sBuffer = ""
    sFilename = ""
End Sub

Private Function separateField(ByVal sFrom As String, ByVal sStart As String, ByVal sEnd As String) As String
    Dim nBegin As Integer, nEnd As Integer
    nBegin = InStr(1, sFrom, sStart, 1)
    If nBegin > 0 Then
        nBegin = nBegin + Len(sStart)
        nEnd = InStr(nBegin, sFrom, sEnd, 1)
        If nEnd = 0 Then
            nEnd = InStr(nBegin, sFrom, vbNewLine, 1)
        End If
        If nEnd = 0 Then
            nEnd = Len(sFrom) + 1
        End If
        separateField = Mid(sFrom, nBegin, nEnd - nBegin)
    End If
End Function

Public Function GetValue(sSection As Variant, sKey As Variant, strDefault As Variant) As String
    Dim nPosSection As Integer
    Dim nPosEndSection As Integer
    Dim sSectionContents As String
    Dim bFound As Boolean
    Dim sValue As String

    If sFilename = "" Then
        Err.Raise STDINI_E_INIFILENOTLOADED, , "STDINI_E_INIFILENOTLOADED, INI file not loaded."
    End If

    sValue = strDefault
    If sBuffer <> "" Then
        nPosSection = InStr(1, sBuffer, "[" & sSection & "]", vbTextCompare)
        If nPosSection > 0 Then
            nPosEndSection = InStr(nPosSection, sBuffer, vbNewLine & "[")
            If nPosEndSection = 0 Then
                nPosEndSection = Len(sBuffer) + 1
            End If
            sSectionContents = Mid(sBuffer, nPosSection, nPosEndSection - nPosSection)
            If InStr(1, sSectionContents, vbNewLine & sKey & "=", vbTextCompare) > 0 Then
                bFound = True
                sValue = separateField(sSectionContents, vbNewLine & sKey & "=", vbNewLine)
            End If
        End If
    End If
    GetValue = sValue
End Function

'    Public Function lambda(sSection, sKey)
'        Dim sResult
'        sResult = GetValue(sSection, sKey, "")
'        Assert sResult <> "", NotSpecified(sSection, sKey)
'        '~ Log "INI @ [" & sSection & "] " & sKey & "=" & sResult
'        lambda = sResult
'    End Function

Public Function HasSection(sSection As Variant) As Boolean
    Dim bSuccess As Boolean
    Dim nOffset As Integer
    nOffset = InStr(1, sBuffer, "[" & sSection & "]", vbTextCompare)
    bSuccess = (nOffset > 0)
    HasSection = bSuccess
End Function

'Public Function NotSpecified(aSection As String, aTag As String) As String
'    NotSpecified = "[" & CStr(aSection) & "] " & CStr(aTag) & " not specified in " & sFilename
'End Function
'end class



