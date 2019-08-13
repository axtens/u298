VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INI3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ___________________________________________________________________
'
'  VBScript File:   IniFileClass.vbs
'  Author:          Frank-Peter Schultze
'
'  Updates:         http://www.fpschultze.de/modules/smartfaq/faq.php?faqid=51
'  Enhancement Req.
'  and Bug Reports: support@fpschultze.de
'
'  Built/Tested On: Windows 2003
'  Requirements:    WSH 1.0+, VBScript 5.0+
'
'  Purpose:         Provides a class to read from/write to ini files
'
'  Last Update:     28-Sep-2006
' ___________________________________________________________________
'
'  This script is a rewritten/improved version of Jean-Luc Antoine's
'  class to accesss ini files, class_ini.vbs. URL of original code:
'  http://www.interclasse.com/scripts/class_ini.php
' ___________________________________________________________________
'

'Class IniFile

    Public Filename
    Public Section
    Public Key

    Private objFso As Scripting.FileSystemObject
    Private objIni


    Private Sub Class_Initialize()

        Set objFso = New Scripting.FileSystemObject '("Scripting.FileSystemObject")

    End Sub


    Private Sub Class_Terminate()

        Set objFso = Nothing

    End Sub


    Private Sub SectionStartEnd(ByRef lngSectionStart, ByRef lngSectionEnd)

        Dim arrContent, i, s

        lngSectionStart = -2
        lngSectionEnd = -1
        arrContent = Split(Content, vbNewLine, -1, 1)
        If (UBound(arrContent) > -1) Then
            If (Section <> "") Then
                For i = LBound(arrContent) To UBound(arrContent)
                    s = Trim(arrContent(i))
                    If (LCase(s) = "[" & LCase(Section) & "]") Then
                        lngSectionStart = i
                    ElseIf (left(s, 1) = "[") And _
                        (right(s, 1) = "]") And _
                        (lngSectionStart >= 0) Then
                        lngSectionEnd = i - 1
                        Exit For
                    End If
                Next
                If (lngSectionStart >= 0) And (lngSectionEnd < 0) Then
                    lngSectionEnd = UBound(arrContent)
                End If
            Else
                lngSectionStart = -1
                lngSectionEnd = UBound(arrContent)
            End If
            If (lngSectionStart > -2) Then
                If (arrContent(lngSectionEnd) = "") Then
                    lngSectionEnd = lngSectionEnd - 1
                End If
            End If
        End If

    End Sub


    Property Get Content()

        Content = ""
        If objFso.FileExists(Filename) Then
            Set objIni = objFso.OpenTextFile(Filename, ForReading)
            Content = objIni.ReadAll
            objIni.Close
            Set objIni = Nothing
        End If

    End Property


    Property Let Content(strContent)

        Set objIni = objFso.CreateTextFile(Filename, True)
        objIni.Write strContent
        objIni.Close
        Set objIni = Nothing

    End Property


    Property Get Value()

        Dim lngSectionStart, lngSectionEnd, lngIndex, arrContent
        Dim strLine, i, s

        Value = Chr(255)
        SectionStartEnd lngSectionStart, lngSectionEnd
        If (lngSectionStart > -2) Then
            arrContent = Split(Content, vbNewLine, -1, 1)
            For lngIndex = lngSectionStart + 1 To lngSectionEnd
                strLine = arrContent(lngIndex)
                i = InStr(1, strLine, "=", 1)
                If (i > 0) Then
                    s = left(strLine, i - 1)
                    s = Trim(s)
                    If (LCase(s) = LCase(Key)) Then
                        Value = Mid(strLine, i + 1)
                        Value = Trim(Value)
                        Exit For
                    End If
                End If
            Next
        End If

    End Property


    Property Let Value(strValue)

        Dim lngSectionStart, lngSectionEnd, arrContent
        Dim lngIndex, lngIndex2, strContent, blnKeyNotFound
        Dim strLine, i, s

        SectionStartEnd lngSectionStart, lngSectionEnd
        If (lngSectionStart < -1) And (strValue <> "") Then
            strContent = Content & vbNewLine _
                & "[" & Section & "]" & vbNewLine _
                    & Key & "=" & strValue
        Else
            blnKeyNotFound = True
            arrContent = Split(Content, vbNewLine, -1, 1)
            For lngIndex = lngSectionStart + 1 To lngSectionEnd
                strLine = arrContent(lngIndex)
                i = InStr(1, strLine, "=", 1)
                If (i > 0) Then
                    s = left(strLine, i - 1)
                    s = Trim(s)
                    If (LCase(s) = LCase(Key)) Then
                        blnKeyNotFound = False
                        If (strValue <> "") Then
                            arrContent(lngIndex) = Key & "=" & strValue
                        Else
                            For lngIndex2 = lngIndex To UBound(arrContent) - 1
                                arrContent(lngIndex2) = arrContent(lngIndex2 + 1)
                            Next
                            ReDim Preserve arrContent(UBound(arrContent) - 1)
                        End If
                        Exit For
                    End If
                End If
            Next
            If blnKeyNotFound And (strValue <> "") Then
                ReDim Preserve arrContent(UBound(arrContent) + 1)
                For lngIndex = UBound(arrContent) To lngSectionEnd + 2 Step -1
                    arrContent(lngIndex) = arrContent(lngIndex - 1)
                Next
                arrContent(lngSectionEnd + 1) = Key & "=" & strValue
            End If
            strContent = arrContent(0)
            For lngIndex = 1 To UBound(arrContent)
                strContent = strContent & vbNewLine & arrContent(lngIndex)
            Next
        End If
        Content = strContent

    End Property

'End Class


