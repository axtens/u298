VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Strings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Const StringMapArraysDifferent As Long = vbObjectError + 1000
Dim ARR As New Arrays

Public Enum AO_COMPARATORS
    AO_NOTEQUAL = 0
    AO_LESSTHAN = 1
    AO_LESSTHANEQUAL = 2
    AO_EQUAL = 3
    AO_GREATERTHANEQUAL = 4
    AO_GREATERTHAN = 5
End Enum
Public Function SplitSet(ByVal sString As Variant, ByVal sSet As Variant) As Variant
' return "(CANT.) DOG, CAT" as Array("(", "CANT", ".) ", "DOG", ", ", "CAT" )
  Dim aRes() As Variant
  Dim i As Long
  Dim sAlphaText As String
  Dim inText As Boolean
  Dim sOtherText As String
  Dim c As String
  
  ReDim aRes(0)
  inText = False
  sAlphaText = ""
  sOtherText = ""
  i = 1
  Do
    c = Mid$(sString, i, 1)
    If InStr(sSet, c) Then
      If inText = False Then
        If sOtherText <> vbNullString Then
          ARR.AAdd_PostIncrement aRes, sOtherText
          sOtherText = ""
        End If
        inText = True
      End If
      sAlphaText = sAlphaText & c
    Else
      If inText = True Then
        If sAlphaText <> vbNullString Then
          ARR.AAdd_PostIncrement aRes, sAlphaText
          sAlphaText = ""
        End If
        inText = False
      End If
      sOtherText = sOtherText & c
    End If
    i = i + 1
    If i > Len(sString) Then
      If inText Then
        ARR.AAdd_PostIncrement aRes, sAlphaText
      Else
        ARR.AAdd_PostIncrement aRes, sOtherText
      End If
      Exit Do
    End If
  Loop
  ARR.DeleteLast aRes
  SplitSet = aRes
End Function



' STRING FUNCTIONS
Function Subst(sText As Variant, ParamArray paArgList() As Variant) As String
    Dim sResult As String
    Dim sTemp As String
    Dim i As Integer
    Dim aArgList As Variant
    aArgList = paArgList
    sResult = sText

    For i = 0 To UBound(aArgList)
        sTemp = aArgList(i)
        If IsNull(sTemp) Then
            sTemp = ""
        End If
        sResult = Replace(sResult, "[" & CStr(i + 1) & "]", sTemp)
    Next

    Subst = sResult
End Function

'Function Subst2(sText As Variant, aArgList As Variant) As String
'    Dim sResult As String
'
'    sResult = Subst(sText, aArgList)
'
'    sResult = Replace(sResult, "\r", vbCr)
'    sResult = Replace(sResult, "\n", vbLf)
'    sResult = Replace(sResult, "\t", vbTab)
'
'    Subst2 = sResult
'End Function

Function Subst2(sText As Variant, ParamArray paArgList() As Variant) As String
    Dim sResult As String
    Dim aArgs As Variant
    aArgs = paArgList
    sResult = Subst(sText, aArgs)
    
    sResult = Replace(sResult, "\r", vbCr)
    sResult = Replace(sResult, "\n", vbLf)
    sResult = Replace(sResult, "\t", vbTab)
    
    Subst2 = sResult
End Function

Public Function ForceExtension(sFilename As Variant, sExtension As Variant) As String
    Dim sResult As String
    Dim nDotPos As Integer

    nDotPos = InStrRev(sFilename, ".")
    If nDotPos = 0 Then
        sResult = sFilename & sExtension
    Else
        sResult = left(sFilename, nDotPos - 1) & sExtension
    End If
    ForceExtension = sResult
End Function

Function DQ(s As Variant) As String
    DQ = Chr(34) & s & Chr(34)
End Function

Function HTMLWrap(sTag As Variant, sContent As Variant) As String
    Dim aFirst As Variant
    aFirst = Split(sTag, " ", 2)
    HTMLWrap = "<" & sTag & ">" & sContent & "</" & aFirst(0) & ">"
End Function

Function RemoveSpaces(sText As Variant) As String
    Dim sResult As String
    Dim s As String
    Dim i As Integer

    sResult = ""
    For i = 1 To Len(sText)
        s = Mid$(sText, i, 1)
        If s <> " " Then
            sResult = sResult & s
        End If
    Next
    RemoveSpaces = sResult
End Function
Function ZeroFill(nNum As Variant, nWidth As Variant) As String
    ZeroFill = right$(String$(nWidth - 1, "0") & CStr(nNum), nWidth)
End Function
Function Zerofill2(nNum As Variant) As String
    Zerofill2 = right$("00" & CStr(nNum), 2)
End Function

Function RightFill(sText As Variant, nLen As Variant, sFill As Variant) As String
    RightFill = left$(sText & String$(nLen, sFill), nLen)
End Function

Function LeftFill(sText As Variant, nLen As Variant, sFill As Variant) As String
    LeftFill = right$(String$(nLen, sFill) & sText, nLen)
End Function

Function StringMap(sText As Variant, aFrom As Variant, vTo As Variant) As String
    Dim sResult As String
    Dim n As Integer
    sResult = sText
    If IsArray(aFrom) And IsArray(vTo) Then
        If UBound(aFrom) <> UBound(vTo) Then
            Err.Raise StringMapArraysDifferent, , "StringMap: From and To arrays of different lengths"
        Else
            For n = 0 To UBound(aFrom)
                sResult = Replace(sResult, aFrom(n), vTo(n))
            Next
        End If
    Else
        If Not IsArray(vTo) Then
            For n = 0 To UBound(aFrom)
                sResult = Replace(sResult, aFrom(n), vTo)
            Next
        Else
            Err.Raise StringMapArraysDifferent, , "StringMap: From and To arrays of different lengths"
        End If
    End If
    StringMap = sResult
End Function

Function LeftOf(sText As Variant, sItem As Variant, Optional bCaseInsens As Boolean = True) As String
    Dim nPos As Integer
    Dim sResult As String
    sResult = sText
    If bCaseInsens Then
        nPos = InStr(LCase$(sText), LCase$(sItem))
    Else
        nPos = InStr(sText, sItem)
    End If
    If nPos > 0 Then
        sResult = left$(sText, nPos - 1)
    End If
    LeftOf = sResult
End Function

Function RightOf(sText As Variant, sItem As Variant, Optional bCaseInsens As Boolean = True) As String
    Dim nPos As Integer
    Dim sResult As String
    sResult = sText
    If bCaseInsens Then
        nPos = InStr(LCase$(sText), LCase$(sItem))
    Else
        nPos = InStr(sText, sItem)
    End If
    If nPos > 0 Then
        sResult = Mid$(sText, nPos + Len(sItem))
    End If
    RightOf = sResult
End Function

Function CountFields(strText As Variant, strDelim As Variant) As Integer
    Dim nOffset As Integer
    Dim nFoundAt As Integer
    Dim nCounter As Integer

    nCounter = 1
    nOffset = 1

    Do
        nFoundAt = InStr(nOffset, strText, strDelim)
        If nFoundAt > 0 Then
                nCounter = nCounter + 1
                nOffset = nFoundAt + Len(strDelim)
        Else
            Exit Do
        End If
    Loop

    CountFields = nCounter
End Function

Function NthField(sText As Variant, sDelimiter As Variant, nReqdField As Variant) As String
    Dim nOffset As Integer
    Dim nFoundAt As Integer
    Dim nCounter As Integer
    Dim sResult As String

    nCounter = 1
    nOffset = 1
    sResult = ""
    
    If nReqdField >= 1 Then

        Do
            nFoundAt = InStr(nOffset, sText, sDelimiter)
            If nFoundAt > 0 Then
                If nCounter = nReqdField Then
                    sResult = Mid$(sText, nOffset, nFoundAt - nOffset)
                    Exit Do
                Else
                    nCounter = nCounter + 1
                    nOffset = nFoundAt + Len(sDelimiter)
                End If
            Else
                If nCounter = nReqdField Then
                    sResult = Mid$(sText, nOffset)
                End If
                Exit Do
            End If
        Loop
    End If
    NthField = sResult
End Function

Function BeginsWith(sText As Variant, sBeginning As Variant, Optional bCaseInsensitive As Boolean = True) As Boolean
    Dim bResult As Boolean

    bResult = False
    If sBeginning = vbNullString Then
        bResult = True
    Else
        If sText <> vbNullString Then
            If bCaseInsensitive = True Then
                bResult = (left$(UCase$(sText), Len(sBeginning)) = UCase$(sBeginning))
            Else
                bResult = (left$(sText, Len(sBeginning)) = sBeginning)
            End If
        End If
    End If

    BeginsWith = bResult
End Function

Function EndsWith(sText As Variant, sEnding As Variant, Optional bCaseInsensitive As Boolean = True) As Boolean
    Dim bResult As Boolean

    bResult = False
    If sEnding = vbNullString Then
        bResult = True
    Else
        If sText <> vbNullString Then
            If bCaseInsensitive = True Then
                bResult = (right$(UCase$(sText), Len(sEnding)) = UCase$(sEnding))
            Else
                bResult = (right$(sText, Len(sEnding)) = sEnding)
            End If
        End If
    End If

    EndsWith = bResult
End Function

Function Contains(sText As Variant, sChunk As Variant, Optional bCaseInsensitive As Boolean = True) As Boolean
    Dim bResult As Boolean

    bResult = False
    If sChunk = vbNullString Then
        bResult = True
    Else
        If sText <> vbNullString Then
            If bCaseInsensitive = True Then
                bResult = (InStr(UCase$(sText), UCase$(sChunk)) > 0)
            Else
                bResult = (InStr(sText, sChunk) > 0)
            End If
        End If
    End If

    Contains = bResult
End Function

Function HexDump(sData As Variant) As String
    Dim sSlice As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim sResult As String
    Dim c As Byte
    Const WIDTH As Integer = 16
    
    On Error Resume Next
    
    sResult = ""
    For i = 1 To Len(sData) Step WIDTH
        sSlice = Mid$(sData, i, WIDTH)
        For j = 1 To WIDTH - 1
            sResult = sResult & right("00" & Hex(Asc(Mid$(sSlice, j, 1))), 2) & " "
            If Err.Number <> 0 Then Exit For
        Next
        sResult = sResult & right("00" & Hex(Asc(Mid$(sSlice, WIDTH, 1))), 2)
        
        sResult = sResult & ": "
        For k = 1 To WIDTH
            c = Asc(Mid$(sSlice, k, 1))
            sResult = sResult & Chr$(c)
        Next
        sResult = sResult & vbNewLine
    Next
    HexDump = sResult
End Function

Function AsString(X As Variant) As String
    Dim sResult As String
    If TypeName(X) = "Null" Then
        sResult = "Null"
    Else
        sResult = CStr(X)
    End If
    AsString = sResult
End Function

Function FirstWord(sText As Variant) As String
    Dim nSpace As Integer
    Dim sTemp As String
    sTemp = LTrim$(sText)
    nSpace = InStr(sTemp, " ")
    If nSpace > 0 Then
        sTemp = left$(sTemp, nSpace - 1)
    End If
    FirstWord = sTemp
End Function

Function LastLineOf(sData As Variant) As String
    Dim aData As Variant
    aData = Split(sData, vbCrLf)
    LastLineOf = aData(UBound(aData))
End Function

Function FirstLineOf(sData As Variant) As String
    FirstLineOf = Split(sData, vbCrLf)(0)
End Function

Function NthLineOf(n As Integer, sData As Variant) As String
    NthLineOf = Split(sData, vbCrLf)(n - 1)
End Function

Function Between(ByRef sText As Variant, sBegin As Variant, sEnd As Variant, Optional bDestructive = False) As String
    Dim nBegin As Integer
    Dim nEnd As Integer
    Dim sResult As String
    
    nBegin = InStr(sText, sBegin)
    If nBegin = 0 Then
        If bDestructive = True Then
            sText = ""
        End If
        Between = ""
        Exit Function
    End If
    nEnd = InStr(nBegin + Len(sBegin), sText, sEnd)
    sResult = Mid$(sText, nBegin + Len(sBegin), nEnd - nBegin - Len(sBegin))
    If bDestructive = True Then
        sText = Mid(sText, nBegin + Len(sBegin) + Len(sResult) + 1)
    End If
    Between = sResult
End Function

Function EndsWithSet(sText As Variant, vEnding As Variant, Optional bCaseInsensitive As Boolean = True) As Boolean
    Dim bResult As Boolean
    Dim i As Integer
    
    bResult = False
    Select Case TypeName(vEnding)
        Case "String"
            bResult = EndsWith(sText, CStr(vEnding), bCaseInsensitive)
        Case "Variant()"
            For i = LBound(vEnding) To UBound(vEnding)
                If EndsWith(sText, CStr(vEnding(i)), bCaseInsensitive) Then
                    bResult = True
                    Exit For
                End If
            Next
    End Select

    EndsWithSet = bResult
End Function

Function LeftOfLast(sText As Variant, sItem As Variant) As String
    Dim nPos As Integer
    Dim sResult As String
    sResult = sText
    nPos = InStrRev(sText, sItem)
    If nPos > 0 Then
        sResult = left$(sText, nPos - 1)
    End If
    LeftOfLast = sResult
End Function

Function RightOfLast(sText As Variant, sItem As Variant) As String
    Dim nPos As Integer
    Dim sResult As String
    sResult = sText
    nPos = InStrRev(sText, sItem)
    If nPos > 0 Then
        sResult = Mid$(sText, nPos + Len(sItem))
    End If
    RightOfLast = sResult
End Function

Function Reverse(sText As Variant) As String
    Dim sResult As String
    Dim nLen As Integer
    Dim i As Integer
    
    nLen = Len(sText)
    sResult = String(nLen, " ")
    For i = 1 To nLen
        Mid(sResult, i, 1) = Mid(sText, nLen - i + 1, 1)
    Next
    Reverse = sResult
End Function


Function ComprehendCSV(sText As Variant) As Variant
    Dim p As Integer
    Dim c As String * 1
    Dim bInquote As Boolean
    Dim aResult() As Variant
    Dim i As Integer
    Dim sItem As String
    
    Const COMMA As String * 1 = ","
    Const DOUBLE_QUOTE As String * 1 = """"
    
    p = 1
    i = 0
    sItem = ""
    Do
        If p > Len(sText) Then Exit Do
        c = Mid$(sText, p, 1)
        If c = Chr(34) Then
            If bInquote Then 'already in a quoted section
                'check the next char
                'if a dq, it's an embedded DQ DQ combination
                'if a comma, it's the end of the field
                'if it's the end of the line, same as for comma
                If p = Len(sText) Then
                    bInquote = False
                ElseIf Mid$(sText, p + 1, 1) = COMMA Then
                    bInquote = False
                ElseIf Mid$(sText, p + 1, 1) = DOUBLE_QUOTE Then
                    sItem = sItem & c
                    p = p + 1
                Else
                    bInquote = False
                End If
            Else
                bInquote = True
            End If
            GoTo nextChar
        End If
        Debug.Print c;
        If c = COMMA Then
            If bInquote Then
                sItem = sItem & c
            Else
                ReDim Preserve aResult(i)
                aResult(i) = sItem
                i = i + 1
                sItem = ""
            End If
            GoTo nextChar
        End If
        sItem = sItem & c
        
nextChar:
        p = p + 1
    Loop
    If sItem <> vbNullString Then
        ReDim Preserve aResult(i)
        aResult(i) = sItem
        i = i + 1
    End If
    ComprehendCSV = aResult
End Function

Public Function AnyOf(sText As Variant, nOperator As AO_COMPARATORS, sChar As Variant) As Boolean
    Dim i As Integer
    Dim c As Variant
    Dim bResult As Boolean
    bResult = False
    For i = 1 To Len(sText)
        c = Mid$(sText, i, 1)
        Select Case nOperator
            Case AO_LESSTHAN
                bResult = (c < sChar)
            Case AO_LESSTHANEQUAL
                bResult = (c <= sChar)
            Case AO_EQUAL
                bResult = (c = sChar)
            Case AO_GREATERTHANEQUAL
                bResult = (c >= sChar)
            Case AO_GREATERTHAN
                bResult = (c > sChar)
        End Select
        If bResult = True Then
            GoTo done
        End If
    Next
done:
    AnyOf = bResult
End Function

Public Function AnyOfList(sText As Variant, nOperator As AO_COMPARATORS, aChar As Variant) As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim c As Variant
    Dim sChar As String
    Dim bResult As Boolean
    bResult = False
    For i = 1 To Len(sText)
        c = Mid$(sText, i, 1)
        For j = LBound(aChar) To UBound(aChar)
            sChar = aChar(j)
            Select Case nOperator
                Case AO_NOTEQUAL
                    bResult = (c <> sChar)
                Case AO_LESSTHAN
                    bResult = (c < sChar)
                Case AO_LESSTHANEQUAL
                    bResult = (c <= sChar)
                Case AO_EQUAL
                    bResult = (c = sChar)
                Case AO_GREATERTHANEQUAL
                    bResult = (c >= sChar)
                Case AO_GREATERTHAN
                    bResult = (c > sChar)
            End Select
            If bResult = True Then
                GoTo done
            End If
        Next
    Next
done:
    AnyOfList = bResult
End Function

Public Function StringToByteArray(ByVal sString As Variant) As Byte()
    Dim aByte() As Byte
    aByte = sString
    StringToByteArray = aByte
End Function

Public Function ByteArrayToString(aBytes() As Byte) As String
    Dim sString As String
    sString = aBytes
    ByteArrayToString = sString
End Function

Public Function StringConversion(s As Variant, n As Integer) As String
    StringConversion = StrConv(s, n)
End Function

'http://www.freevbcode.com/ShowCode.asp?ID=1512
Public Function URLEncode(StringToEncode As Variant, _
    Optional UsePlusRatherThanHexForSpace As Boolean = False) As Variant

Dim TempAns As String
Dim CurChr As Integer
CurChr = 1
Do Until CurChr - 1 = Len(StringToEncode)
  Select Case Asc(Mid(StringToEncode, CurChr, 1))
    Case 48 To 57, 65 To 90, 97 To 122
      TempAns = TempAns & Mid(StringToEncode, CurChr, 1)
    Case 32
      If UsePlusRatherThanHexForSpace = True Then
        TempAns = TempAns & "+"
      Else
        TempAns = TempAns & "%" & Hex(32)
      End If
   Case Else
         TempAns = TempAns & "%" & _
              Format(Hex(Asc(Mid(StringToEncode, _
              CurChr, 1))), "00")
End Select

  CurChr = CurChr + 1
Loop

URLEncode = TempAns
End Function

'http://www.freevbcode.com/ShowCode.asp?ID=1512
Public Function URLDecode(StringToDecode As Variant) As Variant

Dim TempAns As String
Dim CurChr As Integer

CurChr = 1

Do Until CurChr - 1 = Len(StringToDecode)
  Select Case Mid(StringToDecode, CurChr, 1)
    Case "+"
      TempAns = TempAns & " "
    Case "%"
      TempAns = TempAns & Chr(Val("&h" & _
         Mid(StringToDecode, CurChr + 1, 2)))
       CurChr = CurChr + 2
    Case Else
      TempAns = TempAns & Mid(StringToDecode, CurChr, 1)
  End Select

CurChr = CurChr + 1
Loop

URLDecode = TempAns
End Function


' URLDecode function in Perl for reference
' both VB and Perl versions must return same
'
' sub urldecode{
'  local($val)=@_;
'  $val=~s/\+/ /g;
'  $val=~s/%([0-9A-H]{2})/pack('C',hex($1))/ge;
'  return $val;
' }

'Public Function SplitString(ByVal sText As Variant, Optional ByVal sOpens As Variant = Chr$(34) & "'[{", Optional ByVal sCloses As Variant = Chr$(34) & "']}", Optional ByVal bKeep As Boolean = False) As Variant
Public Function SplitString(ByVal sText As Variant) As Variant
    Dim nHere As Long
    Dim cHere As String * 1
    Dim aRes As Variant
    Dim nRes As Long
    Dim bInquote As Boolean
    Dim sString As String
    ReDim aRes(0)
    nHere = 1
    nRes = 0
    Do
        If nHere > Len(sText) Then Exit Do
        cHere = Mid$(sText, nHere, 1)
        If cHere = Chr$(32) Then
            If bInquote Then
                sString = sString & cHere
                GoTo nextChar
            End If
            If sString <> vbNullString Then
                aRes(nRes) = sString
                sString = vbNullString
                nRes = nRes + 1
                ReDim Preserve aRes(nRes)
            End If
            GoTo nextChar
        ElseIf cHere = Chr$(34) Then
            bInquote = Not bInquote
            GoTo nextChar
        Else
            sString = sString & cHere
        End If
nextChar:
        nHere = nHere + 1
    Loop
    If sString <> vbNullString Then
        aRes(nRes) = sString
    End If
    SplitString = aRes
End Function

