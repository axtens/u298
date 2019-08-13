VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Testing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim sSep As Variant
Dim m As New Misc
Dim Files As New Files
Dim WordList As Variant
Dim Misc As New Misc

Public Function Generate(ByVal sPattern As Variant, Optional ByVal nRepeatCount As Variant = 100) As Variant
    Dim sRes As String
    Dim i As Integer
    sRes = ""
    If left$(sPattern, 1) = "?" Or left$(sPattern, 1) = "!" Then
        For i = 1 To nRepeatCount
            sRes = sRes & NewGenerate(sPattern)
        Next
    Else
        'For i = 1 To nRepeatCount
            sRes = sRes & OldGenerate(sPattern, nRepeatCount)
        'Next
    End If
    Generate = sRes
End Function

Private Function NewGenerate(ByVal sPattern As Variant) As Variant
    Dim here As Integer
    Dim c As String * 1
    Dim s As String
    
    Dim nRepCnt As Variant
    Dim sVarTyp As Variant
    Dim sOp As Variant
    Dim sParam As Variant
    Dim aRange As Variant
    Dim sLo As Variant
    Dim sHi As Variant
    
    Dim v As Variant
    
    Dim bRanged As Boolean
    
    Dim rgx As New RegExp
    Dim sbm As SubMatches
    Dim mtc As MatchCollection
    Dim mat As Match
    
    Dim sRes As String
    Dim sFinalRes As String
    
    sFinalRes = vbNullString
    Do
        rgx.MultiLine = False
        rgx.IgnoreCase = True
        rgx.Global = True
        rgx.Pattern = "^!(\d+)(\S)(\S)(.+?);|\?d=([-]*\d+?);"
        
        Set mtc = rgx.Execute(sPattern)
        If mtc.Count = 0 Then
            Exit Do
        End If
        
        Set mat = mtc(0)
        
        'Debug.Print Asc(sSep)
        Debug.Print mat.Value
        
        If left$(mat.Value, 1) = "?" Then
            Set sbm = mat.SubMatches
            v = sbm.Item(4)
            If v = -1 Then
                sSep = vbNullString
            Else
                sSep = Chr$(v)
            End If
            sPattern = Mid(sPattern, Len(mat.Value) + 1)
            GoTo EndLoop
        End If
            
        Set sbm = mat.SubMatches
        nRepCnt = sbm.Item(0)
        sVarTyp = sbm.Item(1)
        sOp = sbm.Item(2)
        sParam = sbm.Item(3)
        bRanged = False
        
        If InStr(sParam, "..") Then
            aRange = Split(sParam, "..")
            sLo = aRange(0)
            sHi = aRange(1)
            bRanged = True
        End If
        
        Dim i As Integer
        sRes = vbNullString
        
        If nRepCnt = "0" Then nRepCnt = Int(Rnd * 100 + 1)
        
        For i = 1 To nRepCnt
        
            Select Case sVarTyp
            Case "i" 'integer
                Select Case sOp
                    Case "<"
                        If bRanged Then Form1.Die "Range not allowed on '<'"
                        sRes = sRes & RandomIntegerLess(sParam)
                    Case ">"
                        If bRanged Then Form1.Die "Range not allowed on '>'"
                        sRes = sRes & RandomIntegerGreater(sParam)
                    Case "="
                        If bRanged Then Form1.Die "Range not allowed on '='"
                        sRes = sRes & sParam
                    Case "@"
                        If bRanged = False Then Form1.Die "No range for ranged integer"
                        sRes = sRes & RandomIntegerRange(sLo, sHi)
                End Select
            Case "I" 'big integer
                Select Case sOp
                    Case "<"
                        If bRanged Then Form1.Die "Range not allowed on '<'"
                        sRes = sRes & m.RandomBigInt(sParam - 1)
                    Case ">"
                        If bRanged Then Form1.Die "Range not allowed on '>'"
                        sRes = sRes & m.RandomBigInt(RandomIntegerGreater(sParam))
                    Case "=" 'in this case, the number of digits
                        If bRanged Then Form1.Die "Range not allowed on '='"
                        sRes = sRes & m.RandomBigInt(sParam)
                    Case "@"
                        If bRanged = False Then Form1.Die "No range for ranged big integer"
                        sRes = sRes & m.RandomBigInt(RandomIntegerRange(sLo, sHi))
                End Select
            Case "f" 'float
                Select Case sOp
                    Case "<"
                        If bRanged Then Form1.Die "Range not allowed on '<'"
                        sRes = sRes & RandomFloatLess(sParam)
                    Case ">"
                        If bRanged Then Form1.Die "Range not allowed on '>'"
                        sRes = sRes & RandomFloatGreater(sParam)
                    Case "="
                        If bRanged Then Form1.Die "Range not allowed on '='"
                        sRes = sRes & sParam
                    Case "@"
                        If bRanged = False Then Form1.Die "No range for ranged float"
                        sRes = sRes & RandomFloatRange(sLo, sHi)
                End Select
            Case "F" 'big float
            Case "c" 'complex
            Case "x" 'character
                Select Case sOp
                    Case "="
                        If bRanged Then Form1.Die "Range not allowed on '='"
                        sRes = sRes & sParam
                    Case "@"
                        If bRanged = False Then Form1.Die "No range for ranged character"
                        sRes = sRes & RandomOfCharacters(sParam)
                End Select
            Case "m" 'matrx
            Case "s" 'string
            Case "d" 'date
            Case "D" 'excel date
            Case "t" 'time
            Case "w" 'word
                Select Case sOp
                    Case "<"
                        If bRanged Then Form1.Die "Range not allowed on '<'"
                        sRes = sRes & RandomWordLess(sParam)
                    Case ">"
                        If bRanged Then Form1.Die "Range not allowed on '>'"
                        sRes = sRes & RandomWordGreater(sParam)
                    Case "="
                        If bRanged Then Form1.Die "Range not allowed on '='"
                        sRes = sRes & sParam
                    Case "@"
                        If bRanged = False Then Form1.Die "No range for ranged float"
                        sRes = sRes & RandomWordRange(sLo, sHi)
                End Select
            End Select
            
            If i < nRepCnt Then
                sRes = sRes & sSep
            Else
                sSep = Chr$(9)
            End If
        Next
        sPattern = Mid(sPattern, Len(mat.Value) + 1)
        If Len(sPattern) > 1 Then
            sFinalRes = sFinalRes & sRes & sSep
        Else
            sFinalRes = sFinalRes & sRes
        End If
EndLoop:
        
    Loop
    NewGenerate = sFinalRes & vbNewLine
End Function

Private Function RandomIntegerLess(vLim As Variant) As Variant
    Dim sRes As Variant
    Dim i As Long
    Do
        i = Int(Rnd * 32767) * IIF(Misc.toss() = True, -1, 1)
        If i < vLim Then
            sRes = i
            Exit Do
        End If
    Loop
    RandomIntegerLess = sRes
End Function

Private Function RandomIntegerGreater(vLim As Variant) As Variant
    Dim sRes As Variant
    Dim i As Long
    Do
        i = Int(Rnd * 32767) * IIF(Misc.toss() = True, -1, 1)
        If i > vLim Then
            sRes = i
            Exit Do
        End If
    Loop
    RandomIntegerGreater = sRes
End Function

Private Function RandomIntegerRange(vLo As Variant, vHi As Variant) As Variant
    Dim sRes As Variant
    Dim i As Long
    Do
        i = Int(Rnd * 32767) * IIF(Misc.toss() = True, -1, 1)
        If i >= vLo And i <= vHi Then
            sRes = i
            Exit Do
        End If
    Loop
    RandomIntegerRange = sRes
End Function

Private Function RandomFloatLess(ByVal vLim As Variant) As Variant
    Dim sRes As Variant
    Dim i As Double
    Dim n As Double
    If vLim = 0 Then
        n = 100
    Else
        n = vLim
    End If
    Do
        i = Rnd * n * IIF(Misc.toss() = True, -1, 1) '* 3.402823E+38
        If i < vLim Then
            sRes = i
            Exit Do
        End If
    Loop
    RandomFloatLess = sRes
End Function

Private Function RandomFloatGreater(ByVal vLim As Variant) As Variant
    Dim sRes As Variant
    Dim i As Double
    Dim n As Double
    If vLim = 0 Then
        n = 100
    Else
        n = vLim
    End If
    Do
        i = Rnd * n * IIF(Misc.toss() = True, -1, 1) '* 3.402823E+38
        If i > vLim Then
            sRes = i
            Exit Do
        End If
    Loop
    RandomFloatGreater = sRes
End Function

Private Function RandomFloatRange(vLo As Variant, vHi As Variant) As Variant
    Dim sRes As Variant
    Dim i As Double
    Do
        i = Rnd * vHi * IIF(Misc.toss() = True, -1, 1) '* 3.402823E+38
        If i >= vLo And i <= vHi Then
            sRes = i
            Exit Do
        End If
    Loop
    RandomFloatRange = sRes
End Function

Private Function RandomWordLess(ByVal vLim As Variant) As Variant
    Dim sRes As Variant
    Dim i As Long
    If vLim < 1 Then vLim = 2
    Do
    i = Rnd * UBound(WordList)
    If i > UBound(WordList) Then i = UBound(WordList)
    sRes = WordList(i)
    If Len(sRes) < vLim Then
        Exit Do
    End If
    Loop
    RandomWordLess = sRes
End Function

Private Function RandomWordGreater(ByVal vLim As Variant) As Variant
   Dim sRes As Variant
    Dim i As Long
    If vLim < 1 Then vLim = 2
    Do
    i = Rnd * UBound(WordList)
    If i > UBound(WordList) Then i = UBound(WordList)
    sRes = WordList(i)
    If Len(sRes) > CLng(vLim) Then
        Exit Do
    End If
    Loop
    RandomWordGreater = sRes
End Function

Private Function RandomWordRange(vLo As Variant, vHi As Variant) As Variant
   Dim sRes As Variant
    Dim i As Long
    If vLo < 0 Then vLo = 0
    If vHi > UBound(WordList) Then vHi = UBound(WordList)
    Do
    i = Rnd * UBound(WordList)
    If i > UBound(WordList) Then i = UBound(WordList)
    sRes = WordList(i)
    If Len(sRes) >= vLo And Len(sRes) >= vHi Then
        Exit Do
    End If
    Loop
    RandomWordRange = sRes
End Function

Private Function RandomOfCharacters(s As Variant) As Variant
    Dim l As Long
    l = Len(s)
    RandomOfCharacters = Mid$(s, Int(Rnd * l + 1), 1)
End Function

Sub LoadWords(sFilename As Variant)
    WordList = Files.ReadFileEToArray(sFilename)
End Sub

Private Function OldGenerate(ByVal sPattern As Variant, ByVal nRepeatCount As Variant) As Variant
    Dim sResult As Variant
    sResult = ""
    
    Dim aPatterns As Variant
    aPatterns = Split(sPattern, ":")
    
    Dim aParts As Variant
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim p As Variant
    Dim vTemp As Variant
    Dim r As Variant
    
    For j = 1 To nRepeatCount
        For i = 0 To UBound(aPatterns)
            aParts = Split(aPatterns(i), "<")
            If UBound(aParts) > 0 Then
                p = aParts(1)
            Else
                p = 1000
            End If
            
            Select Case LCase$(aParts(0))

            Case "c"
                sResult = sResult & m.RandomComplex(p)
            Case "cs"
                sResult = sResult & m.RandomComplexSpaced(p)
            Case "i"
                sResult = sResult & m.RandomInt(p)
            Case "f", "r"
                sResult = sResult & m.RandomReal(p)
            Case "+f", "+r"
                sResult = sResult & Abs(m.RandomReal(p))
            Case "*f", "*r" ' greater than zero
                Do
                     r = Abs(m.RandomReal(p))
                     If r > 0 Then Exit Do
                Loop
                sResult = sResult & r
            Case "li", "bi"
                sResult = sResult & m.RandomBigInt(p)
            Case "lf", "bf", "br", "lr"
                sResult = sResult & m.RandomBigReal(p)
            Case "+i"
                k = m.RandomInt(p)
                sResult = sResult & Abs(k) 'IIF(k < 0, k * -1, k)
            Case "*i" ' greater than zero
                Do
                    k = m.RandomInt(p)
                    If k > 0 Then Exit Do
                Loop
                sResult = sResult & Abs(k) 'IIF(k < 0, k * -1, k)
            Case "s" ' random string, of length in p
                For k = 1 To p
                    sResult = sResult & RandomChar()
                Next
            Case "x" ' random character, from set in p
                sResult = sResult & Mid$(p, Int(Rnd * Len(p)) + 1, 1)
            Case "d" ' random date
                sResult = sResult & CDate(RInt(2 ^ 21))
            Case "dx" 'random excel date
                On Error Resume Next
                vTemp = CVDate(ExcelDate())
                If Err.Number <> 0 Then
                    vTemp = Date
                End If
                sResult = sResult & vTemp
                On Error GoTo 0
                
            Case "t" ' random time
                sResult = sResult & CStr(CDate(Rnd))
            Case Else
                sResult = sResult & aPatterns(i)
            End Select
            If i < UBound(aPatterns) Then
                sResult = sResult & sSep
            End If
        Next
        sResult = sResult & vbNewLine
    Next
    OldGenerate = sResult
End Function

Private Function RandomChar() As String
    Const CHARSET = "abcdefghijklmnopqrstuvwxyz ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim i As Integer
    RandomChar = Mid$(CHARSET, Int(Rnd * Len(CHARSET)) + 1, 1)
End Function


Public Function Separator(ByVal sSeparator As Variant) As Variant
    Dim sPrev As Variant
    sPrev = sSep
    sSep = sSeparator
    Separator = sPrev
End Function

Private Function RInt(nLimit As Long)
    RInt = Int(Rnd * nLimit) + 1
End Function

Private Function ExcelDate() As String
    Dim y1 As Variant
    Dim m1 As Variant
    Dim D1 As Variant
    
    Do
        y1 = RInt(9999)
        If CInt(y1) >= 1900 Then Exit Do
    Loop
    
    ExcelDate = y1 & "-" & CStr(RInt(12)) & "-" & CStr(RInt(31))
    
End Function

Private Sub Class_Initialize()
    Randomize Timer
    sSep = vbTab
End Sub
