Attribute VB_Name = "Complex"
Option Explicit

Public MaxDouble As Double
Public MinDouble As Double
Public NegInfinity As Double
Public PosInfinity As Double
Public QuietNAN As Double

Public PIE As Double

Private Declare Sub CopyMemory Lib "kernel32" _
    Alias "RtlMoveMemory" (Destination As Any, _
    source As Any, ByVal Length As Long)
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function HeapAlloc Lib "kernel32" _
    (ByVal hHeap As Long, ByVal dwFlags As Long, _
     ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" _
    (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Private Declare Sub CopyMemoryWrite Lib "kernel32" Alias _
    "RtlMoveMemory" (ByVal Destination As Long, _
    source As Any, ByVal Length As Long)
Private Declare Sub CopyMemoryRead Lib "kernel32" Alias _
    "RtlMoveMemory" (Destination As Any, _
    ByVal source As Long, ByVal Length As Long)

Private myRegExp As New RegExp
Private myMatches As MatchCollection
Private myMatch As Match

Private mvarInfOption As Boolean 'local copy

Public Property Let InfOption(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.InfOption = 5
    mvarInfOption = vData
End Property

Public Property Get InfOption() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.InfOption
    InfOption = mvarInfOption
End Property

Public Sub Init()
    Dim ptrToDouble As Long
    Dim byteArray(7) As Byte
    Dim i As Integer
    
    byteArray(7) = &H7F 'maxdbl
    byteArray(6) = &HEF
    For i = 0 To 5
        byteArray(i) = &HFF
    Next
    ptrToDouble = VarPtr(MaxDouble)
    CopyMemoryWrite ptrToDouble, byteArray(0), 8

    byteArray(7) = &HFF 'mindbl
    byteArray(6) = &HEF
    For i = 0 To 5
        byteArray(i) = &HFF
    Next
    ptrToDouble = VarPtr(MinDouble)
    CopyMemoryWrite ptrToDouble, byteArray(0), 8
    
    byteArray(7) = &H7F '+inf
    byteArray(6) = &HF0
    For i = 0 To 5
        byteArray(i) = 0
    Next
    ptrToDouble = VarPtr(PosInfinity)
    CopyMemoryWrite ptrToDouble, byteArray(0), 8

    byteArray(7) = &HFF '-inf
    byteArray(6) = &HF0
    For i = 0 To 5
        byteArray(i) = 0
    Next
    ptrToDouble = VarPtr(NegInfinity)
    CopyMemoryWrite ptrToDouble, byteArray(0), 8

    byteArray(7) = &H7F 'quiet NAN
    byteArray(6) = &HF0
    For i = 0 To 5
        byteArray(i) = &HFF
    Next
    ptrToDouble = VarPtr(QuietNAN)
    CopyMemoryWrite ptrToDouble, byteArray(0), 8

    PIE = Atn(1#) * 4#
End Sub

Public Function toString(ByVal c As ComplexNumber) As String
    Dim sSep As String
    If c.Imag <> 0 Then
        If c.Truncation > -1 Then
            c.Imag = Round(c.Imag, c.Truncation)
            c.Real = Round(c.Real, c.Truncation)
        End If
        Select Case Sgn(c.Imag)
        Case -1
            sSep = "-"
        Case 0, 1
            sSep = "+"
        End Select
        sSep = IIF(c.SpaceBefore, " ", vbNullString) & sSep & IIF(c.SpaceAfter, " ", vbNullString)
        toString = CStr(c.Real) & sSep & IIF(c.TrailingImaginary, vbNullString, IIF(c.JayFormat, "j", "i")) & CStr(Abs(c.Imag)) & IIF(c.TrailingImaginary, IIF(c.JayFormat, "j", "i"), vbNullString)
    Else
        If c.Truncation > -1 Then
            c.Real = Round(c.Real, c.Truncation)
        End If
        toString = CStr(c.Real)
    End If
End Function

Public Function SetReal(c As ComplexNumber, r As Double) As ComplexNumber
    c.Real = r
    Set SetReal = c
End Function

Public Function SetImag(c As ComplexNumber, i As Double) As ComplexNumber
    c.Imag = i
    Set SetImag = c
End Function

Public Function Reset(r As Double, i As Double) As ComplexNumber
    Dim c As New ComplexNumber
    c.Real = r
    c.Imag = i
    Set Reset = c
End Function

Public Function Parse(n As Variant) As ComplexNumber
    Dim c As New ComplexNumber
    Dim v1, v2
    If right$(Trim$(n), 1) = "i" Or right$(Trim$(n), 1) = "j" Then
        myRegExp.IgnoreCase = True
        myRegExp.Global = True
        myRegExp.Pattern = "[+\-]?(?:0|[1-9]\d*)\s*?(?:\.\d*)?(?:[eE][+\-]?\d+)?"
        Set myMatches = myRegExp.Execute(n)
        Select Case myMatches.Count
        Case 2
            v1 = myMatches.Item(0).Value
            If left$(v1, 1) = "+" Then v1 = Mid$(v1, 2)
            v2 = myMatches.Item(1).Value
            If left$(v2, 1) = "+" Then v2 = Mid$(v2, 2)
            c.Real = CDbl(v1)
            c.Imag = CDbl(v2)
        Case 1
            v1 = myMatches.Item(0).Value
            If left$(v1, 1) = "+" Then v1 = Mid$(v1, 2)
            c.Real = CDbl(v1)
            c.Imag = 0
        Case Else
            c.Real = 0
            c.Imag = 0
        End Select
    Else
        c.Real = n
        c.Imag = 0
    End If
    If right$(Trim$(n), 1) = "j" Then c.JayFormat = True
    Set Parse = c
End Function

Public Function PolarRad(modl As Variant, arg As Variant) As ComplexNumber
    Dim c As New ComplexNumber
    c.Real = modl * Cos(arg)
    c.Imag = modl * Sin(arg)
    Set PolarRad = c
End Function

Public Function PolarDeg(modl As Variant, ByVal arg As Variant) As ComplexNumber
    Dim c As New ComplexNumber
    arg = (PIE / 180) * arg
    c.Real = modl * Cos(arg)
    c.Imag = modl * Sin(arg)
    Set PolarDeg = c
End Function

Public Function Truncate(c As ComplexNumber, prec As Long) As ComplexNumber
    Dim c0 As New ComplexNumber
    Dim a As Double
    Dim b As Double
    
    If prec > 0 Then
        a = c.Real
        b = c.Imag
        a = Round(a, prec)
        b = Round(b, prec)
        c0.Real = a
        c0.Imag = b
    Else
        c0.Real = c.Real
        c0.Imag = c.Imag
    End If
    
    Set Truncate = c0
End Function

Public Function Plus(c1 As ComplexNumber, c2 As ComplexNumber) As ComplexNumber
    Dim c As New ComplexNumber
    c.Real = c1.Real + c2.Real
    c.Imag = c1.Imag + c2.Imag
    Set Plus = c
End Function

Public Function Zero() As ComplexNumber
    Set Zero = Reset(0, 0)
End Function

Public Function PlusJay() As ComplexNumber
    Set PlusJay = Reset(0, 1)
End Function

Public Function PlusOne() As ComplexNumber
    Set PlusOne = Reset(1, 0)
End Function

Public Function MinusJay() As ComplexNumber
    Set PlusJay = Reset(0, -1)
End Function

Public Function MinusOne() As ComplexNumber
    Set PlusOne = Reset(-1, 0)
End Function

Public Function ComplexPI() As ComplexNumber
    Set ComplexPI = Reset(PIE, 0)
End Function

Public Function TwoPiJay() As ComplexNumber
    Set TwoPiJay = Reset(0, 2 * PIE)
End Function

Public Function PlusInfinity() As ComplexNumber
    Set PlusInfinity = Reset(PosInfinity, PosInfinity)
End Function

Public Function MinusInfinity() As ComplexNumber
    Set MinusInfinity = Reset(NegInfinity, NegInfinity)
End Function

Public Function CopyOf(c As ComplexNumber) As ComplexNumber
    Dim cNew As New ComplexNumber
    cNew.Real = c.Real
    cNew.Imag = c.Imag
    cNew.JayFormat = c.JayFormat
    cNew.SpaceAfter = c.SpaceAfter
    cNew.SpaceBefore = c.SpaceBefore
    cNew.TrailingImaginary = c.TrailingImaginary
    cNew.Truncation = c.Truncation
    Set CopyOf = cNew
End Function

Public Function PlusDouble(c As ComplexNumber, d As Double) As ComplexNumber
    Dim c0 As New ComplexNumber
    Set c0 = Complex.CopyOf(c)
    c0.Real = c.Real + d
    Set PlusDouble = c0
End Function

Public Function DoubleDoublePlus(D1 As Double, D2 As Double) As ComplexNumber
    Set DoubleDoublePlus = Reset(D1 + D2, 0)
End Function

Public Function Minus(c1 As ComplexNumber, c2 As ComplexNumber) As ComplexNumber
    Dim c As New ComplexNumber
    Set c = CopyOf(c1)
    c.Real = c1.Real - c2.Real
    c.Imag = c1.Imag - c2.Imag
    Set Minus = c
End Function

Public Function MinusDouble(c As ComplexNumber, d As Double) As ComplexNumber
    Dim c0 As New ComplexNumber
    Set c0 = Complex.CopyOf(c)
    c0.Real = c.Real - d
    Set MinusDouble = c0
End Function

Public Function DoubleMinus(d As Double, c As ComplexNumber) As ComplexNumber
    Dim c0 As ComplexNumber
    Set c0 = CopyOf(c)
    c0.Real = d - c0.Real
    c0.Imag = -c0.Imag
    Set DoubleMinus = c0
End Function

Public Function DoubleDoubleMinus(D1 As Double, D2 As Double) As ComplexNumber
    Set DoubleDoubleMinus = Reset(D1 - D2, 0)
End Function

Public Function Times(c1 As ComplexNumber, c2 As ComplexNumber) As ComplexNumber
    Dim dTemp1 As Double
    Dim dTemp2 As Double
    
    If Complex.InfOption = True Then
        If IsInfinity(c1) And Not IsZero(c2) Then
            Set Times = Reset(PosInfinity, PosInfinity)
            Exit Function
        End If
        
        If IsInfinity(c2) And Not IsZero(c1) Then
            Set Times = Reset(PosInfinity, PosInfinity)
            Exit Function
        End If
    End If
    
    Set Times = Reset(c1.Real * c2.Real - c1.Imag * c2.Imag, c1.Real * c2.Imag + c1.Imag * c2.Real)

End Function

Public Function TimesDouble(c As ComplexNumber, d As Double) As ComplexNumber
    
    If Complex.InfOption = True Then
        If IsInfinity(c) And Not IsZero(d) Then
            Set TimesDouble = Reset(PosInfinity, PosInfinity)
            Exit Function
        End If
        
        If IsInfinity(d) And Not IsZero(c) Then
            Set TimesDouble = Reset(PosInfinity, PosInfinity)
            Exit Function
        End If
    End If
    Set TimesDouble = Reset(c.Real * d, c.Imag * d)

End Function

Public Function DoubleTimes(d As Double, c As ComplexNumber) As ComplexNumber
    
    If Complex.InfOption = True Then
        If IsInfinity(c) And Not IsZero(d) Then
            Set DoubleTimes = Reset(PosInfinity, PosInfinity)
            Exit Function
        End If
        
        If IsInfinity(d) And Not IsZero(c) Then
            Set DoubleTimes = Reset(PosInfinity, PosInfinity)
            Exit Function
        End If
    End If
    Set DoubleTimes = Reset(d * c.Real, d * c.Imag)

End Function

Public Function DoubleDoubleTimes(D1 As Double, D2 As Double) As ComplexNumber
    Set DoubleDoubleTimes = Reset(D1 * D2, 0)
End Function

Public Function Divide(c1 As ComplexNumber, c2 As ComplexNumber) As ComplexNumber
    Dim denom As Double
    Dim ratio As Double
    
    If InfOption = True Then
        If Not IsInfinity(c1) And IsInfinity(c2) Then
            Set Divide = Reset(0, 0)
            Exit Function
        End If
    End If
    
    If Abs(c2.Real) >= Abs(c2.Imag) Then
        ratio = c2.Imag / c2.Real
        denom = c2.Real + c2.Imag * ratio
        Set Divide = Reset((c1.Real + c1.Imag * ratio) / denom, (c1.Imag - c1.Real * ratio) / denom)
    Else
        ratio = c2.Real / c2.Imag
        denom = c2.Real * ratio + c2.Imag
        Set Divide = Reset((c1.Real * ratio + c1.Imag) / denom, (c1.Imag * ratio - c1.Real) / denom)
    End If
    
End Function

Public Function Divide2(c1 As ComplexNumber, c2 As ComplexNumber) As ComplexNumber
    If InfOption = True Then
        If Not IsInfinity(c1) And IsInfinity(c2) Then
            Set Divide2 = Reset(0, 0)
            Exit Function
        End If
    End If
    Set Divide2 = Times(c1, Conjugate(c2))
End Function

''
Public Function DivideDouble(c As ComplexNumber, d As Double) As ComplexNumber
    Set DivideDouble = Reset(c.Real / d, c.Imag / d)
End Function
'
Public Function DoubleDivide(d As Double, c As ComplexNumber) As ComplexNumber
    Dim a, b
    Dim denom, ratio
    
    If InfOption = True And IsInfinity(c) Then
        Set DoubleDivide = Reset(0, 0)
        Exit Function
    End If

    If Abs(c.Real) >= Abs(c.Imag) Then
        ratio = c.Imag / c.Real
        denom = c.Real + c.Imag * ratio
        Set DoubleDivide = Reset(d / denom, -d * ratio / denom)
    Else
        ratio = c.Real / c.Imag
        denom = c.Real * ratio + c.Imag
        Set DoubleDivide = Reset(d * ratio / denom, -d / denom)
    End If
End Function
'
Public Function Reciprocal(c As ComplexNumber) As ComplexNumber
        Set Reciprocal = DoubleDivide(1, c)
End Function

'
Public Function Negate(c As ComplexNumber) As ComplexNumber
    Set Negate = Reset(-c.Real, -c.Imag)
End Function
'
Public Function Conjugate(c As ComplexNumber) As ComplexNumber
    Set Conjugate = Reset(c.Real, -c.Imag)
End Function
'
''Logicals
Public Function IsReal(c As ComplexNumber) As Boolean
    IsReal = (c.Imag = 0)
End Function
'
Public Function IsZero(c As Variant) As Boolean
    Dim bRes As Boolean
    bRes = False
    If TypeName(c) = "ComplexNumber" Then
        If Abs(c.Real) = 0 And Abs(c.Imag) = 0 Then bRes = True
    Else
        bRes = (Abs(c) = 0)
    End If
    IsZero = bRes
End Function
'
Public Function IsInfinity(c As Variant) As Boolean
    Dim bRes As Boolean
    bRes = False
    If TypeName(c) = "ComplexNumber" Then
        If c.Real = PosInfinity And c.Imag = PosInfinity Then bRes = True
        If c.Real = NegInfinity And c.Imag = NegInfinity Then bRes = True
    Else
        If c = PosInfinity Then bRes = True
        If c = NegInfinity Then bRes = True
    End If
    IsInfinity = bRes
End Function

Public Function IsPlusInfinity(c As Variant) As Boolean
    Dim bRes As Boolean
    bRes = False
    If TypeName(c) = "ComplexNumber" Then
        If c.Real = PosInfinity And c.Imag = PosInfinity Then bRes = True
    Else
        If c = PosInfinity Then bRes = True
    End If
    IsPlusInfinity = bRes
End Function

Public Function IsMinusInfinity(c As Variant) As Boolean
    Dim bRes As Boolean
    bRes = False
    If TypeName(c) = "ComplexNumber" Then
        If c.Real = NegInfinity And c.Imag = NegInfinity Then bRes = True
    Else
        If c = NegInfinity Then bRes = True
    End If
    IsMinusInfinity = bRes
End Function

'
Public Function IsNaN(c As Variant) As Boolean
    Dim bRes As Boolean
    bRes = False
    If TypeName(c) = "ComplexNumber" Then
        If c.Real = QuietNAN And c.Imag = QuietNAN Then bRes = True
    Else
        If c = QuietNAN Then bRes = True
    End If
    IsNaN = bRes
End Function
'
Public Function Equals(c1 As ComplexNumber, c2 As ComplexNumber) As Boolean
    Dim bRes As Boolean
    bRes = False
    If c1.Real = c2.Real And c1.Imag = c2.Imag Then bRes = True
    Equals = bRes
End Function
'
Public Function NotEquals(c1 As ComplexNumber, c2 As ComplexNumber) As Boolean
    Dim bRes As Boolean
    bRes = False
    If c1.Real <> c2.Real And c1.Imag <> c2.Imag Then bRes = True
    NotEquals = bRes
End Function
'
Public Function GreaterThan(c1 As ComplexNumber, c2 As ComplexNumber) As Boolean
    Dim bRes As Boolean
    bRes = False
    If c1.Real > c2.Real Then
        bRes = True
    Else
        If c1.Real = c2.Real Then
            If c1.Imag > c2.Imag Then
                bRes = True
            End If
        End If
    End If
    GreaterThan = bRes
End Function
'
Public Function LessThan(c1 As ComplexNumber, c2 As ComplexNumber) As Boolean
    Dim bRes As Boolean
    bRes = False
    If c1.Real < c2.Real Then
        bRes = True
    Else
        If c1.Real = c2.Real Then
            If c1.Imag < c2.Imag Then
                bRes = True
            End If
        End If
    End If
    LessThan = bRes
End Function
'
Public Function Maximum(c1 As ComplexNumber, c2 As ComplexNumber) As ComplexNumber
    If LessThan(c1, c2) Then
        Set Maximum = CopyOf(c2)
    Else
        Set Maximum = CopyOf(c1)
    End If
End Function
'
Public Function Minimum(c1 As ComplexNumber, c2 As ComplexNumber) As ComplexNumber
    If GreaterThan(c1, c2) Then
        Set Minimum = CopyOf(c2)
    Else
        Set Minimum = CopyOf(c1)
    End If
End Function
'
Public Function Sign(c As ComplexNumber) As Integer
'determines the sign of the number or complex value x. If x is a complex value, the
'result is determined as follows:
'   1, if real(x) > 0 or real(x) = 0 and imag(x) > 0
'   -1, if real(x) < 0 or real(x) = 0 and imag(x) < 0
'   0 otherwise.
    Dim iRes As Integer
    iRes = 0
    If c.Real > 0 Then
        iRes = 1
    ElseIf c.Real < 0 Then
        iRes = -1
    Else
        If c.Real = 0 Then
            If c.Imag > 0 Then
                iRes = 1
            ElseIf c.Imag < 0 Then
                iRes = -1
            End If
        End If
    End If
    Sign = iRes
End Function
'
''mathematical routines
'
Public Function Absolute(c As ComplexNumber) As Double
    Dim rmod As Double
    Dim imod As Double
    Dim ratio As Double
    Dim res As Double
    rmod = Abs(c.Real)
    imod = Abs(c.Imag)
    ratio = 0
    res = 0
    If rmod = 0 Then
        res = imod
    Else
        If imod = 0 Then
            res = rmod
        End If
        If rmod >= imod Then
            ratio = c.Imag / c.Real
            res = rmod * Sqr(1 + ratio * ratio)
        Else
            ratio = c.Real / c.Imag
            res = imod * Sqr(1 + ratio * ratio)
        End If
    End If
    Absolute = res
End Function
'
Public Function SquareAbsolute(c As ComplexNumber) As Double
        SquareAbsolute = c.Real * c.Real + c.Imag * c.Imag
End Function
'
Public Function Argument(c As ComplexNumber) As Double
        Argument = Atan2(c.Imag, c.Real)
End Function
'
Public Function ArgumentDegrees(c As ComplexNumber) As Double
        ArgumentDegrees = Argument(c) * (180 / PIE)
End Function
'
Public Function Exponent(c As ComplexNumber) As ComplexNumber
    Dim LTemp1 As Double
    Dim Ltemp2 As Double
    LTemp1 = Exp(c.Real)
    Ltemp2 = c.Imag
    If Ltemp2 = 0 Then
        Set Exponent = Reset(LTemp1, 0)
    Else
        If c.Real = 0 Then
            Set Exponent = Reset(Cos(Ltemp2), Sin(Ltemp2))
        Else
            Set Exponent = Reset(LTemp1 * Cos(Ltemp2), LTemp1 * Sin(Ltemp2))
        End If
    End If
End Function
'
Public Function Logarithm(c As ComplexNumber) As ComplexNumber
    Set Logarithm = Reset(Log(Absolute(c)), Atan2(c.Imag, c.Real))
End Function
'
'Public Function Logarithm10(c As ComplexNumber) As ComplexNumber
'    Set Logarithm10 = Reset(Log(Absolute(c)) / Log(10), Atan2(c.Imag, c.Real))
'End Function
''
'Public Function LogarithmN(c As ComplexNumber, nBase As Integer) As ComplexNumber
'    Set LogarithmN = Reset(Log(Absolute(c)) / Log(nBase), Atan2(c.Imag, c.Real))
'End Function
'
'
Public Function SquareRoot(c As ComplexNumber) As ComplexNumber
'    Dim a As Double, b As Double, W As Double, ratio As Double, amod As Double, bmod As Double
'    a = c.Real
'    b = c.Imag
'    If b = 0# Then
'        If a >= 0# Then
'            Set SquareRoot = Reset(Sqr(a), 0)
'        Else
'            Set SquareRoot = Reset(0, Sqr(-a))
'        End If
'    Else
'        amod = Abs(a)
'        bmod = Abs(b)
'        If amod >= bmod Then
'            ratio = b / a
'            W = Sqr(amod) * Sqr(0.5 * (1# + Sqr(1# + ratio * ratio)))
'        Else
'            ratio = a / b
'            W = Sqr(bmod) * Sqr(0.5 * (Abs(ratio) + Sqr(1# + ratio * ratio)))
'        End If
'        If a >= 0 Then
'            Set SquareRoot = Reset(W, b / (2# * W))
'        Else
'            If b >= 0# Then
'                Set SquareRoot = Reset(W, b / (2# * W))
'            Else
'                Set SquareRoot = Reset(-W, b / (2# * -W))
'            End If
'        End If
'    End If
    Set SquareRoot = NthRoot(2, c)
End Function
'
Public Function NthRoot(n As Double, c As ComplexNumber) As ComplexNumber
    If n = 0 Then
        Set NthRoot = Reset(PosInfinity, 0)
    Else
        If n = 1 Then
            Set NthRoot = CopyOf(c)
        Else
            Set NthRoot = Exponent(DivideDouble(Logarithm(c), n))
        End If
    End If
End Function
'
Public Function Square(c As ComplexNumber) As ComplexNumber
    Dim a As Double, b As Double
    a = c.Real * c.Real - c.Imag * c.Imag
    b = 2# * c.Real * c.Imag
    Set Square = Reset(a, b)
End Function

'
Public Function Power(c1 As ComplexNumber, c2 As ComplexNumber)
    If IsZero(c1) Then
        If c2.Imag = 0 Then
            If c2.Real = 0 Then
                Set Power = Reset(1#, 0#)
            Else
                If c2.Real > 0# Then
                    Set Power = Reset(0#, 0#)
                Else
                    If c2.Real < 0# Then
                        Set Power = Reset(PosInfinity, 0#)
                    End If
                End If
            End If
        Else
            Set Power = Exponent(Times(c2, Logarithm(c1)))
        End If
    Else
        Set Power = Exponent(Times(c2, Logarithm(c1)))
    End If
End Function
'
''trig
Public Function Sine(c As ComplexNumber) As ComplexNumber
    Set Sine = Reset(Sin(c.Real) * CosH(c.Imag), Cos(c.Real) * SinH(c.Imag))
End Function
'
Public Function Cosine(c As ComplexNumber) As ComplexNumber
    Set Cosine = Reset(Cos(c.Real) * CosH(c.Imag), -Sin(c.Real) * SinH(c.Imag))
End Function
'
Public Function Tangent(c As ComplexNumber) As ComplexNumber
    Dim a As Double
    Dim b As Double
    Dim X As ComplexNumber
    Dim Y As ComplexNumber
    a = c.Real
    b = c.Imag
    Set X = Reset(Sin(a) * CosH(b), Cos(a) * SinH(b))
    Set Y = Reset(Cos(a) * CosH(b), -Sin(a) * SinH(b))
    Set Tangent = Divide(X, Y)
End Function

Public Function Secant(c As ComplexNumber) As ComplexNumber
    Dim a As Double
    Dim b As Double
    a = c.Real
    b = c.Imag
    Set Secant = Reset(Cos(a) * CosH(b), -Sin(a) * SinH(b))
End Function

Public Function Cosecant(c As ComplexNumber) As ComplexNumber
    Dim a As Double
    Dim b As Double
    a = c.Real
    b = c.Imag
    Set Cosecant = Reset(Sin(a) * CosH(b), Cos(a) * SinH(b))
End Function
'
Public Function Cotangent(c As ComplexNumber) As ComplexNumber
    Dim a As Double
    Dim b As Double
    a = c.Real
    b = c.Imag
    Set Cotangent = Divide(Reset(Sin(a) * CosH(b), Cos(a) * SinH(b)), Reset(Cos(a) * CosH(b), -Sin(a) * SinH(b)))
End Function

Public Function Exsecant(c As ComplexNumber) As ComplexNumber
    Set Exsecant = MinusDouble(Secant(c), 1)
End Function

Public Function Versine(c As ComplexNumber) As ComplexNumber
    Set Versine = Minus(PlusOne(), Cosine(c))
End Function

Public Function Coversine(c As ComplexNumber) As ComplexNumber
    Set Coversine = Minus(PlusOne, Sine(c))
End Function

Public Function Haversine(c As ComplexNumber) As ComplexNumber
    Set Haversine = DivideDouble(Versine(c), 2#)
End Function

Public Function InverseTangent(c As ComplexNumber) As ComplexNumber
    Set InverseTangent = DivideDouble(Times(PlusJay(), Logarithm(Divide(Plus(PlusJay(), c), Minus(PlusJay(), c)))), 2#)
End Function
'
Public Function InverseCotangent(c As ComplexNumber) As ComplexNumber
        Set InverseCotangent = InverseTangent(Reciprocal(c))
End Function
'
Public Function InverseHyperbolicCosecant(c As ComplexNumber) As ComplexNumber
    Set InverseHyperbolicCosecant = Logarithm(Divide(PlusDouble(TimesDouble(SquareRoot(PlusDouble(Times(c, c), 1)), Sign(c)), 1), c))
End Function
'
'Public Function ApproximatelyEqual(ParamArray args() As Variant) As Boolean
'    Dim c0 As Complex
'    Dim c1 As Complex
'    Dim c2 As Complex
'    Dim epsilon As Double
'
'    Dim rUB As Double
'    Dim rLB As Double
'    Dim iUB As Double
'    Dim iLB As Double
'
'    Dim cReal As Double
'    Dim cImag As Double
'
'    Dim bRes As Boolean
'    bRes = False
'    Select Case UBound(args)
'        Case 2 'c1, c2, epsilon
'            epsilon = args(2)
'            Set c2 = args(1)
'            Set c1 = args(0)
'
'            cReal = (c1.GetReal)
'            rUB = cReal + epsilon
'            rLB = cReal - epsilon
'            cImag = (c1.GetImag)
'            iUB = cImag + epsilon
'            iLB = cImag - epsilon
'
'            cReal = (c2.GetReal)
'            cImag = (c2.GetImag)
'
'            If rLB <= cReal And cReal <= rUB Then
'                If iLB <= cImag And cImag <= iUB Then
'                    bRes = True
'                End If
'            End If
'
'        Case 1 'c1, epsilon
'            epsilon = args(1)
'            Set c1 = args(0)
'            Set c0 = CopyOf
'
'            cReal = (c0.GetReal)
'            cImag = (c0.GetImag)
'
'            rUB = cReal + epsilon
'            rLB = cReal - epsilon
'            iUB = cImag + epsilon
'            iLB = cImag - epsilon
'
'            cReal = (c1.GetReal)
'            cImag = (c1.GetImag)
'
'            If rLB <= cReal And cReal <= rUB Then
'                If iLB <= cImag And cImag <= iUB Then
'                    bRes = True
'                End If
'            End If
'        Case Else 'error
'    End Select
'    ApproximatelyEqual = bRes
'End Function
'
Private Function Atan2(dy As Variant, dx As Variant) As Variant
    Dim half_pi As Double
    half_pi = PIE / 2
    Dim a As Double

    'use arctan, avoiding division by zero.

    If Abs(dx) > Abs(dy) Then
        a = Atn(dy / dx)
    Else
        a = Atn(dx / dy) '{ pi/4 <= a <= pi/4 }
        If a < 0 Then
            a = -half_pi - a '{ a is negative, so we're adding }
        Else
            a = half_pi - a
        End If
        End If

    If dx < 0 Then
        If dy < 0 Then
            a = a - PIE
        Else
            a = a + PIE
        End If
    End If
    Atan2 = a
End Function '  { atan2 }
'
'
Private Function SinH(a As Variant) As Double
    SinH = 0.5 * (Exp(a) - Exp(-a))
End Function

Private Function CosH(a As Variant) As Double
    CosH = 0.5 * (Exp(a) + Exp(-a))
End Function
'
Private Function PowerDouble(a As ComplexNumber, b As Double) As ComplexNumber
    Dim re As Double
    Dim im As Double
    re = a.Real
    im = a.Imag
    
    If IsZero(a) Then
        If b = 0# Then
            Set PowerDouble = PlusOne()
        Else
            If b > 0# Then
                Set PowerDouble = Zero()
            Else
                If b < 0# Then
                    Set PowerDouble = Reset(PosInfinity, 0)
                End If
            End If
        End If
    Else
        If im = 0# And re > 0# Then
            Set PowerDouble = Reset(re ^ b, 0)
        Else
            Dim c As Double
            Dim th As Double
            c = (re * re + im * im) ^ (b / 2#)
            th = Atan2(im, re)
            Set PowerDouble = Reset(c * Cos(b * th), c * Sin(b * th))
        End If
    End If
End Function

