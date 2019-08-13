VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Misc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim Strings As New Strings

Public Function RandomComplex(Optional nLimit As Variant = 1000, Optional bFloating As Variant = False) As Variant
    Dim sResult As String
    Dim i As Integer
    Dim k As Double
    Dim j As Integer
    Dim parts(1) As String
    
    For j = 0 To 1
        If toss() = True Then
            k = Rand(nLimit) * IIF(toss() = True, -1, 1)
            If k = 0 Then parts(j) = "0" Else parts(j) = k
        Else
            i = Int(Rand(nLimit)) * IIF(toss() = True, -1, 1)
            If i = 0 Then parts(j) = "0" Else parts(j) = i
        End If
    Next
    
    sResult = parts(0) & IIF(toss() = True, "+", "-") & parts(1) & "i"
    sResult = Replace(sResult, "--", "-")
    sResult = Replace(sResult, "+-", "-")
    RandomComplex = sResult
End Function

Public Function RandomComplexSpaced(Optional nLimit As Variant = 1000, Optional bFloating As Variant = False) As Variant
    Dim sResult As String
    Dim i As Integer
    Dim k As Double
    Dim j As Integer
    Dim parts(1) As String
    
    For j = 0 To 1
        If toss() = True Then
            k = Rand(nLimit) * IIF(toss() = True, -1, 1)
            If k = 0 Then parts(j) = "0" Else parts(j) = k
        Else
            i = Int(Rand(nLimit)) * IIF(toss() = True, -1, 1)
            If i = 0 Then parts(j) = "0" Else parts(j) = i
        End If
    Next
    
    sResult = parts(0) & " " & IIF(toss() = True, "+", "-") & parts(1) & "i"
    sResult = Replace(sResult, "--", "-")
    sResult = Replace(sResult, "+-", "-")
    RandomComplexSpaced = sResult
End Function

Public Function RandomInt(Optional nLimit As Variant = 1000) As Variant
    RandomInt = CStr(Int(Rand(nLimit)) * IIF(toss() = True, -1, 1))
End Function


Public Function RandomReal(Optional nLimit As Variant = 1000) As Variant
    Dim nRes As Variant
    'If VarType(nLimit) = vbDouble Then
    If nLimit <> Int(nLimit) Then
        Do
            nRes = Rnd
            If nRes < nLimit Then
                Exit Do
            End If
        Loop
        nRes = nRes * IIF(toss() = True, -1, 1)
    Else
        nRes = CStr((Int(Rand(nLimit)) & "." & Int(Rand(nLimit))) * IIF(toss() = True, -1, 1))
    End If
    RandomReal = nRes
End Function

Public Function RandomBigReal(Optional nLimit As Variant = 20) As Variant
    Dim sResult As String
    Dim i As Integer
    For i = 1 To nLimit
        sResult = sResult & Int(Rand(10))
    Next
    i = Rand(20) + 1
    sResult = left$(sResult, i) & "." & Mid$(sResult, i + 1)
    While left$(sResult, 1) = "0"
        sResult = Mid$(sResult, 2)
    Wend
    sResult = IIF(toss() = True, "", "-") & sResult
    RandomBigReal = sResult
End Function

Public Function RandomBigInt(Optional nLimit As Variant = 20) As Variant
    Dim sResult As String
    Dim i As Integer
    For i = 1 To nLimit
        sResult = sResult & Int(Rand(10))
    Next
    While left$(sResult, 1) = "0"
        sResult = Mid$(sResult, 2)
    Wend
    sResult = IIF(toss() = True, "", "-") & sResult
    RandomBigInt = sResult
End Function

Friend Function toss() As Boolean
    toss = Rnd > 0.5
End Function

Private Function Rand(nLimit As Variant) As String
    Rand = Rnd * nLimit
End Function

Public Function UNSCIENCE(p1 As Variant) As Variant
    Dim nE As Long
    Dim nDot As Long
    Dim sLeft As String
    Dim sRight As String
    Dim nExponent As Long
    Dim sAfterDecimal As String
    Dim sSign As String
    Dim sResult As String
    
    sResult = vbNullString
    
    If left$(p1, 1) = "-" Then
        sSign = "-"
        p1 = Mid$(p1, 2)
    Else
        sSign = vbNullString
    End If
    
    nE = InStr(LCase$(p1), "e")
    
    If nE > 0 Then
        sLeft = left$(p1, nE - 1)
        sRight = Mid$(p1, nE + 1)
        nExponent = Val(sRight)
        If Sgn(nExponent) = 1 Then
            nDot = InStr(sLeft, ".")
            If nDot > 0 Then
                sAfterDecimal = Mid$(sLeft, nDot + 1)
            Else
                sAfterDecimal = vbNullString
            End If
            sResult = Replace$(sLeft, ".", vbNullString) & String$(nExponent - Len(sAfterDecimal), "0")
        Else
            sLeft = Replace$(sLeft, ".", vbNullString)
            sResult = "0." & String$(Abs(nExponent) - 1, "0") & sLeft
        End If
    Else
        sResult = p1
    End If
    
    UNSCIENCE = sSign & sResult

End Function

Private Sub Class_Initialize()
    Randomize Timer
End Sub
