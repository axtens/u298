VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Arrays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim stdstr As New Strings
Public Sub DeleteNth(anArray As Variant, ByVal nElement As Long)
    Dim lIndex As Long
    If nElement = UBound(anArray) Then
        If UBound(anArray) - 1 = -1 Then
            ReDim anArray(0)
        Else
            ReDim Preserve anArray(UBound(anArray) - 1)
        End If
    Else
        If UBound(anArray) > 0 Then
            For lIndex = nElement To UBound(anArray) - 1
                anArray(lIndex) = anArray(lIndex + 1)
            Next
            ReDim Preserve anArray(UBound(anArray) - 1)
        Else
            ReDim anArray(0)
        End If
    End If
End Sub

Public Sub DeleteLast(anArray As Variant)
    DeleteNth anArray, UBound(anArray)
End Sub

Public Sub DeleteFirst(anArray As Variant)
    DeleteNth anArray, LBound(anArray)
End Sub


Function AAdd(ByRef aToThis As Variant, ByVal xAddThis As Variant) As Integer
    Dim nPos As Integer
    nPos = UBound(aToThis) + 1
    ReDim Preserve aToThis(nPos)
    aToThis(nPos) = xAddThis
    AAdd = nPos
End Function

Public Function AAdd_PreIncrement(a As Variant, ByVal X As Variant)
    Dim nPos As Long
    nPos = UBound(a)
    nPos = nPos + 1
    ReDim Preserve a(nPos)
    a(nPos) = X
    AAdd_PreIncrement = nPos
End Function

Public Function AAdd_PostIncrement(a As Variant, ByVal X As Variant)
    Dim nPos As Long
    nPos = UBound(a)
    a(nPos) = X
    nPos = nPos + 1
    ReDim Preserve a(nPos)
    AAdd_PostIncrement = nPos
End Function

Function InArray(sItem As Variant, aList As Variant) As Boolean
    Dim i As Integer
    Dim u As Integer
    Dim bResult As Boolean
    bResult = False
    u = UBound(aList)
    For i = 0 To u
        If sItem = aList(i) Then
            bResult = True
            Exit For
        End If
    Next
    InArray = bResult
End Function

Function StrInArray(sItem As Variant, aList As Variant, Optional bCaseInsens As Boolean = True) As String
    Dim i As Integer
    Dim u As Integer
    Dim bResult As Boolean
    'WScript.Echo "Searching for " & sItem & " in " & Join( aList, "|" ) & " " & IIF( bCaseInsens, "Case Insensitive", "Case Sensitive" )
    bResult = False
    u = UBound(aList)
    For i = 0 To u
        If bCaseInsens = True Then
            bResult = (LCase(sItem) = LCase(aList(i)))
        Else
            bResult = (sItem = aList(i))
        End If
        If bResult = True Then Exit For
    Next
    StrInArray = bResult
End Function

Function AMax(aVec As Variant) As Variant
    Dim vTemp As Variant
    Dim vResult As Variant

    vResult = aVec(0)

    For Each vTemp In aVec
        If vTemp > vResult Then
            vResult = vTemp
        End If
    Next
    AMax = vResult
End Function

Function AMin(aVec As Variant) As Variant
    Dim vTemp As Variant
    Dim vResult As Variant

    vResult = aVec(0)

    For Each vTemp In aVec
        If vTemp < vResult Then
            vResult = vTemp
        End If
    Next
    AMin = vResult
End Function

Function ShowStructure(vToShow As Variant) As String
    ShowStructure = ArrayToString(vToShow, ",", "{", "}")
End Function

Function ArrayToString(vToShow As Variant, Optional sSep As Variant = " ", Optional sOpener As Variant = vbNullString, Optional sCloser As Variant = vbNullString) As String
    Dim nCounter As Integer
    Dim sResult As String
    Dim nUBound As Integer
    Dim nLBound As Integer
    
    sResult = vbNullString

    If IsArray(vToShow) Then
        nUBound = UBound(vToShow)
        nLBound = LBound(vToShow)
        If nUBound >= nLBound Then
            For nCounter = nLBound To nUBound
                sResult = sResult & ArrayToString(vToShow(nCounter))
                If nCounter < nUBound Then sResult = sResult & sSep
            Next
            sResult = sOpener & sResult & sCloser
        Else
            sResult = sOpener & sCloser
        End If
    Else
        sResult = stdstr.AsString(vToShow)
    End If
    ArrayToString = sResult
End Function

Function Reverse1DArray(aInput As Variant) As Variant
    Dim iLB As Integer
    Dim iUB As Integer
    Dim aRes() As Variant
    Dim i As Integer
    
    iLB = LBound(aInput)
    iUB = UBound(aInput)
    ReDim aRes(iUB - iLB)
    For i = iUB To iLB Step -1
        aRes(iUB - i) = aInput(i)
    Next
    Reverse1DArray = aRes
End Function

Function GenerateRangeFromString(sRange As Variant) As Variant
    Dim nTop As Double
    Dim nBot As Double
    Dim nStep As Double
    
    Dim aRange() As String
    Dim aParam() As String
    
    aParam = Split(sRange, ",")
    nStep = 1
    If UBound(aParam) = 1 Then
        nStep = Val(aParam(1))
        If nStep = 0 Then nStep = 1
    End If
    
    aRange = Split(aParam(0), "..")
    
    If UBound(aRange) <= 0 Then
        GenerateRangeFromString = sRange
    Else
        If IsNumeric(aRange(0)) Then
            nTop = Val(aRange(0))
            nBot = Val(aRange(1))
        Else
            nTop = Asc(aRange(0))
            nBot = Asc(aRange(1))
        End If
        GenerateRangeFromString = GenerateRange(nTop, nBot, nStep)
    End If
End Function

Function GenerateRange(nBot As Double, nTop As Double, Optional nStep As Double = 1) As Variant
    Dim aResult() As Variant
    aResult = Array()
    Dim i As Double
    For i = nBot To nTop Step nStep
        AAdd aResult, i
    Next
    GenerateRange = aResult
End Function


Public Function IsVector(a As Variant) As Boolean
'check if argument is a true vector array
    On Error GoTo Error_Handler
    Dim n
    IsVector = False
    n = UBound(a, 1)
    IsVector = True
    n = UBound(a, 2)
    IsVector = False
Error_Handler:
End Function

Public Function IsMatrix(a As Variant) As Boolean
'check if argument is a true matrix array
    On Error GoTo Error_Handler
    Dim n
    IsMatrix = False
    n = UBound(a, 1)
    n = UBound(a, 2)
    IsMatrix = True
Error_Handler:
End Function

Public Function Bounds(a As Variant) As Variant
    If IsMatrix(a) Then
        Bounds = Array(UBound(a, 1), UBound(a, 2))
    ElseIf IsVector(a) Then
        Bounds = UBound(a)
    Else
        Bounds = Empty
    End If
End Function

Public Function GetDimensions(a As Variant) As Variant
    Dim dims() As Integer
    ReDim dims(10)
    Dim i As Integer
    Dim lim As Integer
    lim = -1
    On Error GoTo done
    For i = 0 To UBound(dims)
        dims(i) = UBound(a, i + 1)
        lim = i
    Next
done:
    ReDim Preserve dims(lim)
    GetDimensions = dims
End Function
