VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SymbolTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private sSystemPrefix As String
Private oHash As Collections
Dim stdS As New Strings

Private Sub Class_Initialize()
    Set oHash = New Collections
    sSystemPrefix = "$"
End Sub

Private Sub Class_Terminate()
    Set oHash = Nothing
End Sub

Public Function Increment(sName As Variant) As Integer
    Dim nResult As Integer
    If oHash.Exists(sName) Then
        nResult = Recall(sName) + 1
    Else
        nResult = 1
    End If
    Store sName, nResult
    Increment = nResult
End Function

Public Function Append(sName As Variant, sValue As Variant) As String
    Dim sResult As String
    If oHash.Exists(sName) Then
        sResult = Recall(sName) & "," & sValue
    Else
        sResult = sValue
    End If
    Store sName, sResult
    Append = sResult
End Function

Public Sub Store(sName As Variant, vValue As Variant)
    Dim v As Variable
    If oHash.Exists(sName) Then
        Set v = oHash.Item(sName)
        'oHash.Item(sName) = vValue
        v.Value = vValue
    Else
        oHash.add vValue, sName
    End If
End Sub

Public Function Recall(sName As Variant) As Variant
    If oHash.Exists(sName) Then
        Recall = oHash.Item(sName).Value
    Else
        Recall = Null
    End If
End Function

Public Function Keys() As Variant()
    Keys = oHash.Keys
End Function

Public Function Items() As Variant()
    Items = oHash.Items
End Function

Public Function SymList()
    Dim nCount As Integer
    Dim aResult() As String
    Dim aKeys() As Variant
    Dim aItems() As Variant
    Dim i As Integer

    nCount = oHash.Count
    'WScript.Echo nCount
    ReDim aResult(nCount - 1)
    aKeys = oHash.Keys
    aItems = oHash.Items

    For i = 0 To nCount - 1
        'aResult( 0, i ) = aKeys( i )
        'aResult( 1, i ) = aItems( i )
        aResult(i) = aKeys(i) & "^" & aItems(i)
    Next
    'WScript.Echo UBound( aResult )
    SymList = aResult
End Function

Public Function SymListText() As String
    Dim aList As Variant
    Dim sItem As String
    Dim i As Integer
    Dim sResult As String
    aList = SymList()
    For i = 0 To UBound(aList)
        sItem = Replace(aList(i), "^", "=")
        sResult = sResult & sItem & vbNewLine
    Next
    SymListText = sResult
End Function

Public Function IsSym(sName As Variant) As Boolean
    IsSym = oHash.Exists(sName)
End Function

Public Sub StoreDup(sName As Variant, xValue As Variant)
    Dim i As Integer
    Dim sNewName As String
    sNewName = sName
    i = 1
    If IsSym(sNewName) Then
        Do
            sNewName = sName & "$" & i
            If oHash.Exists(sNewName) = False Then
                Exit Do
            End If
            i = i + 1
        Loop
    End If

    'WScript.Echo "StoreDup", sNewName, xValue
    Store sNewName, xValue
End Sub

Public Sub StoreDup2(sName As Variant, xValue As Variant)
    Dim i As Integer
    Dim sNewName As String
    sNewName = sName
    i = 1
    If IsSym(sNewName) Then
    Store sNewName, Recall(sNewName) & "," & xValue
Else
    Store sNewName, xValue
End If
End Sub

Public Sub Parse(sData As Variant, sSep As Variant)
    Dim aData() As String
    Dim i As Integer
    Dim sLine As Variant
    Dim aLine() As String
    Dim sKey As String
    Dim sValue As String
    
    aData = Split(sData, vbNewLine)
    For Each sLine In aData
        Log sLine
        If sLine <> vbNullString Then
            aLine = Split(sLine, sSep)
            sKey = aLine(0)
            sValue = aLine(1)
            sKey = UCase(RTrim(stdS.RemoveSpaces(sKey)))
            sValue = RTrim(sValue)
            Store sKey, sValue
        End If
    Next
End Sub

Public Sub Remove(sData As Variant)
    If oHash.Exists(sData) Then
        oHash.Remove sData
    End If
End Sub

Public Sub Clear()
    Dim aList() As String
    Dim i As Integer
    Dim sLine As String
    Dim aLine() As String
    aList = SymList()
    For i = 0 To UBound(aList)
        sLine = aList(i)
        aLine = Split(sLine, "^")
        If stdS.BeginsWith(aLine(0), sSystemPrefix, False) = False Then
            Remove aLine(0)
        End If
    Next
End Sub

