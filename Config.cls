VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Dim oFSO As Scripting.FileSystemObject
Dim oStr As Scripting.TextStream
Dim oVars As Collections
Dim str As Strings
Dim Fil As Files
Public gsFilename As String

Private Function DetectFile(sFile As Variant) As String
    Dim sResult As String
    Set oFSO = New Scripting.FileSystemObject
    
    'try find it locally
    sResult = ".\" & sFile
    If oFSO.FileExists(sResult) = True Then GoTo done
    
    'try find it with the application
    sResult = App.Path & "\" & sFile
    If oFSO.FileExists(sResult) = True Then GoTo done
    
    'give up
    sResult = sFile
done:
    DetectFile = sResult
End Function

Public Function Load(sFile As Variant) As Boolean
    Dim sLine As String
    Dim sData As String
    Dim aData() As String
    Dim i As Integer
    Dim aCmd() As String
    Dim nMax As Integer
    Dim bWorked As Boolean
    
    bWorked = False
    
    gsFilename = DetectFile(sFile)
    
    If oFSO.FileExists(gsFilename) Then
        bWorked = True
    Else
        GoTo fail
    End If
    
'    Set oFSO = New Scripting.FileSystemObject
'    Set oStr = oFSO.OpenTextFile(gsFilename, ForReading, False, TristateMixed)
'    sData = oStr.ReadAll
'    oStr.Close

    Set Fil = New Files
    sData = Fil.ReadFileE(gsFilename)
    If sData = vbNullString Then
        bWorked = False
        GoTo fail
    End If
    
    aData = Split(sData, vbNewLine)
    nMax = UBound(aData)
    Do While aData(nMax) = ""
        nMax = nMax - 1
    Loop
    
    For i = 0 To nMax
        sLine = aData(i)
        If left(sLine, 1) = "#" Then
            oVars.add sLine, vbNullChar
        Else
            aCmd = Split(sLine, "=", 2)
            If UBound(aCmd) > 0 Then
                oVars.add aCmd(1), aCmd(0)
            Else
                oVars.add sLine, vbNullChar
            End If
        End If
    Next
fail:
    Load = bWorked
End Function

Public Function LoadW(sFile As Variant) As Boolean
    Dim sData As String
    Dim sLine As String
    Dim aData() As String
    Dim aCmd() As String
    Dim i As Integer
    Dim nMax As Integer
    Dim bWorked As Boolean
    
    gsFilename = DetectFile(sFile)
    
    If oFSO.FileExists(gsFilename) Then
        bWorked = True
    Else
        bWorked = False
        GoTo fail
    End If
    
'    Set oFSO = New Scripting.FileSystemObject
'    Set oStr = oFSO.OpenTextFile(gsFilename, ForReading, False, TristateMixed)
'    sData = oStr.ReadAll
'    oStr.Close
    Set Fil = New Files
    sData = Fil.ReadFileE(gsFilename)
    If sData = vbNullString Then
        bWorked = False
        GoTo fail
    End If
    
    sData = Replace(sData, "\" & vbNewLine, "")
    aData = Split(sData, vbNewLine)
    
    aData = Split(sData, vbNewLine)
    nMax = UBound(aData)
    Do While aData(nMax) = ""
        nMax = nMax - 1
    Loop
    
    For i = 0 To nMax
        sLine = aData(i)
        If left$(sLine, 1) = "#" Then
            oVars.add sLine, vbNullChar
        Else
            aCmd = Split(sLine, "=", 2)
            If UBound(aCmd) > 0 Then
                oVars.add aCmd(1), aCmd(0)
            Else
                oVars.add sLine, vbNullChar
            End If
        End If
    Next
fail:
    LoadW = bWorked
End Function

Public Sub Save(Optional sFile As Variant, Optional bAsAnsi = False)
    Dim v As New Variable
    
    If IsMissing(sFile) Then
        sFile = gsFilename
    End If
    
    Set oFSO = New Scripting.FileSystemObject
    Set oStr = oFSO.OpenTextFile(sFile, ForWriting, True, IIF(bAsAnsi, TristateFalse, TristateTrue))
    
    
    For Each v In oVars.NewEnum
        'MsgBox TypeName(V) & " " & TypeName(V.Name) & " " & TypeName(V.Value)
        'Set v = oVars.Item(v)
        If left$(v.Value, 1) = "#" Or v.Value = "" Then
            oStr.WriteLine v.Value
        Else
            oStr.WriteLine v.Name & "=" & v.Value
        End If
    Next
    oStr.Close
End Sub

Public Sub SaveW(Optional sFile As Variant = "", Optional bAsAnsi = False)
    Dim v As Variable
    Dim sLine As String
    
    If sFile = "" Then
        sFile = gsFilename
    End If
    
    Set oFSO = New Scripting.FileSystemObject
    Set oStr = oFSO.OpenTextFile(sFile, ForWriting, True, IIF(bAsAnsi, TristateFalse, TristateTrue))
    
    For Each v In oVars.NewEnum
        'MsgBox TypeName(V) & " " & TypeName(V.Name) & " " & TypeName(V.Value)
        If left$(v.Value, 1) = "#" Then
            'oStr.WriteLine V.Value
            sLine = v.Value
            WrapLine oStr, sLine, 70
        ElseIf v.Value = "" Then
            oStr.WriteLine
        Else
            'oStr.WriteLine V.Name & "=" & V.Value
            sLine = v.Name & "=" & v.Value
            WrapLine oStr, sLine, 70
        End If
    Next
    oStr.Close
End Sub

Private Sub WrapLine(oStream As Scripting.TextStream, sLine As Variant, nLimit As Integer)
    Dim nPos As Integer
    Dim sChar As String
    Dim nSpacePos As Integer
    
    nSpacePos = InStr(1, sLine, " ")
    
    If Len(sLine) <= nLimit Or nSpacePos = 0 Then
        oStream.WriteLine sLine
        GoTo done
    ElseIf nSpacePos > nLimit Then
        nLimit = nSpacePos
    End If
    
    Do While Len(sLine) > nLimit
        For nPos = nLimit To 1 Step -1
            sChar = Mid$(sLine, nPos, 1)
            If sChar = " " Or sChar = "=" Then
                oStream.WriteLine left$(sLine, nPos) & "\"
                sLine = Mid$(sLine, nPos + 1)
                Exit For
            End If
        Next
    Loop
    If Len(sLine) > 0 Then oStream.WriteLine sLine
done:

End Sub

Public Function List(sFilter As Variant) As String
    Dim aResult() As String
    Dim sLeft As String
    Dim sRight As String
    Dim nPos As Integer
    Dim nCursor As Integer
    
    nCursor = 0
    
    ReDim aResult(1000)
    nPos = InStr(1, sFilter, "*")
    If nPos > 0 Then
        sLeft = left(sFilter, nPos - 1)
        sRight = Mid(sFilter, nPos + 1)
    Else
        sLeft = sFilter
        sRight = ""
    End If
    
    Dim v As Variable
    For Each v In oVars
        If str.BeginsWith(v.Name, sLeft) Then
            If str.EndsWith(v.Name, sRight) Then
                aResult(nCursor) = v.Name
                nCursor = nCursor + 1
            End If
        End If
    Next
    
    ReDim Preserve aResult(nCursor - 1)
    List = Join(aResult, vbTab)
End Function

Public Sub Define(sKey As Variant, sValue As Variant)
    If oVars.Exists(sKey) Then
        oVars.Modify sValue, sKey
    Else
        oVars.add sValue, sKey
    End If
End Sub

'Public Sub Change(sKey As String, sValue As String)
'    oVars.Modify sKey, sValue
'End Sub

Public Function Recall(sKey As Variant) As Variant
    Dim v As Variable
    If oVars.Exists(sKey) Then
        Set v = oVars.Item(sKey)
        Recall = v.Value
    Else
        Error 448
    End If
End Function

Public Function RecallElse(sKey As Variant, sDefault As Variant) As Variant
    Dim v As Variant
    On Error Resume Next
    v = Recall(sKey)
    If Err.Number = 448 Or v = vbNullString Then
        RecallElse = sDefault
    Else
        RecallElse = v
    End If
End Function

Private Sub Class_Initialize()
    Set oVars = New Collections
End Sub

Private Sub Class_Terminate()
    Set oVars = Nothing
End Sub


