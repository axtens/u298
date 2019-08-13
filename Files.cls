VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Files"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim oFSO As Scripting.FileSystemObject
Dim stdS As New Strings

Const WFAPPEND As Long = 1, WFCREATE As Long = 2
Private Const WriteFileUBadMode As Long = vbObjectError + 1001
Private Const WriteFileABadMode As Long = vbObjectError + 1002

Public Enum EFilesClass
    WF_APPEND = WFAPPEND
    WF_CREATE = WFCREATE
End Enum

Public Function ReadFileAToArray(ByVal Filename As Variant) As Variant
    Dim oStream As Scripting.TextStream
    Dim sData As String
    Dim aRes As Variant
    
    Set oFSO = New Scripting.FileSystemObject
    sData = vbNullString

    Set oStream = oFSO.OpenTextFile(Filename, ForReading, False, TristateTrue)
    On Error Resume Next
    sData = oStream.ReadAll
    If Err.Number <> 0 Then
        sData = vbNullString
    End If
    On Error GoTo 0
    oStream.Close

    If sData <> vbNullString Then
        aRes = Split(sData, vbNewLine)
        Do
            If aRes(UBound(aRes)) = vbNullString Then
                ReDim Preserve aRes(UBound(aRes) - 1)
            Else
                Exit Do
            End If
        Loop
    End If
    ReadFileAToArray = aRes
End Function

Public Function ReadFileUToArray(ByVal Filename As Variant) As Variant
    Dim oStream As Scripting.TextStream
    Dim sData As String
    Dim aRes As Variant
    
    Set oFSO = New Scripting.FileSystemObject
    sData = vbNullString

    Set oStream = oFSO.OpenTextFile(Filename, ForReading, False, TristateFalse)
    On Error Resume Next
    sData = oStream.ReadAll
    If Err.Number <> 0 Then
        sData = vbNullString
    End If
    On Error GoTo 0
    oStream.Close

    If sData <> vbNullString Then
        aRes = Split(sData, vbNewLine)
        Do
            If aRes(UBound(aRes)) = vbNullString Then
                ReDim Preserve aRes(UBound(aRes) - 1)
            Else
                Exit Do
            End If
        Loop
    End If
    ReadFileUToArray = aRes
End Function

Public Function ReadFileEToArray(ByVal Filename As Variant) As Variant
    Dim oStream As Scripting.TextStream
    Dim sData As String
    Dim aRes As Variant
    
    Set oFSO = New Scripting.FileSystemObject
    sData = vbNullString

    Set oStream = oFSO.OpenTextFile(Filename, ForReading, False, TristateMixed)
    On Error Resume Next
    sData = oStream.ReadAll
    If Err.Number <> 0 Then
        sData = vbNullString
    End If
    On Error GoTo 0
    oStream.Close

    If sData <> vbNullString Then
        aRes = Split(sData, vbNewLine)
        Do
            If aRes(UBound(aRes)) = vbNullString Then
                ReDim Preserve aRes(UBound(aRes) - 1)
            Else
                Exit Do
            End If
        Loop
    End If
    
'    Dim i As Long
'    For i = 0 To UBound(aRes)
'        aRes(i) = CVar(aRes(i))
'    Next
    
    ReadFileEToArray = aRes
    
End Function

Public Sub DeleteFile(ByVal cFilename As Variant)
    On Error Resume Next
    Kill cFilename
    On Error GoTo 0
End Sub


Function CollectFiles(sDirectory As Variant, sFileType As Variant) As Variant
'    Dim aResult() As Variant
'    Dim oShell As New shell32.Shell
'    Dim oFolder As shell32.Folder
'    Dim cItems As shell32.FolderItems
'    Dim sItem As shell32.FolderItem
'    Dim i As Integer
'    Dim nCount As Integer
'
'    On Error Resume Next
'    Set cItems = oShell.NameSpace(sDirectory).Items
'    If Err.Number = 0 Then
'        ReDim aResult(cItems.Count - 1)
'        nCount = 0
'        For i = 0 To cItems.Count - 1
'            Set sItem = cItems.Item(i)
'            If stdS.EndsWith(sItem.Name, sFileType, True) Then
'                aResult(nCount) = sItem.Name
'                nCount = nCount + 1
'            End If
'        Next
'    End If
'    CollectFiles = aResult
    Dim oFolder As Scripting.Folder
    Dim oFiles As Scripting.Files
    Dim oFile As Scripting.File
    Dim nCount As Integer
    
    Set oFSO = New FileSystemObject
    Set oFolder = oFSO.GetFolder(sDirectory)
    Set oFiles = oFolder.Files
    ReDim aResult(oFiles.Count)
    nCount = -1
    For Each oFile In oFiles
        If stdS.EndsWith(oFile.Name, sFileType, True) Then
            nCount = nCount + 1
            aResult(nCount) = oFSO.BuildPath(oFile.ParentFolder, oFile.Name)
        End If
    Next
    ReDim Preserve aResult(nCount)
    CollectFiles = aResult
End Function

Function ReadFileA(Filename As Variant) As String
    Dim oStream As Scripting.TextStream
    Dim sData As String

    Set oFSO = New Scripting.FileSystemObject
    sData = vbNullString

    Set oStream = oFSO.OpenTextFile(Filename, ForReading, False, TristateTrue)
    On Error Resume Next
    sData = oStream.ReadAll
    If Err.Number <> 0 Then
        sData = vbNullString
    End If
    On Error GoTo 0
    oStream.Close

    ReadFileA = sData
End Function

Function ReadFileU(Filename As Variant) As String
    Dim oStream As Scripting.TextStream
    Dim sData As String
    
    Set oFSO = New Scripting.FileSystemObject
    sData = vbNullString
    
    On Error Resume Next
    Set oStream = oFSO.OpenTextFile(Filename, ForReading, False, TristateTrue)
    If Err.Number = 0 Then
        sData = oStream.ReadAll
        If Err.Number <> 0 Then
            sData = vbNullString
        End If
        oStream.Close
    End If
    On Error GoTo 0

    ReadFileU = sData
End Function

Function ReadFileE(Filename As Variant) As String
    Dim oStream As Scripting.TextStream
    Dim sData As String

    Set oFSO = New Scripting.FileSystemObject
    sData = vbNullString
    On Error Resume Next
    Set oStream = oFSO.OpenTextFile(Filename, ForReading, False, TristateMixed)
    If Err.Number = 0 Then
        sData = oStream.ReadAll
        If Err.Number <> 0 Then
            sData = vbNullString
        End If
        oStream.Close
    End If
    On Error GoTo 0

    ReadFileE = sData
End Function

Function ReadFirstLineA(Filename As Variant) As String
    Dim oStream As Scripting.TextStream
    Dim sData As String

    sData = vbNullString
    Set oFSO = New Scripting.FileSystemObject
    Set oStream = oFSO.OpenTextFile(Filename, ForReading, False, TristateFalse)
    sData = oStream.ReadLine
    oStream.Close

    ReadFirstLineA = sData
End Function

Function ReadFirstLineU(Filename As Variant) As String
    Dim oStream As Scripting.TextStream
    Dim sData As String

    sData = vbNullString
    Set oFSO = New Scripting.FileSystemObject
    Set oStream = oFSO.OpenTextFile(Filename, ForReading, False, TristateTrue)
    sData = oStream.ReadLine
    oStream.Close

    ReadFirstLineU = sData
End Function

Function ReadFirstLineE(Filename As Variant) As String
    Dim oStream As Scripting.TextStream
    Dim sData As String

    sData = vbNullString
    Set oFSO = New Scripting.FileSystemObject
    Set oStream = oFSO.OpenTextFile(Filename, ForReading, False, TristateMixed)
    sData = oStream.ReadLine
    oStream.Close

    ReadFirstLineE = sData
End Function

Function Exists(sFile As Variant) As Boolean
    Set oFSO = New Scripting.FileSystemObject
    Exists = oFSO.FileExists(sFile)
End Function

Sub WriteFileA(sFilename As Variant, sContents As Variant, nMode As Long)
    Dim oStream As Scripting.TextStream
    Set oFSO = New Scripting.FileSystemObject
    If nMode = WFAPPEND Then
        Set oStream = oFSO.OpenTextFile(sFilename, ForAppending, True, TristateFalse)
    ElseIf nMode = WFCREATE Then
        Set oStream = oFSO.OpenTextFile(sFilename, ForWriting, True, TristateFalse)
    Else
        Err.Raise WriteFileABadMode, , "WriteAFile: bad mode (neither append, nor create)"
    End If
    
    oStream.Write sContents
    oStream.Close
    Set oStream = Nothing
End Sub

Sub WriteFileU(sFilename As Variant, sContents As Variant, nMode As Long)
    Dim oStream As Scripting.TextStream
    Set oFSO = New Scripting.FileSystemObject
    If nMode = WFAPPEND Then
        Set oStream = oFSO.OpenTextFile(sFilename, ForAppending, True, TristateTrue)
    ElseIf nMode = WFCREATE Then
        Set oStream = oFSO.OpenTextFile(sFilename, ForWriting, True, TristateTrue)
    Else
        Err.Raise WriteFileUBadMode, , "WriteFileU: bad mode (neither append, nor create)"
    End If
    
    oStream.Write sContents
    oStream.Close
    Set oStream = Nothing
End Sub

Sub WriteFileB(sFilename As Variant, nOffset As Long, vData As Variant)
    Dim i As Integer
    Dim j As Integer
    Dim l As Long
    Dim d As Double
    Dim s As String
    Open sFilename For Binary As #1
    Seek #1, nOffset
    If IsArray(vData) Then
        For i = 0 To UBound(vData)
            Select Case TypeName(vData(i))
            Case Is = "Integer"
                j = vData(i)
                Put #1, , j
            Case Is = "String"
                s = vData(i)
                Put #1, , s
            Case Is = "Long"
                l = vData(i)
                Put #1, , l
            Case Is = "Double"
                d = vData(i)
                Put #1, , d
            End Select
        Next
    Else
        Select Case TypeName(vData)
        Case Is = "Integer"
            j = vData
            Put #1, , j
        Case Is = "String"
            s = vData
            Put #1, , s
        Case Is = "Long"
            l = vData
            Put #1, , l
        Case Is = "Double"
            d = vData
            Put #1, , d
        End Select
    End If
    Close #1
End Sub

'Sub IncludeClass(scriptName)
'    Dim sScript
'    Dim oStream
'    'on error resume next
'    Err.Clear
'    With CreateObject("Scripting.FileSystemobject")
'        Set oStream = .OpenTextFile(scriptName)
'    End With
'    If Err.Number > 0 Then
'        MsgBox "Could not open " & scriptName
'        WScript.Quit
'    End If
'    Err.Clear
'    sScript = oStream.ReadAll()
'    oStream.Close
'    ExecuteGlobal sScript
'    If Err.Number <> 0 Then
'        Die1 "Error in loaded script"
'    End If
'    On Error GoTo 0
'End Sub

'Function ScriptText()
'    ScriptText = WScript.ScriptFullName & ".txt"
'End Function

Function RandomInputFile(sExt As Variant) As String
    Dim tempname As String
    Dim tempfolder As String
    Dim TempFile As String
    Dim n As Integer
    Set oFSO = New Scripting.FileSystemObject
    tempfolder = oFSO.GetSpecialFolder(2) '"C:\temp\rad"
    tempname = oFSO.GetTempName
    n = InStr(tempname, ".")
    RandomInputFile = tempfolder & "\" & left(tempname, n - 1) & sExt
End Function

Function RandomOutputFile(sRandomInputFile As Variant, sExt As Variant) As String
    Dim n As Integer
    n = InStr(sRandomInputFile, ".")
    RandomOutputFile = left(sRandomInputFile, n - 1) & sExt
End Function

Function FileModificationDate(sFile As Variant) As Date
    Set oFSO = New FileSystemObject
    FileModificationDate = oFSO.getfile(sFile).DateLastModified
End Function

Function FileHasBeenModified(sFile As Variant, dThen As Date) As Boolean
    FileHasBeenModified = FileModificationDate(sFile) > dThen
End Function

Function IsUTF16LE(sFile As Variant) As Boolean
    Dim bResult As Boolean
    Dim sBOM As String
    bResult = False
    
    If Exists(sFile) Then
        Open sFile For Input Lock Read As 1 'hFile
        sBOM = Input$(2, 1) 'hFile)
        Close 1 'hFile
    
        If sBOM = Chr(255) & Chr(254) Then
            bResult = True
        End If
    End If
    
    IsUTF16LE = bResult
End Function
