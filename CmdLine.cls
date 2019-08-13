VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CmdLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Attribute ToString.VB_Description = "Returns command line of executing process."
'Attribute ToString.VB_UserMemId = 0
' *********************************************************************
'  Copyright ¬1998-2004 Karl E. Peterson, All Rights Reserved
'  http://www.mvps.org/vb
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************
Option Explicit

Private Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineW" () As Long
Private Declare Function CommandLineToArgv Lib "shell32" Alias "CommandLineToArgvW" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long

' Some calls need to know OS
Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

' Platform ID constants
Private Const VER_PLATFORM_WIN32s As Long = &H0
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = &H1
Private Const VER_PLATFORM_WIN32_NT As Long = &H2

' Member variables
Private m_CmdLine As String
Private m_AppExeName As String
Private m_AppPath As String
Private m_Arguments() As String
Private m_Delimiters As String
Private m_CaseSensitive As Boolean

' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' If this class is used within a VBA project, such as  !
' in Excel or Word, set this constant to True!         !
#Const RunningVBA = False                              '
' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

' *********************************************
'  Initialize/Terminate
' *********************************************
Private Sub Class_Initialize()
   ' Set default values.
   m_CaseSensitive = False
   ' Read command line.
   Me.Refresh
End Sub

Private Sub Class_Terminate()
   ' Nothing to do
End Sub

' *********************************************
'  Public Properties (Read-Only)
' *********************************************
Public Property Get AppExeName(Optional BaseOnly As Boolean = False) As String
   Dim dot As Long
   Dim i As Long
   If BaseOnly Then
      ' Find last dot in name.
      For i = Len(m_AppExeName) To 1 Step -1
         If Mid$(m_AppExeName, i, 1) = "." Then
            dot = i
            Exit For
         End If
      Next i
      ' Chop off extension, if dot was found.
      If dot > 1 Then
         AppExeName = left$(m_AppExeName, dot - 1)
      Else
         AppExeName = m_AppExeName
      End If
   Else
      AppExeName = m_AppExeName
   End If
End Property

Public Property Get AppPath() As String
   AppPath = m_AppPath
End Property

Public Property Get Argument(ByVal Which As Long) As String
   If Which >= 0 And Which <= UBound(m_Arguments) Then
      Argument = m_Arguments(Which)
   End If
End Property

Public Property Get Arguments() As Long
   Arguments = UBound(m_Arguments)
End Property

Public Property Let CaseSensitive(ByVal NewVal As Boolean)
   m_CaseSensitive = NewVal
End Property

Public Property Get CaseSensitive() As Boolean
   CaseSensitive = m_CaseSensitive
End Property

' *********************************************
'  Public Methods
' *********************************************
Public Function FlaggedArg(ByVal Flag As String) As String
   ' This function will scan the argument list, looking for
   ' one that starts with the passed flag. If it's found, and
   ' the passed flag is the entire argument, the following
   ' argument is returned. If the passed flag isn't the entire
   ' argument, the portion following the flag is returned.
   Dim i As Long
   Dim sRet As String
   Dim CompareFlag As VbCompareMethod
   
   ' Convert flag to lowercase if case isn't important.
   If m_CaseSensitive Then
      CompareFlag = vbBinaryCompare
   Else
      CompareFlag = vbTextCompare
   End If
   
   ' Scan arglist, looking for passed flag.
   For i = 1 To UBound(m_Arguments)
      If InStr(1, m_Arguments(i), Flag, CompareFlag) = 1 Then
         ' Base return on whether argument follows directly
         ' after flag, or with space/colon delimiter.
         If Len(m_Arguments(i)) > Len(Flag) Then
            sRet = Mid$(m_Arguments(i), Len(Flag) + 1)
            If Len(sRet) > 1 Then
               If InStr(" :", left$(sRet, 1)) Then
                  ' Trim first character.
                  sRet = Mid$(sRet, 2)
               End If
            End If
         Else
            If i < UBound(m_Arguments) Then
               sRet = m_Arguments(i + 1)
            End If
         End If
         ' All done here.
         Exit For
      End If
   Next i
   
   ' Return results
   FlaggedArg = sRet
End Function

Public Function FlagPresent(ByVal Flag As String) As Long
   ' This function simply scans the argument list,
   ' looking for the passed flag, returns result.
   Dim i As Long
   Dim CompareFlag As VbCompareMethod
   
   ' Convert flag to lowercase if case isn't important.
   If m_CaseSensitive Then
      CompareFlag = vbBinaryCompare
   Else
      CompareFlag = vbTextCompare
   End If
   
   ' Scan arglist, looking for passed flag.
   For i = 1 To UBound(m_Arguments)
      If StrComp(m_Arguments(i), Flag, CompareFlag) = 0 Then
         ' Found it, return matching index.
         FlagPresent = i
         Exit For
      End If
   Next i
End Function

Public Function FlagPresentFromList(ParamArray Flags()) As Boolean
   Dim i As Long
   ' Pass each element in Flags() to FlagPresent,
   ' stopping when one is found.
   For i = LBound(Flags) To UBound(Flags)
      If CBool(Me.FlagPresent(Flags(i))) Then
         FlagPresentFromList = True
         Exit For
      End If
   Next i
End Function

Public Sub Refresh()
   Dim os As OSVERSIONINFO
   Dim i As Long
   
   If Compiled() Then
      ' Get actual command line
      m_CmdLine = PointerToStringW(GetCommandLine())
      
      ' Branch based on OS
      os.dwOSVersionInfoSize = Len(os)
      Call GetVersionEx(os)
      If os.dwPlatformId = VER_PLATFORM_WIN32_NT Then
         Call RefreshNT
      Else 'Win9x
         Call Refresh9x
      End If
      
      ' Parse out path/exename
      If InStr(m_Arguments(0), "\") Then
         For i = Len(m_Arguments(0)) To 1 Step -1
            If Mid$(m_Arguments(0), i, 1) = "\" Then
               m_AppExeName = Mid$(m_Arguments(0), i + 1)
               m_AppPath = left$(m_Arguments(0), i)
               Exit For
            End If
         Next i
      Else
         m_AppExeName = m_Arguments(0)
         m_AppPath = ""
      End If

   Else
      ' Do the best we can with substitutes if this is
      ' running under the Classic VB IDE.
      #If RunningVBA = False Then
         ' If any of the next lines "error out" on you,
         ' or if this class isn't working in the VB IDE,
         ' be sure the RunningVBA #Const is set correctly
         ' at the top of this module.
         m_CmdLine = Trim$(VBA.Command$)
         m_AppExeName = App.ExeName
         m_AppPath = App.Path
         If right$(m_AppPath, 1) <> "\" Then m_AppPath = m_AppPath & "\"
         
         If Len(m_CmdLine) Then
            Call ParseString(m_CmdLine, m_Arguments(), " ")
            ' Shift arguments up one slot.
            ReDim Preserve m_Arguments(0 To UBound(m_Arguments) + 1) As String
            For i = UBound(m_Arguments) To 1 Step -1
               m_Arguments(i) = m_Arguments(i - 1)
            Next i
         Else
            ReDim m_Arguments(0 To 0) As String
         End If
         m_Arguments(0) = m_AppPath & m_AppExeName
      #End If
   End If
End Sub

Public Function toString() As String
   ' Default procedure: just send whole thing
   toString = m_CmdLine
End Function

' *********************************************
'  Private Methods
' *********************************************
Private Function Compiled() As Boolean
   ' Always consider this class to be compiled,
   ' if running under VBA rather than Classic VB.
   #If RunningVBA Then
      Compiled = True
   #Else
      ' Determine if running from EXE/IDE.
      On Error Resume Next
      Debug.Print 1 / 0
      Compiled = (Err.Number = 0)
   #End If
End Function

Private Sub RefreshNT()
   Dim lpData As Long
   Dim lpArgument As Long
   Dim nArgs As Long
   Dim i As Long
   
   ' Use API to return a constructed argument list
   ' which is an array of Unicode strings.
   lpData = CommandLineToArgv(GetCommandLine(), nArgs)
   If lpData Then
      ReDim m_Arguments(0 To nArgs - 1) As String
      ' Extract individual arguments from array, starting
      ' with element 1, because 0 contains the potentially
      ' unqualified appname.
      For i = 1 To nArgs - 1
         lpArgument = PointerToDWord(lpData + (i * 4))
         m_Arguments(i) = PointerToStringW(lpArgument)
      Next i
   End If
   Call GlobalFree(lpData)
   
   ' Get the fully-qualified name of executable.
   lpData = CommandLineToArgv(StrPtr(""), nArgs)
   If lpData Then
      ' Extract individual arguments from array.
      For i = 0 To nArgs - 1
         lpArgument = PointerToDWord(lpData + (i * 4))
         m_Arguments(0) = m_Arguments(0) & PointerToStringW(lpArgument) & " "
      Next i
      m_Arguments(0) = Trim$(m_Arguments(0))
   End If
   Call GlobalFree(lpData)
End Sub

Private Sub Refresh9x()
   Dim ExeName As String
   Dim ExeTerminator As Long
   Dim i As Long
   ' GetCommandLine returns EXE enclosed in quotes
   ' with GUI apps, but not console apps, so lop off
   ' the beginning if needed, and then parse rest.
   ' In console apps, the ExeName ends with first space.
   If InStr(m_CmdLine, """") = 1 Then
      ExeTerminator = InStr(2, m_CmdLine, """")
      ExeName = Mid$(m_CmdLine, 2, ExeTerminator - 2)
   Else
      ExeTerminator = InStr(m_CmdLine, " ")
      ExeName = left$(m_CmdLine, ExeTerminator)
   End If
   
   If Len(Trim$(m_CmdLine)) > ExeTerminator Then
      Call ParseString(Mid$(m_CmdLine, ExeTerminator + 1), m_Arguments(), " ")
      ' Shift arguments up one slot.
      ReDim Preserve m_Arguments(0 To UBound(m_Arguments) + 1) As String
      For i = UBound(m_Arguments) To 1 Step -1
         m_Arguments(i) = m_Arguments(i - 1)
      Next i
   Else
      ReDim m_Arguments(0 To 0) As String
   End If
   ' Insert ExeName into paramarray
   m_Arguments(0) = ExeName
End Sub

Private Sub RefreshIDE()

End Sub

Private Function ParseString(ByVal StrIn As String, StrOut() As String, Optional Delimiters As Variant) As Integer
   Dim InElement As Boolean
   Dim NumEls As Integer
   Dim nPos As Long
   Dim char As String
   Dim Delimit As String
   
   ' Check for missing delimiter, if missing use a space.
   If IsMissing(Delimiters) Then
      Delimiters = " "
   End If
   
   ' InElement serves as a flag to tell if we're currently
   ' processing an element or are in-between.
   InElement = False
   
   ' Process each character in string.
   For nPos = 1 To Len(StrIn)
      ' Extract current character
      char = Mid(StrIn, nPos, 1)
      If InStr(Delimiters, char) = 0 Then
         If Not InElement Then
            ' Set flag to indicate we're moving thru an element,
            ' increment argument counter, expand return array.
            InElement = True
            NumEls = NumEls + 1
            ReDim Preserve StrOut(0 To NumEls - 1)
            StrOut(NumEls - 1) = ""
         End If
         
         ' Append current character to current element.
         StrOut(NumEls - 1) = StrOut(NumEls - 1) & char
      Else
         ' Set flag to indicate we're between arguments.
         InElement = False
      End If
   Next nPos
   
   ' Assign number of arguments as return value for function.
   ParseString = NumEls
End Function

Private Function PointerToDWord(ByVal lpDWord As Long) As Long
   Dim nRet As Long
   If lpDWord Then
      CopyMemory nRet, ByVal lpDWord, 4
      PointerToDWord = nRet
   End If
End Function

Private Function PointerToStringW(lpStringW As Long) As String
   Dim Buffer() As Byte
   Dim nLen As Long
   
   If lpStringW Then
      nLen = lstrlenW(lpStringW) * 2
      If nLen Then
         ReDim Buffer(0 To (nLen - 1)) As Byte
         CopyMemory Buffer(0), ByVal lpStringW, nLen
         PointerToStringW = Buffer
      End If
   End If
End Function



