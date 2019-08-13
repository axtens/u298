VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "System"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim stdS As New Strings
Dim stdF As New Files
Dim stdR As New Registry

Dim oFSO As Scripting.FileSystemObject

Const TEMPORARY_FOLDER As Integer = 2
Private Declare Function GetSystemDefaultUILanguage Lib "kernel32" () As Long
Private Declare Function GetThreadLocale Lib "kernel32" () As Long
Private Declare Sub SleepMSec Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 260
End Type

Private Declare Function RegisterServiceProcess Lib "kernel32" (ByVal ProcessID As Long, ByVal ServiceFlags As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

'http://vb.mvps.org/samples/project.asp?id=console
Private Declare Function GetStdOutHandle Lib "kernel32" _
        Alias "GetStdHandle" _
       (Optional ByVal HandleType As Long = -11) As Long

Private Declare Function WriteFile Lib "kernel32" _
       (ByVal hFile As Long, _
        ByVal lpBuffer As String, _
        ByVal cToWrite As Long, _
        ByRef cWritten As Long, _
        Optional ByVal lpOverlapped As Long) As Long

Private Const SPI_GETWORKAREA = 48

Private Type RECT
  left As Long
  top As Long
  right As Long
  bottom As Long
End Type

Private Declare Function SystemParametersInfo Lib "user32" _
  Alias "SystemParametersInfoA" (ByVal uAction As Long, _
  ByVal uParam As Long, ByRef lpvParam As Any, _
  ByVal fuWinIni As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
  (ByVal hWnd As Long, ByVal lpOperation As String, _
  ByVal lpFile As String, ByVal lpParameters As String, _
  ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function ShellExecuteForExplore Lib "shell32.dll" Alias "ShellExecuteA" _
  (ByVal hWnd As Long, ByVal lpOperation As String, _
  ByVal lpFile As String, lpParameters As Any, _
  lpDirectory As Any, ByVal nShowCmd As Long) As Long

Public Enum EShellShowConstants
    essSW_HIDE = 0
    essSW_MAXIMIZE = 3
    essSW_MINIMIZE = 6
    essSW_SHOWMAXIMIZED = 3
    essSW_SHOWMINIMIZED = 2
    essSW_SHOWNORMAL = 1
    essSW_SHOWNOACTIVATE = 4
    essSW_SHOWNA = 8
    essSW_SHOWMINNOACTIVE = 7
    essSW_SHOWDEFAULT = 10
    essSW_RESTORE = 9
    essSW_SHOW = 5
End Enum

Public Enum EPROCESS_PRIORITY
    EP_IDLE = 64
    EP_BELOW_NORMAL = 16384
    EP_NORMAL = 32
    EP_ABOVE_NORMAL = 32768
    EP_HIGH_PRIORITY = 128
    EP_REAL_TIME = 256
End Enum

Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&
Private Const SE_ERR_ACCESSDENIED = 5        ' access denied
Private Const SE_ERR_ASSOCINCOMPLETE = 27
Private Const SE_ERR_DDEBUSY = 30
Private Const SE_ERR_DDEFAIL = 29
Private Const SE_ERR_DDETIMEOUT = 28
Private Const SE_ERR_DLLNOTFOUND = 32
Private Const SE_ERR_FNF = 2                ' file not found
Private Const SE_ERR_NOASSOC = 31
Private Const SE_ERR_PNF = 3                ' path not found
Private Const SE_ERR_OOM = 8                ' out of memory
Private Const SE_ERR_SHARE = 26

Private Declare Function UuidCreate Lib "rpcrt4.dll" (pId As UUID) As Long
Private Declare Function UuidToString Lib "rpcrt4.dll" Alias "UuidToStringA" (uuidID As UUID, ppUuid As Long) As Long
Private Declare Function RpcStringFree Lib "rpcrt4.dll" Alias "RpcStringFreeA" (ppStringUuid As Long) As Long
Private Declare Function CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal nSize As Long) As Long

Private Type UUID
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4(8) As Byte
End Type

Public Function RTFtoText(ByVal sText As Variant) As Variant
    Dim sResult As Variant
    Load Form1
    On Error Resume Next
    Form1.RichTextBox1.TextRTF = sText
    sResult = Form1.RichTextBox1.Text
    If Err.Number <> 0 Then
        sResult = sText
    End If
    Unload Form1
    RTFtoText = sResult
End Function

Public Function ThreadLocale() As Long
    ThreadLocale = GetThreadLocale()
End Function

Public Function SystemDefaultUILanguage() As Long
    SystemDefaultUILanguage = GetSystemDefaultUILanguage
End Function


Sub Sleep(nMilli As Integer)
    SleepMSec nMilli
End Sub

Function DoEventsSeconds(nSeconds As Integer) As Integer
   Dim Start As Long
   Dim Finish As Long
   Start = Timer   ' Set start time.
   Do While Timer < Start + nSeconds
      DoEvents   ' Yield to other processes.
   Loop
   Finish = Timer   ' Set end time.
   DoEventsSeconds = Finish - Start   ' Calculate total time.
End Function

Sub DoEventsSeconds2(iSeconds As Integer)
    Dim i As Integer
    For i = 1 To iSeconds
        SleepMSec 1000
        DoEvents
    Next
End Sub

Function CaptureDOS(sCommand As Variant, Optional bSynch As Boolean = True)
    Dim sFolder As String
    Dim sName As String
    Dim sTempFile As String
    Dim sResult As String
    Dim oShell As Object
    
    Set oFSO = New Scripting.FileSystemObject
        
    Set oShell = CreateObject("WScript.Shell")
    
    sResult = vbNullString
    
    sFolder = oFSO.GetSpecialFolder(TEMPORARY_FOLDER)
    sName = oFSO.GetTempName
    sTempFile = oFSO.BuildPath(sFolder, sName)
    'Log "*CaptureDOS: sTempFile=" & sTempFile
    oShell.Run sCommand & "> " & stdS.DQ(sTempFile), 0, bSynch
    sResult = stdF.ReadFileE(sTempFile)
    'oFSO.DeleteFile sTempFile
    CaptureDOS = sResult
End Function

'Function SetProcessPriority(sProcess As String, nPriority As EPROCESS_PRIORITY) As Boolean
'    Dim sComputer As String
'    Dim oWMIService As Object
'    Dim cProcesses As SWbemObjectSet
'    Dim oProcess As SWbemObject
'    Dim bDone As Boolean
'
'    bDone = False
'    sComputer = "."
'    Set oWMIService = GetObject("winmgmts:\\" & sComputer & "\root\cimv2")
'
'    Set cProcesses = oWMIService.ExecQuery("Select * from Win32_Process Where Name = '" & sProcess & "'")
'    For Each oProcess In cProcesses
'        oProcess.SetPriority (nPriority)
'        bDone = True
'    Next
'    SetProcessPriority = bDone
'End Function

'Function GetProcessInformation(sProcess As Variant, sElementsCSV As Variant) As Variant
'    Dim sComputer As String
'    Dim oWMIService As Object
'    Dim cProcesses As SWbemObjectSet
'    Dim oProcess As SWbemObject
'    Dim bDone As Boolean
'    Dim aResult As Variant
'    Dim i As Integer
'    Dim j As Integer
'    Dim n As Integer
'    Dim aCSV As Variant
'    Dim aEmbedded As Variant
'    Dim sElem As String
'    Dim vItem As Variant
'
'    bDone = False
'    sComputer = "."
'    Set oWMIService = GetObject("winmgmts:\\" & sComputer & "\root\cimv2")
'
'    Set cProcesses = oWMIService.ExecQuery("Select * from Win32_Process Where Name = '" & sProcess & "'")
'    n = cProcesses.Count - 1
'
'    If n > -1 Then 'there are some
'        aCSV = Split(sElementsCSV, ",")
'
'        ReDim aResult(n, 1)
'        ReDim aEmbedded(UBound(aCSV))
'
'        i = 0
'        For Each oProcess In cProcesses
'            aResult(i, 0) = oProcess.ProcessID
'            For j = 0 To UBound(aCSV)
'                sElem = aCSV(j)
'                Select Case LCase$(sElem)
'                Case "caption": vItem = oProcess.Caption
'                Case "commandline": vItem = oProcess.CommandLine
'                Case "creationclassname": vItem = oProcess.creationclassname
'                Case "creationdate": vItem = oProcess.creationdate
'                Case "cscreationclassname": vItem = oProcess.cscreationclassname
'                Case "csname": vItem = oProcess.csname
'                Case "description": vItem = oProcess.Description
'                Case "executablepath": vItem = oProcess.executablepath
'                Case "executionstate": vItem = oProcess.executionstate
'                Case "handle": vItem = oProcess.Handle
'                Case "handlecount": vItem = oProcess.handlecount
'                Case "installdate": vItem = oProcess.installdate
'                Case "kernelmodetime": vItem = oProcess.kernelmodetime
'                Case "maximumworkingsetsize": vItem = oProcess.maximumworkingsetsize
'                Case "minimumworkingsetsize": vItem = oProcess.minimumworkingsetsize
'                Case "name": vItem = oProcess.Name
'                Case "oscreationclassname": vItem = oProcess.oscreationclassname
'                Case "osname": vItem = oProcess.osname
'                Case "otheroperationcount": vItem = oProcess.otheroperationcount
'                Case "othertransfercount": vItem = oProcess.othertransfercount
'                Case "pagefaults": vItem = oProcess.pagefaults
'                Case "pagefileusage": vItem = oProcess.pagefileusage
'                Case "parentprocessid": vItem = oProcess.parentprocessid
'                Case "peakpagefileusage": vItem = oProcess.peakpagefileusage
'                Case "peakvirtualsize": vItem = oProcess.peakvirtualsize
'                Case "peakworkingsetsize": vItem = oProcess.peakworkingsetsize
'                Case "priority": vItem = oProcess.Priority
'                Case "privatepagecount": vItem = oProcess.privatepagecount
'                Case "processid": vItem = oProcess.ProcessID
'                Case "quotanonpagedpoolusage": vItem = oProcess.quotanonpagedpoolusage
'                Case "quotapagedpoolusage": vItem = oProcess.quotapagedpoolusage
'                Case "quotapeaknonpagedpoolusage": vItem = oProcess.quotapeaknonpagedpoolusage
'                Case "quotapeakpagedpoolusage": vItem = oProcess.quotapeakpagedpoolusage
'                Case "readoperationcount": vItem = oProcess.readoperationcount
'                Case "readtransfercount": vItem = oProcess.readtransfercount
'                Case "sessionid": vItem = oProcess.sessionid
'                Case "status": vItem = oProcess.Status
'                Case "terminationdate": vItem = oProcess.terminationdate
'                Case "threadcount": vItem = oProcess.threadcount
'                Case "usermodetime": vItem = oProcess.usermodetime
'                Case "virtualsize": vItem = oProcess.virtualsize
'                Case "windowsversion": vItem = oProcess.windowsversion
'                Case "workingsetsize": vItem = oProcess.workingsetsize
'                Case "writeoperationcount": vItem = oProcess.writeoperationcount
'                Case "writetransfercount": vItem = oProcess.writetransfercount
'                End Select
'                aEmbedded(j) = vItem
'            Next
'            aResult(i, 1) = aEmbedded
'            i = i + 1
'        Next
'    Else
'        aResult = Empty
'    End If
'    GetProcessInformation = aResult
'End Function

'WatsonR@IntelligenCIA.com comp.lang.visual.basic 18/06/2008 8:43 AM
Public Function KillProcess(ProcessName As String) As Boolean
    Const PROCESS_ALL_ACCESS = 0
    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim exitCode As Long
    Dim myProcess As Long
    Dim AppKill As Boolean
    Dim appCount As Integer
    Dim i As Integer
    On Local Error GoTo Finish
    appCount = 0

    Const TH32CS_SNAPPROCESS As Long = 2&

    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    
    Do While rProcessFound
        i = InStr(1, uProcess.szexeFile, Chr(0))
        szExename = LCase$(left$(uProcess.szexeFile, i - 1))
        If right$(szExename, Len(ProcessName)) = LCase$(ProcessName) Then
            KillProcess = True
            appCount = appCount + 1
            myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
            AppKill = TerminateProcess(myProcess, exitCode)
            Call CloseHandle(myProcess)
        End If
        rProcessFound = ProcessNext(hSnapshot, uProcess)
    Loop
    Call CloseHandle(hSnapshot)
Finish:
End Function

Public Function GetTheWindowsDirectory() As String

    
    Dim strWindowsDir As String        ' Variable to return the path of Windows Directory
    Dim lngWindowsDirLength As Long    ' Variable to return the the lenght of the path
    
    strWindowsDir = Space(250)         ' Initilize the buffer to receive the string
    lngWindowsDirLength = GetWindowsDirectory(strWindowsDir, 250) ' Read the path of the windows directory
    strWindowsDir = left(strWindowsDir, lngWindowsDirLength) ' Extract the windows path from the buffer
    
    GetTheWindowsDirectory = strWindowsDir
    
End Function



'To keep things simple, wrap these two APIs in a single helper function like so:

Function ConsoleWrite(sText As String) As Long
    WriteFile GetStdOutHandle, ByVal sText, Len(sText), ConsoleWrite
End Function

Function ConsoleWriteLine(sText As String) As Long
    Dim sLine As String
    sLine = sText & vbNewLine
    WriteFile GetStdOutHandle, ByVal sLine, Len(sLine), ConsoleWriteLine
End Function


Function DesktopDimensions() As Variant
    Dim wide As Long
    Dim high As Long
    Dim aForm As New Form1
    wide = aForm.ScaleX(Screen.WIDTH, vbTwips, vbPixels)
    high = aForm.ScaleY(Screen.Height, vbTwips, vbPixels)
    DesktopDimensions = Array(wide, high)
End Function

Function AvailableDesktopDimensions() As Variant
    Dim myrect As RECT
    Call SystemParametersInfo(SPI_GETWORKAREA, 0&, myrect, 0&)
    AvailableDesktopDimensions = Array(myrect.left, myrect.top, myrect.right, myrect.bottom)
End Function

'http://www.vbaccelerator.com/codelib/shell/shellex.htm

'To open a the user's browser at a particular internet site
'    ShellEx "http://www.dogma.demon.co.uk", , , , , Me.hWnd
'
'To print a document
'    ShellEx "C:\My Documents\Music\Brown Paper Bag.doc", , , , "print", Me.hWnd
'
'To explore from a folder
'    ShellEx "C:\My Documents\Music", , , , "explore", Me.hWnd
'
'Start VB, minimised, load a project and run it
'    ShellEx _
'       "C:\Program Files\DevStudio\Vb\vb5.exe",_
'        essSW_SHOWMINNOACTIVE , _
'        "/run C:\Stevemac\VB\Develop\BmpTile\gBmpTile.vbg", _
'       , Me.hWnd
'

Public Function ShellEx( _
        ByVal sFile As String, _
        Optional ByVal eShowCmd As EShellShowConstants = essSW_SHOWDEFAULT, _
        Optional ByVal sParameters As String = "", _
        Optional ByVal sDefaultDir As String = "", _
        Optional sOperation As String = "open", _
        Optional Owner As Long = 0 _
    ) As Boolean
Dim lR As Long
Dim lErr As Long, sErr As Long
    If (InStr(UCase$(sFile), ".EXE") <> 0) Then
        eShowCmd = 0
    End If
    On Error Resume Next
    If (sParameters = "") And (sDefaultDir = "") Then
        lR = ShellExecuteForExplore(Owner, sOperation, sFile, 0, 0, essSW_SHOWNORMAL)
    Else
        lR = ShellExecute(Owner, sOperation, sFile, sParameters, sDefaultDir, eShowCmd)
    End If
    If (lR < 0) Or (lR > 32) Then
        ShellEx = True
    Else
        ' raise an appropriate error:
        lErr = vbObjectError + 1048 + lR
        Select Case lR
        Case 0
            lErr = 7: sErr = "Out of memory"
        Case ERROR_FILE_NOT_FOUND
            lErr = 53: sErr = "File not found"
        Case ERROR_PATH_NOT_FOUND
            lErr = 76: sErr = "Path not found"
        Case ERROR_BAD_FORMAT
            sErr = "The executable file is invalid or corrupt"
        Case SE_ERR_ACCESSDENIED
            lErr = 75: sErr = "Path/file access error"
        Case SE_ERR_ASSOCINCOMPLETE
            sErr = "This file type does not have a valid file association."
        Case SE_ERR_DDEBUSY
            lErr = 285: sErr = "The file could not be opened because the target application is busy. Please try again in a moment."
        Case SE_ERR_DDEFAIL
            lErr = 285: sErr = "The file could not be opened because the DDE transaction failed. Please try again in a moment."
        Case SE_ERR_DDETIMEOUT
            lErr = 286: sErr = "The file could not be opened due to time out. Please try again in a moment."
        Case SE_ERR_DLLNOTFOUND
            lErr = 48: sErr = "The specified dynamic-link library was not found."
        Case SE_ERR_FNF
            lErr = 53: sErr = "File not found"
        Case SE_ERR_NOASSOC
            sErr = "No application is associated with this file type."
        Case SE_ERR_OOM
            lErr = 7: sErr = "Out of memory"
        Case SE_ERR_PNF
            lErr = 76: sErr = "Path not found"
        Case SE_ERR_SHARE
            lErr = 75: sErr = "A sharing violation occurred."
        Case Else
            sErr = "An error occurred occurred whilst trying to open or print the selected file."
        End Select
               
        Err.Raise lErr, , App.ExeName & ".GShell", sErr
        ShellEx = False
    End If

End Function

Public Function GetUUID(ByRef sUuid As String) As Boolean
    Const RPC_S_OK As Long = 0
    Const SZ_UUID_LEN As Long = 36
    Dim uuidID As UUID
    Dim sUid As String
    Dim ppUuid As Long
    
    sUid = String(SZ_UUID_LEN, 0)
    If UuidCreate(uuidID) = RPC_S_OK Then
        If UuidToString(uuidID, ppUuid) = RPC_S_OK Then
            CopyMemory ByVal sUid, ByVal ppUuid, SZ_UUID_LEN
            If RpcStringFree(ppUuid) = RPC_S_OK Then
                sUuid = sUid
                GetUUID = True
            End If
        End If
    End If
End Function

Public Function CreateGUID() As String
    Dim sResult As String
    If GetUUID(sResult) Then
        CreateGUID = sResult
    Else
        CreateGUID = vbNullString
    End If
End Function

Public Sub Navigate(ByVal NavTo As String)
  Dim hBrowse As Long
  hBrowse = ShellExecute(0&, "open", NavTo, "", "", essSW_SHOW)
End Sub

Public Sub Cooperate()
    DoEvents
End Sub

Public Sub StopRun()
    End
End Sub

Private Sub Class_Initialize()
    stdR.Needs "WScript.Shell"
    If Err.Number = vbObjectError + 999 Then
        MsgBox "Class 'System' problem: " & Err.Description
        End
    End If

End Sub
