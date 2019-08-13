VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Schedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ---------------------------------------------------------------
' Copyright ©2000 - 2007 Veign, LLC All rights reserved
' ---------------------------------------------------------------
' Distribution: You can freely use this code in your own
' applications, but you many not reproduce
' or publish this code on any web site,
' online service, or distribute as source on
' any media without express permission from Veign.
' ---------------------------------------------------------------

Option Explicit

Private Declare Function GetComputerName Lib "kernel32.dll" _
        Alias "GetComputerNameA" _
        (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function NetScheduleJobAdd Lib "netapi32.dll" _
        (ByVal servername As String, Buffer As Any, _
        Jobid As Long) As Long
Private Declare Function NetScheduleJobDel Lib "netapi32.dll" _
        (ByVal servername As String, ByVal MinJobId As Long, _
        ByVal MaxJobId As Long) As Long


' Schedule structure
Private Type AT_INFO
    JobTime     As Long
    DaysOfMonth As Long
    DaysOfWeek  As Byte
    Flags       As Byte
    dummy       As Integer
    Command     As String
End Type

'Schedule constants
'If you set this flag, the job runs, and
'continues to run, on each day for which a
'corresponding bit is set in the DaysOfMonth
'or DaysOfWeek member. The job is not deleted
'after it executes.
'If this flag is clear, the job runs only once
'for each bit set in these members. The job is
'deleted after it executes once.
Const JOB_RUN_PERIODICALLY = &H1

'If you set this flag, the job executes at the
'first occurrence of JobTime at the computer
'where the job is queued.
'Setting this flag is equivalent to setting
'the bit for the current day in the DaysOfMonth
'member.
Const JOB_ADD_CURRENT_DATE = &H8

'If you set this flag, the job does not run interactively.
'If this flag is clear, the job runs interactively.
Const JOB_NONINTERACTIVE = &H10

Const NERR_Success = 0

'Day Constants
Public Enum ENUM_WEEKDAYS
    enMonday = 1
    enTuesday = 2
    enWednesday = 4
    enThursday = 8
    enFriday = 16
    enSaturday = 32
    enSunday = 64
End Enum


Public Function AddTask(strTime As Variant, strCommand As Variant, _
                        Optional enDaysInWeek As ENUM_WEEKDAYS = -1, _
                        Optional strDaysInMonth As Variant = "", _
                        Optional RunInteractive As Boolean = True, _
                        Optional ReOccuring As Boolean = True) As Long

On Error GoTo Hell

'Default return value (failure)
AddTask = -1

'Convert the computer name to unicode
Dim strComputerName As String
strComputerName = StrConv(GetNameOfComputer(), vbUnicode)

'Setup Task Structure
Dim udtAtInfo As AT_INFO
With udtAtInfo
    .JobTime = (Hour(CDate(strTime)) * 3600 + Minute(CDate(strTime)) * 60) * 1000

    'Set the task period
    If enDaysInWeek > -1 Then
        'Get the days of the week from the constants
        .DaysOfWeek = enDaysInWeek
    ElseIf Len(strDaysInMonth) > 0 Then
        'Loop through all days of the month passed
        ' and set the bits
        Dim i As Long
        Dim strDates() As String
        strDates = Split(strDaysInMonth, ",")
        For i = 0 To UBound(strDates)
            .DaysOfMonth = .DaysOfMonth + 2 ^ (strDates(i) - 1)
        Next
    End If

    'Set Flags
    If ReOccuring Then
        .Flags = JOB_RUN_PERIODICALLY
    End If

    If Not RunInteractive Then
        .Flags = .Flags Or JOB_NONINTERACTIVE
    End If

    'Set the command to run
    .Command = StrConv(strCommand, vbUnicode)
End With

'Schedule
Dim lngJobID As Long
Dim retval As Long
retval = NetScheduleJobAdd(strComputerName, udtAtInfo, lngJobID)

'Check for success
If retval = NERR_Success Then
    AddTask = lngJobID
Else
    AddTask = -1 * retval
End If

Exit_For:
    On Error GoTo 0
    Exit Function

Hell:
     GoTo Exit_For

End Function

Public Function DeleteTask(lngID As Long) As Boolean

On Error GoTo Hell

'Convert the computer name to unicode
Dim strComputerName As String
strComputerName = StrConv(GetNameOfComputer(), vbUnicode)

'Delete Task
Dim retval As Long
retval = NetScheduleJobDel(strComputerName, lngID, lngID)

'Return success
If retval = NERR_Success Then
    DeleteTask = True
Else
    DeleteTask = False
End If

Exit_For:
    On Error GoTo 0
    Exit Function

Hell:
     GoTo Exit_For

End Function

Public Function GetNameOfComputer() As String

Dim retval As Long

'Create a string buffer for the computer name
Dim strCompName As String
strCompName = Space(255)

'Retrieve the Computer name
retval = GetComputerName(strCompName, 255)

'Remove the trailing null character from the string
GetNameOfComputer = left(strCompName, InStr(strCompName, vbNullChar) - 1)

End Function




