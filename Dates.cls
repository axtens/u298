VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim stdstr As New Strings

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer

    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer

    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer

End Type

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(31) As Integer

    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long

End Type

Private Declare Function GetTimeZoneInformation _
        Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Const TIME_ZONE_ID_INVALID& = &HFFFFFFFF
Private Const TIME_ZONE_ID_STANDARD& = 1
Private Const TIME_ZONE_ID_UNKNOWN& = 0

Private Const TIME_ZONE_ID_DAYLIGHT& = 2


Private Declare Function InternetTimeToSystemTime Lib "wininet.dll" _
       (ByVal lpszTime As String, _
        ByRef pst As SYSTEMTIME, _
        ByVal dwReserved As Long) _
        As Long

 
Function dbDate(vDate As Variant) As String
    '~ http://www.aspfaq.com/show.asp?id=2260
    Dim dDate As Date
    dDate = CDate(vDate)
    dbDate = Year(dDate) & "-" & _
             stdstr.Zerofill2(Month(dDate)) & "-" & _
             stdstr.Zerofill2(Day(dDate)) & " " & _
             stdstr.Zerofill2(Hour(dDate)) & ":" & _
             stdstr.Zerofill2(Minute(dDate)) & ":" & _
             stdstr.Zerofill2(Second(dDate))
End Function


Public Function GetGmtTime(Optional StartingDate As Variant) As Date

    'Parameters: StartingDate (Optional).  The function will figure
    'out GMT time based on StartingDate
    'If StartingDate is not provided, the current time will be used
    
    Dim Difference As Long

    
    Difference = GetTimeDifference()
    
    If IsMissing(StartingDate) Then
        'use current time
        GetGmtTime = DateAdd("s", -Difference, Now)
    Else
        'use StartingDate

        GetGmtTime = DateAdd("s", -Difference, StartingDate)
    End If
End Function

Public Function GetTimeDifference() As Long

    'Returns  the time difference between
    'local & GMT time in seconds.
    'If the  result is negative, your time zone
    'lags behind GMT zone.
    'If the  result is positive, your time zone is ahead.
    
    Dim tz As TIME_ZONE_INFORMATION
    Dim retcode As Long

    Dim Difference As Long
    
    'retrieve the time zone information
    retcode = GetTimeZoneInformation(tz)
    
    'convert to seconds

    Difference = -tz.Bias * 60
    'cache the result

    GetTimeDifference = Difference
    
    'if we are in daylight  saving time, apply the bias.
    If retcode = TIME_ZONE_ID_DAYLIGHT& Then

        If tz.DaylightDate.wMonth <> 0 Then
            'if tz.DaylightDate.wMonth = 0 then the daylight
            'saving time change doesn't occur
            GetTimeDifference = Difference - tz.DaylightBias * 60
        End If

    End If
    
End Function

Public Function GetTimeHere(gmtTime As Date) As Date

    'Parameters:    gmtTime - Provides the time & date
    'from which to make calculations
    'Returns the time in your local time zone
    'which corresposponds to GMT time
    
    Dim Differerence As Long

    
    Differerence = GetTimeDifference()
    GetTimeHere = DateAdd("s", Differerence, gmtTime)
    
End Function

Public Function InternetTimeToVbLocalTime(ByVal DateString As String) As Date

    'Currently we process 2 formats
    'Rfc822 and Iso8601
    
    'Iso8601 is either 1997-07-16T19:20:30+01:00 (25 bytes) or 1997-07-16T19:20:30Z (20 bytes)
    'Rfc822 is Tue, 23 Sep 2003 13:21:00 -07:00 (32 bytes) or Tue, 23 Sep 2003 13:21:00 GMT (29 bytes)
    
    'The key difference is that Iso8661 time has a latin letter T in position 11

    
    
    DateString = Trim$(DateString)
    
    If Mid$(DateString, 11, 1) = "T" Then
        InternetTimeToVbLocalTime = Iso8601TimeToLocalVbTime(DateString)
    Else
        InternetTimeToVbLocalTime = Rfc822TimeToLocalVbTime(DateString)
    End If

    
End Function

Private Function Iso8601TimeToLocalVbTime(ByVal sIso8601 As String) As Date

    'format of the time is similar to this: 1997-07-16T19:20:30+01:00
    'or                                     1997-07-16T19:20:30Z or 2003-10-09T09:40:46Z
    'where Z is UTC (aka GMT time)
    
    'formatting breakdown
    '                1012141618202224
    '       1997-07-16T19:20:30+01:00

    '       1234567891113151719212325
    
    Dim sYear As String
    Dim sMonth As String

    Dim sDay As String
    Dim sHour As String
    Dim sMinute As String

    Dim sSecond As String
    Dim sTimeZone As String
    Dim dtDateTime As Date

    Dim bSign As Boolean
    Dim dGMT As Long

    sYear = left$(sIso8601, 4)
    sMonth = Mid$(sIso8601, 6, 2)
    sDay = Mid$(sIso8601, 9, 2)
    sHour = Mid$(sIso8601, 12, 2)
    sMinute = Mid$(sIso8601, 15, 2)
    sSecond = Mid$(sIso8601, 18, 2)
    sTimeZone = Mid$(sIso8601, 20)
    
    dtDateTime = CDate(DateSerial(sYear, sMonth, sDay) & " " & _
                                     TimeSerial(sHour, sMinute, sSecond))
                                     
                                         
    'replace Z with +00:00 for easier processing
    sTimeZone = Replace(sTimeZone, "Z", "+00:00", , , vbTextCompare)
    
    'get the size
    bSign = IIF(left$(sTimeZone, 1) = "+", True, False)
    

    'grab the hour & minutes
    dGMT = Val(Mid$(sTimeZone, Len(sTimeZone) - 3, 2)) + (CInt(right$(sTimeZone, 2)) * 100 / 60)
    
    If bSign Then
        dtDateTime = DateAdd("H", -dGMT, dtDateTime)
    Else

        dtDateTime = DateAdd("H", dGMT, dtDateTime)
    End If
    
    Iso8601TimeToLocalVbTime = GetTimeHere(dtDateTime)
    
    
    
End Function

Private Function Rfc822TimeToLocalVbTime(sRfc822 As String) As Date

    Dim uSystemTime As SYSTEMTIME
    Dim sWWW As String
    Dim iHours As Integer

    Dim dGMT As Long
    Dim sHourDifferential As String
    Dim dtDateTime As Date

    
    Dim sSign As String
    Dim bSign As Boolean
    Dim sEscapedTime As String

    Dim sTimeZoneString As String
    Dim iPos As Integer
    'true = positive

    'false = negative
    
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    sWWW = sRfc822
    
    If InStr(1, sWWW, "GMT", vbTextCompare) > 0 Then
        sWWW = Replace(sWWW, "GMT", "+0000")
    End If

    'check to make sure that the time zone is included

    If Len(Trim$(sWWW)) = 25 Then
        'add time zone
        sWWW = sWWW & " +0000"
    End If

    
    If (InStr(1, sWWW, ",") = 0) Then sWWW = "Thu, " & sWWW
    Call InternetTimeToSystemTime(sWWW, uSystemTime, 0&)
    With uSystemTime
        dtDateTime = CDate(DateSerial(.wYear, .wMonth, .wDay) & " " & _
                                     TimeSerial(.wHour, .wMinute, .wSecond))
    End With

    
    'get the sign from the back end
    'remove colons, in case the time is 07:00 instead of 0700
    sEscapedTime = Replace(sWWW, ":", "")
    
    sSign = Mid$(sEscapedTime, Len(sEscapedTime) - 4, 1)
    bSign = IIF(sSign = "-", False, True)
    
    'grab the hour & minutes

    iPos = InStrRev(sWWW, " ")
    If iPos > 0 Then
        'get rid of the space and the +/- sign
        sTimeZoneString = Mid$(sWWW, iPos + 2)
        
        'escape it
        sTimeZoneString = Replace(sTimeZoneString, ":", "")
        sTimeZoneString = Replace(sTimeZoneString, " ", "")
        

        'at this point we should have the following: 0700
        dGMT = Val(left$(sTimeZoneString, 2)) + Val(right$(sWWW, 2)) * 100 / 60
        
        'dGMT = Val(Mid$(sWWW, Len(sWWW) - 3, 2)) + (CInt(Right$(sWWW, 2)) * 100 / 60)
    Else
        dGMT = 0
    End If

        
        
    If bSign Then
        dtDateTime = DateAdd("H", -dGMT, dtDateTime)
    Else
        dtDateTime = DateAdd("H", dGMT, dtDateTime)
    End If
    
    Rfc822TimeToLocalVbTime = GetTimeHere(dtDateTime)
    

End Function

