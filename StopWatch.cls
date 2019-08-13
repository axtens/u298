VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private dThen As Date, dNow As Date
Private bTicking As Boolean

Private Sub Class_Initialize()
    bTicking = False
End Sub

Public Sub Start()
    If bTicking = False Then
        bTicking = True
        dThen = Now
    Else
        MsgBox "StopWatch already ticking"
    End If
End Sub

Public Sub Finish()
    If bTicking = True Then
        bTicking = False
        dNow = Now
    Else
        MsgBox "StopWatch not ticking"
    End If
End Sub

Public Function StartTime() As Date
    If bTicking = True Then
        StartTime = dThen
    Else
        MsgBox "StopWatch not started"
        StartTime = Empty
    End If
End Function

Public Function FinishTime() As Date
    If bTicking = False Then
        FinishTime = dNow
    Else
        MsgBox "StopWatch still ticking"
        FinishTime = Empty
    End If
End Function

Public Function Seconds() As Integer
    If bTicking = False Then
        Seconds = DateDiff("s", dThen, dNow)
    Else
        MsgBox "StopWatching still ticking"
        Seconds = Empty
    End If
End Function

Public Function LapTime() As String
    Dim dTemp As Date, nHour As Integer, nMin As Integer, nSec As Integer
    If bTicking = True Then
        dTemp = DateAdd("s", DateDiff("s", dThen, Now), Empty)
        nHour = Hour(dTemp)
        nMin = Minute(dTemp)
        nSec = Second(dTemp)
        LapTime = nHour & "h " & nMin & "m " & nSec & "s"
    Else
        MsgBox "StopWatch not ticking"
        LapTime = Empty
    End If
End Function
