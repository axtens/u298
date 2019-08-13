VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShellSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'~ http://www.iti.fh-flensburg.de/lang/algorithmen/sortieren/shell/shell.htm
Dim a() As Variant
Dim n As Integer

Sub shellsort(a0 As Variant, Optional bAscending As Boolean = True)
    a = a0
    n = UBound(a) + 1
    shell_sort bAscending
End Sub

Public Function sorted() As Variant()
    sorted = a
End Function

Private Sub shell_sort(bAscending As Boolean)

    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim h As Long
    Dim t As Variant
    Dim cols(15) As Long
    cols(0) = 1391376
    cols(1) = 463792
    cols(2) = 198768
    cols(3) = 86961
    cols(4) = 33936
    cols(5) = 13776
    cols(6) = 4592
    cols(7) = 1968
    cols(8) = 861
    cols(9) = 336
    cols(10) = 112
    cols(11) = 48
    cols(12) = 21
    cols(13) = 7
    cols(14) = 3
    cols(15) = 1
    
    For k = 0 To 15 ' UBound(cols)
        h = cols(k)
        
        i = h
        Do
            If i >= n Then
                Exit Do
            End If
            
            j = i
            t = a(i)
            Do
                If j < h Then
                    Exit Do
                End If
                If bAscending = True Then
                    If a(j - h) <= t Then
                        Exit Do
                    End If
                Else
                    If a(j - h) > t Then
                        Exit Do
                    End If
                End If
                a(j) = a(j - h)
                j = j - h
                
            Loop
            
            a(j) = t
            
            i = i + 1
        Loop
    Next
        
End Sub

