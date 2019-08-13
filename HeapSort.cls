VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HeapSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'~ http://www.iti.fh-flensburg.de/lang/algorithmen/sortieren/heap/heapen.htm
Dim aVector() As Variant
Dim n As Long
'Dim bAsc As Boolean

Public Sub heapsort(aVec As Variant) ', Optional bAscending As Boolean = True)
    aVector = aVec
    n = UBound(aVector) + 1
'    bAsc = bAscending
    heap_sort
End Sub

Public Function sorted() As Variant()
    sorted = aVector
End Function

Private Sub heap_sort()
    buildheap
    Do While n > 1
        n = n - 1
        exchange 0, n
        downheap 0
    Loop
End Sub

Private Sub buildheap()
    Dim v As Long
    For v = Int(n / 2) - 1 To 0 Step -1
        downheap v
    Next
End Sub

Private Sub downheap(v As Long)
    Dim w As Long
    w = 2 * v + 1   '// first descendant of v
    Do While w < n
        If w + 1 < n Then   '// is there a second descendant?
            'If bAsc Then
                If aVector(w + 1) > aVector(w) Then
                    w = w + 1
                End If
            'Else
            '    If aVector(w + 1) <= aVector(w) Then
            '        w = w + 1
            '    End If
            'End If
        End If

        '// w is the descendant of v with maximum label
        'If bAsc Then
            If aVector(v) >= aVector(w) Then
                Exit Sub  '// v has heap property
            End If
        'Else
        '    If aVector(v) < aVector(w) Then
        '        Exit Sub  '// v has heap property
        '    End If
        'End If
        '// otherwise
        exchange v, w  '// exchange labels of v and w
        v = w          '// continue
        w = 2 * v + 1
    Loop
End Sub

Private Sub exchange(i As Long, j As Long)
    Dim t As Variant
    t = aVector(i)
    aVector(i) = aVector(j)
    aVector(j) = t
End Sub

