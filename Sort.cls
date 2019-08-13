VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'~ http://www.iti.fh-flensburg.de/lang/algorithmen/sortieren/shell/shell.htm
Dim a() As Variant
Dim n As Long
Dim hi As Long
Dim lo As Long
Public Enum SSOF_COMPARISON
    SSOF_NUMERICAL = 0
    SSOF_STRING = 1
End Enum
Dim str As New Strings
'Public Sub SortOnField(sMethod As String, aVec As Variant, ByVal nField As Integer, ByVal sFieldSep As Variant, Optional ByVal nComparisonType As SSOF_COMPARISON = SSOF_NUMERICAL)
'    a = aVec
'    Select Case LCase$(sMethod)
'    Case "shellsort"
'        n = UBound(a) + 1
'        shell_sort_on_field nField, sFieldSep, nComparisonType
'    Case "quicksort"
'        lo = LBound(aVec)
'        hi = UBound(aVec)
'        quick_sort_on_field nField, sFieldSep, nComparisonType
'    Case "heapsort"
'        n = UBound(a) + 1
'        heap_sort_on_field nField, sFieldSep, nComparisonType
'    End Select
'End Sub

Public Sub shellSortOnField(aVec As Variant, ByVal nField As Integer, ByVal sFieldSep As Variant, Optional ByVal nComparisonType As SSOF_COMPARISON = SSOF_NUMERICAL)
    a = aVec
    n = UBound(a) + 1
    shell_sort_on_field nField, sFieldSep, nComparisonType
End Sub


Sub shellsort(A0 As Variant, Optional bAscending As Boolean = True)
    a = A0
    n = UBound(a) + 1
    shell_sort bAscending
End Sub

Public Sub quicksort(aVec As Variant) ', Optional bAscending As Boolean = True)
    a = aVec
    lo = LBound(aVec)
    hi = UBound(aVec)
'    bAsc = bAscending
    Quick_Sort a, lo, hi
End Sub

Public Sub heapsort(aVec As Variant) ', Optional bAscending As Boolean = True)
    a = aVec
    n = UBound(a) + 1
'    bAsc = bAscending
    heap_sort
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

Private Sub shell_sort_on_field(nField As Integer, sDelim As Variant, nType As SSOF_COMPARISON)
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
                If nType = SSOF_NUMERICAL Then
                    If Val(str.NthField(a(j - h), sDelim, nField)) > Val(str.NthField(t, sDelim, nField)) Then
                        Exit Do
                    End If
                Else
                    If str.NthField(a(j - h), sDelim, nField) > str.NthField(t, sDelim, nField) Then
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
            If a(w + 1) > a(w) Then
                w = w + 1
            End If
        End If

        If a(v) >= a(w) Then
            Exit Sub  '// v has heap property
        End If
        exchange v, w  '// exchange labels of v and w
        v = w          '// continue
        w = 2 * v + 1
    Loop
End Sub

Private Sub exchange(i As Long, j As Long)
    Dim t As Variant
    t = a(i)
    a(i) = a(j)
    a(j) = t
End Sub


Private Sub Quick_Sort(vec As Variant, loBound As Long, hiBound As Long)
  Dim pivot, loSwap, hiSwap, temp


  '== This procedure is adapted from the algorithm given in:
  '==    Data Abstractions & Structures using C++ by
  '==    Mark Headington and David Riley, pg. 586
  '== Quicksort is the fastest array sorting routine for
  '== unordered arrays.  Its big O is  n log n


    '~ http://4guysfromrolla.com/webtech/012799-2.shtml


  '== Two items to sort
  If hiBound - loBound = 1 Then
    If vec(loBound) > vec(hiBound) Then
      temp = vec(loBound)
      vec(loBound) = vec(hiBound)
      vec(hiBound) = temp
    End If
  End If


  '== Three or more items to sort
  pivot = vec(Int((loBound + hiBound) / 2))
  vec(Int((loBound + hiBound) / 2)) = vec(loBound)
  vec(loBound) = pivot
  loSwap = loBound + 1
  hiSwap = hiBound


  Do
    '== Find the right loSwap
    While loSwap < hiSwap And vec(loSwap) <= pivot
      loSwap = loSwap + 1
    Wend
    '== Find the right hiSwap
    While vec(hiSwap) > pivot
      hiSwap = hiSwap - 1
    Wend
    '== Swap values if loSwap is less then hiSwap
    If loSwap < hiSwap Then
      temp = vec(loSwap)
      vec(loSwap) = vec(hiSwap)
      vec(hiSwap) = temp
    End If
  Loop While loSwap < hiSwap


  vec(loBound) = vec(hiSwap)
  vec(hiSwap) = pivot


  '== Recursively call function .. the beauty of Quicksort
    '== 2 or more items in first section
    If loBound < (hiSwap - 1) Then Call Quick_Sort(vec, loBound, hiSwap - 1)
    '== 2 or more items in second section
    If hiSwap + 1 < hiBound Then Call Quick_Sort(vec, hiSwap + 1, hiBound)


End Sub  'QuickSort

