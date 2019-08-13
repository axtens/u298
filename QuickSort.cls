VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QuickSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim aVector() As Variant
Dim lo As Long
Dim hi As Long

Public Sub quicksort(aVec As Variant) ', Optional bAscending As Boolean = True)
    aVector = aVec
    lo = LBound(aVec)
    hi = UBound(aVec)
'    bAsc = bAscending
    Quick_Sort aVector, lo, hi
End Sub

Public Function sorted() As Variant()
    sorted = aVector
End Function

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
