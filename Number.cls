VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Number"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
     Option Explicit

      Private Const OFFSET_4 = 4294967296#
      Private Const MAXINT_4 = 2147483647
      Private Const OFFSET_2 = 65536
      Private Const MAXINT_2 = 32767

      Public Function UnsignedToLong(Value As Double) As Long
        If Value < 0 Or Value >= OFFSET_4 Then Error 6 ' Overflow
        If Value <= MAXINT_4 Then
          UnsignedToLong = Value
        Else
          UnsignedToLong = Value - OFFSET_4
        End If
      End Function

      Public Function LongToUnsigned(Value As Long) As Double
        If Value < 0 Then
          LongToUnsigned = Value + OFFSET_4
        Else
          LongToUnsigned = Value
        End If
      End Function

      Public Function UnsignedToInteger(Value As Long) As Integer
        If Value < 0 Or Value >= OFFSET_2 Then Error 6 ' Overflow
        If Value <= MAXINT_2 Then
          UnsignedToInteger = Value
        Else
          UnsignedToInteger = Value - OFFSET_2
        End If
      End Function

      Public Function IntegerToUnsigned(Value As Integer) As Long
        If Value < 0 Then
          IntegerToUnsigned = Value + OFFSET_2
        Else
          IntegerToUnsigned = Value
        End If
      End Function
                
Public Function DecimalToBinary(DecimalValue As Long, _
    MinimumDigits As Integer) As String

' Returns a string containing the binary
' representation of a positive integer.

Dim result As String
Dim ExtraDigitsNeeded As Integer

' Make sure value is not negative.
DecimalValue = Abs(DecimalValue)

' Construct the binary value.

Do
    result = CStr(DecimalValue Mod 2) & result
    DecimalValue = DecimalValue \ 2
Loop While DecimalValue > 0

' Add leading zeros if needed.
ExtraDigitsNeeded = MinimumDigits - Len(result)
If ExtraDigitsNeeded > 0 Then
    result = String(ExtraDigitsNeeded, "0") & result
End If

DecimalToBinary = result

End Function
