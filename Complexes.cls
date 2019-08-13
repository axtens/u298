VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Complexes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Function Plus(ByVal c1 As Variant, ByVal c2 As Variant) As Variant
    Plus = Complex.toString(Complex.Plus(Complex.Parse(c1), Complex.Parse(c2)))
End Function

Public Function Minus(ByVal c1 As Variant, ByVal c2 As Variant) As Variant
    Minus = Complex.toString(Complex.Minus(Complex.Parse(c1), Complex.Parse(c2)))
End Function

Public Function Times(ByVal c1 As Variant, ByVal c2 As Variant) As Variant
    Times = Complex.toString(Complex.Times(Complex.Parse(c1), Complex.Parse(c2)))
End Function

Public Function Divide(ByVal c1 As Variant, ByVal c2 As Variant) As Variant
    Divide = Complex.toString(Complex.Divide(Complex.Parse(c1), Complex.Parse(c2)))
End Function

Public Function Sine(ByVal s As Variant) As Variant
    Sine = Complex.toString(Complex.Sine(Complex.Parse(s)))
End Function

Public Function Truncate(ByVal s As Variant, prec As Long) As Variant
    Truncate = Complex.toString(Complex.Truncate(Complex.Parse(s), prec))
End Function

Private Sub Class_Initialize()
    Complex.Init
End Sub
