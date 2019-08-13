VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComplexNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarReal As Double 'local copy
Private mvarImag As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarJayFormat As Boolean 'local copy
Private mvarSpaceBefore As Boolean 'local copy
Private mvarSpaceAfter As Boolean 'local copy
Private mvarTrailingImaginary As Boolean 'local copy
'local variable(s) to hold property value(s)
'local variable(s) to hold property value(s)
Private mvarTruncation As Integer 'local copy
Public Property Let Truncation(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Truncation = 5
    mvarTruncation = vData
End Property


Public Property Get Truncation() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Truncation
    Truncation = mvarTruncation
End Property




Public Property Let TrailingImaginary(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TrailingImaginary = 5
    mvarTrailingImaginary = vData
End Property

Public Property Get TrailingImaginary() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TrailingImaginary
    TrailingImaginary = mvarTrailingImaginary
End Property

Public Property Let SpaceAfter(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SpaceAfter = 5
    mvarSpaceAfter = vData
End Property

Public Property Get SpaceAfter() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SpaceAfter
    SpaceAfter = mvarSpaceAfter
End Property

Public Property Let SpaceBefore(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SpaceBefore = 5
    mvarSpaceBefore = vData
End Property

Public Property Get SpaceBefore() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SpaceBefore
    SpaceBefore = mvarSpaceBefore
End Property

Public Property Let JayFormat(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.JayFormat = 5
    mvarJayFormat = vData
End Property

Public Property Get JayFormat() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.JayFormat
    JayFormat = mvarJayFormat
End Property

Public Property Let Imag(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Imag = 5
    mvarImag = vData
End Property

Public Property Get Imag() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Imag
    Imag = mvarImag
End Property

Public Property Let Real(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Real = 5
    mvarReal = vData
End Property

Public Property Get Real() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Real
    Real = mvarReal
End Property

Private Sub Class_Initialize()
    mvarReal = 0
    mvarImag = 0
    mvarJayFormat = False
    mvarSpaceBefore = False
    mvarSpaceAfter = False
    mvarTrailingImaginary = True
    mvarTruncation = -1
End Sub
