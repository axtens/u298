VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Collections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private stdV As New Variable
Public AddOnly As Boolean
Private mCol As Collection


Public Function add(vData As Variant, Optional sKey As Variant) As Variable
    Dim objNewMember As Variable
    'create a new object
    Set objNewMember = New Variable

    'set the properties passed into the method
    objNewMember.Value = vData
    objNewMember.Name = sKey
    If Exists(sKey) Then
        If Not AddOnly Then
            Modify vData, sKey
        End If
    Else
        If IsMissing(sKey) Then
            mCol.add objNewMember
        Else
            mCol.add objNewMember, sKey
        End If
    End If
    'return the object created
    Set Define = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As Variable
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    Set Item = mCol(vntIndexKey)
End Property

Public Function Exists(vntIndexKey As Variant) As Boolean
    Dim v As Variable
    On Error Resume Next
    Set v = mCol(vntIndexKey)
    Exists = (Err.Number = 0)
End Function

Public Function Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Function


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub

Public Sub Modify(sValue As Variant, sKey As Variant)
    Dim v As Variable
    'Set V = New stdV
    Set v = mCol.Item(sKey)
    v.Value = sValue
    'mCol.Remove sKey
    'Add sValue, sKey
End Sub

Public Property Get NewEnum() As IUnknown
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property

Public Function Keys() As Variant
    Dim v As Variable
    Dim i As Long
    Dim aResult() As Variant
    ReDim aResult(Count() - 1)
    i = 0
    For Each v In mCol
        aResult(i) = v.Name
        i = i + 1
    Next
    Keys = aResult
End Function


Public Function Items() As Variant
    Dim v As Variable
    Dim i As Long
    Dim aResult() As Variant
    ReDim aResult(Count() - 1)
    i = 0
    For Each v In mCol
        aResult(i) = v.Value
        i = i + 1
    Next
    Items = aResult
End Function


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
    AddOnly = True
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub


