VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Domains"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'http://www.windowsitpro.com/WindowsScripting/Article/ArticleID/4734/4734.html
'----------------------------------------------------------------
' Function:     GetAvailableNTDomains
' Description:  Returns a listing of available NT Domains
' Parameters:   None
' Returns:      Variant array of NT Domain names.
'----------------------------------------------------------------
'Public Function GetAvailableNTDomains() As Variant()
'  Dim objIADsContainer      As ActiveDs.IADsContainer
'  Dim objIADsDomain         As ActiveDs.IADsDomain
'  Dim vReturn() As Variant               ' Variant
'
'  ReDim vReturn(0)
'
'  Set objIADsContainer = GetObject("WinNT:")
'  For Each objIADsDomain In objIADsContainer
'    If Trim(vReturn(0)) <> "" Then
'      ReDim Preserve vReturn(UBound(vReturn) + 1)
'    End If
'    vReturn(UBound(vReturn)) = objIADsDomain.Name
'  Next
'
'  GetAvailableNTDomains = vReturn
'End Function

