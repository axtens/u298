VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Drives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'from WatsonR@IntelligenCIA.com on comp.lang.visual.basic
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" _
 (ByVal nDrive As String) As Long
Const DRIVE_REMOVABLE = 2
Const DRIVE_REMOTE = 4
Const DRIVE_RAMDISK = 6
Const DRIVE_FIXED = 3
Const DRIVE_CDROM = 5

'syntax: DriveType("a:") or DriveType(dir1.path) etc..,

Function DriveType(ByVal sDrv As String) As String
    Dim sDrive As String
    Dim iDrive As Integer
    sDrive = left$(sDrv, 2)
    iDrive = GetDriveType(sDrive)
    Select Case iDrive
        Case DRIVE_REMOVABLE
            DriveType = "REMOVABLE"
        Case DRIVE_FIXED
            DriveType = "FIXED"
        Case DRIVE_REMOTE
            DriveType = "REMOTE"
        Case DRIVE_RAMDISK
            DriveType = "RAMDISK"
        Case DRIVE_CDROM
            DriveType = "CDROM"
        Case Else
            DriveType = "UNKNOWN"
    End Select
End Function

