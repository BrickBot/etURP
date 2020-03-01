VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TabCaption           As String
Public TabLeft              As Long
Public TabTop               As Long
Public TabWidth             As Long
Public TabHeight            As Long
Public TabPicture           As StdPicture

Public Property Get TabIcon() As StdPicture
    Set TabIcon = TabPicture
End Property

Public Property Set TabIcon(ByVal NewTabPicture As StdPicture)
    Set TabPicture = NewTabPicture
End Property

Private Sub Class_Terminate()
    Set TabPicture = Nothing
End Sub
