VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XPPanels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public HasPicture                       As Boolean
Public Icon                             As Integer
Public Psate                            As P_state
Public Caption                          As String
Public key                              As String
Enum P_state
    Opened = 0
    Closed = 1
    fixed = 2
End Enum
Public ppicture                         As Variant
Public PicturePath                      As String

