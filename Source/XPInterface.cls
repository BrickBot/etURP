VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XPInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'       XPStyle Interface

Private XPMsgBoxLib As New ClsDialogBox
Private XPManifestLib As New ClsXPManifest


Public Sub AboutXP()
    XPMsgBoxLib.AboutBox , , , , , , , , , , "XP Controls"
End Sub

Public Function XPInputBox(msgText As String, MsgTitle As String)
    XPInputBox = XPMsgBoxLib.InputBoxXP(msgText, MsgTitle)
End Function

Public Function XPMsgBox(msgText As String, MsgTitle As String, FontBold As Boolean, MsgType As MsgXPType, MsgIcon As MsgIconType, Optional MsgBOT As String, Optional MsgBTT As String, Optional MsgBThT As String)
Debug.Print "MSGBOX - " & Chr(32) & msgText & Chr(32)
    XPMsgBox = XPMsgBoxLib.MsgBoxXP(msgText, MsgTitle, MsgType, MsgIcon, , , , FontBold, MsgBOT, MsgBTT, MsgBThT)
End Function

Public Sub XPMsgExitSystem()
    XPMsgBoxLib.MsgExitSystem
End Sub

Public Sub XPChangeMsgBoxBackColour(r, G, B)
    On Error Resume Next
    XPMsgBoxLib.ChangeMsgBackColor RGB(r, G, B)
End Sub

Public Sub XPShutdown()
    XPManifestLib.Shutdown
End Sub

Public Sub XPCloseCDDoor()
    XPManifestLib.CloseCDRoom
End Sub

Public Sub XPOpenCDDoor()
    XPManifestLib.OpenCDRoom
End Sub

Public Function XPGetUserName()
    XPGetUserName = XPManifestLib.Get_UserName
End Function

Public Function XPGetComputerName()
    XPGetComputerName = XPManifestLib.GetCommName
End Function

Public Function XPGetSystemPath()
    XPGetSystemPath = XPManifestLib.GetSysPath
End Function

Public Sub XPOpenSystemInfo()
    XPManifestLib.SysInfo
End Sub

Private Sub Class_Initialize()
    Set XPMsgBoxLib = New ClsDialogBox
    Set XPManifestLib = New ClsXPManifest
End Sub
