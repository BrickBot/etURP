VERSION 5.00
Begin VB.Form RemoteControl 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RCX/Scout Remote"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   2730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      Caption         =   "3"
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      Caption         =   "2"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "1"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Beep"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "P5"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "P4"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "P3"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "P2"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "P1"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stop"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   2760
      Width           =   495
   End
End
Attribute VB_Name = "RemoteControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Enum Remote
  kRemoteKeysReleased = "$0000"
  kRemotePBMessage1 = "$0100"
  kRemotePBMessage2 = "$0200"
  kRemotePBMessage3 = "$0400"
  kRemoteOutAForward = "$0800"
  kRemoteOutBForward = "$1000"
  kRemoteOutCForward = "$2000"
  kRemoteOutABackward = "$4000"
  kRemoteOutBBackward = "$8000"
  kRemoteOutCBackward = "$0001"
  kRemoteSelProgram1 = "$0002"
  kRemoteSelProgram2 = "$0004"
  kRemoteSelProgram3 = "$0008"
  kRemoteSelProgram4 = "$0010"
  kRemoteSelProgram5 = "$0020"
  kRemoteStopOutOff = "$0040"
  kRemotePlayASound = "$0080"
End Enum

Private Sub Command1_Click()
SendRemote kRemoteStopOutOff
End Sub

Private Sub Command10_Click()
SendRemote kRemotePBMessage3
End Sub

Private Sub Command2_Click()
SendRemote kRemoteSelProgram1
End Sub

Sub SendRemote(RCommand As Remote)
closecom
ShellExecute Me.hwnd, "Open", "nqc.exe", " -SCOM" & about.Spirit.ComPortNo & " -remote " & RCommand & " 1", App.Path & "\bin\", 0
REOPENCOMWHENDONE
End Sub

Function REOPENCOMWHENDONE()
a = Format(Now, "ss")
While B - a < 2
B = Format(Now, "ss")
DoEvents
Wend

While (a > 0) Or (B > 0)
a = FindWindow("", "NQC")
B = FindWindow("", "nqc")
DoEvents
Wend

about.Spirit.InitComm
End Function

Function closecom()
about.Spirit.CloseComm
End Function

Function FindWindow(ByVal sClassName As String, ByVal sWindowName As String) As Long
If Len(sClassName) = 0 Then
   xHwnd = M_FindWindow(0&, sWindowName)
ElseIf Len(sWindowName) = 0 Then
   xHwnd = M_FindWindow(sClassName, 0&)
Else
   xHwnd = M_FindWindow(sClassName, sWindowName)
End If
FindWindow = xHwnd
End Function

Private Sub Command3_Click()
SendRemote kRemoteSelProgram2
End Sub

Private Sub Command4_Click()
SendRemote kRemoteSelProgram3
End Sub

Private Sub Command5_Click()
SendRemote kRemoteSelProgram4
End Sub

Private Sub Command6_Click()
SendRemote kRemoteSelProgram5
End Sub

Private Sub Command7_Click()
SendRemote kRemotePlayASound
End Sub

Private Sub Command8_Click()
SendRemote kRemotePBMessage1
End Sub

Private Sub Command9_Click()
SendRemote kRemotePBMessage2
End Sub
