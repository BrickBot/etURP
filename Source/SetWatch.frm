VERSION 5.00
Begin VB.Form SetWatch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set RCX Watch"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3810
   Icon            =   "SetWatch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin etUCP.chameleonButton Command2 
      Height          =   495
      Left            =   1920
      TabIndex        =   10
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Set To Selected Time"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "SetWatch.frx":030A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Command1 
      Height          =   495
      Left            =   1920
      TabIndex        =   9
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Set To System Time"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "SetWatch.frx":0326
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.VScrollBar VScroll4 
      Height          =   855
      Left            =   840
      Max             =   9
      TabIndex        =   3
      Top             =   240
      Width           =   255
   End
   Begin VB.VScrollBar VScroll3 
      Height          =   855
      Left            =   600
      Max             =   9
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   855
      Left            =   360
      Max             =   9
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   855
      Left            =   120
      Max             =   9
      TabIndex        =   0
      Top             =   240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Left            =   1150
      Top             =   480
      Width           =   700
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   555
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "0"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   555
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   530
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   550
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "0"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   550
      Width           =   255
   End
End
Attribute VB_Name = "SetWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function M_FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer
Dim xHwnd           As Long

Function closecom()
    About.Spirit.CloseComm
    SetWatch.Enabled = False
End Function

Function REOPENCOMWHENDONE()
    a = Format(Now, "ss")

    While B - a < 5
        B = Format(Now, "ss")
    Wend


    While (a > 0) Or (B > 0)
        a = FindWindow("", "NQC")
        B = FindWindow("", "nqc")
    Wend

    About.Spirit.InitComm
    SetWatch.Enabled = True
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


Private Sub Command1_Click()
    HoursTemp = Hour(Now)
    MinsTemp = Minute(Now)
    VScroll1.Value = Int(Mid(HoursTemp, 1, 1))
    VScroll2.Value = Int(Mid(HoursTemp, 2, 1))
    VScroll3.Value = Int(Mid(MinsTemp, 1, 1))
    VScroll4.Value = Int(Mid(MinsTemp, 2, 1))
    About.Spirit.SetWatch Hour(Now), Minute(Now)
End Sub

Private Sub Command2_Click()
    About.Spirit.SetWatch VScroll1.Value & VScroll2.Value, VScroll3.Value & VScroll4.Value
End Sub

Private Sub Form_Load()
LogText "Load - SETWATCH"
    
    On Error Resume Next
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Private Sub VScroll1_Change()
    Label1.Caption = VScroll1.Value
End Sub

Private Sub VScroll2_Change()
    Label2.Caption = VScroll2.Value
End Sub

Private Sub VScroll3_Change()
    Label3.Caption = VScroll3.Value
End Sub

Private Sub VScroll4_Change()
    Label4.Caption = VScroll4.Value
End Sub
