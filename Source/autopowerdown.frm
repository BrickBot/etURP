VERSION 5.00
Begin VB.Form PowerDownTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Power Down Time"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2985
   Icon            =   "autopowerdown.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   2985
   StartUpPosition =   3  'Windows Default
   Begin etUCP.chameleonButton chameleonButton1 
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Force Set"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "autopowerdown.frx":030A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox ASet 
      BackColor       =   &H00C0C0C0&
      Caption         =   "AutoSet"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   240
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.VScrollBar PDT 
      Height          =   975
      Left            =   240
      Max             =   60
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.Label PowerDown 
      BackColor       =   &H00C0C0C0&
      Caption         =   "0 (Indefinite)"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "PowerDownTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ASet_Click()
    If ASet.Value = 1 Then
        chameleonButton1.Caption = "Force Set"
    Else
        chameleonButton1.Caption = "Set PDT"
    End If
End Sub

Private Sub chameleonButton1_Click()
    About.Spirit.PBPowerdownTime PDT.Value
End Sub

Private Sub Form_Load()
LogText "Load - PWRDWNTIME"
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
    DrawXPCtl Me
End Sub

Private Sub PDT_Change()
    PowerDown.Caption = PDT.Value & " (Minutes)"
    If PDT.Value = 1 Then PowerDown.Caption = "1 (Minute)"
    If PDT.Value = 0 Then PowerDown.Caption = "0 (Indefinite)"
    If ASet.Value = 1 Then
        About.Spirit.PBPowerdownTime PDT.Value
    End If
End Sub

