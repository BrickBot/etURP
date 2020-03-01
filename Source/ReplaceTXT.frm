VERSION 5.00
Begin VB.Form ReplaceTXT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Replace"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   Icon            =   "ReplaceTXT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox CaseSens 
      Caption         =   "Case Sensitive"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   960
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.TextBox RText 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   480
      Width           =   3135
   End
   Begin VB.TextBox FText 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   3135
   End
   Begin etUCP.chameleonButton chameleonButton1 
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Replace"
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
      MICON           =   "ReplaceTXT.frx":030A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton chameleonButton2 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Cancel"
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
      MICON           =   "ReplaceTXT.frx":0326
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "String to Replace With:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "String to Search For:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "ReplaceTXT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReplaceActiveForm As frmDocument

Private Sub chameleonButton1_Click()
If CaseSens.Value = 1 Then
ReplaceActiveForm.ReplaceProgramText FText.Text, RText.Text, True
Else
ReplaceActiveForm.ReplaceProgramText FText.Text, RText.Text, False
End If

Me.Hide
End Sub

Private Sub chameleonButton2_Click()
Me.Hide
End Sub

Private Sub Form_Load()
    DrawXPCtl Me
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Sub ReplaceText(AForm As Form)
Me.Caption = "Replace: " & AForm.Caption
Set ReplaceActiveForm = AForm
Me.Show
End Sub
