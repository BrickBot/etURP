VERSION 5.00
Begin VB.Form SaveFile 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "SaveFile"
   ClientHeight    =   1815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin etUCP.chameleonButton chameleonButton1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Save"
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
      MICON           =   "SaveFile.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.Titlebar Titlebar1 
      Height          =   1815
      Left            =   0
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3201
   End
   Begin etUCP.chameleonButton chameleonButton2 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Discard"
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
      MICON           =   "SaveFile.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton chameleonButton3 
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "SaveFile.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Save Changes to """" ?"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   3495
   End
End
Attribute VB_Name = "SaveFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chameleonButton1_Click()
Me.Hide
XPSaveAnswer = "Save"
Unload Me
End Sub

Private Sub chameleonButton2_Click()
Me.Hide
XPSaveAnswer = "Discard"
Unload Me
End Sub

Private Sub chameleonButton3_Click()
Me.Hide
XPSaveAnswer = "Cancel"
Unload Me
End Sub

Private Sub Form_Load()
Me.Titlebar1.Caption = "Save Changes - " & SAVEFileName
Me.Label1.Caption = "Save changes to " & Chr(34) & SAVEFileName & Chr(34) & " ?"
Me.Titlebar1.ShowMaximize = False
Me.Titlebar1.ShowMinimize = False
SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
XPSaveAnswer = "-"
End Sub

Private Sub Titlebar1_Closed()
Me.Hide
XPSaveAnswer = "Cancel"
Unload Me
End Sub
