VERSION 5.00
Object = "{3AC5FA56-A608-11D5-AD7C-B629F13B4140}#1.0#0"; "USEFULL.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5580
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      Height          =   2070
      Left            =   45
      TabIndex        =   0
      Top             =   -15
      Width           =   5460
      Begin Usefull.USEFULLctrl uc 
         Left            =   4200
         Top             =   1800
         _ExtentX        =   476
         _ExtentY        =   794
      End
      Begin VB.Timer Timer1 
         Interval        =   1500
         Left            =   120
         Top             =   480
      End
      Begin VB.PictureBox picLogo 
         BorderStyle     =   0  'None
         Height          =   585
         Left            =   3240
         Picture         =   "frmSplash.frx":0000
         ScaleHeight     =   585
         ScaleWidth      =   615
         TabIndex        =   2
         Top             =   1335
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "By Dean Camera"
         BeginProperty Font 
            Name            =   "Bradley Hand ITC"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1665
         TabIndex        =   8
         Top             =   1665
         Width           =   2535
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "LicenseTo"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Tag             =   "LicenseTo"
         Top             =   0
         Visible         =   0   'False
         Width           =   6855
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         Caption         =   "PRODUCT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         TabIndex        =   7
         Tag             =   "Product"
         Top             =   960
         Width           =   5220
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "En-Tech"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1920
         TabIndex        =   6
         Tag             =   "CompanyProduct"
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6075
         TabIndex        =   5
         Tag             =   "Version"
         Top             =   2760
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lblCompany 
         Caption         =   "Company"
         Height          =   255
         Left            =   4710
         TabIndex        =   4
         Tag             =   "Company"
         Top             =   3330
         Width           =   2415
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright"
         Height          =   255
         Left            =   4710
         TabIndex        =   3
         Tag             =   "Copyright"
         Top             =   3120
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
uc.ACTN_PRG_StayOnTop True
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
End Sub

Private Sub Timer1_Timer()
    Unload frmSplash
    fMainForm.Show
End Sub
