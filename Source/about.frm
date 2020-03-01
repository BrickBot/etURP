VERSION 5.00
Object = "{D6CD40C0-A522-11D0-9800-D3C9B35D2C47}#1.0#0"; "SPIRIT.OCX"
Begin VB.Form about 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About En-Tech URC"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin etUCP.chameleonButton Command3 
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "About OlsenXP"
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
      MICON           =   "about.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Command2 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "About Spirit"
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
      MICON           =   "about.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin SPIRITLib.Spirit Spirit 
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   873
         _StockProps     =   0
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"about.frx":0038
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1335
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "En-Tech is in NO WAY connected to the LEGO company. Use this program at your OWN RISK. This program is FREEWARE."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   0
         TabIndex        =   3
         Top             =   2400
         Width           =   3135
      End
      Begin VB.Label Version 
         Alignment       =   2  'Center
         Caption         =   "Version"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "                   En-Tech                             Ultimate Robot Programmer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
   End
   Begin etUCP.chameleonButton Command1 
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   3120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Ok"
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
      MICON           =   "about.frx":010B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XPManifestLib As New ClsXPManifest

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Hide
    Command1.Enabled = False
    Command1.Enabled = True
End Sub

Private Sub Command2_Click()
    Spirit.AboutBox
End Sub

Private Sub Command3_Click()
    XPLib.AboutXP
End Sub

Private Sub Form_Load()
LogText "Load - ABOUT"
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
    version.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    DrawXPCtl Me
End Sub

Private Sub Spirit_AsyncronBrickError(ByVal Number As Integer, Description As String)
    fMainForm.sbStatusBar.Panels(1).text = "SPIRIT error number " & Number & ". " & Description
End Sub

Private Sub Spirit_DownloadDone(ByVal ErrorCode As Integer, ByVal DownloadNo As Integer)
    If ErrorCode = 0 Then
        firmware.UnlockDownloadedFirmware
    Else
        firmware.WarnDownloadError
    End If
End Sub

Private Sub Spirit_DownloadSize(ByVal DownloadSizeInBytes As Long)
    FirmwareByteSize = DownloadSizeInBytes
End Sub

Private Sub Spirit_downloadStatus(ByVal timeInMS As Long, ByVal sizeInBytes As Long, ByVal taskNo As Integer)
    FirmwareByteSize = sizeInBytes
    FirmwareMaxSecs = (timeInMS / 60)
    If DownloadingFirmware = True Then
        firmware.UpdateFirmwareSpecs
        DownloadingFirmware = False
    End If
End Sub

Private Sub Spirit_DownloadTime(ByVal DownloadTimeInMS As Long)
    FirmwareMaxSecs = (DownloadTimeInMS / 60)
End Sub

Private Sub Spirit_DownloadTimeAndSize(ByVal timeInDeciSeconds As Long, ByVal sizeInBytes As Long)
    FirmwareMaxSecs = timeInDeciSeconds
    FirmwareByteSize = sizeInBytes
End Sub
