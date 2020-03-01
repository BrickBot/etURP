VERSION 5.00
Begin VB.Form firmware 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Firmware Download"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4380
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer FirmwareTimer 
      Interval        =   1000
      Left            =   1800
      Top             =   2160
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DO NOT TURN OFF THE RCX UNTIL DOWNLOAD IS FINISHED"
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
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Width           =   4335
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   4320
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Left            =   3840
      Top             =   960
      Width           =   255
   End
   Begin VB.Image Downloading 
      Height          =   270
      Left            =   3900
      Picture         =   "firmwaredownload.frx":0000
      Top             =   1005
      Width           =   255
   End
   Begin VB.Label FirmwareSize 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Firmware Size"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Firmware Size (Bytes):"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label FirmwareName 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Firmware Filename"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Firmware Filename:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4320
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "FIRMWARE DOWNLOADING, PLEASE WAIT..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label SE 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "0 Seconds Elapsed"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   4095
   End
End
Attribute VB_Name = "firmware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SecsElapsed As Integer

Private Sub FirmwareTimer_Timer()
    SecsElapsed = SecsElapsed + 1

    SE.Caption = SecsElapsed & " Seconds Elapsed."

    Downloading.Visible = Not Downloading.Visible
End Sub

Private Sub Form_Load()
LogText "Load - FIRMWARE"
    Me.Visible = False
    SecsElapsed = 0
    Me.Icon = MSprogrammer.Icon
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
    Me.Visible = True
End Sub

Sub DownloadFirmware()
    On Error Resume Next
    SecsElapsed = 0
    DownloadingFirmware = True
    Me.Visible = True
    FDownloadTemp = About.Spirit.DownloadFirmware(FIRMLOCATION)
    If FDownloadTemp <> True Then
        XPLib.XPMsgBox "Firmware Download failed. Move the RCX closer and try again.", "Firmware Download", True, XP_OKOnly, msg_Critical
        Me.FirmwareTimer.Enabled = False
        Me.Visible = False
        Me.Hide
    End If
    Me.FirmwareTimer.Enabled = True
End Sub

Sub UnlockDownloadedFirmware()
    FDownloadTemp = About.Spirit.UnlockFirmware("Do you byte, when I knock?")
    If UCase(Mid(FDownloadTemp, 1, 10)) = "DOWNLOAD F" Then
        XPLib.XPMsgBox "Firmware Download failed. Move the RCX closer and try again.", "Firmware Download", True, XP_OKOnly, msg_Critical
    Else
        NOFIRM = False
    End If
    Me.FirmwareTimer.Enabled = False
    Me.Visible = False
    Me.Hide
End Sub

Sub WarnDownloadError()
    XPLib.XPMsgBox "Firmware Download failed. Move the RCX closer and try again.", "Firmware Download", True, XP_OKOnly, msg_Critical
    Me.FirmwareTimer.Enabled = False
    Me.Visible = False
    Me.Hide
End Sub

Sub UpdateFirmwareSpecs()
    FirmwareName = FIRMNAME
    FirmwareSize.Caption = FirmwareByteSize
End Sub

Private Sub Image1_Click()

End Sub

Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

