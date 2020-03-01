VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SearchRCX 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search for Brick"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   Icon            =   "searchRCX.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin etUCP.chameleonButton Command2 
      Height          =   375
      Left            =   1200
      TabIndex        =   14
      Top             =   2400
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
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "searchRCX.frx":030A
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
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
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
      MICON           =   "searchRCX.frx":0326
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Brick Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   2400
      TabIndex        =   7
      Top             =   120
      Width           =   1815
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Other"
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   1575
         Begin VB.OptionButton CMaster 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cybermaster"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "RCX"
         Height          =   975
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1575
         Begin VB.OptionButton Option6 
            BackColor       =   &H00C0C0C0&
            Caption         =   "RCX"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            ToolTipText     =   "Use RCX with 1.0 or 1.5 Firmware."
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00C0C0C0&
            Caption         =   "RCX2"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            ToolTipText     =   "Use RCX with 2.0 Firmware."
            Top             =   580
            Width           =   735
         End
      End
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2865
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Port Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Auto"
         Height          =   375
         Left            =   480
         TabIndex        =   5
         ToolTipText     =   "Automatically select COMPORT. This can be slower than selecting the COMPORT from the list."
         Top             =   1680
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Comport 4"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Comport 3"
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Comport 2"
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Comport 1"
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
End
Attribute VB_Name = "SearchRCX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Port As String

Private Sub AutoCheck_Click()
    CMaster.Value = False
End Sub

Private Sub CMaster_Click()
    If CMaster.Value = True Then
        BrickType = CYBERMASTER
        Option6.Value = False
        Option7.Value = False
        AutoCheck.Value = False
    End If
End Sub

Private Sub Command1_Click()
    If Option6.Value = True Then BrickType = RCX
    If Option7.Value = True Then BrickType = RCX2
    If CMaster.Value = True Then BrickType = CYBERMASTER

    Select Case BrickType
        Case CYBERMASTER
            About.Spirit.LinkType = Cable
            About.Spirit.PBrick = Spirit
        Case RCX Or RCX2
            About.Spirit.LinkType = InfraRed
            About.Spirit.PBrick = RCX
    End Select

    COMOPENCLOSED = 1
    Command1.Enabled = False
    Command2.Enabled = False
    bestuse = 0
    found = 0

    If Option1.Value = True Then
        SB.SimpleText = "Searching for Brick/Tower"
        About.Spirit.ComPortNo = 1
        If About.Spirit.InitComm Then
            If About.Spirit.TowerAndCableConnected = True And About.Spirit.PBAliveOrNot = False Then
                SB.SimpleText = "Tower Found."
                bestuse = 1
            End If
            If About.Spirit.TowerAndCableConnected And About.Spirit.PBAliveOrNot Then
                found = 1
                SB.SimpleText = "Brick Found."
            End If
        End If
        If bestuse = 0 And found = 0 Then
            XPLib.XPMsgBox "Brick/Tower Not Found on Selected Port.", "Search for Brick", False, XP_OKOnly, msg_Critical
            Command1.Enabled = True
            Command2.Enabled = True
            SB.SimpleText = "Port open failed."
            Exit Sub
        End If
    End If

    If Option2.Value = True Then
        SB.SimpleText = "Searching for Brick/Tower"
        About.Spirit.ComPortNo = 2
        If About.Spirit.InitComm Then
            If About.Spirit.TowerAndCableConnected = True And About.Spirit.PBAliveOrNot = False Then
                SB.SimpleText = "Tower Found."
                bestuse = 1
            End If
            If About.Spirit.TowerAndCableConnected And About.Spirit.PBAliveOrNot Then
                found = 1
                SB.SimpleText = "Brick Found."
            End If
        End If
        If bestuse = 0 And found = 0 Then
            XPLib.XPMsgBox "Brick/Tower Not Found on Selected Port.", "Search for Brick", False, XP_OKOnly, msg_Critical
            Command1.Enabled = True
            Command2.Enabled = True
            Exit Sub
        End If
    End If

    If Option3.Value = True Then
        SB.SimpleText = "Searching for Brick/Tower"
        About.Spirit.ComPortNo = 3
        If About.Spirit.InitComm Then
            If About.Spirit.TowerAndCableConnected = True And About.Spirit.PBAliveOrNot = False Then
                SB.SimpleText = "Tower Found."
                bestuse = 1
            End If
            If About.Spirit.TowerAndCableConnected And About.Spirit.PBAliveOrNot Then
                found = 1
                SB.SimpleText = "Brick Found."
            End If
        End If
        If bestuse = 0 And found = 0 Then
            XPLib.XPMsgBox "Brick/Tower Not Found on Selected Port.", "Search for Brick", False, XP_OKOnly, msg_Critical
            Command1.Enabled = True
            Command2.Enabled = True
            Exit Sub
        End If
    End If

    If Option4.Value = True Then
        SB.SimpleText = "Searching for Brick/Tower"
        About.Spirit.ComPortNo = 4
        If About.Spirit.InitComm Then
            If About.Spirit.TowerAndCableConnected = True And About.Spirit.PBAliveOrNot = False Then
                SB.SimpleText = "Tower Found."
                bestuse = 1
            End If
            If About.Spirit.TowerAndCableConnected And About.Spirit.PBAliveOrNot Then
                found = 1
                SB.SimpleText = "Brick Found."
            End If
        End If
        If bestuse = 0 And found = 0 Then
            XPLib.XPMsgBox "Brick/Tower Not Found on Selected Port.", "Search for Brick", False, XP_OKOnly, msg_Critical
            Command1.Enabled = True
            Command2.Enabled = True
            Exit Sub
        End If
    End If

    ' Auto Select
    If Option5.Value = True Then
        For I = 1 To 4
            SB.SimpleText = "Searching COM" & I
            About.Spirit.ComPortNo = I
            If About.Spirit.InitComm Then
                If About.Spirit.PBAliveOrNot And About.Spirit.TowerAndCableConnected Then
                    GoTo Continue
                    Exit For
                Else
                    About.Spirit.CloseComm
                End If
            End If
        Next I
        
        Command1.Enabled = True
        Command2.Enabled = True
        XPLib.XPMsgBox "Brick Not Found!", "Search for Brick", False, XP_OKOnly, msg_Critical
        SB.SimpleText = "Search failed."
        Exit Sub
    End If
    
Continue:

    COMOPEN = True
    FixTB = True
    
    MyCom = "COM" & Str(About.Spirit.ComPortNo)

    If found = 0 And bestuse = 1 Then
        XPLib.XPMsgBox "Tower, but not brick found. If brick is on, move it closer to the Tower and try again.", "Search For Brick", False, XP_OKOnly, msg_Critical
        Command1.Enabled = True
        Command2.Enabled = True
        Exit Sub
    End If

    FirmTemp = About.Spirit.UnlockPBrick
    If Right(FirmTemp, 5) = "00.00" Then NOFIRM = True

    If NOFIRM = False Then
        w = About.Spirit.Poll(14, 0)
        SetWatch.Label1.Caption = Int(w / 60)
        SetWatch.Label3.Caption = (Int(w / 60) - w / 60)
    Else
        FixTB = False
        XPLib.XPMsgBox "The RCX has no firmware. You will need to download this from the Tools menu.", "No Firmware Warning", False, XP_OKOnly, msg_Information
    End If

    If fMainForm Is Nothing Then
        Set fMainForm = New frmMain
        Load fMainForm
        fMainForm.tbToolBar.Enabled = True
    Else
        fMainForm.FlickMode
        fMainForm.Show
        fMainForm.tbToolBar.Enabled = True
    End If

    If NOFIRM = False Then
        fMainForm.tbToolBar.Visible = False
    Else
        fMainForm.tbToolBar.Visible = True
    End If

    Command1.Enabled = True
    Command2.Enabled = True

    Me.Hide
    Unload Me
End Sub

Private Sub Command2_Click()
    BrickType = NONE
    COMOPENCLOSED = 0
    COMOPEN = False
    DisableTB = True

    XPLib.XPMsgBox "No Brick used, some features will be disabled.", "No Brick Used", False, XP_OKOnly, msg_Exclamation

    Me.Hide
    Unload Me
    On Error Resume Next
    If fMainForm Is Nothing Then
        Set fMainForm = New frmMain
        Load fMainForm
        fMainForm.Show
        fMainForm.tbToolBar.Enabled = False
        frmMain.Show
    Else
        fMainForm.tbToolBar.Enabled = False
        fMainForm.FlickMode
        fMainForm.Show
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next

If Right(App.Path, 1) = "\" Then BetweenChar = "" Else BetweenChar = "\"
Kill App.Path & BetweenChar & "etURP.log"

LogText "*****START - etURP Startup Log (" & Now & ")*****"
        
    SearchRCX.Visible = False

    Set XPLib = New XPInterface
    MyCommand = Command

    DrawXPCtl Me

    Static ProgramOpened As Boolean

    If ProgramOpened = False Then
        Static ShowSearch As Integer
        ShowSearch = GetSetting("En-Tech URP", "Options", "StartupBrickSearch", 1)
        If ShowSearch = 0 Then
            frmMain.Show
            SearchRCX.Hide
            SearchRCX.Visible = True
            Exit Sub
        End If
    End If

    If MyCommand <> "" Then
        NoAllowSRCX = True
        BrickType = NONE
        COMOPENCLOSED = 0
        COMOPEN = False
        DisableTB = True

        Me.Hide
        Unload Me
        On Error Resume Next
        If fMainForm Is Nothing Then
            Set fMainForm = New frmMain
            Load fMainForm
            fMainForm.Show
            fMainForm.tbToolBar.Enabled = False
            frmMain.Show
        Else
            fMainForm.tbToolBar.Enabled = False
            fMainForm.FlickMode
            fMainForm.Show
        End If
        Exit Sub
    End If

    ProgramOpened = True

    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
    COMOPEN = False
    BrickType = NONE

    SetWatch.Hide
    For I = 0 To Forms.Count - 1
        If Forms(I).Caption <> Me.Caption Then Forms(I).Hide
    Next

    If App.PrevInstance = True Then
        XPLib.XPMsgBox "Unable to start etURP, etURP is already running.", "etURP Start Faliure", True, XP_OKOnly, msg_Critical
        LogText "FATAL STARTUP ERROR - Unable to start etURP, etURP is already running."
        LogText "*****END - etURP Startup Log (" & Now & ")*****"
        End
    End If

    SearchRCX.Visible = True
    SearchRCX.Show
End Sub

Private Sub Option6_Click()
    BrickType = RCX
    CMaster.Value = False
End Sub

Private Sub Option7_Click()
    BrickType = RCX2
    CMaster.Value = False
End Sub

