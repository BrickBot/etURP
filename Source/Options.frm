VERSION 5.00
Object = "{C8A61D56-D8DC-11D2-8064-9D6F06504DA8}#1.1#0"; "AXCOLCTL.OCX"
Begin VB.Form Options 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "En-Tech URP Options"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin etUCP.xpWellsTab OpTab 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   3836
      Alignment       =   0
      TabHeight       =   25
      BackColor       =   14741744
      ForeColor       =   -2147483630
      ForeColorActive =   9982008
      ForeColorHot    =   16711680
      FrameColor      =   8421504
      MaskColor       =   16711935
      SelectedTab     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumberOfTabs    =   3
      TabWidth1       =   60
      TabText1        =   "Editor"
      TabPicture1     =   "Options.frx":0000
      TabWidth2       =   60
      TabText2        =   "Startup"
      TabPicture2     =   "Options.frx":02B6
      TabWidth3       =   60
      TabText3        =   "Syntax"
      TabPicture3     =   "Options.frx":0590
      Begin VB.PictureBox OptionsTab 
         BackColor       =   &H00E0F0F0&
         BorderStyle     =   0  'None
         Height          =   1575
         Index           =   1
         Left            =   120
         ScaleHeight     =   1575
         ScaleWidth      =   3615
         TabIndex        =   4
         Top             =   480
         Width           =   3615
         Begin etUCP.chameleonButton chameleonButton1 
            Height          =   375
            Left            =   600
            TabIndex        =   23
            Top             =   1200
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            BTYPE           =   14
            TX              =   "Associate Files to etURP"
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
            MICON           =   "Options.frx":0846
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.CheckBox PBFT 
            BackColor       =   &H00E0F0F0&
            Caption         =   "Play Tune When Brick Found"
            Height          =   255
            Left            =   600
            TabIndex        =   22
            Top             =   120
            Width           =   2415
         End
         Begin VB.CheckBox STON 
            BackColor       =   &H00E0F0F0&
            Caption         =   "Show Templates on Startup"
            Height          =   255
            Left            =   600
            TabIndex        =   21
            Top             =   840
            Width           =   2415
         End
         Begin etUCP.chameleonButton egbutton 
            Height          =   255
            Left            =   3240
            TabIndex        =   20
            Top             =   1320
            Visible         =   0   'False
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BTYPE           =   14
            TX              =   "!"
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
            MICON           =   "Options.frx":0862
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.CheckBox SFBOS 
            BackColor       =   &H00E0F0F0&
            Caption         =   "Search For Brick On Startup"
            Height          =   255
            Left            =   600
            TabIndex        =   5
            Top             =   480
            Width           =   2415
         End
      End
      Begin VB.PictureBox OptionsTab 
         BackColor       =   &H00E0F0F0&
         BorderStyle     =   0  'None
         Height          =   1575
         Index           =   0
         Left            =   120
         ScaleHeight     =   1575
         ScaleWidth      =   3615
         TabIndex        =   1
         Top             =   480
         Width           =   3615
         Begin VB.OptionButton CASOLD 
            BackColor       =   &H00E0F0F0&
            Caption         =   "Cursor at End of Loaded Documents"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   1080
            Width           =   3615
         End
         Begin VB.OptionButton CAEOLD 
            BackColor       =   &H00E0F0F0&
            Caption         =   "Cursor at Start of Loaded Documents"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   840
            Width           =   3615
         End
         Begin VB.CheckBox USCC 
            BackColor       =   &H00E0F0F0&
            Caption         =   "Use Syntax Colour Coding"
            Height          =   255
            Left            =   600
            TabIndex        =   3
            Top             =   0
            Width           =   2655
         End
         Begin VB.CheckBox SGTLE 
            BackColor       =   &H00E0F0F0&
            Caption         =   "Show ""GoTo Line..."" Errors"
            Height          =   315
            Left            =   600
            TabIndex        =   2
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.PictureBox OptionsTab 
         BackColor       =   &H00E0F0F0&
         BorderStyle     =   0  'None
         Height          =   1575
         Index           =   2
         Left            =   120
         ScaleHeight     =   1575
         ScaleWidth      =   3615
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   3615
         Begin ImgColorPicker.ColorPalette CRes 
            Height          =   330
            Left            =   1080
            TabIndex        =   10
            Top             =   120
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   582
            BoxSize         =   5
         End
         Begin ImgColorPicker.ColorPalette CFunc 
            Height          =   330
            Left            =   1080
            TabIndex        =   11
            Top             =   480
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   582
            BoxSize         =   5
         End
         Begin ImgColorPicker.ColorPalette CComm 
            Height          =   330
            Left            =   1080
            TabIndex        =   12
            Top             =   840
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   582
            BoxSize         =   5
         End
         Begin ImgColorPicker.ColorPalette CKWrd 
            Height          =   330
            Left            =   1080
            TabIndex        =   13
            Top             =   1200
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   582
            BoxSize         =   5
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0F0F0&
            Caption         =   "Keyword :"
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label3 
            BackColor       =   &H00E0F0F0&
            Caption         =   "Comment :"
            Height          =   375
            Left            =   240
            TabIndex        =   9
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E0F0F0&
            Caption         =   "Function :"
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0F0F0&
            Caption         =   "Reserved :"
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Top             =   120
            Width           =   1095
         End
      End
   End
   Begin etUCP.chameleonButton Command2 
      Height          =   375
      Left            =   1560
      TabIndex        =   19
      Top             =   2400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Defaults"
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
      MICON           =   "Options.frx":087E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Command3 
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Top             =   2400
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "Options.frx":089A
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
      Height          =   375
      Left            =   120
      TabIndex        =   17
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
      MICON           =   "Options.frx":08B6
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
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()

End Sub

Private Sub chameleonButton1_Click()
Files.AssociateFile "etURP", "En-Tech Ultimate Robot Programmer for LEGO Mindstorms and Cybermaster", ".nqc"
Files.AssociateFile "etURP", "En-Tech Ultimate Robot Programmer for LEGO Mindstorms and Cybermaster", ".nqh"
End Sub

Private Sub Command1_Click()
    Me.Enabled = False
    SaveValues

    'Extra Code, needed to apply options
    GlobalRefresh
    '-----------------------------------
    Me.Enabled = True
    Me.Hide
End Sub

Private Sub Command2_Click()
    GetDefaults
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift = 1 And Button > 1 Then
        egbutton.Visible = True
    End If
End Sub

Private Sub Command3_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub egbutton_Click()
    EasterEgg
End Sub

Private Sub Form_Load()
LogText "Load - OPTIONS"
    Me.Visible = False
    DrawXPCtl Me
    Me.Icon = MSprogrammer.Icon
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3

    OpTab.DrawTab

    GetValues

    OpTab.SelectedTab = 1

    Me.Visible = True
End Sub

Sub FixOptions()
    Me.Hide
    GetValues
    Command1_Click
    Unload Me
End Sub

Sub GetValues()
    ' GET VALUES
    USCC.Value = GetSetting("En-Tech URP", "Editor", "SyntaxColour", 1)
    SGTLE.Value = GetSetting("En-Tech URP", "Editor", "GoToLineErrors", 1)
    PBFT.Value = GetSetting("En-Tech URP", "Startup", "BrickFoundTune", 1)
    SFBOS.Value = GetSetting("En-Tech URP", "Startup", "StartupBrickSearch", 1)
    STON.Value = GetSetting("En-Tech URP", "Startup", "AutoShowTemplates", 1)

    CRes.SelectedColor = GetSetting("En-Tech URP", "Syntax", "Reserved", 32)
    CFunc.SelectedColor = GetSetting("En-Tech URP", "Syntax", "Function", 22)
    CComm.SelectedColor = GetSetting("En-Tech URP", "Syntax", "Comment", 28)
    CKWrd.SelectedColor = GetSetting("En-Tech URP", "Syntax", "KeyWord", 25)
        
    CursorTemp = GetSetting("En-Tech URP", "Editor", "Cursor", 1)
    If CursorTemp = 1 Then
        CAEOLD.Value = True
    Else
        CASOLD.Value = True
    End If
End Sub

Sub GetDefaults()
    ' DEFAULTS
    USCC.Value = 1
    SGTLE.Value = 1
    PBFT.Value = 1
    SFBOS.Value = 1
    STON.Value = 1

    CAEOLD.Value = True

    CRes.SelectedColor = 32
    CFunc.SelectedColor = 22
    CComm.SelectedColor = 28
    CKWrd.SelectedColor = 25
End Sub

Sub SaveValues()
    ' SAVE VALUES
    SaveSetting "En-Tech URP", "Editor", "SyntaxColour", USCC.Value
    SaveSetting "En-Tech URP", "Editor", "GoToLineErrors", SGTLE.Value
    SaveSetting "En-Tech URP", "Startup", "BrickFoundTune", PBFT.Value
    SaveSetting "En-Tech URP", "Startup", "StartupBrickSearch", SFBOS.Value
    SaveSetting "En-Tech URP", "Startup", "AutoShowTemplates", STON.Value

    SaveSetting "En-Tech URP", "Syntax", "Reserved", Me.CRes.SelectedColor
    SaveSetting "En-Tech URP", "Syntax", "Function", Me.CFunc.SelectedColor
    SaveSetting "En-Tech URP", "Syntax", "Comment", Me.CComm.SelectedColor
    SaveSetting "En-Tech URP", "Syntax", "KeyWord", Me.CKWrd.SelectedColor
    SaveSetting "En-Tech URP", "Syntax", "ReservedC", CComm.Colors(Me.CRes.SelectedColor)
    SaveSetting "En-Tech URP", "Syntax", "FunctionC", CComm.Colors(Me.CFunc.SelectedColor)
    SaveSetting "En-Tech URP", "Syntax", "CommentC", CComm.Colors(Me.CComm.SelectedColor)
    SaveSetting "En-Tech URP", "Syntax", "KeyWordC", CComm.Colors(Me.CKWrd.SelectedColor)

    If CAEOLD.Value = True Then
        SaveSetting "En-Tech URP", "Options", "Cursor", 1
    Else
        SaveSetting "En-Tech URP", "Options", "Cursor", 0
    End If
End Sub

Private Sub OpTab_TabPressed(PreviousTab As Integer)
    For I = 0 To OptionsTab.Count - 1
        OptionsTab(I).Visible = False
    Next I

    OptionsTab(OpTab.SelectedTab - 1).Visible = True
End Sub

Sub GlobalRefresh()
    For I = 1 To Forms.Count - 1
        If Forms(I).Tag = "PROGRAM" Then
            Forms(I).rtftext.NewColorValue ColorComment, CComm.Colors(CComm.SelectedColor)
            Forms(I).rtftext.NewColorValue ColorFuncObj, CFunc.Colors(CFunc.SelectedColor)
            Forms(I).rtftext.NewColorValue ColorReserved, CRes.Colors(CRes.SelectedColor)
            Forms(I).rtftext.NewColorValue ColorKeyword, CKWrd.Colors(CKWrd.SelectedColor)

            Forms(I).rtftext.HighlightRefresh
        End If
    Next I
End Sub

