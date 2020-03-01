VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form MSprogrammer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MicroScout Programmer - .nqc Creator   (C) Dean Camera, 2003"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7260
   FillColor       =   &H8000000F&
   ForeColor       =   &H80000013&
   Icon            =   "MSprogrammer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin etUCP.chameleonButton Command6 
      Height          =   320
      Left            =   5280
      TabIndex        =   31
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Translate"
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
      MICON           =   "MSprogrammer.frx":030A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Command4 
      Height          =   320
      Left            =   4800
      TabIndex        =   29
      Top             =   2880
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Open..."
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
      MICON           =   "MSprogrammer.frx":0326
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
      Left            =   2280
      TabIndex        =   28
      Top             =   3120
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Back To Editor"
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
      MICON           =   "MSprogrammer.frx":0342
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
      TabIndex        =   27
      Top             =   3120
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Copy To Clipboard"
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
      MICON           =   "MSprogrammer.frx":035E
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
      TabIndex        =   26
      Top             =   3120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Generate NQC Code"
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
      MICON           =   "MSprogrammer.frx":037A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   3960
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Autorun While Downloading"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   3050
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0E0FF&
      Height          =   2895
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   4455
   End
   Begin RichTextLib.RichTextBox text1 
      Height          =   2895
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5106
      _Version        =   393217
      BackColor       =   16761024
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"MSprogrammer.frx":0396
   End
   Begin VB.Frame Frame1 
      Height          =   375
      Left            =   -240
      TabIndex        =   24
      Top             =   3480
      Width           =   7575
      Begin VB.Label Label23 
         BackColor       =   &H00C0C0C0&
         Caption         =   " **** Automatic Translation when saving and loading ****     **** Editor is NOT case-sensitive ****"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   190
         Width           =   7335
      End
   End
   Begin etUCP.chameleonButton Command5 
      Height          =   315
      Left            =   5880
      TabIndex        =   30
      Top             =   2880
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Save..."
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
      MICON           =   "MSprogrammer.frx":049E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Caption         =   "<None>"
      Height          =   255
      Left            =   6120
      TabIndex        =   22
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Caption         =   "<None>"
      Height          =   255
      Left            =   6120
      TabIndex        =   21
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Caption         =   "<.5,1,2,5>"
      Height          =   255
      Left            =   6120
      TabIndex        =   20
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   "<None>"
      Height          =   255
      Left            =   6120
      TabIndex        =   19
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Caption         =   "<None>"
      Height          =   255
      Left            =   6120
      TabIndex        =   18
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "<None>"
      Height          =   255
      Left            =   6120
      TabIndex        =   17
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "<.5,1,2,5>"
      Height          =   255
      Left            =   6120
      TabIndex        =   16
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "<1,2,3,4,5>"
      Height          =   255
      Left            =   6120
      TabIndex        =   15
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "Parameters:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6120
      TabIndex        =   14
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "<A,B,C,1,2,3>"
      Height          =   255
      Left            =   6120
      TabIndex        =   13
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Put cursor over command for info."
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   12
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label10 
      Caption         =   "Wait For Light"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4800
      TabIndex        =   11
      ToolTipText     =   "Waits until light is shone onto the sensor."
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label9 
      Caption         =   "Seek Light"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      ToolTipText     =   "The MicroScout seeks light with its motor and light sensor."
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "Motor Back"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      ToolTipText     =   "Switches the motor on in the reverse direction for a set period of time, as specified in the parameter."
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "Motor Reverse"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      ToolTipText     =   "Turns the motor on in the reverse direction."
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "Motor Forward"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      ToolTipText     =   "Turns the motor on in the forward direction."
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Motor Off"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4800
      TabIndex        =   6
      ToolTipText     =   "Switches the motor off."
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Motor On"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      ToolTipText     =   "Switches the motor on for a set period of time, as specified in the parameter."
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Beep"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4800
      TabIndex        =   4
      ToolTipText     =   "Performs one of five different beeps. These are inbult into the MicroScout."
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Port"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      ToolTipText     =   "Selects output port. Light sensors can be used on sensor ports. This command MUST go first."
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Commands:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "MSprogrammer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AutoRunTemp, VLLHEADER, TextTMP, PORTCHAR As String
Dim fixed

Private Sub Command1_Click()
    On Error Resume Next
    PORTCHAR = "a"

    If Asc(Mid(text1.text, Len(text1.text) - 1, 1)) <> Asc(vbNewLine) Then
        text1.text = text1.text & vbNewLine
    End If

    a = 1
    For I = 1 To Len(text1.text)
        If Asc(Mid(text1.text, I, 1)) = Asc(vbNewLine) Then
            fixed = fixed & FixCommand(Mid(text1.text, a, I - a))
            a = I
        End If
    Next I

    If Check1.Value = 1 Then
        AutoRunTemp = vbNewLine & VLLHEADER & "(MSVLL_D_RUN);" & vbNewLine
    Else
        AutoRunTemp = ""
    End If

    VLLHEADER = "vll_" & PORTCHAR

    Text2.text = "#define VLLP" & UCase(PORTCHAR) & vbNewLine & "#include " & Chr(34) & "libvll.nqc" & Chr(34) & vbNewLine & vbNewLine & "task main()" & vbNewLine & "{" & vbNewLine & VLLHEADER & "(MSVLL_D_DEL_SCRIPT);" & vbNewLine & fixed & vbNewLine & AutoRunTemp & "}"
    Text2.Visible = True

    Command2.Visible = True
    Command3.Visible = True
End Sub

Function FixCommand(Command)
    FixCommand = ""
    VLLHEADER = "vll_" & PORTCHAR

    Command = UCase(Command)

    If InStr(1, Command, "PORT A") Then PORTCHAR = "a"
    If InStr(1, Command, "PORT B") Then PORTCHAR = "b"
    If InStr(1, Command, "PORT C") Then PORTCHAR = "c"
    If InStr(1, Command, "PORT 1") Then PORTCHAR = "1"
    If InStr(1, Command, "PORT 2") Then PORTCHAR = "2"
    If InStr(1, Command, "PORT 3") Then PORTCHAR = "3"


    If InStr(1, Command, "MOTOR FORWARD") Then FixCommand = vbNewLine & VLLHEADER & "(MSVLL_D_FWD);"
    If InStr(1, Command, "MOTOR REVERSE") Then FixCommand = vbNewLine & VLLHEADER & "(MSVLL_D_RWD);"
    If InStr(1, Command, "MOTOR ON 1") Then FixCommand = vbNewLine & VLLHEADER & "(MSVLL_S_FWD1);"
    If InStr(1, Command, "MOTOR ON .5") Then FixCommand = vbNewLine & VLLHEADER & "(MSVLL_S_FWD05);"
    If InStr(1, Command, "MOTOR ON 2") Then FixCommand = vbNewLine & VLLHEADER & "(MSVLL_S_FWD2);"
    If InStr(1, Command, "MOTOR ON 3") Then FixCommand = vbNewLine & VLLHEADER & "(MSVLL_S_FWD3);"
    If InStr(1, Command, "MOTOR ON 4") Then FixCommand = vbNewLine & VLLHEADER & "(MSVLL_S_FWD4);"
    If InStr(1, Command, "MOTOR ON 5") Then FixCommand = vbNewLine & VLLHEADER & "(MSVLL_S_FWD5);"
    If InStr(1, Command, "MOTOR OFF") Then FixCommand = vbNewLine & VLLHEADER & "(MSVLL_D_STOP);"
    If InStr(1, Command, "MOTOR BACK 1") Then FixCommand = vbNewLine & VLLHEADER & "(MSVLL_S_RWD1);"
    If InStr(1, Command, "MOTOR BACK .5") Then FixCommand = vbNewLine & VLLHEADER & "(MSVLL_S_RWD05);"
    If InStr(1, Command, "MOTOR BACK 2") Then FixCommand = vbNewLine & VLLHEADER & "(MSVLL_S_RWD2);"
    If InStr(1, Command, "MOTOR BACK 3") Then FixCommand = vbNewLine & VLLHEADER & "(MSVLL_S_RWD3);"
    If InStr(1, Command, "MOTOR BACK 4") Then FixCommand = vbNewLine & VLLHEADER & "(MSVLL_S_RWD4);"
    If InStr(1, Command, "MOTOR BACK 5") Then FixCommand = vbNewLine & VLLHEADER & "(MSVLL_S_RWD5);"
    If InStr(1, Command, "BEEP 1") Then FixCommand = vbNewLine & VLLHEADER & "(MSVLL_D_BEEP1);"
    If InStr(1, Command, "BEEP 2") Then FixCommand = vbNewLine & VLLHEADER & "(MSVLL_D_BEEP2);"
    If InStr(1, Command, "BEEP 3") Then FixCommand = vbNewLine & VLLHEADER & "(MSVLL_D_BEEP3);"
    If InStr(1, Command, "BEEP 4") Then FixCommand = vbNewLine & VLLHEADER & "(MSVLL_D_BEEP4);"
    If InStr(1, Command, "BEEP 5") Then FixCommand = vbNewLine & VLLHEADER & "(MSVLL_D_BEEP5);"
    If InStr(1, Command, "SEEK LIGHT") Then FixCommand = vbNewLine & VLLHEADER & "(MSVLL_S_SEEK_LIGHT);"
    If InStr(1, Command, "WAIT FOR LIGHT") Then FixCommand = vbNewLine & VLLHEADER & "(MSVLL_S_WAIT4_LIGHT);"
End Function

Private Sub Command2_Click()
    Clipboard.Clear
    Clipboard.SetText Text2.text
End Sub

Private Sub Command3_Click()
    Command2.Visible = False
    Command3.Visible = False
    Text2.Visible = False
End Sub

Private Sub Command4_Click()
    CD.DialogTitle = "Load NQC File"
    CD.Filter = "NQC Files (.nqc)|*.nqc"
    CD.ShowOpen
    If CD.FileName <> "" Then
        Text2.Visible = False
        Command2.Visible = False
        Command3.Visible = False
        text1.LoadFile CD.FileName
        Translate
    End If
End Sub

Private Sub Command5_Click()
    CD.DefaultExt = ".nqc"
    CD.DialogTitle = "Save NQC File"
    CD.Filter = "NQC Files (.nqc)|*.nqc"
    CD.ShowSave
    If CD.FileName <> "" Then
        TextTMP = text1.text
        Command1_Click
        Text2.Visible = False
        Command2.Visible = False
        Command3.Visible = False
        text1.text = Text2.text
        Open CD.FileName For Output As #1
        Print #1, text1.text
        Close #1
        text1.text = TextTMP
    End If
End Sub

Function Translate()
    On Error Resume Next
    PORTCHAR = "a"

    Text2.text = text1.text

    a = 1
    For I = 1 To Len(Text2.text)
        If Asc(Mid(Text2.text, I, 1)) = Asc(vbNewLine) Then
            fixed = fixed & BackFixCommand(Mid(Text2.text, a, I - a))
            a = I + 1
        End If
    Next I

    text1.text = fixed
    Text2.Visible = False

    If Asc(Mid(text1.text, Len(text1.text) - 1, 1)) <> Asc(vbNewLine) Then
        text1.text = text1.text & vbNewLine
    End If

    text1.text = Mid(text1.text, 3)
End Function

Function BackFixCommand(Command)
    BackFixCommand = ""

    If Len(Command) > 40 Then Command = Mid(Command, 1, 40)

    Command = UCase(Command)

    If InStr(1, Command, "#DEFINE VLLP1") Then BackFixCommand = vbNewLine & "Port 1"
    If InStr(1, Command, "#DEFINE VLLP2") Then BackFixCommand = vbNewLine & "Port 2"
    If InStr(1, Command, "#DEFINE VLLP3") Then BackFixCommand = vbNewLine & "Port 3"
    If InStr(1, Command, "#DEFINE VLLPA") Then BackFixCommand = vbNewLine & "Port A"
    If InStr(1, Command, "#DEFINE VLLPB") Then BackFixCommand = vbNewLine & "Port B"
    If InStr(1, Command, "#DEFINE VLLPC") Then BackFixCommand = vbNewLine & "Port C"

    If InStr(1, Command, "MSVLL_D_FWD") Then BackFixCommand = vbNewLine & "Motor Forward"
    If InStr(1, Command, "MSVLL_D_RWD") Then BackFixCommand = vbNewLine & "Motor Reverse"
    If InStr(1, Command, "MSVLL_S_FWD1") Then BackFixCommand = vbNewLine & "Motor On 1"
    If InStr(1, Command, "MSVLL_S_FWD05") Then BackFixCommand = vbNewLine & "Motor On .5"
    If InStr(1, Command, "MSVLL_S_FWD2") Then BackFixCommand = vbNewLine & "Motor On 2"
    If InStr(1, Command, "MSVLL_S_FWD5") Then BackFixCommand = vbNewLine & "Motor On 5"
    If InStr(1, Command, "MSVLL_D_STOP") Then BackFixCommand = vbNewLine & "Motor Off"
    If InStr(1, Command, "MSVLL_S_RWD1") Then BackFixCommand = vbNewLine & "Motor Back 1"
    If InStr(1, Command, "MSVLL_S_RWD05") Then BackFixCommand = vbNewLine & "Motor Back .5"
    If InStr(1, Command, "MSVLL_S_RWD2") Then BackFixCommand = vbNewLine & "Motor Back 2"
    If InStr(1, Command, "MSVLL_S_RWD5") Then BackFixCommand = vbNewLine & "Motor Back 5"
    If InStr(1, Command, "MSVLL_D_BEEP1") Then BackFixCommand = vbNewLine & "Beep 1"
    If InStr(1, Command, "MSVLL_D_BEEP2") Then BackFixCommand = vbNewLine & "Beep 2"
    If InStr(1, Command, "MSVLL_D_BEEP3") Then BackFixCommand = vbNewLine & "Beep 3"
    If InStr(1, Command, "MSVLL_D_BEEP4") Then BackFixCommand = vbNewLine & "Beep 4"
    If InStr(1, Command, "MSVLL_D_BEEP5") Then BackFixCommand = vbNewLine & "Beep 5"
    If InStr(1, Command, "MSVLL_S_SEEK_LIGHT") Then BackFixCommand = vbNewLine & "Seek Light"
    If InStr(1, Command, "MSVLL_S_WAIT4_LIGHT") Then BackFixCommand = vbNewLine & "Wait For Light"
End Function

Private Sub Command6_Click()
    Translate
End Sub

Private Sub Form_Load()
LogText "Load - MSPROGRAMMER"
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
    DrawXPCtl Me
End Sub

