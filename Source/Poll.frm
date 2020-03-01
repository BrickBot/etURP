VERSION 5.00
Begin VB.Form Poll 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RCX Variables"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   Icon            =   "Poll.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sensors"
      Height          =   1815
      Left            =   3360
      TabIndex        =   68
      Top             =   4200
      Width           =   4215
      Begin VB.CheckBox PollItem 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sensor 2"
         Height          =   255
         Index           =   33
         Left            =   120
         TabIndex        =   77
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox PollItem 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sensor 3"
         Height          =   255
         Index           =   32
         Left            =   120
         TabIndex        =   76
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox PollText 
         Height          =   285
         Index           =   33
         Left            =   3000
         TabIndex        =   75
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox PollText 
         Height          =   285
         Index           =   34
         Left            =   3000
         TabIndex        =   74
         Top             =   360
         Width           =   975
      End
      Begin VB.CheckBox PollItem 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sensor 1"
         Height          =   255
         Index           =   34
         Left            =   120
         TabIndex        =   73
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox PollText 
         Height          =   285
         Index           =   32
         Left            =   3000
         TabIndex        =   72
         Top             =   1320
         Width           =   975
      End
      Begin VB.ComboBox Sensor3 
         Height          =   315
         Left            =   1200
         TabIndex        =   71
         Text            =   "Sensor3"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox Sensor2 
         Height          =   315
         Left            =   1200
         TabIndex        =   70
         Text            =   "Sensor2"
         Top             =   840
         Width           =   1695
      End
      Begin VB.ComboBox Sensor1 
         Height          =   315
         Left            =   1200
         TabIndex        =   69
         Text            =   "Sensor1"
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Poll Selection"
      Height          =   975
      Left            =   120
      TabIndex        =   67
      Top             =   5040
      Width           =   2655
      Begin etUCP.chameleonButton Command1 
         Height          =   255
         Left            =   1440
         TabIndex        =   80
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Select None"
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
         MICON           =   "Poll.frx":030A
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
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Select All"
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
         MICON           =   "Poll.frx":0326
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin etUCP.chameleonButton PVariables 
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Poll Variables"
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
         MICON           =   "Poll.frx":0342
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Auto Poll"
      Height          =   735
      Left            =   120
      TabIndex        =   64
      Top             =   4200
      Width           =   2655
      Begin VB.ComboBox PollSecs 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         IntegralHeight  =   0   'False
         ItemData        =   "Poll.frx":035E
         Left            =   1150
         List            =   "Poll.frx":0360
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Auto Poll (                     Secs.)"
         Height          =   315
         Left            =   120
         TabIndex        =   66
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Timer AutoPoll 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8160
      Top             =   2640
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   31
      Left            =   6240
      TabIndex        =   63
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   31
      Left            =   5160
      TabIndex        =   62
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   30
      Left            =   5160
      TabIndex        =   61
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   30
      Left            =   6240
      TabIndex        =   60
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   29
      Left            =   5160
      TabIndex        =   59
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   29
      Left            =   6240
      TabIndex        =   58
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   28
      Left            =   5160
      TabIndex        =   57
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   28
      Left            =   6240
      TabIndex        =   56
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   27
      Left            =   5160
      TabIndex        =   55
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   27
      Left            =   6240
      TabIndex        =   54
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   26
      Left            =   5160
      TabIndex        =   53
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   26
      Left            =   6240
      TabIndex        =   52
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   25
      Left            =   5160
      TabIndex        =   51
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   25
      Left            =   6240
      TabIndex        =   50
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   24
      Left            =   5160
      TabIndex        =   49
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   24
      Left            =   6240
      TabIndex        =   48
      Top             =   840
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   23
      Left            =   5160
      TabIndex        =   47
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   23
      Left            =   6240
      TabIndex        =   46
      Top             =   480
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   22
      Left            =   5160
      TabIndex        =   45
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   22
      Left            =   6240
      TabIndex        =   44
      Top             =   120
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   21
      Left            =   2640
      TabIndex        =   43
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   21
      Left            =   3720
      TabIndex        =   42
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   20
      Left            =   2640
      TabIndex        =   41
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   20
      Left            =   3720
      TabIndex        =   40
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   19
      Left            =   2640
      TabIndex        =   39
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   19
      Left            =   3720
      TabIndex        =   38
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   18
      Left            =   2640
      TabIndex        =   37
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   18
      Left            =   3720
      TabIndex        =   36
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   17
      Left            =   2640
      TabIndex        =   35
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   17
      Left            =   3720
      TabIndex        =   34
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   16
      Left            =   2640
      TabIndex        =   33
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   16
      Left            =   3720
      TabIndex        =   32
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   15
      Left            =   2640
      TabIndex        =   31
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   15
      Left            =   3720
      TabIndex        =   30
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   14
      Left            =   2640
      TabIndex        =   29
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   14
      Left            =   3720
      TabIndex        =   28
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   13
      Left            =   2640
      TabIndex        =   27
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   13
      Left            =   3720
      TabIndex        =   26
      Top             =   840
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   12
      Left            =   2640
      TabIndex        =   25
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   12
      Left            =   3720
      TabIndex        =   24
      Top             =   480
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   11
      Left            =   2640
      TabIndex        =   23
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   11
      Left            =   3720
      TabIndex        =   22
      Top             =   120
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   21
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   10
      Left            =   1200
      TabIndex        =   20
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   19
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   9
      Left            =   1200
      TabIndex        =   18
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   8
      Left            =   1200
      TabIndex        =   16
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   1200
      TabIndex        =   14
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   1200
      TabIndex        =   12
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   1200
      TabIndex        =   10
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   1200
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   1200
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   1200
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox PollText 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CheckBox PollItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Poll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Busy As Boolean

Private Sub AutoPoll_Timer()
    PVariables_Click
End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then

        If PollSecs.text = "" Then
            Check1.Value = 0
            Check2.Value = 0
            Exit Sub
        End If

        If PollSecs.text = ".5" Then
            temp = 500
        Else
            temp = (1000 * Int(PollSecs.text))
        End If

        AutoPoll.Interval = temp

        AutoPoll.Enabled = True
    Else
        AutoPoll.Enabled = False
        AutoPoll.Interval = 1
    End If
End Sub

Private Sub Check2_Click()
    Check1.Value = Check2.Value
    Check1_Click
End Sub

Private Sub Command1_Click()
    For I = 0 To 31
        Me.PollItem.Item(I).Value = 1
    Next I
End Sub

Private Sub Command2_Click()
    For I = 0 To 31
        Me.PollItem.Item(I).Value = 0
    Next I
End Sub

Private Sub Form_Load()
LogText "Load - POLL"
Me.Visible = False
    DrawXPCtl Me
Me.Visible = True
End Sub

Private Sub PVariables_Click()
    If Busy = False Then
        Busy = True
        If COMOPEN = True Then
            If About.Spirit.TowerAndCableConnected And About.Spirit.PBAliveOrNot Then

                For I = 0 To 31
                    If Me.PollItem(I).Value = 1 Then Me.PollText(I) = About.Spirit.Poll(0, I)

                    DoEvents
                Next I

                If Me.PollItem(32).Value = 1 Then PollText(32).text = About.Spirit.Poll(9, 0)
                If Me.PollItem(33).Value = 1 Then PollText(33).text = About.Spirit.Poll(9, 1)
                If Me.PollItem(34).Value = 1 Then PollText(34).text = About.Spirit.Poll(9, 2)

            End If
        End If
        Busy = False
    End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Sensor1_Click()
    If Sensor1.text = "None (Raw)" Then
        About.Spirit.SetSensorType 0, 0
    End If

    If Sensor1.text = "Touch" Then
        About.Spirit.SetSensorType 0, 1
    End If

    If Sensor1.text = "Temperature" Then
        About.Spirit.SetSensorType 0, 2
    End If

    If Sensor1.text = "Light" Then
        About.Spirit.SetSensorType 0, 3
    End If

    If Sensor1.text = "Rotation" Then
        About.Spirit.SetSensorType 0, 4
    End If
End Sub

Private Sub Sensor2_Click()
    If Sensor2.text = "None (Raw)" Then
        About.Spirit.SetSensorType 1, 0
    End If

    If Sensor2.text = "Touch" Then
        About.Spirit.SetSensorType 1, 1
    End If

    If Sensor2.text = "Temperature" Then
        About.Spirit.SetSensorType 1, 2
    End If

    If Sensor2.text = "Light" Then
        About.Spirit.SetSensorType 1, 3
    End If

    If Sensor2.text = "Rotation" Then
        About.Spirit.SetSensorType 1, 4
    End If
End Sub

Private Sub Sensor3_Click()
    If Sensor3.text = "None (Raw)" Then
        About.Spirit.SetSensorType 2, 0
    End If

    If Sensor3.text = "Touch" Then
        About.Spirit.SetSensorType 2, 1
    End If

    If Sensor3.text = "Temperature" Then
        About.Spirit.SetSensorType 2, 2
    End If

    If Sensor3.text = "Light" Then
        About.Spirit.SetSensorType 2, 3
    End If

    If Sensor3.text = "Rotation" Then
        About.Spirit.SetSensorType 2, 4
    End If
End Sub


Sub ActivatePoll()

    PollSecs.Clear
    Sensor1.Clear
    Sensor2.Clear
    Sensor3.Clear

    PollSecs.AddItem ".5"
    PollSecs.AddItem "1"
    PollSecs.AddItem "2"
    PollSecs.AddItem "3"
    PollSecs.AddItem "4"
    PollSecs.AddItem "5"

    Sensor1.AddItem "None (Raw)"
    Sensor1.AddItem "Touch"
    Sensor1.AddItem "Temperature"
    Sensor1.AddItem "Light"
    Sensor1.AddItem "Rotation"

    Sensor2.AddItem "None (Raw)"
    Sensor2.AddItem "Touch"
    Sensor2.AddItem "Temperature"
    Sensor2.AddItem "Light"
    Sensor2.AddItem "Rotation"

    Sensor3.AddItem "None (Raw)"
    Sensor3.AddItem "Touch"
    Sensor3.AddItem "Temperature"
    Sensor3.AddItem "Light"
    Sensor3.AddItem "Rotation"

    For I = 0 To 31
        Me.PollItem.Item(I).Caption = "Var " & I
        Me.PollText.Item(I).text = "?"
        Me.PollText.Item(I).Locked = True
    Next I

    Me.PollText.Item(32).text = "?"
    Me.PollText.Item(33).text = "?"
    Me.PollText.Item(34).text = "?"

    Static Ione, Itwo, Ithree As Integer

    Ione = About.Spirit.Poll(10, 0)
    Itwo = About.Spirit.Poll(10, 1)
    Ithree = About.Spirit.Poll(10, 2)

    Static SOneTemp, STwoTemp, SThreeTemp As String
    SOneTemp = "None (Raw)"
    STwoTemp = "None (Raw)"
    SThreeTemp = "None (Raw)"

    Select Case Ione
        Case 1
            SOneTemp = "None (Raw)"
        Case 2
            SOneTemp = "Touch"
        Case 3
            SOneTemp = "Temperature"
        Case 4
            SOneTemp = "Light"
        Case 5
            SOneTemp = "Rotation"
    End Select

    Select Case Itwo
        Case 1
            STwoTemp = "None (Raw)"
        Case 2
            STwoTemp = "Touch"
        Case 3
            STwoTemp = "Temperature"
        Case 4
            STwoTemp = "Light"
        Case 5
            STwoTemp = "Rotation"
    End Select

    Select Case Ithree
        Case 1
            SThreeTemp = "None (Raw)"
        Case 2
            SThreeTemp = "Touch"
        Case 3
            SThreeTemp = "Temperature"
        Case 4
            SThreeTemp = "Light"
        Case 5
            SThreeTemp = "Rotation"
    End Select

    Sensor1.text = SOneTemp
    Sensor2.text = STwoTemp
    Sensor3.text = SThreeTemp
End Sub
