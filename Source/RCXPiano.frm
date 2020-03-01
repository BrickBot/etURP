VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form RCXPiano 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Brick Piano"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "RCXPiano.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5595
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox Notes 
      Height          =   1575
      Left            =   120
      TabIndex        =   47
      Top             =   3960
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2778
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"RCXPiano.frx":030A
   End
   Begin etUCP.chameleonButton chameleonButton1 
      Height          =   255
      Left            =   120
      TabIndex        =   46
      Top             =   3480
      Width           =   375
      _ExtentX        =   661
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
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   255
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   255
      MPTR            =   1
      MICON           =   "RCXPiano.frx":0412
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
      Left            =   2760
      TabIndex        =   44
      Top             =   2520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Clear"
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
      MICON           =   "RCXPiano.frx":042E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Command5 
      Height          =   735
      Left            =   4560
      TabIndex        =   43
      Top             =   2160
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Play"
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
      MICON           =   "RCXPiano.frx":044A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   285
      Left            =   2160
      TabIndex        =   41
      Top             =   3480
      Width           =   165
      _ExtentX        =   291
      _ExtentY        =   503
      _Version        =   393216
      Value           =   12
      BuddyControl    =   "VWaitTime"
      BuddyDispid     =   196610
      OrigLeft        =   1200
      OrigTop         =   3360
      OrigRight       =   1440
      OrigBottom      =   3855
      Max             =   20
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox VNoteTime 
      Height          =   285
      Left            =   3870
      Locked          =   -1  'True
      TabIndex        =   39
      Text            =   "10"
      Top             =   3480
      Width           =   390
   End
   Begin VB.TextBox VWaitTime 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "12"
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox BlackPianoKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   9
      Left            =   4800
      ScaleHeight     =   825
      ScaleWidth      =   225
      TabIndex        =   36
      Tag             =   "0155|0311|0622|1244|2489|4978|9956"
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox PianoKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   13
      Left            =   4920
      ScaleHeight     =   1425
      ScaleWidth      =   345
      TabIndex        =   35
      Tag             =   "0165|0330|0659|1318|2637|5274|10548"
      Top             =   240
      Width           =   375
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Select Piano Language"
      Height          =   975
      Left            =   1080
      TabIndex        =   30
      Top             =   1440
      Width           =   3495
      Begin etUCP.chameleonButton Command4 
         Height          =   255
         Left            =   960
         TabIndex        =   48
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Use"
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
         MICON           =   "RCXPiano.frx":0466
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.OptionButton UseNQC 
         Caption         =   "NQC"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton UseLASM 
         Caption         =   "LASM"
         Height          =   255
         Left            =   1200
         TabIndex        =   32
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton UseMScript 
         Caption         =   "Mindscript"
         Height          =   255
         Left            =   2160
         TabIndex        =   31
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Include ""Task Main"" Header"
      Height          =   255
      Left            =   2640
      TabIndex        =   29
      Top             =   3000
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Rest"
      Height          =   615
      Left            =   2040
      MaskColor       =   &H00000000&
      Picture         =   "RCXPiano.frx":0482
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2400
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Note Length"
      Height          =   1095
      Left            =   240
      TabIndex        =   22
      Top             =   2160
      Width           =   1695
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1/16"
         Height          =   255
         Left            =   840
         TabIndex        =   27
         Tag             =   "1"
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1/8"
         Height          =   255
         Left            =   840
         TabIndex        =   26
         Tag             =   "2"
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1/4"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Tag             =   "4"
         Top             =   720
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1/2"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Tag             =   "8"
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1/1"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Tag             =   "16"
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.PictureBox BlackPianoKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   8
      Left            =   4440
      ScaleHeight     =   825
      ScaleWidth      =   225
      TabIndex        =   21
      Tag             =   "0138|0277|0554|1109|2217|4435|8870"
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox BlackPianoKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   7
      Left            =   3720
      ScaleHeight     =   825
      ScaleWidth      =   225
      TabIndex        =   20
      Tag             =   "0116|0233|0466|0932|1865|3729|7459"
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox BlackPianoKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   6
      Left            =   3360
      ScaleHeight     =   825
      ScaleWidth      =   225
      TabIndex        =   19
      Tag             =   "0104|0208|0415|0831|1661|3322|6645"
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox BlackPianoKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   5
      Left            =   3000
      ScaleHeight     =   825
      ScaleWidth      =   225
      TabIndex        =   18
      Tag             =   "0092|0185|0370|0740|1480|2960|5920"
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox BlackPianoKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   4
      Left            =   2280
      ScaleHeight     =   825
      ScaleWidth      =   225
      TabIndex        =   17
      Tag             =   "0078|0155|0311|0622|1244|2489|4978"
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox BlackPianoKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   3
      Left            =   1920
      ScaleHeight     =   825
      ScaleWidth      =   225
      TabIndex        =   16
      Tag             =   "0069|0138|0277|0554|1109|2217|4435"
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox BlackPianoKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   2
      Left            =   1200
      ScaleHeight     =   825
      ScaleWidth      =   225
      TabIndex        =   15
      Tag             =   "0058|0116|0233|0466|0932|1865|3729"
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox BlackPianoKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   1
      Left            =   840
      ScaleHeight     =   825
      ScaleWidth      =   225
      TabIndex        =   14
      Tag             =   "0052|0104|0208|0415|0831|1661|3322"
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox BlackPianoKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   0
      Left            =   480
      ScaleHeight     =   825
      ScaleWidth      =   225
      TabIndex        =   13
      Tag             =   "0046|0092|0185|0370|0740|1480|2960"
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox PianoKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   12
      Left            =   4560
      ScaleHeight     =   1425
      ScaleWidth      =   345
      TabIndex        =   12
      Tag             =   "0147|0294|0587|1175|2349|4699|9397"
      Top             =   240
      Width           =   375
   End
   Begin VB.PictureBox PianoKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   11
      Left            =   4200
      ScaleHeight     =   1425
      ScaleWidth      =   345
      TabIndex        =   11
      Tag             =   "0131|0262|0523|1046|2093|4186|8372"
      Top             =   240
      Width           =   375
   End
   Begin VB.PictureBox PianoKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   10
      Left            =   3840
      ScaleHeight     =   1425
      ScaleWidth      =   345
      TabIndex        =   10
      Tag             =   "0123|0247|0494|0988|1975|3951|7902"
      Top             =   240
      Width           =   375
   End
   Begin VB.PictureBox PianoKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   9
      Left            =   3480
      ScaleHeight     =   1425
      ScaleWidth      =   345
      TabIndex        =   9
      Tag             =   "0110|0220|0440|0880|1760|3520|7040"
      Top             =   240
      Width           =   375
   End
   Begin VB.PictureBox PianoKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   8
      Left            =   3120
      ScaleHeight     =   1425
      ScaleWidth      =   345
      TabIndex        =   8
      Tag             =   "0098|0196|0392|0784|1568|3136|6272"
      Top             =   240
      Width           =   375
   End
   Begin VB.PictureBox PianoKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   7
      Left            =   2760
      ScaleHeight     =   1425
      ScaleWidth      =   345
      TabIndex        =   7
      Tag             =   "0087|0175|0349|0698|1397|2794|5588|2794|5588"
      Top             =   240
      Width           =   375
   End
   Begin VB.PictureBox PianoKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   6
      Left            =   2400
      ScaleHeight     =   1425
      ScaleWidth      =   345
      TabIndex        =   6
      Tag             =   "0082|0165|0330|0659|1318|2637|5270"
      Top             =   240
      Width           =   375
   End
   Begin VB.PictureBox PianoKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   4
      Left            =   1680
      ScaleHeight     =   1425
      ScaleWidth      =   345
      TabIndex        =   4
      Tag             =   "0065|0131|0262|0523|1046|2093|4186"
      Top             =   240
      Width           =   375
   End
   Begin VB.PictureBox PianoKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   3
      Left            =   1320
      ScaleHeight     =   1425
      ScaleWidth      =   345
      TabIndex        =   3
      Tag             =   "0062|0123|0247|0494|0988|1975|3951"
      Top             =   240
      Width           =   375
   End
   Begin VB.PictureBox PianoKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   2
      Left            =   960
      ScaleHeight     =   1425
      ScaleWidth      =   345
      TabIndex        =   2
      Tag             =   "0055|0110|0220|0440|0880|1760|3520"
      Top             =   240
      Width           =   375
   End
   Begin VB.PictureBox PianoKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   1
      Left            =   600
      ScaleHeight     =   1425
      ScaleWidth      =   345
      TabIndex        =   1
      Tag             =   "0049|0098|0196|0392|0784|1568|3136"
      Top             =   240
      Width           =   375
   End
   Begin VB.PictureBox PianoKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   0
      Left            =   240
      ScaleHeight     =   1425
      ScaleWidth      =   345
      TabIndex        =   0
      Tag             =   "0044|0087|0175|0349|0698|1397|2794"
      Top             =   240
      Width           =   375
   End
   Begin VB.PictureBox PianoKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   5
      Left            =   2040
      ScaleHeight     =   1425
      ScaleWidth      =   345
      TabIndex        =   5
      Tag             =   "0073|0147|0294|0587|1175|2349|4699"
      Top             =   240
      Width           =   375
   End
   Begin MSComctlLib.Slider Transpose 
      Height          =   195
      Left            =   240
      TabIndex        =   34
      Top             =   1800
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   344
      _Version        =   393216
      LargeChange     =   1
      Max             =   6
      SelStart        =   3
      Value           =   3
   End
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   285
      Left            =   4261
      TabIndex        =   42
      Top             =   3480
      Width           =   240
      _ExtentX        =   291
      _ExtentY        =   503
      _Version        =   393216
      Value           =   10
      BuddyControl    =   "VNoteTime"
      BuddyDispid     =   196609
      OrigLeft        =   1200
      OrigTop         =   3360
      OrigRight       =   1440
      OrigBottom      =   3855
      Max             =   20
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin etUCP.chameleonButton Command2 
      Height          =   375
      Left            =   2760
      TabIndex        =   45
      Top             =   2160
      Width           =   1575
      _ExtentX        =   2778
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
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "RCXPiano.frx":06F4
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
      Caption         =   "Note Time"
      Height          =   255
      Left            =   3000
      TabIndex        =   40
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Wait Time"
      Height          =   255
      Left            =   960
      TabIndex        =   38
      Top             =   3480
      Width           =   1335
   End
End
Attribute VB_Name = "RCXPiano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NoteInt, NoteTime As Integer

Dim TempNotes As New Collection
Dim TempWaits As New Collection

Private Sub BlackPianoKey_Click(Index As Integer)
    On Error Resume Next
    NoteInt = Mid(BlackPianoKey(Index).Tag, (Transpose.Value * 5) + 1, 4)

    NoteInt = Int(NoteInt)

    TempNotes.Add NoteInt
    TempWaits.Add (Int(Me.VNoteTime.text) * NoteTime)
    TempNotes.Add 1
    TempWaits.Add (Val(VWaitTime.text) * NoteTime)

    If UseNQC.Value = True Then
        Notes.text = Notes.text & vbNewLine & "PlayTone(" & NoteInt & "," & NoteTime & "*__NOTETIME); Wait(" & NoteTime & "*__WAITTIME);"
    End If

    If UseLASM.Value = True Then
        Notes.text = Notes.text & vbNewLine & "playt " & NoteInt & "," & (Val(VNoteTime.text) * NoteTime) & vbNewLine & "wait 2," & (NoteTime * Int(VWaitTime.text))
    End If

    If UseMScript.Value = True Then
        Notes.text = Notes.text & vbNewLine & "tone " & NoteInt & " for " & (Val(VNoteTime.text) * NoteTime) & vbNewLine & "wait " & (NoteTime * Int(VWaitTime.text))
    End If

    Notes.SelStart = Len(Notes.text)

    If COMOPEN = True Then If About.Spirit.TowerAndCableConnected And About.Spirit.PBAliveOrNot Then About.Spirit.PlayTone NoteInt, (Int(VNoteTime.text) * NoteTime)
End Sub

Private Sub chameleonButton1_Click()
    If Me.Height = 6030 Then
        Me.Height = 4200
    Else
        Me.Height = 6030
    End If
End Sub

Private Sub Command1_Click()
    If UseNQC.Value = True Then
        Notes.text = Notes.text & vbNewLine & "Wait(" & NoteTime & "*__WAITTIME);"
    End If

    If UseLASM.Value = True Then
        Notes.text = Notes.text & vbNewLine & "Wait 2," & (Val(VWaitTime.text) * NoteTime)
    End If

    If UseMScript.Value = True Then
        Notes.text = Notes.text & vbNewLine & "Wait " & (Val(VWaitTime.text) * NoteTime)
    End If

    TempNotes.Add 1
    TempWaits.Add (Val(VWaitTime.text) * NoteTime)
End Sub

Private Sub Command2_Click()
    Clipboard.Clear

    If Check1.Value = 1 And Check1.Enabled = True Then
        Clipboard.SetText "task main() {" & vbNewLine & "#define __NOTETIME   " & Val(VNoteTime.text) & vbNewLine & "#define __WAITTIME   " & Val(VWaitTime.text) & vbNewLine & vbNewLine & Notes.text & vbNewLine & "}"
    Else
        Clipboard.SetText "#define __NOTETIME   " & Val(VNoteTime.text) & vbNewLine & "#define __WAITTIME   " & Val(VWaitTime.text) & vbnewlien & vbNewLine & Notes.text
    End If
End Sub

Private Sub Command3_Click()
    Notes.text = ""

    On Error Resume Next

    For I = 1 To 5
        For X = 0 To TempNotes.Count
            TempNotes.Remove X
        Next
        For X = 0 To TempWaits.Count
            TempWaits.Remove X
        Next
    Next I

End Sub

Private Sub Command4_Click()
    If UseNQC.Value = True Then Check1.Enabled = True Else Check1.Enabled = False

    Frame2.Visible = False

    Frame1.Enabled = True
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    Check1.Enabled = True

    VWaitTime.Enabled = True
    VNoteTime.Enabled = True
    UpDown1.Enabled = True
    UpDown2.Enabled = True

    For I = 0 To 13
        PianoKey(I).Enabled = True
        PianoKey(I).Backcolor = vbWhite
    Next

    For I = 0 To 9
        BlackPianoKey(I).Enabled = True
        BlackPianoKey(I).Backcolor = vbBlack
    Next

    Option1.Enabled = True
    Option2.Enabled = True
    Option3.Enabled = True
    Option4.Enabled = True
    Option5.Enabled = True
End Sub

Private Sub Command5_Click()
    On Error Resume Next
    If COMOPEN = True Then
        If About.Spirit.TowerAndCableConnected And About.Spirit.PBAliveOrNot Then
            For I = 1 To TempNotes.Count
                About.Spirit.PlayTone TempNotes.Item(I), TempWaits.Item(I)
            Next I
        Else
            XPLib.XPMsgBox "RCX Not Detected!", "Play Tones", False, XP_OKOnly, msg_Exclamation
        End If
    End If
End Sub

Private Sub Form_Load()
LogText "Load - RCXPIANO"
    Me.Hide

    Me.Height = 4200

    DrawXPCtl Me

    Command1.Backcolor = RGB(236, 235, 230)

    Frame1.Enabled = False
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Check1.Enabled = False

    UpDown1.Enabled = False
    UpDown2.Enabled = False
    VWaitTime.Enabled = False
    VNoteTime.Enabled = False

    For I = 0 To 13
        PianoKey(I).Enabled = False
        PianoKey(I).Backcolor = RGB(160, 160, 160)
    Next

    For I = 0 To 9
        BlackPianoKey(I).Enabled = False
        BlackPianoKey(I).Backcolor = RGB(160, 160, 160)
    Next

    Option1.Enabled = False
    Option2.Enabled = False
    Option3.Enabled = False
    Option4.Enabled = False
    Option5.Enabled = False

    Frame2.Visible = True

    NoteTime = Option3.Tag
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3

    Me.Show


    ' REMOVE FOR LANGUAGE SELECTION
    Command4_Click
End Sub

Private Sub Option1_Click()
    NoteTime = Option1.Tag
End Sub

Private Sub Option2_Click()
    NoteTime = Option2.Tag
End Sub

Private Sub Option3_Click()
    NoteTime = Option3.Tag
End Sub

Private Sub Option4_Click()
    NoteTime = Option4.Tag
End Sub

Private Sub Option5_Click()
    NoteTime = Option5.Tag
End Sub

Private Sub PianoKey_Click(Index As Integer)
    On Error Resume Next
    NoteInt = Mid(PianoKey(Index).Tag, (Transpose.Value * 5) + 1, 4)

    If Index = 13 And Transpose.Value = Transpose.Max Then NoteInt = Mid(PianoKey(Index).Tag, (Transpose.Value * 5) + 1, 5)

    NoteInt = Int(NoteInt)

    TempNotes.Add NoteInt
    TempWaits.Add (Int(Me.VNoteTime.text) * NoteTime)
    TempNotes.Add 1
    TempWaits.Add (Val(VWaitTime.text) * NoteTime)

    If UseNQC.Value = True Then
        Notes.text = Notes.text & vbNewLine & "PlayTone(" & NoteInt & "," & NoteTime & "*__NOTETIME); Wait(" & NoteTime & "*__WAITTIME);"
    End If

    If UseLASM.Value = True Then
        Notes.text = Notes.text & vbNewLine & "playt " & NoteInt & "," & (Val(VNoteTime.text) * NoteTime) & vbNewLine & "wait 2," & (NoteTime * Int(VWaitTime.text))
    End If

    If UseMScript.Value = True Then
        Notes.text = Notes.text & vbNewLine & "tone " & NoteInt & " for " & (Val(VNoteTime.text) * NoteTime) & vbNewLine & "wait " & (NoteTime * Int(VWaitTime.text))
    End If

    Notes.SelStart = Len(Notes.text)

    If COMOPEN = True Then If About.Spirit.TowerAndCableConnected And About.Spirit.PBAliveOrNot Then About.Spirit.PlayTone NoteInt, (Int(VNoteTime.text) * NoteTime)
End Sub

