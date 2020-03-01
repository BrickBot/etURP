VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form MMap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Memory Map"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin etUCP.chameleonButton Command1 
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   7200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Refresh"
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
      MICON           =   "MemMap.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer DelayLoad 
      Interval        =   500
      Left            =   1800
      Top             =   7320
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   12303
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"MemMap.frx":001C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "MMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Map As Variant

Private Sub Command1_Click()
    DelayLoad_Timer
End Sub

Private Sub DelayLoad_Timer()
    Me.Caption = "Memory Map - " & BrickType

    Map = About.Spirit.MemMap
    Me.RichTextBox1.text = ""

    temp = Int(Right(Me.Caption, 1))

    Select Case temp
        Case 1
            GetRCXMap
        Case 2
            GetRCXMap
        Case Is > 2
            RichTextBox1.Font.Bold = True
            RichTextBox1.text = "                                       ***Brick Memory Map***" & vbNewLine
            RichTextBox1.Font.Bold = False
            RichTextBox1.text = RichTextBox1.text & vbNewLine & "Memory Used = " & Map(9)
            RichTextBox1.text = RichTextBox1.text & vbNewLine & "Top Of Memory = " & Map(10)
            RichTextBox1.text = RichTextBox1.text & vbNewLine & "Memory Avalible = " & Int(Map(10)) - Int(Map(9))
            Me.Caption = "Memory Map - Brick"
    End Select

    DelayLoad.Enabled = False
End Sub

Private Sub Form_Load()
LogText "Load - MMAP"
    Me.Visible = False
    DrawXPCtl Me
    Me.Icon = MSprogrammer.Icon
    RichTextBox1.text = "Loading, Please Wait..."
    Me.Visible = True
End Sub

Sub GetRCXMap()
    RichTextBox1.text = "                                   ***RCX" & Int(BrickType) & " Memory Map***"
    AddSubPointers
    RichTextBox1.text = RichTextBox1.text & vbNewLine & "Datalog Start Pointer = " & Map(91)
    RichTextBox1.text = RichTextBox1.text & vbNewLine & "Datalog Current Pointer = " & Map(92)
    RichTextBox1.text = RichTextBox1.text & vbNewLine & "Memory Used = " & Map(93)
    RichTextBox1.text = RichTextBox1.text & vbNewLine & "Top Of Memory = " & Map(94)
    RichTextBox1.text = RichTextBox1.text & vbNewLine & "Memory Avalible = " & Int(Map(94)) - Int(Map(93))
    Me.Caption = Me.Caption & "."
    Me.Caption = "Memory Map - RCX/RCX2"
End Sub

Function AddSubPointers()
    Me.Caption = "Memory Map - LOADING: ."
    AddLine " "
    AddLine "                                       Subroutine Pointers:", False
    AddLine vbNewLine, False

    For I = 0 To 4
        Me.Caption = Me.Caption & "."
        AddLine "Program " & (I + 1) & ":"
        For s = 0 To 7
            z = Int(1 + 8 * I + s)
            AddLine Map(z) & " ", False
        Next s
    Next I

    AddLine " "

    AddLine "                                           Task Pointers:", False
    AddLine vbNewLine, False

    For I = 0 To 4
        Me.Caption = Me.Caption & "."
        AddLine "Program " & (I + 1) & ":"
        For s = 0 To 9
            z = Int(41 + 10 * I + s)
            AddLine Map(z) & " ", False
        Next s
    Next I

    AddLine " "
End Function

Function AddLine(text As String, Optional EndNewLine As Boolean = True)
    If EndNewLine = True Then temp = vbNewLine
    RichTextBox1.text = RichTextBox1.text & temp & text & temp
End Function

Private Sub Form_Unload(Cancel As Integer)
    DelayLoad.Enabled = True
End Sub
