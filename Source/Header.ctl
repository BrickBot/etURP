VERSION 5.00
Begin VB.UserControl Header 
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   2325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4665
   MousePointer    =   99  'Custom
   ScaleHeight     =   2325
   ScaleWidth      =   4665
   Begin VB.Image Image1 
      Height          =   375
      Left            =   120
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LBLCAPTION 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C65D21&
      Height          =   210
      Left            =   600
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   200
      Width           =   2775
   End
   Begin VB.Image imgright 
      Height          =   375
      Left            =   3000
      MousePointer    =   99  'Custom
      Picture         =   "Header.ctx":0000
      Top             =   120
      Width           =   1590
   End
   Begin VB.Shape pichead 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1560
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Header"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event click()
Public numpanel                         As Integer
Private pnlstate                        As P_state

Public Property Let Caption(newcaption As String)

LBLCAPTION.Caption = newcaption

End Property

Public Property Let Headerstyle(newvalue As P_state)

pnlstate = newvalue

If numpanel = 1 Then

    Select Case newvalue
        Case Opened
            imgright.Picture = LoadResPicture(104, 0)
        Case Closed
            imgright.Picture = LoadResPicture(105, 0)
        Case Fixed
            imgright.Picture = LoadResPicture(106, 0)
    End Select
            pichead.FillColor = &HC45518
            pichead.BorderColor = &HC45518
            LBLCAPTION.ForeColor = &HFFFFFF

Else

    Select Case newvalue
        Case Opened
            imgright.Picture = LoadResPicture(101, 0)
        Case Closed
            imgright.Picture = LoadResPicture(102, 0)
        Case Fixed
            imgright.Picture = LoadResPicture(103, 0)
    End Select
            pichead.FillColor = &HFFFFFF
            pichead.BorderColor = &HFFFFFF
            LBLCAPTION.ForeColor = &HC65D21

End If

End Property

Public Function MOVECAPTION(moved As Boolean)
If moved = True Then
    LBLCAPTION.Left = 120
Else
    LBLCAPTION.Left = 600
End If

End Function

Private Sub ImgCmd_Click()

RaiseEvent click

End Sub

Private Sub ImgCmd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If pnlstate = Fixed Then

    ImgCmd.MousePointer = 0

End If

End Sub

Public Function SethPic(pic As StdPicture)

Set Image1.Picture = pic

End Function

Private Sub Image1_Click()
RaiseEvent click
End Sub

Private Sub imgright_Click()
    RaiseEvent click
End Sub

Private Sub LBLCAPTION_Click()
    RaiseEvent click
End Sub

Private Sub UserControl_Click()
    RaiseEvent click
End Sub

Private Sub UserControl_Resize()

On Error Resume Next
With UserControl
    .imgright.Move .ScaleWidth - imgright.Width, 120
    .pichead.Move 0, 120, .ScaleWidth - .imgright.Width, .imgright.Height
    .Height = .imgright.Top + .imgright.Height
End With

End Sub

