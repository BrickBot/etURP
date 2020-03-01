VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDocument 
   Caption         =   "Program #"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3690
   ScaleWidth      =   4680
   Tag             =   "PROGRAM"
   Begin etUCP.rtbSyntax rtftext 
      Height          =   2895
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5106
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmDocument.frx":030A
      RightMargin     =   1.00000e5
   End
   Begin MSComctlLib.ImageList ProgIcons 
      Left            =   240
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":0642
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox Errors 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Errors encountered in your program. Double-Click to hide."
      Top             =   3000
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1085
      _Version        =   393217
      BackColor       =   12632256
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmDocument.frx":0A96
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
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sFile As String
Dim ValTemp As Integer
Dim MeChanged As Boolean
Dim PRIVATELineMem As Integer

Private Sub Errors_DblClick()
    Errors.Visible = False
    Form_Resize
End Sub

Private Sub Form_Activate()
rtftext_SelChange
End Sub

Private Sub Form_GotFocus()
rtftext_SelChange
End Sub

Private Sub Form_Load()
    DrawXPCtl Me

LogText "Load - DOCUMENT"

    If Mid(Me.Caption, 1, 9) <> "Program #" Then
        Me.rtftext.LoadFile Me.Caption
        Me.rtftext.HighlightRefresh
    End If
    
        rtftext.ToolTip = "Current Cursor Position: Line " & rtftext.CountLines(True)
        CurrentLine = Int(rtftext.CountLines(True))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MeChanged = rtftext.IsChanged

    If rtftext.text = "" Then
        Unload Me
    Else
        If MeChanged = True Then
            If rtftext.text = "" Then
                Unload Me
            Else
                sc = XPLib.XPMsgBox("Save Changes?", Me.Caption, False, XP_Custom, msg_Question, "Discard", "Save")
                If sc = False Then
                    Unload Me
                ElseIf sc = True Then
                    Dim sFile As String
                    If Left$(Me.Caption, 8) = "Document" Then
                        With dlgCommonDialog
                            .DialogTitle = "Save"
                            .CancelError = False
                            .Filter = "Not Quite C Files (*.nqc, *.nqh)|*.nqc;*.nqh|LASM Files (*.asm)|*.asm|Mindscript Files (*.rcx2, *.lsc)|*.rcx2;*.lsc|All Files (*.*)|*.*"
                            .ShowSave
                            If Len(.FileName) = 0 Then
                                Exit Sub
                            End If
                            sFile = .FileName
                        End With
                        Me.rtftext.SaveFile sFile
                    Else
                        sFile = Me.Caption
                        Me.rtftext.SaveFile sFile
                    End If
                Else
                    Exit Sub
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Errors.Visible = True Then
        rtftext.Left = 10
        rtftext.Top = 10
        rtftext.Height = Me.Height - 1550
        rtftext.Width = Me.Width - 150
        Errors.Top = Me.Height - 1500
        Errors.Height = Me.Height - Errors.Top - 400
        Errors.Left = 10
        Errors.Width = Me.Width - 150
    Else
        rtftext.Left = 10
        rtftext.Top = 10
        rtftext.Height = Me.Height - 400
        rtftext.Width = Me.Width - 150
    End If
End Sub

Sub GoToLine(LineNumber)
    On Error Resume Next
    LineNumber = LineNumber + 1
    If LineNumber = 2 Then
        rtftext.SelStart = 1
        Me.SetFocus
        Exit Sub
    End If

    lines = 1
    For I = 1 To Len(rtftext.text)
        If Asc(Mid(rtftext.text, I, 1)) = 13 Then lines = lines + 1
        If lines = LineNumber - 1 Then
            rtftext.SelStart = I + 1
            Me.SetFocus
            Exit Sub
        End If
    Next

    LineNumber = LineNumber - 1

    ValTemp = GetSetting("En-Tech URP", "Options", "GoToLineErrors", 1)
    If ValTemp = 1 Then
        XPLib.XPMsgBox "Line number " & LineNumber & " not found." & vbNewLine & "Maximum line number in current program is " & lines & ".", "GoTo Line...", False, XP_OKOnly, msg_Exclamation
        Me.rtftext.SetFocus
    End If
End Sub

Function ShowErrorsBox()
    Errors.Visible = True
    Form_Resize
    rtftext.SetFocus
End Function

Function RefreshColour()
    rtftext.HighlightRefresh
End Function

Public Function ReplaceProgramText(Word As String, Replace As String, CaseSensitive As Boolean) As String
entire = rtftext.text

If Word = Replace Then
XPLib.XPMsgBox "Replace Text Cannot be Identical to Find Text!", "Replace...", False, XP_OKOnly, msg_Exclamation
Exit Function
End If

If UCase(Word) = UCase(Replace) Then CaseSensitive = True

Dim Replaces As Integer
    
    Dim I As Integer
    I = 1
    Dim LeftPart
    Do While True
If CaseSensitive Then
        I = InStr(1, entire, Word, vbBinaryCompare)
Else
        I = InStr(1, entire, Word, vbTextCompare)
End If
        If I = 0 Then
            Exit Do
        Else
            LeftPart = Left(entire, I - 1)
            entire = LeftPart & Replace & Right(entire, Len(entire) - Len(Word) - Len(LeftPart))
            Replaces = Replaces + 1
            If Replaces > 10000 Then Exit Do
        End If
    Loop
rtftext.text = entire

XPLib.XPMsgBox Replaces & " Replaces were made.", "Replace...", False, XP_OKOnly, msg_Information
End Function

Function FindProgramText(text As String, CaseSensitive As Boolean, Again As Boolean)
Static inp As String

inp = rtftext.text

If CaseSensitive = False Then
text = UCase(text)
inp = UCase(inp)
End If

On Error Resume Next

If Again = True Then
For I = (rtftext.SelStart + Len(text)) To Len(rtftext.text)
If Mid(inp, I, Len(text)) = text Then
On Error GoTo FixErr
z = I - 1
If z < 0 Then z = 0
rtftext.SelStart = z
rtftext.SelLength = Len(text)
Exit For
End If
Next I
Else
For I = 1 To Len(rtftext.text)
If Mid(inp, I, Len(text)) = text Then
z = I - 1
If z < 0 Then z = 0
On Error GoTo FixErr
rtftext.SelStart = z
rtftext.SelLength = Len(text)
Exit For
End If
Next I
End If

Me.rtftext.SetFocus
Exit Function
FixErr:
rtftext.SelStart = I
rtftext.SelLength = Len(text)
Me.rtftext.SetFocus
End Function

Private Sub rtftext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If rtftext.SelStart <> Int(PRIVATELineMem) Then
rtftext_SelChange
End If
End Sub

Private Sub rtftext_SelChange()
If Mid(Me.Caption, 1, 9) <> "(LOADING)" Then
LCount = rtftext.CountLines(True)
rtftext.ToolTip = "Current Cursor Position: Line " & LCount
CurrentLine = LCount
PRIVATELineMem = LCount
End If
End Sub

