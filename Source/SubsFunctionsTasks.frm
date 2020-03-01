VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E77A939D-0153-4E20-A213-57934F00CE3D}#1.0#0"; "SIDEMENUS.OCX"
Begin VB.Form SubsFunctionsTasks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subs/Tasks"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin XPSideMenus.XPsidemenu XPSM 
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5953
      Speed           =   0
   End
   Begin etUCP.chameleonButton chameleonButton1 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   3600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "SubsFunctionsTasks.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ImageList SFTimages 
      Left            =   1680
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   12632256
      ImageWidth      =   11
      ImageHeight     =   12
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SubsFunctionsTasks.frx":001C
            Key             =   "Sub"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SubsFunctionsTasks.frx":0220
            Key             =   "Func"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SubsFunctionsTasks.frx":0424
            Key             =   "Task"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "SubsFunctionsTasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SubsLines, TasksLines, FuncLines As Integer
Dim CurrentForm As frmDocument
Dim Funcs, Subs, Tasks As Integer

Sub GotoTSF(ActiveForm As frmDocument)
    Me.Show
    Me.Visible = False
    GetTSFNames ActiveForm
    Set CurrentForm = ActiveForm
    Me.Visible = True
End Sub

Sub GetTSFNames(ActiveForm As frmDocument)

    SubsLines = 1
    TasksLines = 1
    FuncLines = 1
    
 Funcs = 0
 Subs = 0
 Tasks = 0
    
    ActiveForm.Visible = False
    If ActiveForm.Caption = "Program #" Then
    Unload ActiveForm
    Set ActiveForm = Nothing
    End If
    
    If ActiveForm Is Nothing Then
        XPLib.XPMsgBox "No currently active program.", "Goto Func/Sub/Task...", False, XP_OKOnly, msg_Critical
        Me.Hide
        Exit Sub
    End If
    
    ActiveForm.Visible = True

    On Error Resume Next

    Me.Icon = MSprogrammer.Icon
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3

    s = ActiveForm.rtftext.text

    Dim lStartPos As Long
    Dim lEndPos As Long
    
    lStartPos = 1
    
    s = s & vbCrLf
    lEndPos = InStr(lStartPos, s, vbCrLf)
    Do While lEndPos > 0
        ParseLine Mid(s, lStartPos, lEndPos - lStartPos), RTBPos + lStartPos
        SubsLines = SubsLines + 1
        TasksLines = TasksLines + 1
        FuncLines = FuncLines + 1
        lStartPos = lEndPos + Len(vbCrLf)
        lEndPos = InStr(lStartPos, s, vbCrLf)
        DoEvents
    Loop

If Tasks = 0 Then Me.XPSM.ChangePANEL "T", "T", "Tasks (task...)", fixed
If Subs = 0 Then Me.XPSM.ChangePANEL "S", "S", "Subroutines (sub...)", fixed
If Funcs = 0 Then Me.XPSM.ChangePANEL "F", "F", "Functions (void...)", fixed
End Sub


Private Function ParseLine(Data, LEnd)
    For I = 1 To Len(Data)
        If Mid(Data, I, 2) = "//" Or Mid(Data, I, 1) = "*" Then Exit Function
        If LCase(Mid(Data, I, 4)) = "task" Then
            For z = I + 1 To Len(Data)
                If Mid(Data, z, 1) = " " Or Mid(Data, z, 1) = "{" Or Mid(Data, z, 1) = Chr(13) Then
                    XPSM.AddHyper "T" & TasksLines, "T", "(Line " & TasksLines & "): " & RemoveJunk(Mid(Data, (I + 5), Len(Data))), True, Hyperlink
                    Tasks = Tasks + 1
                    Exit Function
                End If
            Next z
        End If
    Next I

    For I = 1 To Len(Data)
        If Mid(Data, I, 2) = "//" Or Mid(Data, I, 1) = "*" Then Exit Function
        If LCase(Mid(Data, I, 3)) = "sub" Then
            For z = I + 1 To Len(Data)
                If Mid(Data, z, 1) = " " Or Mid(Data, z, 1) = "{" Or Mid(Data, z, 1) = Chr(13) Then
                    XPSM.AddHyper "S" & SubsLines, "S", "(Line " & SubsLines & "): " & RemoveJunk(Mid(Data, (I + 4), Len(Data))), True, Hyperlink
                    Subs = Subs + 1
                    Exit Function
                End If
            Next z
        End If
    Next I
    
 For I = 1 To Len(Data)
        If Mid(Data, I, 2) = "//" Or Mid(Data, I, 1) = "*" Then Exit Function
        If LCase(Mid(Data, I, 4)) = "void" Then
            For z = I + 1 To Len(Data)
                If Mid(Data, z, 1) = " " Or Mid(Data, z, 1) = "{" Or Mid(Data, z, 1) = Chr(13) Then
                    XPSM.AddHyper "F" & FuncLines, "F", "(Line " & FuncLines & "): " & RemoveJunk(Mid(Data, (I + 5), Len(Data))), True, Hyperlink
                    Funcs = Funcs + 1
                    Exit Function
                End If
            Next z
        End If
    Next I
End Function


Function RemoveJunk(text As String)
    For I = 1 To Len(text)
        If Right(text, 1) = Chr(13) Then
            RemoveJunk = Mid(text, 1, (Len(text) - 1))
        End If

        If Asc(Mid(text, I, 1)) <> 13 And Asc(Mid(text, I, 1)) <> 10 And Mid(text, I, 1) <> "{" And Mid(text, I, 1) <> " " And Mid(text, I, 1) <> "(" Then
            temptxt = temptxt & Mid(text, I, 1)
        Else
            RemoveJunk = temptxt
            Exit Function
        End If
    Next

    RemoveJunk = temptxt
End Function

Private Sub chameleonButton1_Click()
    XPSM.RemoveAll
    XPSM.Addpanel "F", "Functions (void...)", Closed, False
    XPSM.Addpanel "S", "Subroutines (sub...)", Closed, False
    XPSM.Addpanel "T", "Tasks (task...)", Closed, False
    GetTSFNames CurrentForm
End Sub

Private Sub Form_Load()
LogText "Load - SBS/FNCS/TSKS"
    
    XPSM.RemoveAll
    XPSM.Addpanel "T", "Tasks (task...)", Closed, False
    XPSM.Addpanel "F", "Functions (void...)", Closed, False
    XPSM.Addpanel "S", "Subroutines (sub...)", Closed, False
End Sub

Private Sub Image1_Click()

End Sub

Private Sub XPSM_HyperClick(key As String)
key = Mid(key, 2)
CurrentForm.GoToLine Int(key)
End Sub
