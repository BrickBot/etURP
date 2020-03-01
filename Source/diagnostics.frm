VERSION 5.00
Begin VB.Form diagnostics 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Diagnostics - Updated Every 5 Secs."
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "RCX"
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   3495
      Begin VB.Label FreeMem 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1320
         TabIndex        =   15
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Free Memory:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label prgm 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Currently Selected Program:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label version 
         BackColor       =   &H00C0C0C0&
         Caption         =   "------------"
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Version (ROM/RAM):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Power (mv):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tower"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.Label Comport 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Timer AutoUpdate 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3120
      Top             =   1920
   End
   Begin VB.Frame Frame3 
      Caption         =   "RCX"
      Height          =   1455
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Width           =   3255
   End
End
Attribute VB_Name = "diagnostics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private XPManifestLib As ClsXPManifest


Private Sub AutoUpdate_Timer()
    On Error Resume Next
    
Me.Caption = "Diagnostics - Updating..."
    
    If BrickType = RCX Or BrickType = RCX2 Then Frame2.Visible = True Else Frame2.Visible = False
    temp1 = ""
    temp2 = ""
    If About.Spirit.TowerAndCableConnected Then temp1 = "Tower Connected"
    If About.Spirit.TowerAlive Then temp2 = "Tower Alive"
    If temp1 <> "" And temp2 <> "" Then
        Label2.Caption = temp1 & ", " & temp2
    Else
        Label2.Caption = temp1 & temp2
    End If

    Comport.Caption = About.Spirit.ComPortNo

    If BrickType = RCX Or BrickType = RCX2 Then
        If Int(Label4.Caption) <> 9000 Then Label4.Caption = Label4.Caption & " (" & Mid(((100 / 9000) * Int(Label4.Caption)), 1, 2) & "%)" Else Label4.Caption = Label4.Caption & " (100%)"
        If About.Spirit.PBAliveOrNot Then
            Label6.Caption = "Active"
            SetVersionText
            Label4.Caption = About.Spirit.PBBattery
            prgm.Caption = About.Spirit.Poll(8, 0) + 1
        Else
            Label6.Caption = "Not On/Out of Range"
            version.Caption = "?"
            Label4.Caption = "?"
            prgm.Caption = "?"
            version.Caption = "?"
        End If
        Map = About.Spirit.MemMap
        FreeMem.Caption = Int(Map(94)) - Int(Map(93))
    Else
        Label6.Caption = "Not an RCX brick."
        version.Caption = "?"
        Label4.Caption = "?"
        prgm.Caption = "?"
        version.Caption = "?"
    End If
    
Me.Caption = "Diagnostics - Updated Every 5 Secs."
End Sub

Private Sub Form_Load()
LogText "Load - DIAGNOSTICS"
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
 End Sub

Function SetVersionText()
    version.Caption = About.Spirit.UnlockPBrick
    If BrickType = RCX Then version.Caption = version.Caption & " (RCX)"
    If BrickType = RCX2 Then version.Caption = version.Caption & " (RCX2)"
End Function

Sub ActivateDiagnostics()
Me.AutoUpdate.Enabled = True
AutoUpdate_Timer
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.AutoUpdate.Enabled = False
End Sub

