VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RegSvr - En-Tech Easy Register (*.ocx, *.dll)"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox Files 
      Height          =   1455
      Left            =   120
      Pattern         =   "*.ocx;*.dll"
      TabIndex        =   1
      Top             =   1920
      Width           =   6255
   End
   Begin VB.TextBox Log 
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "REGISTER.frx":0000
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SItem

Private Sub Form_Load()
Log.Text = "Registering Components:"
Files.Path = App.Path
On Error Resume Next
For i = 0 To Files.ListCount - 1
Files.Selected(i) = True
SItem = "regsvr32.exe " & Chr(34) & App.Path & Files.FileName & Chr(34) & " /s"
If Right(SItem, 3) <> "OCX" Or Right(SItem, 3) <> "DLL" Then
Shell SItem
Log.Text = Log.Text & vbNewLine & "REGSVR32.exe: " & Files.FileName
End If
Next

Shell App.Path & "Regocx32.exe *.ocx"

Log.Text = Log.Text & vbNewLine & "Registering Complete."
End Sub

Private Sub Text1_Change()

End Sub
