VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "En-Tech URP"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9885
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "etURP"
   Moveable        =   0   'False
   NegotiateToolbars=   0   'False
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   5520
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   31
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":041E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0532
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0646
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":075A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":086E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0982
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A96
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0EE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":132E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":177A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1FC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2BB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2FB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":33AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":37AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3BA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3FA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":439E
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":479A
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4B96
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4EB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":52AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":56AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5AA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5EA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":629E
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":669A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList SystemImg 
      Left            =   4920
      Top             =   1440
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
            Picture         =   "frmMain.frx":6A96
            Key             =   "etURP"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6AF4
            Key             =   "Program"
         EndProperty
      EndProperty
   End
   Begin VB.Timer TBCheck 
      Interval        =   330
      Left            =   3960
      Top             =   1440
   End
   Begin MSComctlLib.Toolbar SmallTB 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FindBrick"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tNQC"
                  Text            =   "NQC"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tLASM"
                  Text            =   "LASM"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tMScript"
                  Text            =   "MindScript"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Compile"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MicroscoutProgrammer"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "RCXPiano"
            ImageIndex      =   22
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   2925
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   10
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7373
            MinWidth        =   441
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   353
            MinWidth        =   353
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Line:"
            TextSave        =   "Line:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   353
            MinWidth        =   353
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   441
            MinWidth        =   441
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            Text            =   "NONE"
            TextSave        =   "NONE"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   2117
            MinWidth        =   2117
            Text            =   "NONE"
            TextSave        =   "NONE"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   4440
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   41
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "Create New Program."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open Program."
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save Program."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print Program."
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FindBrick"
            Object.ToolTipText     =   "Find Brick"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "ChangeCType"
            Object.ToolTipText     =   "Change the Code Compiler Used."
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tNQC"
                  Text            =   "NQC"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tLASM"
                  Text            =   "LASM"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tMScript"
                  Text            =   "MindScript"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Compile"
            Object.ToolTipText     =   "Compile Program."
            ImageIndex      =   8
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Download"
            Object.ToolTipText     =   "Download Program."
            ImageIndex      =   10
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DownloadRun"
            Object.ToolTipText     =   "Download and Run Program."
            ImageIndex      =   11
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Run"
            Object.ToolTipText     =   "Run Selected Program."
            ImageIndex      =   9
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DownloadFirm"
            Object.ToolTipText     =   "Download Firmware."
            ImageIndex      =   28
         EndProperty
         BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Watch"
            Object.ToolTipText     =   "Set the RCX's Time."
            ImageIndex      =   12
         EndProperty
         BeginProperty Button33 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PowerDown"
            Object.ToolTipText     =   "Set the Power Down Time for the RCX."
            ImageIndex      =   18
         EndProperty
         BeginProperty Button34 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Diagnostics"
            Object.ToolTipText     =   "Check the Tower and Brick's Condition."
            ImageIndex      =   13
         EndProperty
         BeginProperty Button35 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ClearMem"
            Object.ToolTipText     =   "Clear Brick Memory."
            ImageIndex      =   15
         EndProperty
         BeginProperty Button36 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MMap"
            Object.ToolTipText     =   "Look at the Brick's Memory Map."
            ImageIndex      =   27
         EndProperty
         BeginProperty Button37 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "RCXVars"
            Object.ToolTipText     =   "Look at the Brick's Variables."
            ImageIndex      =   29
         EndProperty
         BeginProperty Button38 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button39 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DirectControl"
            Object.ToolTipText     =   "Directly Control the Brick."
            ImageIndex      =   19
         EndProperty
         BeginProperty Button40 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MicroscoutProgrammer"
            Object.ToolTipText     =   "Program the MicroScout."
            ImageIndex      =   20
         EndProperty
         BeginProperty Button41 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "RCXPiano"
            Object.ToolTipText     =   "Compose Tunes and Songs for the Brick"
            ImageIndex      =   22
         EndProperty
      EndProperty
      MousePointer    =   4
      Begin VB.ComboBox ProgramNum 
         Height          =   315
         Left            =   4280
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu MNUFileMenuSideBar 
         Caption         =   "{SIDEBAR|TRANSPARENT|BOLD|TEXT:En-Tech<BR>Ultimate Robot Programmer|FCOLOR:vbwhite|UNDERLINE|BCOLOR:&H00F2A762&|GCOLOR:vbBlack}"
      End
      Begin VB.Menu Sep6 
         Caption         =   "-{Raised}Program"
      End
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New{IMG:I1}"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open...{IMG:I2}"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-{Raised}File"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save{IMG:3}"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-{Raised}Printer"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Set&up..."
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print...{IMG:I4}"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-{Raised}"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t{IMG:I5}"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy{IMG:I6}"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste{IMG:I7}"
         Shortcut        =   ^V
      End
      Begin VB.Menu Sep8 
         Caption         =   "-{Raised}Find/Replace"
         Index           =   0
      End
      Begin VB.Menu MNUFindText 
         Caption         =   "Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu MNUReplaceText 
         Caption         =   "Replace..."
         Shortcut        =   ^R
      End
      Begin VB.Menu Sep5 
         Caption         =   "-{Raised}etURP"
      End
      Begin VB.Menu MNUOptions 
         Caption         =   "Options..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu OpSep 
         Caption         =   "-{Raised}Program"
      End
      Begin VB.Menu MNUGoto 
         Caption         =   "Goto Line..."
         Shortcut        =   ^G
      End
      Begin VB.Menu MNUGotoTFS 
         Caption         =   "List Subs/Functions/Tasks..."
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "&New Window"
      End
      Begin VB.Menu mnuWindowBar0 
         Caption         =   "-{Raised}Windows"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu WindowSep 
         Caption         =   "-{Raised}Templates"
      End
      Begin VB.Menu MNUSTemplates 
         Caption         =   "Show Templates"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuCOMPILE 
      Caption         =   "Compile"
      Begin VB.Menu mnuCOMPILEPROGRAM 
         Caption         =   "Compile{IMG:I8}"
         Shortcut        =   {F5}
      End
      Begin VB.Menu MNUDownLoadRun 
         Caption         =   "Download && Run{IMG:I11}"
         Shortcut        =   {F6}
      End
      Begin VB.Menu MNUDownLoad 
         Caption         =   "Download{IMG:I10}"
         Shortcut        =   {F7}
      End
      Begin VB.Menu MNUSelProgram 
         Caption         =   "Program"
         Begin VB.Menu MNUP1 
            Caption         =   "Program 1"
            Checked         =   -1  'True
         End
         Begin VB.Menu MNUP2 
            Caption         =   "Program 2"
         End
         Begin VB.Menu MNUP3 
            Caption         =   "Program 3"
         End
         Begin VB.Menu MNUP4 
            Caption         =   "Program 4"
         End
         Begin VB.Menu MNUP5 
            Caption         =   "Program 5"
         End
      End
   End
   Begin VB.Menu MNUTools 
      Caption         =   "Tools"
      Begin VB.Menu Sep7 
         Caption         =   "-{Raised}Direct Interface/Misc."
         Index           =   0
      End
      Begin VB.Menu MNUmsprgmr 
         Caption         =   "MicroScout Programmer{IMG:I20}"
      End
      Begin VB.Menu MNURcxPiano 
         Caption         =   "RCX Piano{IMG:I22}"
      End
      Begin VB.Menu MNUDControl 
         Caption         =   "Direct Control{IMG:I19}"
      End
      Begin VB.Menu MNURCXvar 
         Caption         =   "RCX Variables{IMG:I29}"
      End
      Begin VB.Menu Sep 
         Caption         =   "-{Raised}Brick Information"
      End
      Begin VB.Menu MNURCXWatch 
         Caption         =   "Set RCX Watch{IMG:I12}"
      End
      Begin VB.Menu MNUSetPDTime 
         Caption         =   "Set RCX Power Down Time{IMG:I18}"
      End
      Begin VB.Menu MNUDiagnostics 
         Caption         =   "Diagnostics{IMG:I13}"
      End
      Begin VB.Menu MNUMemMap 
         Caption         =   "Memory Map{IMG:I27}"
      End
      Begin VB.Menu MNUCLRMEM 
         Caption         =   "Clear RCX Memory{IMG:I15}"
      End
      Begin VB.Menu Sep3 
         Caption         =   "-{Raised}Communication"
      End
      Begin VB.Menu MNUOCom 
         Caption         =   "Open Communication{IMG:I17}"
         Enabled         =   0   'False
      End
      Begin VB.Menu MNUCCom 
         Caption         =   "Close Communication{IMG:I16}"
      End
      Begin VB.Menu Sep4 
         Caption         =   "-{Raised}IR/Radio"
      End
      Begin VB.Menu MNUshortrange 
         Caption         =   "Set RCX to Short Range{IMG:32}"
         Enabled         =   0   'False
      End
      Begin VB.Menu MNUlongrange 
         Caption         =   "Set RCX to Long Range{IMG:33}"
      End
      Begin VB.Menu MNUTOFF 
         Caption         =   "Turn Off RCX"
      End
      Begin VB.Menu sep2 
         Caption         =   "-{Raised}Brick"
      End
      Begin VB.Menu MNUFBrick 
         Caption         =   "Find Brick{IMG:I21}"
      End
      Begin VB.Menu MNUdf 
         Caption         =   "Download Firmware{IMG:30}"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   For Debugging, Change to FALSE to avoid crashes and inability to switch to VB.
Const AllowSubClassMenu As Boolean = True
'---------------------------------------------------------------------------------


Dim SHORTLONGRANGE As Integer
Dim COLOURTEMP
Dim tbc

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Const EM_UNDO = &HC7
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Public Enum CodeType
NQC = 0
LASM = 1
MindScript = 3
End Enum

Dim CompileCodeType As CodeType
Private Const WM_SETREDRAW = &HB

Dim MyTips As New cTips

Dim xHwnd           As Long

Private Sub Command1_Click()
    COMOPEN = True
End Sub

Private Sub MDIForm_Load()
On Error GoTo FailTXT
    
    Randomize

LogText "Main Form Loading:"

Me.Top = 0
Me.Left = 0

LogText "Menu Subclass"

If AllowSubClassMenu = True Then
Set MyTips = New cTips
SetMenu hWnd, imlToolbarIcons, MyTips, lv_MDIchildForm_WithMenus
modMenus.HighlightGradient = True
modMenus.DefaultIcon = Me.SystemImg.ListImages(1).Picture
modMenus.HighlightDisabledMenuItems = True
modMenus.CheckMarksXPstyle = True
End If

    CheckTemplatesUnload = False

LogText "Templates"

    temp = GetSetting("En-Tech URP", "Startup", "AutoShowTemplates", 1)
    If temp = 1 Then Templates.Show: Me.MNUSTemplates.Checked = True

    If fMainForm Is Nothing Then
        Set fMainForm = Me
    End If

LogText "XPLib and Options"

    Set XPLib = New XPInterface

    Me.Icon = MSprogrammer.Icon
    Options.FixOptions

LogText "LOADSTUFF Subroutine"

    LoadStuff

    COMOPENCLOSED = 1

    TBCheck_Timer
    Me.Show

    If MyCommand = "" Then
LogText "Loading New Document"
        LoadNewDoc
    Else
LogText "Loading Command Line Document"
        LoadFileFromCommand
    End If
  
Exit Sub
FailTXT:
XPLib.XPMsgBox "Error Number " & Err.Number & " occurred. " & Err.Description, "etURP Start Error", True, XP_OKOnly, msg_Critical
LogText "ERROR: " & Err.Number & ": " & Err.Description
End Sub


Private Sub LoadNewDoc()
LogText "LOADNEWDOC Subroutine"
    Static lDocumentCount As Long
    Dim frmD As frmDocument
    lDocumentCount = lDocumentCount + 1
    Set frmD = New frmDocument
    frmD.Caption = "Program " & lDocumentCount
    frmD.Show
End Sub

Private Sub LoadFileFromCommand()
    
sFile = Command

    LoadNewDoc
            
    ActiveForm.Caption = "(LOADING) " & sFile
    ActiveForm.Icon = ActiveForm.ProgIcons.ListImages(2).Picture
    ActiveForm.rtftext.Backcolor = RGB(200, 200, 200)
    ActiveForm.rtftext.LoadFile sFile
    ActiveForm.Icon = ActiveForm.ProgIcons.ListImages(1).Picture
    ActiveForm.rtftext.Backcolor = vbWhite
    ActiveForm.rtftext.Tag = sFile
        
    CursorTemp = GetSetting("En-Tech URP", "Options", "Cursor", 1)

    If CursorTemp = 1 Then
        ActiveForm.rtftext.SelStart = 0
        ActiveForm.rtftext.SelLength = 0
    Else
        ActiveForm.rtftext.SelStart = Len(ActiveForm.rtftext.text) - 1
        ActiveForm.rtftext.SelLength = 0
    End If
              
    ActiveForm.Caption = sFile
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Set MyTips = Nothing
    
LogText "*****END - etURP Startup Log (" & Now & ")*****"

    Dim I As Integer
    While Forms.Count > 1
        I = 0
        While Forms(I).Caption = Me.Caption
            I = I + 1
        Wend

        Unload Forms(I)
    Wend
End Sub

Private Sub MNUCCom_Click()
    sbStatusBar.Panels(1).text = "Communication Closed."
    About.Spirit.CloseComm
    COMOPENCLOSED = 0
    MNUOCom.Enabled = True
    MNUCCom.Enabled = False
    tbToolBar.Enabled = False

    Me.MNUdf.Enabled = False
    Me.MNUDownLoad.Enabled = False
    Me.MNUDownLoadRun.Enabled = False
    Me.MNURCXWatch.Enabled = False
    Me.MNUTOFF.Enabled = False
End Sub

Private Sub MNUCLRMEM_Click()
    temp = XPLib.XPMsgBox("Warning! This will erase ALL RCX Programs!", "En-Tech URP", False, XP_Custom, msg_Exclamation, "Cancel", "Erase")
    If temp = True Then
        For I = 0 To 4
            About.Spirit.SelectPrgm I
            About.Spirit.ClearAllEvents
            About.Spirit.DeleteAllSubs
            About.Spirit.DeleteAllTasks
        Next
        About.Spirit.SetDatalog 0
        ProgramNum_Click
    End If
End Sub

Private Sub mnuCOMPILEPROGRAM_Click()
    On Error Resume Next
    ActiveForm.ShowErrorsBox

    TempFileName = Chr(34) & App.Path & "\bin\tempprg.nqc" & Chr(34)
    ASMTempFileName = Chr(34) & App.Path & "\bin\tempprg.asm" & Chr(34)
    LSCTempFileName = Chr(34) & App.Path & "\bin\tempprg.lsc" & Chr(34)

    includepath = ActiveForm.rtftext.Tag
    If includepath = "" Then includepath = App.Path & "\bin\"

    If CompileCodeType = NQC Then
        Kill App.Path & "\bin\tempprg.nqc"
        Open App.Path & "\bin\tempprg.nqc" For Output As #1
        Print #1, ActiveForm.rtftext.text
        Close #1
        MainSubs.ExecCmd "nqc.exe -Etemp.log -I" & Chr(34) & includepath & Chr(34) & " " & TempFileName
        WaitForNQC
        ActiveForm.Errors.LoadFile App.Path & "\temp.log"
    End If

    If CompileCodeType = LASM Then
        Kill App.Path & "\bin\tempprg.asm"
        Open App.Path & "\bin\tempprg.asm" For Output As #1
        Print #1, ActiveForm.rtftext.text
        Close #1

        MainSubs.ExecCmd "lcc32.exe" & BCandLASMBrickTypeHeader & " C=LASM -E=temp.log " & ASMTempFileName
        WaitForLCC
        ActiveForm.Errors.LoadFile App.Path & "\bin\temp.log"
    End If

    If CompileCodeType = MindScript Then
        Kill App.Path & "\bin\tempprg.lsc"
        Open App.Path & "\bin\tempprg.lsc" For Output As #1
        Print #1, ActiveForm.rtftext.text
        Close #1

        MainSubs.ExecCmd "lcc32.exe" & BCandLASMBrickTypeHeader & " C=MindScript -E=temp.log " & LSCTempFileName
        WaitForLCC
        ActiveForm.Errors.LoadFile App.Path & "\bin\temp.log"
    End If

    Kill App.Path & "\bin\tempprg.nqc"
    Kill App.Path & "\bin\tempprg.asm"
    Kill App.Path & "\bin\tempprg.lsc"
    Kill App.Path & "temp.log"
    Kill App.Path & "\bin\temp.log"


    ActiveForm.rtftxt.SetFocus
End Sub

Private Sub MNUDControl_Click()
    DirectControl.Show
End Sub

Private Sub MNUdf_Click()
    If COMOPEN = True Then
        dlgCommonDialog.Filter = "Firmware Files (*.lgo)|*.lgo"
        dlgCommonDialog.ShowOpen
        If dlgCommonDialog.FileName <> "" Then
            If About.Spirit.TowerAndCableConnected And About.Spirit.PBAliveOrNot Then
                FIRMNAME = dlgCommonDialog.FileTitle
                FIRMLOCATION = dlgCommonDialog.FileName
                firmware.Show
                firmware.DownloadFirmware
            End If
        End If
    End If
End Sub

Private Sub MNUDiagnostics_Click()
    diagnostics.Show
    diagnostics.ActivateDiagnostics
End Sub

Private Sub MNUDownLoad_Click()
    On Error Resume Next

    includepath = ActiveForm.rtftext.Tag
    If includepath = "" Then includepath = App.Path & "\bin\"

    If CompileCodeType = NQC Then
        Open "\bin\tempprg.nqc" For Output As #1
        Print #1, ActiveForm.rtftext.text
        Close #1
        closecom
        MainSubs.ExecCmd "nqc.exe" & BrickTypeHeader & " -SCOM" & About.Spirit.ComPortNo & " -I" & Chr(34) & includepath & Chr(34) & " -d -pgm " & Right(ProgramNum.text, 1) & " tempprg.nqc"
        REOPENCOMWHENDONE
    End If

    If CompileCodeType = LASM Then
        Open App.Path & "\bin\tempprg.asm" For Output As #1
        Print #1, ActiveForm.rtftext.text
        Close #1
        closecom
        MainSubs.ExecCmd "lcc32.exe" & BCandLASMBrickTypeHeader & " -S=" & About.Spirit.ComPortNo & " C=LASM -d tempprg.asm"
        LCCREOPENCOMWHENDONE
        ActiveForm.Errors.LoadFile App.Path & "\bin\temp.log"
    End If

    If CompileCodeType = MindScript Then
        Open App.Path & "\bin\tempprg.lsc" For Output As #1
        Print #1, ActiveForm.rtftext.text
        Close #1
        closecom
        MainSubs.ExecCmd "lcc32.exe" & BCandLASMBrickTypeHeader & " -S=" & About.Spirit.ComPortNo & " C=MindScript -d tempprg.lsc"
        LCCREOPENCOMWHENDONE
        ActiveForm.Errors.LoadFile App.Path & "\bin\temp.log"
    End If

    Kill App.Path & "\bin\tempprg.nqc"
    Kill App.Path & "\bin\tempprg.asm"
    Kill App.Path & "\bin\tempprg.lsc"
End Sub

Private Sub MNUDownLoadRun_Click()
    On Error Resume Next

    includepath = ActiveForm.rtftext.Tag
    If includepath = "" Then includepath = App.Path & "\bin\"

    If CompileCodeType = NQC Then
        Open "\bin\tempprg.nqc" For Output As #1
        Print #1, ActiveForm.rtftext.text
        Close #1
        closecom
        MainSubs.ExecCmd "nqc.exe" & BrickTypeHeader & " -SCOM" & About.Spirit.ComPortNo & " -I" & Chr(34) & includepath & Chr(34) & " -d -pgm " & Right(ProgramNum.text, 1) & " tempprg.nqc -run"
        REOPENCOMWHENDONE
    End If

    If CompileCodeType = LASM Then
        Open App.Path & "\bin\tempprg.asm" For Output As #1
        Print #1, ActiveForm.rtftext.text
        Close #1
        closecom
        MainSubs.ExecCmd "lcc32.exe" & BCandLASMBrickTypeHeader & " -S=" & About.Spirit.ComPortNo & " C=LASM -d tempprg.asm"
        LCCREOPENCOMWHENDONE
        ActiveForm.Errors.LoadFile App.Path & "\bin\temp.log"
    End If

    If CompileCodeType = MindScript Then
        Open App.Path & "\bin\tempprg.lsc" For Output As #1
        Print #1, ActiveForm.rtftext.text
        Close #1
        closecom
        MainSubs.ExecCmd "lcc32.exe" & BCandLASMBrickTypeHeader & " -S=" & About.Spirit.ComPortNo & " C=MindScript -d tempprg.lsc"
        LCCREOPENCOMWHENDONE
        ActiveForm.Errors.LoadFile App.Path & "\bin\temp.log"
    End If

    Kill App.Path & "\bin\tempprg.nqc"
    Kill App.Path & "\bin\tempprg.asm"
    Kill App.Path & "\bin\tempprg.lsc"
End Sub

Private Sub mnuEditPaste_Click()
    On Error Resume Next
    '    ActiveForm.rtfText.SelRTF = Clipboard.GetText

End Sub

Private Sub MNUFBrick_Click()
    If COMOPEN = True And COMOPENCLOSED = 1 Then About.Spirit.CloseComm: COMOPEN = False
    SearchRCX.Show
    Me.Hide
End Sub

Private Sub MNUFileMenuSideBar_Click()
About.Show
End Sub

Private Sub MNUFindText_Click()
Static TXTStr As String
    
    If ActiveForm Is Nothing Then
        XPLib.XPMsgBox "No Windows Open.", "Goto Line Number...", False, XP_OKOnly, msg_Exclamation
        Exit Sub
    End If

FindTXT.FindText ActiveForm
End Sub

Private Sub MNUgoto_Click()
    If ActiveForm Is Nothing Then
        XPLib.XPMsgBox "No Windows Open.", "Goto Line Number...", False, XP_OKOnly, msg_Exclamation
        Exit Sub
    End If

    temp = InputBox("Please enter line number.", "Goto Line Number...")
    If temp = "" Then Exit Sub
    If Int(temp) > 1 Then ActiveForm.GoToLine (Int(temp))
End Sub

Private Sub MNUGotoTFS_Click()
    SubsFunctionsTasks.GotoTSF ActiveForm
End Sub

Private Sub MNUlongrange_Click()
    SHORTLONGRANGE = 0
    About.Spirit.PBTxPower (1)
    MNUlongrange.Enabled = False
    MNUshortrange.Enabled = True
End Sub

Private Sub MNUMMap_Click()
    MMap.Show
End Sub

Private Sub MNUMemMap_Click()
    MMap.Show
End Sub

Private Sub MNUmsprgmr_Click()
    MSprogrammer.Show
End Sub

Private Sub MNUOCom_Click()
    COMOPENCLOSED = 1

    sbStatusBar.Panels(1).text = "Communication Open."
    About.Spirit.InitComm
    MNUOCom.Enabled = False
    MNUCCom.Enabled = True
    tbToolBar.Enabled = True

    Me.MNUdf.Enabled = True
    Me.MNUDownLoad.Enabled = True
    Me.MNUDownLoadRun.Enabled = True
    Me.MNURCXWatch.Enabled = True
    Me.MNUTOFF.Enabled = True
End Sub

Private Sub MNUOptions_Click()
    Options.Show
End Sub

Private Sub MNUP1_Click()
    About.Spirit.SelectPrgm 1
    MNUP1.Checked = True
    MNUP2.Checked = False
    MNUP3.Checked = False
    MNUP4.Checked = False
    MNUP5.Checked = False
    ProgramNum.text = "Program 1"
End Sub

Private Sub MNUP2_Click()
    About.Spirit.SelectPrgm 2
    MNUP1.Checked = False
    MNUP2.Checked = True
    MNUP3.Checked = False
    MNUP4.Checked = False
    MNUP5.Checked = False
    ProgramNum.text = "Program 2"
End Sub

Private Sub MNUP3_Click()
    About.Spirit.SelectPrgm 3
    MNUP1.Checked = False
    MNUP2.Checked = False
    MNUP3.Checked = True
    MNUP4.Checked = False
    MNUP5.Checked = False
    ProgramNum.text = "Program 3"
End Sub

Private Sub MNUP4_Click()
    About.Spirit.SelectPrgm 4
    MNUP1.Checked = False
    MNUP2.Checked = False
    MNUP3.Checked = False
    MNUP4.Checked = True
    MNUP5.Checked = False
    ProgramNum.text = "Program 4"
End Sub

Private Sub MNUP5_Click()
    About.Spirit.SelectPrgm 5
    MNUP1.Checked = False
    MNUP2.Checked = False
    MNUP3.Checked = False
    MNUP4.Checked = False
    MNUP5.Checked = True
    ProgramNum.text = "Program 5"
End Sub

Private Sub MNURcxPiano_Click()
    RCXPiano.Show
End Sub

Private Sub MNURCXvar_Click()
    Poll.Show
    Poll.ActivatePoll
End Sub

Private Sub MNURCXWatch_Click()
    SetWatch.Show
End Sub

Private Sub MNUReplaceText_Click()
ReplaceTXT.ReplaceText ActiveForm
End Sub

Private Sub MNUSetPDTime_Click()
    PowerDownTime.Show
End Sub

Private Sub MNUshortrange_Click()
    SHORTLONGRANGE = 1
    About.Spirit.PBTxPower (0)
    MNUlongrange.Enabled = True
    MNUshortrange.Enabled = False
End Sub

Private Sub MNUSTemplates_Click()
    If Templates.Visible = True Then
        MNUSTemplates.Checked = False
        Templates.Visible = False
        CheckTemplatesUnload = False
    Else
        MNUSTemplates.Checked = True
        Templates.Visible = True
    End If
End Sub

Private Sub MNUTOFF_Click()
    About.Spirit.PBTurnOff
    COMOPEN = False
End Sub

Private Sub Port_Click()
    If Port.text <> "AUTO" Then
        MNUCCom_Click
        About.Spirit.ComPortNo = Right(Port.text, 1)
        MNUOCom_Click
        About.Spirit.PBAliveOrNot
    Else
        found = 0
        bestuse = 0
        For I = 1 To 4
            About.Spirit.ComPortNo = I
            Me.sbStatusBar.Panels(1).text = "Checking Port " & I & "..."
            If About.Spirit.InitComm Then
                If About.Spirit.TowerAndCableConnected = True And About.Spirit.PBAliveOrNot = False Then
                    Me.sbStatusBar.Panels(I).text = "Found tower, but not RCX. Searching for RCX on other ports."
                    bestuse = I
                End If
                If About.Spirit.TowerAndCableConnected And About.Spirit.PBAliveOrNot Then
                    Me.sbStatusBar.Panels(1).text = "RCX Found on port " & I & "."
                    found = 1
                    Exit For
                End If
            End If
        Next
        If found = 0 Then Me.sbStatusBar.Panels(1).text = "RCX Not Found."
        If found = 0 And bestuse <> 0 Then
            Me.sbStatusBar.Panels(1).text = "RCX Not Found, but Tower Found. Using COMPORT " & bestuse & " with Valid Tower."
            About.Spirit.ComPortNo = bestuse
            About.Spirit.InitComm
        End If
    End If
End Sub

Private Sub MNUUltiTax_Click()
    If ActiveForm Is Nothing Then
        XPLib.XPMsgBox "No Windows Open.", "Goto Line Number...", False, XP_OKOnly, msg_Exclamation
        Exit Sub
    End If
    
UltiTax.UseUltiTax ActiveForm.rtftext.text
End Sub

Private Sub ProgramNum_Click()
    If COMOPEN = True Then
        If BrickType = RCX Or BrickType = RCX2 Then
            About.Spirit.SelectPrgm (Right(ProgramNum.text, 1) - 1)
            Select Case Int((Right(ProgramNum.text, 1) - 1))
                Case 1
                    MNUP1.Checked = True
                    MNUP2.Checked = False
                    MNUP3.Checked = False
                    MNUP4.Checked = False
                    MNUP5.Checked = False
                Case 2
                    MNUP1.Checked = False
                    MNUP2.Checked = True
                    MNUP3.Checked = False
                    MNUP4.Checked = False
                    MNUP5.Checked = False
                Case 3
                    MNUP1.Checked = False
                    MNUP2.Checked = False
                    MNUP3.Checked = True
                    MNUP4.Checked = False
                    MNUP5.Checked = False
                Case 4
                    MNUP1.Checked = False
                    MNUP2.Checked = False
                    MNUP3.Checked = False
                    MNUP4.Checked = True
                    MNUP5.Checked = False
                Case 5
                    MNUP1.Checked = False
                    MNUP2.Checked = False
                    MNUP3.Checked = False
                    MNUP4.Checked = False
                    MNUP5.Checked = True
            End Select
        End If
    End If
End Sub

Private Sub SmallTB_ButtonClick(ByVal Button As MSComctlLib.Button)
    tbToolBar_ButtonClick Button
End Sub

Private Sub SmallTB_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    tbToolBar_ButtonMenuClick ButtonMenu
End Sub

Private Sub TBCheck_Timer()
    If COMOPEN = False Or NOFIRM = True Then 'Little Mode
        Me.tbToolBar.Visible = False
        Me.SmallTB.Visible = True
        Me.DisableEnableCommands False
        Me.tbToolBar.Enabled = False
        MNUlongrange.Enabled = False
        MNUshortrange.Enabled = False
        MNUOCom.Enabled = False
        MNUCCom.Enabled = False
    Else ' Big Mode
        If SHORTLONGRANGE = 0 Then SRTemp = True: LRTemp = False
        If SHORTLONGRANGE = 1 Then SRTemp = False: LRTemp = True

        MNUlongrange.Enabled = LRTemp
        MNUshortrange.Enabled = SRTemp

        If NOFIRM = True Then
            MNUlongrange.Enabled = False
            MNUshortrange.Enabled = False
        End If

        Me.tbToolBar.Visible = True
        Me.SmallTB.Visible = False
        Me.DisableEnableCommands True
        Me.tbToolBar.Enabled = True

        If COMOPENCLOSED = 0 Then mOCom = True: mCCom = False
        If COMOPENCLOSED = 1 Then mOCom = False: mCCom = True
        If NOFIRM = True Then
            MNUOCom.Enabled = False
            MNUCCom.Enabled = False
        Else
            MNUOCom.Enabled = mOCom
            MNUCCom.Enabled = mCCom
        End If
    End If

    If AFAddText <> "" Then
        AddAFText AFAddText
        AFAddText = ""
    End If

    If FixSTemplatesMenu = True Then
        MNUSTemplates_Click
        FixSTemplatesMenu = False
    End If
    
    If CurrentLine <> 0 Then
        Me.sbStatusBar.Panels(3).text = "Line: " & CurrentLine
        CurrentLine = 0
    End If
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.key
        Case "New"
            LoadNewDoc
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Compile"
            mnuCOMPILEPROGRAM_Click
        Case "Download"
            MNUDownLoad_Click
        Case "DownloadRun"
            MNUDownLoadRun_Click
        Case "Run"
            closecom
            MainSubs.ExecCmd "nqc.exe" & BrickTypeHeader & " -SCOM" & About.Spirit.ComPortNo & " -run"
            REOPENCOMWHENDONE
        Case "Watch"
            MNURCXWatch_Click
        Case "Diagnostics"
            diagnostics.Show
            diagnostics.ActivateDiagnostics
        Case "VRemote"
            VirtualRemote.Show
        Case "ClearMem"
            MNUCLRMEM_Click
        Case "PowerDown"
            MNUSetPDTime_Click
        Case "DirectControl"
            MNUDControl_Click
        Case "MicroscoutProgrammer"
            MNUmsprgmr_Click
        Case "FindBrick"
            If COMOPEN = True Then About.Spirit.CloseComm: COMOPEN = False
            SearchRCX.Show
            Me.Hide
        Case "RCXPiano"
            RCXPiano.Show
        Case "MMap"
            MNUMemMap_Click
        Case "DownloadFirm"
            MNUdf_Click
        Case "RCXVars"
            MNURCXvar_Click
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
    About.Show
End Sub

Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuWindowNewWindow_Click()
    LoadNewDoc
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuEditCopy_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtftext.SelRTF

End Sub

Private Sub mnuEditCut_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtftext.SelRTF
    ActiveForm.rtftext.SelText = vbNullString

End Sub



Private Sub mnuFileExit_Click()
    Unload Me
End Sub


Private Sub mnuFilePrint_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    

    With dlgCommonDialog
        .DialogTitle = "Print"
        .CancelError = True
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
        If ActiveForm.rtftext.SelLength = 0 Then
            .Flags = .Flags + cdlPDAllPages
        Else
            .Flags = .Flags + cdlPDSelection
        End If
        .ShowPrinter
        If Err <> MSComDlg.cdlCancel Then
            ActiveForm.rtftext.SelPrint .hDC
        End If
    End With

End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "Page Setup"
        .CancelError = True
        .ShowPrinter
    End With

End Sub

Private Sub mnuFileSaveAs_Click()
    Dim sFile As String
    

    If ActiveForm Is Nothing Then Exit Sub
    

    With dlgCommonDialog
        .DialogTitle = "Save As"
        .CancelError = False
        .Filter = "Not Quite C Files (*.nqc, *.nqh)|*.nqc;*.nqh|LASM Files (*.asm)|*.asm|Mindscript Files (*.rcx2, *.lsc)|*.rcx2;*.lsc|All Files (*.*)|*.*"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    ActiveForm.Caption = sFile
    ActiveForm.rtftext.SaveFile sFile, &H1

End Sub

Private Sub mnuFileSave_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub

    If ActiveForm.rtftext.text <> vbNullString Then
        Dim sFile As String
        If Left$(ActiveForm.Caption, 7) = "Program" Then
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
            ActiveForm.rtftext.SaveFile sFile, &H1
        Else
            sFile = ActiveForm.Caption
            ActiveForm.rtftext.SaveFile sFile
        End If
    End If
End Sub

Private Sub mnuFileClose_Click()
    If ActiveForm Is Nothing Then
    Else
        Unload ActiveForm
    End If
End Sub

Private Sub mnuFileOpen_Click()
    Dim sFile As String

    With dlgCommonDialog
        .FileName = ""
        .DialogTitle = "Open"
        .CancelError = False
        .Filter = "Not Quite C Files (*.nqc, *.nqh)|*.nqc;*.nqh|LASM Files (*.asm)|*.asm|Mindscript Files (*.rcx2, *.lsc)|*.rcx2;*.lsc|All Files (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    
    LoadNewDoc
            
    ActiveForm.Caption = "(LOADING) " & sFile
    ActiveForm.Icon = ActiveForm.ProgIcons.ListImages(2).Picture
    ActiveForm.rtftext.Backcolor = RGB(200, 200, 200)
    ActiveForm.rtftext.LoadFile sFile
    ActiveForm.Icon = ActiveForm.ProgIcons.ListImages(1).Picture
    ActiveForm.rtftext.Backcolor = vbWhite
    ActiveForm.rtftext.Tag = GetFilePath(dlgCommonDialog.FileName)
        
    CursorTemp = GetSetting("En-Tech URP", "Options", "Cursor", 1)

    If CursorTemp = 1 Then
        ActiveForm.rtftext.SelStart = 0
        ActiveForm.rtftext.SelLength = 0
    Else
        ActiveForm.rtftext.SelStart = Len(ActiveForm.rtftext.text) - 1
        ActiveForm.rtftext.SelLength = 0
    End If
              
    ActiveForm.Caption = sFile
End Sub

Private Sub mnuFileNew_Click()
    LoadNewDoc
End Sub


Function REOPENCOMWHENDONE()
    Me.sbStatusBar.Panels(1).text = "Waiting for NQC.exe to finish..."

    tbToolBar.Enabled = True
    mnuCOMPILE.Enabled = True

    Me.MNUdf.Enabled = True
    Me.MNUDownLoad.Enabled = True
    Me.MNUDownLoadRun.Enabled = True
    Me.MNURCXWatch.Enabled = True
    Me.MNUTOFF.Enabled = True

    MNUOCom_Click
End Function

Function closecom()
    MNUCCom_Click
    tbToolBar.Enabled = False

    Me.MNUdf.Enabled = False
    Me.MNUDownLoad.Enabled = False
    Me.MNUDownLoadRun.Enabled = False
    Me.MNURCXWatch.Enabled = False
    Me.MNUTOFF.Enabled = False
End Function

Function FindWindow(ByVal sClassName As String, ByVal sWindowName As String) As Long
    If Len(sClassName) = 0 Then
        xHwnd = M_FindWindow(0&, sWindowName)
    ElseIf Len(sWindowName) = 0 Then
        xHwnd = M_FindWindow(sClassName, 0&)
    Else
        xHwnd = M_FindWindow(sClassName, sWindowName)
    End If
    FindWindow = xHwnd
End Function

Function DisableEnableCommands(Enabled As Boolean)
    MNUdf.Enabled = Enabled

    If NOFIRM = True Then
        MNUdf.Enabled = True
        Enabled = False
    End If

    MNUDownLoadRun.Enabled = Enabled
    MNUDownLoad.Enabled = Enabled
    MNUSelProgram.Enabled = Enabled
    MNUCLRMEM.Enabled = Enabled
    MNUDControl.Enabled = Enabled
    MNUDiagnostics.Enabled = Enabled
    MNURCXWatch.Enabled = Enabled
    MNUSelProgram.Enabled = Enabled
    MNUSetPDTiEnabled = Enabled
    MNUTOFF.Enabled = Enabled
    MNUSetPDTime.Enabled = Enabled
    MNURCXvar.Enabled = Enabled
    MNUMemMap.Enabled = Enabled
End Function

Sub FlickMode()
    LoadStuff
End Sub

Function WaitForNQC()
    Me.sbStatusBar.Panels(1).text = "Waiting for NQC.exe to finish..."

    While (a > 0) Or (B > 0) Or (X < 1000)
        a = FindWindow("", "NQC")
        B = FindWindow("", "nqc")
        X = X + 1
        DoEvents
    Wend

    Me.sbStatusBar.Panels(1).text = "Compile Done."
End Function

Private Sub tbToolBar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.key
        Case "tMScript"
            tbToolBar.Buttons(12).Image = 26
            SmallTB.Buttons(12).Image = 26
            CompileCodeType = MindScript
        Case "tLASM"
            tbToolBar.Buttons(12).Image = 24
            SmallTB.Buttons(12).Image = 24
            CompileCodeType = LASM
        Case "tNQC"
            tbToolBar.Buttons(12).Image = 25
            SmallTB.Buttons(12).Image = 25
            CompileCodeType = NQC
    End Select
End Sub

Function WaitForLCC()
    Me.sbStatusBar.Panels(1).text = "Waiting for LCC32.exe to finish..."

    While (a > 0) Or (X < 1000)
        a = FindWindow("", "lcc32")
        X = X + 1
        DoEvents
    Wend

    Me.sbStatusBar.Panels(1).text = "Compile Done."
End Function

Function LCCREOPENCOMWHENDONE()
    Me.sbStatusBar.Panels(1).text = "Waiting for LCC32.exe to finish..."

    a = Format(Now, "ss")
    While B - a < 2
        B = Format(Now, "ss")
        DoEvents
    Wend

    While (a > 0) Or (B > 0) Or (X < 1000)
        B = FindWindow("", "lcc32")
        X = X + 1
        DoEvents
    Wend

    tbToolBar.Enabled = True
    mnuCOMPILE.Enabled = True

    Me.MNUdf.Enabled = True
    Me.MNUDownLoad.Enabled = True
    Me.MNUDownLoadRun.Enabled = True
    Me.MNURCXWatch.Enabled = True
    Me.MNUTOFF.Enabled = True

    MNUOCom_Click
End Function

Function GetFilePath(FileNameAndPath As String)
    For I = 1 To Len(FileNameAndPath) Step -1
        a = Mid(FileNameAndPath, I, 1)
        If a = "\" Then
            GetFilePath = Mid(FileNameAndPath, 1, I)
            Exit For
        End If
    Next

End Function

Sub LoadStuff()
    CompileCodeType = NQC

LogText "Brick Type Selection"
    
    Select Case BrickType
        Case CYBERMASTER
            About.Spirit.LinkType = Cable
            About.Spirit.PBrick = Spirit
            sbStatusBar.Panels(8).text = "CABLE"
            LogText "SPIRIT - Initialised (Cybermaster, CABLE)"
        Case RCX
            About.Spirit.LinkType = InfraRed
            About.Spirit.PBrick = RCX
            sbStatusBar.Panels(8).text = "IR"
            LogText "SPIRIT - Initialised (RCX, IR)"
        Case RCX2
            About.Spirit.LinkType = InfraRed
            About.Spirit.PBrick = RCX
            sbStatusBar.Panels(8).text = "IR"
            LogText "SPIRIT - Initialised (RCX2, IR)"
        Case Else
            sbStatusBar.Panels(8).text = "NONE"
            LogText "SPIRIT - No Brick Initialised"
    End Select

    MNUshortrange.Enabled = False
    MNUlongrange.Enabled = False

LogText "Populate Program Combo"

    If COMOPEN = True Then
        sbStatusBar.Panels(1).text = "Communication Open."
        If BrickType = RCX Or BrickType = RCX2 Then
            ProgramNum.AddItem "Program 1"
            ProgramNum.AddItem "Program 2"
            ProgramNum.AddItem "Program 3"
            ProgramNum.AddItem "Program 4"
            ProgramNum.AddItem "Program 5"

            ProgramNum.text = "Program 1"
            If NOFIRM = False Then About.Spirit.SelectPrgm 1
        Else
            ProgramNum.AddItem "<CM Prgm>"
            ProgramNum.text = "<CM Prgm>"
        End If
    Else
        sbStatusBar.Panels(1).text = "Communication Closed."

        If BrickType = CYBERMASTER Then
            ProgramNum.AddItem "<CM Prgm>"
            ProgramNum.text = "<CM Prgm>"
        ElseIf BrickType = RCX Or BrickType = RCX2 Then
            ProgramNum.AddItem "Program 1"
            ProgramNum.AddItem "Program 2"
            ProgramNum.AddItem "Program 3"
            ProgramNum.AddItem "Program 4"
            ProgramNum.AddItem "Program 5"

            ProgramNum.text = "Program 1"
        Else ' No Brick Selected
            ProgramNum.AddItem "<NONE>"
            ProgramNum.text = "<NONE>"
        End If

        Me.tbToolBar.Enabled = False
    End If

LogText "Brick Tune"

    Static PlayTune As Integer

 PlayTune = GetSetting("En-Tech URP", "Startup", "BrickFoundTune", 1)

    If COMOPEN = True And PlayTune = 1 And NOFIRM = False Then
        If BrickType = RCX Or BrickType = RCX2 Then
            About.Spirit.PlayTone 200, (8)
            About.Spirit.PlayTone 250, (8)
            About.Spirit.PlayTone 300, (8)
            About.Spirit.PlayTone 350, (8)
            About.Spirit.PlayTone 300, (8)
            About.Spirit.PlayTone 350, (8)
            About.Spirit.PlayTone 400, (8)
            About.Spirit.PlayTone 450, (8)
            a = Format(Now, "ss")

            While B - a < 1
                B = Format(Now, "ss")
            Wend

            About.Spirit.PlayTone 1000, (20)
            If BrickType = RCX2 Then About.Spirit.PlayTone 1010, (20)
        End If
    End If
    
    If BrickType = RCX Then sbStatusBar.Panels(7).text = "RCX"
    If BrickType = RCX2 Then sbStatusBar.Panels(7).text = "RCX2"
    If BrickType = CYBERMASTER Then sbStatusBar.Panels(7).text = "CM"
    If BrickType = NONE Then sbStatusBar.Panels(7).text = "NONE"

LogText "Main Form Loaded."
End Sub

Function AddAFText(text As String)
    If ActiveForm Is Nothing Then Exit Function

    ActiveForm.rtftext.text = ActiveForm.rtftext.text & vbNewLine & text
End Function
