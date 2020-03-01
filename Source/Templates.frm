VERSION 5.00
Begin VB.Form Templates 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Templates"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   Begin etUCP.ThreeDLineLR ThreeDLineLR1 
      Height          =   90
      Left            =   0
      TabIndex        =   20
      Top             =   1440
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   159
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   22
      Left            =   1440
      TabIndex        =   25
      Top             =   2040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "StopAllTasks();"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   21
      Left            =   1440
      TabIndex        =   24
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "stop ..;"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   20
      Left            =   1440
      TabIndex        =   23
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "start ..;"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   31
      Left            =   1440
      TabIndex        =   36
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "ClearTimer(..);"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   28
      Left            =   1440
      TabIndex        =   32
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "ClearMessage();"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":0070
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   29
      Left            =   1440
      TabIndex        =   33
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "SendMessage(..);"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":008C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   26
      Left            =   0
      TabIndex        =   30
      Top             =   4920
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "ClearSensor(..);"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":00A8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   32
      Left            =   0
      TabIndex        =   38
      Top             =   5640
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "PlayTone(..,..);"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":00C4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.ThreeDLineLR ThreeDLineLR5 
      Height          =   90
      Left            =   0
      TabIndex        =   37
      Top             =   5280
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   159
   End
   Begin etUCP.ThreeDLineLR ThreeDLineLR3 
      Height          =   90
      Left            =   0
      TabIndex        =   27
      Top             =   4080
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   159
   End
   Begin etUCP.ThreeDLineUD ThreeDLineUD2 
      Height          =   4095
      Left            =   1320
      TabIndex        =   21
      Top             =   0
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   7223
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   19
      Left            =   120
      TabIndex        =   19
      Top             =   3480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "until (..){..};"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":00E0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   18
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "until (..);"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":00FC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "repeat (..){..}"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":0118
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "do{..}while(..)"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":0134
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.ThreeDLineLR ThreeDLineLR2 
      Height          =   90
      Left            =   0
      TabIndex        =   22
      Top             =   2400
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   159
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "If (..){..}else{..}"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":0150
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "If (..){..}"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":016C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "If (..) ..;"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":0188
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   11
      Left            =   1440
      TabIndex        =   11
      Top             =   3720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "Wait(..);"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":01A4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   10
      Left            =   1440
      TabIndex        =   10
      Top             =   3480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "SetPower(..,..);"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":01C0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   9
      Left            =   1440
      TabIndex        =   9
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "Float(..);"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":01DC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   8
      Left            =   1440
      TabIndex        =   8
      Top             =   3000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "Off(..);"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":01F8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   7
      Left            =   1440
      TabIndex        =   7
      Top             =   2760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "OnRev(..);"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":0214
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "#include .."
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":0230
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "#define .."
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":024C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "int ..;"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":0268
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "void ..(..){..}"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":0284
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   6
      Left            =   1440
      TabIndex        =   6
      Top             =   2520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "OnFwd(..);"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":02A0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "while (..){..}"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":02BC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "task ..(){..}"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":02D8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "sub ..(){..}"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":02F4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   23
      Left            =   0
      TabIndex        =   26
      Top             =   4680
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "SetSensor(..,SENSOR_ROTATION);"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":0310
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   24
      Left            =   0
      TabIndex        =   28
      Top             =   4440
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "SetSensor(..,SENSOR_LIGHT);"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":032C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   25
      Left            =   0
      TabIndex        =   29
      Top             =   4200
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "SetSensor(..,SENSOR_TOUCH);"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":0348
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   27
      Left            =   1440
      TabIndex        =   31
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "Message()"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":0364
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   30
      Left            =   1440
      TabIndex        =   35
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "Timer(..)"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":0380
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.chameleonButton Template 
      Height          =   255
      Index           =   33
      Left            =   0
      TabIndex        =   39
      Top             =   5400
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      BTYPE           =   13
      TX              =   "PlaySound(..);"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Templates.frx":039C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin etUCP.ThreeDLineLR ThreeDLineLR4 
      Height          =   90
      Left            =   1320
      TabIndex        =   34
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   159
   End
End
Attribute VB_Name = "Templates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
LogText "Load - TEMPLATES"
    
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
    Me.Left = Screen.Width - Me.Width - 10
    Me.Top = Screen.Height - Me.Height - 1110
End Sub

Private Sub Form_Terminate()
    If CheckTemplatesUnload = True Then
        FixSTemplatesMenu = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FixSTemplatesMenu = True
End Sub

Private Sub Template_Click(Index As Integer)
    Select Case Index
        Case 0
            AFAddText = "task ""name""()" & vbNewLine & "{" & vbNewLine & "  ""statements""" & vbNewLine & "}"
        Case 1
            AFAddText = "sub ""name""()" & vbNewLine & "{" & vbNewLine & "  ""statements""" & vbNewLine & "}"
        Case 2
            AFAddText = "void ""name""(""Arguments"")" & vbNewLine & "{" & vbNewLine & "  ""statements""" & vbNewLine & "}"
        Case 3
            AFAddText = "int ""name"";"
        Case 4
            AFAddText = "#define ""macro"""
        Case 5
            AFAddText = "#include ""file"""
        Case 6
            AFAddText = "OnFwd(""outs"");"
        Case 7
            AFAddText = "OnRev(""outs"");"
        Case 8
            AFAddText = "Off(""outs"");"
        Case 9
            AFAddText = "Float(""outs"");"
        Case 10
            AFAddText = "SetPower(""outs"",""speed"");"
        Case 11
            AFAddText = "Wait(""ticks"");"
        Case 12
            AFAddText = "if (""condition"") ""statement"";"
        Case 13
            AFAddText = "if (""condition"")" & vbNewLine & "{" & vbNewLine & "  ""statements""" & vbNewLine & "}"
        Case 14
            AFAddText = "if (""condition"")" & vbNewLine & "{" & vbNewLine & "  ""statements""" & vbNewLine & "}" & vbNewLine & "else" & vbNewLine & "{" & vbNewLine & "  ""statements""" & vbNewLine & "}"
        Case 15
            AFAddText = "while (""condition"")" & vbNewLine & "{" & vbNewLine & "  ""statements""" & vbNewLine & "}"
        Case 16
            AFAddText = "do" & vbNewLine & "{" & vbNewLine & "  ""statements""" & vbNewLine & "}" & vbNewLine & "while (""condition"")"
        Case 17
            AFAddText = "repeat (""value"")" & vbNewLine & "{" & vbNewLine & "  ""statements""" & vbNewLine & "}"
        Case 18
            AFAddText = "until (""condition"");"
        Case 19
            AFAddText = "until (""condition"")" & vbNewLine & "{" & vbNewLine & "  ""statements""" & vbNewLine & "}"
        Case 20
            AFAddText = "start ""taskname"";"
        Case 21
            AFAddText = "stop ""taskname"";"
        Case 22
            AFAddText = "StopAllTasks();"
        Case 23
            AFAddText = "SetSensor(""sensor"",SENSOR_ROTATION);"
        Case 24
            AFAddText = "SetSensor(""sensor"",SENSOR_LIGHT);"
        Case 25
            AFAddText = "SetSensor(""sensor"",SENSOR_TOUCH);"
        Case 26
            AFAddText = "ClearSensor(""sensor"");"
        Case 27
            AFAddText = "Message();"
        Case 28
            AFAddText = "ClearMessage();"
        Case 29
            AFAddText = "SendMessage(""value"");"
        Case 30
            AFAddText = "Timer(""number"")"
        Case 31
            AFAddText = "ClearTimer(""number"");"
        Case 32
            AFAddText = "PlayTone(""freq"",""ticks"");"
        Case 33
            AFAddText = "PlaySound(""numb"");"
    End Select
End Sub


