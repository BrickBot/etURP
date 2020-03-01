VERSION 5.00
Begin VB.UserControl ThreeDLineLR 
   BackStyle       =   0  'Transparent
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1800
   ScaleHeight     =   255
   ScaleWidth      =   1800
   ToolboxBitmap   =   "ThreeD_LR.ctx":0000
   Begin VB.Line LineB 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   6  'Inside Solid
      X1              =   0
      X2              =   1800
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line LineA 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   1800
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "ThreeDLineLR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Simple Left-To-Right 3D Line Control, (C) Dean Camera, 2003


Private Sub UserControl_Resize()
    UserControl.Height = 90
    LineA.X2 = UserControl.Width
    LineB.Y1 = 15
    LineB.Y2 = 15
    LineB.X2 = UserControl.Width
End Sub
