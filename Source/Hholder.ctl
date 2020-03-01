VERSION 5.00
Begin VB.UserControl Hholder 
   BackColor       =   &H00F7DFD6&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "Hholder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event resize()

Public Event Mousemove()
Public Function Resizeme(Hasimage As Boolean, dropspeed As Integer)

Dim ctl                                 As Control
Dim ctlcount                            As Integer
Dim ctlicount                           As Integer
Dim oldheight As Long
    ' Examine every control.

    For Each ctl In UserControl.ContainedControls

        If (TypeOf ctl Is Label) Then
            
            If ctl.Enabled = True Then
                
                ctlcount = ctlcount + 1
            
            End If
                
                If ctlcount = 1 Then

                    If Hasimage = True Then

                        ctl.Top = 1400

                    Else

                        ctl.Top = 100

                    End If

                Else

                    If Hasimage = True Then

                        ctl.Top = ((ctl.Height) * (ctlcount - 1)) + (100 * ctlcount) + 1300

                    Else

                        ctl.Top = ((ctl.Height) * (ctlcount - 1)) + (100 * ctlcount)

                    End If

                End If

        ctl.Left = 380
        Dim myheight As Single

                        If Hasimage = True Then

                            myheight = (ctl.Height * ctlcount) + (100 * (ctlcount + 3)) + 1300

                        Else

                            myheight = (ctl.Height * ctlcount) + (100 * (ctlcount + 2))

                        End If

    End If

    Next ctl

If dropspeed = 0 Then
UserControl.Height = myheight
Else
oldheight = 1
 spdfact = 0
 I = 0
 Do
    I = I + 1
    UserControl.Height = I * spdfact
    spdfact = spdfact + 5
 Loop While UserControl.Height <= myheight
End If

End Function

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent Mousemove
End Sub

Private Sub UserControl_Resize()
    RaiseEvent resize
End Sub

