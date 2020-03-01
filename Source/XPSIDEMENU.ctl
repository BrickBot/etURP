VERSION 5.00
Begin VB.UserControl XPsidemenu 
   Alignable       =   -1  'True
   BackColor       =   &H00C65D21&
   ClientHeight    =   9045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7770
   MousePointer    =   9  'Size W E
   ScaleHeight     =   9045
   ScaleWidth      =   7770
   Begin VB.VScrollBar VScroll1 
      Height          =   855
      Left            =   5685
      TabIndex        =   4
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox piccontainer 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   720
      ScaleHeight     =   6255
      ScaleWidth      =   3255
      TabIndex        =   0
      Top             =   1440
      Width           =   3255
      Begin VB.PictureBox Lsplitter 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   1920
         ScaleHeight     =   1095
         ScaleWidth      =   15
         TabIndex        =   5
         Top             =   2400
         Visible         =   0   'False
         Width           =   15
      End
      Begin XPSideMenus.Header Header 
         Height          =   495
         Index           =   0
         Left            =   225
         TabIndex        =   1
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
      End
      Begin XPSideMenus.Hholder Hholder 
         Height          =   30
         Index           =   0
         Left            =   480
         TabIndex        =   2
         Top             =   1080
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   53
         Begin VB.Label Hhyperlink 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caption"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C45518&
            Height          =   210
            Index           =   0
            Left            =   240
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Tag             =   "0"
            Top             =   1560
            Width           =   540
         End
         Begin VB.Image HHimage 
            Height          =   240
            Index           =   0
            Left            =   360
            Top             =   840
            Width           =   240
         End
         Begin VB.Image Himage 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   1100
            Index           =   0
            Left            =   1080
            MousePointer    =   99  'Custom
            Stretch         =   -1  'True
            Top             =   720
            Width           =   950
         End
      End
      Begin VB.Shape Border 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   6  'Inside Solid
         Height          =   375
         Left            =   1920
         Top             =   1320
         Width           =   375
      End
   End
End
Attribute VB_Name = "XPsidemenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Event HyperClick(key As String)
Event PictureClick(key As String)
Public Himagelist                       As ImageList
Public Pimagelist                       As ImageList
Private Hcol                            As Collection
Private Pcol                            As Collection
Private mg                              As clsGradient
Private P_showone                       As Boolean

Private iFullFormHeigth As Integer
Private iFullFormWidth As Integer
Private oldvPos As Integer
Private oldhPos As Integer

Private ANISPEED As Integer
Private DORESIZE As Boolean
Private XP_resizable As Boolean

Public Property Get Resizable() As Boolean
    Resizable = XP_resizable
End Property

Public Property Let Resizable(ByVal vNewValue As Boolean)
    XP_resizable = vNewValue
    PropertyChanged "Resizable"
    UserControl_Resize
End Property

Public Function AddHyper(key As String, ParentKey As String, Caption As String, Enabled As Boolean, H_style As H_type, Optional Icon As Integer, Optional Tooltip As String)
On Error Resume Next
Dim hObj                                As XPHypers
    Dim keyindex                        As Integer
    Dim Numhyper                        As Integer
    Dim Numlabel                        As Integer
    Dim pobj                            As XPPanels
Set hObj = New XPHypers
    hObj.key = key
    hObj.Caption = Caption
    hObj.Enabled = Enabled
    hObj.hstyle = H_style
    hObj.Icon = Icon
    hObj.ParentKey = ParentKey
    hObj.Tooltip = Tooltip
    'Add to Collection
    Hcol.Add hObj
    keyindex = getindexfromkey(ParentKey)

    

        Load Hhyperlink(Hhyperlink.Count)
        
            Numhyper = Hhyperlink.Count - 1
                Set Hhyperlink(Numhyper).Container = Hholder(keyindex)
            Hhyperlink(Numhyper).Visible = hObj.Enabled
            Hhyperlink(Numhyper).Caption = hObj.Caption
            Hhyperlink(Numhyper).Enabled = hObj.Enabled
            Hhyperlink(Numhyper).Tag = key
            Hhyperlink(Numhyper).ToolTipText = hObj.Tooltip

           If hObj.hstyle = 0 Then
                Load HHimage(HHimage.Count)
                Set HHimage(Numhyper).Container = Hholder(keyindex)
                HHimage(Numhyper).Visible = hObj.Enabled
                HHimage(Numhyper).Picture = Himagelist.ListImages.Item(hObj.Icon).Picture
           ElseIf hObj.hstyle = 1 Then
                Hhyperlink(Numhyper).ForeColor = vbBlack
                Hhyperlink(Numhyper).MousePointer = 0
           End If
            
      
         
      
   

    Set pobj = Pcol(keyindex)
            If pobj.Psate <> Closed Then
           Hholder(keyindex).Resizeme pobj.HasPicture, ANISPEED
           End If
           

    'Clear From Memory
    Set hObj = Nothing
    Set pobj = Nothing
    'Add to PropertyBag


End Function

Public Function Addpanel(key As String, Caption As String, Pstate As P_state, HasPicture As Boolean, Optional Icon As Integer, Optional ppicture As Variant, Optional Pic_path As String)

    Dim keyindex                        As Integer
Dim pobj                                As XPPanels
On Error Resume Next
Set pobj = New XPPanels
    pobj.key = key
    pobj.Caption = Caption
    pobj.Icon = Icon
    pobj.Psate = Pstate

    
Set pobj.ppicture = ppicture
    pobj.PicturePath = Pic_path
    pobj.HasPicture = HasPicture
    'Add to Collection
    Pcol.Add pobj
    'Add to PropertyBag
    PropertyChanged "Addpanel"

        keyindex = getindexfromkey(key)
        X = Pcol.Count
            Set pobj = Pcol.Item(X)
        Load Header(X)
        Load Hholder(X)
        Header(X).numpanel = X
        Header(X).Move 200, 0 - Header(X).Height, piccontainer.ScaleWidth - (Header(X).Left * 2)

        Header(X).Headerstyle = pobj.Psate
        Header(X).Top = Header(X - 1).Top + Header(X - 1).Height + 60 + Hholder(X - 1).Height
        Hholder(X).Top = Header(X).Top + Header(X).Height
        Hholder(X).Left = Header(X).Left
        Hholder(X).Width = Header(X).Width
        Hholder(X).Height = 0
        Header(X).Visible = True
        If pobj.Psate = Closed Then Hholder(X).Visible = False Else Hholder(X).Visible = True
        Header(X).SethPic Pimagelist.ListImages.Item(pobj.Icon).Picture
        
        Header(X).Caption = pobj.Caption
        If pobj.Icon = 0 Then
            Header(X).MOVECAPTION True
        Else
            Header(X).MOVECAPTION False
        End If
    If pobj.HasPicture = True Then

        Load Himage(Himage.Count)
        Set Himage(Himage.Count - 1).Container = Hholder(keyindex)
        Himage(Himage.Count - 1).Top = 100
        Himage(Himage.Count - 1).Visible = True
        Himage(Himage.Count - 1).Tag = pobj.key
        If pobj.PicturePath = "" Then
            Set Himage(Himage.Count - 1).Picture = pobj.ppicture
        Else
            Himage(Himage.Count - 1).Picture = LoadPicture(pobj.PicturePath)
        End If
    End If

    'Clear from Memory
    Set pobj = Nothing

End Function

Public Function ChangePANEL(key As String, Optional newKey As String, Optional Caption As String, Optional Pstate As P_state, Optional Icon As Integer)

Dim pobj                                As XPPanels
On Error Resume Next
X = getindexfromkey(key)
Set pobj = Pcol.Item(X)

pobj.Caption = Caption
pobj.Icon = Icon
pobj.Psate = Pstate

If Len(newKey) > 0 Then

    pobj.key = newKey

End If

    Header(X).numpanel = X
    Header(X).Headerstyle = pobj.Psate
    Header(X).SethPic Pimagelist.ListImages.Item(pobj.Icon).Picture
    Header(X).Caption = pobj.Caption

End Function

Public Function ChangeHyper(key As String, Optional Caption As String, Optional Enabled As Boolean, Optional H_style As H_type, Optional Icon As Integer, Optional Tooltip As String)

Dim hObj                                As XPHypers
On Error Resume Next
X = gethindexfromkey(key)
Set hObj = Hcol.Item(X)

hObj.Caption = Caption
hObj.Enabled = Enabled
hObj.hstyle = H_style
    hObj.Icon = Icon
hObj.Tooltip = Tooltip

Dim test As Boolean




            Hhyperlink(X).Visible = hObj.Enabled
            HHimage(X).Visible = hObj.Enabled
            HHimage(X).Picture = Himagelist.ListImages.Item(hObj.Icon).Picture
            Hhyperlink(X).Caption = hObj.Caption
            Hhyperlink(X).Enabled = hObj.Enabled
            Hhyperlink(X).Tag = key
            Hhyperlink(X).ToolTipText = hObj.Tooltip


            If hObj.hstyle = 0 Then
                Set HHimage(X).Container = Hholder(keyindex)
                HHimage(X).Visible = hObj.Enabled
                HHimage(X).Picture = Himagelist.ListImages.Item(hObj.Icon).Picture
           ElseIf hObj.hstyle = 1 Then
                Hhyperlink(X).ForeColor = vbBlack
                Hhyperlink(X).MousePointer = 0
                HHimage(X).Visible = False
           End If

   
keyindex = getindexfromkey(hObj.ParentKey)
    Set pobj = Pcol(keyindex)

      
            If pobj.Psate <> Closed Then
           Hholder(keyindex).Resizeme pobj.HasPicture, ANISPEED
           End If
     

End Function

Private Function getindexfromkey(key As String) As Integer

Dim pobj                                As XPPanels

For I = 1 To Pcol.Count
    Set pobj = Pcol(I)

    If pobj.key = key Then

        getindexfromkey = I

    End If

Next I

End Function

Private Function gethindexfromkey(key As String) As Integer

Dim hObj                                As XPHypers

For I = 1 To Hcol.Count
    Set hObj = Hcol(I)

    If hObj.key = key Then

        gethindexfromkey = I

    End If

Next I

End Function

Public Function RemoveAll()
On Error Resume Next
For I = 1 To Hcol.Count
    Hcol.Remove (I)
    Unload HHimage(I)
    Unload Hhyperlink(I)
    
Next I
For z = 1 To Pcol.Count
    Pcol.Remove (z)
    Unload Himage(z)
    Unload Hholder(z)
    Unload Header(z)
Next z
Set Pcol = Nothing
Set Hcol = Nothing
    Set Pcol = New Collection
    Set Hcol = New Collection

End Function

Public Property Get Showone() As Boolean
Attribute Showone.VB_Description = "When true All panels are closed Except the selected Panels and those that are set to Fixed"
    Showone = P_showone
End Property

Public Property Let Showone(newvalue As Boolean)
    P_showone = newvalue
    PropertyChanged "Showone"
End Property


Private Sub Header_click(Index As Integer)
    On Error Resume Next
    Dim pobj                                As XPPanels
    Dim oldheight As Long
    Set pobj = Pcol(Index)

    
        If pobj.Psate = Opened Then
            pobj.Psate = Closed
             spdfact = 1
             I = Hholder(Index).Height
             Do
                I = I - 1 * spdfact
                Hholder(Index).Height = I
                spdfact = spdfact + 5
                    'Exit Resize Loop if the min Height is reached
                    If oldheight = Hholder(Index).Height Then Exit Do
                    oldheight = Hholder(Index).Height
             Loop Until Hholder(Index).Height <= 1
             Hholder(Index).Visible = False
     
        ElseIf pobj.Psate = Closed Then
            Hholder(Index).Visible = True
            pobj.Psate = Opened
            Hholder(Index).Resizeme pobj.HasPicture, ANISPEED
    
        End If
    
    Header(Index).Headerstyle = pobj.Psate

If pobj.Psate <> Fixed Then
    For I = 1 To Pcol.Count
        Set pobj = Pcol(I)
            If P_showone = True Then
                If I <> Index Then
                        If pobj.Psate = Opened Then
                            pobj.Psate = Closed
                            Header(I).Headerstyle = pobj.Psate
                            For E = 1 To Hholder(I).Height
                                Hholder(I).Height = Hholder(I).Height - ANISPEED
                            Next E
                        End If
                End If
            End If
    Next I
End If
Set pobj = Nothing

End Sub

Private Sub Hholder_Mousemove(Index As Integer)
For I = 1 To Hhyperlink.Count - 1
    Hhyperlink(I).FontUnderline = False

Next I
End Sub

Private Sub Hholder_resize(Index As Integer)

On Error Resume Next
For I = 0 To Header.Count - 1

    If I >= Index Then

    Header(I + 1).Top = Hholder(I).Top + Hholder(I).Height + 60
    Hholder(I + 1).Top = Header(I + 1).Top + Header(I + 1).Height

    End If

Next I
 For I = 1 To Himage.Count - 1
    Himage(I).Left = (Hholder(Index).Width / 2) - (Himage(I).Width / 2)
    Next I
    DoEvents
 GetFullSize

    UserControl_Resize
    
             For I = 1 To HHimage.Count - 1
                HHimage(I).Top = Hhyperlink(I).Top
                HHimage(I).Left = 50
            Next I
   
End Sub





Private Sub GetFullSize()
Dim ctl As Control
Dim fullhtemp As Integer
Dim fullvtemp As Integer

fullhtemp = 0
fullvtemp = 0
For I = 1 To Hholder.Count - 1
        If Hholder(I).Top + Hholder(I).Height > fullvtemp Then fullvtemp = Hholder(I).Top + Hholder(I).Height
Next
iFullFormHeigth = fullvtemp
End Sub
Private Sub pScrollForm()


For I = 1 To Header.Count - 1
Header(I).Top = Header(I).Top + oldvPos - VScroll1.Value
Hholder(I).Top = Hholder(I).Top + oldvPos - VScroll1.Value

Next I


oldvPos = VScroll1.Value

End Sub
Private Sub Hhyperlink_Click(Index As Integer)
If Hhyperlink(Index).MousePointer = 99 Then
    RaiseEvent HyperClick(Hhyperlink(Index).Tag)
End If
End Sub

Public Property Get Showborder() As Boolean

Showborder = Border.Visible

End Property

Public Property Let Showborder(newvalue As Boolean)

Border.Visible = newvalue
PropertyChanged "Showborder"

End Property

Private Sub Hhyperlink_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Hhyperlink(Index).ForeColor <> vbBlack Then
    Hhyperlink(Index).FontUnderline = True
End If
End Sub

Private Sub Himage_Click(Index As Integer)
RaiseEvent PictureClick(Himage(Index).Tag)

End Sub



Private Sub piccontainer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl.MousePointer = 0

End Sub

Private Sub piccontainer_Resize()
On Error Resume Next

With piccontainer
    Border.Move .ScaleLeft, .ScaleTop, .ScaleWidth, .ScaleHeight
    Header(0).Move 200, 0 - Header(0).Height, .ScaleWidth - (Header(0).Left * 2)

End With
For I = 1 To Header.Count - 1
Header(I).Width = piccontainer.ScaleWidth - (Header(I).Left * 2)
Hholder(I).Width = Header(I).Width

Next I



With mg
.Color1 = &HE6A17A
.Color2 = &HD67764
.Angle = 270
.Draw piccontainer

End With
piccontainer.Refresh

End Sub


Private Sub UserControl_Initialize()

    Set Pcol = New Collection
    Set Hcol = New Collection
    Set mg = New clsGradient
GetFullSize

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'let user resize control at runetime
DORESIZE = True

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'let user resize control at runetime
UserControl.MousePointer = 9
If DORESIZE = True Then
    If X > UserControl.Width Then
        UserControl.Parent.Refresh
        UserControl.Parent.Line (X, 0)-(X, UserControl.Parent.Height), &H404040, B
    ElseIf X <= UserControl.Width Then
        UserControl.Parent.Refresh
        piccontainer.Refresh
        With Lsplitter
            .Left = X
            '.X2 = X
            .Top = 0
            .Height = UserControl.Height
            .Visible = True
        End With
    End If
End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'let user resize control at runtime
If DORESIZE = True Then
    UserControl.Width = 0 + X
    UserControl.Parent.Refresh
    Lsplitter.Visible = False
    DORESIZE = False
End If

End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

Border.Visible = PropBag.ReadProperty("ShowBorder", True)
P_showone = PropBag.ReadProperty("Showone", False)
ANISPEED = PropBag.ReadProperty("Speed", 35)
XP_resizable = PropBag.ReadProperty("Resizable", False)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
Dim Resizegrip As Single
VScroll1.Visible = (iFullFormHeigth - UserControl.Height) >= 0
With UserControl
    If XP_resizable = True Then
        Resizegrip = 25
    ElseIf XP_resizable = False Then
        Resizegrip = 0
    End If
End With


With UserControl
If (iFullFormHeigth - UserControl.Height) >= 0 Then
   piccontainer.Move .ScaleLeft, .ScaleTop, .ScaleWidth - VScroll1.Width - Resizegrip, .ScaleHeight
Else
    piccontainer.Move .ScaleLeft, .ScaleTop, .ScaleWidth - Resizegrip, .ScaleHeight
End If


End With
VScroll1.Move UserControl.ScaleWidth - VScroll1.Width - Resizegrip, 0, VScroll1.Width, UserControl.ScaleHeight


 If VScroll1.Enabled Then
        With VScroll1
            .Min = 0
            .Max = (iFullFormHeigth - UserControl.Height) + 100
            .SmallChange = Screen.TwipsPerPixelY * 10
            .LargeChange = UserControl.ScaleHeight
        End With
End If

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

PropBag.WriteProperty "ShowBorder", Border.Visible, True
PropBag.WriteProperty "Showone", P_showone, False
PropBag.WriteProperty "Speed", ANISPEED, 35
PropBag.WriteProperty "Resizable", XP_resizable, False
End Sub




Private Sub VScroll1_Change()
 Call pScrollForm
End Sub


Private Sub VScroll1_Scroll()
Call pScrollForm
End Sub



Public Property Get Speed() As Integer
    Speed = ANISPEED
End Property

Public Property Let Speed(ByVal vNewValue As Integer)
    ANISPEED = vNewValue
    PropertyChanged "Speed"
End Property
