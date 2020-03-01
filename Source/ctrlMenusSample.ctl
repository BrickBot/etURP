VERSION 5.00
Begin VB.UserControl ctrlMenusSample 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1605
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   765
   ScaleWidth      =   1605
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1680
      Picture         =   "ctrlMenusSample.ctx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Use the HelpHook && SampleMenus_IDE control properties"
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   1500
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "ctrlMenusSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' This control will allow you to view your menus while in IDE mode with some exceptions.
' Read the notes below...

' Exceptions: If the form where this control is attached is in IDE, controls cannot be fully referenced since
' no instance of the form is created while it is in IDE. Therefore...
' 1. Menu icons will not be displayed. However, a default image will be displayed in its place.
' 2. Image sidebars will not be displayed for same reasons. A default sidebar image will be displayed in its place
' 3. Only visible menus will be displayed while in IDE. To see how non-visible menus will look, change their
'       visible property to True
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' This control should NOT be included in your compiled application. It has no purpose once the form is run
' outside of IDE.  Additionally, there are a few items that should not be called/set from your application and
' are designed only to work with/for this user control.  Those routines are...
' 1. clsMenuItems.RestoreMenus
' 2. modMenus.AmInIDE (variable) -- never set to true. The control will set/reset as needed.
' 3. modMenus.DefaultIcon (variable). Only applies if AmInIDE is set to True
' One final note. Changing the module name from modMenus to another name will require reflecting the
' name change in this control only. Changing the modMenus routine SetMenu would also
' require a change in this control
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

Private TestFont As StdFont
Private bHelpHook As Boolean
Private hParent As Long                     ' handle of form that control is attached to
Private bIDEhook As Boolean             ' status of subclassing
Private bInfoMsgShown As Boolean    ' indicates general info message has been displayed or not
Public Enum HelpTopics
    General = 0
    Alignments = 1
    Colors = 2
    FontAttributes = 3
    Images = 4
    NoScroll = 5
    ListComboBoxes = 6
    SeparatorBars = 7
    SidebarText = 8
    Tips = 9
    Colors_Menu = 10
    DayOfMonth_Menu = 11
    Drives_Menu = 12
    Fonts_Menu = 13
    Months_Menu = 14
    States_Menu = 15
    WeekDays_Menu = 16
End Enum
    
Private Sub Image1_Click()

End Sub

Private Sub UserControl_Initialize()
Set TestFont = New StdFont
TestFont.Name = modMenus.MenuFontName
TestFont.Size = modMenus.MenuFontSize
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
' ensure control doesn't hook form when form is run
If Ambient.UserMode Then HookParent False
End Sub

Private Sub UserControl_Show()
' ensure control doesn't hook form when form is run
If Ambient.UserMode = False Then HookParent False
End Sub

Private Sub UserControl_Terminate()
' ensure hook is removed when form closes
If hParent Then HookParent False
Set TestFont = Nothing
End Sub

Private Sub HookParent(bSet As Boolean)
' either hook/unhook form
bIDEhook = bSet
If bSet Then                                    ' hooking form
    If Not bInfoMsgShown Then          ' display this message if not displayed once
        MsgBox "1. Images will not display while in IDE (default image instead)" & vbCrLf & _
        "2. Image sidebars will show LaVolpe fox head while in IDE" & vbCrLf & _
        "3. Non-Visible menus will not be displayed while in IDE" & vbCrLf & _
        "4. Combo & List box items won't be displayed on submenus" & vbCrLf & vbCrLf & _
        "Note: DO NOT leave user control in your application -- remove it when not needed." & vbCrLf & vbCrLf & _
        "All of the above will work properly once application is run.", vbInformation + vbOKOnly, "Notes while in IDE"
        bInfoMsgShown = True            ' prevent displaying message again
    End If
    modMenus.AmInIDE = True         ' set flag to use default icon & restore menus when unhooked
    modMenus.DefaultIcon = UserControl.Image1.Picture.handle    ' set default picture for icons
    Debug.Print "hooking form via the user control"
    hParent = Parent.hWnd               ' cache parent form hwnd
    SetMenu hParent                        ' hook form per request, but don't provide ImageList control or tips option
Else
    Debug.Print "Usercontrol unhooking form if necessary"
    If hParent Then CleanClass hParent  ' unhook form & restore menu items
    hParent = 0                                    ' reset variables
    modMenus.AmInIDE = False
End If
End Sub

Public Property Let SampleMenus_IDE(bSample As Boolean)
Attribute SampleMenus_IDE.VB_Description = "Toggle viewing menus in design mode."
' hook/unhook form per user request
HookParent bSample
End Property
Public Property Get SampleMenus_IDE() As Boolean
' return status of subclassing/hooking
SampleMenus_IDE = bIDEhook
End Property

Public Property Get HelpHook() As Boolean
Attribute HelpHook.VB_Description = "General help information about showing menus in design mode"
' just provided so user can select True from the HelpHook property
HelpHook = bHelpHook
End Property

Public Property Let HiLiteDisabledItems(bHiLite As Boolean)
Attribute HiLiteDisabledItems.VB_Description = "Toggle highlighting or not highlighting disabled menu items. Keyboard navigation always highlights disabled items."
    modMenus.HighlightDisabledMenuItems = bHiLite
End Property
Public Property Get HiLiteDisabledItems() As Boolean
    HiLiteDisabledItems = modMenus.HighlightDisabledMenuItems
End Property
Public Property Let HiliteGradient(bGradient As Boolean)
Attribute HiliteGradient.VB_Description = "Toggle between solid or gradient background highlighting of menu items"
    modMenus.HighlightGradient = bGradient
End Property
Public Property Get HiliteGradient() As Boolean
    HiliteGradient = modMenus.HighlightGradient
End Property
Public Property Let HiliteItalicized(bItalics As Boolean)
Attribute HiliteItalicized.VB_Description = "Toggle between italicizing menu items when they become highlighted. Disabled items are never italicized"
    modMenus.ItalicizeSelectedItems = bItalics
End Property
Public Property Get HiliteItalicized() As Boolean
    HiliteItalicized = modMenus.ItalicizeSelectedItems
End Property
Public Property Set SetTestFont(ByRef xFont As Font)
    Set TestFont = xFont
    modMenus.MenuFontName = TestFont.Name
    modMenus.MenuFontSize = TestFont.Size
End Property
Public Property Get SetTestFont() As Font
    Set SetTestFont = TestFont
End Property

Public Property Let HelpHook(bHelp As Boolean)
If bHelp = True Then
    ' user wants a little help, provide help and simultaneously hook/unhook form as requested
    Dim iResponse As Integer
    iResponse = MsgBox("Turn sampling on off via the SampleMenus_IDE property." & vbCrLf & vbCrLf & _
    "DO NOT close VB while sampling!! You can close a form window without worry." & vbCrLf & vbCrLf & _
    "Also, don't compile with this control attached." & vbCrLf & vbCrLf & "Start/Stop Sampling?", vbYesNoCancel, "Menu Previewer")
    If iResponse = vbYes Then
        ' user wants to toggle the hooking action; otherwise no hooking action changes
        If bIDEhook Then HookParent False Else HookParent True
    End If
End If
bHelpHook = False       ' show the HelpHook property as False again
End Property

Public Property Get HelpFormatting() As HelpTopics
Attribute HelpFormatting.VB_Description = "General help pertaining to formatting of menu captions"
HelpFormatting = General
End Property
Public Property Let HelpFormatting(Topic As HelpTopics)
Dim sMsg As String, sTitle As String
Select Case Topic
Case General
    sTitle = "General Formatting"
    sMsg = "The coded part of the caption can come before or after the actual menu caption." & vbCrLf & _
        "Coded part must be enclosed in { } & each flag separated by a pipe symbol ( | )." & vbCrLf & vbCrLf & _
        "If a flag requires a value, you format as FLAG:Value  ..." & vbCrLf & vbCrLf
    sMsg = sMsg & "Align: used in text/image sidebars only" & vbCrLf & _
        "Text: used in text sidebars" & vbCrLf & _
        "Files: used with LB:" & vbCrLf & _
        "LB: or CB: used to display list/combo boxes" & vbCrLf & _
        "Font: used in text sidebars" & vbCrLf & _
        "FSize: used in text sidebars" & vbCrLf & _
        "MinFSize: used in text sidebars" & vbCrLf & _
        "IMG: used in any menu item, except separator bars" & vbCrLf & _
        "Tip: used in any menu item" & vbCrLf & _
        "BColor: used in text/image sidebars" & vbCrLf & _
        "GColor: used in text/image sidebars" & vbCrLf & _
        "FColor: used in text sidebars" & vbCrLf & vbCrLf
    sMsg = sMsg & "The following do not have values & are entered only by name:" & vbCrLf & vbCrLf
    sMsg = sMsg & "Default used only in normal menu captions" & vbCrLf & _
        "Bold used in text sidebars" & vbCrLf & _
        "Italic used in text sidebars" & vbCrLf & _
        "Underline used in text sidebars" & vbCrLf & _
        "NoScroll used in text/image sidebars" & vbCrLf & _
        "Raised used only with separator bars" & vbCrLf & _
        "ImgBKG used only with normal menu items" & vbCrLf & _
        "Transparent used only with image sidebars (bitmap images only)"
Case Colors
    sTitle = "Help Formatting: Colors"
    sMsg = "BackColor. Flag is BColor:[color]" & vbCrLf & _
        "  Enter color of vbNull for menu back color" & vbCrLf & _
        "  For Image Sidebars enter -1 for back color of image" & vbCrLf & vbCrLf
    sMsg = sMsg & "ForeColor. Flag is FColor:[color]" & vbCrLf & _
        "  Only applies to Text Sidebars" & vbCrLf & vbCrLf
    sMsg = sMsg & "Gradients. Flag is GColor:[color]" & vbCrLf & _
        "  Applies to Image & Text Sidebars only." & vbCrLf & _
        "  The gradient is from BackColor to Gradient Color" & vbCrLf & _
        "  Enter color of vbNull for menu back color"
Case Images
    sTitle = "Help Formatting: Images"
    sMsg = "Image. Flag is IMG:[image ID]" & vbCrLf & vbCrLf
    sMsg = sMsg & "If image is an ImageList item, then" & vbCrLf & _
           "  - format is I[index]. Example for 2nd item: IMG:I2" & vbCrLf & vbCrLf
    sMsg = sMsg & "If image is a control on the form, then" & vbCrLf & _
           "  - format is [control name]. Example: IMG:myImage(1)" & vbCrLf & _
           "    Note: control must be on same form as the menu." & vbCrLf & vbCrLf
    sMsg = sMsg & "If image is a memory object, then" & vbCrLf & _
           "  - format is [handle]. Example: IMG:826982" & vbCrLf & vbCrLf
    sMsg = sMsg & "Special image flags:" & vbCrLf & vbCrLf & _
        "ImgBkg (not for sidebars) is used if the menu item is a bitmap and you do not want" & vbCrLf & _
        "    the image background to be transparent. By default bitmaps are made transparent." & vbCrLf & _
        "Transparent (only for image sidebars). Will make the sidebar image background transparent if possible."
        
Case SeparatorBars
    sTitle = "Help Formatting: Separator Bars"
    sMsg = "Non-Text separator bars: -[{Raised}]" & vbCrLf & _
        "  Caption must be a hyphen. The flag {Raised} is optional." & vbCrLf & vbCrLf
    sMsg = sMsg & "Text separator bars: -Caption[{Raised}]" & vbCrLf & _
        "  Caption is preceeded by a hyphen. The flag {Raised} is optional." & vbCrLf & vbCrLf
    sMsg = sMsg & "If {Raised} is provided, then the bar has a raised effect," & vbCrLf & _
        "otherwise the bar will have the standard sunken effect."
Case SidebarText
    sTitle = "Help Formatting: Text Sidebars"
    sMsg = "The text in a text-type sidebar can be anything you like." & vbCrLf & _
        "There are 3 special flags that can be added to the text to assist in line breaks." & vbCrLf & vbCrLf
    sMsg = sMsg & "<br!> between two words forces a line break" & vbCrLf & _
                "<br0> between two words prevents a line break" & vbCrLf & _
                "<br>  between two words is a break preference if needed." & vbCrLf & vbCrLf
    sMsg = sMsg & "The program will try to fit the text on one line in the sidebar (assuming no <br!> encountered)." & vbCrLf & _
        "If not possible, the font will be reduced to the authorized minimum font size" & vbCrLf & vbCrLf & _
        "If text still won't fit on one line, then the program will force line breaks in this order." & vbCrLf & _
        " 1. <br!> is always a line break" & vbCrLf & _
        " 2. <br> preferred break point" & vbCrLf & _
        " 3. Spaces are used as the last resort." & vbCrLf & vbCrLf & _
        "Also see help topic: FontAttributes"
Case FontAttributes
    sTitle = "Help Formatting: Font Attributes"
    sMsg = "Following attribute only applies to normal menu captions" & vbCrLf & vbCrLf
    sMsg = sMsg & "Flag: Default  This bolds the menu item." & vbCrLf & _
        "  Note: If a popup menu and an item was made default, it will automatically be bolded." & vbCrLf & vbCrLf
    sMsg = sMsg & "Following applies only to Text Sidebars" & vbCrLf & vbCrLf
    sMsg = sMsg & "Flag: Bold makes text bold" & vbCrLf & _
        "Flag: Italic makes text italicized" & vbCrLf & _
        "Flag: Underline makes text underlined" & vbCrLf & _
        "Flag: Font:[fontname] makes text display using that font name" & vbCrLf & _
        "Flag: FSize:[fontsize] makes text try to display using that font size" & vbCrLf & _
        "Flag: MinFSize:[fontsize] prevents scaling text below that font size when fitting text in sidebar." & vbCrLf & vbCrLf
    sMsg = sMsg & "Also see help topic: SidebarText"
Case NoScroll
    sTitle = "Help Formatting: NoScroll Flag"
    sMsg = "In order to display sidebars when a menu would scroll, the program converts " & _
        "scrolling sidebars to columns/panels. If this is not desirable, then you can " & _
        "add the flag NoScroll which will prevent the sidebar from displaying when " & _
        "menu scrolls and allow the menu to remain scrolling vs columns." & vbCrLf & vbCrLf
    sMsg = sMsg & "If this flag is set, should a menu height revert to a size where it " & _
        "would no longer scroll, the sidebar will reappear."
Case Alignments
    sTitle = "Help Formatting: Alignment"
    sMsg = "This flag only applies to sidebars." & vbCrLf & vbCrLf
    sMsg = sMsg & "Align:Bot will align the sidebar to the bottom of the menu" & vbCrLf & _
        "Align:Top will align the sidebar to the top of the menu" & vbCrLf & _
        "Align:Ctr will align sidebar in center of menu." & vbCrLf & _
        "  - Note: This flag is not required. Center alignment is default."
Case Tips
    sTitle = "Help Formatting: Tips"
    sMsg = "The TIP flag can be used with any menu item." & vbCrLf & vbCrLf
    sMsg = sMsg & "In order for tips to be displayed, you need to initialize cTips " & _
        "within the form that owns the menu and then call SetMenu and pass the " & _
        "initialized class to that function. Read the remarks in cTips." & vbCrLf & vbCrLf
    sMsg = sMsg & "Simply provide the flag Tip: followed by the tip you want displayed." & vbCrLf & vbCrLf
    sMsg = sMsg & "The tip will be returned to your form via the DisplayTip event for the cTip class."
Case States_Menu
    sTitle = "Custom Menu: States"
    sMsg = "This menu will display all 50 states and DC." & vbCrLf & vbCrLf
    sMsg = sMsg & "cTips CustomSelection Category: State" & vbCrLf & vbCrLf
    sMsg = sMsg & "Caption: {lvStates:[S]:ID:[id] | [other flags]}" & vbCrLf & vbCrLf & _
        "[S] is the 2-letter state you want shown as checked" & vbCrLf & _
        "  - if no item is to be checked, supply -1 for [S]" & vbCrLf & vbCrLf
    sMsg = sMsg & "[id] is a user-defined string value passed back thru cTips." & vbCrLf & vbCrLf
    sMsg = sMsg & "Example: United States{lvStates:IL:ID:mnuXyz|IMG:i5}" & vbCrLf & vbCrLf
    sMsg = sMsg & "Return Value: The 2-letter abbreviation of the state the user clicked"
Case Months_Menu
    sTitle = "Custom Menu: Months"
    sMsg = "This menu will display the 12 months of the year in 3 styles" & vbCrLf & vbCrLf
    sMsg = sMsg & "cTips CustomSelection Category: Month" & vbCrLf & vbCrLf
    sMsg = sMsg & "Caption: {lvMonths:[M]:Year:[Y]:Group:[G]:ID:[id] | [other flags]}" & vbCrLf & vbCrLf & _
        "[M] is the numerical value (1-12) of month to show as checked" & vbCrLf & _
        "  - supply -1 for [M] to not have any month checked" & vbCrLf & _
        "  - supply 0 for [M] to check the current month" & vbCrLf & vbCrLf & _
        "Year:[Y] is optional. It is the year for the month to display" & vbCrLf & _
        "  - supply 0 for [Y] or don't include flag to default to current year." & vbCrLf & vbCrLf & _
        "Group:[G] is optional. It is how the months will be displayed" & vbCrLf & _
        "  - If flag not provided, default is simple alphabetical list of months" & vbCrLf & _
        "  - Set [G] as CYQtr to group months by calendar year quarters" & vbCrLf & _
        "  - Set [G] as FYQtr to group months by fiscal year quarters" & vbCrLf & vbCrLf
    sMsg = sMsg & "[id] is a user-defined string value passed back thru cTips." & vbCrLf & vbCrLf
    sMsg = sMsg & "Example: Month to Report{lvMonths:-1:Group:FYQtr:ID:mnuXyz|IMG:i5}" & vbCrLf & vbCrLf
    sMsg = sMsg & "Return Value: The numerical value (1-12) of the month the user clicked"
Case WeekDays_Menu
    sTitle = "Custom Menu: Days of the Week"
    sMsg = "This menu will display the 7 days of the week" & vbCrLf & vbCrLf
    sMsg = sMsg & "cTips CustomSelection Category: WeekDay" & vbCrLf & vbCrLf
    sMsg = sMsg & "Caption: {lvDays:[D]:ID:[id] | [other flags]}" & vbCrLf & vbCrLf & _
        "[D] is the numerical value (1-7) of the weekday you want shown as checked" & vbCrLf & _
        "  - supply -1 for [D] to not have any day checked" & vbCrLf & _
        "  - supply 0 for [D] to check the current day of the week" & vbCrLf & vbCrLf
    sMsg = sMsg & "[id] is a user-defined string value passed back thru cTips." & vbCrLf & vbCrLf
    sMsg = sMsg & "Example: Work Schedule For...{lvDays:0:ID:mnuXyz|IMG:i5}" & vbCrLf & vbCrLf
    sMsg = sMsg & "Return Value: The numerical value (1-7) of the day the user clicked" & vbCrLf & _
        "Note: Whether Sunday or Monday is the 1st day of the week is dependent upon system settings"
Case DayOfMonth_Menu
    sTitle = "Custom Menu: Days of a Month"
    sMsg = "This menu will display the all days of a month" & vbCrLf & vbCrLf
    sMsg = sMsg & "cTips CustomSelection Category: DayOfMonth" & vbCrLf & vbCrLf
    sMsg = sMsg & "Caption: {lvMonth:[M]:Year:[Y]:Day:[D]:ID:[id] | [other flags]}" & vbCrLf & vbCrLf & _
        "[M] is the numerical value (1-12) of month to display" & vbCrLf & _
        "  - supply 0 for [M] to deault to current month" & vbCrLf & vbCrLf & _
        "Year:[Y] is optional. It is the year for the month to display" & vbCrLf & _
        "  - supply 0 for [Y] or don't include flag to default to current year." & vbCrLf & vbCrLf & _
        "Day:[D] is optional. It is the day of the month to show as checked" & vbCrLf & _
        "  - supply -1 for [D] or don't include flag to not check any date." & vbCrLf & _
        "  - supply 0 for [D] to check current date" & vbCrLf & vbCrLf
    sMsg = sMsg & "[id] is a user-defined string value passed back thru cTips." & vbCrLf & vbCrLf
    sMsg = sMsg & "Example: Orders for Month{lvMonth:0:Day:-1:ID:mnuXyz|IMG:i5}" & vbCrLf & vbCrLf
    sMsg = sMsg & "Return Value: A date value of the day the user clicked"
Case Colors_Menu
    sTitle = "Custom Menu: 24 Color Options"
    sMsg = "This menu will display the 24 color options. The list includes 23 colors " & vbCrLf & _
        "an additional option to select a color not in the list" & vbCrLf & vbCrLf
    sMsg = sMsg & "cTips CustomSelection Category: Color" & vbCrLf & vbCrLf
    sMsg = sMsg & "Caption: {lvColors:[C]:ID:[id] | [other flags]}" & vbCrLf & vbCrLf & _
        "[C] is the numerical value of the color you want shown as checked" & vbCrLf & _
        "  - supply -1 for [C] to not have any color checked" & vbCrLf & vbCrLf
    sMsg = sMsg & "[id] is a user-defined string value passed back thru cTips." & vbCrLf & vbCrLf
    sMsg = sMsg & "Example: Back Color{lvColors:vbWhite:ID:mnuXyz|IMG:i5}" & vbCrLf & vbCrLf
    sMsg = sMsg & "Return Value: The numerical value of the color the user clicked" & vbCrLf & vbCrLf & _
        "Note: If the value returned is -1 then the user canceled selection of a custom color"
Case Fonts_Menu
    sTitle = "Custom Menu: Fonts"
    sMsg = "This menu can display the all installed fonts and group them in a few ways" & vbCrLf & vbCrLf
    sMsg = sMsg & "cTips CustomSelection Category: Font" & vbCrLf & vbCrLf
    sMsg = sMsg & "Caption: {lvFonts:[F]:Type:[T]:Group:[G]:ID:[id] | [other flags]}" & vbCrLf & _
        "[F] is the name of the font to display" & vbCrLf & _
        "  - supply -1 for [F] to not have any font checked" & vbCrLf & vbCrLf & _
        "Type:[T] is optional. Default is to display all fonts" & vbCrLf & _
        "  - supply System for [T] to only display system fonts." & vbCrLf & _
        "  - supply TrueType for [T] to only display true-type fonts." & vbCrLf & vbCrLf & _
        "Group:[G] is optional. This option filters fonts to display." & vbCrLf & _
        "  - supply range for [G] to filter the fonts. The range is always in the format" & vbCrLf & _
        "    of Start Letter-End Letter. Example: L-V." & vbCrLf & vbCrLf
    sMsg = sMsg & "[id] is a user-defined string value passed back thru cTips." & vbCrLf & vbCrLf
    sMsg = sMsg & "Example: TrueType Fonts L-V{lvFonts:-1:Type:TrueType:Group:L-V:ID:mnuXyz:IMG:i5}" & vbCrLf & vbCrLf
    sMsg = sMsg & "Return Value: A string value of the font name the user clicked"
Case ListComboBoxes
    sTitle = "Help Formatting: List & Combo Boxes"
    sMsg = "You can bring the contents of a list box, file list or combo box to a menu." & vbCrLf & vbCrLf
    sMsg = sMsg & "Following are flags specifically for these types of menus." & vbCrLf & vbCrLf
    sMsg = sMsg & "LB:[list box or file list ID] The ID can be control name or hWnd" & vbCrLf & _
                "CB:[combo box ID] The ID can be control name or hWnd" & vbCrLf & _
                "Files:[Path] Only for LB: If the list box contains file names, then " & vbCrLf & _
                "     the program can display their icons if you include the above flag and also " & vbCrLf & _
                "     provide the optional value for the [Path]." & vbCrLf & _
                "     [Path] If the paths are included in the file names supply -1, otherwise, " & vbCrLf & _
                "            supply the path name in this value" & vbCrLf & vbCrLf
    sMsg = sMsg & "Notes: Combo boxes. When user selects a combo box item from the menu, then " & vbCrLf & _
                "       the combo box will receive a Click event." & vbCrLf & _
                "       List boxes. Same as combo boxes. In addition, multi-select listboxes are supported." & vbCrLf & vbCrLf
    sMsg = sMsg & "Finally. Owner-drawn list boxes and combo boxes are not supported at this time."
Case Drives_Menu
    sTitle = "Custom Menu: Drive Listing"
    sMsg = "This menu will display the drives on your computer" & vbCrLf & vbCrLf
    sMsg = sMsg & "cTips CustomSelection Category: Drive" & vbCrLf & vbCrLf
    sMsg = sMsg & "Caption: {lvDrives:[D]:ID:[id] | [other flags]}" & vbCrLf & vbCrLf & _
        "[D] is the string value of the drive you want shown as checked" & vbCrLf & _
        "  - supply -1 for [D] to not have any drive checked" & vbCrLf & vbCrLf
    sMsg = sMsg & "[id] is a user-defined string value passed back thru cTips." & vbCrLf & vbCrLf
    sMsg = sMsg & "Example: Select a Drive{lvDrives:C:\:ID:mnuXyz|IMG:i5}" & vbCrLf & vbCrLf
    sMsg = sMsg & "Return Value: The string value of the drive the user clicked"
End Select

MsgBox sMsg, vbInformation + vbOKOnly, sTitle
   
End Property
