VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl rtbSyntax 
   ClientHeight    =   2895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4440
   ScaleHeight     =   2895
   ScaleWidth      =   4440
   ToolboxBitmap   =   "rtbSyntax.ctx":0000
   Begin VB.TextBox SaveText 
      Height          =   615
      Left            =   3360
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin RichTextLib.RichTextBox OriginalText 
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"rtbSyntax.ctx":0532
   End
   Begin VB.PictureBox k 
      Height          =   375
      Left            =   2160
      ScaleHeight     =   315
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox f 
      Height          =   375
      Left            =   1560
      ScaleHeight     =   315
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox r 
      Height          =   375
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox c 
      Height          =   375
      Left            =   360
      ScaleHeight     =   315
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   735
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      _Version        =   393217
      ScrollBars      =   3
      RightMargin     =   1.00000e5
      TextRTF         =   $"rtbSyntax.ctx":0638
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "rtbSyntax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' rtbSyntax Control.
' Aaron Bennear
' April 7, 2002
'
' Adds syntax highlighting to the RichTextBox control. A design goal was to
' accomplish this with only one file, the ctl file. (There is an associated
' ctx file, but this is not required and simply supplies the toolbox icon.)
' This makes the control a drop in replacement for the standard RichTextBox,
' and very easy to add to a project. Therefore, the list of keywords
' highlighted is composed of constants rather than loaded from a file. Also,
' the syntax parsing is not object oriented and does not use classes.
'
' The list of keywords to highlight can be modified. However, the syntax
' highlighting is designed to handle VBScript code. It therefore does not
' handle non-VB conventions such as multiline comments.
'
' As much as possible, the control uses delegation to expose the properties
' and events of the underlying RichTextBox control. The Data and OLE related
' properties proved problematic and are not supported. Also, properties that
' are read-only at run time can not be delegated from user control to
' constituent control. For RichTextBox these are: Appearance, DisableNoScroll,
' MultiLine, and Scrollbars. Of course, these property limitations can dealt
' with by editing the control.
'
' The parser scans line by line. Within each line the text is read left to
' right and the type of syntax context is kept track of. As the context
' changes, from keyword to string for instance, coloring is done for the
' subsection just completed.
'
' As the text in the RichTextBox changes, the parsing needs to be redone. To
' improve performance, an attempt is made to only reparse the lines that are
' changed. This is done by keeping track of the current and previous point of
' insertion. Also, an API call is used to disable the repainting of the
' RichTextBox as it is being colored, to prevent unsightly selection changes
' from flashing by.
'
' The tedious task of matching positions in the text being parsed with its
' absolute position in the RichTextBox is made more so by the fact that VB
' string functions are 1-indexed while the RichTextBox is 0-indexed.
'
' The control exposed only one additionaly method: HighlightRefresh. This
' reparses all the text. It should not be necessary to call this function
' from code using the control, but is provided just in case.
'

' MODIFIED BY DEAN CAMERA, 2003, FOR etURP

Option Explicit

Dim UseSTax As Integer

Dim ScrollMin
Dim KillSub As Boolean

Const HASHLINE = "#"
Const COMMENT = "//"
Const DELIMITER = vbTab & " ()+,"
Const RESERVED As String = " __NQC__ __event_src __res __sensor __type ANIMATION abs acquire asm break case catch const continue do default else false for if inline int monitor repeat return sign start stop sub switch task true until void while "
Const FUNC_OBJa As String = " SetSensor SetSensorType SetSensorMode ClearSensor SensorValue SensorType SensorMode SensorValueBool SensorValueRaw SetSensorLowerLimit SetSensorUpperLimit SetSensorHysteresis CalibrateSensor SetOutput SetDirection SetPower OutputStatus On Off Float Fwd Rev Toggle OnFwd OnRev OnFor SetGlobalOutput SetGlobalDirection SetMaxPower GlobalOutputStatus PlaySound PlayTone MuteSound UnmuteSound ClearSound SelectSounds SelectDisplay SetUserDisplay Message ClearMessage SendMessage SetTxPower SetSerialComm SetSerialPacket SetSerialData SerialData SendSerial SendVLL ClearTimer Timer SetTimer FastTimer ClearCounter IncCounter DecCounter Counter SetPriority ActiveEvents CurrentEvents Event SetEvent ClearEvent ClearAllEvents EventState CalibrateEvent SetUpperLimit UpperLimit SetLowerLimit LowerLimit SetHysteresis Hysteresis SetClickTime ClickTime SetClickCounter ClickCounter"
Const FUNC_OBJb As String = "SetSensor ClickTime SetCounterLimit SetTimerLimit CreateDatalog AddToDatalog UploadDatalog Wait StopAllTasks Random SetRandomSeed SetSleepTime SleepNow Program SelectProgram BatteryLevel FirmwareVersion Watch SetWatch SetScoutRules ScoutRules SetScoutMode SetEventFeedback EventFeedback SetLight Drive OnWait OnWaitDifferent ClearTachoCounter TachoCount TachoSpeed ExternalMotorRunning AGC ClearRelationTable SetAnimation SetLED VLL LED "
Const FUNC_OBJ As String = FUNC_OBJa & FUNC_OBJb

Const KEYWORDa As String = " SENSOR_1 SENSOR_2 SENSOR_3 SENSOR_L SENSOR_M SENSOR_R SENSOR_MODE_RAW SENSOR_MODE_BOOL SENSOR_MODE_EDGE SENSOR_MODE_PULSE SENSOR_MODE_PERCENT SENSOR_MODE_CELSIUS SENSOR_MODE_FAHRENHEIT SENSOR_MODE_ROTATION SENSOR_TYPE_NONE SENSOR_TYPE_TOUCH SENSOR_TYPE_TEMPERATURE SENSOR_TYPE_LIGHT SENSOR_TYPE_ROTATION SENSOR_TOUCH SENSOR_LIGHT SENSOR_ROTATION SENSOR_PULSE SENSOR_EDGE SENSOR_CELSIUS SENSOR_FAHRENHEIT OUT_A OUT_B OUT_C OUT_L OUT_R OUT_X OUT_OFF OUT_ON OUT_FLOAT OUT_REV OUT_TOGGLE OUT_FWD OUT_FULL OUT_LOW OUT_HALF SOUND_CLICK SOUND_DOUBLE_BEEP SOUND_DOWN SOUND_UP SOUND_LOW_BEEP SOUND_FAST_UP DIRSPEED DISPLAY_WATCH DISPLAY_SENSOR_1 DISPLAY_SENSOR_2 DISPLAY_SENSOR_3 DISPLAY_OUT_A DISPLAY_OUT_B DISPLAY_OUT_C DISPLAY_USER TX_POWER_LO TX_POWER_HI SERIAL_COMM_DEFAULT SERIAL_COMM_4800 SERIAL_COMM_DUTY25 SERIAL_COMM_76KHZ SERIAL_PACKET_DEFAULT SERIAL_PACKET_PREAMBLE SERIAL_PACKET_NEGATED SERIAL_PACKET_CHECKSUM SERIAL_PACKET_RCX ACQUIRE_OUT_A ACQUIRE_OUT_B ACQUIRE_OUT_C"
Const KEYWORDb As String = " ACQUIRE_SOUND ACQUIRE_USER_1 ACQUIRE_USER_2 ACQUIRE_USER_3 ACQUIRE_USER_4 ACQUIRE_LED EVENT_TYPE_PRESSED EVENT_TYPE_RELEASED EVENT_TYPE_PULSE EVENT_TYPE_EDGE EVENT_TYPE_FASTCHANGE EVENT_TYPE_LOW EVENT_TYPE_NORMAL EVENT_TYPE_HIGH EVENT_TYPE_CLICK EVENT_TYPE_DOUBLECLICK EVENT_TYPE_MESSAGE EVENT_TYPE_ENTRY_FOUND EVENT_TYPE_MSG_DISCARD EVENT_TYPE_MSG_RECEIVED EVENT_TYPE_VLL_MSG_RECEIVED EVENT_TYPE_ENTRY_CHANGED EVENT_1_PRESSED EVENT_1_RELEASED EVENT_2_PRESSED EVENT_2_RELEASED EVENT_LIGHT_HIGH EVENT_LIGHT_NORMAL EVENT_LIGHT_LOW EVENT_LIGHT_CLICK EVENT_LIGHT_DOUBLECLICK EVENT_COUNTER_0 EVENT_COUNTER_1 EVENT_TIMER_0 EVENT_TIMER_1 EVENT_TIMER_2 EVENT_MESSAGE LED_MODE_ON LED_MODE_BLINK LED_MODE_DURATION LED_MODE_SCALE LED_MODE_SCALE_BLINK LED_MODE_SCALE_DURATION LED_MODE_RED_SCALE LED_MODE_RED_SCALE_BLINK LED_MODE_GREEN_SCALE LED_MODE_GREEN_SCALE_BLINK LED_MODE_YELLOW LED_MODE_YELLOW_BLINK LED_MODE_YELLOW_DURATION LED_MODE_VLL LED_MODE_VLL_BLINK LED_MODE_VLL_DURATION ANIMATION_SCAN"
Const KEYWORDc As String = " ANIMATION_SPARKLE ANIMATION_FLASH ANIMATION_RED_TO_GREEN ANIMATION_GREEN_TO_RED ANIMATION_POINT_FORWARD ANIMATION_ALARM ANIMATION_THINKING "
Const KEYWORD As String = KEYWORDa & KEYWORDb & KEYWORDc

Const KEYWORD_PAD As String = " "

Public Enum ColorIndex
colour_KEYWORD = 0
colour_FUNC_OBJ = 1
colour_RESERVED = 2
colour_COMMENT = 3
End Enum

Const RGB_STRING As String = "100,0,0"
Const RGB_HASHLINE As String = "100,0,0"

Enum SyntaxTypes
    ColorComment = 0
    ColorString = 1
    ColorReserved = 2
    ColorFuncObj = 3
    ColorDelimiter = 4
    ColorNormal = 5
    ColorKeyword = 6
    colorhashline = 7
End Enum

' Global variable used to suppress parsing until the end of a series of
' changes. Or, in the Change event itself to prevent cascaded Change events.
Private mbInChange As Boolean

' RGB values derived from constants
Private mrgbComment As Long
Private mrgbString As Long
Private mrgbReserved As Long
Private mrgbFuncObj As Long
Private mrgbDelimiter As Long
Private mrgbKeyword As Long
Private mrgbNormal As Long
Private mrgbHashLine As Long

Dim RGB_COMMENT As String
Dim RGB_FUNC_OBJ As String
Dim RGB_NORMAL As String
Dim RGB_DELIMITER As String
Dim RGB_RESERVED As String
Dim RGB_KEYWORD As String

' One WinAPI call. Used to suppress repainting during parsing.
Private Const WM_SETREDRAW = &HB
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' Keeping track of current and previous insertion point. Used to determine
' what portion of text has changed.
Private mlPrevSelStart As Long
Private mlCurSelStart As Long

'
' Delegation code
'

'Default Property Values:
Const m_def_ForeColor = 0
Const m_def_hWnd = 0
'Property Variables:
Dim m_ForeColor As Long
Dim m_hWnd As Long
'Event Declarations:
Event Change() 'MappingInfo=rtb,rtb,-1,Change
Attribute Change.VB_Description = "Indicates that the contents of a control have changed."
Event Click() 'MappingInfo=rtb,rtb,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=rtb,rtb,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=rtb,rtb,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=rtb,rtb,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=rtb,rtb,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=rtb,rtb,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses a mouse button."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=rtb,rtb,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=rtb,rtb,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user presses and releases a mouse button."
Event SelChange() 'MappingInfo=rtb,rtb,-1,SelChange
Attribute SelChange.VB_Description = "Occurs when the current selection of text in the RichTextBox control has changed or the insertion point has moved."

'
' Sub UserControl_Initialize
' Position constituate control, call initialization.
'
Private Sub UserControl_Initialize()
    r.Backcolor = GetSetting("En-Tech URP", "Syntax", "ReservedC", RGB(0, 0, 0))
    f.Backcolor = GetSetting("En-Tech URP", "Syntax", "FunctionC", RGB(0, 128, 0))
    c.Backcolor = GetSetting("En-Tech URP", "Syntax", "CommentC", RGB(0, 0, 200))
    k.Backcolor = GetSetting("En-Tech URP", "Syntax", "KeyWordC", RGB(150, 0, 0))

    RGB_COMMENT = "0,128,0"
    RGB_FUNC_OBJ = "0,0,200"
    RGB_NORMAL = "0,0,0"
    RGB_DELIMITER = "0,0,0"
    RGB_RESERVED = "0,0,0"
    RGB_KEYWORD = "0,100,0"

    rtb.Top = 0
    rtb.Left = 0
    InitParser
    mlPrevSelStart = 0
End Sub

Sub NewColorValue(WColor As SyntaxTypes, ColorString As String)
    On Error Resume Next
    If WColor = ColorComment Then
        c.Backcolor = ColorString
    ElseIf WColor = ColorFuncObj Then
        f.Backcolor = ColorString
    ElseIf WColor = ColorNormal Then
        RGB_NORMAL = ColorString
    ElseIf WColor = ColorDelimiter Then
        RGB_DELIMITER = ColorString
    ElseIf WColor = ColorKeyword Then
        k.Backcolor = ColorString
    ElseIf WColor = ColorReserved Then
        r.Backcolor = ColorString
    End If

    InitParser
End Sub

'
' Sub InitParser
' Derive color values.
'

Private Sub InitParser()
    Dim vArr

    vArr = Split(RGB_HASHLINE, ",")
    mrgbHashLine = RGB(vArr(0), vArr(1), vArr(2))

    vArr = Split(RGB_COMMENT, ",")
    mrgbComment = c.Backcolor

    vArr = Split(RGB_STRING, ",")
    mrgbString = RGB(vArr(0), vArr(1), vArr(2))

    vArr = Split(RGB_RESERVED, ",")
    mrgbReserved = r.Backcolor

    vArr = Split(RGB_FUNC_OBJ, ",")
    mrgbFuncObj = f.Backcolor

    vArr = Split(RGB_DELIMITER, ",")
    mrgbDelimiter = RGB(vArr(0), vArr(1), vArr(2))

    vArr = Split(RGB_NORMAL, ",")
    mrgbNormal = RGB(vArr(0), vArr(1), vArr(2))

    vArr = Split(RGB_KEYWORD, ",")
    mrgbKeyword = k.Backcolor
End Sub

Sub HighlightRefreshChange()
    UseSTax = GetSetting("En-Tech URP", "Options", "SyntaxColour", 1)
    If UseSTax = 0 Then Exit Sub

    If mbInChange = True Then
        ' change is being blocked or deferred
        GoTo ExitSub
    End If

    ' suppress change events generated during this change event
    '
    mbInChange = True


    Dim srtbText As String      ' working string
    ' add final cariage return so last line is processed
    srtbText = rtb.Text & vbCrLf

    ' preserve selection and restore at end
    '
    Dim lOrigSelStart As Long
    Dim lOrigSelLength As Long
    lOrigSelStart = rtb.SelStart
    lOrigSelLength = rtb.SelLength


    Dim lStartPos As Long
    Dim lEndPos As Long

    If mlPrevSelStart < rtb.SelStart Then
        lStartPos = mlPrevSelStart
        lEndPos = rtb.SelStart
    Else
        lStartPos = rtb.SelStart
        lEndPos = mlPrevSelStart
    End If


    If lStartPos > 1 Then
        ' set start position to beginning of line
        If InStrRev(srtbText, vbCrLf, lStartPos - 1) > 0 Then
            lStartPos = InStrRev(srtbText, vbCrLf, lStartPos - 1) + Len(vbCrLf) - 1
        Else
            lStartPos = 0
        End If
    Else
        lStartPos = 0
    End If

    ' set end position to end of line
    If InStr(lEndPos + 1, srtbText, vbCrLf) > 0 Then
        lEndPos = InStr(rtb.SelStart + 1, srtbText, vbCrLf) - 1
    Else
        lEndPos = Len(srtbText) - 1
    End If


    ' send affected text to the parser, along with its position in the
    ' RichTextBox

    Dim X As Long

    If lStartPos <> lEndPos Then
        X = SendMessage(rtb.hWnd, WM_SETREDRAW, 0, 0)
        ParseLines rtb.Text, rtb, lStartPos
    End If

    rtb.SelStart = lOrigSelStart
    rtb.SelLength = lOrigSelLength

    mbInChange = False

ExitSub:
End Sub

'
' Sub rtb_Change
' Determine the changed region and feed to the parser.
'
Private Sub rtb_Change()

    RaiseEvent Change

    UseSTax = GetSetting("En-Tech URP", "Options", "SyntaxColour", 1)
    If UseSTax = 0 Then Exit Sub

    If mbInChange = True Then
        ' change is being blocked or deferred
        GoTo ExitSub
    End If

    ' suppress change events generated during this change event
    '
    mbInChange = "True"


    Dim srtbText As String      ' working string
    ' add final cariage return so last line is processed
    srtbText = rtb.Text & vbCrLf

    ' preserve selection and restore at end
    '
    Dim lOrigSelStart As Long
    Dim lOrigSelLength As Long
    lOrigSelStart = rtb.SelStart
    lOrigSelLength = rtb.SelLength


    Dim lStartPos As Long
    Dim lEndPos As Long

    If mlPrevSelStart < rtb.SelStart Then
        lStartPos = mlPrevSelStart
        lEndPos = rtb.SelStart
    Else
        lStartPos = rtb.SelStart
        lEndPos = mlPrevSelStart
    End If


    If lStartPos > 1 Then
        ' set start position to beginning of line
        If InStrRev(srtbText, vbCrLf, lStartPos - 1) > 0 Then
            lStartPos = InStrRev(srtbText, vbCrLf, lStartPos - 1) + Len(vbCrLf) - 1
        Else
            lStartPos = 0
        End If
    Else
        lStartPos = 0
    End If

    ' set end position to end of line
    If InStr(lEndPos + 1, srtbText, vbCrLf) > 0 Then
        lEndPos = InStr(rtb.SelStart + 1, srtbText, vbCrLf) - 1
    Else
        lEndPos = Len(srtbText) - 1
    End If


    Dim X As Long

    'prevent textbox from repainting
    X = SendMessage(rtb.hWnd, WM_SETREDRAW, 0, 0)

    ' send affected text to the parser, along with its position in the
    ' RichTextBox
    If lStartPos <> lEndPos Then
        ParseLines Mid(srtbText, lStartPos + 1, lEndPos - lStartPos), rtb, lStartPos
    End If

    rtb.SelStart = lOrigSelStart
    rtb.SelLength = lOrigSelLength

    'allow texbox to repaint
    X = SendMessage(rtb.hWnd, WM_SETREDRAW, 1, 0)
    'force repaint
    rtb.Refresh

    mbInChange = False

ExitSub:

End Sub

'
' Sub rtb_SelChange
' Keep track of previous SelStart to allow determination of
' affected region.
'
Private Sub rtb_SelChange()
    RaiseEvent SelChange

    mlPrevSelStart = mlCurSelStart
    mlCurSelStart = rtb.SelStart


End Sub

'
' Sub rtb_KeyDown
' Normally, tabbing leaves the control, but instead, we want to insert
' tab into edited text.
'
Private Sub rtb_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)

    If KeyCode = Asc(vbTab) Then  ' TAB key was pressed.
        ' Ignore the TAB key, so focus doesn't leave the control
        KeyCode = 0

        ' Replace selected text with the tab character
        rtb.SelText = vbTab
    End If


End Sub

'
' Sub HighlightRefresh
' Manipulate tracked previous selection and current selection
' to force reparsing of entire text.
'
Public Sub HighlightRefresh()
    'prevent textbox from repainting
    Dim X As Long
    X = SendMessage(rtb.hWnd, WM_SETREDRAW, 0, 0)

    Dim lOrigSelStart As Long
    Dim lOrigSelLength As Long
    lOrigSelStart = rtb.SelStart
    lOrigSelLength = rtb.SelLength

    mlPrevSelStart = 0

    rtb.SelStart = 0
    rtb.SelLength = Len(rtb.Text)
    rtb.SelBold = False
    rtb.SelColor = RGB(0, 0, 0)

    UseSTax = GetSetting("En-Tech URP", "Options", "SyntaxColour", 1)
    If UseSTax = 1 Then
        HighlightRefreshChange
    End If

    Dim CursorTemp As Integer
    CursorTemp = GetSetting("En-Tech URP", "Options", "Cursor", 1)

    If CursorTemp = 1 Then
        rtb.SelStart = 0
        rtb.SelLength = 0
    Else
        rtb.SelStart = Len(rtb.Text)
        rtb.SelLength = 0
    End If

    'allow texbox to repaint
    X = SendMessage(rtb.hWnd, WM_SETREDRAW, 1, 0)
    'force repaint
    rtb.Refresh

    If CursorTemp = 1 Then
        rtb.SelStart = 0
        rtb.SelLength = 0
    Else
        rtb.SelStart = Len(rtb.Text)
        rtb.SelLength = 0
    End If
End Sub

'
' Sub ParseLines
' Feed text, line by line, to the parser.
'
Private Sub ParseLines(ByVal s As String, rtb As RichTextBox, ByVal RTBPos As Long)
    Dim lStartPos As Long
    Dim lEndPos As Long

    lStartPos = 1

    s = s & vbCrLf
    lEndPos = InStr(lStartPos, s, vbCrLf)
    Do While lEndPos > 0
        ParseLine Mid(s, lStartPos, lEndPos - lStartPos), rtb, RTBPos + lStartPos - 1
        lStartPos = lEndPos + Len(vbCrLf)
        lEndPos = InStr(lStartPos, s, vbCrLf)
    Loop
End Sub

'
' Sub ParseLine
' Lines are treated independently. Parseline is the main parsing code. Scan
' line from left to right, emitting text to be colored.
'
Private Sub ParseLine(ByVal s As String, rtb As RichTextBox, ByVal RTBPos As Long)

    Dim bInString As Boolean    ' are we in a quoted string?
    bInString = False

    Dim bInWord As Boolean      ' are we in a word? (not a string, comment,
    ' or delimiter)
    bInWord = False

    Dim sCurString As String        ' the current set of characters
    Dim lCurStringStart As Long     '   - where it starts
    Dim scurchar As String          ' the current character

    Dim I As Long

    Static LineTemp As String
    Static SingSChar As String
    Static InHashLine As Boolean

    KillSub = False
    InHashLine = False

    For I = 1 To Len(s)
        scurchar = Mid(s, I, 1)
        LineTemp = s
        If scurchar = HASHLINE Then
            If Not bInString Then
                If bInWord Then
                    ' before we encounterd the line we were processing a word
                    Highlight rtb, ParseWord(sCurString), lCurStringStart + RTBPos - 1, I - lCurStringStart
                    sCurString = ""
                    bInWord = True
                    InHashLine = True
                End If
                Highlight rtb, colorhashline, I + RTBPos - 1, Len(s) - I + 1
                KillSub = True
            End If
        End If

        SingSChar = Mid(LineTemp, I, 1)
        scurchar = Mid(LineTemp, I, 2)
        If scurchar = COMMENT Then
            ' if comment character occurs within a quoted string, it doesn't
            ' count
            If Not bInString Then
                ' this is a comment. we are done with the line
                If bInWord Then
                    ' before we encounterd the comment we were processing a word
                    If InHashLine = True Then
                        Highlight rtb, colorhashline, lCurStringStart + RTBPos - 1, I - lCurStringStart
                    Else
                        Highlight rtb, ParseWord(sCurString), lCurStringStart + RTBPos - 1, I - lCurStringStart
                    End If
                    sCurString = ""
                    bInWord = False
                End If
                Highlight rtb, ColorComment, I + RTBPos - 1, Len(s) - I + 1
                GoTo ExitSub    ' rest of line is comment
            End If
        End If

        If KillSub = True Then GoTo Next_i

        scurchar = Mid(s, I, 1)

        If scurchar = """" Then
            ' if not already in a string, then this quote begins a string
            ' otherwise, we are in a string, and this quote ends it
            If bInString Then
                sCurString = sCurString & scurchar
                Highlight rtb, ColorString, lCurStringStart + RTBPos - 1, I - lCurStringStart + 1
                sCurString = ""
                bInString = False
            Else
                If bInWord Then
                    ' before we encounterd the string we were processing a word
                    Highlight rtb, ParseWord(sCurString), lCurStringStart + RTBPos - 1, I - lCurStringStart
                    sCurString = ""
                    bInWord = False
                End If

                bInString = True
                sCurString = scurchar
                lCurStringStart = I
            End If

            GoTo Next_i ' get next character
        End If

        If InStr(1, DELIMITER, scurchar) > 0 Then
            If bInWord Then
                ' before we encounterd the delimiter we were processing a word
                Highlight rtb, ParseWord(sCurString), lCurStringStart + RTBPos - 1, I - lCurStringStart
                sCurString = ""
                bInWord = False
                Highlight rtb, ParseWord(sCurString), I + RTBPos - 1, Len(s) - I + 1
            End If

            Highlight rtb, ColorDelimiter, I + RTBPos - 1, 1
            GoTo Next_i
        End If

        If (Not bInWord) And (Not bInString) Then
            bInWord = True
            sCurString = scurchar
            lCurStringStart = I

            GoTo Next_i ' get next character
        End If

        ' add current character to the "word" we are in the middle of
        sCurString = sCurString & scurchar
Next_i:                             ' VB style continue
    Next

    If bInString Then
        ' before we encounterd the end of the line we were processing a string
        Highlight rtb, ColorString, lCurStringStart + RTBPos - 1, I - lCurStringStart
    ElseIf bInWord Then
        ' before we encounterd the end of the line we were processing a word
        Highlight rtb, ParseWord(sCurString), lCurStringStart + RTBPos - 1, I - lCurStringStart
    End If

ExitSub:
    Exit Sub
End Sub

'
' Function ParseWord
' Determine color for this word by checking for its existence in the keyword
' lists. The word being checked it padded with spaces to prevent matches
' with substrings of keywords.
'
Private Function ParseWord(ByVal Word As String) As SyntaxTypes
    If InStr(1, RESERVED, KEYWORD_PAD & Word & KEYWORD_PAD, vbTextCompare) > 0 Then
        ParseWord = ColorReserved
    ElseIf InStr(1, FUNC_OBJ, KEYWORD_PAD & Word & KEYWORD_PAD, vbTextCompare) > 0 Then
        ParseWord = ColorFuncObj
    ElseIf InStr(1, KEYWORD, KEYWORD_PAD & Word & KEYWORD_PAD, vbTextCompare) > 0 Then
        ParseWord = ColorKeyword
    Else
        ParseWord = ColorNormal
    End If
End Function

'
' Sub Highlight
' Color this range in the RichTextBox. Note that you could also apply bold,
' italic, etc. to the selection at the same time.
'
Private Sub Highlight(rtb As RichTextBox, SyntaxType As SyntaxTypes, StartPos As Long, Length As Long)

    rtb.SelStart = StartPos
    rtb.SelLength = Length
    rtb.SelBold = False
    Select Case SyntaxType
        Case SyntaxTypes.ColorComment
            rtb.SelColor = mrgbComment
        Case SyntaxTypes.ColorString
            rtb.SelColor = mrgbString
        Case SyntaxTypes.ColorReserved
            rtb.SelColor = mrgbReserved
            rtb.SelBold = True
        Case SyntaxTypes.ColorFuncObj
            rtb.SelColor = mrgbFuncObj
        Case SyntaxTypes.ColorDelimiter
            rtb.SelColor = mrgbDelimiter
        Case SyntaxTypes.ColorKeyword
            rtb.SelColor = mrgbKeyword
        Case SyntaxTypes.colorhashline
            rtb.SelColor = mrgbHashLine
        Case Else
            rtb.SelColor = mrgbNormal
    End Select

    'If InHashLine = True Then rtb.SelColor = mrgbHashLine
End Sub

'
' Sub UserControl_Resize
' Constituate control is always same size as user control.
'
Private Sub UserControl_Resize()
    rtb.Width = UserControl.ScaleWidth
    rtb.Height = UserControl.ScaleHeight
End Sub

' *****************************************************************************
' Properties
' For the most part this code is generated by the VB ActiveX Control Wizard
' *****************************************************************************


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,AutoVerbMenu
Public Property Get AutoVerbMenu() As Boolean
Attribute AutoVerbMenu.VB_Description = "Returns/sets a value that indicating whether the selected object's verbs will be displayed in a popup menu when the right mouse button is clicked."
    AutoVerbMenu = rtb.AutoVerbMenu
End Property

Public Property Let AutoVerbMenu(ByVal New_AutoVerbMenu As Boolean)
    rtb.AutoVerbMenu() = New_AutoVerbMenu
    PropertyChanged "AutoVerbMenu"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,BackColor
Public Property Get Backcolor() As OLE_COLOR
Attribute Backcolor.VB_Description = "Returns/sets the background color of an object."
    Backcolor = rtb.Backcolor
End Property

Public Property Let Backcolor(ByVal New_BackColor As OLE_COLOR)
    rtb.Backcolor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = rtb.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConstants)
    rtb.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,BulletIndent
Public Property Get BulletIndent() As Single
Attribute BulletIndent.VB_Description = "Returns or sets the amount of indent used in a RichTextBox control when SelBullet is set to True."
    BulletIndent = rtb.BulletIndent
End Property

Public Property Let BulletIndent(ByVal New_BulletIndent As Single)
    rtb.BulletIndent() = New_BulletIndent
    PropertyChanged "BulletIndent"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = rtb.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    rtb.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,FileName
Public Property Get FileName() As String
Attribute FileName.VB_Description = "Returns/sets the filename of the file loaded into the RichTextBox control at design time."
    FileName = rtb.FileName
End Property

Public Property Let FileName(ByVal New_FileName As String)
    rtb.FileName() = New_FileName

    PropertyChanged "FileName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = rtb.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set rtb.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,HideSelection
Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Returns/sets a value that specifies if the selected item remains highlighted when a control loses focus."
    HideSelection = rtb.HideSelection
End Property

Public Property Let HideSelection(ByVal New_HideSelection As Boolean)
    rtb.HideSelection() = New_HideSelection
    PropertyChanged "HideSelection"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = m_hWnd
End Property

Public Property Let hWnd(ByVal New_hWnd As Long)
    m_hWnd = New_hWnd
    PropertyChanged "hWnd"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Returns/sets a value indicating whether the contents in a RichTextBox control can be edited."
    Locked = rtb.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    rtb.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets a value indicating whether there is a maximum number of characters a RichTextBox control can hold and, if so, specifies the maximum number of characters."
    MaxLength = rtb.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    rtb.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = rtb.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set rtb.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets a value indicating the type of mouse pointer displayed when the mouse is over the control at run time."
    MousePointer = rtb.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    rtb.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,RightMargin
Public Property Get RightMargin() As Single
Attribute RightMargin.VB_Description = "Sets the right margin used for textwrap, centering, etc."
    RightMargin = rtb.RightMargin
End Property

Public Property Let RightMargin(ByVal New_RightMargin As Single)
    rtb.RightMargin() = New_RightMargin
    PropertyChanged "RightMargin"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in an object."
    Text = rtb.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    mbInChange = True
    rtb.Text() = New_Text
    mbInChange = False
    HighlightRefresh

    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,Find
Public Function Find(ByVal bstrString As String, Optional ByVal vStart As Variant, Optional ByVal vEnd As Variant, Optional ByVal vOptions As Variant) As Long
Attribute Find.VB_Description = "Searches the text in a RichTextBox control for a given string."
    Find = rtb.Find(bstrString, vStart, vEnd, vOptions)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,GetLineFromChar
Public Function GetLineFromChar(ByVal lChar As Long) As Long
Attribute GetLineFromChar.VB_Description = "Returns the number of the line containing a specified character position in a RichTextBox control."
    GetLineFromChar = rtb.GetLineFromChar(lChar)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,LoadFile
Public Sub LoadFile(ByVal bstrFilename As String, Optional ByVal vFileType As Variant)
Attribute LoadFile.VB_Description = "Loads an .RTF file or text file into a RichTextBox control."
    OriginalText.LoadFile bstrFilename
    rtb.LoadFile bstrFilename
    HighlightRefresh
End Sub

Public Function IsChanged() As Boolean
    IsChanged = True
    If rtb.Text = OriginalText.Text Then IsChanged = False
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a control."
    rtb.Refresh
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,SaveFile
Public Sub SaveFile(ByVal bstrFilename As String, Optional ByVal vFlags As Variant)
Attribute SaveFile.VB_Description = "Saves the contents of a RichTextBox control to a file."
    SaveText.Text = rtb.Text

    Open bstrFilename For Output As #1
    Print #1, SaveText.Text
    Close #1
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,SelPrint
Public Sub SelPrint(ByVal lHDC As Long, Optional ByVal vStartDoc As Variant)
Attribute SelPrint.VB_Description = "Sends formatted text in a RichTextBox control to a device for printing."
    rtb.SelPrint lHDC, vStartDoc
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,Span
Public Sub Span(ByVal bstrCharacterSet As String, Optional ByVal vForward As Variant, Optional ByVal vNegate As Variant)
Attribute Span.VB_Description = "Selects text in a RichTextBox control based on a set of specified characters."
    rtb.Span bstrCharacterSet, vForward, vNegate
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,UpTo
Public Sub UpTo(ByVal bstrCharacterSet As String, Optional ByVal vForward As Variant, Optional ByVal vNegate As Variant)
Attribute UpTo.VB_Description = "Moves the insertion point up to, but not including, the first character that is a member of the specified character set in a RichTextBox control."
    rtb.UpTo bstrCharacterSet, vForward, vNegate
End Sub

Private Sub rtb_Click()
    RaiseEvent Click
End Sub

Private Sub rtb_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub rtb_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub rtb_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub rtb_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub rtb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub rtb_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_ForeColor = m_def_ForeColor
    m_hWnd = m_def_hWnd
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ' prevent parsing while file is loading
    mbInChange = True

    rtb.AutoVerbMenu = PropBag.ReadProperty("AutoVerbMenu", False)
    rtb.Backcolor = PropBag.ReadProperty("BackColor", &H80000005)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    rtb.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    rtb.BulletIndent = PropBag.ReadProperty("BulletIndent", 0)
    rtb.Enabled = PropBag.ReadProperty("Enabled", True)
    rtb.FileName = PropBag.ReadProperty("FileName", "")
    Set rtb.Font = PropBag.ReadProperty("Font", Ambient.Font)
    rtb.HideSelection = PropBag.ReadProperty("HideSelection", True)
    m_hWnd = PropBag.ReadProperty("hWnd", m_def_hWnd)
    rtb.Locked = PropBag.ReadProperty("Locked", False)
    rtb.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    rtb.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    rtb.RightMargin = PropBag.ReadProperty("RightMargin", 0)
    rtb.Text = PropBag.ReadProperty("Text", "")

    mbInChange = False
    HighlightRefresh

    rtb.ToolTipText = PropBag.ReadProperty("ToolTip", "")
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("AutoVerbMenu", rtb.AutoVerbMenu, False)
    Call PropBag.WriteProperty("BackColor", rtb.Backcolor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("BorderStyle", rtb.BorderStyle, 1)
    Call PropBag.WriteProperty("BulletIndent", rtb.BulletIndent, 0)
    Call PropBag.WriteProperty("Enabled", rtb.Enabled, True)
    Call PropBag.WriteProperty("FileName", rtb.FileName, "")
    Call PropBag.WriteProperty("Font", rtb.Font, Ambient.Font)
    Call PropBag.WriteProperty("HideSelection", rtb.HideSelection, True)
    Call PropBag.WriteProperty("hWnd", m_hWnd, m_def_hWnd)
    Call PropBag.WriteProperty("Locked", rtb.Locked, False)
    Call PropBag.WriteProperty("MaxLength", rtb.MaxLength, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", rtb.MousePointer, 0)
    Call PropBag.WriteProperty("RightMargin", rtb.RightMargin, 0)
    Call PropBag.WriteProperty("Text", rtb.Text, "")
    Call PropBag.WriteProperty("ToolTip", rtb.ToolTipText, "")
End Sub

' *****************************************************************************
' Run Time Only Properties
' NOT generated by the ActiveX Control Wizard. Each of these procedures has
' its Procedure Attribute "Don't Show In Property Browser" set to true.
' *****************************************************************************

Public Property Get SelAlignment() As SelAlignmentConstants
Attribute SelAlignment.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If

    SelAlignment = rtb.SelAlignment
End Property

Public Property Let SelAlignment(ByVal New_SelAlignment As SelAlignmentConstants)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If

    rtb.SelAlignment = New_SelAlignment
End Property

Public Property Get SelBold() As Boolean
Attribute SelBold.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If

    SelBold = rtb.SelBold
End Property

Public Property Let SelBold(ByVal New_SelBold As Boolean)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If

    rtb.SelBold = New_SelBold
End Property

Public Property Get SelItalic() As Boolean
Attribute SelItalic.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If

    SelItalic = rtb.SelItalic
End Property

Public Property Let SelItalic(ByVal New_SelItalic As Boolean)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If

    rtb.SelItalic = New_SelItalic
End Property

Public Property Get SelStrikethru() As Boolean
Attribute SelStrikethru.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If

    SelStrikethru = rtb.SelStrikethru
End Property

Public Property Let SelStrikethru(ByVal New_SelStrikethru As Boolean)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If

    rtb.SelStrikethru = New_SelStrikethru
End Property

Public Property Get SelUnderline() As Boolean
Attribute SelUnderline.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If

    SelUnderline = rtb.SelUnderline
End Property

Public Property Let SelUnderline(ByVal New_SelUnderline As Boolean)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If

    rtb.SelUnderline = New_SelUnderline
End Property

Public Property Get SelBullet() As Variant
Attribute SelBullet.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If

    SelBullet = rtb.SelBullet
End Property

Public Property Let SelBullet(ByVal New_SelBullet As Variant)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If

    rtb.SelBullet = New_SelBullet
End Property

Public Property Get SelCharOffset() As Variant
Attribute SelCharOffset.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If

    SelCharOffset = rtb.SelCharOffset
End Property

Public Property Let SelCharOffset(ByVal New_SelCharOffset As Variant)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If

    rtb.SelCharOffset = New_SelCharOffset
End Property

Public Property Get SelRTF() As String
Attribute SelRTF.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If

    SelRTF = rtb.SelRTF
End Property

Public Property Let SelRTF(ByVal New_SelRTF As String)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If

    rtb.SelRTF = New_SelRTF
End Property

Public Property Get SelTabCount() As Integer
Attribute SelTabCount.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If

    SelTabCount = rtb.SelTabCount
End Property

Public Property Let SelTabCount(ByVal New_SelTabCount As Integer)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If

    rtb.SelTabCount = New_SelTabCount
End Property

Public Property Get SelTabs(Index As Integer) As Integer
Attribute SelTabs.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If

    SelTabs = rtb.SelTabs(Index)
End Property

Public Property Let SelTabs(Index As Integer, ByVal New_SelTabs As Integer)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If

    rtb.SelTabs(Index) = New_SelTabs
End Property

Public Property Get SelColor() As Variant
Attribute SelColor.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If

    SelColor = rtb.SelColor
End Property

Public Property Let SelColor(ByVal New_SelColor As Variant)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If

    rtb.SelColor = New_SelColor
End Property

Public Property Get SelHangingIndent() As Integer
Attribute SelHangingIndent.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If

    SelHangingIndent = rtb.SelHangingIndent
End Property

Public Property Let SelHangingIndent(ByVal New_SelHangingIndent As Integer)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If

    rtb.SelHangingIndent = New_SelHangingIndent
End Property

Public Property Get SelIndent() As Integer
Attribute SelIndent.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If

    SelIndent = rtb.SelIndent
End Property

Public Property Let SelIndent(ByVal New_SelIndent As Integer)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If

    rtb.SelIndent = New_SelIndent
End Property

Public Property Get SelRightIndent() As Integer
Attribute SelRightIndent.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If

    SelRightIndent = rtb.SelRightIndent
End Property

Public Property Let SelRightIndent(ByVal New_SelRightIndent As Integer)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If

    rtb.SelRightIndent = New_SelRightIndent
End Property

Public Property Get SelLength() As Long
Attribute SelLength.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If

    SelLength = rtb.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If

    rtb.SelLength = New_SelLength
End Property

Public Property Get SelStart() As Long
Attribute SelStart.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If

    SelStart = rtb.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If

    rtb.SelStart = New_SelStart
End Property

Public Property Get SelText() As String
Attribute SelText.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If

    SelText = rtb.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If

    rtb.SelText = New_SelText
End Property

Public Property Get SelProtected() As Variant
Attribute SelProtected.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If

    SelProtected = rtb.SelProtected
End Property

Public Property Let SelProtected(ByVal New_SelProtected As Variant)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If

    rtb.SelProtected = New_SelProtected
End Property

Public Property Get TextRTF() As String
Attribute TextRTF.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    TextRTF = rtb.TextRTF
End Property

Public Property Let TextRTF(ByVal New_TextRTF As String)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If

    mbInChange = True
    rtb.TextRTF = New_TextRTF
    mbInChange = False

    HighlightRefresh

End Property

Function AddZeroes(Amount)
    If Amount = 0 Then Exit Function
    Static v As Integer
    Static temp As String
    temp = ""

    For v = 1 To Amount
        temp = temp & "0"
    Next

    AddZeroes = temp
End Function

Function CountLines(Optional ToCursor As Boolean)
    Dim txt As String

    If ToCursor = True Then
        txt = Mid(rtb.Text, 1, (rtb.SelStart + rtb.SelLength))
    Else
        txt = rtb.Text
    End If

    Static lines, z As Integer
    lines = 1
    For z = 1 To Len(txt)
        If Asc(Mid(rtb.Text, z, 1)) = 13 Then
            lines = lines + 1
        End If
        DoEvents
    Next z

    CountLines = lines
End Function

Sub ChangeEvent()
    rtb_Change
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,ToolTipText
Public Property Get ToolTip() As String
Attribute ToolTip.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTip = rtb.ToolTipText
End Property

Public Property Let ToolTip(ByVal New_ToolTip As String)
    rtb.ToolTipText() = New_ToolTip
    PropertyChanged "ToolTip"
End Property

