VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpiritDLL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'                            -----------------------
'                            EASY LEGO DLL INTERFACE
'                            (C) Dean Camera, 2003
'                            -----------------------

' --------------------------------------------------------------------------------
' --------------------------------------------------------------------------------
' --------------------------------------------------------------------------------

Private WithEvents DLLSpirit As LEGOVPBrickLib.VPBrick
Attribute DLLSpirit.VB_VarHelpID = -1

Public Enum PortNum
AUTO = 0
COM1 = 1
COM2 = 2
COM3 = 3
COM4 = 4
COM5 = 5
COM6 = 6
COM7 = 7
COM8 = 8
USB1 = 9
USB2 = 10
End Enum

Public Enum LPBrickType
Brick_RCX2 = 0
Brick_Scout = 1
Brick_SpyBot = 2
Brick_MicroScout = 3
End Enum

Dim Brick As LEGOVPBrickLib.BrickTypes

Dim LastErrorInt As Integer
Dim LastErrorDescription As String

Private Sub Class_Initialize()
Set DLLSpirit = New LEGOVPBrickLib.VPBrick
End Sub

' --------------------------------------------------------------------------------
' --------------------------------------------------------------------------------
' --------------------------------------------------------------------------------
' --------------------------------------------------------------------------------

' FUNCTION NAME: GetLastErrorDescription
' DESCRIPTION: Gets the description of the last DLL error.
' RETURNS: Last error description or "No Error or Wrong Last Number!"

Function GetLastErrorDescription(ErrorNumber As Integer) As Integer
If ErrorNumber = LastErrorInt Then
GetLastErrorDescription = LastErrorDescription
Else
GetLastErrorDescription = "No Error or Wrong Last Number!"
End If
End Function

' FUNCTION NAME: InitComm
' DESCRIPTION: Initialises the port. Must be called before other commands.
' RETURNS: 1 if successful, error number if not.

Function InitCom(PBrickType As LPBrickType, BrickPort As PortNum) As Integer
On Error GoTo Damn

Select Case PBrickType
Case 1
Brick = Scout
Case 2
Brick = Spybot
Case 3
Brick = MicroScout
Case Else
Brick = RCX2
End Select

DLLSpirit.BrickType = Brick

Select Case BrickPort
Case 0                              'Auto
UsePort$ = vbNullString
DLLSpirit.FindPort UsePort$
DLLSpirit.Open UsePort$
Case Is > 8                         'USB
DLLSpirit.Open "\\.\USB" & Int(BrickPort)
Case Else                           'COM
DLLSpirit.Open "\\.\COM" & Int(BrickPort)
End Select
InitCom = 1
Exit Function
Damn:
InitCom = Err.Number
LastErrorInt = Err.Number
LastErrorDescription = Err.Description
End Function

' FUNCTION NAME: CheckStatus
' DESCRIPTION: Check the satus of the brick, to see if it is able to recieve a
'              command.
' RETURNS: 1 if successful, error number if not.

Function CheckStatus(PlayReadySound As Boolean) As Integer
On Error GoTo Damn
  nBrickType = DLLSpirit.Status(CheckBrickType)
  DLLSpirit.SetBrickType (nBrickType)
  nStatus = DLLSpirit.Status(BrickStatus)
  If nStatus = StatusReady Then
    If nBrickType = Scout Then
    If PlayReadySound = True Then DLLSpirit.Execute "sound 25"
    Else
    If PlayReadySound = True Then DLLSpirit.Execute "sound 3"
    End If
    CheckStatus = 1
  Else
    CheckStatus = 0
  End If
  
Exit Function
Damn:
InitCom = Err.Number
LastErrorInt = Err.Number
LastErrorDescription = Err.Description
End Function

' FUNCTION NAME: PlayTone
' DESCRIPTION: Plays the selected tone for the selected duration.
' RETURNS: 1 if successful, error number if not.

Function PlayTone(Frequency As Integer, Duration As Integer) As Integer
On Error GoTo Damn
DLLSpirit.Execute "playt " & Frequency & "," & Duration
PlayTone = 1
Exit Function
Damn:
PlayTone = Err.Number
LastErrorInt = Err.Number
LastErrorDescription = Err.Description
End Function

' FUNCTION NAME: CloseComm
' DESCRIPTION: Closes communication with the Brick/Tower. Must be called last.
' RETURNS: 1 if successful, 0 if not.

Function CloseComm() As Integer
On Error GoTo Damn
DLLSpirit.Close
CloseComm = 1
Exit Function
Damn:
CloseComm = 0
End Function

' FUNCTION NAME: ExecuteCommand
' DESCRIPTION: Executes a LASM, MindSript or VLL command.
' RETURNS: 1 if successful, error number if not.

Function ExecuteCommand(Command$, TypeOfCommand As LEGOVPBrickLib.CodeType) As Integer
On Error GoTo Damn
DLLSpirit.Execute Command$, TypeOfCommand
ExecuteCommand = 1
Exit Function
Damn:
ExecuteCommand = Err.Number
LastErrorInt = Err.Number
LastErrorDescription = Err.Description
End Function

' FUNCTION NAME: DownloadProgram
' DESCRIPTION: Executes a LASM, MindSript or VLL program.
' RETURNS: 1 if successful, error number on DLL error, or "#" & errorpos.

Function DownloadProgram(Program As String, PCode As LEGOVPBrickLib.CodeType, ProgramSlot As Integer)
On Error GoTo Damn
DLLSpirit.ProgramSlot = ProgramSlot
DLLSpirit.Download Program, PCode, ErrorPos
If ErrorPos = 0 Then
DownloadProgram = 1
Else
DownloadProgram = "#" & ErrorPos
End If
Exit Function
Damn:
DownloadProgram = Err.Number
LastErrorInt = Err.Number
LastErrorDescription = Err.Description
End Function
