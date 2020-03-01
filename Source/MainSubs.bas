Attribute VB_Name = "MainSubs"
Global FixSTemplatesMenu, CheckTemplatesUnload  As Boolean
Global DownloadingFirmware As Boolean
Global AFAddText As String

Global CurrentLine As Integer

Global FirmwareMaxSecs, FirmwareByteSize As Integer
Global FIRMNAME, FIRMLOCATION As String

Global NOFIRM As Boolean

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public fMainForm As frmMain
Public Declare Function M_FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer
Global MyCom As String
Global COMOPEN As Boolean
Global FixTB, DisableTB As Boolean
Global NoAllowSRCX As Boolean

Global MyCommand As String

Public Const MINSIZE = 10000

Global BrickType As PBricks

Global LRTemp, SRTemp, mOCom, mCCom As Boolean
Global COMOPENCLOSED As Integer

Global XPLib As New XPInterface

Public Enum PBricks
    NONE = 0
    RCX = 1
    RCX2 = 2
    CYBERMASTER = 4
End Enum

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
         hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
         lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
         lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
         ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
         ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
         lpStartupInfo As STARTUPINFO, lpProcessInformation As _
         PROCESS_INFORMATION) As Long

Private Declare Function CloseHandle Lib "kernel32" _
         (ByVal hObject As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
         (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Const HIGH_PRIORITY_CLASS = &H80
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

Private Type STARTUPINFO

    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Public Function ExecCmd(cmdline As String, Optional NotInBinDir As Boolean) As Long
    If NotInBinDir = False Then cmdline = App.Path & "\bin\" & cmdline

    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim lngRC As Long
    ' Initialize the STARTUPINFO structure:
    start.cb = Len(start)
    ' Start the shelled application:
    lngRC = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, _
       HIGH_PRIORITY_CLASS, 0&, 0&, start, proc)
    ' Wait for the shelled application to finish:
    lngRC = WaitForSingleObject(proc.hProcess, INFINITE)
    Call GetExitCodeProcess(proc.hProcess, lngRC)
    Call CloseHandle(proc.hProcess)
    ExecCmd = lngRC
End Function

Function BrickTypeHeader()
    BrickTypeHeader = ""
    If BrickType = RCX Then BrickTypeHeader = "-TRCX"
    If BrickType = RCX2 Then BrickTypeHeader = "-TRCX2"
    If BrickType = CYBERMASTER Then BrickTypeHeader = "-TCM"
End Function

Function BCandLASMBrickTypeHeader()
    If BrickType = RCX Then BrickTypeHeader = "-T=RCX"
    If BrickType = RCX2 Then BrickTypeHeader = "-T=RCX2"
    If BrickType = RCX2 Then BrickTypeHeader = "-T=CM"
End Function

Sub EasterEgg()
    COMOPEN = True
End Sub

Sub LogText(text)
Debug.Print text
If Right(App.Path, 1) = "\" Then BetweenChar = "" Else BetweenChar = "\"
Open App.Path & BetweenChar & "etURP.log" For Append As #9
Print #9, text
Close #9
End Sub
