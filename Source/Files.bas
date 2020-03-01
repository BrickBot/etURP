Attribute VB_Name = "Files"
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
' Return codes from Registration functions.
Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1&
Const ERROR_BADKEY = 2&
Const ERROR_CANTOPEN = 3&
Const ERROR_CANTREAD = 4&
Const ERROR_CANTWRITE = 5&
Const ERROR_OUTOFMEMORY = 6&
Const ERROR_INVALID_PARAMETER = 7&
Const ERROR_ACCESS_DENIED = 8&
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const MAX_PATH = 260&
Private Const REG_SZ = 1

Function AssociateFile(ProgramName As String, ProgramDesc As String, FileExt As String)
Debug.Print "ASSOCIATE FILE - " & ProgramName & " called."
    
    Dim sKeyName As String   'Holds Key Name in registry.
    Dim sKeyValue As String  'Holds Key Value in registry.
    Dim ret&                 'Holds error status if any from API calls.
    Dim lphKey&              'Holds created key handle from RegCreateKey.

    'This creates a Root entry called "MyApp".
    sKeyName = ProgramName
    sKeyValue = ProgramDesc
    ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
    ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)
    'This creates a Root entry called .BAR associated with "MyApp".
    sKeyName = FileExt
    sKeyValue = ProgramName
    ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
    ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)
    'This sets the command line for "MyApp".
    sKeyName = ProgramName
    If Right(App.Path, 1) = "\" Then BetweenChar = "" Else BetweenChar = "\"
    sKeyValue = App.Path & BetweenChar & App.EXEName & " %1"
    ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
    ret& = RegSetValue&(lphKey&, "shell\open\command", REG_SZ, sKeyValue, MAX_PATH)
    sKeyValue = App.Path & BetweenChar & "Bin\etURP.ico"
    ret& = RegSetValue&(lphKey&, "DefaultIcon", REG_SZ, sKeyValue, MAX_PATH)
End Function
