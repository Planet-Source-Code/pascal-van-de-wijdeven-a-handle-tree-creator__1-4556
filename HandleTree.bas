Attribute VB_Name = "Module1"
Option Explicit

Public Type RECT
  Left    As Long
  Top     As Long
  Right   As Long
  Bottom  As Long
End Type

Public Type POINTAPI
  x       As Long
  Y       As Long
End Type

Public Type WINDOWPLACEMENT
  Length            As Long
  flags             As Long
  showCmd           As Long
  ptMinPosition     As POINTAPI
  ptMaxPosition     As POINTAPI
  rcNormalPosition  As RECT
End Type

Public phnd As Long

Public kk As RECT
Public pk As RECT

Public Const LB_SETTABSTOPS = &H192
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_TRANSPARENT = &H20&
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWME = SWP_FRAMECHANGED Or _
SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_NOTOPMOST = -2
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)

Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5

Public Const WS_OVERLAPPED = &H0&
Public Const WS_POPUP = &H80000000
Public Const WS_CHILD = &H40000000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_DISABLED = &H8000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Public Const WS_BORDER = &H800000
Public Const WS_DLGFRAME = &H400000
Public Const WS_VSCROLL = &H200000
Public Const WS_HSCROLL = &H100000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_GROUP = &H20000
Public Const WS_TABSTOP = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_CHILDWINDOW = (WS_CHILD)
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_ACCEPTFILES = &H10&

Public Const WH_MIN = (-1)
Public Const WH_MSGFILTER = (-1)
Public Const WH_JOURNALRECORD = 0
Public Const WH_JOURNALPLAYBACK = 1
Public Const WH_KEYBOARD = 2
Public Const WH_GETMESSAGE = 3
Public Const WH_CALLWNDPROC = 4
Public Const WH_CBT = 5
Public Const WH_SYSMSGFILTER = 6
Public Const WH_MOUSE = 7
Public Const WH_HARDWARE = 8
Public Const WH_DEBUG = 9
Public Const WH_SHELL = 10
Public Const WH_FOREGROUNDIDLE = 11
Public Const WH_MAX = 11
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
   (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
   
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Declare Function GetWindow Lib "user32" _
   (ByVal hWnd As Long, ByVal wCmd As Long) As Long

Public Declare Function GetWindowPlacement Lib "user32" _
   (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

Public Declare Function GetWindowRect Lib "user32" _
   (ByVal hWnd As Long, lpRect As RECT) As Long

Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
   (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
   (ByVal hWnd As Long, ByVal lpClassName As String, _
    ByVal nMaxCount As Long) As Long

Public Declare Function MoveWindow Lib "user32" _
   (ByVal hWnd As Long, _
    ByVal x As Long, ByVal Y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal bRepaint As Long) As Long

Public Declare Function SendMessageArray Lib "user32" Alias "SendMessageA" _
   (ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Public Declare Function SetForegroundWindow Lib "user32" _
   (ByVal hWnd As Long) As Long

Public Declare Function SetWindowPlacement Lib "user32" _
   (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

Public Declare Function ShowWindow Lib "user32" _
   (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
Declare Function GetLastError Lib "kernel32" () As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public Const TH32CS_SNAPPROCESS As Long = 2&
Public Const MAX_PATH As Long = 260

Public Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwflags As Long
szexeFile As String * MAX_PATH
End Type

Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlgas As Long, ByVal lProcessID As Long) As Long

Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long

Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long

Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Public Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Integer, ByVal Y As Integer, ByVal hIcon As Long) As Long
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long 'hinst- The instance handle of the application calling ExtractIcon.  Should be the name of your EXE file, or VB.EXE at runtime
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long 'lpModuleName- The filename of a module, to get the handle of it.

Public Function GetExeFromHandle(hWnd As Long) As String
Dim threadID As Long, processID As Long, hSnapshot As Long
Dim uProcess As PROCESSENTRY32, rProcessFound As Long
Dim i As Integer, szExename As String
' Get ID for window thread
threadID = GetWindowThreadProcessId(hWnd, processID)
' Check if valid
If threadID = 0 Or processID = 0 Then Exit Function
' Create snapshot of current processes
hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
' Check if snapshot is valid
If hSnapshot = -1 Then Exit Function
'Initialize uProcess with correct size
uProcess.dwSize = Len(uProcess)
'Start looping through processes
rProcessFound = ProcessFirst(hSnapshot, uProcess)
Do While rProcessFound
If uProcess.th32ProcessID = processID Then
'Found it, now get name of exefile
i = InStr(1, uProcess.szexeFile, Chr(0))
If i > 0 Then szExename = Left$(uProcess.szexeFile, i - 1)
Exit Do
Else
'Wrong ID, so continue looping
rProcessFound = ProcessNext(hSnapshot, uProcess)
End If
Loop
Call CloseHandle(hSnapshot)
GetExeFromHandle = szExename
End Function
