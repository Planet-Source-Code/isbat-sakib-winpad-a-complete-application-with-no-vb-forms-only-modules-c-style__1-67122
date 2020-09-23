Attribute VB_Name = "modDeclares"
'******************************************************************************
'******This code is downloaded from Planet Source Code
'****** Author: Isbat Sakib
'****** Email: iamsakib@gmail.com
'******************************************************************************


Option Explicit

'Class Styles
Public Const CS_VREDRAW             As Long = &H1
Public Const CS_HREDRAW             As Long = &H2

'Used for default processing
Public Const CW_USEDEFAULT          As Long = &H80000000

'Window Styles
Public Const WS_CHILD               As Long = &H40000000
Public Const WS_VISIBLE             As Long = &H10000000
Public Const WS_OVERLAPPED          As Long = &H0
Public Const WS_CAPTION             As Long = &HC00000  ' WS_BORDER Or WS_DLGFRAME
Public Const WS_SYSMENU             As Long = &H80000
Public Const WS_THICKFRAME          As Long = &H40000
Public Const WS_MINIMIZEBOX         As Long = &H20000
Public Const WS_MAXIMIZEBOX         As Long = &H10000
Public Const WS_OVERLAPPEDWINDOW    As Long = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_VSCROLL             As Long = &H200000
Public Const WS_HSCROLL             As Long = &H100000
Public Const WS_EX_CLIENTEDGE       As Long = &H200
Public Const WS_EX_ACCEPTFILES      As Long = &H10&

'Edit Control Styles
Public Const ES_MULTILINE           As Long = &H4
Public Const ES_AUTOHSCROLL         As Long = &H80
Public Const ES_AUTOVSCROLL         As Long = &H40
Public Const ES_LEFT                As Long = &H0
Public Const ES_NOHIDESEL           As Long = &H100&
Public Const ES_WANTRETURN          As Long = &H1000&

'Various Messages

Public Const WM_DESTROY             As Long = &H2
Public Const WM_LBUTTONDOWN         As Long = &H201
Public Const WM_LBUTTONUP           As Long = &H202
Public Const WM_CREATE              As Long = &H1
Public Const WM_CLOSE               As Long = &H10
Public Const WM_LBUTTONDBLCLK       As Long = &H203
Public Const WM_RBUTTONDOWN         As Long = &H204
Public Const WM_RBUTTONUP           As Long = &H205
Public Const WM_NCLBUTTONDOWN       As Long = &HA1
Public Const WM_COMMAND             As Long = &H111
Public Const WM_PAINT               As Long = &HF
Public Const WM_SETFONT             As Long = &H30
Public Const WM_GETFONT             As Long = &H31
Public Const WM_SETFOCUS            As Long = &H7
Public Const WM_SIZE                As Long = &H5
Public Const WM_USER                As Long = &H400
Public Const WM_QUIT                As Long = &H12
Public Const WM_UNDO                As Long = &H304
Public Const WM_INITMENUPOPUP       As Long = &H117
Public Const WM_PASTE               As Long = &H302
Public Const WM_CUT                 As Long = &H300
Public Const WM_COPY                As Long = &H301
Public Const WM_CLEAR               As Long = &H303
Public Const WM_GETTEXTLENGTH       As Long = &HE
Public Const WM_GETTEXT             As Long = &HD
Public Const WM_SETTEXT             As Long = &HC
Public Const WM_SETREDRAW           As Long = &HB
Public Const WM_DROPFILES           As Long = &H233
Public Const WM_CTLCOLOREDIT        As Long = &H133


Public Const EN_CHANGE              As Long = &H300
Public Const EN_UPDATE              As Long = &H400
Public Const EN_ERRSPACE            As Long = &H500
Public Const EN_MAXTEXT             As Long = &H501
Public Const EM_CANUNDO             As Long = &HC6
Public Const EM_SETMARGINS          As Long = &HD3
Public Const EC_RIGHTMARGIN         As Long = &H2
Public Const EM_GETSEL              As Long = &HB0
Public Const EM_SETSEL              As Long = &HB1
Public Const EM_SCROLLCARET         As Long = &HB7
Public Const EM_REPLACESEL          As Long = &HC2
Public Const EM_LINEFROMCHAR        As Long = &HC9
Public Const EM_LINEINDEX           As Long = &HBB
Public Const EM_LINELENGTH          As Long = &HC1
Public Const EM_LIMITTEXT           As Long = &HC5


'Clipboard Constants
Public Const CF_TEXT                As Long = 1

'Color Type
Public Const COLOR_WINDOW           As Long = 5

'Standard Cursor ID
Public Const IDC_ARROW              As Long = 32512

'Standard Icon ID
Public Const IDI_APPLICATION        As Long = 32512

'ShowWindow() Command
Public Const SW_SHOWNORMAL          As Long = &H1

'MessageBox() Flags
Public Const MB_OK                  As Long = &H0
Public Const MB_ICONEXCLAMATION     As Long = &H30
Public Const MB_ABORTRETRYIGNORE    As Long = &H2
Public Const MB_DEFBUTTON1          As Long = &H0
Public Const MB_DEFBUTTON2          As Long = &H100
Public Const MB_DEFBUTTON3          As Long = &H200
Public Const MB_ICONASTERISK        As Long = &H40
Public Const MB_ICONHAND            As Long = &H10
Public Const MB_ICONINFORMATION     As Long = MB_ICONASTERISK
Public Const MB_ICONQUESTION        As Long = &H20
Public Const MB_ICONSTOP            As Long = MB_ICONHAND
Public Const MB_OKCANCEL            As Long = &H1
Public Const MB_RETRYCANCEL         As Long = &H5
Public Const MB_YESNO               As Long = &H4
Public Const MB_YESNOCANCEL         As Long = &H3
Public Const MB_YES                 As Long = &H6
Public Const MB_NO                  As Long = &H7
Public Const IDCANCEL               As Long = 2
Public Const IDYES                  As Long = 6

'Window Field Offsets
Public Const GWL_HINSTANCE          As Long = (-6)
Public Const GWL_EXSTYLE            As Long = (-20)
Public Const GWL_STYLE              As Long = (-16)

'Menu Flags
Public Const MF_STRING              As Long = &H0&
Public Const MF_POPUP               As Long = &H10&
Public Const MF_SEPARATOR           As Long = &H800&
Public Const MF_BYCOMMAND           As Long = &H0
Public Const MF_GRAYED              As Long = &H1
Public Const MF_ENABLED             As Long = &H0
Public Const MF_CHECKED             As Long = &H8&
Public Const MF_UNCHECKED           As Long = &H0&

'Used in creating font
Public Const LF_FACESIZE            As Long = 32
Public Const FW_NORMAL              As Long = 400
Public Const ANSI_CHARSET           As Long = 0
Public Const OUT_DEFAULT_PRECIS     As Long = 0
Public Const CLIP_DEFAULT_PRECIS    As Long = 0
Public Const DEFAULT_QUALITY        As Long = 0

'Used in RedrawWindow API call
Public Const RDW_ERASE              As Long = &H4
Public Const RDW_INVALIDATE         As Long = &H1
Public Const RDW_ERASENOW           As Long = &H200

'Locale Information
Public Const LOCALE_USER_DEFAULT    As Long = &H400
Public Const TIME_NOSECONDS         As Long = &H2
Public Const DATE_SHORTDATE         As Long = &H1

'File Constants
Public Const GENERIC_READ           As Long = &H80000000
Public Const GENERIC_WRITE          As Long = &H40000000
Public Const CREATE_NEW             As Long = 1
Public Const OPEN_EXISTING          As Long = 3
Public Const TRUNCATE_EXISTING      As Long = 5
Public Const FILE_ATTRIBUTE_NORMAL  As Long = &H80
Public Const FILE_ATTRIBUTE_READONLY As Long = &H1
Public Const INVALID_HANDLE_VALUE   As Long = -1
Public Const FILE_SHARE_READ        As Long = &H1
Public Const FILE_SHARE_WRITE       As Long = &H2

Public Const TRANSPARENT            As Long = 1
Public Const OPAQUE                 As Long = 2

Public Const FNOINVERT              As Byte = &H2
Public Const FVIRTKEY               As Byte = 1
Public Const FCONTROL               As Byte = &H8
Public Const VK_F1                  As Integer = &H70
Public Const VK_F2                  As Integer = &H71
Public Const VK_F3                  As Integer = &H72
Public Const VK_F5                  As Integer = &H74


'Used for registry access
Public Const REG_SZ                 As Long = 1
Public Const REG_DWORD              As Long = 4
Public Const REG_OPTION_NON_VOLATILE As Long = 0
Public Const HKEY_LOCAL_MACHINE     As Long = &H80000002
Public Const KEY_ALL_ACCESS         As Long = &H3F

Public Enum RegistryDataType
    RegistryString = REG_SZ
    RegistryLong = REG_DWORD
End Enum

'Default Font
Public Const SYSTEM_FONT            As Long = 13

'Application Defined IDs
Public Const ID_FILE_NEW            As Long = 9001
Public Const ID_FILE_OPEN           As Long = 9002
Public Const ID_FILE_SAVE           As Long = 9003
Public Const ID_FILE_SAVEAS         As Long = 9004
Public Const ID_FILE_PAGESETUP      As Long = 9005
Public Const ID_FILE_PRINT          As Long = 9006
Public Const ID_FILE_EXIT           As Long = 9007
Public Const ID_EDIT_UNDO           As Long = 9008
Public Const ID_EDIT_CUT            As Long = 9009
Public Const ID_EDIT_COPY           As Long = 9010
Public Const ID_EDIT_PASTE          As Long = 9011
Public Const ID_EDIT_DELETE         As Long = 9012
Public Const ID_EDIT_SELECTALL      As Long = 9013
Public Const ID_EDIT_TIMEDATE       As Long = 9014
Public Const ID_EDIT_WORDWRAP       As Long = 9015
Public Const ID_EDIT_SETFONT        As Long = 9016
Public Const ID_SEARCH_FIND         As Long = 9017
Public Const ID_SEARCH_FINDNEXT     As Long = 9018
Public Const ID_HELP_HELPTOPICS     As Long = 9019
Public Const ID_HELP_ABOUT          As Long = 9020
Public Const IDC_MAIN_EDIT          As Long = 101


Public Type WNDCLASS
    style                           As Long
    lpfnwndproc                     As Long
    cbClsextra                      As Long
    cbWndExtra2                     As Long
    hInstance                       As Long
    hIcon                           As Long
    hCursor                         As Long
    hbrBackground                   As Long
    lpszMenuName                    As String
    lpszClassName                   As String
End Type

Public Type POINTAPI
    x                               As Long
    y                               As Long
End Type

Public Type RECT
    Left                            As Long
    Top                             As Long
    Right                           As Long
    Bottom                          As Long
End Type

Public Type msg
    hwnd                            As Long
    Message                         As Long
    wParam                          As Long
    lParam                          As Long
    time                            As Long
    pt                              As POINTAPI
End Type

Public Type SYSTEMTIME
    wYear                           As Integer
    wMonth                          As Integer
    wDayOfWeek                      As Integer
    wDay                            As Integer
    wHour                           As Integer
    wMinute                         As Integer
    wSecond                         As Integer
    wMilliseconds                   As Integer
End Type

Public Type LOGFONT
    lfHeight                        As Long
    lfWidth                         As Long
    lfEscapement                    As Long
    lfOrientation                   As Long
    lfWeight                        As Long
    lfItalic                        As Byte
    lfUnderline                     As Byte
    lfStrikeOut                     As Byte
    lfCharSet                       As Byte
    lfOutPrecision                  As Byte
    lfClipPrecision                 As Byte
    lfQuality                       As Byte
    lfPitchAndFamily                As Byte
    lfFaceName(0 To LF_FACESIZE - 1) As Byte
End Type

Public Type ACCEL
    fVirt                           As Byte
    Key                             As Integer
    cmd                             As Integer
End Type


'Global Class Name
Public Const gClassName             As String = "WinPadClass"

'Global Application Name
Public Const gAppName               As String = "WinPad"

'Global hwnd of the Application
Public gHwnd                        As Long

'Global hwnd of the Edit Box
Public gEditHwnd                    As Long

'Global handle of menubar
Public ghMenu                       As Long

'Global handle of the font used by edit box
Public ghFont                       As Long

'Global variable to check if edit box contents are changed
Public gEditChanged                 As Boolean

'Global FilePath
Public gPathOfFile                  As String

'Global FileName
Public gNameOfFile                  As String

'Global Text Color
Public gTextColor                   As Long

'Global Background brush for edit box
Public gBkBrush                     As Long

'Global Accelerator Table
Public gAccTable                    As Long

'Declares and Subs
Public Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Public Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Public Declare Function TranslateMessage Lib "user32" (lpMsg As msg) As Long
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As msg) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function LoadCursorByNum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function LoadIconByNum Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Long) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Public Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Public Declare Function CreateMenu Lib "user32" () As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function AppendMenuByString Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Public Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Public Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Public Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Public Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SYSTEMTIME, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long
Public Declare Function GetTimeFormat Lib "kernel32" Alias "GetTimeFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpTime As SYSTEMTIME, ByVal lpFormat As String, ByVal lpTimeStr As String, ByVal cchTime As Long) As Long
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Public Declare Function IsDialogMessage Lib "user32" Alias "IsDialogMessageA" (ByVal hDlg As Long, lpMsg As msg) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function CheckMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
Public Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegSetValueExByString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExByLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Long, ByVal cbData As Long) As Long
Public Declare Function RegQueryValueExByString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegQueryValueExByLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Public Declare Sub DragFinish Lib "shell32.dll" (ByVal hDrop As Long)
Public Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function DestroyAcceleratorTable Lib "user32" (ByVal haccel As Long) As Long
Public Declare Function CreateAcceleratorTable Lib "user32" Alias "CreateAcceleratorTableA" (lpaccl As ACCEL, ByVal cEntries As Long) As Long
Public Declare Function TranslateAccelerator Lib "user32" Alias "TranslateAcceleratorA" (ByVal hwnd As Long, ByVal hAccTable As Long, lpMsg As msg) As Long
Public Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineA" () As String
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetMapMode Lib "gdi32.dll" (ByVal hdc As Long) As Long
