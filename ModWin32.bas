Attribute VB_Name = "ModWin32"
' License: This program is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License as published by
' the Free Software Foundation; either version 3 of the License, or (at your
' option) any later version. This program is distributed in the hope that it
' will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty
' of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General
' Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

'© 2008-2011 Thomas James
'tmj2005@gmail.com

Public Const VK_F1 = &H70
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B
Public Const VK_F13 = &H7C
Public Const VK_F14 = &H7D
Public Const VK_F15 = &H7E
Public Const VK_F16 = &H7F
Public Const VK_F17 = &H80
Public Const VK_F18 = &H81
Public Const VK_F19 = &H82
Public Const VK_F20 = &H83
Public Const VK_F21 = &H84
Public Const VK_F22 = &H85
Public Const VK_F23 = &H86
Public Const VK_F24 = &H87
Public Const VK_KEY_0 = &H30
Public Const VK_KEY_1 = &H31
Public Const VK_KEY_2 = &H32
Public Const VK_KEY_3 = &H33
Public Const VK_KEY_4 = &H34
Public Const VK_KEY_5 = &H35
Public Const VK_KEY_6 = &H36
Public Const VK_KEY_7 = &H37
Public Const VK_KEY_8 = &H38
Public Const VK_KEY_9 = &H39
Public Const VK_KEY_A = &H41
Public Const VK_KEY_B = &H42
Public Const VK_KEY_C = &H43
Public Const VK_KEY_D = &H44
Public Const VK_KEY_E = &H45
Public Const VK_KEY_F = &H46
Public Const VK_KEY_G = &H47
Public Const VK_KEY_H = &H48
Public Const VK_KEY_I = &H49
Public Const VK_KEY_J = &H4A
Public Const VK_KEY_K = &H4B
Public Const VK_KEY_L = &H4C
Public Const VK_KEY_M = &H4D
Public Const VK_KEY_N = &H4E
Public Const VK_KEY_O = &H4F
Public Const VK_KEY_P = &H50
Public Const VK_KEY_Q = &H51
Public Const VK_KEY_R = &H52
Public Const VK_KEY_S = &H53
Public Const VK_KEY_T = &H54
Public Const VK_KEY_U = &H55
Public Const VK_KEY_V = &H56
Public Const VK_KEY_W = &H57
Public Const VK_KEY_X = &H58
Public Const VK_KEY_Y = &H59
Public Const VK_KEY_Z = &H5A
Public Const VK_NUMPAD0 = &H60
Public Const VK_NUMPAD1 = &H61
Public Const VK_NUMPAD2 = &H62
Public Const VK_NUMPAD3 = &H63
Public Const VK_NUMPAD4 = &H64
Public Const VK_NUMPAD5 = &H65
Public Const VK_NUMPAD6 = &H66
Public Const VK_NUMPAD7 = &H67
Public Const VK_NUMPAD8 = &H68
Public Const VK_NUMPAD9 = &H69
Public Const VK_ADD = &H6B 'numpad
Public Const VK_DECIMAL = &H6E 'numpad
Public Const VK_DIVIDE = &H6F 'numpad
Public Const VK_MULTIPLY = &H6A 'numpad
Public Const VK_SUBTRACT = &H6D 'numpad
Public Const VK_SPACE = &H20 'Space
Public Const VK_INSERT = &H2D 'Insert
Public Const VK_HOME = &H24 'Home
Public Const VK_PRIOR = &H21 'Page Up
Public Const VK_DELETE = &H2E 'Delete
Public Const VK_END = &H23 'End
Public Const VK_NEXT = &H22 'Page Down
Public Const VK_UP = &H26 'Arrow Up
Public Const VK_DOWN = &H28 'Arrow Down
Public Const VK_LEFT = &H25 'Arrow Left
Public Const VK_RIGHT = &H27 'Arrow Right
Public Const VK_TAB = &H9  'Tab
Public Const VK_RETURN = &HD  'Enter
Public Const VK_ESCAPE = &H1B 'Esc
Public Const VK_SCROLL = &H91 'Scroll Lock
Public Const VK_NUMLOCK = &H90 'Num Lock
Public Const VK_PAUSE = &H13 'Pause
Public Const VK_SNAPSHOT = &H2C 'Print Screen
Public Const VK_LWIN = &H5B 'Left Win
Public Const VK_RWIN = &H5C 'Right Win
Public Const VK_APPS = &H5D 'Context Menu
Public Const VK_BACK = &H8 ' Backspace
Public Const VK_OEM_8 = &HDF

Public Const VK_CONTROL = &H11
Public Const VK_LCONTROL = &HA2
Public Const VK_RCONTROL = &HA3
Public Const VK_MENU = &H12
Public Const VK_LMENU = &HA4
Public Const VK_RMENU = &HA5
Public Const VK_SHIFT = &H10
Public Const VK_LSHIFT = &HA0
Public Const VK_RSHIFT = &HA1
Public Const VK_CAPITAL = &H14

Public Const WM_CHAR = &H102
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105

Public Const MB_ICONHAND = &H10& ' = (vbCritical)
Public Const MB_ICONQUESTION = &H20& ' = (vbQuestion)
Public Const MB_ICONEXCLAMATION = &H30& ' = (vbExclamation)
Public Const MB_ICONASTERISK = &H40& ' = (vbInformation)
Public Const MB_OK = &H0&

Public Const KEYEVENTF_KEYDOWN = &H0
Public Const KEYEVENTF_KEYUP = &H2
Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_UNICODE = &H4
Public Const INPUT_KEYBOARD = &H1

Public Const COLOR_WINDOW = &H5
Public Const COLOR_BTNFACE = &HF

Public Const E_INVALIDARG = &H80070057
Public Const CSIDL_LOCAL_APPDATA = &H1C
Public Const SW_SHOWDEFAULT = &HA

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_CLOSE = &H10
Public Const WM_PASTE = &H302

Public Const IDC_HAND = 32649&
Public Const IDC_ARROW = 32512&
Public Const IDC_WAIT = 32514&

Public Const AddIcon = &H0
Public Const ModifyIcon = &H1
Public Const DeleteIcon = &H2
Public Const MessageFlag = &H1
Public Const IconFlag = &H2
Public Const TipFlag = &H4

Public Const WH_KEYBOARD_LL = 13&
Public Const HC_ACTION = 0&

Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Public Declare Function SetKeyboardState Lib "user32" (kbArray As KeyboardBytes) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Public Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Public Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Public Declare Function MapVirtualKeyEx Lib "user32" Alias "MapVirtualKeyExA" (ByVal wCode As Long, ByVal wMapType As Long, ByVal dwhkl As Long) As Long
Public Declare Function ToAsciiEx Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpKeyState As Byte, lpChar As Long, ByVal uFlags As Long, ByVal dwhkl As Long) As Long
Public Declare Function SendInput Lib "user32" (ByVal nInputs As Long, pInputs As Any, ByVal cbSize As Long) As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Public Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal Message As Long, data As NotifyIconData) As Boolean
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Public Declare Function SysAllocStringLen Lib "oleaut32" (ByVal Ptr As Long, ByVal Length As Long) As Long
Public Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, lpUsedDefaultChar As Long) As Long
Public Declare Sub PutMem4 Lib "msvbvm60" (ByVal Ptr As Long, ByVal Value As Long)

Public Type NotifyIconData
  Size As Long
  handle As Long
  id As Long
  Flags As Long
  CallBackMessage As Long
  Icon As Long
  Tip As String * 64
End Type

Public Type KBDLLHOOKSTRUCT
    vkCode As Long
    ScanCode As Long
    Flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Public Type KeyboardBytes
    kbByte(0 To 255) As Byte
End Type

Public Type KeyboardInput
   dwType As Long
   wVK As Integer
   wScan As Integer
   dwFlags As Long
   dwTime As Long
   dwExtraInfo As Long
   dwPadding As Currency
End Type
