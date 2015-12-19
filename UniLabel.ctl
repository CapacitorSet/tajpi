VERSION 5.00
Begin VB.UserControl UniLabel 
   CanGetFocus     =   0   'False
   ClientHeight    =   15
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15
   ForwardFocus    =   -1  'True
   MaskColor       =   &H00FFFFFF&
   ScaleHeight     =   1
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1
   ToolboxBitmap   =   "UniLabel.ctx":0000
End
Attribute VB_Name = "UniLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Option Explicit

' Selfsub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' Local variables/constants: must declare these regardless if using subclassing, hooking, callbacks
    Private z_scFunk            As Collection   'hWnd/thunk-address collection; initialized as needed
    Private z_hkFunk            As Collection   'hook/thunk-address collection; initialized as needed
    Private z_cbFunk            As Collection   'callback/thunk-address collection; initialized as needed
    Private Const IDX_INDEX     As Long = 2     'index of the subclassed hWnd OR hook type
    Private Const IDX_PREVPROC  As Long = 9     'Thunk data index of the original WndProc
    Private Const IDX_BTABLE    As Long = 11    'Thunk data index of the Before table for messages
    Private Const IDX_ATABLE    As Long = 12    'Thunk data index of the After table for messages
    Private Const IDX_CALLBACKORDINAL As Long = 36 ' Ubound(callback thunkdata)+1, index of the callback

  ' Declarations:
    Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
    Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
    Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
    Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
    Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
    Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
    Private Enum eThunkType
        SubclassThunk = 0
        HookThunk = 1
        CallbackThunk = 2
    End Enum

    '-Selfsub specific declarations----------------------------------------------------------------------------
    Private Enum eMsgWhen                                                   'When to callback
      MSG_BEFORE = 1                                                        'Callback before the original WndProc
      MSG_AFTER = 2                                                         'Callback after the original WndProc
      MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                            'Callback before and after the original WndProc
    End Enum
    
    ' see ssc_Subclass for complete listing of indexes and what they relate to
    Private Const IDX_PARM_USER As Long = 13    'Thunk data index of the User-defined callback parameter data index
    Private Const IDX_UNICODE   As Long = 107   'Must be UBound(subclass thunkdata)+1; index for unicode support
    Private Const MSG_ENTRIES   As Long = 32    'Number of msg table entries. Set to 1 if using ALL_MESSAGES for all subclassed windows
    Private Const ALL_MESSAGES  As Long = -1    'All messages will callback
    
    Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function CallWindowProcW Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
    Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hWnd As Long) As Long
    Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    
    '-SelfHook specific declarations----------------------------------------------------------------------------
    Private Declare Function SetWindowsHookExA Lib "user32" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Function SetWindowsHookExW Lib "user32" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
    Private Declare Function CallNextHookEx Lib "user32.dll" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
    Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
    
    Private Enum eHookType  ' http://msdn2.microsoft.com/en-us/library/ms644990.aspx
      WH_MSGFILTER = -1
      WH_JOURNALRECORD = 0
      WH_JOURNALPLAYBACK = 1
      WH_KEYBOARD = 2
      WH_GETMESSAGE = 3
      WH_CALLWNDPROC = 4
      WH_CBT = 5
      WH_SYSMSGFILTER = 6
      WH_MOUSE = 7
      WH_DEBUG = 9
      WH_SHELL = 10
      WH_FOREGROUNDIDLE = 11
      WH_CALLWNDPROCRET = 12
      WH_KEYBOARD_LL = 13       ' NT/2000/XP+ only, Global hook only
      WH_MOUSE_LL = 14          ' NT/2000/XP+ only, Global hook only
    End Enum
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long

Private Declare Function GetTextAlign Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32" (ByVal hDC As Long, ByVal wFlags As Long) As Long

' BackStyle property
Public Enum ulbBackStyle
    ulbTransparent
    ulbOpaque
End Enum

' Mouse button constants
Public Enum ulbMouseButtonConstants
    ulbNoButton = 0
    ulbLeftButton = vbLeftButton
    ulbRightButton = vbRightButton
    ulbLeftAndRight = vbLeftButton Or vbRightButton
    ulbMiddleButton = vbMiddleButton
    ulbLeftAndMiddle = vbLeftButton Or vbMiddleButton
    ulbRightAndMiddle = vbRightButton Or vbMiddleButton
    ulbAllButtons = vbLeftButton Or vbRightButton Or vbMiddleButton
End Enum

' Shift constants
Public Enum ulbShiftConstants
    ulbNoneMask = 0
    ulbShiftMask = vbShiftMask
    ulbCtrlMask = vbCtrlMask
    ulbShiftCtrlMask = vbShiftMask Or vbCtrlMask
    ulbAltMask = vbAltMask
    ulbShiftAltMask = vbShiftMask Or vbAltMask
    ulbCtrlAltMask = vbCtrlMask Or vbAltMask
    ulbShiftCtrlAltMask = vbShiftMask Or vbCtrlMask Or vbAltMask
End Enum

' events    Tools > Procedure attributes...
Public Event Change()
Public Event Click(Button As MouseButtonConstants)
Public Event DblClick(Button As MouseButtonConstants)
Public Event MouseDown(Button As ulbMouseButtonConstants, Shift As ulbShiftConstants, X As Single, Y As Single)
Public Event MouseEnter()
Public Event MouseLeave()
Public Event MouseMove(Button As ulbMouseButtonConstants, Shift As ulbShiftConstants, X As Single, Y As Single)
Public Event MouseUp(Button As ulbMouseButtonConstants, Shift As ulbShiftConstants, X As Single, Y As Single)

' Window messages
Private Const WM_DESTROY = &H2&
Private Const WM_EXITSIZEMOVE As Long = &H232&
Private Const WM_LBUTTONDBLCLK = &H203&
Private Const WM_LBUTTONDOWN = &H201&
Private Const WM_LBUTTONUP = &H202&
Private Const WM_MBUTTONDBLCLK = &H209&
Private Const WM_MBUTTONDOWN = &H207&
Private Const WM_MBUTTONUP = &H208&
Private Const WM_MOUSEMOVE = &H200&
Private Const WM_MOUSELEAVE = &H2A3&
Private Const WM_PAINT = &HF&
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204&
Private Const WM_RBUTTONUP = &H205&
Private Const WM_SHOWWINDOW = &H18&
Private Const WM_STYLECHANGED = &H7D&
Private Const WM_SYSCOLORCHANGE = &H15&
Private Const WM_THEMECHANGED = &H31A&

' API drawtext constants
Private Const DT_CALCRECT = &H400&
Private Const DT_CENTER = &H1&
Private Const DT_LEFT = &H0&
Private Const DT_NOCLIP = &H100&
Private Const DT_NOPREFIX = &H800
Private Const DT_RIGHT = &H2&
Private Const DT_SINGLELINE = &H20
Private Const DT_WORDBREAK = &H10&

' API font constants
Private Const FW_BOLD = 700&
Private Const FW_NORMAL = 400&
Private Const LF_FACESIZE = 32&
Private Const LOGPIXELSX = 88&
Private Const LOGPIXELSY = 90&

' WM_MOUSEMOVE and others
Private Const MK_LBUTTON = &H1&
Private Const MK_RBUTTON = &H2&
Private Const MK_SHIFT = &H4&
Private Const MK_CONTROL = &H8&
Private Const MK_MBUTTON = &H10&

' API text alignment
Private Const TA_RTLREADING = &H100&

' API used by ExtTextOutW
Private Const ETO_CLIPPED As Integer = &H4

' API font structure
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(31) As Byte
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type PAINTSTRUCT
    hDC As Long
    fErase As Long
    rcPaint As RECT
    fRestore As Long
    fIncUpdate As Long
    rgbReserved(31) As Byte
End Type

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize As Long
    dwFlags As TRACKMOUSEEVENT_FLAGS
    hWndTrack As Long
    dwHoverTime As Long
End Type

' API declarations: generic
Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetUpdateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function TrackMouseEventUser32 Lib "user32" Alias "TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

' API declarations: ANSI/Unicode
Private Declare Function DrawTextA Lib "user32" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function ExtTextOutW Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As Long, ByVal nCount As Long, lpDx As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByVal lpVersionInformation As Long) As Long

' public properties                     Tools > Procedure attributes...
Dim m_Alignment As AlignmentConstants   ' Alignment
Dim m_AutoSize As Boolean               ' AutoSize
Dim m_BackColor As OLE_COLOR            ' BackColor
Dim m_BackStyle As ulbBackStyle         ' BackStyle
Dim m_Caption As String                 ' Caption
Dim m_Font As IFontDisp                 ' Font
Dim m_ForeColor As OLE_COLOR            ' Forecolor
Dim m_PaddingBottom As Byte             ' PaddingBottom
Dim m_PaddingLeft As Byte               ' PaddingLeft
Dim m_PaddingRight As Byte              ' PaddingRight
Dim m_PaddingTop As Byte                ' PaddingTop
Dim m_RightToLeft As Boolean            ' RightToLeft
Dim m_UseEvents As Boolean              ' UseEvents
Dim m_UseMnemonic As Boolean            ' UseMnemonic
Dim m_WordWrap As Boolean               ' WordWrap

' public functions
Dim m_MouseOver As Boolean              ' Indicates whether mouse is hovering over control

' private "properties"
Dim m_BackClr As Long                   ' Label background color as pure RGB
Dim m_Container As RECT                 ' Container full draw area
Dim m_ContainerhWnd As Long
Dim m_ContainerIsParent As Boolean
Dim m_ContainerIsVisible As Boolean
Dim m_ContainerRC As RECT               ' Control draw area within the container
Dim m_ForeClr As Long                   ' Label text color as pure RGB
Dim m_Format As Long                    ' DrawText format
Dim m_hFnt As Long                      ' Label API font object
Dim m_Partial As RECT                   ' Partial draw area
Dim m_PartialDraw As Boolean            ' Partial draw area is partial if True
Dim m_RC As RECT                        ' Label draw area in container
Dim m_ScaleMode As ScaleModeConstants   ' Container ScaleMode
Dim m_TrackComCtl As Boolean
Dim m_TrackUser32 As Boolean
Dim m_WideMode As Long                  ' Wide version API available

' for fixing XP Theme problem with a certain version of comctl32.dll
Dim m_FreeShell32 As Boolean
Dim m_Shell32 As Long

Dim IDE_DesignTime As Boolean           ' True if in IDE design time

Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets the alignment of a control's text."
Attribute Alignment.VB_ProcData.VB_Invoke_Property = ";Misc"
    Alignment = m_Alignment
End Property
Public Property Let Alignment(ByVal NewValue As AlignmentConstants)
    m_Alignment = NewValue
    ' update style
    Private_SetFormat
    ' repaint
    Private_Redraw
End Property
Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "Determines if a control is automatically resized to display its entire contents."
Attribute AutoSize.VB_ProcData.VB_Invoke_Property = ";Position"
Attribute AutoSize.VB_UserMemId = -500
    AutoSize = m_AutoSize
End Property
Public Property Let AutoSize(ByVal NewValue As Boolean)
    m_AutoSize = NewValue
    ' update style
    Private_SetFormat
    ' control might resize; if not, just redraw
    If m_AutoSize Then Private_SetSize Else Private_Redraw
End Property
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
    m_BackColor = NewValue
    If m_BackColor < 0 Then m_BackClr = GetSysColor(m_BackColor And &HFF&) Else m_BackClr = m_BackColor
    ' only redraw if the background is visible
    If m_BackStyle = ulbOpaque Then Private_Redraw 0&
End Property
Public Property Get BackStyle() As ulbBackStyle
Attribute BackStyle.VB_Description = "Indicates whether a Label is transparent or opaque."
Attribute BackStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackStyle.VB_UserMemId = -502
    BackStyle = m_BackStyle
End Property
Public Property Let BackStyle(ByVal NewValue As ulbBackStyle)
    m_BackStyle = NewValue
    ' redraw
    If m_BackStyle = ulbTransparent Then GetClientRect UserControl.ContainerHwnd, m_Container
    UserControl.BackStyle = m_BackStyle
End Property
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in a Label."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
Attribute Caption.VB_MemberFlags = "200"
    Caption = m_Caption
End Property
Public Property Let Caption(ByRef NewValue As String)
    m_Caption = NewValue
    ' control might resize; if not, just redraw
    If m_AutoSize Then Private_SetSize Else Private_Redraw
    ' if using events...
    If m_UseEvents And (Not IDE_DesignTime) Then RaiseEvent Change
End Property
Public Property Get Font() As IFontDisp
Attribute Font.VB_Description = "Returns/sets a Font object."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property
Public Property Set Font(ByRef NewValue As IFontDisp)
    Set m_Font = NewValue
    ' change font
    Private_SetFont
    ' control might resize; if not, just redraw
    If m_AutoSize Then Private_SetSize Else Private_Redraw
End Property
Public Property Get FontBold() As Boolean
    FontBold = m_Font.Bold
End Property
Public Property Let FontBold(ByVal NewValue As Boolean)
    m_Font.Bold = NewValue
    ' change font
    Private_SetFont
    ' control might resize; if not, just redraw
    If m_AutoSize Then Private_SetSize Else Private_Redraw
End Property
Public Property Get FontItalic() As Boolean
    FontItalic = m_Font.Italic
End Property
Public Property Let FontItalic(ByVal NewValue As Boolean)
    m_Font.Italic = NewValue
    ' change font
    Private_SetFont
    ' control might resize; if not, just redraw
    If m_AutoSize Then Private_SetSize Else Private_Redraw
End Property
Public Property Get FontName() As String
    FontName = m_Font.Name
End Property
Public Property Let FontName(ByRef NewValue As String)
    m_Font.Name = NewValue
    ' change font
    Private_SetFont
    ' control might resize; if not, just redraw
    If m_AutoSize Then Private_SetSize Else Private_Redraw
End Property
Public Property Get FontSize() As Single
    FontSize = m_Font.Size
End Property
Public Property Let FontSize(ByVal NewValue As Single)
    m_Font.Size = NewValue
    ' change font
    Private_SetFont
    ' control might resize; if not, just redraw
    If m_AutoSize Then Private_SetSize Else Private_Redraw
End Property
Public Property Get FontStrike() As Boolean
    FontStrike = m_Font.Strikethrough
End Property
Public Property Let FontStrike(ByVal NewValue As Boolean)
    m_Font.Strikethrough = NewValue
    ' change font
    Private_SetFont
    ' control might resize; if not, just redraw
    If m_AutoSize Then Private_SetSize Else Private_Redraw
End Property
Public Property Get FontUnderline() As Boolean
    FontUnderline = m_Font.Underline
End Property
Public Property Let FontUnderline(ByVal NewValue As Boolean)
    m_Font.Underline = NewValue
    ' change font
    Private_SetFont
    ' control might resize; if not, just redraw
    If m_AutoSize Then Private_SetSize Else Private_Redraw
End Property
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the text color used to display text and graphics in an object."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
    m_ForeColor = NewValue
    If m_ForeColor < 0 Then m_ForeClr = GetSysColor(m_ForeColor And &HFF&) Else m_ForeClr = m_ForeColor
    ' redraw
    Private_Redraw
End Property
Public Property Get MouseIcon() As IPictureDisp
Attribute MouseIcon.VB_ProcData.VB_Invoke_Property = ";Misc"
    Set MouseIcon = UserControl.MouseIcon
End Property
Public Property Set MouseIcon(ByRef NewValue As IPictureDisp)
    Set UserControl.MouseIcon = NewValue
End Property
Public Function MouseOver() As Boolean
    MouseOver = m_MouseOver
End Function
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_ProcData.VB_Invoke_Property = ";Misc"
    MousePointer = UserControl.MousePointer
End Property
Public Property Let MousePointer(ByVal NewValue As MousePointerConstants)
    UserControl.MousePointer = NewValue
End Property
Public Property Get PaddingBottom() As Byte
Attribute PaddingBottom.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PaddingBottom = m_PaddingBottom
End Property
Public Property Let PaddingBottom(ByVal NewValue As Byte)
    m_PaddingBottom = NewValue
    ' control might resize; check first for autosize
    If m_AutoSize Then Private_SetSize Else Private_Redraw
End Property
Public Property Get PaddingLeft() As Byte
Attribute PaddingLeft.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PaddingLeft = m_PaddingLeft
End Property
Public Property Let PaddingLeft(ByVal NewValue As Byte)
    m_PaddingLeft = NewValue
    ' control might resize; check first for autosize
    If m_AutoSize Then Private_SetSize Else Private_Redraw
End Property
Public Property Get PaddingRight() As Byte
Attribute PaddingRight.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PaddingRight = m_PaddingRight
End Property
Public Property Let PaddingRight(ByVal NewValue As Byte)
    m_PaddingRight = NewValue
    ' control might resize; check first for autosize
    If m_AutoSize Then Private_SetSize Else Private_Redraw
End Property
Public Property Get PaddingTop() As Byte
Attribute PaddingTop.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PaddingTop = m_PaddingTop
End Property
Public Property Let PaddingTop(ByVal NewValue As Byte)
    m_PaddingTop = NewValue
    ' control might resize; check first for autosize
    If m_AutoSize Then Private_SetSize Else Private_Redraw
End Property
Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute RightToLeft.VB_UserMemId = -611
    RightToLeft = m_RightToLeft
End Property
Public Property Let RightToLeft(ByVal NewValue As Boolean)
    m_RightToLeft = NewValue
    ' set text alignment
    If m_RightToLeft Then
        SetTextAlign UserControl.hDC, (GetTextAlign(UserControl.hDC) Or TA_RTLREADING)
    Else
        SetTextAlign UserControl.hDC, (GetTextAlign(UserControl.hDC) And Not TA_RTLREADING)
    End If
    ' redraw
    Private_Redraw
End Property
Public Sub SetColor(ByVal BackColor As OLE_COLOR, ByVal ForeColor As OLE_COLOR)
    m_BackColor = BackColor
    m_ForeColor = ForeColor
    If m_BackColor < 0 Then m_BackClr = GetSysColor(m_BackColor And &HFF&) Else m_BackClr = m_BackColor
    If m_ForeColor < 0 Then m_ForeClr = GetSysColor(m_ForeColor And &HFF&) Else m_ForeClr = m_ForeColor
    ' redraw
    Private_Redraw
End Sub
Public Sub SetPadding(ByVal Left As Byte, ByVal Top As Byte, ByVal Right As Byte, ByVal Bottom As Byte)
    m_PaddingLeft = Left
    m_PaddingTop = Top
    m_PaddingRight = Right
    m_PaddingBottom = Bottom
    ' control might resize; check first for autosize
    If m_AutoSize Then Private_SetSize Else Private_Redraw
End Sub
Public Property Get UseEvents() As Boolean
Attribute UseEvents.VB_ProcData.VB_Invoke_Property = ";Misc"
    UseEvents = m_UseEvents
End Property
Public Property Let UseEvents(ByVal NewValue As Boolean)
    m_UseEvents = NewValue
End Property
Public Property Get UseMnemonic() As Boolean
Attribute UseMnemonic.VB_ProcData.VB_Invoke_Property = ";Misc"
    UseMnemonic = m_UseMnemonic
End Property
Public Property Let UseMnemonic(ByVal NewValue As Boolean)
    m_UseMnemonic = NewValue
    ' update style
    Private_SetFormat
    ' control might resize; check first for autosize
    If m_AutoSize Then Private_SetSize Else Private_Redraw
End Property
Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_Description = "Returns/sets a value that determines whether a control expands to fit the text in its Caption."
Attribute WordWrap.VB_ProcData.VB_Invoke_Property = ";Misc"
    WordWrap = m_WordWrap
End Property
Public Property Let WordWrap(ByVal NewValue As Boolean)
    m_WordWrap = NewValue
    ' update style
    Private_SetFormat
    ' control might resize; check first for autosize
    If m_AutoSize Then Private_SetSize Else Private_Redraw
End Property

Private Sub UserControl_AmbientChanged(PropertyName As String)
    ' see if container scalemode has changed
    If LenB(PropertyName) = 20 Then If PropertyName = "ScaleUnits" Then m_ScaleMode = Private_GetContainerScaleMode
End Sub
Private Sub UserControl_Initialize()
    If Not Private_InIDE Then
        ' this will fix a problem with some versions of comctl32.dll when using XP Themes
        ' http://www.vbaccelerator.com/home/vb/Code/Libraries/XP_Visual_Styles/Preventing_Crashes_at_Shutdown/article.asp
        m_Shell32 = GetModuleHandleA("shell32.dll")
        If m_Shell32 = 0 Then m_Shell32 = LoadLibraryA("shell32.dll"): m_FreeShell32 = True
    End If
    ' NT supports Unicode; yet make sure the function is supported
    ' I made testing under Windows 98 and even TextOutW failed to show Unicode
    If Private_IsWinNT Then m_WideMode = Private_IsFunctionSupported("DrawTextW", "user32.dll")
    ' see if mouse tracking is supported (WM_MOUSELEAVE)
    m_TrackUser32 = Private_IsFunctionSupported("TrackMouseEvent", "user32.dll")
    If Not m_TrackUser32 Then m_TrackComCtl = Private_IsFunctionSupported("_TrackMouseEvent", "comctl32.dll")
End Sub
Private Sub UserControl_InitProperties()
    ' design time?
    IDE_DesignTime = (Not UserControl.Ambient.UserMode)
    ' container scalemode
    m_ScaleMode = Private_GetContainerScaleMode
    ' init colors
    m_BackColor = UserControl.Ambient.BackColor
    m_ForeColor = UserControl.Ambient.ForeColor
    If m_BackColor < 0 Then m_BackClr = GetSysColor(m_BackColor And &HFF&) Else m_BackClr = m_BackColor
    If m_ForeColor < 0 Then m_ForeClr = GetSysColor(m_ForeColor And &HFF&) Else m_ForeClr = m_ForeColor
    ' initial properties
    m_Alignment = vbLeftJustify
    m_AutoSize = False
    m_BackStyle = ulbOpaque
    m_Caption = UserControl.Ambient.DisplayName
    Set m_Font = UserControl.Ambient.Font
    m_UseEvents = True
    m_UseMnemonic = True
    m_WordWrap = False
    ' initialize
    Private_StartSubclass
    Private_SetFont
    Private_SetFormat
    Private_SetSize
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ' design time?
    IDE_DesignTime = (Not UserControl.Ambient.UserMode)
    ' container scalemode
    m_ScaleMode = Private_GetContainerScaleMode
    ' load settings
    m_Alignment = PropBag.ReadProperty("Alignment", vbAlignLeft)
    m_AutoSize = PropBag.ReadProperty("AutoSize", False)
    m_BackColor = PropBag.ReadProperty("BackColor", UserControl.Ambient.BackColor)
    m_BackStyle = PropBag.ReadProperty("BackStyle", ulbOpaque)
    UserControl.BackStyle = m_BackStyle
    m_Caption = PropBag.ReadProperty("Caption", UserControl.Ambient.DisplayName)
    Set m_Font = PropBag.ReadProperty("Font", UserControl.Ambient.Font)
    m_ForeColor = PropBag.ReadProperty("ForeColor", UserControl.Ambient.ForeColor)
    Set UserControl.MouseIcon = PropBag.ReadProperty("MouseIcon", UserControl.MouseIcon)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", UserControl.MousePointer)
    m_PaddingBottom = PropBag.ReadProperty("PaddingBottom", 0)
    m_PaddingLeft = PropBag.ReadProperty("PaddingLeft", 0)
    m_PaddingRight = PropBag.ReadProperty("PaddingRight", 0)
    m_PaddingTop = PropBag.ReadProperty("PaddingTop", 0)
    m_RightToLeft = PropBag.ReadProperty("RightToLeft", False)
    m_UseEvents = PropBag.ReadProperty("UseEvents", True)
    m_UseMnemonic = PropBag.ReadProperty("UseMnemonic", True)
    m_WordWrap = PropBag.ReadProperty("WordWrap", False)
    ' init colors
    If m_BackColor < 0 Then m_BackClr = GetSysColor(m_BackColor And &HFF&) Else m_BackClr = m_BackColor
    If m_ForeColor < 0 Then m_ForeClr = GetSysColor(m_ForeColor And &HFF&) Else m_ForeClr = m_ForeColor
    ' set text alignment
    If m_RightToLeft Then
        SetTextAlign UserControl.hDC, (GetTextAlign(UserControl.hDC) Or TA_RTLREADING)
    Else
        SetTextAlign UserControl.hDC, (GetTextAlign(UserControl.hDC) And Not TA_RTLREADING)
    End If
    ' initialize
    Private_StartSubclass
    Private_SetFont
    Private_SetFormat
    Private_SetSize
End Sub
Private Sub UserControl_Resize()
    If m_RC.Bottom <> UserControl.ScaleHeight Or m_RC.Right <> UserControl.ScaleWidth Then
        ' resize draw rectangle
        m_RC.Bottom = UserControl.ScaleHeight
        m_RC.Right = UserControl.ScaleWidth
        ' see if autosize...
        If m_AutoSize Then Private_SetSize
    End If
End Sub
Private Sub UserControl_Terminate()
    ' quit subclassing
    ssc_Terminate
    ' destroy what we have created
    If m_hFnt <> 0 Then DeleteObject m_hFnt
    ' unload shell32 if it was loaded by this control
    If m_FreeShell32 Then FreeLibrary m_Shell32
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    ' save settings
    PropBag.WriteProperty "Alignment", m_Alignment
    PropBag.WriteProperty "AutoSize", m_AutoSize
    PropBag.WriteProperty "BackColor", m_BackColor
    PropBag.WriteProperty "BackStyle", m_BackStyle
    PropBag.WriteProperty "Caption", m_Caption
    PropBag.WriteProperty "Font", m_Font
    PropBag.WriteProperty "ForeColor", m_ForeColor
    PropBag.WriteProperty "MouseIcon", UserControl.MouseIcon
    PropBag.WriteProperty "MousePointer", UserControl.MousePointer
    PropBag.WriteProperty "PaddingBottom", m_PaddingBottom
    PropBag.WriteProperty "PaddingLeft", m_PaddingLeft
    PropBag.WriteProperty "PaddingRight", m_PaddingRight
    PropBag.WriteProperty "PaddingTop", m_PaddingTop
    PropBag.WriteProperty "RightToLeft", m_RightToLeft
    PropBag.WriteProperty "UseEvents", m_UseEvents
    PropBag.WriteProperty "UseMnemonic", m_UseMnemonic
    PropBag.WriteProperty "WordWrap", m_WordWrap
End Sub

'-SelfSub code------------------------------------------------------------------------------------
'-The following routines are exclusively for the ssc_Subclass routines----------------------------
Private Function ssc_Subclass(ByVal lng_hWnd As Long, _
                    Optional ByVal lParamUser As Long = 0, _
                    Optional ByVal nOrdinal As Long = 1, _
                    Optional ByVal oCallback As Object = Nothing, _
                    Optional ByVal bIdeSafety As Boolean = True, _
                    Optional ByRef bUnicode As Boolean = False, _
                    Optional ByVal bIsAPIwindow As Boolean = False) As Boolean 'Subclass the specified window handle

    '*************************************************************************************************
    '* lng_hWnd   - Handle of the window to subclass
    '* lParamUser - Optional, user-defined callback parameter
    '* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
    '* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* bIdeSafety - Optional, enable/disable IDE safety measures. There is not reason to set this to False
    '* bUnicode - Optional, if True, Unicode API calls should be made to the window vs ANSI calls
    '*            Parameter is byRef and its return value should be checked to know if ANSI to be used or not
    '* bIsAPIwindow - Optional, if True DestroyWindow will be called if IDE ENDs
    '*****************************************************************************************
    '** Subclass.asm - subclassing thunk
    '**
    '** Paul_Caton@hotmail.com
    '** Copyright free, use and abuse as you see fit.
    '**
    '** v2.0 Re-write by LaVolpe, based mostly on Paul Caton's original thunks....... 20070720
    '** .... Reorganized & provided following additional logic
    '** ....... Unsubclassing only occurs after thunk is no longer recursed
    '** ....... Flag used to bypass callbacks until unsubclassing can occur
    '** ....... Timer used as delay mechanism to free thunk memory afer unsubclassing occurs
    '** .............. Prevents crash when one window subclassed multiple times
    '** .............. More END safe, even if END occurs within the subclass procedure
    '** ....... Added ability to destroy API windows when IDE terminates
    '** ....... Added auto-unsubclass when WM_NCDESTROY received
    '*****************************************************************************************
    ' Subclassing procedure must be declared identical to the one at the end of this class (Sample at Ordinal #1)

    Dim z_Sc(0 To IDX_UNICODE) As Long                 'Thunk machine-code initialised here
    
    Const SUB_NAME      As String = "ssc_Subclass"     'This routine's name
    Const CODE_LEN      As Long = 4 * IDX_UNICODE + 4  'Thunk length in bytes
    Const PAGE_RWX      As Long = &H40&                'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&              'Commit allocated memory
    Const MEM_RELEASE   As Long = &H8000&              'Release allocated memory flag
    Const GWL_WNDPROC   As Long = -4                   'SetWindowsLong WndProc index
    Const WNDPROC_OFF   As Long = &H60                 'Thunk offset to the WndProc execution address
    Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1)) 'Bytes to allocate per thunk, data + code + msg tables
    
  ' This is the complete listing of thunk offset values and what they point/relate to.
  ' Those rem'd out are used elsewhere or are initialized in Declarations section
  
  'Const IDX_RECURSION  As Long = 0     'Thunk data index of callback recursion count
  'Const IDX_SHUTDOWN   As Long = 1     'Thunk data index of the termination flag
  'Const IDX_INDEX      As Long = 2     'Thunk data index of the subclassed hWnd
   Const IDX_EBMODE     As Long = 3     'Thunk data index of the EbMode function address
   Const IDX_CWP        As Long = 4     'Thunk data index of the CallWindowProc function address
   Const IDX_SWL        As Long = 5     'Thunk data index of the SetWindowsLong function address
   Const IDX_FREE       As Long = 6     'Thunk data index of the VirtualFree function address
   Const IDX_BADPTR     As Long = 7     'Thunk data index of the IsBadCodePtr function address
   Const IDX_OWNER      As Long = 8     'Thunk data index of the Owner object's vTable address
  'Const IDX_PREVPROC   As Long = 9     'Thunk data index of the original WndProc
   Const IDX_CALLBACK   As Long = 10    'Thunk data index of the callback method address
  'Const IDX_BTABLE     As Long = 11    'Thunk data index of the Before table
  'Const IDX_ATABLE     As Long = 12    'Thunk data index of the After table
  'Const IDX_PARM_USER  As Long = 13    'Thunk data index of the User-defined callback parameter data index
   Const IDX_DW         As Long = 14    'Thunk data index of the DestroyWinodw function address
   Const IDX_ST         As Long = 15    'Thunk data index of the SetTimer function address
   Const IDX_KT         As Long = 16    'Thunk data index of the KillTimer function address
   Const IDX_EBX_TMR    As Long = 20    'Thunk code patch index of the thunk data for the delay timer
   Const IDX_EBX        As Long = 26    'Thunk code patch index of the thunk data
  'Const IDX_UNICODE    As Long = xx    'Must be UBound(subclass thunkdata)+1; index for unicode support
    
    Dim z_ScMem       As Long           'Thunk base address
    Dim nAddr         As Long
    Dim nID           As Long
    Dim nMyID         As Long
    Dim bIDE          As Boolean

    If IsWindow(lng_hWnd) = 0 Then      'Ensure the window handle is valid
        zError SUB_NAME, "Invalid window handle"
        Exit Function
    End If
    
    nMyID = GetCurrentProcessId                         'Get this process's ID
    GetWindowThreadProcessId lng_hWnd, nID              'Get the process ID associated with the window handle
    If nID <> nMyID Then                                'Ensure that the window handle doesn't belong to another process
        zError SUB_NAME, "Window handle belongs to another process"
        Exit Function
    End If
    
    If oCallback Is Nothing Then Set oCallback = Me     'If the user hasn't specified the callback owner
    
    nAddr = zAddressOf(oCallback, nOrdinal)             'Get the address of the specified ordinal method
    If nAddr = 0 Then                                   'Ensure that we've found the ordinal method
        zError SUB_NAME, "Callback method not found"
        Exit Function
    End If
        
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
    
    If z_ScMem <> 0 Then                                  'Ensure the allocation succeeded
    
      If z_scFunk Is Nothing Then Set z_scFunk = New Collection 'If this is the first time through, do the one-time initialization
      On Error GoTo CatchDoubleSub                              'Catch double subclassing
        z_scFunk.Add z_ScMem, "h" & lng_hWnd                    'Add the hWnd/thunk-address to the collection
      On Error GoTo 0
      
   'z_Sc (0) thru z_Sc(17) are used as storage for the thunks & IDX_ constants above relate to these thunk positions which are filled in below
    z_Sc(18) = &HD231C031: z_Sc(19) = &HBBE58960: z_Sc(21) = &H21E8F631: z_Sc(22) = &HE9000001: z_Sc(23) = &H12C&: z_Sc(24) = &HD231C031: z_Sc(25) = &HBBE58960: z_Sc(27) = &H3FFF631: z_Sc(28) = &H75047339: z_Sc(29) = &H2873FF23: z_Sc(30) = &H751C53FF: z_Sc(31) = &HC433913: z_Sc(32) = &H53FF2274: z_Sc(33) = &H13D0C: z_Sc(34) = &H18740000: z_Sc(35) = &H875C085: z_Sc(36) = &H820443C7: z_Sc(37) = &H90000000: z_Sc(38) = &H87E8&: z_Sc(39) = &H22E900: z_Sc(40) = &H90900000: z_Sc(41) = &H2C7B8B4A: z_Sc(42) = &HE81C7589: z_Sc(43) = &H90&: z_Sc(44) = &H75147539: z_Sc(45) = &H6AE80F: z_Sc(46) = &HD2310000: z_Sc(47) = &HE8307B8B: z_Sc(48) = &H7C&: z_Sc(49) = &H7D810BFF: z_Sc(50) = &H8228&: z_Sc(51) = &HC7097500: z_Sc(52) = &H80000443: z_Sc(53) = &H90900000: z_Sc(54) = &H44753339: z_Sc(55) = &H74047339: z_Sc(56) = &H2473FF3F: z_Sc(57) = &HFFFFFC68
    z_Sc(58) = &H2475FFFF: z_Sc(59) = &H811453FF: z_Sc(60) = &H82047B: z_Sc(61) = &HC750000: z_Sc(62) = &H74387339: z_Sc(63) = &H2475FF07: z_Sc(64) = &H903853FF: z_Sc(65) = &H81445B89: z_Sc(66) = &H484443: z_Sc(67) = &H73FF0000: z_Sc(68) = &H646844: z_Sc(69) = &H56560000: z_Sc(70) = &H893C53FF: z_Sc(71) = &H90904443: z_Sc(72) = &H10C261: z_Sc(73) = &H53E8&: z_Sc(74) = &H3075FF00: z_Sc(75) = &HFF2C75FF: z_Sc(76) = &H75FF2875: z_Sc(77) = &H2473FF24: z_Sc(78) = &H891053FF: z_Sc(79) = &H90C31C45: z_Sc(80) = &H34E30F8B: z_Sc(81) = &H1078C985: z_Sc(82) = &H4C781: z_Sc(83) = &H458B0000: z_Sc(84) = &H75AFF228: z_Sc(85) = &H90909023: z_Sc(86) = &H8D144D8D: z_Sc(87) = &H8D503443: z_Sc(88) = &H75FF1C45: z_Sc(89) = &H2C75FF30: z_Sc(90) = &HFF2875FF: z_Sc(91) = &H51502475: z_Sc(92) = &H2073FF52: z_Sc(93) = &H902853FF: z_Sc(94) = &H909090C3: z_Sc(95) = &H74447339: z_Sc(96) = &H4473FFF7
    z_Sc(97) = &H4053FF56: z_Sc(98) = &HC3447389: z_Sc(99) = &H89285D89: z_Sc(100) = &H45C72C75: z_Sc(101) = &H800030: z_Sc(102) = &H20458B00: z_Sc(103) = &H89145D89: z_Sc(104) = &H81612445: z_Sc(105) = &H4C4&: z_Sc(106) = &H1862FF00

    ' cache callback related pointers & offsets
      z_Sc(IDX_EBX) = z_ScMem                                                 'Patch the thunk data address
      z_Sc(IDX_EBX_TMR) = z_ScMem                                             'Patch the thunk data address
      z_Sc(IDX_INDEX) = lng_hWnd                                              'Store the window handle in the thunk data
      z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                                   'Store the address of the before table in the thunk data
      z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table in the thunk data
      z_Sc(IDX_OWNER) = ObjPtr(oCallback)                                     'Store the callback owner's object address in the thunk data
      z_Sc(IDX_CALLBACK) = nAddr                                              'Store the callback address in the thunk data
      z_Sc(IDX_PARM_USER) = lParamUser                                        'Store the lParamUser callback parameter in the thunk data
      
      ' validate unicode request & cache unicode usage
      If bUnicode Then bUnicode = (IsWindowUnicode(lng_hWnd) <> 0&)
      z_Sc(IDX_UNICODE) = bUnicode                                            'Store whether the window is using unicode calls or not
      
      ' get function pointers for the thunk
      If bIdeSafety = True Then                                               'If the user wants IDE protection
          Debug.Assert zInIDE(bIDE)
          If bIDE = True Then z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode", bUnicode) 'Store the EbMode function address in the thunk data
                                                        '^^ vb5 users, change vba6 to vba5
      End If
      If bIsAPIwindow Then                                                    'If user wants DestroyWindow sent should IDE end
          z_Sc(IDX_DW) = zFnAddr("user32", "DestroyWindow", bUnicode)
      End If
      z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree", bUnicode)           'Store the VirtualFree function address in the thunk data
      z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", bUnicode)        'Store the IsBadCodePtr function address in the thunk data
      z_Sc(IDX_ST) = zFnAddr("user32", "SetTimer", bUnicode)                  'Store the SetTimer function address in the thunk data
      z_Sc(IDX_KT) = zFnAddr("user32", "KillTimer", bUnicode)                 'Store the KillTimer function address in the thunk data
      
      If bUnicode Then
          z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcW", bUnicode)      'Store CallWindowProc function address in the thunk data
          z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongW", bUnicode)       'Store the SetWindowLong function address in the thunk data
          RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                    'Copy the thunk code/data to the allocated memory
          z_Sc(IDX_PREVPROC) = SetWindowLongW(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF) 'Set the new WndProc, return the address of the original WndProc
      Else
          z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA", bUnicode)      'Store CallWindowProc function address in the thunk data
          z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA", bUnicode)       'Store the SetWindowLong function address in the thunk data
          RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                    'Copy the thunk code/data to the allocated memory
          z_Sc(IDX_PREVPROC) = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF) 'Set the new WndProc, return the address of the original WndProc
      End If
      If z_Sc(IDX_PREVPROC) = 0 Then                                          'Ensure the new WndProc was set correctly
          zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
          GoTo ReleaseMemory
      End If
      'Store the original WndProc address in the thunk data
      RtlMoveMemory z_ScMem + IDX_PREVPROC * 4, VarPtr(z_Sc(IDX_PREVPROC)), 4&
      ssc_Subclass = True                                                     'Indicate success
      
    Else
        zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
        
    End If

 Exit Function                                                                'Exit ssc_Subclass
    
CatchDoubleSub:
 zError SUB_NAME, "Window handle is already subclassed"
      
ReleaseMemory:
      VirtualFree z_ScMem, 0, MEM_RELEASE                                     'ssc_Subclass has failed after memory allocation, so release the memory
      
End Function

'Terminate all subclassing
Private Sub ssc_Terminate()
    ' can be made public, can be removed & zTerminateThunks can be called instead
    zTerminateThunks SubclassThunk
End Sub

'UnSubclass the specified window handle
Private Sub ssc_UnSubclass(ByVal lng_hWnd As Long)
    ' can be made public, can be removed & zUnthunk can be called instead
    zUnThunk lng_hWnd, SubclassThunk
End Sub

'Add the message value to the window handle's specified callback table
Private Sub ssc_AddMsg(ByVal lng_hWnd As Long, ByVal When As eMsgWhen, ParamArray Messages() As Variant)
    
    Dim z_ScMem       As Long                                   'Thunk base address
    
    z_ScMem = zMap_VFunction(lng_hWnd, SubclassThunk)           'Ensure that the thunk hasn't already released its memory
    If z_ScMem Then
      Dim M As Long
      For M = LBound(Messages) To UBound(Messages)
        Select Case VarType(Messages(M))                        ' ensure no strings, arrays, doubles, objects, etc are passed
        Case vbByte, vbInteger, vbLong
            If When And MSG_BEFORE Then                         'If the message is to be added to the before original WndProc table...
              If zAddMsg(Messages(M), IDX_BTABLE, z_ScMem) = False Then 'Add the message to the before table
                When = (When And Not MSG_BEFORE)
              End If
            End If
            If When And MSG_AFTER Then                          'If message is to be added to the after original WndProc table...
              If zAddMsg(Messages(M), IDX_ATABLE, z_ScMem) = False Then 'Add the message to the after table
                When = (When And Not MSG_AFTER)
              End If
            End If
        End Select
      Next
    End If
End Sub

'Delete the message value from the window handle's specified callback table
Private Sub ssc_DelMsg(ByVal lng_hWnd As Long, ByVal When As eMsgWhen, ParamArray Messages() As Variant)
    
    Dim z_ScMem       As Long                                                   'Thunk base address
    
    z_ScMem = zMap_VFunction(lng_hWnd, SubclassThunk)                           'Ensure that the thunk hasn't already released its memory
    If z_ScMem Then
      Dim M As Long
      For M = LBound(Messages) To UBound(Messages)                              ' ensure no strings, arrays, doubles, objects, etc are passed
        Select Case VarType(Messages(M))
        Case vbByte, vbInteger, vbLong
            If When And MSG_BEFORE Then                                         'If the message is to be removed from the before original WndProc table...
              zDelMsg Messages(M), IDX_BTABLE, z_ScMem                          'Remove the message to the before table
            End If
            If When And MSG_AFTER Then                                          'If message is to be removed from the after original WndProc table...
              zDelMsg Messages(M), IDX_ATABLE, z_ScMem                          'Remove the message to the after table
            End If
        End Select
      Next
    End If
End Sub

'Call the original WndProc
Private Function ssc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ' can be made public, can be removed if you will not use this in your window procedure
    Dim z_ScMem       As Long                           'Thunk base address
    z_ScMem = zMap_VFunction(lng_hWnd, SubclassThunk)
    If z_ScMem Then                                     'Ensure that the thunk hasn't already released its memory
        If zData(IDX_UNICODE, z_ScMem) Then
            ssc_CallOrigWndProc = CallWindowProcW(zData(IDX_PREVPROC, z_ScMem), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
        Else
            ssc_CallOrigWndProc = CallWindowProcA(zData(IDX_PREVPROC, z_ScMem), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
        End If
    End If
End Function

'Get the subclasser lParamUser callback parameter
Private Function zGet_lParamUser(ByVal hWnd_Hook_ID As Long, ByVal vType As eThunkType) As Long
    ' can be removed if you never will retrieve or replace the user-defined parameter
    If vType <> CallbackThunk Then
        Dim z_ScMem       As Long                                       'Thunk base address
        z_ScMem = zMap_VFunction(hWnd_Hook_ID, vType)
        If z_ScMem Then                                                 'Ensure that the thunk hasn't already released its memory
          zGet_lParamUser = zData(IDX_PARM_USER, z_ScMem)               'Get the lParamUser callback parameter
        End If
    End If
End Function

'Let the subclasser lParamUser callback parameter
Private Sub zSet_lParamUser(ByVal hWnd_Hook_ID As Long, ByVal vType As eThunkType, ByVal NewValue As Long)
    ' can be removed if you never will retrieve or replace the user-defined parameter
    If vType <> CallbackThunk Then
        Dim z_ScMem       As Long                                       'Thunk base address
        z_ScMem = zMap_VFunction(hWnd_Hook_ID, vType)
        If z_ScMem Then                                                 'Ensure that the thunk hasn't already released its memory
          zData(IDX_PARM_USER, z_ScMem) = NewValue                      'Set the lParamUser callback parameter
        End If
    End If
End Sub

'Add the message to the specified table of the window handle
Private Function zAddMsg(ByVal uMsg As Long, ByVal nTable As Long, ByVal z_ScMem As Long) As Boolean
      Dim nCount As Long                                                        'Table entry count
      Dim nBase  As Long
      Dim i      As Long                                                        'Loop index
    
      zAddMsg = True
      nBase = zData(nTable, z_ScMem)                                            'Map zData() to the specified table
      
      If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table...
        nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES
      Else
        
        nCount = zData(0, nBase)                                                'Get the current table entry count
        For i = 1 To nCount                                                     'Loop through the table entries
          If zData(i, nBase) = 0 Then                                           'If the element is free...
            zData(i, nBase) = uMsg                                              'Use this element
            GoTo Bail                                                           'Bail
          ElseIf zData(i, nBase) = uMsg Then                                    'If the message is already in the table...
            GoTo Bail                                                           'Bail
          End If
        Next i                                                                  'Next message table entry
    
        nCount = i                                                             'On drop through: i = nCount + 1, the new table entry count
        If nCount > MSG_ENTRIES Then                                           'Check for message table overflow
          zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
          zAddMsg = False
          GoTo Bail
        End If
        
        zData(nCount, nBase) = uMsg                                            'Store the message in the appended table entry
      End If
    
      zData(0, nBase) = nCount                                                 'Store the new table entry count
Bail:
End Function

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long, ByVal z_ScMem As Long)
      Dim nCount As Long                                                        'Table entry count
      Dim nBase  As Long
      Dim i      As Long                                                        'Loop index
    
      nBase = zData(nTable, z_ScMem)                                            'Map zData() to the specified table
    
      If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
        zData(0, nBase) = 0                                                     'Zero the table entry count
      Else
        nCount = zData(0, nBase)                                                'Get the table entry count
        
        For i = 1 To nCount                                                     'Loop through the table entries
          If zData(i, nBase) = uMsg Then                                        'If the message is found...
            zData(i, nBase) = 0                                                 'Null the msg value -- also frees the element for re-use
            GoTo Bail                                                           'Bail
          End If
        Next i                                                                  'Next message table entry
        
       ' zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
      End If
Bail:
End Sub

'-The following routines are used for each of the three types of thunks ----------------------------

'Maps zData() to the memory address for the specified thunk type
Private Function zMap_VFunction(vFuncTarget As Long, _
                                vType As eThunkType, _
                                Optional oCallback As Object, _
                                Optional bIgnoreErrors As Boolean) As Long
    
    Dim thunkCol As Collection
    Dim colID As String
    Dim z_ScMem       As Long         'Thunk base address
    
    If vType = CallbackThunk Then
        Set thunkCol = z_cbFunk
        If oCallback Is Nothing Then Set oCallback = Me
        colID = "h" & ObjPtr(oCallback) & "." & vFuncTarget
    ElseIf vType = HookThunk Then
        Set thunkCol = z_hkFunk
        colID = "h" & vFuncTarget
    ElseIf vType = SubclassThunk Then
        Set thunkCol = z_scFunk
        colID = "h" & vFuncTarget
    Else
        zError "zMap_Vfunction", "Invalid thunk type passed"
        Exit Function
    End If
    
    If thunkCol Is Nothing Then
        zError "zMap_VFunction", "Thunk hasn't been initialized"
    Else
        If thunkCol.Count Then
            On Error GoTo Catch
            z_ScMem = thunkCol(colID)               'Get the thunk address
            If IsBadCodePtr(z_ScMem) Then z_ScMem = 0&
            zMap_VFunction = z_ScMem
        End If
    End If
    Exit Function                                               'Exit returning the thunk address
    
Catch:
    ' error ignored when zUnThunk is called, error handled there
    If Not bIgnoreErrors Then zError "zMap_VFunction", "Thunk type for " & vType & " does not exist"
End Function

' sets/retrieves data at the specified offset for the specified memory address
Private Property Get zData(ByVal nIndex As Long, ByVal z_ScMem As Long) As Long
  RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4
End Property

Private Property Let zData(ByVal nIndex As Long, ByVal z_ScMem As Long, ByVal nValue As Long)
  RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4
End Property

'Error handler
Private Sub zError(ByRef sRoutine As String, ByVal sMsg As String)
  ' Note. These two lines can be rem'd out if you so desire. But don't remove the routine
  App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
  MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String, ByVal asUnicode As Boolean) As Long
  If asUnicode Then
    zFnAddr = GetProcAddress(GetModuleHandleW(StrPtr(sDLL)), sProc)         'Get the specified procedure address
  Else
    zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                 'Get the specified procedure address
  End If
  Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
  ' ^^ FYI VB5 users. Search for zFnAddr("vba6", "EbMode") and replace with zFnAddr("vba5", "EbMode")
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
    ' Note: used both in subclassing and hooking routines
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim i     As Long                                                         'Loop index
  Dim J     As Long                                                         'Loop limit
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, i, bSub) Then                                 'Probe for a Class method
    If Not zProbe(nAddr + &H6F8, i, bSub) Then                              'Probe for a Form method
      If Not zProbe(nAddr + &H710, i, bSub) Then                            'Probe for a PropertyPage method
        If Not zProbe(nAddr + &H7A4, i, bSub) Then                          'Probe for a UserControl method
            Exit Function                                                   'Bail...
        End If
      End If
    End If
  End If
  
  i = i + 4                                                                 'Bump to the next entry
  J = i + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
  Do While i < J
    RtlMoveMemory VarPtr(nAddr), i, 4                                       'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If

    RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    
    i = i + 4                                                               'Next vTable entry
  Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32                                                       'Probe eight entries
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Do                                                             'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop
End Function

Private Function zInIDE(ByRef bIDE As Boolean) As Boolean
    ' only called in IDE, never called when compiled
    bIDE = True
    zInIDE = bIDE
End Function

Private Sub zUnThunk(ByVal thunkID As Long, ByVal vType As eThunkType, Optional ByVal oCallback As Object)

    ' thunkID, depends on vType:
    '   - Subclassing:  the hWnd of the window subclassed
    '   - Hooking:      the hook type created
    '   - Callbacks:    the ordinal of the callback
    '       ensure KillTimer is already called, if any callback used for SetTimer
    ' oCallback only used when vType is CallbackThunk

    Const IDX_SHUTDOWN  As Long = 1
    Const MEM_RELEASE As Long = &H8000&             'Release allocated memory flag
    
    Dim z_ScMem       As Long                       'Thunk base address
    
    z_ScMem = zMap_VFunction(thunkID, vType, oCallback, True)
    Select Case vType
    Case SubclassThunk
        If z_ScMem Then                                 'Ensure that the thunk hasn't already released its memory
            zData(IDX_SHUTDOWN, z_ScMem) = 1            'Set the shutdown indicator
            zDelMsg ALL_MESSAGES, IDX_BTABLE, z_ScMem   'Delete all before messages
            zDelMsg ALL_MESSAGES, IDX_ATABLE, z_ScMem   'Delete all after messages
        End If
        z_scFunk.Remove "h" & thunkID                   'Remove the specified thunk from the collection
        
    Case HookThunk
        If z_ScMem Then                                 'Ensure that the thunk hasn't already released its memory
            ' if not unhooked, then unhook now
            If zData(IDX_SHUTDOWN, z_ScMem) = 0 Then UnhookWindowsHookEx zData(IDX_PREVPROC, z_ScMem)
            If zData(0, z_ScMem) = 0 Then               ' not recursing then
                VirtualFree z_ScMem, 0, MEM_RELEASE     'Release allocated memory
                z_hkFunk.Remove "h" & thunkID           'Remove the specified thunk from the collection
            Else
                zData(IDX_SHUTDOWN, z_ScMem) = 1        ' Set the shutdown indicator
                zData(IDX_ATABLE, z_ScMem) = 0          ' want no more After messages
                zData(IDX_BTABLE, z_ScMem) = 0          ' want no more Before messages
                ' when zTerminate is called this thunk's memory will be released
            End If
        Else
            z_hkFunk.Remove "h" & thunkID       'Remove the specified thunk from the collection
        End If
    Case CallbackThunk
        If z_ScMem Then                         'Ensure that the thunk hasn't already released its memory
            VirtualFree z_ScMem, 0, MEM_RELEASE 'Release allocated memory
        End If
        z_cbFunk.Remove "h" & ObjPtr(oCallback) & "." & thunkID           'Remove the specified thunk from the collection
    End Select

End Sub

Private Sub zTerminateThunks(ByVal vType As eThunkType)

    ' Terminates all thunks of a specific type
    ' Any subclassing, hooking, recurring callbacks should have already been canceled

    Dim i As Long
    Dim oCallback As Object
    Dim thunkCol As Collection
    Dim z_ScMem       As Long                           'Thunk base address
    Const INDX_OWNER As Long = 0
    
    Select Case vType
    Case SubclassThunk
        Set thunkCol = z_scFunk
    Case HookThunk
        Set thunkCol = z_hkFunk
    Case CallbackThunk
        Set thunkCol = z_cbFunk
    Case Else
        Exit Sub
    End Select
    
    If Not (thunkCol Is Nothing) Then                 'Ensure that hooking has been started
      With thunkCol
        For i = .Count To 1 Step -1                   'Loop through the collection of hook types in reverse order
          z_ScMem = .Item(i)                          'Get the thunk address
          If IsBadCodePtr(z_ScMem) = 0 Then           'Ensure that the thunk hasn't already released its memory
            Select Case vType
                Case SubclassThunk
                    zUnThunk zData(IDX_INDEX, z_ScMem), SubclassThunk    'Unsubclass
                Case HookThunk
                    zUnThunk zData(IDX_INDEX, z_ScMem), HookThunk        'Unhook
                Case CallbackThunk
                    ' zUnThunk expects object not pointer, convert pointer to object
                    RtlMoveMemory VarPtr(oCallback), VarPtr(zData(INDX_OWNER, z_ScMem)), 4&
                    zUnThunk zData(IDX_CALLBACKORDINAL, z_ScMem), CallbackThunk, oCallback ' release callback
                    ' remove the object pointer reference
                    RtlMoveMemory VarPtr(oCallback), VarPtr(INDX_OWNER), 4&
            End Select
          End If
        Next i                                        'Next member of the collection
      End With
      Set thunkCol = Nothing                         'Destroy the hook/thunk-address collection
    End If


End Sub

Private Function Private_GetContainerScaleMode() As ScaleModeConstants
    ' this should be called only when we know scalemode has changed
    Select Case UserControl.Ambient.ScaleUnits
        Case "Twip"
            Private_GetContainerScaleMode = vbTwips
        Case "Point"
            Private_GetContainerScaleMode = vbPoints
        Case "Pixel"
            Private_GetContainerScaleMode = vbPixels
        Case "Character"
            Private_GetContainerScaleMode = vbCharacters
        Case "Inch"
            Private_GetContainerScaleMode = vbInches
        Case "Millimeter"
            Private_GetContainerScaleMode = vbMillimeters
        Case "Centimeter"
            Private_GetContainerScaleMode = vbCentimeters
        Case "User"
            ' prevent user scalemode
            UserControl.Extender.Container.ScaleMode = vbTwips
            Private_GetContainerScaleMode = vbTwips
    End Select
End Function
Private Function Private_InIDE() As Boolean: Debug.Assert True Xor Private_InIDEcheck(Private_InIDE): End Function
Private Function Private_InIDEcheck(IDE As Boolean) As Boolean: IDE = True: End Function
Private Function Private_IsFunctionSupported(ByRef FunctionName As String, ByRef ModuleName As String) As Boolean
    Dim lngModule As Long, blnUnload As Boolean
    ' get handle to module
    lngModule = GetModuleHandleA(ModuleName)
    ' if getting the handle failed...
    If lngModule = 0 Then
        ' try loading the module
        lngModule = LoadLibraryA(ModuleName)
        ' we have to unload it too if that succeeded
        blnUnload = (lngModule <> 0)
    End If
    ' now if we have a handle to module...
    If lngModule <> 0 Then
        ' see if the queried function is supported; return True if so, False if not
        Private_IsFunctionSupported = (GetProcAddress(lngModule, FunctionName) <> 0)
        ' see if we have to unload the module
        If blnUnload Then FreeLibrary lngModule
    End If
End Function
Private Function Private_IsWinNT() As Boolean
    ' this is actually OSVERSIONINFOEX
    Dim VersionInfo(36) As Long
    ' size for version information
    VersionInfo(0) = 148
    ' if passing version information doesn't fail...
    If GetVersionEx(VarPtr(VersionInfo(0))) <> 0 Then
        ' see if the version is NT
        Private_IsWinNT = (VersionInfo(4) = 2)
    End If
End Function
Private Sub Private_Redraw(Optional EraseBG As Long = -1&)
    If m_BackStyle = ulbOpaque Then
        InvalidateRect UserControl.hWnd, m_RC, EraseBG
    Else
        m_ContainerRC.Left = ScaleX(UserControl.Extender.Left, m_ScaleMode, vbPixels)
        m_ContainerRC.Top = ScaleX(UserControl.Extender.Top, m_ScaleMode, vbPixels)
        m_ContainerRC.Right = m_ContainerRC.Left + UserControl.ScaleWidth
        m_ContainerRC.Bottom = m_ContainerRC.Top + UserControl.ScaleHeight
        InvalidateRect UserControl.ContainerHwnd, m_ContainerRC, EraseBG
    End If
End Sub
Private Sub Private_SetFont()
    Dim uLF As LOGFONT, lngLen As Long
    ' destroy old font if it exists
    If m_hFnt <> 0 Then DeleteObject m_hFnt
    ' initialize font settings
    With m_Font
        ' determine length of font name
        If Len(.Name) >= LF_FACESIZE Then lngLen = LF_FACESIZE Else lngLen = Len(.Name)
        ' copy maximum allowed length
        CopyMemory uLF.lfFaceName(0), ByVal .Name, lngLen
        ' set other font settings
        uLF.lfHeight = -MulDiv(.Size, GetDeviceCaps(UserControl.hDC, LOGPIXELSY), 72)
        uLF.lfItalic = .Italic
        If Not .Bold Then uLF.lfWeight = FW_NORMAL Else uLF.lfWeight = FW_BOLD
        uLF.lfUnderline = .Underline
        uLF.lfStrikeOut = .Strikethrough
        uLF.lfCharSet = .Charset
    End With
    ' create new font
    m_hFnt = CreateFontIndirect(uLF)
End Sub
Private Sub Private_SetFormat()
    ' initialize with alignment
    If m_Alignment = vbLeftJustify Then
        m_Format = DT_LEFT
    ElseIf m_Alignment = vbCenter Then
        m_Format = DT_CENTER
    Else
        m_Format = DT_RIGHT
    End If
    If Not m_UseMnemonic Then m_Format = m_Format Or DT_NOPREFIX
    ' word wrapping?
    If m_WordWrap Then
        ' word wrap
        m_Format = m_Format Or DT_WORDBREAK
    ElseIf m_AutoSize Then
        ' force to size if autosizing and not wrapping
        m_Format = m_Format Or DT_NOCLIP Or DT_SINGLELINE
    Else
        m_Format = m_Format Or DT_SINGLELINE
    End If
End Sub
Private Sub Private_SetSize()
    Dim RC As RECT, lngTemp As Long, sngNewWidth As Single, sngNewHeight As Single
    ' see if control needs to be resized
    If m_AutoSize Then
        ' oh yeah baby, lets change the size
        RC.Right = UserControl.ScaleWidth - m_PaddingLeft - m_PaddingRight
        ' change font
        lngTemp = SelectObject(UserControl.hDC, m_hFnt)
        ' get size
        If m_WideMode Then
            ' Unicode
            DrawTextW UserControl.hDC, StrPtr(m_Caption), Len(m_Caption), RC, m_Format Or DT_CALCRECT
        Else
            ' ANSI
            DrawTextA UserControl.hDC, m_Caption, Len(m_Caption), RC, m_Format Or DT_CALCRECT
        End If
        ' restore original font
        SelectObject UserControl.hDC, lngTemp
        ' see how to change control size
        If m_WordWrap Then
            ' word wrapping; just change height
            m_RC.Bottom = RC.Bottom + m_PaddingTop + m_PaddingBottom
            m_RC.Right = UserControl.ScaleWidth
            ' cause redraw
            sngNewHeight = ScaleY(m_RC.Bottom, vbPixels, vbTwips)
            If UserControl.Height <> sngNewHeight Then
                UserControl.Height = sngNewHeight
                If m_BackStyle = ulbTransparent Then Private_Redraw
            End If
        Else
            ' no word wrapping: change width and height
            m_RC.Right = RC.Right + m_PaddingLeft + m_PaddingRight
            m_RC.Bottom = RC.Bottom + m_PaddingTop + m_PaddingBottom
            ' calculate new size
            sngNewWidth = ScaleX(m_RC.Right, vbPixels, vbTwips)
            sngNewHeight = ScaleY(m_RC.Bottom, vbPixels, vbTwips)
            ' check if size changes
            If (UserControl.Height <> sngNewHeight) Or (UserControl.Width <> sngNewWidth) Then
                ' causes redraw
                UserControl.Size sngNewWidth, sngNewHeight
                If m_BackStyle = ulbTransparent Then Private_Redraw
            End If
        End If
    Else
        m_RC.Bottom = UserControl.ScaleHeight
        m_RC.Right = UserControl.ScaleWidth
        ' redraw
        Private_Redraw
    End If
End Sub
Private Sub Private_StartSubclass()
    ' try to start subclassing
    If ssc_Subclass(UserControl.hWnd, , 1, , Not IDE_DesignTime) Then
        ' mouse events
        ssc_AddMsg UserControl.hWnd, MSG_BEFORE, WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN
        ssc_AddMsg UserControl.hWnd, MSG_AFTER, WM_LBUTTONDBLCLK, WM_MBUTTONDBLCLK, WM_RBUTTONDBLCLK, _
            WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP, WM_MOUSEMOVE
        ' if mouse tracking is supported...
        If m_TrackUser32 Or m_TrackComCtl Then ssc_AddMsg UserControl.hWnd, MSG_AFTER, WM_MOUSELEAVE
        ' other messages
        ssc_AddMsg UserControl.hWnd, MSG_BEFORE, WM_DESTROY, WM_PAINT
        ssc_AddMsg UserControl.hWnd, MSG_AFTER, WM_SHOWWINDOW, WM_STYLECHANGED, WM_SYSCOLORCHANGE, WM_THEMECHANGED
    Else
        Debug.Print UserControl.Ambient.DisplayName & ": Subclassing failed!"
    End If
End Sub
' this must be before the last procedure! search for SelfSub if you need reasoning behind this
Private Sub Private_WndProcContainer(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lParamUser As Long)
    Dim lngPaintDC As Long
    Dim lnghDC As Long, lngBitmap As Long, lngBrush As Long, lngOldBrush As Long
    Dim rcDraw As RECT, rcRedraw As RECT, ptDraw As POINTAPI, lngTemp As Long, RC As RECT
    ' we only use this if we are in transparent mode
    If m_BackStyle = ulbOpaque Then Exit Sub
    ' BEFORE OR AFTER?
    If bBefore Then
        If uMsg = WM_PAINT Then
            ' get the partial update rectangle
            If GetUpdateRect(lng_hWnd, m_Partial, -1&) Then
                ' see if it is a partial draw area
                m_PartialDraw = (m_Partial.Left > m_Container.Left) Or (m_Partial.Top > m_Container.Top) Or (m_Partial.Right < m_Container.Right) Or (m_Partial.Bottom < m_Container.Bottom)
            End If
        ElseIf uMsg = WM_LBUTTONDOWN Or uMsg = WM_RBUTTONDOWN Then
            m_ContainerRC.Left = ScaleX(UserControl.Extender.Left, m_ScaleMode, vbPixels)
            m_ContainerRC.Top = ScaleX(UserControl.Extender.Top, m_ScaleMode, vbPixels)
            m_ContainerRC.Right = m_ContainerRC.Left + UserControl.ScaleWidth
            m_ContainerRC.Bottom = m_ContainerRC.Top + UserControl.ScaleHeight
            InvalidateRect lng_hWnd, m_ContainerRC, -1&
        End If
    ' BEFORE
    Else
    ' AFTER
        If uMsg = WM_PAINT Then
            m_ContainerRC.Left = ScaleX(UserControl.Extender.Left, m_ScaleMode, vbPixels)
            m_ContainerRC.Top = ScaleX(UserControl.Extender.Top, m_ScaleMode, vbPixels)
            m_ContainerRC.Right = m_ContainerRC.Left + UserControl.ScaleWidth
            m_ContainerRC.Bottom = m_ContainerRC.Top + UserControl.ScaleHeight
            lngPaintDC = GetDC(lng_hWnd)
            ' see if it is a partial draw
            If m_PartialDraw Then
                ' determine if the redraw area includes this area
                If m_Partial.Top >= m_ContainerRC.Bottom Then Exit Sub
                If m_Partial.Left >= m_ContainerRC.Right Then Exit Sub
                If m_Partial.Right <= m_ContainerRC.Left Then Exit Sub
                If m_Partial.Bottom <= m_ContainerRC.Top Then Exit Sub
                ' calculate redraw rect
                If m_ContainerRC.Left < m_Partial.Left Then
                    rcRedraw.Left = m_Partial.Left
                Else
                    rcRedraw.Left = m_ContainerRC.Left
                End If
                If m_ContainerRC.Top < m_Partial.Top Then
                    rcRedraw.Top = m_Partial.Top
                Else
                    rcRedraw.Top = m_ContainerRC.Top
                End If
                If m_ContainerRC.Right > m_Partial.Right Then
                    rcRedraw.Right = m_Partial.Right
                Else
                    rcRedraw.Right = m_ContainerRC.Right
                End If
                If m_ContainerRC.Bottom > m_Partial.Bottom Then
                    rcRedraw.Bottom = m_Partial.Bottom
                Else
                    rcRedraw.Bottom = m_ContainerRC.Bottom
                End If
                ' calculate bitmap drawing rect
                rcDraw.Right = rcRedraw.Right - rcRedraw.Left
                rcDraw.Bottom = rcRedraw.Bottom - rcRedraw.Top
                ' create a new DC to draw into
                lnghDC = CreateCompatibleDC(lngPaintDC)
                ' we can't have negative margins for wordwrapped drawtext...
                ' and non-left aligned text has issues
                If (Not m_WordWrap) And (m_Alignment = vbLeftJustify) Then
                    ' make a minimal bitmap to the DC
                    lngBitmap = CreateCompatibleBitmap(lngPaintDC, rcDraw.Right, rcDraw.Bottom)
                    ' destroy old 1 x 1 bitmap, set new bitmap
                    DeleteObject SelectObject(lnghDC, lngBitmap)
                    ' then draw some stuff based on what setting we use
                    If m_BackStyle = ulbOpaque Then
                        ' set brush color
                        lngBrush = CreateSolidBrush(m_BackClr)
                        ' switch brush
                        lngOldBrush = SelectObject(lnghDC, lngBrush)
                        ' fill with new color
                        FillRect lnghDC, rcDraw, lngBrush
                        ' restore original brush
                        lngBrush = SelectObject(lnghDC, lngOldBrush)
                        ' delete own brush
                        DeleteObject lngBrush
                        ' text background color
                        SetBkColor lnghDC, m_BackClr
                    Else
                        ' get current background piece to the image
                        BitBlt lnghDC, 0, 0, rcDraw.Right, rcDraw.Bottom, lngPaintDC, rcRedraw.Left, rcRedraw.Top, vbSrcCopy
                        ' text background transparent
                        SetBkMode lnghDC, 0&
                    End If
                    ' draw text only if there is text
                    If LenB(m_Caption) > 0 Then
                        ' change these for text drawing
                        rcDraw.Left = m_ContainerRC.Left - rcRedraw.Left + m_PaddingLeft
                        rcDraw.Top = m_ContainerRC.Top - rcRedraw.Top + m_PaddingTop
                        rcDraw.Right = rcDraw.Right + m_PaddingLeft
                        rcDraw.Bottom = rcDraw.Bottom + m_PaddingTop
                        ' text color
                        SetTextColor lnghDC, m_ForeClr
                        ' change font
                        lngTemp = SelectObject(lnghDC, m_hFnt)
                        ' draw the text
                        If m_WideMode Then
                            DrawTextW lnghDC, StrPtr(m_Caption), Len(m_Caption), rcDraw, m_Format
                        Else
                            ExtTextOutW lnghDC, rcDraw.Left, rcDraw.Top, ETO_CLIPPED, rcDraw, StrPtr(m_Caption), Len(m_Caption), ByVal 0&
                            'DrawTextA lnghDC, m_Caption, Len(m_Caption), rcDraw, m_Format
                        End If
                        ' restore original font
                        SelectObject lnghDC, lngTemp
                    End If
                    ' copy end result to the container
                    BitBlt lngPaintDC, rcRedraw.Left, rcRedraw.Top, rcDraw.Right, rcDraw.Bottom, lnghDC, 0, 0, vbSrcCopy
                ' WM_PAINT continues...
                Else
                    ' full width draw... otherwise wrapping messes up
                    rcDraw.Right = m_ContainerRC.Right - m_ContainerRC.Left
                    ' make a full bitmap to the DC
                    lngBitmap = CreateCompatibleBitmap(lngPaintDC, rcDraw.Right, rcDraw.Bottom)
                    ' destroy old 1 x 1 bitmap, set new bitmap
                    DeleteObject SelectObject(lnghDC, lngBitmap)
                    ' then draw some stuff based on what setting we use
                    If m_BackStyle = ulbOpaque Then
                        ' set brush color
                        lngBrush = CreateSolidBrush(m_BackClr)
                        ' switch brush
                        lngOldBrush = SelectObject(lnghDC, lngBrush)
                        ' fill with new color
                        FillRect lnghDC, rcDraw, lngBrush
                        ' restore original brush
                        lngBrush = SelectObject(lnghDC, lngOldBrush)
                        ' delete own brush
                        DeleteObject lngBrush
                        ' text background color
                        SetBkColor lnghDC, m_BackClr
                    Else
                        ' get current background piece to the image
                        BitBlt lnghDC, rcRedraw.Left - m_ContainerRC.Left, 0, rcRedraw.Right - rcRedraw.Left, rcDraw.Bottom, lngPaintDC, rcRedraw.Left, rcRedraw.Top, vbSrcCopy
                        ' text background transparent
                        SetBkMode lnghDC, 0&
                    End If
                    ' draw text only if there is text
                    If LenB(m_Caption) > 0 Then
                        rcDraw.Left = rcDraw.Left + m_PaddingLeft
                        rcDraw.Top = m_ContainerRC.Top - rcRedraw.Top + m_PaddingTop
                        rcDraw.Right = rcDraw.Right - m_PaddingRight
                        rcDraw.Bottom = rcDraw.Bottom + m_PaddingTop
                        ' text color
                        SetTextColor lnghDC, m_ForeClr
                        ' change font
                        lngTemp = SelectObject(lnghDC, m_hFnt)
                        ' draw the text
                        If m_WideMode Then
                            DrawTextW lnghDC, StrPtr(m_Caption), Len(m_Caption), rcDraw, m_Format
                        Else
                            DrawTextA lnghDC, m_Caption, Len(m_Caption), rcDraw, m_Format
                        End If
                        ' restore original font
                        SelectObject lnghDC, lngTemp
                    End If
                    ' copy end result to the container
                    BitBlt lngPaintDC, rcRedraw.Left, rcRedraw.Top, rcRedraw.Right - rcRedraw.Left, rcDraw.Bottom, lnghDC, rcRedraw.Left - m_ContainerRC.Left, 0, vbSrcCopy
                End If
                ' and remove the device context
                DeleteDC lnghDC
            ' WM_PAINT continues...
            Else
                ' full redraw, with background or not?
                If m_BackStyle = ulbOpaque Then
                    ' set brush color
                    lngBrush = CreateSolidBrush(m_BackClr)
                    lngOldBrush = SelectObject(lngPaintDC, lngBrush)
                    ' fill with new color
                    FillRect lngPaintDC, m_ContainerRC, lngBrush
                    lngBrush = SelectObject(lngPaintDC, lngOldBrush)
                    DeleteObject lngBrush
                    ' text background color
                    SetBkColor lngPaintDC, m_BackClr
                Else
                    ' text background transparent
                    SetBkMode lngPaintDC, 0&
                End If
                ' draw text only if there is text
                If LenB(m_Caption) > 0 Then
                    Dim TextRC As RECT
                    TextRC.Left = m_ContainerRC.Left + m_PaddingLeft
                    TextRC.Top = m_ContainerRC.Top + m_PaddingTop
                    TextRC.Right = m_ContainerRC.Right - m_PaddingRight
                    TextRC.Bottom = m_ContainerRC.Bottom - m_PaddingBottom
                    ' text color
                    SetTextColor lngPaintDC, m_ForeClr
                    ' change font
                    lngTemp = SelectObject(lngPaintDC, m_hFnt)
                    ' draw the text
                    If m_WideMode Then
                        DrawTextW lngPaintDC, StrPtr(m_Caption), Len(m_Caption), TextRC, m_Format
                    ElseIf (Not m_WordWrap) And (m_Alignment = vbLeftJustify) Then
                        ExtTextOutW lngPaintDC, TextRC.Left, TextRC.Top, ETO_CLIPPED, TextRC, StrPtr(m_Caption), Len(m_Caption), ByVal 0&
                    Else
                        DrawTextA lngPaintDC, m_Caption, Len(m_Caption), TextRC, m_Format
                    End If
                    ' restore original font
                    SelectObject lngPaintDC, lngTemp
                End If
            End If
        ElseIf uMsg = WM_LBUTTONUP Then
            InvalidateRect lng_hWnd, m_ContainerRC, -1&
            m_ContainerRC.Left = ScaleX(UserControl.Extender.Left, m_ScaleMode, vbPixels)
            m_ContainerRC.Top = ScaleX(UserControl.Extender.Top, m_ScaleMode, vbPixels)
            m_ContainerRC.Right = m_ContainerRC.Left + UserControl.ScaleWidth
            m_ContainerRC.Bottom = m_ContainerRC.Top + UserControl.ScaleHeight
            InvalidateRect lng_hWnd, m_ContainerRC, -1&
        ElseIf uMsg = WM_EXITSIZEMOVE Then
            ' get container full client draw area
            GetClientRect lng_hWnd, m_Container
        ElseIf uMsg = WM_SHOWWINDOW Then
            m_ContainerIsVisible = (wParam = 1&)
            If m_ContainerIsVisible Then
                Debug.Print "Container is visible!"
            Else
                Debug.Print "Container is invisible!"
            End If
        End If
    End If
End Sub
' this must be the last procedure! search for SelfSub if you need reasoning behind this
Private Sub Private_WndProcUserControl(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lParamUser As Long)
    Dim TME As TRACKMOUSEEVENT_STRUCT, PS As PAINTSTRUCT, lngPaintDC As Long
    Dim lnghDC As Long, lngBitmap As Long, lngBrush As Long, lngOldBrush As Long
    Dim rcDraw As RECT, rcRedraw As RECT, ptDraw As POINTAPI, lngTemp As Long, RC As RECT
    Dim XY(1) As Integer, Button As ulbMouseButtonConstants, Shift As ulbShiftConstants
    ' BEFORE or AFTER?
    If bBefore Then
        If uMsg = WM_PAINT Then
            ' we do not draw if we are in transparent mode
            If m_BackStyle = ulbTransparent Then Exit Sub
            ' check what area is about to be redrawn
            If GetUpdateRect(lng_hWnd, m_Partial, -1&) Then
                ' see if it is a partial draw area
                m_PartialDraw = (m_Partial.Left > 0) Or (m_Partial.Top > 0) Or (m_Partial.Right < m_RC.Right) Or (m_Partial.Bottom < m_RC.Bottom)
                ' start painting
                lngPaintDC = BeginPaint(lng_hWnd, PS)
                ' do painting
                ' see if it is a partial draw
                If m_PartialDraw Then
                    ' determine if the redraw area includes this area
                    If m_Partial.Top >= m_RC.Bottom Then Exit Sub
                    If m_Partial.Left >= m_RC.Right Then Exit Sub
                    If m_Partial.Right <= m_RC.Left Then Exit Sub
                    If m_Partial.Bottom <= m_RC.Top Then Exit Sub
                    ' calculate redraw rect
                    If m_RC.Left < m_Partial.Left Then
                        rcRedraw.Left = m_Partial.Left
                    Else
                        rcRedraw.Left = m_RC.Left
                    End If
                    If m_RC.Top < m_Partial.Top Then
                        rcRedraw.Top = m_Partial.Top
                    Else
                        rcRedraw.Top = m_RC.Top
                    End If
                    If m_RC.Right > m_Partial.Right Then
                        rcRedraw.Right = m_Partial.Right
                    Else
                        rcRedraw.Right = m_RC.Right
                    End If
                    If m_RC.Bottom > m_Partial.Bottom Then
                        rcRedraw.Bottom = m_Partial.Bottom
                    Else
                        rcRedraw.Bottom = m_RC.Bottom
                    End If
                    ' calculate bitmap drawing rect
                    rcDraw.Right = rcRedraw.Right - rcRedraw.Left
                    rcDraw.Bottom = rcRedraw.Bottom - rcRedraw.Top
                    ' create a new DC to draw into
                    lnghDC = CreateCompatibleDC(lngPaintDC)
                    ' we can't have negative margins for wordwrapped drawtext...
                    ' and non-left aligned text has issues
                    If (Not m_WordWrap) And (m_Alignment = vbLeftJustify) Then
                        ' make a minimal bitmap to the DC
                        lngBitmap = CreateCompatibleBitmap(lngPaintDC, rcDraw.Right, rcDraw.Bottom)
                        ' destroy old 1 x 1 bitmap, set new bitmap
                        DeleteObject SelectObject(lnghDC, lngBitmap)
                        ' then draw some stuff based on what setting we use
                        If m_BackStyle = ulbOpaque Then
                            ' set brush color
                            lngBrush = CreateSolidBrush(m_BackClr)
                            ' switch brush
                            lngOldBrush = SelectObject(lnghDC, lngBrush)
                            ' fill with new color
                            FillRect lnghDC, rcDraw, lngBrush
                            ' restore original brush
                            lngBrush = SelectObject(lnghDC, lngOldBrush)
                            ' delete own brush
                            DeleteObject lngBrush
                            ' text background color
                            SetBkColor lnghDC, m_BackClr
                        Else
                            ' get current background piece to the image
                            BitBlt lnghDC, 0, 0, rcDraw.Right, rcDraw.Bottom, lngPaintDC, rcRedraw.Left, rcRedraw.Top, vbSrcCopy
                            ' text background transparent
                            SetBkMode lnghDC, 0&
                        End If
                        ' draw text only if there is text
                        If LenB(m_Caption) > 0 Then
                            ' change these for text drawing
                            rcDraw.Left = m_RC.Left - rcRedraw.Left + m_PaddingLeft
                            rcDraw.Top = m_RC.Top - rcRedraw.Top + m_PaddingTop
                            rcDraw.Right = rcDraw.Right + m_PaddingLeft
                            rcDraw.Bottom = rcDraw.Bottom + m_PaddingTop
                            ' text color
                            SetTextColor lnghDC, m_ForeClr
                            ' change font
                            lngTemp = SelectObject(lnghDC, m_hFnt)
                            ' draw the text
                            If m_WideMode Then
                                DrawTextW lnghDC, StrPtr(m_Caption), Len(m_Caption), rcDraw, m_Format
                            Else
                                ExtTextOutW lnghDC, rcDraw.Left, rcDraw.Top, ETO_CLIPPED, rcDraw, StrPtr(m_Caption), Len(m_Caption), ByVal 0&
                                'DrawTextA lnghDC, m_Caption, Len(m_Caption), rcDraw, m_Format
                            End If
                            ' restore original font
                            SelectObject lnghDC, lngTemp
                        End If
                        ' copy end result to the container
                        BitBlt lngPaintDC, rcRedraw.Left, rcRedraw.Top, rcDraw.Right, rcDraw.Bottom, lnghDC, 0, 0, vbSrcCopy
                    ' WM_PAINT continues...
                    Else
                        ' full width draw... otherwise wrapping messes up
                        rcDraw.Right = m_RC.Right - m_RC.Left
                        ' make a full bitmap to the DC
                        lngBitmap = CreateCompatibleBitmap(lngPaintDC, rcDraw.Right, rcDraw.Bottom)
                        ' destroy old 1 x 1 bitmap, set new bitmap
                        DeleteObject SelectObject(lnghDC, lngBitmap)
                        ' then draw some stuff based on what setting we use
                        If m_BackStyle = ulbOpaque Then
                            ' set brush color
                            lngBrush = CreateSolidBrush(m_BackClr)
                            ' switch brush
                            lngOldBrush = SelectObject(lnghDC, lngBrush)
                            ' fill with new color
                            FillRect lnghDC, rcDraw, lngBrush
                            ' restore original brush
                            lngBrush = SelectObject(lnghDC, lngOldBrush)
                            ' delete own brush
                            DeleteObject lngBrush
                            ' text background color
                            SetBkColor lnghDC, m_BackClr
                        Else
                            ' get current background piece to the image
                            BitBlt lnghDC, rcRedraw.Left - m_RC.Left, 0, rcRedraw.Right - rcRedraw.Left, rcDraw.Bottom, lngPaintDC, rcRedraw.Left, rcRedraw.Top, vbSrcCopy
                            ' text background transparent
                            SetBkMode lnghDC, 0&
                        End If
                        ' draw text only if there is text
                        If LenB(m_Caption) > 0 Then
                            rcDraw.Left = rcDraw.Left + m_PaddingLeft
                            rcDraw.Top = m_RC.Top - rcRedraw.Top + m_PaddingTop
                            rcDraw.Right = rcDraw.Right - m_PaddingRight
                            rcDraw.Bottom = rcDraw.Bottom + m_PaddingTop
                            ' text color
                            SetTextColor lnghDC, m_ForeClr
                            ' change font
                            lngTemp = SelectObject(lnghDC, m_hFnt)
                            ' draw the text
                            If m_WideMode Then
                                DrawTextW lnghDC, StrPtr(m_Caption), Len(m_Caption), rcDraw, m_Format
                            Else
                                DrawTextA lnghDC, m_Caption, Len(m_Caption), rcDraw, m_Format
                            End If
                            ' restore original font
                            SelectObject lnghDC, lngTemp
                        End If
                        ' copy end result to the container
                        BitBlt lngPaintDC, rcRedraw.Left, rcRedraw.Top, rcRedraw.Right - rcRedraw.Left, rcDraw.Bottom, lnghDC, rcRedraw.Left - m_RC.Left, 0, vbSrcCopy
                    End If
                    ' and remove the device context
                    DeleteDC lnghDC
                ' WM_PAINT continues...
                Else
                    ' full redraw, with background or not?
                    If m_BackStyle = ulbOpaque Then
                        ' set brush color
                        lngBrush = CreateSolidBrush(m_BackClr)
                        lngOldBrush = SelectObject(lngPaintDC, lngBrush)
                        ' fill with new color
                        FillRect lngPaintDC, m_RC, lngBrush
                        lngBrush = SelectObject(lngPaintDC, lngOldBrush)
                        DeleteObject lngBrush
                        ' text background color
                        SetBkColor lngPaintDC, m_BackClr
                    Else
                        ' text background transparent
                        SetBkMode lngPaintDC, 0&
                    End If
                    ' draw text only if there is text
                    If LenB(m_Caption) > 0 Then
                        Dim TextRC As RECT
                        TextRC.Left = m_PaddingLeft
                        TextRC.Top = m_PaddingTop
                        TextRC.Right = m_RC.Right - m_PaddingRight
                        TextRC.Bottom = m_RC.Bottom - m_PaddingBottom
                        ' text color
                        SetTextColor lngPaintDC, m_ForeClr
                        ' change font
                        lngTemp = SelectObject(lngPaintDC, m_hFnt)
                        ' draw the text
                        If m_WideMode Then
                            DrawTextW lngPaintDC, StrPtr(m_Caption), Len(m_Caption), TextRC, m_Format
                        ElseIf (Not m_WordWrap) And (m_Alignment = vbLeftJustify) Then
                            ExtTextOutW lngPaintDC, TextRC.Left, TextRC.Top, ETO_CLIPPED, TextRC, StrPtr(m_Caption), Len(m_Caption), ByVal 0&
                        Else
                            DrawTextA lngPaintDC, m_Caption, Len(m_Caption), TextRC, m_Format
                        End If
                        ' restore original font
                        SelectObject lngPaintDC, lngTemp
                    End If
                End If
                ' end painting
                EndPaint lng_hWnd, PS
                ' mark handled
                bHandled = True
            End If
            ' end of WM_PAINT
        ElseIf uMsg = WM_DESTROY Then
            ssc_UnSubclass UserControl.ContainerHwnd
        ElseIf m_UseEvents Then
            Select Case uMsg
                Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN
                    ' this horribility is one reason to not use events...
                    CopyMemory XY(0), lParam, 4&
                    ' see key states
                    If (wParam And MK_SHIFT) = MK_SHIFT Then Shift = ulbShiftMask
                    If (wParam And MK_CONTROL) = MK_CONTROL Then Shift = Shift Or ulbCtrlMask
                    ' button states
                    If (wParam And MK_LBUTTON) = MK_LBUTTON Then Button = ulbLeftButton
                    If (wParam And MK_MBUTTON) = MK_MBUTTON Then Button = Button Or ulbMiddleButton
                    If (wParam And MK_RBUTTON) = MK_RBUTTON Then Button = Button Or ulbRightButton
                    ' raise the event
                    RaiseEvent MouseDown(Button, Shift, ScaleX(XY(0), vbPixels, m_ScaleMode), ScaleY(XY(1), vbPixels, m_ScaleMode))
            End Select
        End If
    ' BEFORE
    Else
    ' AFTER
        Select Case uMsg
            Case WM_MOUSEMOVE
                ' see if entering into the control
                If Not m_MouseOver Then
                    ' initialize TrackMouseEvent structure
                    TME.cbSize = Len(TME)
                    TME.dwFlags = TME_LEAVE
                    TME.hWndTrack = lng_hWnd
                    ' see which tracking API is available, if any
                    If m_TrackUser32 Then
                        TrackMouseEventUser32 TME
                    ElseIf m_TrackComCtl Then
                        TrackMouseEventComCtl TME
                    End If
                    ' set mouseover
                    m_MouseOver = True
                    ' raise event if using events
                    If m_UseEvents And (Not IDE_DesignTime) Then RaiseEvent MouseEnter
                End If
                ' see if need to raise events...
                If m_UseEvents Then
                    ' this horribility is one reason to not use events...
                    CopyMemory XY(0), lParam, 4&
                    ' see key states
                    If (wParam And MK_SHIFT) = MK_SHIFT Then Shift = ulbShiftMask
                    If (wParam And MK_CONTROL) = MK_CONTROL Then Shift = Shift Or ulbCtrlMask
                    ' button states
                    If (wParam And MK_LBUTTON) = MK_LBUTTON Then Button = ulbLeftButton
                    If (wParam And MK_MBUTTON) = MK_MBUTTON Then Button = Button Or ulbMiddleButton
                    If (wParam And MK_RBUTTON) = MK_RBUTTON Then Button = Button Or ulbRightButton
                    ' raise the event
                    RaiseEvent MouseMove(Button, Shift, ScaleX(XY(0), vbPixels, m_ScaleMode), ScaleY(XY(1), vbPixels, m_ScaleMode))
                End If
            Case WM_MOUSELEAVE
                m_MouseOver = False
                ' raise event if using events
                If m_UseEvents And (Not IDE_DesignTime) Then RaiseEvent MouseLeave
            Case WM_LBUTTONUP
                If m_UseEvents Then
                    ' this horribility is one reason to not use events...
                    CopyMemory XY(0), lParam, 4&
                    ' see key states
                    If (wParam And MK_SHIFT) = MK_SHIFT Then Shift = ulbShiftMask
                    If (wParam And MK_CONTROL) = MK_CONTROL Then Shift = Shift Or ulbCtrlMask
                    ' button states
                    If (wParam And MK_LBUTTON) = MK_LBUTTON Then Button = ulbLeftButton
                    If (wParam And MK_MBUTTON) = MK_MBUTTON Then Button = Button Or ulbMiddleButton
                    If (wParam And MK_RBUTTON) = MK_RBUTTON Then Button = Button Or ulbRightButton
                    ' raise the event
                    RaiseEvent MouseUp(Button, Shift, ScaleX(XY(0), vbPixels, m_ScaleMode), ScaleY(XY(1), vbPixels, m_ScaleMode))
                    ' click
                    RaiseEvent Click(vbLeftButton)
                End If
            Case WM_MBUTTONUP
                If m_UseEvents Then
                    ' this horribility is one reason to not use events...
                    CopyMemory XY(0), lParam, 4&
                    ' see key states
                    If (wParam And MK_SHIFT) = MK_SHIFT Then Shift = ulbShiftMask
                    If (wParam And MK_CONTROL) = MK_CONTROL Then Shift = Shift Or ulbCtrlMask
                    ' button states
                    If (wParam And MK_LBUTTON) = MK_LBUTTON Then Button = ulbLeftButton
                    If (wParam And MK_MBUTTON) = MK_MBUTTON Then Button = Button Or ulbMiddleButton
                    If (wParam And MK_RBUTTON) = MK_RBUTTON Then Button = Button Or ulbRightButton
                    ' raise the event
                    RaiseEvent MouseUp(Button, Shift, ScaleX(XY(0), vbPixels, m_ScaleMode), ScaleY(XY(1), vbPixels, m_ScaleMode))
                    ' click
                    RaiseEvent Click(vbMiddleButton)
                End If
            Case WM_RBUTTONUP
                If m_UseEvents Then
                    ' this horribility is one reason to not use events...
                    CopyMemory XY(0), lParam, 4&
                    ' see key states
                    If (wParam And MK_SHIFT) = MK_SHIFT Then Shift = ulbShiftMask
                    If (wParam And MK_CONTROL) = MK_CONTROL Then Shift = Shift Or ulbCtrlMask
                    ' button states
                    If (wParam And MK_LBUTTON) = MK_LBUTTON Then Button = ulbLeftButton
                    If (wParam And MK_MBUTTON) = MK_MBUTTON Then Button = Button Or ulbMiddleButton
                    If (wParam And MK_RBUTTON) = MK_RBUTTON Then Button = Button Or ulbRightButton
                    ' raise the event
                    RaiseEvent MouseUp(Button, Shift, ScaleX(XY(0), vbPixels, m_ScaleMode), ScaleY(XY(1), vbPixels, m_ScaleMode))
                    ' click
                    RaiseEvent Click(vbRightButton)
                End If
            Case WM_LBUTTONDBLCLK
                If m_UseEvents Then RaiseEvent DblClick(vbLeftButton)
            Case WM_MBUTTONDBLCLK
                If m_UseEvents Then RaiseEvent DblClick(vbMiddleButton)
            Case WM_RBUTTONDBLCLK
                If m_UseEvents Then RaiseEvent DblClick(vbRightButton)
            Case WM_STYLECHANGED, WM_SYSCOLORCHANGE, WM_THEMECHANGED
                ' if using system colors...
                If (m_BackColor < 0) Or (m_ForeColor < 0) Then
                    ' update colors as necessary
                    If m_BackColor < 0 Then m_BackClr = GetSysColor(m_BackColor And &HFF&) Else m_BackClr = m_BackColor
                    If m_ForeColor < 0 Then m_ForeClr = GetSysColor(m_ForeColor And &HFF&) Else m_ForeClr = m_ForeColor
                    ' force redraw
                    Private_Redraw
                End If
            Case WM_SHOWWINDOW
                If wParam = 1& Then
                    If m_ContainerhWnd <> 0 Then ssc_UnSubclass m_ContainerhWnd
                    m_ContainerhWnd = GetParent(UserControl.hWnd)
                    ' try to subclass container
                    If ssc_Subclass(m_ContainerhWnd, , 2, , Not IDE_DesignTime) Then
                        ssc_AddMsg m_ContainerhWnd, MSG_AFTER, WM_EXITSIZEMOVE
                        ssc_AddMsg m_ContainerhWnd, MSG_BEFORE_AFTER, WM_PAINT
                        ssc_AddMsg m_ContainerhWnd, MSG_AFTER, WM_SHOWWINDOW
                        If IDE_DesignTime Then
                            ssc_AddMsg m_ContainerhWnd, MSG_BEFORE, WM_LBUTTONDOWN, WM_RBUTTONDOWN
                            ssc_AddMsg m_ContainerhWnd, MSG_AFTER, WM_LBUTTONUP
                        End If
                        On Error Resume Next
                        m_ContainerIsParent = (UserControl.Parent.hWnd = m_ContainerhWnd)
                        If Err.Number <> 0 Then m_ContainerIsParent = False
                        On Error GoTo 0
                        Debug.Print UserControl.Ambient.DisplayName & ": Started subclassing!"
                    Else
                        m_ContainerhWnd = 0
                        Debug.Print UserControl.Ambient.DisplayName & ": Container subclassing failed!"
                    End If
                ElseIf m_ContainerhWnd <> 0 Then
                    Debug.Print UserControl.Ambient.DisplayName & ": Ended subclassing!"
                    ' unsubclass
                    If (Not m_ContainerIsParent) Or IDE_DesignTime Then ssc_UnSubclass m_ContainerhWnd
                    m_ContainerhWnd = 0
                Else
                    Debug.Print UserControl.Ambient.DisplayName & ": Not subclassed!"
                End If
        End Select
    End If
End Sub
