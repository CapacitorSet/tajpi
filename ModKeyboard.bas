Attribute VB_Name = "ModKeyboard"
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

Public data As NotifyIconData
Public capsLockOn As Boolean
Public shiftDown As Boolean
Public lShiftDown As Boolean
Public rShiftDown As Boolean
Public ctrlDown As Boolean
Public lCtrlDown As Boolean
Public rCtrlDown As Boolean
Public OEM8Down As Boolean
Public altDown As Boolean
Public lAltDown As Boolean
Public altGrDown As Boolean
Public allowWinUp As Boolean
Public upperCase As Boolean
Public buffer As String
Public lastCharSuffixed As String
Public lastSuffixChar As String
Public lastAutoEuAu As String
Public lastPrefixKey As String
Public deadKey As String
Public cyrillic As Boolean
Public russian As Boolean
Public hook As Long
Public foregroundWindow As Long
Public remoteThreadId As Long
Public currentThreadId As Long
Public focusedWindow As Long

Dim map As Long
Dim layout As Long
Dim keys(255) As Byte
Dim ascii As Long
Dim key As String
Dim unshifted As String
Dim sKey As String

Public Function LowLevelKeyboardProc(ByVal nCode As Long, _
                                     ByVal wParam As Long, _
                                     lParam As KBDLLHOOKSTRUCT) As Long
         
    On Error GoTo Finish
    
    If nCode = HC_ACTION Then
    
        If wParam = WM_KEYUP Then
            If lParam.vkCode = VK_OEM_8 Then
                OEM8Down = False
            End If
        End If
        
        'not interested in these keys
        If lParam.vkCode = VK_LSHIFT Or _
           lParam.vkCode = VK_RSHIFT Or _
           lParam.vkCode = VK_LMENU Or _
           lParam.vkCode = VK_RMENU Or _
           lParam.vkCode = VK_LCONTROL Or _
           lParam.vkCode = VK_RCONTROL Or _
           lParam.vkCode = VK_CAPITAL Or _
           lParam.vkCode = VK_NUMLOCK Then
            GoTo Finish
        End If
                                                        
        If wParam = WM_KEYDOWN Or wParam = WM_SYSKEYDOWN Then
                               
            ctrlDown = (GetKeyState(VK_CONTROL) And &H80) = &H80
            lCtrlDown = (GetKeyState(VK_LCONTROL) And &H80) = &H80
            rCtrlDown = (GetKeyState(VK_RCONTROL) And &H80) = &H80
            altDown = (GetKeyState(VK_MENU) And &H80) = &H80
            lAltDown = (GetKeyState(VK_LMENU) And &H80) = &H80
            altGrDown = (GetKeyState(VK_RMENU) And &H80) = &H80
            lShiftDown = (GetKeyState(VK_LSHIFT) And &H80) = &H80
            rShiftDown = (GetKeyState(VK_RSHIFT) And &H80) = &H80
            shiftDown = (GetKeyState(VK_SHIFT) And &H80) = &H80
            capsLockOn = (GetKeyState(VK_CAPITAL) And &H1) = &H1
            upperCase = IIf(capsLockOn, Not shiftDown, shiftDown)
            
            If lParam.vkCode = VK_OEM_8 Then
                OEM8Down = True
            End If
            
            'get layout of focused control
            foregroundWindow = GetForegroundWindow()
            remoteThreadId = GetWindowThreadProcessId(foregroundWindow, 0)
            currentThreadId = GetCurrentThreadId()
            Call AttachThreadInput(remoteThreadId, currentThreadId, 1)
            focusedWindow = GetFocus()
            Call AttachThreadInput(remoteThreadId, currentThreadId, 0)
            layout = GetKeyboardLayout(GetWindowThreadProcessId(focusedWindow, 0))
            
            If lParam.vkCode = VK_OEM_8 And layout = FRENCH_CANADIAN_MULTILINGUAL_STANDARD Then
                GoTo Finish
            End If
            
            'map the key based on layout
            map = MapVirtualKeyEx(lParam.vkCode, 2, layout)
            key = ChrW(map And &HFFF)
                
            unshifted = key
                      
            'SCRIPT
            If scriptLoaded Then
                If ProcessHotKey(lParam.vkCode, key) Then
                    LowLevelKeyboardProc = 1
                    Exit Function
                End If
            End If
            
            'SWITCH ON/OFF
            If FrmConfig.CmbKeys.ListIndex > 0 Then
                If lParam.vkCode = FrmConfig.CmbKeys.ItemData(FrmConfig.CmbKeys.ListIndex) Or UCase$(key) = FrmConfig.CmbKeys.Text Then
            
                    If ctrlDown = (FrmConfig.ChkCtrl.Value = 1) And _
                       altDown = (FrmConfig.ChkAlt.Value = 1) And _
                       shiftDown = (FrmConfig.ChkShift.Value = 1) Then
                        
                        Call FrmMain.SetTajpiEnabled(Not FrmMain.MenuActive.Checked, True)
                        cyrillic = False
                        russian = False
                        LowLevelKeyboardProc = 1
                        Exit Function
                    
                    End If
                End If
            End If
                 
            'is tajpi active?
            If Not FrmMain.MenuActive.Checked Then
                GoTo Finish
            End If
            
            'make sure input to direct keys text boxes are not shifted non-alpha chars
            If FrmConfig.DirectKeyBoxFocused() Then
                If shiftDown And Not (key Like "[A-Za-z]" Or _
                                      lParam.vkCode = VK_BACK Or _
                                      lParam.vkCode = VK_LEFT Or _
                                      lParam.vkCode = VK_RIGHT) Then
                    LowLevelKeyboardProc = 1
                    Exit Function
                End If
            End If
            
            'ignore keypress if a textbox on the config form was focused
            If (foregroundWindow = FrmConfig.hWnd Or foregroundWindow = FrmClipboardMethod.hWnd) And _
                FrmConfig.TextBoxFocused Then
                GoTo Finish
            End If
            
            'check for complex dead key as MapVirtualKeyEx() does not report these
            Dim cDeadKey As String
            cDeadKey = ComplexDeadKey(key, layout)
                                        
            'dead-key handling
            If deadKey <> "" Then
                
                'key pressed after a dead key. we ignore these and allow the dead key combination to be processed normally,
                'unless the dead key was a configured prefix and the key pressed after it was an accentable esperanto character
                If FrmConfig.ChkPrefixes.Value And IsAccentable(key) And IsPrefixKey(deadKey) Then
                    
                    'accent the character
                    Send (EOKey(IIf(upperCase, key, LCase$(key)), _
                          FrmConfig.RBtnUnicode.Value, _
                          FrmConfig.ChkEntityCodes.Value, _
                          (cyrillic Or russian)))
                    
                    'flush the dead key char from the internal buffer
                    Call ToAsciiEx(VK_SPACE, lParam.ScanCode, keys(0), ascii, 0&, layout)
                    
                    ClearBuffer
                    LowLevelKeyboardProc = 1
                    Exit Function
                
                Else
                    
                    ClearBuffer
                    GoTo Finish
                    
                End If
                
            ElseIf cDeadKey <> "" Then
            
                'complex dead key
                deadKey = cDeadKey
                GoTo Finish
            
            ElseIf map < 0 Then
                
                'regular dead key
                deadKey = key
                GoTo Finish
                      
            Else
            
                deadKey = ""
            
            End If
            
            'ALT-GR
            If (altGrDown Or (ctrlDown And altDown)) And FrmConfig.ChkAltGr.Value And IsAccentable(key) Then
                                        
                'knock out alt-gr so as not to interfere with the character
                If altGrDown Then
                    keybd_event VK_RMENU, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
                End If
                
                'knock out ctrl and alt so as not to interfere with the character
                If ctrlDown And altDown Then
                    keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
                    keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0
                End If
                                
                ' accent the character
                Send (EOKey(IIf(upperCase, key, LCase$(key)), _
                      FrmConfig.RBtnUnicode.Value, _
                      FrmConfig.ChkEntityCodes.Value, _
                      (cyrillic Or russian)))
                                               
                ClearBuffer
                LowLevelKeyboardProc = 1
                Exit Function
            End If
                                   
            'translate the key
            Call GetKeyboardState(keys(0))
            Call ToAsciiEx(lParam.vkCode, lParam.ScanCode, keys(0), ascii, 0&, layout)
                                  
            'build the buffer
            If ascii And map <> 0 Then
                key = Chr(ascii)
            Else
                key = " "
            End If
            buffer = buffer & key
            If Len(buffer) > 2 Then
                buffer = Right$(buffer, 2)
            End If
                             
            'DIRECT KEYS
            If FrmConfig.ChkDirectKeys.Value Then
                Dim i As Integer
                For i = FrmConfig.TxtDirectKey.LBound To FrmConfig.TxtDirectKey.UBound
                    If FrmConfig.TxtDirectKey(i).Text <> "" Then
                        If UCase$(unshifted) = FrmConfig.TxtDirectKey(i).Text Then
                            If altGrDown Or (ctrlDown And altDown) Then
                                
                                ' if altGr is pressed then allow the actual key to go through rather
                                ' than input an accented character. we knock out altGr first though, so
                                ' that it can actually be displayed
                                If altGrDown Then
                                    keybd_event VK_RMENU, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
                                End If
                                If ctrlDown And altDown Then
                                    keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
                                    keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0
                                End If
                                
                                ClearBuffer
                                GoTo Finish
                            
                            Else
                            
                                sKey = FrmConfig.TxtDirectKey(i).Tag
                                sKey = EOKey(IIf(upperCase, UCase$(sKey), sKey), _
                                             FrmConfig.RBtnUnicode.Value, _
                                             FrmConfig.ChkEntityCodes.Value, _
                                             (cyrillic Or russian))
                                
                                Send (sKey)
                                lastPrefixKey = ""
                                LowLevelKeyboardProc = 1
                                Exit Function
                            
                            End If
                        End If
                    End If
                Next
            End If
            If cyrillic Or russian Then
                If Not IsSuffixKey(UCase$(key)) And Not IsDirectKey(UCase$(key)) Then
                    sKey = CyrillicKey(key, FrmConfig.ChkEntityCodes.Value)
                    If sKey <> "" Then
                        Send (sKey)
                        lastCharSuffixed = ""
                        LowLevelKeyboardProc = 1
                        Exit Function
                    End If
                End If
            End If
            
            Dim char1, char1U As String
            Dim char2, char2U As String
            If Len(buffer) = 2 Then
                char1 = Mid$(buffer, 1, 1)
                char2 = Mid$(buffer, 2, 1)
                char1U = UCase$(char1)
                char2U = UCase$(char2)
            ElseIf Len(buffer) = 1 Then
                char1 = ""
                char2 = Mid$(buffer, 1, 1)
                char1U = ""
                char2U = UCase$(char2)
            End If
            
            'AUTO AU/EU
            If FrmConfig.ChkAutomaticAuEu.Value Then
                
                If char2U = "U" Then
                    
                    If lastAutoEuAu <> "" And FrmConfig.ChkSuffixesRepeat.Value Then
                        
                        'remove the accent
                        SendBack
                        If cyrillic Or russian Then
                            Send (CyrillicKey(lastAutoEuAu))
                        Else
                            Send (lastAutoEuAu)
                        End If
                            
                        ClearBuffer
                        LowLevelKeyboardProc = 1
                        Exit Function
                    
                    ElseIf (char1U = "A" Or char1U = "E") Then
                        
                        'send the accented char
                        Send (EOKey(char2, _
                             FrmConfig.RBtnUnicode.Value, _
                             FrmConfig.ChkEntityCodes.Value, _
                             (cyrillic Or russian)))
                        ClearBuffer
                        lastAutoEuAu = char2
                        LowLevelKeyboardProc = 1
                        Exit Function
                        
                    End If
                
                Else
                    lastAutoEuAu = ""
                End If
            
            Else
                lastAutoEuAu = ""
            End If
                          
            'SUFFIXES
            If FrmConfig.ChkSuffixes.Value And IsSuffixKey(char2U) Then
                If lastCharSuffixed <> "" And FrmConfig.ChkSuffixesRepeat.Value And char1U = char2U Then
                    'remove the accent
                    SendBack
                    If cyrillic Or russian Then
                        Send (CyrillicKey(lastCharSuffixed) & CyrillicKey(char2))
                    Else
                        Send (lastCharSuffixed & char2)
                    End If
                                                        
                    ClearBuffer
                    buffer = char2
                    LowLevelKeyboardProc = 1
                    Exit Function
                    
                ElseIf IsAccentable(char1) And char1 <> lastSuffixChar Then
                    'accent the character
                    SendBack
                    Send (EOKey(char1, _
                          FrmConfig.RBtnUnicode.Value, _
                          FrmConfig.ChkEntityCodes.Value, _
                          (cyrillic Or russian)))
                    ClearBuffer
                    If FrmConfig.ChkSuffixesRepeat.Value Then
                        buffer = char2
                    End If
                    lastCharSuffixed = char1
                    lastSuffixChar = char2
                    LowLevelKeyboardProc = 1
                    Exit Function
                
                Else
                    lastCharSuffixed = ""
                    lastSuffixChar = ""
                End If
                                        
            Else
                lastCharSuffixed = ""
                lastSuffixChar = ""
            End If
            
            'PREFIXES
            If FrmConfig.ChkPrefixes.Value Then
                
                If FrmConfig.ChkInvisibleSuffix.Value Then
                    If IsPrefixKey(char2U) And lastPrefixKey = "" Then
                        lastPrefixKey = char2
                        LowLevelKeyboardProc = 1
                        Exit Function
                    End If
                End If
               
                If IsAccentable(char2) Then
                                        
                    If IsPrefixKey(char1U) Then
                        
                        'accent the character
                        If lastPrefixKey = "" Then
                            SendBack
                        End If
                        Send (EOKey(char2, _
                             FrmConfig.RBtnUnicode.Value, _
                             FrmConfig.ChkEntityCodes.Value, _
                             (cyrillic Or russian)))
                        ClearBuffer
                        LowLevelKeyboardProc = 1
                        Exit Function
            
                    End If
                
                End If
        
            End If
            
            If lastPrefixKey <> "" And FrmConfig.ChkInvisibleSuffix.Value Then
                Send (lastPrefixKey)
                If key = " " Then
                    lastPrefixKey = ""
                    LowLevelKeyboardProc = 1
                    Exit Function
                End If
            End If
            lastPrefixKey = ""
        
        End If
    
    End If
    
Finish:
           
    LowLevelKeyboardProc = CallNextHookEx(hook, nCode, wParam, lParam)
                                              
   
End Function

Public Sub InstallHook()

    hook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0&)

End Sub

Public Sub UninstallHook()

    Call UnhookWindowsHookEx(hook)

End Sub

Public Sub ClearBuffer()

    buffer = ""
    lastPrefixKey = ""
    lastAutoEuAu = ""
    lastCharSuffixed = ""
    lastSuffixChar = ""
    deadKey = ""

End Sub
