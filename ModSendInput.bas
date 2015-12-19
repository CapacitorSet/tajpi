Attribute VB_Name = "ModSendInput"
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

Public Sub SendBack()

    UninstallHook
    
    Call keybd_event(VK_BACK, 0, 0, 0)
    Call keybd_event(VK_BACK, 0, KEYEVENTF_KEYUP, 0)
        
    InstallHook
 
End Sub

Public Sub SendCtrlX()

    UninstallHook
    
     'knock out ctrl
    Call keybd_event(VK_LCONTROL, 0, KEYEVENTF_KEYUP, 0)
    Call keybd_event(VK_RCONTROL, 0, KEYEVENTF_KEYUP, 0)
    
    'knock out alt
    Call keybd_event(VK_LMENU, 0, KEYEVENTF_KEYUP, 0)
    Call keybd_event(VK_RMENU, 0, KEYEVENTF_KEYUP, 0)
    
    'knock out shift
    Call keybd_event(VK_LSHIFT, 0, KEYEVENTF_KEYUP, 0)
    Call keybd_event(VK_RSHIFT, 0, KEYEVENTF_KEYUP, 0)
       
    'send ctrl+x
    Call keybd_event(VK_LCONTROL, 0, 0, 0)
    Call keybd_event(VK_KEY_X, 0, 0, 0)
    Call keybd_event(VK_KEY_X, 0, KEYEVENTF_KEYUP, 0)
    Call keybd_event(VK_LCONTROL, 0, KEYEVENTF_KEYUP, 0)
    
    InstallHook
 
End Sub

Public Sub SendCtrlV()

    UninstallHook
    
     'knock out the shift key as capital V doesn't work in many applications, must be lowercase v
    Call keybd_event(VK_LSHIFT, 0, KEYEVENTF_KEYUP, 0)
    Call keybd_event(VK_RSHIFT, 0, KEYEVENTF_KEYUP, 0)
    
    'knock out alt-gr so that pasted character can appear
    Call keybd_event(VK_RMENU, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
    
    'knock out ctrl and alt so as not to interfere with the character
    keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
    keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0
    
    'send ctrl+v
    Call keybd_event(VK_LCONTROL, 0, 0, 0)
    Call keybd_event(VK_KEY_V, 0, 0, 0)
    Call keybd_event(VK_KEY_V, 0, KEYEVENTF_KEYUP, 0)
    Call keybd_event(VK_LCONTROL, 0, KEYEVENTF_KEYUP, 0)
    
    'restore shift
    If lShiftDown Then
        Call keybd_event(VK_LSHIFT, 0, 0, 0)
    End If
    If rShiftDown Then
        Call keybd_event(VK_RSHIFT, 0, 0, 0)
    End If
        
    InstallHook
 
End Sub

Public Sub SendShiftInsert()

    UninstallHook
    
    'knock out alt-gr so that pasted character can appear
    Call keybd_event(VK_RMENU, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
    
    'knock out ctrl and alt so as not to interfere with the character
    keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
    keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0
    
    'send shift+insert
    Call keybd_event(VK_LSHIFT, 0, 0, 0)
    Call keybd_event(VK_INSERT, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(VK_INSERT, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
    Call keybd_event(VK_LSHIFT, 0, KEYEVENTF_KEYUP, 0)
    
    'restore shift
    If lShiftDown Then
        Call keybd_event(VK_LSHIFT, 0, 0, 0)
    End If
        
    InstallHook
 
End Sub

Public Sub SendString(data As String)
            
    UninstallHook
    
    Dim i As Integer
    Dim char As String
    Dim kbinput(2) As KeyboardInput
    
    For i = 1 To Len(data)
        
        char = Mid$(data, i, 1)
            
        kbinput(0).dwType = INPUT_KEYBOARD
        kbinput(0).wVK = 0
        kbinput(0).wScan = AscW(char)
        kbinput(0).dwFlags = KEYEVENTF_UNICODE
        kbinput(0).dwTime = 0
        kbinput(0).dwExtraInfo = 0
        
        kbinput(1).dwType = INPUT_KEYBOARD
        kbinput(1).wVK = 0
        kbinput(1).wScan = AscW(char)
        kbinput(1).dwFlags = KEYEVENTF_UNICODE Or KEYEVENTF_KEYUP
        kbinput(1).dwTime = 0
        kbinput(1).dwExtraInfo = 0

        Call SendInput(2, kbinput(0), Len(kbinput(0)))
    Next
    
    InstallHook
    
End Sub

Public Sub PasteSend(ByVal data As String, Optional ByVal useConfigPasteMethod As Boolean = True)

    If FrmClipboardMethod.ChkRestore.Value Then
        If Not FrmMain.Timer1.enabled Then
            clipsave = GetClipboardUnicode(FrmMain.hWnd)
        End If
    End If
    
    Clipboard.Clear
    If SetClipboardUnicode(FrmMain.hWnd, data) Then
        
        If useConfigPasteMethod Then
        
            If FrmClipboardMethod.RBtnCtrlV.Value Then
                SendCtrlV
            Else
                SendShiftInsert
            End If
            
        Else
            
            SendCtrlV
        
        End If
        
    End If
    
    If FrmClipboardMethod.ChkRestore.Value Then
        FrmMain.Timer1.enabled = False
        FrmMain.Timer1.enabled = True
    End If

End Sub

Public Sub Send(ByVal data As String)
    
    On Error Resume Next
    
    If FrmConfig.ChkPaste.Value Then
        PasteSend (data)
    Else
        SendString (data)
    End If
    
End Sub

