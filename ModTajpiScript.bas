Attribute VB_Name = "ModTajpiScript"
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

Public script() As String
Public scriptLoaded As Boolean
Public scriptOn As Boolean

Public Sub LoadScript()

On Error GoTo HandleError
    
    Dim path As String
    path = App.path & "\Tajpi.skr"
            
    If FileExists(path) Then
   
        Dim B() As Byte
        Dim strBuff As String
        Open path For Binary Access Read As #1
        If LOF(1) Then
        
            ReDim B(LOF(1) - 1)
            Get #1, , B
            strBuff = RemoveBOM(StrConvFromUTF8(CStr(B)))
            script = Split(strBuff, vbCrLf)
            
            Dim i As Integer
            For i = LBound(script) To UBound(script)
                script(i) = Escape(DeComment(script(i)))
            Next
            
        End If
        Close #1

        scriptLoaded = True
        scriptOn = True
    
    Else
        
        scriptLoaded = False
        scriptOn = False
        
    End If
    
    Exit Sub
    
HandleError:
    
    scriptLoaded = False
    scriptOn = False
    
End Sub

Public Function Escape(ByVal str As String) As String
    Dim i As Long, curChar As String, EscapeMode As Boolean
    
    For i = 1 To Len(str)
        curChar = Mid$(str, i, 1)
        If EscapeMode = False Then
            If curChar = "\" Then
                EscapeMode = True
            Else
                Escape = Escape & curChar
            End If
        Else
            If curChar = "\" Then
                Escape = Escape & "\"
            ElseIf curChar = "n" Then
                Escape = Escape & vbCrLf
            ElseIf curChar = "t" Then
                Escape = Escape & vbTab
            Else
                Escape = Escape & curChar
            End If
            EscapeMode = False
        End If
    Next i
    
End Function

Public Function BlankEscapes(ByVal str As String) As String
    Dim i As Long, curChar As String, EscapeMode As Boolean
    
    For i = 1 To Len(str)
        curChar = Mid$(str, i, 1)
        If EscapeMode = False Then
            If curChar = "\" Then
                EscapeMode = True
            Else
                BlankEscapes = BlankEscapes & curChar
            End If
        Else
            BlankEscapes = BlankEscapes & "--"
            EscapeMode = False
        End If
    Next i
    
End Function

Public Function DeComment(line As String) As String

    DeComment = line
    Dim i As Integer
    Dim blanked As String
    blanked = BlankEscapes(line)
    i = InStr(blanked, ";")
    If i > 0 Then
        DeComment = Left$(line, i - 1)
    End If
    
End Function

Public Sub ProcessCommand(command As String)
    
    Dim clip As String
    
    If Left$(command, 1) = "_" And Len(command) > 1 Then
        
        If scriptOn Then
        
            Dim Text As String
            Text = Mid$(command, 2)
            Call PasteSend(Text, False)
           
        End If
    
    ElseIf IsNumeric(command) Then
           
        If scriptOn Then
           
            SendString (ChrW(Val(command)))
        
        End If
        
    Else
        
        command = Trim$(UCase$(command))
        
        Select Case command
        
            Case "RUSA"
                If scriptOn Then
                    russian = Not russian
                    If russian Then
                        cyrillic = False
                    End If
                End If
            
            Case "CIRILA"
                If scriptOn Then
                    cyrillic = Not cyrillic
                    If cyrillic Then
                        russian = False
                    End If
                End If
                      
            Case "AGORDI"
                If scriptOn Then
                    FrmConfig.Display
                End If
            
            Case "HELPO"
                If scriptOn Then
                    ShowHelp
                End If
            
            Case "HELPO_EO"
                If scriptOn Then
                    ShowHelp ("Esperanto")
                End If
                
            Case "HELPO_EN"
                If scriptOn Then
                    ShowHelp ("English")
                End If
                
            Case "ESPERANTO"
                If scriptOn Then
                    Call SetLanguage("Esperanto", True)
                    SaveConfig
                End If
            
            Case "ENGLISH"
                If scriptOn Then
                    Call SetLanguage("English", True)
                    SaveConfig
                End If
            
            Case "PRI_TAJPI"
                If scriptOn Then
                    FrmAbout.Show
                End If
            
            Case "ELIRI"
                If scriptOn Then
                    FrmMain.UnloadMe
                    End
                End If
            
            Case "SSIGI_ELEKTITAN"
                If scriptOn Then
                    Clipboard.Clear
                    SendCtrlX
                    MySleep (100)
                    clip = GetClipboardUnicode(FrmMain.hWnd)
                    clip = XToEsperanto(clip)
                    Clipboard.Clear
                    Call SetClipboardUnicode(FrmMain.hWnd, clip)
                    SendCtrlV
                End If
                
            Case "SSIGI_TONDUJON"
                If scriptOn Then
                    clip = GetClipboardUnicode(FrmMain.hWnd)
                    clip = XToEsperanto(clip)
                    Clipboard.Clear
                    Call SetClipboardUnicode(FrmMain.hWnd, clip)
                End If
                
            Case "XIGI_ELEKTITAN"
                If scriptOn Then
                    Clipboard.Clear
                    SendCtrlX
                    MySleep (100)
                    clip = GetClipboardUnicode(FrmMain.hWnd)
                    clip = EsperantoToX(clip)
                    Clipboard.Clear
                    Call SetClipboardUnicode(FrmMain.hWnd, clip)
                    SendCtrlV
                End If
                
            Case "XIGI_TONDUJON"
                If scriptOn Then
                    clip = GetClipboardUnicode(FrmMain.hWnd)
                    clip = EsperantoToX(clip)
                    Clipboard.Clear
                    Call SetClipboardUnicode(FrmMain.hWnd, clip)
                End If
            
            Case "HIGI_ELEKTITAN"
                If scriptOn Then
                    Clipboard.Clear
                    SendCtrlX
                    MySleep (100)
                    clip = GetClipboardUnicode(FrmMain.hWnd)
                    clip = EsperantoToH(clip)
                    Clipboard.Clear
                    Call SetClipboardUnicode(FrmMain.hWnd, clip)
                    SendCtrlV
                End If
            
            Case "HIGI_TONDUJON"
                If scriptOn Then
                    clip = GetClipboardUnicode(FrmMain.hWnd)
                    clip = EsperantoToH(clip)
                    Clipboard.Clear
                    Call SetClipboardUnicode(FrmMain.hWnd, clip)
                End If
            
            Case "HTMLIGI_ELEKTITAN"
                If scriptOn Then
                    Clipboard.Clear
                    SendCtrlX
                    MySleep (100)
                    clip = GetClipboardUnicode(FrmMain.hWnd)
                    clip = ToHTMLCodes(clip)
                    Clipboard.Clear
                    Call SetClipboardUnicode(FrmMain.hWnd, clip)
                    SendCtrlV
                End If
            
            Case "HTMLIGI_TONDUJON"
                If scriptOn Then
                    clip = GetClipboardUnicode(FrmMain.hWnd)
                    clip = ToHTMLCodes(clip)
                    Clipboard.Clear
                    Call SetClipboardUnicode(FrmMain.hWnd, clip)
                End If
            
            Case "UNIKODO"
                If scriptOn Then
                    FrmConfig.RBtnLatin3.Value = False
                    FrmConfig.RBtnUnicode.Value = True
                    SaveConfig
                End If
                
            Case "LATINA3"
                If scriptOn Then
                    FrmConfig.RBtnLatin3.Value = True
                    FrmConfig.RBtnUnicode.Value = False
                    SaveConfig
                End If
                        
            Case "ALGLUI"
                If scriptOn Then
                    If FrmConfig.ChkPaste.Value = 1 Then
                        FrmConfig.ChkPaste.Value = 0
                    Else
                        FrmConfig.ChkPaste.Value = 1
                    End If
                    SaveConfig
                End If
            
            Case "HTML"
                If scriptOn Then
                    If FrmConfig.ChkEntityCodes.Value = 1 Then
                        FrmConfig.ChkEntityCodes.Value = 0
                    Else
                        FrmConfig.ChkEntityCodes.Value = 1
                    End If
                    SaveConfig
                End If
            
            Case "KLAVKOMANDO"
                If scriptOn Then
                    Call FrmMain.SetTajpiEnabled(Not FrmMain.MenuActive.Checked, True)
                    russian = False
                    cyrillic = False
                End If
                
            Case "SKRIPTO"
                scriptOn = Not scriptOn
            
            
        End Select
    End If
    
End Sub

Public Function ProcessHotKey(vkCode As Long, key As String) As Boolean
    
    ProcessHotKey = False
    
    Dim i As Integer
    For i = LBound(script) To UBound(script)
    
        Dim line As String
        line = script(i)
        
        If line <> "" Then
                       
            Dim splits() As String
            splits = Split(line, "::", 2)
            
            If UBound(splits) > 0 Then
                        
                Dim hotkey As String
                Dim command As String
                Dim keyName As String
                hotkey = UCase$(splits(0))
                command = splits(1)
                keyName = hotkey
                keyName = Replace(keyName, "<", "")
                keyName = Replace(keyName, ">", "")
                keyName = Replace(keyName, "^", "")
                keyName = Replace(keyName, "!", "")
                keyName = Replace(keyName, "+", "")
                
                If UCase$(key) = keyName Or vkCode = VKFromKeyName(keyName) Then
                                            
                    Dim ctrl As Boolean: ctrl = False
                    Dim lCtrl As Boolean: lCtrl = False
                    Dim rCtrl As Boolean: rCtrl = False
                    Dim alt As Boolean: alt = False
                    Dim lAlt As Boolean: lAlt = False
                    Dim rAlt As Boolean: rAlt = False
                    Dim Shift As Boolean: Shift = False
                    Dim lShift As Boolean: lShift = False
                    Dim rShift As Boolean: rShift = False
                    
                    If InStr(hotkey, "<^") > 0 Then
                        lCtrl = True
                        hotkey = Replace(hotkey, "<^", "")
                    End If
                    If InStr(hotkey, ">^") > 0 Then
                        rCtrl = True
                        hotkey = Replace(hotkey, ">^", "")
                    End If
                    If InStr(hotkey, "^") > 0 Then
                        ctrl = True
                        hotkey = Replace(hotkey, "^", "")
                    End If
                    
                    If InStr(hotkey, "<!") > 0 Then
                        lAlt = True
                        hotkey = Replace(hotkey, "<!", "")
                    End If
                    If InStr(hotkey, ">!") > 0 Then
                        rAlt = True
                        hotkey = Replace(hotkey, ">!", "")
                    End If
                    If InStr(hotkey, "!") > 0 Then
                        alt = True
                        hotkey = Replace(hotkey, "!", "")
                    End If
                    
                    If InStr(hotkey, "<+") > 0 Then
                        lShift = True
                        hotkey = Replace(hotkey, "<+", "")
                    End If
                    If InStr(hotkey, ">+") > 0 Then
                        rShift = True
                        hotkey = Replace(hotkey, ">+", "")
                    End If
                    If InStr(hotkey, "+") > 0 Then
                        Shift = True
                        hotkey = Replace(hotkey, "+", "")
                    End If
                    
                    If ((lCtrlDown = lCtrl) Or (lCtrlDown And ctrl) Or altGrDown) And _
                       ((rCtrlDown = rCtrl) Or (rCtrlDown And ctrl)) And _
                       ((ctrlDown = ctrl) Or (ctrlDown And lCtrl) Or (ctrlDown And rCtrl) Or altGrDown) And _
                       ((lAltDown = lAlt) Or (lAltDown And alt)) And _
                       ((altGrDown = rAlt) Or (altGrDown And alt)) And _
                       ((altDown = alt) Or (altDown And lAlt) Or (altDown And rAlt)) And _
                       ((lShiftDown = lShift) Or (lShiftDown And Shift)) And _
                       ((rShiftDown = rShift) Or (rShiftDown And Shift)) And _
                       ((shiftDown = Shift) Or (shiftDown And lShift) Or (shiftDown And rShift)) Then
                        
                        If Left$(command, 1) = "_" Then
                        
                            ProcessCommand (command)
                        
                        Else
                            
                            Dim commands() As String
                            commands = Split(command, ",")
                            Dim J As Integer
                            For J = LBound(commands) To UBound(commands)
                                ProcessCommand (commands(J))
                            Next
                        
                        End If
                        ProcessHotKey = True
                        ClearBuffer
                    End If
                    
                
                End If
                
            End If
            
        End If
    
    Next

End Function
