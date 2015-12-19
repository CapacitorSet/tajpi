Attribute VB_Name = "ModUtil"
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

Const accentable As String = "cCgGhHjJsSuU"
Const accentableW As String = "cCgGhHjJsSwW"
Public ACTIVE_TOOLTIP As String
Public INACTIVE_TOOLTIP As String
Public PASTE_MESSAGE_1 As String
Public PASTE_MESSAGE_2 As String
Public LANGUAGE As String
Public vkhash As New Collection

Public Function RemoveBOM(ByVal str As String, Optional ByVal utf8 As Boolean = True) As String

    RemoveBOM = str
    
    If utf8 Then
        If AscW(str) = &HFEFF Then
            RemoveBOM = Mid$(str, 2)
        End If
    Else
        If Left$(str, 2) = Chr(&HFF) & Chr(&HFE) Then
            RemoveBOM = Mid$(str, 3)
        ElseIf Left$(str, 3) = Chr(&HEF) & Chr(&HBB) & Chr(&HBF) Then
            RemoveBOM = Mid$(str, 4)
        End If
    End If

End Function

Public Function StrConvFromUTF8(Text As String) As String
    
    ' get length
    Dim lngLen As Long, lngPtr As Long: lngLen = LenB(Text)
    ' has any?
    If lngLen Then
        ' create a BSTR over twice that length
        lngPtr = SysAllocStringLen(0, lngLen * 1.25)
        ' place it in output variable
        PutMem4 VarPtr(StrConvFromUTF8), lngPtr
        ' convert & get output length
        lngLen = MultiByteToWideChar(65001, 0, ByVal StrPtr(Text), lngLen, ByVal lngPtr, LenB(StrConvFromUTF8))
        ' resize the buffer
        StrConvFromUTF8 = Left$(StrConvFromUTF8, lngLen)
    End If

End Function

Public Function StrConvToUTF8(Text As String) As String
    
    ' get length
    Dim lngLen As Long, lngPtr As Long: lngLen = LenB(Text)
    ' has any?
    If lngLen Then
        ' create a BSTR over twice that length
        lngPtr = SysAllocStringLen(0, lngLen * 1.25)
        ' place it in output variable
        PutMem4 VarPtr(StrConvToUTF8), lngPtr
        ' convert & get output length
        lngLen = WideCharToMultiByte(65001, 0, ByVal StrPtr(Text), Len(Text), ByVal lngPtr, LenB(StrConvToUTF8), ByVal 0&, ByVal 0&)
        ' resize the buffer
        StrConvToUTF8 = LeftB$(StrConvToUTF8, lngLen)
    End If

End Function

Public Sub ShowHelp(Optional ByVal lang As String)
    
    Dim file As String
    Dim RC As Long
    
    If lang = "" Then
        lang = LANGUAGE
    End If
    
    If lang = "English" Then
        file = App.path & "\Helpo (angla).chm"
    Else
        file = App.path & "\Helpo.chm"
    End If

    RC = ShellExecute(0, "Open", file, 0&, 0&, SW_SHOWDEFAULT)

End Sub

Public Sub SetEnabled(ByVal enabled As Boolean, ByRef ctrl As Control)

    ctrl.enabled = enabled
    If TypeName(ctrl) = "TextBox" Then
        If enabled Then
            ctrl.BackColor = GetSysColor(COLOR_WINDOW)
        Else
            ctrl.BackColor = GetSysColor(COLOR_BTNFACE)
        End If
    End If

End Sub

Public Function FileExists(ByVal strPath As String) As Boolean
    
    If Dir$(strPath) <> "" Then
        FileExists = True
    Else
        FileExists = False
    End If

End Function

Public Function GetDataPath() As String
    
    Dim sPath As String
    Dim RetVal As Long
        
    sPath = String$(260, 0)
    RetVal = SHGetFolderPath(0, CSIDL_LOCAL_APPDATA, 0, 0, sPath)
    
    Select Case RetVal
        Case 0
            sPath = Left$(sPath, InStr(1, sPath, Chr(0)) - 1)
        Case 1, E_INVALIDARG
            sPath = App.path
    End Select
    
    GetDataPath = sPath

End Function

Public Function IsPrefixKey(ByVal key As String) As Boolean
    If key <> "" Then
        IsPrefixKey = (InStr(1, FrmConfig.TxtPrefixes.Text, key) > 0)
    End If
End Function

Public Function IsSuffixKey(ByVal key As String) As Boolean
    If key <> "" Then
        IsSuffixKey = (InStr(1, FrmConfig.TxtSuffixes.Text, key) > 0)
    End If
End Function

Public Function IsDirectKey(ByVal key As String) As Boolean
    If key <> "" Then
        Dim i As Integer
        For i = FrmConfig.TxtDirectKey.LBound To FrmConfig.TxtDirectKey.UBound
            If FrmConfig.TxtDirectKey(i).Text <> "" Then
                If key = FrmConfig.TxtDirectKey(i).Text Then
                    IsDirectKey = True
                End If
            End If
        Next
    End If
End Function

Public Function IsAccentable(ByVal key As String) As Boolean
    
    If key <> "" Then
        IsAccentable = (InStr(1, IIf(FrmConfig.ChkW.Value, accentableW, accentable), key) > 0)
    End If

End Function

Public Function ToHTMLCodes(str As String) As String

    Dim i As Integer
    For i = 1 To Len(str)
        ToHTMLCodes = ToHTMLCodes & "&#" & AscW(Mid$(str, i, 1)) & ";"
    Next

End Function

Public Function XToEsperanto(str As String) As String

    XToEsperanto = str
    XToEsperanto = Replace(XToEsperanto, "cx", EOKey("c"))
    XToEsperanto = Replace(XToEsperanto, "Cx", EOKey("C"))
    XToEsperanto = Replace(XToEsperanto, "gx", EOKey("g"))
    XToEsperanto = Replace(XToEsperanto, "Gx", EOKey("G"))
    XToEsperanto = Replace(XToEsperanto, "hx", EOKey("h"))
    XToEsperanto = Replace(XToEsperanto, "Hx", EOKey("H"))
    XToEsperanto = Replace(XToEsperanto, "jx", EOKey("j"))
    XToEsperanto = Replace(XToEsperanto, "Jx", EOKey("J"))
    XToEsperanto = Replace(XToEsperanto, "sx", EOKey("s"))
    XToEsperanto = Replace(XToEsperanto, "Sx", EOKey("S"))
    XToEsperanto = Replace(XToEsperanto, "ux", EOKey("u"))
    XToEsperanto = Replace(XToEsperanto, "Ux", EOKey("U"))
    
    XToEsperanto = Replace(XToEsperanto, "cX", EOKey("c"))
    XToEsperanto = Replace(XToEsperanto, "CX", EOKey("C"))
    XToEsperanto = Replace(XToEsperanto, "gX", EOKey("g"))
    XToEsperanto = Replace(XToEsperanto, "GX", EOKey("G"))
    XToEsperanto = Replace(XToEsperanto, "hX", EOKey("h"))
    XToEsperanto = Replace(XToEsperanto, "HX", EOKey("H"))
    XToEsperanto = Replace(XToEsperanto, "jX", EOKey("j"))
    XToEsperanto = Replace(XToEsperanto, "JX", EOKey("J"))
    XToEsperanto = Replace(XToEsperanto, "sX", EOKey("s"))
    XToEsperanto = Replace(XToEsperanto, "SX", EOKey("S"))
    XToEsperanto = Replace(XToEsperanto, "uX", EOKey("u"))
    XToEsperanto = Replace(XToEsperanto, "UX", EOKey("U"))

End Function

Public Function EsperantoToX(str As String) As String

    EsperantoToX = str
    EsperantoToX = Replace(EsperantoToX, EOKey("c"), "cx")
    EsperantoToX = Replace(EsperantoToX, EOKey("C"), "Cx")
    EsperantoToX = Replace(EsperantoToX, EOKey("g"), "gx")
    EsperantoToX = Replace(EsperantoToX, EOKey("G"), "Gx")
    EsperantoToX = Replace(EsperantoToX, EOKey("h"), "hx")
    EsperantoToX = Replace(EsperantoToX, EOKey("H"), "Hx")
    EsperantoToX = Replace(EsperantoToX, EOKey("j"), "jx")
    EsperantoToX = Replace(EsperantoToX, EOKey("J"), "Jx")
    EsperantoToX = Replace(EsperantoToX, EOKey("s"), "sx")
    EsperantoToX = Replace(EsperantoToX, EOKey("S"), "Sx")
    EsperantoToX = Replace(EsperantoToX, EOKey("u"), "ux")
    EsperantoToX = Replace(EsperantoToX, EOKey("U"), "Ux")

End Function

Public Function EsperantoToH(str As String) As String

    EsperantoToH = str
    EsperantoToH = Replace(EsperantoToH, EOKey("c"), "ch")
    EsperantoToH = Replace(EsperantoToH, EOKey("C"), "Ch")
    EsperantoToH = Replace(EsperantoToH, EOKey("g"), "gh")
    EsperantoToH = Replace(EsperantoToH, EOKey("G"), "Gh")
    EsperantoToH = Replace(EsperantoToH, EOKey("h"), "hh")
    EsperantoToH = Replace(EsperantoToH, EOKey("H"), "Hh")
    EsperantoToH = Replace(EsperantoToH, EOKey("j"), "jh")
    EsperantoToH = Replace(EsperantoToH, EOKey("J"), "Jh")
    EsperantoToH = Replace(EsperantoToH, EOKey("s"), "sh")
    EsperantoToH = Replace(EsperantoToH, EOKey("S"), "Sh")
    EsperantoToH = Replace(EsperantoToH, EOKey("u"), "u")
    EsperantoToH = Replace(EsperantoToH, EOKey("U"), "U")

End Function

Public Sub MySleep(milliseconds As Integer)
    
    Dim i As Integer
    For i = 0 To milliseconds
        Sleep (1)
        DoEvents
    Next

End Sub

Public Function LoadRes(i As Integer) As String
    
    LoadRes = XToEsperanto(LoadResString(i))
    
End Function

Public Sub SetLanguage(ByVal lang As String, Optional ByVal events As Boolean = False)

    On Error GoTo Finish
    
    LANGUAGE = lang
        
    Dim i As Integer
    If lang = "English" Then
        i = 201
    Else
        i = 401
    End If
  
    FrmAbout.Caption = LoadRes(i): i = i + 1
    FrmAbout.Label2.Caption = LoadRes(i): i = i + 1
    FrmAbout.Lbl1.Caption = ""
    FrmAbout.Lbl1.Caption = LoadRes(i): i = i + 1
    FrmAbout.Lbl2.Caption = ""
    FrmAbout.Lbl2.Caption = LoadRes(i): i = i + 1
    FrmAbout.Lbl3.Caption = ""
    FrmAbout.Lbl3.Caption = LoadRes(i): i = i + 1
    FrmAbout.BtnOK.Caption = LoadRes(i): i = i + 1
    
    FrmClipboardMethod.Caption = LoadRes(i): i = i + 1
    FrmClipboardMethod.Label1.Caption = LoadRes(i): i = i + 1
    FrmClipboardMethod.RBtnShiftInsert.Caption = LoadRes(i): i = i + 1
    FrmClipboardMethod.ChkRestore.Caption = LoadRes(i): i = i + 1
    FrmClipboardMethod.Label2.Caption = LoadRes(i): i = i + 1
    FrmClipboardMethod.Label3.Caption = LoadRes(i): i = i + 1
    FrmClipboardMethod.BtnOK.Caption = LoadRes(i): i = i + 1
    
    FrmConfig.Caption = LoadRes(i): i = i + 1
    FrmConfig.Frame1.Caption = LoadRes(i): i = i + 1
    FrmConfig.ChkDirectKeys.Caption = LoadRes(i): i = i + 1
    FrmConfig.ChkPrefixes.Caption = LoadRes(i): i = i + 1
    FrmConfig.ChkInvisibleSuffix.Caption = LoadRes(i): i = i + 1
    FrmConfig.ChkSuffixes.Caption = LoadRes(i): i = i + 1
    FrmConfig.ChkSuffixesRepeat.Caption = LoadRes(i): i = i + 1
    FrmConfig.ChkAltGr.Caption = LoadRes(i): i = i + 1
    FrmConfig.ChkAutomaticAuEu.Caption = LoadRes(i)
    FrmConfig.LblAutomaticAuEu.Caption = ""
    FrmConfig.LblAutomaticAuEu.Caption = LoadRes(i): i = i + 1
    FrmConfig.Frame2.Caption = LoadRes(i): i = i + 1
    FrmConfig.RBtnUnicode.Caption = LoadRes(i): i = i + 1
    FrmConfig.RBtnLatin3.Caption = LoadRes(i): i = i + 1
    FrmConfig.ChkPaste.Caption = LoadRes(i): i = i + 1
    FrmConfig.ChkEntityCodes.Caption = LoadRes(i): i = i + 1
    FrmConfig.Frame3.Caption = LoadRes(i): i = i + 1
    FrmConfig.ChkStartActive.Caption = LoadRes(i): i = i + 1
    FrmConfig.ChkAutomaticStart.Caption = LoadRes(i)
    FrmConfig.LblAutomaticStart.Caption = ""
    FrmConfig.LblAutomaticStart.Caption = LoadRes(i): i = i + 1
    FrmConfig.Frame4.Caption = LoadRes(i): i = i + 1
    FrmConfig.BtnOK.Caption = LoadRes(i): i = i + 1
    FrmConfig.BtnCancel.Caption = LoadRes(i): i = i + 1
    FrmConfig.BtnHelp.Caption = LoadRes(i): i = i + 1
    FrmConfig.Frame5.Caption = LoadRes(i): i = i + 1
    FrmConfig.ChkW.Caption = LoadRes(i)
    FrmConfig.LblW.Caption = ""
    FrmConfig.LblW.Caption = LoadRes(i): i = i + 1
        
    FrmMain.MenuActive.Caption = LoadRes(i): i = i + 1
    FrmMain.MenuConfig.Caption = LoadRes(i): i = i + 1
    FrmMain.MenuHelp.Caption = LoadRes(i): i = i + 1
    FrmMain.MenuAbout.Caption = LoadRes(i): i = i + 1
    FrmMain.MenuExit.Caption = LoadRes(i): i = i + 1
    
    ACTIVE_TOOLTIP = LoadRes(i): i = i + 1
    INACTIVE_TOOLTIP = LoadRes(i): i = i + 1
    PASTE_MESSAGE_1 = LoadRes(i): i = i + 1
    PASTE_MESSAGE_2 = LoadRes(i): i = i + 1
    
    If FrmMain.MenuActive.Checked Then
        FrmMain.SetEnabledIcon
    Else
        FrmMain.SetDisabledIcon
    End If
    
Finish:

    If events Then
        DoEvents
    End If

End Sub

Public Function VKFromKeyName(keyName As String) As String
    
    On Error Resume Next
    VKFromKeyName = 0
    VKFromKeyName = vkhash(Replace(UCase$(keyName), " ", ""))
    
End Function

Public Sub BuildVKHash()
    
    vkhash.Add VK_SPACE, "SPACE"
    vkhash.Add VK_TAB, "TAB"
    vkhash.Add VK_RETURN, "ENTER"
    vkhash.Add VK_ESCAPE, "ESCAPE"
    vkhash.Add VK_SCROLL, "SCROLLLOCK"
    vkhash.Add VK_NUMLOCK, "NUMLOCK"
    vkhash.Add VK_CAPITAL, "CAPSLOCK"
    vkhash.Add VK_PAUSE, "PAUSE"
    vkhash.Add VK_SNAPSHOT, "PRINTSCR"
    vkhash.Add VK_INSERT, "INSERT"
    vkhash.Add VK_HOME, "HOME"
    vkhash.Add VK_DELETE, "DELETE"
    vkhash.Add VK_END, "END"
    vkhash.Add VK_NEXT, "PAGEUP"
    vkhash.Add VK_PRIOR, "PAGEDOWN"
    vkhash.Add VK_UP, "ARROWUP"
    vkhash.Add VK_DOWN, "ARROWDOWN"
    vkhash.Add VK_LEFT, "ARROWLEFT"
    vkhash.Add VK_RIGHT, "ARROWRIGHT"
    vkhash.Add VK_LWIN, "WINL"
    vkhash.Add VK_RWIN, "WINR"
    vkhash.Add VK_APPS, "MENU"
    
    vkhash.Add VK_KEY_A, "A"
    vkhash.Add VK_KEY_B, "B"
    vkhash.Add VK_KEY_C, "C"
    vkhash.Add VK_KEY_D, "D"
    vkhash.Add VK_KEY_E, "E"
    vkhash.Add VK_KEY_F, "F"
    vkhash.Add VK_KEY_G, "G"
    vkhash.Add VK_KEY_H, "H"
    vkhash.Add VK_KEY_I, "I"
    vkhash.Add VK_KEY_J, "J"
    vkhash.Add VK_KEY_K, "K"
    vkhash.Add VK_KEY_L, "L"
    vkhash.Add VK_KEY_M, "M"
    vkhash.Add VK_KEY_N, "N"
    vkhash.Add VK_KEY_O, "O"
    vkhash.Add VK_KEY_P, "P"
    vkhash.Add VK_KEY_Q, "Q"
    vkhash.Add VK_KEY_R, "R"
    vkhash.Add VK_KEY_S, "S"
    vkhash.Add VK_KEY_T, "T"
    vkhash.Add VK_KEY_U, "U"
    vkhash.Add VK_KEY_V, "V"
    vkhash.Add VK_KEY_W, "W"
    vkhash.Add VK_KEY_X, "X"
    vkhash.Add VK_KEY_Y, "Y"
    vkhash.Add VK_KEY_Z, "Z"
    
    vkhash.Add VK_KEY_1, "1"
    vkhash.Add VK_KEY_2, "2"
    vkhash.Add VK_KEY_3, "3"
    vkhash.Add VK_KEY_4, "4"
    vkhash.Add VK_KEY_5, "5"
    vkhash.Add VK_KEY_6, "6"
    vkhash.Add VK_KEY_7, "7"
    vkhash.Add VK_KEY_8, "8"
    vkhash.Add VK_KEY_9, "9"
    vkhash.Add VK_KEY_0, "0"
    
    vkhash.Add VK_F1, "F1"
    vkhash.Add VK_F2, "F2"
    vkhash.Add VK_F3, "F3"
    vkhash.Add VK_F4, "F4"
    vkhash.Add VK_F5, "F5"
    vkhash.Add VK_F6, "F6"
    vkhash.Add VK_F7, "F7"
    vkhash.Add VK_F8, "F8"
    vkhash.Add VK_F9, "F9"
    vkhash.Add VK_F10, "F10"
    vkhash.Add VK_F11, "F11"
    vkhash.Add VK_F12, "F12"
    vkhash.Add VK_F13, "F13"
    vkhash.Add VK_F14, "F14"
    vkhash.Add VK_F15, "F15"
    vkhash.Add VK_F16, "F16"
    vkhash.Add VK_F17, "F17"
    vkhash.Add VK_F18, "F18"
    vkhash.Add VK_F19, "F19"
    vkhash.Add VK_F20, "F20"
    vkhash.Add VK_F21, "F21"
    vkhash.Add VK_F22, "F22"
    vkhash.Add VK_F23, "F23"
    vkhash.Add VK_F24, "F24"
    
    vkhash.Add VK_NUMPAD1, "NUMPAD1"
    vkhash.Add VK_NUMPAD2, "NUMPAD2"
    vkhash.Add VK_NUMPAD3, "NUMPAD3"
    vkhash.Add VK_NUMPAD4, "NUMPAD4"
    vkhash.Add VK_NUMPAD5, "NUMPAD5"
    vkhash.Add VK_NUMPAD6, "NUMPAD6"
    vkhash.Add VK_NUMPAD7, "NUMPAD7"
    vkhash.Add VK_NUMPAD8, "NUMPAD8"
    vkhash.Add VK_NUMPAD9, "NUMPAD9"
    vkhash.Add VK_NUMPAD0, "NUMPAD0"
    vkhash.Add VK_ADD, "NUMPAD+"
    vkhash.Add VK_DECIMAL, "NUMPAD."
    vkhash.Add VK_DIVIDE, "NUMPAD/"
    vkhash.Add VK_MULTIPLY, "NUMPAD*"
    vkhash.Add VK_SUBTRACT, "NUMPAD-"

End Sub

Public Function CyrillicKey(ByVal char As String, Optional ByVal entityCode = False) As String
    
    CyrillicKey = ""
    
    Select Case char
        Case "a"
            CyrillicKey = IIf(entityCode, "&#1072;", ChrW$(1072))
        Case "A"
            CyrillicKey = IIf(entityCode, "&#1040;", ChrW$(1040))
        Case "b"
            CyrillicKey = IIf(entityCode, "&#1073;", ChrW$(1073))
        Case "B"
            CyrillicKey = IIf(entityCode, "&#1041;", ChrW$(1041))
        Case "c"
            CyrillicKey = IIf(entityCode, "&#1094;", ChrW$(1094))
        Case "C"
            CyrillicKey = IIf(entityCode, "&#1062;", ChrW$(1062))
        Case "d"
            CyrillicKey = IIf(entityCode, "&#1076;", ChrW$(1076))
        Case "D"
            CyrillicKey = IIf(entityCode, "&#1044;", ChrW$(1044))
        Case "e"
            CyrillicKey = IIf(entityCode, "&#1077;", ChrW$(1077))
        Case "E"
            CyrillicKey = IIf(entityCode, "&#1045;", ChrW$(1045))
        Case "f"
            CyrillicKey = IIf(entityCode, "&#1092;", ChrW$(1092))
        Case "F"
            CyrillicKey = IIf(entityCode, "&#1060;", ChrW$(1060))
        Case "g"
            CyrillicKey = IIf(entityCode, "&#1075;", ChrW$(1075))
        Case "G"
            CyrillicKey = IIf(entityCode, "&#1043;", ChrW$(1043))
        Case "h"
            CyrillicKey = IIf(entityCode, "&#1093;", ChrW$(1093))
        Case "H"
            CyrillicKey = IIf(entityCode, "&#1061;", ChrW$(1061))
        Case "i"
            CyrillicKey = IIf(entityCode, "&#1080;", ChrW$(1080))
        Case "I"
            CyrillicKey = IIf(entityCode, "&#1048;", ChrW$(1048))
        Case "j"
            CyrillicKey = IIf(entityCode, "&#1081;", ChrW$(1081))
        Case "J"
            CyrillicKey = IIf(entityCode, "&#1049;", ChrW$(1049))
        Case "k"
            CyrillicKey = IIf(entityCode, "&#1082;", ChrW$(1082))
        Case "K"
            CyrillicKey = IIf(entityCode, "&#1050;", ChrW$(1050))
        Case "l"
            CyrillicKey = IIf(entityCode, "&#1083;", ChrW$(1083))
        Case "L"
            CyrillicKey = IIf(entityCode, "&#1051;", ChrW$(1051))
        Case "m"
            CyrillicKey = IIf(entityCode, "&#1084;", ChrW$(1084))
        Case "M"
            CyrillicKey = IIf(entityCode, "&#1052;", ChrW$(1052))
        Case "n"
            CyrillicKey = IIf(entityCode, "&#1085;", ChrW$(1085))
        Case "N"
            CyrillicKey = IIf(entityCode, "&#1053;", ChrW$(1053))
        Case "o"
            CyrillicKey = IIf(entityCode, "&#1086;", ChrW$(1086))
        Case "O"
            CyrillicKey = IIf(entityCode, "&#1054;", ChrW$(1054))
        Case "p"
            CyrillicKey = IIf(entityCode, "&#1087;", ChrW$(1087))
        Case "P"
            CyrillicKey = IIf(entityCode, "&#1055;", ChrW$(1055))
        Case "r"
            CyrillicKey = IIf(entityCode, "&#1088;", ChrW$(1088))
        Case "R"
            CyrillicKey = IIf(entityCode, "&#1056;", ChrW$(1056))
        Case "s"
            CyrillicKey = IIf(entityCode, "&#1089;", ChrW$(1089))
        Case "S"
            CyrillicKey = IIf(entityCode, "&#1057;", ChrW$(1057))
        Case "t"
            CyrillicKey = IIf(entityCode, "&#1090;", ChrW$(1090))
        Case "T"
            CyrillicKey = IIf(entityCode, "&#1058;", ChrW$(1058))
        Case "u"
            CyrillicKey = IIf(entityCode, "&#1091;", ChrW$(1091))
        Case "U"
            CyrillicKey = IIf(entityCode, "&#1059;", ChrW$(1059))
        Case "v"
            CyrillicKey = IIf(entityCode, "&#1074;", ChrW$(1074))
        Case "V"
            CyrillicKey = IIf(entityCode, "&#1042;", ChrW$(1042))
        Case "z"
            CyrillicKey = IIf(entityCode, "&#1079;", ChrW$(1079))
        Case "Z"
            CyrillicKey = IIf(entityCode, "&#1047;", ChrW$(1047))
        
            
    End Select

End Function

Public Function EOKey(ByVal char As String, Optional ByVal unicode As Boolean = True, Optional ByVal entityCode = False, Optional ByVal cyrillic As Boolean = False) As String

    EOKey = ""

    If cyrillic Then
        
        Select Case char
            Case "c"
                EOKey = IIf(entityCode, "&#1095;", ChrW$(1095))
            Case "C"
                EOKey = IIf(entityCode, "&#1063;", ChrW$(1063))
            Case "g"
                If russian Then
                    EOKey = IIf(entityCode, "&#1076;&#1078;", ChrW$(1076) + ChrW$(1078))
                Else
                    EOKey = IIf(entityCode, "&#1107;", ChrW$(1107))
                End If
            Case "G"
                If russian Then
                    EOKey = IIf(entityCode, "&#1044;&#1046;", ChrW$(1044) + ChrW$(1046))
                Else
                    EOKey = IIf(entityCode, "&#1027;", ChrW$(1027))
                End If
            Case "h"
                If russian Then
                    EOKey = IIf(entityCode, "&#1093;", ChrW$(1093))
                Else
                    EOKey = IIf(entityCode, "&#1115;", ChrW$(1115))
                End If
            Case "H"
                If russian Then
                    EOKey = IIf(entityCode, "&#1061;", ChrW$(1061))
                Else
                    EOKey = IIf(entityCode, "&#1035;", ChrW$(1035))
                End If
            Case "j"
                EOKey = IIf(entityCode, "&#1078;", ChrW$(1078))
            Case "J"
                EOKey = IIf(entityCode, "&#1046;", ChrW$(1046))
            Case "s"
                EOKey = IIf(entityCode, "&#1096;", ChrW$(1096))
            Case "S"
                EOKey = IIf(entityCode, "&#1064;", ChrW$(1064))
            Case "u"
                If russian Then
                    EOKey = IIf(entityCode, "&#1091;", ChrW$(1091))
                Else
                    EOKey = IIf(entityCode, "&#1118;", ChrW$(1118))
                End If
            Case "U"
                If russian Then
                    EOKey = IIf(entityCode, "&#1059;", ChrW$(1059))
                Else
                    EOKey = IIf(entityCode, "&#1038;", ChrW$(1038))
                End If
            Case "w"
                If russian Then
                    EOKey = IIf(entityCode, "&#1091;", ChrW$(1091))
                Else
                    EOKey = IIf(entityCode, "&#1118;", ChrW$(1118))
                End If
            Case "W"
                If russian Then
                    EOKey = IIf(entityCode, "&#1059;", ChrW$(1059))
                Else
                    EOKey = IIf(entityCode, "&#1038;", ChrW$(1038))
                End If
        End Select
    
    Else

        If unicode Then
    
            Select Case char
                Case "c"
                    EOKey = IIf(entityCode, "&#265;", ChrW$(265))
                Case "C"
                    EOKey = IIf(entityCode, "&#264;", ChrW$(264))
                Case "g"
                    EOKey = IIf(entityCode, "&#285;", ChrW$(285))
                Case "G"
                    EOKey = IIf(entityCode, "&#284;", ChrW$(284))
                Case "h"
                    EOKey = IIf(entityCode, "&#293;", ChrW$(293))
                Case "H"
                    EOKey = IIf(entityCode, "&#292;", ChrW$(292))
                Case "j"
                    EOKey = IIf(entityCode, "&#309;", ChrW$(309))
                Case "J"
                    EOKey = IIf(entityCode, "&#308;", ChrW$(308))
                Case "s"
                    EOKey = IIf(entityCode, "&#349;", ChrW$(349))
                Case "S"
                    EOKey = IIf(entityCode, "&#348;", ChrW$(348))
                Case "u"
                    EOKey = IIf(entityCode, "&#365;", ChrW$(365))
                Case "U"
                    EOKey = IIf(entityCode, "&#364;", ChrW$(364))
                Case "w"
                    EOKey = IIf(entityCode, "&#365;", ChrW$(365))
                Case "W"
                    EOKey = IIf(entityCode, "&#364;", ChrW$(364))
            End Select
    
        Else
    
            Select Case char
                Case "c"
                    EOKey = IIf(entityCode, "&#230;", ChrW$(230))
                Case "C"
                    EOKey = IIf(entityCode, "&#198;", ChrW$(198))
                Case "g"
                    EOKey = IIf(entityCode, "&#248;", ChrW$(248))
                Case "G"
                    EOKey = IIf(entityCode, "&#216;", ChrW$(216))
                Case "h"
                    EOKey = IIf(entityCode, "&#182;", ChrW$(182))
                Case "H"
                    EOKey = IIf(entityCode, "&#166;", ChrW$(166))
                Case "j"
                    EOKey = IIf(entityCode, "&#188;", ChrW$(188))
                Case "J"
                    EOKey = IIf(entityCode, "&#172;", ChrW$(172))
                Case "s"
                    EOKey = IIf(entityCode, "&#254;", ChrW$(254))
                Case "S"
                    EOKey = IIf(entityCode, "&#222;", ChrW$(222))
                Case "u"
                    EOKey = IIf(entityCode, "&#253;", ChrW$(253))
                Case "U"
                    EOKey = IIf(entityCode, "&#221;", ChrW$(221))
                Case "w"
                    EOKey = IIf(entityCode, "&#253;", ChrW$(253))
                Case "W"
                    EOKey = IIf(entityCode, "&#221;", ChrW$(221))
            End Select
    
        End If
    
    End If
    
End Function
