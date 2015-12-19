Attribute VB_Name = "ModConfig"
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

Public Sub SaveConfig()
    
    On Error Resume Next
    Dim fileName As String
    fileName = GetDataPath() & "\Tajpi.ini"

    Call WritePrivateProfileString("Tajpi", "bRektaj", CStr(FrmConfig.ChkDirectKeys.Value), fileName)
    Call WritePrivateProfileString("Tajpi", "sRektaC", FrmConfig.TxtDirectKey(0).Text, fileName)
    Call WritePrivateProfileString("Tajpi", "sRektaG", FrmConfig.TxtDirectKey(1).Text, fileName)
    Call WritePrivateProfileString("Tajpi", "sRektaH", FrmConfig.TxtDirectKey(2).Text, fileName)
    Call WritePrivateProfileString("Tajpi", "sRektaJ", FrmConfig.TxtDirectKey(3).Text, fileName)
    Call WritePrivateProfileString("Tajpi", "sRektaS", FrmConfig.TxtDirectKey(4).Text, fileName)
    Call WritePrivateProfileString("Tajpi", "sRektaU", FrmConfig.TxtDirectKey(5).Text, fileName)
    Call WritePrivateProfileString("Tajpi", "bPrefiksoj", CStr(FrmConfig.ChkPrefixes.Value), fileName)
    Call WritePrivateProfileString("Tajpi", "sPrefiksoj", FrmConfig.TxtPrefixes.Text, fileName)
    Call WritePrivateProfileString("Tajpi", "bMalvidebligi", CStr(FrmConfig.ChkInvisibleSuffix.Value), fileName)
    Call WritePrivateProfileString("Tajpi", "bSufiksoj", CStr(FrmConfig.ChkSuffixes.Value), fileName)
    Call WritePrivateProfileString("Tajpi", "sSufiksoj", FrmConfig.TxtSuffixes.Text, fileName)
    Call WritePrivateProfileString("Tajpi", "bRipeto", CStr(FrmConfig.ChkSuffixesRepeat.Value), fileName)
    Call WritePrivateProfileString("Tajpi", "bAltGr", CStr(FrmConfig.ChkAltGr.Value), fileName)
    Call WritePrivateProfileString("Tajpi", "bAuxtomataAuEu", CStr(FrmConfig.ChkAutomaticAuEu.Value), fileName)
    Call WritePrivateProfileString("Tajpi", "bUnikodo", CStr(IIf(FrmConfig.RBtnUnicode.Value, "1", "0")), fileName)
    Call WritePrivateProfileString("Tajpi", "bLatin3", CStr(IIf(FrmConfig.RBtnLatin3.Value, "1", "0")), fileName)
    Call WritePrivateProfileString("Tajpi", "bAlgluu", CStr(FrmConfig.ChkPaste.Value), fileName)
    Call WritePrivateProfileString("Tajpi", "bHTML", CStr(FrmConfig.ChkEntityCodes.Value), fileName)
    Call WritePrivateProfileString("Tajpi", "bStartuAktiva", CStr(FrmConfig.ChkStartActive.Value), fileName)
    Call WritePrivateProfileString("Tajpi", "bAuxtomataStarto", CStr(FrmConfig.ChkAutomaticStart.Value), fileName)
    Call WritePrivateProfileString("Tajpi", "bCtrl", CStr(FrmConfig.ChkCtrl.Value), fileName)
    Call WritePrivateProfileString("Tajpi", "bAlt", CStr(FrmConfig.ChkAlt.Value), fileName)
    Call WritePrivateProfileString("Tajpi", "bShift", CStr(FrmConfig.ChkShift.Value), fileName)
    Call WritePrivateProfileString("Tajpi", "iKlavkomando", CStr(FrmConfig.CmbKeys.ListIndex), fileName)
    Call WritePrivateProfileString("Tajpi", "bShiftInsert", CStr(IIf(FrmClipboardMethod.RBtnShiftInsert.Value, "1", "0")), fileName)
    Call WritePrivateProfileString("Tajpi", "bCtrlV", CStr(IIf(FrmClipboardMethod.RBtnCtrlV.Value, "1", "0")), fileName)
    Call WritePrivateProfileString("Tajpi", "bRestore", CStr(FrmClipboardMethod.ChkRestore.Value), fileName)
    Call WritePrivateProfileString("Tajpi", "iDelay", FrmClipboardMethod.TxtDelay.Text, fileName)
    Call WritePrivateProfileString("Tajpi", "sLanguage", LANGUAGE, fileName)
    Call WritePrivateProfileString("Tajpi", "bPrefixW", CStr(FrmConfig.ChkW.Value), fileName)
    
    If FrmConfig.ChkAutomaticStart.Value = 1 Then
        regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "Tajpi", App.path & "\Tajpi.exe"
    Else
        regDelete_Sub_Key HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "Tajpi"
    End If

End Sub

Public Sub LoadConfig()

    On Error Resume Next
    Dim fileName As String
    fileName = GetDataPath() & "\Tajpi.ini"

    Dim Value As String
    Value = Space$(255)
    Dim Length As String

    Length = GetPrivateProfileString("Tajpi", "bRektaj", "0", Value, 255, fileName)
    If Length > 0 Then FrmConfig.ChkDirectKeys.Value = Val(Left$(Value, Length))

    Length = GetPrivateProfileString("Tajpi", "sRektaC", "X", Value, 255, fileName)
    If Length > 0 Then FrmConfig.TxtDirectKey(0).Text = Left$(Value, Length)

    Length = GetPrivateProfileString("Tajpi", "sRektaG", "Y", Value, 255, fileName)
    If Length > 0 Then FrmConfig.TxtDirectKey(1).Text = Left$(Value, Length)

    Length = GetPrivateProfileString("Tajpi", "sRektaH", "]", Value, 255, fileName)
    If Length > 0 Then FrmConfig.TxtDirectKey(2).Text = Left$(Value, Length)

    Length = GetPrivateProfileString("Tajpi", "sRektaJ", "[", Value, 255, fileName)
    If Length > 0 Then FrmConfig.TxtDirectKey(3).Text = Left$(Value, Length)

    Length = GetPrivateProfileString("Tajpi", "sRektaS", "Q", Value, 255, fileName)
    If Length > 0 Then FrmConfig.TxtDirectKey(4).Text = Left$(Value, Length)

    Length = GetPrivateProfileString("Tajpi", "sRektaU", "W", Value, 255, fileName)
    If Length > 0 Then FrmConfig.TxtDirectKey(5).Text = Left$(Value, Length)

    Length = GetPrivateProfileString("Tajpi", "bPrefiksoj", "0", Value, 255, fileName)
    If Length > 0 Then FrmConfig.ChkPrefixes.Value = Val(Left$(Value, Length))

    Length = GetPrivateProfileString("Tajpi", "sPrefiksoj", ";", Value, 255, fileName)
    If Length > 0 Then FrmConfig.TxtPrefixes.Text = Left$(Value, Length)
    
    Length = GetPrivateProfileString("Tajpi", "bMalvidebligi", "0", Value, 255, fileName)
    If Length > 0 Then FrmConfig.ChkInvisibleSuffix.Value = Val(Left$(Value, Length))

    Length = GetPrivateProfileString("Tajpi", "bSufiksoj", "1", Value, 255, fileName)
    If Length > 0 Then FrmConfig.ChkSuffixes.Value = Val(Left$(Value, Length))

    Length = GetPrivateProfileString("Tajpi", "sSufiksoj", "XH^", Value, 255, fileName)
    If Length > 0 Then FrmConfig.TxtSuffixes.Text = Left$(Value, Length)

    Length = GetPrivateProfileString("Tajpi", "bRipeto", "1", Value, 255, fileName)
    If Length > 0 Then FrmConfig.ChkSuffixesRepeat.Value = Val(Left$(Value, Length))
    
    Length = GetPrivateProfileString("Tajpi", "bAltGr", "0", Value, 255, fileName)
    If Length > 0 Then FrmConfig.ChkAltGr.Value = Val(Left$(Value, Length))

    Length = GetPrivateProfileString("Tajpi", "bAuxtomataAuEu", "0", Value, 255, fileName)
    If Length > 0 Then FrmConfig.ChkAutomaticAuEu.Value = Val(Left$(Value, Length))

    Length = GetPrivateProfileString("Tajpi", "bUnikodo", "1", Value, 255, fileName)
    If Length > 0 Then FrmConfig.RBtnUnicode.Value = Val(Left$(Value, Length))

    Length = GetPrivateProfileString("Tajpi", "bLatin3", "0", Value, 255, fileName)
    If Length > 0 Then FrmConfig.RBtnLatin3.Value = Val(Left$(Value, Length))

    Length = GetPrivateProfileString("Tajpi", "bAlgluu", "0", Value, 255, fileName)
    If Length > 0 Then FrmConfig.ChkPaste.Value = Val(Left$(Value, Length))

    Length = GetPrivateProfileString("Tajpi", "bHTML", "0", Value, 255, fileName)
    If Length > 0 Then FrmConfig.ChkEntityCodes.Value = Val(Left$(Value, Length))

    Length = GetPrivateProfileString("Tajpi", "bStartuAktiva", "1", Value, 255, fileName)
    If Length > 0 Then FrmConfig.ChkStartActive.Value = Val(Left$(Value, Length))

    Length = GetPrivateProfileString("Tajpi", "bAuxtomataStarto", "1", Value, 255, fileName)
    If Length > 0 Then FrmConfig.ChkAutomaticStart.Value = Val(Left$(Value, Length))

    Length = GetPrivateProfileString("Tajpi", "bCtrl", "1", Value, 255, fileName)
    If Length > 0 Then FrmConfig.ChkCtrl.Value = Val(Left$(Value, Length))

    Length = GetPrivateProfileString("Tajpi", "bAlt", "0", Value, 255, fileName)
    If Length > 0 Then FrmConfig.ChkAlt.Value = Val(Left$(Value, Length))

    Length = GetPrivateProfileString("Tajpi", "bShift", "0", Value, 255, fileName)
    If Length > 0 Then FrmConfig.ChkShift.Value = Val(Left$(Value, Length))

    Length = GetPrivateProfileString("Tajpi", "iKlavkomando", "1", Value, 255, fileName)
    If Length > 0 Then FrmConfig.CmbKeys.ListIndex = Val(Left$(Value, Length))
  
    Length = GetPrivateProfileString("Tajpi", "bShiftInsert", "1", Value, 255, fileName)
    If Length > 0 Then FrmClipboardMethod.RBtnShiftInsert.Value = Val(Left$(Value, Length))

    Length = GetPrivateProfileString("Tajpi", "bCtrlV", "0", Value, 255, fileName)
    If Length > 0 Then FrmClipboardMethod.RBtnCtrlV.Value = Val(Left$(Value, Length))

    Length = GetPrivateProfileString("Tajpi", "bRestore", "1", Value, 255, fileName)
    If Length > 0 Then FrmClipboardMethod.ChkRestore.Value = Val(Left$(Value, Length))
    SetEnabled FrmClipboardMethod.ChkRestore.Value, FrmClipboardMethod.TxtDelay

    Length = GetPrivateProfileString("Tajpi", "iDelay", "500", Value, 255, fileName)
    If Length > 0 Then FrmClipboardMethod.TxtDelay.Text = Left$(Value, Length)
    FrmMain.Timer1.Interval = Val(Left$(Value, Length))
    
    Length = GetPrivateProfileString("Tajpi", "sLanguage", "Esperanto", Value, 255, fileName)
    If Length > 0 Then
        LANGUAGE = Left$(Value, Length)
    Else
        LANGUAGE = "Esperanto"
    End If
    
    Length = GetPrivateProfileString("Tajpi", "bPrefixW", "0", Value, 255, fileName)
    If Length > 0 Then FrmConfig.ChkW.Value = Val(Left$(Value, Length))
    

End Sub
