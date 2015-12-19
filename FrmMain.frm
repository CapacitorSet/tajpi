VERSION 5.00
Begin VB.Form FrmMain 
   ClientHeight    =   660
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   1740
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   660
   ScaleWidth      =   1740
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   75
   End
   Begin VB.Image ImgDisabledIcon 
      Height          =   240
      Left            =   120
      Picture         =   "FrmMain.frx":0E42
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu MenuTajpi 
      Caption         =   "Tajpi"
      Begin VB.Menu MenuActive 
         Caption         =   "&Aktiva"
      End
      Begin VB.Menu MenuConfig 
         Caption         =   "Ago&rdi"
      End
      Begin VB.Menu MenuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuHelp 
         Caption         =   "H&elpo"
      End
      Begin VB.Menu MenuAbout 
         Caption         =   "&Pri Tajpi"
      End
      Begin VB.Menu MenuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuExit 
         Caption         =   "&Halti"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Private Sub Form_Load()
   
    If App.PrevInstance Then
        Dim handle As Long
        handle = FindWindow(vbNullString, "Tajpi Cxefa Fenestro")
        If handle <> 0 Then
            Call PostMessage(handle, WM_CLOSE, 0&, 0&)
        End If
    End If
    
    Me.Caption = "Tajpi Cxefa Fenestro"

    BuildVKHash
    FrmConfig.Prepare
    AddIconToTray
    InstallHook

    Dim iniFile As String
    iniFile = GetDataPath() & "\Tajpi.ini"
    
    LoadScript
    LoadConfig
    SetLanguage (LANGUAGE)
    SetTajpiEnabled (FrmConfig.ChkStartActive.Value = 1)
    
    If Not FileExists(iniFile) Then

        SaveConfig
        FrmConfig.Visible = True
        MsgBox "Bonvolu agordi Tajpi-on por unua uzo.", vbOKOnly & vbInformation, "Tajpi"

    End If
        
    SetForegroundWindow Me.hWnd
    PostMessage Me.hWnd, 0, 0&, 0&
   
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim Message As Long
       
    If Me.ScaleMode = vbPixels Then
        Message = X
    Else
        Message = X / Screen.TwipsPerPixelX
    End If
    
    Select Case Message
        
        Case WM_LBUTTONDBLCLK
            Call SetTajpiEnabled(Not MenuActive.Checked, True)
            
        Case WM_RBUTTONDOWN
            SetForegroundWindow Me.hWnd
            Call Me.PopupMenu(Me.MenuTajpi, , , , Me.MenuActive)
            PostMessage Me.hWnd, 0, 0&, 0&
            
    End Select

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    UnloadMe
    
End Sub

Public Sub UnloadMe()

    Unload FrmAbout
    Unload FrmConfig
    Unload FrmClipboardMethod
    DeleteIconFromTray
    UninstallHook

End Sub

Public Sub SetTajpiEnabled(ByVal enabled As Boolean, Optional ByVal beep As Boolean = False)

    If enabled Then
        
        ClearBuffer
        SetEnabledIcon
        If beep Then
            Call MessageBeep(MB_ICONEXCLAMATION)
        End If
                  
    Else
        SetDisabledIcon
        If beep Then
            Call MessageBeep(MB_ICONHAND)
        End If
    End If
    
    MenuActive.Checked = enabled
    
End Sub

Public Sub SetDisabledIcon()
    
    On Error Resume Next
    data.Icon = ImgDisabledIcon.Picture
    data.Tip = INACTIVE_TOOLTIP & vbNullChar
    Call Shell_NotifyIcon(ModifyIcon, data)

End Sub

Public Sub SetEnabledIcon()
    
    On Error Resume Next
    data.Icon = Icon
    data.Tip = ACTIVE_TOOLTIP & vbNullChar
    Call Shell_NotifyIcon(ModifyIcon, data)

End Sub

Public Sub DeleteIconFromTray()
    
    Call Shell_NotifyIcon(DeleteIcon, data)

End Sub

Public Sub AddIconToTray()
    
    data.Size = Len(data)
    data.handle = hWnd
    data.id = vbNull
    data.Flags = IconFlag Or TipFlag Or MessageFlag
    data.CallBackMessage = WM_MOUSEMOVE
    data.Icon = Icon
    data.Tip = "Tajpi" & vbNullChar
    Call Shell_NotifyIcon(AddIcon, data)

End Sub

Private Sub Form_Resize()
    
    If Me.WindowState = vbMinimized Or Me.WindowState = vbMaximized Then
        Me.Visible = False
    End If
    
End Sub

Private Sub MenuAbout_Click()
    
    FrmAbout.Show
    
End Sub

Private Sub MenuActive_Click()

    Call SetTajpiEnabled(Not MenuActive.Checked, True)

End Sub

Private Sub MenuConfig_Click()
    
    FrmConfig.Display

End Sub

Private Sub MenuExit_Click()

    Unload Me
    End

End Sub

Private Sub MenuHelp_Click()

    ShowHelp

End Sub

Private Sub Timer1_Timer()
    
    Timer1.enabled = False
    Call SetClipboardUnicode(hWnd, clipsave)
    clipsave = ""

End Sub

