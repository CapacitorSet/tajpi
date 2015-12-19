VERSION 5.00
Begin VB.Form FrmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tajpi: Agordo"
   ClientHeight    =   5730
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   5040
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmConfig.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "FrmConfig"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   382
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   336
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "Lingvo"
      Height          =   630
      Left            =   3840
      TabIndex        =   43
      Top             =   4335
      Width           =   1110
      Begin VB.PictureBox PicEnglish 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   165
         Left            =   660
         Picture         =   "FrmConfig.frx":0E42
         ScaleHeight     =   165
         ScaleWidth      =   240
         TabIndex        =   45
         TabStop         =   0   'False
         ToolTipText     =   "English"
         Top             =   285
         Width           =   240
      End
      Begin VB.PictureBox PicEsperanto 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   165
         Left            =   195
         Picture         =   "FrmConfig.frx":1218
         ScaleHeight     =   165
         ScaleWidth      =   240
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "Esperanto"
         Top             =   285
         Width           =   240
      End
   End
   Begin VB.CommandButton BtnHelp 
      Caption         =   "Helpo [F1]"
      Height          =   360
      Left            =   3840
      TabIndex        =   32
      Top             =   5280
      Width           =   1110
   End
   Begin VB.CommandButton BtnCancel 
      Caption         =   "&Nuligi"
      Height          =   360
      Left            =   3840
      TabIndex        =   31
      Top             =   615
      Width           =   1110
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "&Bone"
      Default         =   -1  'True
      Height          =   360
      Left            =   3840
      TabIndex        =   30
      Top             =   135
      Width           =   1110
   End
   Begin VB.Frame Frame4 
      Caption         =   "Klavkomando"
      Height          =   630
      Left            =   90
      TabIndex        =   42
      Top             =   5010
      Width           =   3630
      Begin VB.CheckBox ChkCtrl 
         Caption         =   "&Ctrl"
         Height          =   270
         Left            =   165
         TabIndex        =   26
         Top             =   255
         Width           =   585
      End
      Begin VB.CheckBox ChkAlt 
         Caption         =   "A&lt"
         Height          =   270
         Left            =   840
         TabIndex        =   27
         Top             =   255
         Width           =   585
      End
      Begin VB.CheckBox ChkShift 
         Caption         =   "Shi&ft"
         Height          =   270
         Left            =   1455
         TabIndex        =   28
         Top             =   255
         Width           =   645
      End
      Begin VB.ComboBox CmbKeys 
         Height          =   315
         ItemData        =   "FrmConfig.frx":1591
         Left            =   2220
         List            =   "FrmConfig.frx":1593
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   210
         Width           =   1260
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Starto"
      Height          =   630
      Left            =   90
      TabIndex        =   41
      Top             =   4335
      Width           =   3630
      Begin VB.CheckBox ChkStartActive 
         Caption         =   "S&tarti aktiva"
         Height          =   345
         Left            =   165
         TabIndex        =   23
         Top             =   210
         Width           =   1515
      End
      Begin Tajpi.UniLabel LblAutomaticStart 
         Height          =   195
         Left            =   1965
         TabIndex        =   25
         Top             =   270
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   344
         Alignment       =   0
         AutoSize        =   -1  'True
         BackColor       =   -2147483633
         BackStyle       =   1
         Caption         =   "Starti auto&mate"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         MouseIcon       =   "FrmConfig.frx":1595
         MousePointer    =   0
         PaddingBottom   =   0
         PaddingLeft     =   0
         PaddingRight    =   0
         PaddingTop      =   0
         RightToLeft     =   0   'False
         UseEvents       =   -1  'True
         UseMnemonic     =   -1  'True
         WordWrap        =   0   'False
      End
      Begin VB.CheckBox ChkAutomaticStart 
         Caption         =   "Starti automate"
         Height          =   285
         Left            =   1695
         TabIndex        =   24
         Top             =   240
         Width           =   1860
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Enigo"
      Height          =   1275
      Left            =   90
      TabIndex        =   40
      Top             =   3015
      Width           =   3630
      Begin VB.CommandButton BtnMethod 
         Caption         =   "..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2250
         TabIndex        =   21
         Top             =   600
         Width           =   270
      End
      Begin VB.CheckBox ChkEntityCodes 
         Caption         =   "&HTML-aj surogatoj"
         Height          =   390
         Left            =   165
         TabIndex        =   22
         Top             =   840
         Width           =   3045
      End
      Begin VB.OptionButton RBtnUnicode 
         Caption         =   "&Unikodo"
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   285
         Value           =   -1  'True
         Width           =   1410
      End
      Begin VB.OptionButton RBtnLatin3 
         Caption         =   "Latina-&3"
         Height          =   195
         Left            =   1665
         TabIndex        =   19
         Top             =   285
         Width           =   1530
      End
      Begin VB.CheckBox ChkPaste 
         Caption         =   "Alglui liter&on el tondujo"
         Height          =   375
         Left            =   165
         TabIndex        =   20
         Top             =   525
         Width           =   2160
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Klavaro"
      Height          =   2925
      Left            =   90
      TabIndex        =   33
      Top             =   45
      Width           =   3630
      Begin Tajpi.UniLabel LblW 
         Height          =   195
         Left            =   2435
         TabIndex        =   17
         Top             =   2565
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   344
         Alignment       =   0
         AutoSize        =   -1  'True
         BackColor       =   -2147483633
         BackStyle       =   1
         Caption         =   "&w por u"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         MouseIcon       =   "FrmConfig.frx":15B1
         MousePointer    =   0
         PaddingBottom   =   0
         PaddingLeft     =   0
         PaddingRight    =   0
         PaddingTop      =   0
         RightToLeft     =   0   'False
         UseEvents       =   -1  'True
         UseMnemonic     =   -1  'True
         WordWrap        =   0   'False
      End
      Begin VB.CheckBox ChkW 
         Caption         =   "&w por u"
         Height          =   345
         Left            =   2160
         TabIndex        =   16
         Top             =   2505
         Width           =   945
      End
      Begin VB.CheckBox ChkAltGr 
         Caption         =   "Alt &Gr + supersignenda litero"
         Height          =   330
         Left            =   150
         TabIndex        =   13
         Top             =   2175
         Width           =   3180
      End
      Begin VB.TextBox TxtSuffixes 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   11
         Top             =   1485
         Width           =   1425
      End
      Begin VB.CheckBox ChkSuffixesRepeat 
         Caption         =   "Rip&eto de sufikso forigas supersignon"
         Enabled         =   0   'False
         Height          =   300
         Left            =   375
         TabIndex        =   12
         Top             =   1845
         Width           =   3225
      End
      Begin VB.TextBox TxtPrefixes 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   8
         Top             =   810
         Width           =   1425
      End
      Begin VB.CheckBox ChkInvisibleSuffix 
         Caption         =   "Igi mal&videbla la prefiksan literon"
         Enabled         =   0   'False
         Height          =   315
         Left            =   375
         TabIndex        =   9
         Top             =   1155
         Width           =   3090
      End
      Begin VB.CheckBox ChkPrefixes 
         Caption         =   "&Prefiksoj"
         Height          =   375
         Left            =   150
         TabIndex        =   7
         Top             =   780
         Width           =   1575
      End
      Begin VB.CheckBox ChkDirectKeys 
         Caption         =   "&Rektaj klavoj"
         Height          =   375
         Left            =   150
         TabIndex        =   0
         Top             =   405
         Width           =   1275
      End
      Begin VB.CheckBox ChkSuffixes 
         Caption         =   "&Sufiksoj"
         Height          =   375
         Left            =   150
         TabIndex        =   10
         Top             =   1455
         Width           =   1590
      End
      Begin VB.TextBox TxtDirectKey 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1560
         MaxLength       =   1
         TabIndex        =   1
         Tag             =   "c"
         Top             =   450
         Width           =   220
      End
      Begin VB.TextBox TxtDirectKey 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   2
         Tag             =   "g"
         Top             =   450
         Width           =   220
      End
      Begin VB.TextBox TxtDirectKey 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   2040
         MaxLength       =   1
         TabIndex        =   3
         Tag             =   "h"
         Top             =   450
         Width           =   220
      End
      Begin VB.TextBox TxtDirectKey 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   2280
         MaxLength       =   1
         TabIndex        =   4
         Tag             =   "j"
         Top             =   450
         Width           =   220
      End
      Begin VB.TextBox TxtDirectKey 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   2520
         MaxLength       =   1
         TabIndex        =   5
         Tag             =   "s"
         Top             =   450
         Width           =   220
      End
      Begin VB.TextBox TxtDirectKey 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   6
         Tag             =   "u"
         Top             =   450
         Width           =   220
      End
      Begin Tajpi.UniLabel LblAutomaticAuEu 
         Height          =   195
         Left            =   420
         TabIndex        =   15
         Top             =   2565
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   344
         Alignment       =   0
         AutoSize        =   -1  'True
         BackColor       =   -2147483633
         BackStyle       =   1
         Caption         =   "&Automata au/eu"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         MouseIcon       =   "FrmConfig.frx":15CD
         MousePointer    =   0
         PaddingBottom   =   0
         PaddingLeft     =   0
         PaddingRight    =   0
         PaddingTop      =   0
         RightToLeft     =   0   'False
         UseEvents       =   -1  'True
         UseMnemonic     =   -1  'True
         WordWrap        =   0   'False
      End
      Begin Tajpi.UniLabel LblCx 
         Height          =   195
         Left            =   1620
         TabIndex        =   34
         Top             =   195
         Width           =   90
         _ExtentX        =   159
         _ExtentY        =   344
         Alignment       =   2
         AutoSize        =   -1  'True
         BackColor       =   -2147483633
         BackStyle       =   1
         Caption         =   "c"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         MouseIcon       =   "FrmConfig.frx":15E9
         MousePointer    =   0
         PaddingBottom   =   0
         PaddingLeft     =   0
         PaddingRight    =   0
         PaddingTop      =   0
         RightToLeft     =   0   'False
         UseEvents       =   -1  'True
         UseMnemonic     =   -1  'True
         WordWrap        =   0   'False
      End
      Begin Tajpi.UniLabel LblGx 
         Height          =   195
         Left            =   1860
         TabIndex        =   35
         Top             =   195
         Width           =   90
         _ExtentX        =   159
         _ExtentY        =   344
         Alignment       =   2
         AutoSize        =   -1  'True
         BackColor       =   -2147483633
         BackStyle       =   1
         Caption         =   "g"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         MouseIcon       =   "FrmConfig.frx":1605
         MousePointer    =   0
         PaddingBottom   =   0
         PaddingLeft     =   0
         PaddingRight    =   0
         PaddingTop      =   0
         RightToLeft     =   0   'False
         UseEvents       =   -1  'True
         UseMnemonic     =   -1  'True
         WordWrap        =   0   'False
      End
      Begin Tajpi.UniLabel LblHx 
         Height          =   195
         Left            =   2100
         TabIndex        =   36
         Top             =   195
         Width           =   90
         _ExtentX        =   159
         _ExtentY        =   344
         Alignment       =   2
         AutoSize        =   -1  'True
         BackColor       =   -2147483633
         BackStyle       =   1
         Caption         =   "h"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         MouseIcon       =   "FrmConfig.frx":1621
         MousePointer    =   0
         PaddingBottom   =   0
         PaddingLeft     =   0
         PaddingRight    =   0
         PaddingTop      =   0
         RightToLeft     =   0   'False
         UseEvents       =   -1  'True
         UseMnemonic     =   -1  'True
         WordWrap        =   0   'False
      End
      Begin Tajpi.UniLabel LblJx 
         Height          =   195
         Left            =   2355
         TabIndex        =   37
         Top             =   195
         Width           =   90
         _ExtentX        =   159
         _ExtentY        =   344
         Alignment       =   2
         AutoSize        =   0   'False
         BackColor       =   -2147483633
         BackStyle       =   1
         Caption         =   "j"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         MouseIcon       =   "FrmConfig.frx":163D
         MousePointer    =   0
         PaddingBottom   =   0
         PaddingLeft     =   0
         PaddingRight    =   0
         PaddingTop      =   0
         RightToLeft     =   0   'False
         UseEvents       =   -1  'True
         UseMnemonic     =   -1  'True
         WordWrap        =   0   'False
      End
      Begin Tajpi.UniLabel LblSx 
         Height          =   195
         Left            =   2595
         TabIndex        =   38
         Top             =   195
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   344
         Alignment       =   2
         AutoSize        =   -1  'True
         BackColor       =   -2147483633
         BackStyle       =   1
         Caption         =   "s"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         MouseIcon       =   "FrmConfig.frx":1659
         MousePointer    =   0
         PaddingBottom   =   0
         PaddingLeft     =   0
         PaddingRight    =   0
         PaddingTop      =   0
         RightToLeft     =   0   'False
         UseEvents       =   -1  'True
         UseMnemonic     =   -1  'True
         WordWrap        =   0   'False
      End
      Begin Tajpi.UniLabel LblUx 
         Height          =   195
         Left            =   2820
         TabIndex        =   39
         Top             =   195
         Width           =   90
         _ExtentX        =   159
         _ExtentY        =   344
         Alignment       =   2
         AutoSize        =   -1  'True
         BackColor       =   -2147483633
         BackStyle       =   1
         Caption         =   "u"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         MouseIcon       =   "FrmConfig.frx":1675
         MousePointer    =   0
         PaddingBottom   =   0
         PaddingLeft     =   0
         PaddingRight    =   0
         PaddingTop      =   0
         RightToLeft     =   0   'False
         UseEvents       =   -1  'True
         UseMnemonic     =   -1  'True
         WordWrap        =   0   'False
      End
      Begin VB.CheckBox ChkAutomaticAuEu 
         Caption         =   "&Automata au/eu"
         Height          =   345
         Left            =   150
         TabIndex        =   14
         Top             =   2505
         Width           =   1665
      End
   End
End
Attribute VB_Name = "FrmConfig"
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

Option Explicit

Private Sub BtnOK_Click()
        
    Hide
    SaveConfig
    ClearBuffer

End Sub

Private Sub BtnCancel_Click()
    
    Hide
    LoadConfig
    
End Sub

Private Sub BtnHelp_Click()
    
    ShowHelp
    
End Sub

Private Sub BtnMethod_Click()
    
    FrmClipboardMethod.Show vbModal, Me
    ChkPaste.SetFocus

End Sub

Private Sub ChkDirectKeys_Click()
    
    Dim i As Integer
    For i = TxtDirectKey.LBound To TxtDirectKey.UBound
        SetEnabled ChkDirectKeys.Value, TxtDirectKey(i)
    Next

End Sub

Private Sub ChkPrefixes_Click()

    SetEnabled ChkPrefixes.Value, TxtPrefixes
    SetEnabled ChkPrefixes.Value, ChkInvisibleSuffix

End Sub

Private Sub ChkSuffixes_Click()

    SetEnabled ChkSuffixes.Value, TxtSuffixes
    SetEnabled ChkSuffixes.Value, ChkSuffixesRepeat

End Sub

Private Sub ChkPaste_Click()
    
    BtnMethod.enabled = ChkPaste.Value

End Sub

Private Sub LblAutomaticAuEu_Click(Button As MouseButtonConstants)

    ChkAutomaticAuEu.Value = IIf(ChkAutomaticAuEu.Value = 0, 1, 0)
    ChkAutomaticAuEu.SetFocus

End Sub

Private Sub LblAutomaticStart_Click(Button As MouseButtonConstants)
    
    ChkAutomaticStart.Value = IIf(ChkAutomaticStart.Value = 0, 1, 0)
    ChkAutomaticStart.SetFocus

End Sub

Private Sub LblW_Click(Button As MouseButtonConstants)
    
    ChkW.Value = IIf(ChkW.Value = 0, 1, 0)
    ChkW.SetFocus

End Sub

Private Sub PicEsperanto_Click()
    
    Call SetLanguage("Esperanto", True)
    SaveConfig
    
End Sub

Private Sub PicEsperanto_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetCursor LoadCursor(0, IDC_HAND)
    
End Sub

Private Sub PicEnglish_Click()

    Call SetLanguage("English", True)
    SaveConfig
    
End Sub

Private Sub PicEnglish_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetCursor LoadCursor(0, IDC_HAND)
    
End Sub

Private Sub TxtDirectKey_KeyPress(Index As Integer, KeyAscii As Integer)
        
    Dim upper As String
    upper = UCase$(Chr$(KeyAscii))
    If upper <> "" Then
        KeyAscii = Asc(upper)
    End If

End Sub

Private Sub TxtPrefixes_KeyPress(KeyAscii As Integer)

    Dim upper As String
    upper = UCase$(Chr$(KeyAscii))
    If upper <> "" Then
        KeyAscii = Asc(upper)
    End If
    
End Sub

Private Sub TxtSuffixes_KeyPress(KeyAscii As Integer)

    Dim upper As String
    upper = UCase$(Chr$(KeyAscii))
    If upper <> "" Then
        KeyAscii = Asc(upper)
    End If
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then
        ShowHelp
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   
    If UnloadMode = vbFormControlMenu Then
        Me.Hide
        LoadConfig
        Cancel = 1
    End If

End Sub

Public Function DirectKeyBoxFocused() As Boolean

    If Me.Visible And (Not Me.ActiveControl Is Nothing) Then
        Dim i As Integer
        For i = Me.TxtDirectKey.LBound To Me.TxtDirectKey.UBound
            If Me.ActiveControl = Me.TxtDirectKey(i) Then
                DirectKeyBoxFocused = True
            End If
        Next
    End If
    
    DirectKeyBoxFocused = False

End Function

Public Function TextBoxFocused() As Boolean

    If Me.Visible And (Not Me.ActiveControl Is Nothing) Then
        If DirectKeyBoxFocused Then
            TextBoxFocused = True
        End If
        If Me.ActiveControl = Me.TxtPrefixes Or Me.ActiveControl = Me.TxtSuffixes Then
            TextBoxFocused = True
        End If
        If FrmClipboardMethod.Visible And Not FrmClipboardMethod.ActiveControl Is Nothing Then
            If FrmClipboardMethod.ActiveControl = FrmClipboardMethod.TxtDelay Then
                TextBoxFocused = True
            End If
        End If
    End If
  
End Function

Public Sub Display()

    LoadConfig
    Visible = True
   
End Sub

Public Sub Prepare()
    
    LblCx.Caption = EOKey("c")
    LblGx.Caption = EOKey("g")
    LblHx.Caption = EOKey("h")
    LblJx.Caption = EOKey("j")
    LblSx.Caption = EOKey("s")
    LblUx.Caption = EOKey("u")
        
    CmbKeys.Clear
    CmbKeys.AddItem ""
    Dim i As Integer
    Dim name As String
    Dim vk As Integer
    For i = 601 To 709
        name = LoadRes(i)
        vk = VKFromKeyName(name)
        CmbKeys.AddItem (name)
        If vk <> 0 Then
            CmbKeys.ItemData(CmbKeys.NewIndex) = vk
        End If
    Next
        
End Sub
