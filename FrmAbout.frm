VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pri Tajpi"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3165
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAbout.frx":0000
   LinkTopic       =   "FrmAbout"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   3165
   StartUpPosition =   2  'CenterScreen
   Begin Tajpi.UniLabel LblEmail 
      Height          =   240
      Left            =   1065
      TabIndex        =   3
      Top             =   1350
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   423
      Alignment       =   0
      AutoSize        =   0   'False
      BackColor       =   -2147483633
      BackStyle       =   1
      Caption         =   "tmj2005@gmail.com"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12582912
      MouseIcon       =   "FrmAbout.frx":0E42
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
   Begin Tajpi.UniLabel Lbl2 
      Height          =   210
      Left            =   180
      TabIndex        =   9
      Top             =   1350
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   370
      Alignment       =   0
      AutoSize        =   0   'False
      BackColor       =   -2147483633
      BackStyle       =   1
      Caption         =   "Retposto:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      MouseIcon       =   "FrmAbout.frx":0E5E
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
   Begin Tajpi.UniLabel LblGPL 
      Height          =   225
      Left            =   2295
      TabIndex        =   2
      Top             =   840
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   397
      Alignment       =   0
      AutoSize        =   0   'False
      BackColor       =   -2147483633
      BackStyle       =   1
      Caption         =   "GPL"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12582912
      MouseIcon       =   "FrmAbout.frx":0E7A
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
   Begin Tajpi.UniLabel Lbl1 
      Height          =   225
      Left            =   180
      TabIndex        =   8
      Top             =   840
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   397
      Alignment       =   0
      AutoSize        =   0   'False
      BackColor       =   -2147483633
      BackStyle       =   1
      Caption         =   "Eldonata lau la kondicoj de la"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      MouseIcon       =   "FrmAbout.frx":0FDC
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
   Begin Tajpi.UniLabel LblWeb 
      Height          =   225
      Left            =   1080
      TabIndex        =   4
      Top             =   1665
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   397
      Alignment       =   0
      AutoSize        =   0   'False
      BackColor       =   -2147483633
      BackStyle       =   1
      Caption         =   "http://tajpi.webhop.net"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12582912
      MouseIcon       =   "FrmAbout.frx":0FF8
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
   Begin VB.CommandButton BtnOK 
      Caption         =   "&Bone"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2070
      TabIndex        =   1
      Top             =   2205
      Width           =   945
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2535
      Picture         =   "FrmAbout.frx":1014
      ScaleHeight     =   405
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   165
      Width           =   480
   End
   Begin Tajpi.UniLabel Lbl3 
      Height          =   210
      Left            =   180
      TabIndex        =   10
      Top             =   1665
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   370
      Alignment       =   0
      AutoSize        =   0   'False
      BackColor       =   -2147483633
      BackStyle       =   1
      Caption         =   "Retpagaro:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      MouseIcon       =   "FrmAbout.frx":17B6
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
   Begin VB.Label Label3 
      Caption         =   "© 2008-2012 Thomas James"
      Height          =   225
      Left            =   165
      TabIndex        =   7
      Top             =   615
      Width           =   2310
   End
   Begin VB.Label Label2 
      Caption         =   "Klavarilo por esperantistoj"
      Height          =   225
      Left            =   180
      TabIndex        =   6
      Top             =   390
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Tajpi v2.97"
      Height          =   225
      Left            =   180
      TabIndex        =   5
      Top             =   165
      Width           =   1095
   End
End
Attribute VB_Name = "FrmAbout"
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

Private Sub BtnOK_Click()
    
    Hide

End Sub

Private Sub LblEmail_Click(Button As MouseButtonConstants)

    Call ShellExecute(Me.hWnd, "open", "mailto:tmj2005@gmail.com", 0&, 0&, SW_SHOWNORMAL)

End Sub

Private Sub LblEmail_MouseMove(Button As ulbMouseButtonConstants, Shift As ulbShiftConstants, X As Single, Y As Single)

    SetCursor LoadCursor(0, IDC_HAND)

End Sub


Private Sub LblGPL_Click(Button As MouseButtonConstants)
    
    Call ShellExecute(Me.hWnd, "open", "http://www.gnu.org/copyleft/gpl.html", 0&, 0&, SW_SHOWNORMAL)

End Sub

Private Sub LblGPL_MouseMove(Button As ulbMouseButtonConstants, Shift As ulbShiftConstants, X As Single, Y As Single)

    SetCursor LoadCursor(0, IDC_HAND)

End Sub

Private Sub LblWeb_Click(Button As MouseButtonConstants)

    Call ShellExecute(Me.hWnd, "open", "http://tajpi.webhop.net", 0&, 0&, SW_SHOWNORMAL)

End Sub

Private Sub LblWeb_MouseMove(Button As ulbMouseButtonConstants, Shift As ulbShiftConstants, X As Single, Y As Single)

    SetCursor LoadCursor(0, IDC_HAND)
    
End Sub
