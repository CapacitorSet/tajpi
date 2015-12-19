VERSION 5.00
Begin VB.Form FrmClipboardMethod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alglua Metodo"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3435
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
   Icon            =   "FrmClipboardMethod.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   173
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   229
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TxtDelay 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   300
      Left            =   870
      TabIndex        =   3
      Top             =   1575
      Width           =   555
   End
   Begin VB.CheckBox ChkRestore 
      Caption         =   "&Remeti enhavon de tondujo post algluo"
      Height          =   300
      Left            =   165
      TabIndex        =   2
      Top             =   1230
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin VB.OptionButton RBtnCtrlV 
      Caption         =   "&Ctrl + V"
      Height          =   285
      Left            =   150
      TabIndex        =   1
      Top             =   885
      Width           =   1170
   End
   Begin VB.OptionButton RBtnShiftInsert 
      Caption         =   "&Shift + Insert (rekomendinda)"
      Height          =   300
      Left            =   150
      TabIndex        =   0
      Top             =   615
      Value           =   -1  'True
      Width           =   2685
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "&Bone"
      Default         =   -1  'True
      Height          =   345
      Left            =   2400
      TabIndex        =   4
      Top             =   2115
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "milisekundoj"
      Height          =   195
      Left            =   1500
      TabIndex        =   7
      Top             =   1620
      Width           =   1200
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Post"
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   1620
      Width           =   690
   End
   Begin VB.Label Label1 
      Caption         =   "Elektu la metodon por alglui la supersignan literon el la tondujo."
      Height          =   480
      Left            =   180
      TabIndex        =   5
      Top             =   135
      Width           =   3120
   End
End
Attribute VB_Name = "FrmClipboardMethod"
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

    If ChkRestore.Value Then
        If TxtDelay.Text = "" Then
            MsgBox PASTE_MESSAGE_1, vbOKOnly & vbExclamation, "Tajpi"
            TxtDelay.SetFocus
        ElseIf TxtDelay.Text = "0" Or Val(TxtDelay.Text) > 10000 Then
            MsgBox PASTE_MESSAGE_2, vbOKOnly & vbExclamation, "Tajpi"
            TxtDelay.SetFocus
        Else
            FrmMain.Timer1.Interval = Val(TxtDelay.Text)
            Hide
        End If
    Else
        Hide
    End If
    
End Sub

Private Sub ChkRestore_Click()

    SetEnabled ChkRestore.Value, TxtDelay

End Sub

Private Sub TxtDelay_KeyPress(KeyAscii As Integer)
        
    If (KeyAscii < 48 Or KeyAscii > 57) And _
       KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub
