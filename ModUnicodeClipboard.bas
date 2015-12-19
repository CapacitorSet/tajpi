Attribute VB_Name = "ModUnicodeClipboard"
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

Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Const CF_UNICODETEXT = 13
Private Const GMEM_DDESHARE As Long = &H2000

Public clipsave As String

Public Function GetClipboardUnicode(hWnd As Long) As String
    
    On Error GoTo HandleError:

    ' Returns a byte array containing binary data on the clipboard for
    ' format lFormatID:
    Dim hMem As Long, lSize As Long, lPtr As Long
    Dim sReturn As String

    If OpenClipboard(hWnd) Then

        If IsClipboardFormatAvailable(CF_UNICODETEXT) = 0 Then
            CloseClipboard
            Exit Function
        End If

        hMem = GetClipboardData(CF_UNICODETEXT)
        ' If success:
        If (hMem <> 0) Then
            ' Get the size of this memory block:
            lSize = GlobalSize(hMem)
            ' Get a pointer to the memory:
            lPtr = GlobalLock(hMem)
            If (lSize > 0) Then
                ' Resize the byte array to hold the data:
                sReturn = String$(lSize \ 2 + 1, 0)
                ' Copy from the pointer into the array:
                CopyMemory ByVal StrPtr(sReturn), ByVal lPtr, lSize
            End If
            ' Unlock the memory block:
            GlobalUnlock hMem
            ' Success:
            GetClipboardUnicode = Left$(sReturn, Len(sReturn) - 2)
            ' Don't free the memory - it belongs to the clipboard.
        End If

        CloseClipboard
    End If
    
HandleError:

End Function

Public Function SetClipboardUnicode(hWnd As Long, sUniText As String) As Boolean
    
    On Error GoTo HandleError:
    
    ' Puts the binary data contained in bData() onto the clipboard under
    ' format lFormatID:
    Dim lSize As Long
    Dim lPtr As Long
    Dim hMem As Long
    
    ' Determine the size of the binary data to write:
    If OpenClipboard(hWnd) Then
        lSize = LenB(sUniText) + 2
        ' Generate global memory to hold this:
        hMem = GlobalAlloc(GMEM_DDESHARE, lSize)
        If (hMem <> 0) Then
            ' Get pointer to the memory block:
            lPtr = GlobalLock(hMem)
            ' Copy the data into the memory block:
            CopyMemory ByVal lPtr, ByVal StrPtr(sUniText), lSize - 2
            ' Unlock the memory block.
            GlobalUnlock hMem
            
            ' Now set the clipboard data:
            If (SetClipboardData(CF_UNICODETEXT, hMem) <> 0) Then
                ' Success:
                SetClipboardUnicode = True
            End If
        End If
        ' We don't free the memory because the clipboard takes
        ' care of that now.
        CloseClipboard
    End If

HandleError:
    
End Function
