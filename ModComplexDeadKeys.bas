Attribute VB_Name = "ModComplexDeadKeys"
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

'This module deals with complex dead keys (dead keys which require use of a modifier key,
'such as Alt-Gr, Shift, or others). It is necessary because MapVirtualKeyEx() does not report dead
'keys which require a modifier. We have to handle these separately :(
'One way to do this would be to load the appropriate keyboard .dll file at runtime and examine it
'to determine what dead keys are present and what modifiers they use, if any. This however would require
'loading a 10k .dll file on every key press, which is a bit heavy. So instead we will refer to a hard
'coded table containing the complex dead key profile of the most common keymaps. Others may need to be
'added over time. Not an elegant solution but it does the job well enough.
'Of course all of this can be avoided by hooking WM_CHAR. But this requires .dll injection, which Tajpi
'aims to avoid by having a single .exe file.

Option Explicit

Public Const DANISH_DENMARK = 67503110
Public Const DUTCH_NETHERLANDS = 68355091
Public Const ESTONIAN_ESTONIA = 69534757
Public Const FAROESE = 70779960
Public Const FINNISH = 67830795
Public Const FINNISH_WITH_SAMI = -265485301
Public Const FRENCH_BELGIAN = 135006220
Public Const FRENCH_BELGIAN_COMMA = -266467316
Public Const FRENCH_CANADIAN = 269028364
Public Const FRENCH_CANADIAN_MULTILINGUAL_STANDARD = -266335220
Public Const FRENCH_FRANCE = 67896332
Public Const GERMAN_GERMANY = 67568647
Public Const GERMAN_AUSTRIA = 67570695
Public Const GREEK_GREECE = 67634184
Public Const ICELANDIC_ICELAND = 68092943
Public Const IRISH_IRELAND = 403245116
Public Const PORTUGUESE_BRAZILIAN_ABNT = 68551702
Public Const PORTUGUESE_BRAZILIAN_ABNT2 = -267385834
Public Const PORTUGUESE_PORTUGAL = 135661590
Public Const ROMANIAN_STANDARD = -257620968
Public Const SPANISH_SPAIN_INTERNATIONAL_SORT = 67767306
Public Const SPANISH_SPAIN_TRADITIONAL_SORT = 67765258
Public Const SWEDISH_SWEDEN = 69010461
Public Const TURKISH_Q = 69141535
Public Const UNITED_KINGDOM_EXTENDED = 72484873
Public Const UNITED_STATES_INTERNATIONAL = -268368887

Public Function ComplexDeadKey(ByVal key As String, ByVal layout As Long) As String

    ComplexDeadKey = ""
    
    Select Case layout
        
        Case DANISH_DENMARK
            If shiftDown Then
                Select Case key
                    Case "´"
                        ComplexDeadKey = "`"
                    Case "¨"
                        ComplexDeadKey = "^"
                End Select
            End If
        
        Case DUTCH_NETHERLANDS
            If shiftDown Then
                Select Case key
                    Case "°"
                        ComplexDeadKey = "~"
                    Case "¨"
                        ComplexDeadKey = "^"
                    Case "´"
                        ComplexDeadKey = "`"
                End Select
            End If
            If key = "°" And (altGrDown Or (ctrlDown And altDown)) Then
                ComplexDeadKey = "¸"
            End If
            
        Case ESTONIAN_ESTONIA
            If shiftDown Then
                Select Case key
                    Case "´"
                        ComplexDeadKey = "`"
                    Case "Ž"
                        ComplexDeadKey = "~"
                End Select
            End If
        
        Case FAROESE
            If altGrDown Or (ctrlDown And altDown) Then
                Select Case key
                    Case "å"
                        ComplexDeadKey = "¨"
                    Case "ð"
                        ComplexDeadKey = "~"
                    Case "ø"
                        ComplexDeadKey = "^"
                End Select
            ElseIf key = "´" And shiftDown Then
                ComplexDeadKey = "`"
            End If
            
        Case FINNISH
        Case FINNISH_WITH_SAMI
            If shiftDown Then
                Select Case key
                    Case "´"
                        ComplexDeadKey = "`"
                    Case "¨"
                        ComplexDeadKey = "^"
                End Select
            ElseIf key = "¨" And (altGrDown Or (ctrlDown And altDown)) Then
                ComplexDeadKey = "~"
            End If
            
        Case FRENCH_CANADIAN
            If key = "é" And (altGrDown Or (ctrlDown And altDown)) Then
                ComplexDeadKey = "´"
            ElseIf key = "¸" And shiftDown Then
                ComplexDeadKey = "¨"
            End If
            
        Case FRENCH_BELGIAN
        Case FRENCH_BELGIAN_COMMA
            If altGrDown Or (ctrlDown And altDown) Then
                Select Case key
                    Case "ù"
                        ComplexDeadKey = "´"
                    Case "µ"
                        ComplexDeadKey = "`"
                    Case "="
                        ComplexDeadKey = "~"
                End Select
            ElseIf key = "^" And shiftDown Then
                ComplexDeadKey = "¨"
            End If
        
        Case FRENCH_CANADIAN_MULTILINGUAL_STANDARD
            Select Case key
                Case "="
                    If OEM8Down Then
                        If shiftDown Then
                            ComplexDeadKey = key
                        Else
                            ComplexDeadKey = "¸"
                        End If
                    End If
                Case "^"
                    If altGrDown Or (ctrlDown And altDown) Then
                        ComplexDeadKey = "`"
                    ElseIf shiftDown Then
                        If OEM8Down Then
                            ComplexDeadKey = "°"
                        Else
                            ComplexDeadKey = "¨"
                        End If
                    Else
                        ComplexDeadKey = key
                    End If
                Case "ç"
                    If altGrDown Or (ctrlDown And altDown) Then
                        ComplexDeadKey = "~"
                    ElseIf OEM8Down And shiftDown Then
                        ComplexDeadKey = "¯"
                    End If
                Case ";"
                    If OEM8Down Then
                        ComplexDeadKey = "´"
                    End If
                Case "è"
                    If OEM8Down And shiftDown Then
                        ComplexDeadKey = key
                    End If
                Case "à"
                    If OEM8Down And shiftDown Then
                        ComplexDeadKey = key
                    End If
                Case "é"
                    If OEM8Down And shiftDown Then
                        ComplexDeadKey = "·"
                    End If
            End Select
            
        Case FRENCH_FRANCE
            If key = "^" And shiftDown Then
                ComplexDeadKey = "¨"
            End If
            
        Case GERMAN_AUSTRIA
        Case GERMAN_GERMANY
            If key = "´" And shiftDown Then
                ComplexDeadKey = "`"
            End If
            
        Case GREEK_GREECE
            Select Case key
                Case "ò"
                    If shiftDown Then
                        ComplexDeadKey = "¡"
                    End If
                Case "´"
                    If shiftDown Then
                        ComplexDeadKey = "¨"
                    ElseIf altGrDown Or (ctrlDown And altDown) Then
                        ComplexDeadKey = "¡"
                    End If
            End Select
                    
        Case ICELANDIC_ICELAND
            If key = "°" And shiftDown Then
                ComplexDeadKey = "¨"
            ElseIf key = "´" And (altGrDown Or (ctrlDown And altDown)) Then
                ComplexDeadKey = "^"
            End If
        
        Case IRISH_IRELAND
            If key = "'" And (altGrDown Or (ctrlDown And altDown)) Then
                ComplexDeadKey = "´"
            End If
                    
        Case PORTUGUESE_BRAZILIAN_ABNT
        Case PORTUGUESE_BRAZILIAN_ABNT2
            If shiftDown Then
                Select Case key
                    Case "6"
                        ComplexDeadKey = "¨"
                    Case "´"
                        ComplexDeadKey = "`"
                    Case "~"
                        ComplexDeadKey = "^"
                End Select
            End If
        
        Case PORTUGUESE_PORTUGAL
            If shiftDown Then
                Select Case key
                    Case "´"
                        ComplexDeadKey = "`"
                    Case "~"
                        ComplexDeadKey = "^"
                End Select
            ElseIf key = "+" And (altGrDown Or (ctrlDown And altDown)) Then
                ComplexDeadKey = "¨"
            End If
        
        Case ROMANIAN_STANDARD
            If altGrDown Or (ctrlDown And altDown) Then
                Select Case key
                    Case "1"
                        ComplexDeadKey = "~"
                    Case "2"
                        ComplexDeadKey = "¡"
                    Case "3"
                        ComplexDeadKey = "^"
                    Case "4"
                        ComplexDeadKey = "¢"
                    Case "5"
                        ComplexDeadKey = "°"
                    Case "6"
                        ComplexDeadKey = "²"
                    Case "7"
                        ComplexDeadKey = "`"
                    Case "8"
                        ComplexDeadKey = "·"
                    Case "9"
                        ComplexDeadKey = "´"
                    Case "0"
                        ComplexDeadKey = "½"
                    Case "-"
                        ComplexDeadKey = "¨"
                    Case "="
                        ComplexDeadKey = "¸"
                End Select
            End If
            
        Case SPANISH_SPAIN_INTERNATIONAL_SORT
        Case SPANISH_SPAIN_TRADITIONAL_SORT
            If shiftDown Then
                Select Case key
                    Case "`"
                        ComplexDeadKey = "^"
                    Case "´"
                        ComplexDeadKey = "¨"
                End Select
            ElseIf key = "4" And (altGrDown Or (ctrlDown And altDown)) Then
                ComplexDeadKey = "~"
            End If
            
        Case SWEDISH_SWEDEN
            If shiftDown Then
                Select Case key
                    Case "´"
                        ComplexDeadKey = "`"
                    Case "¨"
                        ComplexDeadKey = "^"
                End Select
            ElseIf key = "¨" And (altGrDown Or (ctrlDown And altDown)) Then
                ComplexDeadKey = "~"
            End If
        
        Case TURKISH_Q
            If key = "3" And shiftDown Then
                ComplexDeadKey = "^"
            ElseIf altGrDown Or (ctrlDown And altDown) Then
                Select Case key
                    Case "ð"
                        ComplexDeadKey = "¨"
                    Case "ü"
                        ComplexDeadKey = "~"
                    Case "þ"
                        ComplexDeadKey = "´"
                    Case ","
                        ComplexDeadKey = "`"
                End Select
            End If
            
        Case UNITED_KINGDOM_EXTENDED
            If altGrDown Or (ctrlDown And altDown) Then
                Select Case key
                    Case "2"
                        ComplexDeadKey = "¨"
                    Case "6"
                        ComplexDeadKey = "^"
                    Case "#"
                        ComplexDeadKey = "~"
                End Select
                
            End If
        
        Case UNITED_STATES_INTERNATIONAL
            If shiftDown Then
                Select Case key
                    Case "6"
                        ComplexDeadKey = "^"
                    Case "`"
                        ComplexDeadKey = "~"
                    Case "'"
                        ComplexDeadKey = """"
                End Select
            End If
        
    End Select

End Function
