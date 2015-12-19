Attribute VB_Name = "ModRegistry"
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

Private m_lngRetVal As Long
  
Private Const REG_NONE As Long = 0
Private Const REG_SZ As Long = 1
Private Const REG_EXPAND_SZ As Long = 2
Private Const REG_BINARY As Long = 3
Private Const REG_DWORD As Long = 4
Private Const REG_DWORD_LITTLE_ENDIAN As Long = 4
Private Const REG_DWORD_BIG_ENDIAN As Long = 5
Private Const REG_LINK As Long = 6
Private Const REG_MULTI_SZ As Long = 7
Private Const REG_RESOURCE_LIST As Long = 8
Private Const REG_FULL_RESOURCE_DESCRIPTOR As Long = 9
Private Const REG_RESOURCE_REQUIREMENTS_LIST As Long = 10

Private Const KEY_QUERY_VALUE As Long = &H1
Private Const KEY_SET_VALUE As Long = &H2
Private Const KEY_CREATE_SUB_KEY As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Private Const KEY_NOTIFY As Long = &H10
Private Const KEY_CREATE_LINK As Long = &H20
Private Const KEY_ALL_ACCESS As Long = &H3F

Public Const HKEY_CLASSES_ROOT As Long = &H80000000
Public Const HKEY_CURRENT_USER As Long = &H80000001
Public Const HKEY_LOCAL_MACHINE As Long = &H80000002
Public Const HKEY_USERS As Long = &H80000003
Public Const HKEY_PERFORMANCE_DATA As Long = &H80000004
Public Const HKEY_CURRENT_CONFIG As Long = &H80000005
Public Const HKEY_DYN_DATA As Long = &H80000006

Private Const ERROR_SUCCESS As Long = 0
Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const ERROR_NO_MORE_ITEMS As Long = 259

Private Const REG_OPTION_NON_VOLATILE As Long = 0
Private Const REG_OPTION_VOLATILE As Long = &H1

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal lngRootKey As Long) As Long

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
          (ByVal lngRootKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
          (ByVal lngRootKey As Long, ByVal lpSubKey As String) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
          (ByVal lngRootKey As Long, ByVal lpValueName As String) As Long

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
          (ByVal lngRootKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
          (ByVal lngRootKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
           lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
          (ByVal lngRootKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
           ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Function regDelete_Sub_Key(ByVal lngRootKey As Long, _
                                  ByVal strRegKeyPath As String, _
                                  ByVal strRegSubKey As String)
    
    Dim lngKeyHandle As Long
  
    If regDoes_Key_Exist(lngRootKey, strRegKeyPath) Then
    
        m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
        m_lngRetVal = RegDeleteValue(lngKeyHandle, strRegSubKey)
        m_lngRetVal = RegCloseKey(lngKeyHandle)
    End If
  
End Function

Public Function regDoes_Key_Exist(ByVal lngRootKey As Long, _
                                  ByVal strRegKeyPath As String) As Boolean

    Dim lngKeyHandle As Long
    lngKeyHandle = 0
    m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
    If lngKeyHandle = 0 Then
        regDoes_Key_Exist = False
    Else
        regDoes_Key_Exist = True
    End If
    m_lngRetVal = RegCloseKey(lngKeyHandle)

End Function

Public Function regQuery_A_Key(ByVal lngRootKey As Long, _
                               ByVal strRegKeyPath As String, _
                               ByVal strRegSubKey As String) As Variant

    Dim intPosition As Integer
    Dim lngKeyHandle As Long
    Dim lngDataType As Long
    Dim lngBufferSize As Long
    Dim lngBuffer As Long
    Dim strBuffer As String

    lngKeyHandle = 0
    lngBufferSize = 0

    m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)

    If lngKeyHandle = 0 Then
        regQuery_A_Key = ""
        m_lngRetVal = RegCloseKey(lngKeyHandle)
        Exit Function
    End If

    m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, _
                           lngDataType, ByVal 0&, lngBufferSize)

    If lngKeyHandle = 0 Then
        regQuery_A_Key = ""
        m_lngRetVal = RegCloseKey(lngKeyHandle)
        Exit Function
    End If

    Select Case lngDataType
           Case REG_SZ:

                strBuffer = Space$(lngBufferSize)

                m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, 0&, _
                                       ByVal strBuffer, lngBufferSize)

                If m_lngRetVal <> ERROR_SUCCESS Then
                    regQuery_A_Key = ""
                Else

                    intPosition = InStr(1, strBuffer, Chr(0))
                    If intPosition > 0 Then
                        regQuery_A_Key = Left$(strBuffer, intPosition - 1)
                    Else
                        regQuery_A_Key = strBuffer
                    End If
                End If

           Case REG_DWORD:
                m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, _
                                       lngBuffer, 4&)

                If m_lngRetVal <> ERROR_SUCCESS Then
                    regQuery_A_Key = ""
                Else
                    regQuery_A_Key = lngBuffer
                End If

           Case Else:
                regQuery_A_Key = ""
    End Select

    m_lngRetVal = RegCloseKey(lngKeyHandle)

End Function

Public Sub regCreate_Key_Value(ByVal lngRootKey As Long, ByVal strRegKeyPath As String, _
                               ByVal strRegSubKey As String, varRegData As Variant)
    
    Dim lngKeyHandle As Long
    Dim lngDataType As Long
    Dim lngKeyValue As Long
    Dim strKeyValue As String
    
    If IsNumeric(varRegData) Then
        lngDataType = REG_DWORD
    Else
        lngDataType = REG_SZ
    End If
    
    m_lngRetVal = RegCreateKey(lngRootKey, strRegKeyPath, lngKeyHandle)
      
    Select Case lngDataType
           Case REG_SZ:
                strKeyValue = Trim(varRegData) & Chr(0)
                m_lngRetVal = RegSetValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, _
                                            ByVal strKeyValue, Len(strKeyValue))
                                     
           Case REG_DWORD:
                lngKeyValue = CLng(varRegData)
                m_lngRetVal = RegSetValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, _
                                            lngKeyValue, 4&)
                                     
    End Select
    
    m_lngRetVal = RegCloseKey(lngKeyHandle)
  
End Sub

Public Function regCreate_A_Key(ByVal lngRootKey As Long, ByVal strRegKeyPath As String)

    Dim lngKeyHandle As Long

    m_lngRetVal = RegCreateKey(lngRootKey, strRegKeyPath, lngKeyHandle)

    m_lngRetVal = RegCloseKey(lngKeyHandle)

End Function

Public Function regDelete_A_Key(ByVal lngRootKey As Long, _
                                ByVal strRegKeyPath As String, _
                                ByVal strRegKeyName As String) As Boolean

    Dim lngKeyHandle As Long

    regDelete_A_Key = False

    If regDoes_Key_Exist(lngRootKey, strRegKeyPath) Then

        m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)

        m_lngRetVal = RegDeleteKey(lngKeyHandle, strRegKeyName)

        If m_lngRetVal = 0 Then regDelete_A_Key = True

        m_lngRetVal = RegCloseKey(lngKeyHandle)
    End If

End Function

