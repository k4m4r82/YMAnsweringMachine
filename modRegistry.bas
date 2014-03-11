Attribute VB_Name = "modRegistry"
Option Explicit

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002

Private Const REG_SZ = 1 ' Unicode nul terminated string

Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Private Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    
    On Error Resume Next
    
    'retrieve nformation about the key
    lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = REG_SZ Then
            'Create a buffer
            strBuf = String(lDataBufSize, Chr$(0))
            'retrieve the key's content
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then
                'Remove the unnecessary chr$(0)'s
                RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If
        End If
    End If
End Function

Public Function getFromWindowsRegistry(hKey As Long, strPath As String, strValue As String, Optional defValue As String = "") As String
    Dim ret As Long
    Dim tmp As String
    
    On Error Resume Next
    
    'Open the key
    RegOpenKey hKey, strPath, ret
    'Get the key's content
    tmp = RegQueryStringValue(ret, strValue)
    'Close the key
    RegCloseKey ret
    
    If Not (Len(tmp) > 0) Then tmp = defValue
    
    getFromWindowsRegistry = tmp
End Function

Public Sub saveToWindowsRegistry(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim ret As Long
    
    On Error Resume Next
    
    'Create a new key
    RegCreateKey hKey, strPath, ret
    'Save a string to the key
    RegSetValueEx ret, strValue, 0, REG_SZ, ByVal strData, Len(strData)
    'close the key
    RegCloseKey ret
End Sub


