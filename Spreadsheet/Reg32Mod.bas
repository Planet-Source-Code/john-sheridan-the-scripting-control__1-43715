Attribute VB_Name = "Reg32Mod"
              'ADVAPI32 Registry API Bas File.
' This file was not writen by me but I like to thank who did write it

' --------------------------------------------------------------------
' ADVAPI32
' --------------------------------------------------------------------
' function prototypes, constants, and type definitions
' for Windows 32-bit Registry API

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&

' Registry API prototypes
Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_DWORD = 4                      ' 32-bit number
Public sKeys As Collection

Public Sub SaveKey(hKey As Long, strPath As String)
Dim Keyhand&
    r = RegCreateKey(hKey, strPath, Keyhand&)
    r = RegCloseKey(Keyhand&)
End Sub

Public Function GetString(hKey As Long, strPath As String, strValue As String)

Dim Keyhand As Long
Dim datatype As Long
Dim lResult As Long
Dim strBuf As String
Dim lDataBufSize As Long
Dim intZeroPos As Integer
r = RegOpenKey(hKey, strPath, Keyhand)
lResult = RegQueryValueEx(Keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
If lValueType = REG_SZ Then
    strBuf = String(lDataBufSize, " ")
    lResult = RegQueryValueEx(Keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        intZeroPos = InStr(strBuf, Chr$(0))
        If intZeroPos > 0 Then
            GetString = Left$(strBuf, intZeroPos - 1)
        Else
            GetString = strBuf
        End If
    End If
End If
End Function


Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strdata As String)
Dim Keyhand As Long
Dim r As Long
    r = RegCreateKey(hKey, strPath, Keyhand)
    r = RegSetValueEx(Keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(Keyhand)
End Sub


Function GetDWord(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
Dim lResult As Long
Dim lValueType As Long
Dim lBuf As Long
Dim lDataBufSize As Long
Dim r As Long
Dim Keyhand As Long

r = RegOpenKey(hKey, strPath, Keyhand)

 ' Get length/data type
lDataBufSize = 4
    
lResult = RegQueryValueEx(Keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)

If lResult = ERROR_SUCCESS Then
    If lValueType = REG_DWORD Then
        GetDWord = lBuf
    End If
'Else
'    Call errlog("GetDWORD-" & strPath, False)
End If

r = RegCloseKey(Keyhand)
    
End Function

Function SaveDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    Dim lResult As Long
    Dim Keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, Keyhand)
    lResult = RegSetValueEx(Keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    'If lResult <> error_success Then Call errlog("SetDWORD", False)
    r = RegCloseKey(Keyhand)
End Function

Public Function DeleteKey(ByVal hKey As Long, ByVal StrKey As String)
Dim r As Long
    r = RegDeleteKey(hKey, StrKey)
    
End Function

Public Function DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
Dim Keyhand As Long
    r = RegOpenKey(hKey, strPath, Keyhand)
    r = RegDeleteValue(Keyhand, strValue)
    r = RegCloseKey(Keyhand)
End Function

Public Sub GetKeyNames(ByVal hKey As Long, ByVal strPath As String)
Dim Cnt As Long, StrBuff As String, StrKey As String, TKey As Long
    RegOpenKey hKey, strPath, TKey
    Do
        StrBuff = String(255, vbNullChar)
        If RegEnumKeyEx(TKey, Cnt, StrBuff, 255, 0, vbNullString, 0, ByVal 0&) <> 0 Then Exit Do
        Cnt = Cnt + 1
        StrKey = Left(StrBuff, InStr(StrBuff, vbNullChar) - 1)
        sKeys.Add StrKey
    Loop
End Sub



