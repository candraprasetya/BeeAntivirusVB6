Attribute VB_Name = "basReg"
Option Explicit

Private lReg As Long
Private KeyHandle As Long
Private lResult As Long
Private lValueType As Long
Private lDataBufSize As Long

Private Const ERROR_SUCCESS = 0&
Private Const REG_SZ = 1
Private Const REG_DWORD = 4
Const KEY_READ = ((&H20000 Or &H1 Or &H8 Or &H10) And (Not &H100000))
Private Const KEY_QUERY_VALUE = &H1

' API yang berhubungan dengan Registry
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

' Enumerisasi Main Key
Public Enum MainKey
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum

Public Function CreateKeyReg(hKey As MainKey, spath As String) As Long
    
    lReg = RegCreateKey(hKey, spath, KeyHandle)
    lReg = RegCloseKey(KeyHandle)
    
End Function

Public Function GetStringValue(hKey As MainKey, spath As String, sValue As String) As String
    Dim sBuff As String
    Dim intZeroPos As Integer
    
    lReg = RegOpenKey(hKey, spath, KeyHandle)
    lResult = RegQueryValueEx(KeyHandle, sValue, 0&, lValueType, ByVal 0&, lDataBufSize)

    If lValueType = REG_SZ Then
        sBuff = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(KeyHandle, sValue, 0&, 0&, ByVal sBuff, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(sBuff, Chr$(0))
            If intZeroPos > 0 Then
                GetStringValue = Left$(sBuff, intZeroPos - 1)
            Else
                GetStringValue = sBuff
            End If
        End If
    End If
    
End Function

Public Function SetStringValue(hKey As MainKey, spath As String, sValue As String, sData As String) As Long
    
    lReg = RegCreateKey(hKey, spath, KeyHandle)
    lReg = RegSetValueEx(KeyHandle, sValue, 0, REG_SZ, ByVal sData, Len(sData))
    lReg = RegCloseKey(KeyHandle)
    
End Function

Function GetDwordValue(ByVal hKey As MainKey, ByVal spath As String, ByVal sValueName As String) As Long
    
    Dim lBuff As Long
    
    lReg = RegOpenKey(hKey, spath, KeyHandle)
    lDataBufSize = 4
    lResult = RegQueryValueEx(KeyHandle, sValueName, 0&, lValueType, lBuff, lDataBufSize)

    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            GetDwordValue = lBuff
        End If
    End If
    
    lReg = RegCloseKey(KeyHandle)
    
End Function

Public Function SetDwordValue(ByVal hKey As MainKey, ByVal spath As String, ByVal sValueName As String, ByVal lData As Long) As Long
    
    lReg = RegCreateKey(hKey, spath, KeyHandle)
    lResult = RegSetValueEx(KeyHandle, sValueName, 0&, REG_DWORD, lData, 4)
    lReg = RegCloseKey(KeyHandle)
    
End Function

Public Function DeleteKey(ByVal hKey As MainKey, ByVal sKey As String) As Long
    lReg = RegDeleteKey(hKey, sKey)
End Function

Public Function DeleteValue(ByVal hKey As MainKey, ByVal spath As String, ByVal sValue As String) As Long
    lReg = RegOpenKey(hKey, spath, KeyHandle)
    lReg = RegDeleteValue(KeyHandle, sValue)
    lReg = RegCloseKey(KeyHandle)
End Function

Public Function SingkatanMainKey(sSingkatannya As String) As MainKey
Select Case sSingkatannya
    Case "HKLM"
    SingkatanMainKey = HKEY_LOCAL_MACHINE
    Case "HKCU"
    SingkatanMainKey = HKEY_CURRENT_USER
    Case "HKCR"
    SingkatanMainKey = HKEY_CLASSES_ROOT
    Case "HKU"
    SingkatanMainKey = HKEY_USERS
End Select
End Function

