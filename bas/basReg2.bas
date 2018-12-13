Attribute VB_Name = "basReg"
Option Explicit

Private lReg            As Long
Private KeyHandle       As Long
Private lResult         As Long
Private lValueType      As Long
Private lDataBufSize    As Long

Private Const ERROR_SUCCESS = 0&
Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2

Private Const REG_DWORD = 4
Const KEY_READ = ((&H20000 Or &H1 Or &H8 Or &H10) And (Not &H100000))

' API yang berhubungan dengan Registry
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Private Declare Function RegCreateKeyW Lib "advapi32.dll" (ByVal hkey As Long, ByVal lpSubKey As Long, ByVal pz_phkResult As Long) As Long '---###FIX!IT###---:phkResult--->ByVal pz_phkResult As Long
Private Declare Function RegDeleteKeyW Lib "advapi32.dll" (ByVal hkey As Long, ByVal lpSubKey As Long) As Long
Private Declare Function RegDeleteValueW Lib "advapi32.dll" (ByVal hkey As Long, ByVal pz_lpValueName As Long) As Long
Private Declare Function RegEnumValueW Lib "advapi32.dll" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpValueName As Long, ByVal pz_lpcbValueName As Long, ByVal lpReserved As Long, ByVal pz_lpType As Long, ByVal pz_lpData As Long, ByVal pz_lpcbData As Long) As Long '---###FIX!IT###---:variabel yang ditambah diawali dgn "pz_" diubah menjadi "ByVal * As Long" (sebagai pointer).
Private Declare Function RegOpenKeyW Lib "advapi32.dll" (ByVal hkey As Long, ByVal lpSubKey As Long, ByVal pz_phkResult As Long) As Long '---###FIX!IT###---:variabel yang ditambah diawali dgn "pz_" diubah menjadi "ByVal * As Long" (sebagai pointer).
Private Declare Function RegOpenKeyExW Lib "advapi32.dll" (ByVal hkey As Long, ByVal lpSubKey As Long, ByVal ulOptions As Long, ByVal samDesired As Long, ByVal pz_phkResult As Long) As Long '---###FIX!IT###---:variabel yang ditambah diawali dgn "pz_" diubah menjadi "ByVal * As Long" (sebagai pointer).
Private Declare Function RegQueryValueExW Lib "advapi32.dll" (ByVal hkey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, ByVal pz_lpType As Long, ByVal pz_lpData As Long, ByVal pz_lpcbData As Long) As Long '---###FIX!IT###---:variabel yang ditambah diawali dgn "pz_" diubah menjadi "ByVal * As Long" (sebagai pointer).
Private Declare Function RegSetValueExW Lib "advapi32.dll" (ByVal hkey As Long, ByVal lpValueName As Long, ByVal Reserved As Long, ByVal dwType As Long, ByVal pz_lpData As Long, ByVal cbData As Long) As Long '---###FIX!IT###---:variabel yang ditambah diawali dgn "pz_" diubah menjadi "ByVal * As Long" (sebagai pointer).
Dim cImgList  As gComCtl


Public Function CreateKeyReg(ByVal hkey As Long, ByRef sPath As String) As Long '---?tanya?:nggak kepakai? '---###FIX!IT###---:optimisasi variabel string,ganti ByRef aja,kalau cuman buat dibaca, bukan ditulis / dimodifikasi ulang.
    
    lReg = RegCreateKeyW(hkey, StrPtr(sPath), VarPtr(KeyHandle)) '---###FIX!IT###---:KeyHandle--->VarPtr(KeyHandle)
    lReg = RegCloseKey(KeyHandle)
    
End Function

Public Function GetStringValue(ByVal hkey As Long, ByRef sPath As String, ByRef sValue As String) As String '---###FIX!IT###---:optimisasi variabel string,ganti ByRef aja,kalau cuman buat dibaca, bukan ditulis / dimodifikasi ulang.
    Dim sBuff As String
    Dim intZeroPos As Integer
    
    lReg = RegOpenKeyW(hkey, StrPtr(sPath), VarPtr(KeyHandle)) '---###FIX!IT###---:KeyHandle--->VarPtr(KeyHandle)
    lResult = RegQueryValueExW(KeyHandle, StrPtr(sValue), 0&, VarPtr(lValueType), 0&, VarPtr(lDataBufSize)) '---###FIX!IT###---:jadikan pointer-ke-value dari variabel.

    If lValueType = REG_SZ Then
        sBuff = String$(lDataBufSize / 2, Chr$(32)) '---###FIX!IT###---:character unicode adalah 2 bytes perchar,fungsi string(param1,param2)menghitung char dengan: 1count=1char (di kode),1char=2bytes (di memori).
        lResult = RegQueryValueExW(KeyHandle, StrPtr(sValue), 0&, VarPtr(lValueType), StrPtr(sBuff), VarPtr(lDataBufSize)) '---###FIX!IT###---:lValueType harus sama.untuk fungsi "RegQueryValueExW" yang diminta adalah ukuran buffer dalam bytes,bukan chars.
        If lResult = ERROR_SUCCESS Then
            '---###FIX!IT###---:kalau ukuran buffer yang disyaratkan dan yang dialokasikan sama, sepertinya fungsi trimnullchars di bawah ini tidak terpakai lagi---:
            intZeroPos = InStr(sBuff, Chr$(0))
            If intZeroPos > 0 Then '---?tanya?TrimNullChars.
                GetStringValue = Replace(sBuff, Chr(0), "") 'Left$(sBuff, intZeroPos - 1)
            Else
                GetStringValue = sBuff
            End If
            '------------------;
            'GetStringValue = sBuff
        End If
    End If
    'MsgBox "TEST_TRIMMER:[" & GetStringValue & "]"
End Function

Public Function SetStringValue(ByVal hkey As Long, ByRef sPath As String, ByRef sValue As String, ByRef sData As String) As Long '---###FIX!IT###---:optimisasi variabel string,ganti ByRef aja,kalau cuman buat dibaca, bukan ditulis / dimodifikasi ulang.
    
    lReg = RegCreateKeyW(hkey, StrPtr(sPath), VarPtr(KeyHandle)) '---###FIX!IT###---:KeyHandle--->VarPtr(KeyHandle)
    lReg = RegSetValueExW(KeyHandle, StrPtr(sValue), 0, REG_SZ, StrPtr(sData), LenB(sData)) '---###FIX!IT###---:jadikan pointer-ke-value dari variabel.Len(?) jadi LenB(?),untuk fungsi "RegSetValueExW" yang diminta adalah ukuran buffer dalam bytes, bukan chars.
    lReg = RegCloseKey(KeyHandle)
    
End Function

Public Function SetExpandValue(ByVal hkey As Long, ByRef sPath As String, ByRef sValue As String, ByRef sData As String) As Long '---###FIX!IT###---:optimisasi variabel string,ganti ByRef aja,kalau cuman buat dibaca, bukan ditulis / dimodifikasi ulang.
    
    lReg = RegCreateKeyW(hkey, StrPtr(sPath), VarPtr(KeyHandle)) '---###FIX!IT###---:KeyHandle--->VarPtr(KeyHandle)
    lReg = RegSetValueExW(KeyHandle, StrPtr(sValue), 0, REG_EXPAND_SZ, StrPtr(sData), LenB(sData)) '---###FIX!IT###---:jadikan pointer-ke-value dari variabel.Len(?) jadi LenB(?),untuk fungsi "RegSetValueExW" yang diminta adalah ukuran buffer dalam bytes, bukan chars.
    lReg = RegCloseKey(KeyHandle)
    
End Function

Function GetDwordValue(ByVal hkey As Long, ByRef sPath As String, ByRef sValueName As String) As Long '---###FIX!IT###---:optimisasi variabel string,ganti ByRef aja,kalau cuman buat dibaca, bukan ditulis / dimodifikasi ulang.
    
    Dim lBuff As Long
    
    lReg = RegOpenKeyW(hkey, StrPtr(sPath), VarPtr(KeyHandle)) '---###FIX!IT###---:KeyHandle--->VarPtr(KeyHandle)
    lDataBufSize = 4 '?info?DWORD.
    lResult = RegQueryValueExW(KeyHandle, StrPtr(sValueName), 0&, VarPtr(lValueType), VarPtr(lBuff), VarPtr(lDataBufSize)) '---###FIX!IT###---:lBuff adalah variant-variable,bukan string.

    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            GetDwordValue = lBuff
        End If
    End If
    
    lReg = RegCloseKey(KeyHandle)
    
End Function

Public Function SetDwordValue(ByVal hkey As Long, ByRef sPath As String, ByRef sValueName As String, ByVal lData As Long) As Long '---###FIX!IT###---:optimisasi variabel string,ganti ByRef aja,kalau cuman buat dibaca, bukan ditulis / dimodifikasi ulang.
    
    lReg = RegCreateKeyW(hkey, StrPtr(sPath), VarPtr(KeyHandle)) '---###FIX!IT###---:KeyHandle--->VarPtr(KeyHandle)
    lResult = RegSetValueExW(KeyHandle, StrPtr(sValueName), 0&, REG_DWORD, VarPtr(lData), 4) '---###FIX!IT###---:jadikan pointer-ke-value dari variabel.
    lReg = RegCloseKey(KeyHandle)
    
End Function

Public Function DeleteKey(ByVal hkey As Long, ByRef sKey As String) As Long '---###FIX!IT###---:optimisasi variabel string,ganti ByRef aja,kalau cuman buat dibaca, bukan ditulis / dimodifikasi ulang.
    lReg = RegDeleteKeyW(hkey, StrPtr(sKey))
End Function

Public Function DeleteValue(ByVal hkey As Long, ByRef sPath As String, ByRef sValue As String) As Long '---###FIX!IT###---:optimisasi variabel string,ganti ByRef aja,kalau cuman buat dibaca, bukan ditulis / dimodifikasi ulang.
    lReg = RegOpenKeyW(hkey, StrPtr(sPath), VarPtr(KeyHandle)) '---###FIX!IT###---:KeyHandle--->VarPtr(KeyHandle)
    lReg = RegDeleteValueW(KeyHandle, StrPtr(sValue))
    lReg = RegCloseKey(KeyHandle)
End Function

Public Function RegEnumStr(ByVal MainKey As Long, ByRef sPath As String, ByRef sValue() As String, ByRef sData() As String) As Long
On Error Resume Next
Dim iHKey As Long, iHasil As Long, Num As Long, ValLen As Long '<---biasakan inisialisasi variabel dengan mencantumkan tipe variabel yang jelas.
Dim nOutDataLength  As Long
Dim szOutDataValue As String
Dim StrValue As String
Dim pOpType As Long

iHasil = RegOpenKeyExW(MainKey, StrPtr(sPath), 0, KEY_READ, VarPtr(iHKey))
If iHasil <> 0& Then ' 0& = ERROR_SUCCESS
    Exit Function
End If
Num = 0
ReDim sValue(100) As String
ReDim sData(100) As String

pOpType = REG_SZ
Do
    ValLen = 2048  ' Penampung Panjang Value Maximal aja
    StrValue = String$(ValLen, 0)
    nOutDataLength = 2048
    szOutDataValue = String$(nOutDataLength, 0)
    iHasil = RegEnumValueW(iHKey, Num, StrPtr(StrValue), VarPtr(ValLen), 0&, VarPtr(pOpType), StrPtr(szOutDataValue), VarPtr(nOutDataLength))
    If iHasil = 0& Then
        Num = Num + 1
        StrValue = Left$(StrValue, ValLen) '---dalam chars.
        szOutDataValue = Left$(szOutDataValue, (nOutDataLength / 2) - 1) '---dalam bytes,nullchars dibuang.
        sValue(Num) = StrValue
        sData(Num) = szOutDataValue
        '---------------------------;
    End If
Loop While iHasil = 0& ' 0& = ERROR_SUCCESS
    RegEnumStr = Num
    StrValue = vbNullString
    szOutDataValue = vbNullString
Call RegCloseKey(iHKey)
End Function

'---Catatan
' Belum bisa buffer RUNDLL.EXE PathDll,Param -> ah kurang penting untuk mendapatkan startup virus di reg
Public Function BufferStartupPath(sFile As String) As String
Dim sTmp        As String
Dim sTmp2       As String
Dim sSpecial    As String
Dim sFirstFol   As String
Dim sPathCad    As String
Dim nNum        As Long
Dim iCount      As Long

If ValidFile(sFile) = False Then
    ' dapatkan awal dari drive:\
    If InStr(sFile, ":\") > 0 Then sTmp = Mid$(sFile, InStr(sFile, ":\") - 1)
    sTmp = Replace(sTmp, Chr(34), "")
    If ValidFile(sTmp) = True Then GoTo KLIMAKS
    
    If InStr(sFile, Chr(34)) > 0 Then
        sTmp2 = Replace(sFile, Chr(34), "")
        sPathCad = Right$(sFile, Len(sFile) - InStrRev(sFile, Chr(34)))
        sTmp2 = Left$(sTmp, Len(sTmp) - Len(sPathCad))
        If ValidFile(sTmp2) = True Then sTmp = sTmp2: GoTo KLIMAKS
    End If
    ' Hilangkan /[param] --- contoh C:\Memeil\Jelek.exe /s
    nNum = InStr(sFile, "/")
    If nNum > 0 Then
       sTmp = Left$(sFile, nNum - 1)
    Else
       ' Hilangkan -[param] --- contoh C:\Memeil\Jelek.exe -start
       nNum = InStr(StrReverse(sFile), "-") ' ambil terkanan pertama [karena namafile atau folder boleh -]
       If nNum > 0 Then
          sTmp = Left$(sFile, Len(sFile) - nNum)
       Else
          sTmp = sFile
       End If
    End If
    
    Do
        'jika ada spasi terkanan hilangkan
        If Right(sTmp, 1) = Chr(32) Then
            sTmp = Left$(sTmp, Len(sTmp) - 1)
        Else
            sTmp = sTmp
        End If
    Loop While Right$(sTmp, 1) = Chr(32) ' hapus sampai char terkanan bukan spasi

    ' klo ada chr(34) / [""] -- buang aj
    sTmp = Replace(sTmp, Chr(34), "")
    
    '-------Lalu buat jaga-jaga kalo auto path misal SOUNDMAN.EXE
    If InStr(sTmp, "\") = 0 Then
       sSpecial = GetSpecFolder(WINDOWS_DIR) ' coba di windows dulu
       If ValidFile(sSpecial & "\" & sTmp) = True Then
          sTmp = sSpecial & "\" & sTmp
       Else
          sSpecial = GetSpecFolder(SYSTEM_DIR) ' coba di system32 sekarang
          If ValidFile(sSpecial & "\" & sTmp) = True Then
             sTmp = sSpecial & "\" & sTmp
          End If
       End If
    End If
    sFirstFol = GetPath(sTmp, FirstFolder)
    If ValidFile(GetSpecFolder(PROGRAM_FILE) + Right$(sTmp, Len(sTmp) - Len(sFirstFol))) = True Then
    sTmp = GetSpecFolder(PROGRAM_FILE) + Right$(sTmp, Len(sTmp) - Len(sFirstFol))
    ElseIf ValidFile(GetSpecFolder(SYSTEM_DIR) + Right$(sTmp, Len(sTmp) - Len(sFirstFol))) = True Then
    sTmp = GetSpecFolder(SYSTEM_DIR) + Right$(sTmp, Len(sTmp) - Len(sFirstFol))
    ElseIf ValidFile(GetSpecFolder(WINDOWS_DIR) + Right$(sTmp, Len(sTmp) - Len(sFirstFol))) = True Then
    sTmp = GetSpecFolder(WINDOWS_DIR) + Right$(sTmp, Len(sTmp) - Len(sFirstFol))
    End If
    
    
    If ValidFile(sTmp) = True Then sTmp = sTmp Else sTmp = "Failed !"
Else
    sTmp = sFile
End If

KLIMAKS:
' Klimaks :D
BufferStartupPath = sTmp
End Function


' Buat Enum startup
Public Function EnumRegStartup(ByRef sFileStart() As String, bWithCommon As Boolean) As Long
Dim nJum         As Long
Dim nLong        As Long
Dim nStart       As Long
Dim nCount       As Long
Dim sName        As String
Dim sFile        As String
Dim ArrFile()    As String
Dim sPathReg(7)  As String
Dim sKeyRegN(7)  As String
Dim sValueName() As String
Dim sValueData() As String


ReDim sFileStart(100) As String ' karena blum tahu secara pasti berap Startup-nya

sKeyRegN(0) = "HKCU"
sPathReg(0) = "Software\Microsoft\Windows\CurrentVersion\Run"

sKeyRegN(1) = "HKLM"
sPathReg(1) = "Software\Microsoft\Windows\CurrentVersion\Run"

sKeyRegN(2) = "HKLM"
sPathReg(2) = "Software\Microsoft\Windows\CurrentVersion\Run-"

sKeyRegN(3) = "HKLM"
sPathReg(3) = "Software\Microsoft\Windows\CurrentVersion\RunOnce"

sKeyRegN(4) = "HKLM"
sPathReg(4) = "Software\Microsoft\Windows\CurrentVersion\RunOnceEx"


For nStart = 0 To 4
    nJum = RegEnumStr(SingkatanKey(sKeyRegN(nStart)), sPathReg(nStart), sValueName(), sValueData())
    For nLong = 1 To nJum
        sFile = BufferStartupPath(sValueData(nLong))
        If ValidFile(sFile) = True Then
           sFileStart(nCount) = sFile
           nCount = nCount + 1
        End If
    Next
    ' hayoo habis dipakai diset ulang dulu
    nLong = 1
    Erase sValueName
    Erase sValueData
Next
' Ditambah Reg Start-Up Singgle
sPathReg(5) = "SOFTWARE\microsoft\Windows NT\CurrentVersion\Winlogon"
sKeyRegN(5) = "HKLM"
sFile = GetStringValue(SingkatanKey(sKeyRegN(5)), sPathReg(5), "Shell")
If UCase(sFile) <> "EXPLORER.EXE" Then ' berarti ada tuh
   sFile = Mid$(sFile, InStr(sFile, Chr(32)) + 1)
   If ValidFile(sFile) = True Then
      sFileStart(nCount) = sFile
      nCount = nCount + 1
   End If
End If

sPathReg(6) = "SOFTWARE\microsoft\Windows NT\CurrentVersion\Winlogon"
sKeyRegN(6) = "HKLM"
sFile = GetStringValue(SingkatanKey(sKeyRegN(6)), sPathReg(6), "Userinit")
sName = GetSpecFolder(SYSTEM_DIR) & "\userinit.exe"
If UCase(sFile) <> UCase(sName) Then  ' berarti ada tuh
   sFile = Replace(UCase(sFile), UCase(sName) & ",", "")
   sFile = BuangSpaceAwal(sFile)
   If ValidFile(sFile) = True Then
      sFileStart(nCount) = sFile
      nCount = nCount + 1
   End If
End If

sPathReg(7) = "SOFTWARE\microsoft\Windows NT\CurrentVersion\Windows"
sKeyRegN(7) = "HKLM"
sFile = GetStringValue(SingkatanKey(sKeyRegN(7)), sPathReg(7), "Load")
If ValidFile(sFile) = True Then
   sFileStart(nCount) = sFile
   nCount = nCount + 1
End If

If bWithCommon = True Then
   nJum = GetFile(GetSpecFolder(USER_STARTUP), ArrFile)
   For nLong = 1 To nJum
       sFileStart(nCount) = ArrFile(nLong - 1)
       nCount = nCount + 1
   Next
   nLong = 1 ' reset
   nJum = GetFile(GetSpecFolder(ALL_USER_STARTUP), ArrFile)
   For nLong = 1 To nJum
       sFileStart(nCount) = ArrFile(nLong - 1)
       nCount = nCount + 1
   Next

End If

EnumRegStartup = nCount

End Function

Public Function SingkatanKey(sKey As String) As Long

Select Case sKey
    Case "HKCR"
        SingkatanKey = &H80000000
    Case "HKCU"
        SingkatanKey = &H80000001
    Case "HKLM"
        SingkatanKey = &H80000002
    Case "HKU"
        SingkatanKey = &H80000003
End Select

End Function

Public Sub GetRegStartup(ByRef Lv As ucListView)
Dim nJum         As Long
Dim nLong        As Long
Dim nStart       As Long
Dim lngItem      As Long
Dim sFile        As String
Dim sPathReg(7)  As String
Dim sKeyRegN(7)  As String
Dim sValueName() As String
Dim sValueData() As String
Dim sName        As String
Dim NamaVrz      As String
Dim sStatus      As String
Dim AUStartUp    As String
Dim UStartup     As String
Dim FSO          As Object
Dim FileNow      As String
Dim sFileUS      As Object
Dim sEvaluType   As String
Dim sDisCUPath   As String
Dim sDisAUPath   As String
Set cImgList = New gComCtl

sDisCUPath = App.Path & "\Plus Fitur\Startup Manager\Disable-CU"
sDisAUPath = App.Path & "\Plus Fitur\Startup Manager\Disable-AU"
sKeyRegN(0) = "HKCU"
sPathReg(0) = "Software\Microsoft\Windows\CurrentVersion\Run"

sKeyRegN(1) = "HKLM"
sPathReg(1) = "Software\Microsoft\Windows\CurrentVersion\Run"

sKeyRegN(2) = "HKLM"
sPathReg(2) = "Software\Microsoft\Windows\CurrentVersion\Run-"

sKeyRegN(3) = "HKLM"
sPathReg(3) = "Software\Microsoft\Windows\CurrentVersion\RunOnce"

sKeyRegN(4) = "HKLM"
sPathReg(4) = "Software\Microsoft\Windows\CurrentVersion\RunOnceEx"

sKeyRegN(5) = "HKCU"
sPathReg(5) = "Software\Microsoft\Windows\CurrentVersion\Run-"

Set Lv.ImageList = cImgList.NewImageList(16, 16, imlColor32)

With Lv
.ListItems.Clear
For nStart = 0 To 5
    nJum = RegEnumStr(SingkatanKey(sKeyRegN(nStart)), sPathReg(nStart), sValueName(), sValueData())
    For nLong = 1 To nJum
        sFile = BufferStartupPath(sValueData(nLong))
        If ValidFile(sFile) = True Then
           DrawIco sFile, frMain.picBuff, ricnSmall
           Lv.ImageList.AddFromDc frMain.picBuff.hdc, 16, 16
           If nStart = 2 Or nStart = 5 Then
            sStatus = "Disable"
           Else
            sStatus = "Enable"
           End If
           If sValueName(nLong) = "Bee Antivirus" Or sValueName(nLong) = "avast" Then
            sEvaluType = "Necessary"
           Else
            sEvaluType = "Optional"
           End If
           Lv.ListItems.Add , sValueName(nLong), , (Lv.ImageList.IconCount - 1), , , , , Array(sKeyRegN(nStart) & "\" & sPathReg(nStart), sValueData(nLong), sFile, sEvaluType, sStatus)
        Else
           WriteFileUniSim GetSpecFolder(USER_DOC) & "\XXXXXXXX", "Unne"
           DrawIco GetSpecFolder(USER_DOC) & "\XXXXXXXX", frMain.picBuff, ricnSmall
           HapusFile GetSpecFolder(USER_DOC) & "\XXXXXXXX"
           Lv.ImageList.AddFromDc frMain.picBuff.hdc, 16, 16
           If nStart = 2 Or nStart = 5 Then
            sStatus = "Disable"
           Else
            sStatus = "Enable"
           End If
            sEvaluType = "Unnecessary"
           Lv.ListItems.Add , sValueName(nLong), , (Lv.ImageList.IconCount - 1), , , , , Array(sKeyRegN(nStart) & "\" & sPathReg(nStart), sValueData(nLong), sFile, sEvaluType, sStatus)
        End If
    Next
Next
UStartup = GetSpecFolder(USER_STARTUP)
AUStartUp = GetSpecFolder(ALL_USER_STARTUP)

On Error GoTo KELUAR:
    Set FSO = CreateObject("Scripting.FileSystemObject")
    For Each sFileUS In FSO.GetFolder(UStartup).Files
        DoEvents
        FileNow = sFileUS
        If LCase(GetPath(FileNow, FileName)) <> "desktop.ini" Then
            DrawIco FileNow, frMain.picBuff, ricnSmall
            Lv.ImageList.AddFromDc frMain.picBuff.hdc, 16, 16
            Lv.ListItems.Add , GetPath(FileNow, JustName), , (Lv.ImageList.IconCount - 1), , , , , Array(UStartup, "None", FileNow, "Optional", "Enable")
        End If
    Next

    Set FSO = CreateObject("Scripting.FileSystemObject")
    For Each sFileUS In FSO.GetFolder(AUStartUp).Files
        DoEvents
        FileNow = sFileUS
        If LCase(GetPath(FileNow, FileName)) <> "desktop.ini" Then
            DrawIco FileNow, frMain.picBuff, ricnSmall
            Lv.ImageList.AddFromDc frMain.picBuff.hdc, 16, 16
            Lv.ListItems.Add , GetPath(FileNow, JustName), , (Lv.ImageList.IconCount - 1), , , , , Array(AUStartUp, "None", FileNow, "Optional", "Enable")
        End If
    Next
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    For Each sFileUS In FSO.GetFolder(sDisCUPath).Files
        DoEvents
        FileNow = sFileUS
        If LCase(GetPath(FileNow, FileName)) <> "desktop.ini" Then
            DrawIco FileNow, frMain.picBuff, ricnSmall
            Lv.ImageList.AddFromDc frMain.picBuff.hdc, 16, 16
            Lv.ListItems.Add , GetPath(FileNow, JustName), , (Lv.ImageList.IconCount - 1), , , , , Array(UStartup, "None", FileNow, "Optional", "Disable")
        End If
    Next
    Set FSO = CreateObject("Scripting.FileSystemObject")
    For Each sFileUS In FSO.GetFolder(sDisAUPath).Files
        DoEvents
        FileNow = sFileUS
        If LCase(GetPath(FileNow, FileName)) <> "desktop.ini" Then
            DrawIco FileNow, frMain.picBuff, ricnSmall
            Lv.ImageList.AddFromDc frMain.picBuff.hdc, 16, 16
            Lv.ListItems.Add , GetPath(FileNow, JustName), , (Lv.ImageList.IconCount - 1), , , , , Array(AUStartUp, "None", FileNow, "Optional", "Disable")
        End If
    Next
KELUAR:
End With
End Sub

