Attribute VB_Name = "basConfig"
Public Enum HKEYS_Constants
  HKEY_CLS_ROOT = &H80000000
  HKEY_CUR_USER = &H80000001
  HKEY_LOC_MAC = &H80000002
  HKEY_USERZ = &H80000003
  HKEY_PERF_DATU = &H80000004
End Enum

Private ExKey As HKEYS_Constants
Private ExPathFiles As String
Private ExPathFolders As String
Private ExPathShortCuts As String
Dim xSXor As classSimpleXOR
Dim xHuff As classHuffman
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Boolean) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long

Public Sub DisableClose(Frm As Form, Optional Disable As Boolean = True)

    Dim hMenu As Long
    Dim nCount As Long
    
    If Disable Then
        hMenu = GetSystemMenu(Frm.hWnd, False)
        nCount = GetMenuItemCount(hMenu)
        
        Call RemoveMenu(hMenu, nCount - 1, MF_REMOVE Or MF_BYPOSITION)
        Call RemoveMenu(hMenu, nCount - 2, MF_REMOVE Or MF_BYPOSITION)
    
        DrawMenuBar Frm.hWnd
    Else
        GetSystemMenu Frm.hWnd, True
        DrawMenuBar Frm.hWnd
    End If

End Sub

Public Function GetSettingF(Path As String) As Boolean
Static hFileP As Long, OutConfig() As Byte, sTempC As Long, IsiFile As String, tmpC() As String, count As Long
hFileP = GetHandleFile(Path)
GetSettingF = False
'On Error GoTo l_AkhirC
If ValidFile(Path) = False Then Exit Function
    GetSettingF = True
    Call ReadUnicodeFile2(hFileP, 1, FileLen(Path), OutConfig())
    Set xSXor = New classSimpleXOR
    Call xSXor.DecryptByte(OutConfig(), "DAXAAngKOt")
    Set xSXor = Nothing
    Set xHuff = New classHuffman
    Call xHuff.DecodeByte(OutConfig(), UBound(OutConfig()) + 1)
    Set xHuff = Nothing
IsiFile = StrConv(OutConfig, vbUnicode)
tmpC = Split(IsiFile, Chr(13))
sTempC = UBound(tmpC())
count = 1
For count = 1 To sTempC
    frMain.ck(count).value = tmpC(count - 1)
Next
l_AkhirC:
TutupFile hFileP
End Function
Public Sub SaveSettingF(Path As String)
Static IsiSetting As String, OutConfig() As Byte, sHfileP As Long
sHfileP = GetHandleFile(Path)
IsiSetting = frMain.ck(1).value & Chr(13) & frMain.ck(2).value & Chr(13) & frMain.ck(3).value & Chr(13) & frMain.ck(4).value & Chr(13) & frMain.ck(5).value & Chr(13) & frMain.ck(6).value & Chr(13) & frMain.ck(7).value & Chr(13) & frMain.ck(8).value & Chr(13)
WriteFileUniSim Path, IsiSetting
Call ReadUnicodeFile2(sHfileP, 1, FileLen(Path), OutConfig())
Set xHuff = New classHuffman
Call xHuff.EncodeByte(OutConfig(), UBound(OutConfig()) + 1)
Set xHuff = Nothing
Set xSXor = New classSimpleXOR
Call xSXor.EncryptByte(OutConfig(), "DAXAAngKOt")
Set xSXor = Nothing
WriteUnicodeFile Path, 1, OutConfig()
TutupFile sHfileP
End Sub

Public Function RunSU(Index As Integer)
If Index = 1 Then
    SetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Bee Antivirus", App.Path & "\" & App.EXEName & ".exe /U %1Å"
frMain.ck(4).value = 1
frRTP.mnROS.Checked = True
Else
    DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Bee Antivirus"
frMain.ck(4).value = 0
frRTP.mnROS.Checked = False
End If
End Function

Public Function ApplySetting()
With frMain
    If .ck(4).value = 1 Then
        RunSU 1
    Else
        RunSU 0
    End If
    
    If .ck(5).value = 1 Then
        Install_CMenu
    Else
        Uninstall_CMenu
    End If

    If .ck(6).value = 1 Then
        frMain.tmFD.Enabled = True
    Else
        frMain.tmFD.Enabled = False
    End If
    
    If .ck(7).value = 1 Then
        ScanRTPmod = True
        frRTP.mnEP.Checked = True
    Else
        ScanRTPmod = False
        frRTP.mnEP.Checked = False
    End If
    
End With
End Function

Public Sub Install_CMenu() ' shell menu
  CreateKeyReg &H80000001, "Software\CI-Soft Software\Bee Antivirus"
  SetDwordValue &H80000001, "Software\CI-Soft Software\Bee Antivirus", "ContextMenu", "1"
  SetStringValue &H80000001, "Software\CI-Soft Software\Bee Antivirus", "Exe", App.Path & "\" & App.EXEName & ".exe"
  Shell "regsvr32.exe /s " & Chr(34) & App.Path & "\Bee Shell Extension.dll" & Chr(34), vbHide
  RegData
End Sub

Public Sub Uninstall_CMenu() ' cabut shell menu
  DeleteKey &H80000001, "Software\CI-Soft Software\Bee Antivirus"
  DeleteKey &H80000000, "*\shellex\ContextMenuHandlers\Bee - Av"
  DeleteKey &H80000000, "Folder\shellex\ContextMenuHandlers\Bee - Av"
  DeleteKey &H80000000, "lnkfile\shellex\ContextMenuHandlers\Bee - Av"
  Shell "regsvr32.exe /u /s " & Chr(34) & App.Path & "\Bee Shell Extension.dll" & Chr(34), vbHide
End Sub

Private Sub RegData()
  Dim cr As New cRegistry
  Dim sKey() As String, lCount As Long, ZX As Long
  Dim sVal As String
  
  Dim sTemp As String
  ExKey = HKEY_CLS_ROOT
  ExPathFiles = "*\shellex\ContextMenuHandlers\Bee - Av"
  ExPathFolders = "Folder\shellex\ContextMenuHandlers\Bee - Av"
  ExPathShortCuts = "lnkfile\shellex\ContextMenuHandlers\Bee - Av"
  cr.ClassKey = HKEY_CLASSES_ROOT
  cr.SectionKey = "CLSID"
  
  cr.EnumerateSections sKey(), lCount
  
  For ZX = 1 To lCount

    sTemp = "CLSID\" & sKey(ZX)
    sVal = GetStringValue(ExKey, sTemp, "")
    If sVal = "Bee Shell Extension" Then
    
      Debug.Print sKey(ZX)
      SetStringValue ExKey, ExPathFiles, "", sKey(ZX)
      SetStringValue ExKey, ExPathFolders, "", sKey(ZX)
      SetStringValue ExKey, ExPathShortCuts, "", sKey(ZX)
      Exit Sub
      
    End If
    
  Next
  
  
  Set cr = Nothing
End Sub
