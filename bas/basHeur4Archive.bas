Attribute VB_Name = "basHeur4Archive"
Public IsiDataArc() As Byte

Private Function GetChecksum(PackedData() As Byte, ArcSize As Long) As String
On Error Resume Next
Dim iCount      As Long

If ArcSize <= 0 Then
    GetChecksum = vbNullString
    Exit Function
End If

' Walapun 300 tapi variatif
If ArcSize > 4300 Then
    Call ReadUniFileArc(PackedData, 4000, 300, ArcSize)
Else
    '[APTX] hilangkan salin antar variable yang nggak perlu
    If TheSize > 300 Then
        Call ReadUniFileArc(PackedData, 1, 300, ArcSize)
    Else
        Call ReadUniFileArc(PackedData, 1, CStr(ArcSize), ArcSize)
    End If
End If

For iCount = 0 To 299 Step 10
    GetChecksum = GetChecksum & Hex$(IsiDataArc(iCount))
Next iCount

GetChecksum = StrReverse(GetChecksum) ' dibalik  supaya variatif

'If MYCeksum = String(Len(MYCeksum), "0") Then MYCeksum = "Tolong pakai ceksum Cadangan !"
Erase DataOut
End Function

Private Function GetChecksum2(sPackedData() As Byte) As String
On Error Resume Next
Dim sTmpFile As String
Dim IsiFile  As String
Dim iCount   As Integer


sTmpFile = StrConv(sPackedData(), vbUnicode)

If Len(sTmpFile) = 0 Then
    GetChecksum2 = vbNullString
    Exit Function
End If

If Len(sTmpFile) > 4300 Then
    IsiFile = Mid(sTmpFile, 4000)
Else
    IsiFile = sTmpFile
End If

For iCount = 1 To 300 Step 10
    GetChecksum2 = GetChecksum2 & Hex$(Asc(Mid(IsiFile, iCount, 1)))
Next iCount

GetChecksum2 = StrReverse(GetChecksum2) ' dibalik  supaya variatif
End Function

Private Function GetChecksumCadangan(tPackedData() As Byte, TheSize As Long) As String
On Error Resume Next
Dim DataOut()   As Byte
Dim TmpCount    As Long
Dim iCount      As Long

If TheSize <= 0 Then
    GetChecksumCadangan = vbNullString
    Exit Function
End If

If TheSize > 400 Then ' ambil lebh banyak krena hanya cdangan
    Call ReadUniFileArc(tPackedData, 1, 400, TheSize)
Else
    Call ReadUniFileArc(tPackedData, 1, CStr(TheSize), TheSize)
End If

For iCount = 0 To 199
    TmpCount = TmpCount + IsiDataArc(iCount) ^ 2.2
Next iCount

GetChecksumCadangan = Hex$(TmpCount)
TmpCount = 0

For iCount = 200 To 399
    TmpCount = TmpCount + IsiDataArc(iCount) ^ 2.2
Next iCount

GetChecksumCadangan = GetChecksumCadangan & Hex$(TmpCount)

Erase DataOut
End Function

Private Function CekByString(xData As String, sNameFile As String) As String
Static Udata As String
Udata = UCase(xData)
sNameFile = UCase(sNameFile)
Select Case Left(xData, 2)
Case "MZ"
    If InStr(Udata, "IMISSYOU") > 0 Then
    CekByString = "Win32/Runonce"
    Exit Function
    End If
    
    If InStr(xData, "µí§¶ýÚÿ×Ðþÿÿ·hþÿÿÿÿÿï¡ùÿÿÿÿÿÿÿÿÿÿÿÿ€") > 0 Then
    CekByString = "Win32/Alman.A"
    Exit Function
    End If
    
    If InStr(xData, "4xÛ 35‰úPC§ãàn†¡út‚t(ZŠð ÐøÈÔ¯éú²/") > 0 Then
    CekByString = "Win32/Alman.B"
    Exit Function
    End If
    
    If InStr(xData, "¨¶ÁX") > 0 Then
    CekByString = "Win32/Alman.B"
    Exit Function
    End If
    
    If InStr(xData, "D’Rich_D") > 0 Or InStr(xData, "D’RichßD’") > 0 Then
    CekByString = "Win32/Oliga"
    Exit Function
    End If
    
    If InStr(xData, "(ÚI") > 0 Then
    CekByString = "Win32/Sysyer"
    Exit Function
    End If
    
    If InStr(xData, "àhtnrnog") > 0 Then
    CekByString = "Win32/Vitro"
    Exit Function
    End If
    
    If InStr(xData, "Ò@Òæ¿\ÒRichç¿\Ò") > 0 Then
    CekByString = "Win32/Ramnit.L"
    Exit Function
    End If
    
    If InStr(xData, "RRSPè”") > 0 Then
    CekByString = "Win32/TeddyBear"
    Exit Function
    End If
    
    If InStr(xData, "Že©óž=9¨lßìnßdYôl") > 0 Then
    CekByString = "Win32/Alman.B"
    Exit Function
    End If
    
    If InStr(xData, "Êh¡7Tívœ?") > 0 Then
    CekByString = "Win32/Service"
    Exit Function
    End If
    
    If InStr(xData, "`ª`oMØ!5·=") > 0 Then
    CekByString = "Win32/Runonce"
    Exit Function
    End If
    
    If InStr(xData, "X5O!P%@AP") > 0 Then
    CekByString = "Eicar Not Virus !!!"
    Exit Function
    End If
    
    If InStr(xData, "xµT’h©Lþ") > 0 Then
    CekByString = "Win32/Service"
    Exit Function
    End If
    
    If InStr(xData, "è³¶ûÿ‰Eð3Ò") > 0 Then
    CekByString = "Win32/Spooler"
    Exit Function
    End If
    
    If InStr(xData, "H") = 529 And InStr(xData, "w") = 541 And InStr(xData, "!") = 551 And InStr(xData, "C") = 557 Then
    CekByString = "Win32/Sality"
    Exit Function
    End If
Case "MS"
    If InStr(Udata, "MSFT") > 0 Then
    CekByString = "Win32:Malware-Gen"
    Exit Function
    End If
    
Case "HE"
    If InStr(Udata, "IMISSYOU") > 0 Then
    CekByString = "FakeMail.R[Vrt]"
    Exit Function
    End If
    
End Select

If isProperFile(sNameFile, "HTML HTM HTA") = True Then
If InStr(Udata, "IMISSYOU") > 0 Then
CekByString = "Infected HTML"
Exit Function
End If
End If

If isProperFile(sNameFile, "INF") = True Then
If InStr(Udata, "[AUTORUN]") > 0 And InStr(Udata, ".PIF") > 0 Then
CekByString = "Malware Autorun"
Exit Function
End If
End If

If isProperFile(sNameFile, "INF") = True Then
If InStr(Udata, "[AUTORUN]") > 0 And InStr(Udata, ".BAT") > 0 Then
CekByString = "Trojan.Autorun.lnk"
Exit Function
End If
End If

If isProperFile(sNameFile, "INF") = True Then
If InStr(Udata, "[AUTORUN]") > 0 And InStr(Udata, ".CMD") > 0 Then
CekByString = "Trojan.Autorun.Ink"
Exit Function
End If
End If

If isProperFile(sNameFile, "INF") = True Then
If InStr(Udata, "[AUTORUN]") > 0 And InStr(Udata, "SHELL") > 0 Then
CekByString = "Trojan.Autorun.Ink"
Exit Function
End If
End If

If isProperFile(sNameFile, "INF") = True Then
If InStr(Udata, "[AUTORUN]") > 0 And InStr(Udata, "OPEN\COMMAND") > 0 Then
CekByString = "Malware Autorun"
Exit Function
End If
End If

If isProperFile(sNameFile, "INF") = True Then
If InStr(Udata, "[AUTORUN]") > 0 And InStr(Udata, "EXPLORE\COMMAND") > 0 Then
CekByString = "Trojan.Autorun.Ink"
Exit Function
End If
End If

If isProperFile(sNameFile, "INF") = True Then
If InStr(Udata, "WSCRIPT") > 0 Then
CekByString = "Trojan.Autorun.Ink"
Exit Function
End If
End If

If isProperFile(sNameFile, "LNK") = True Then
If InStr(Udata, ".SCR") > 0 Or InStr(Udata, ".VBS") > 0 Or InStr(Udata, ".DB") > 0 Or InStr(Udata, "//E:VBSCRIPT") > 0 Then
CekByString = "Gen-LnkVirut[lnk]"
Exit Function
End If
End If

If isProperFile(sNameFile, "BAT CMD") = True Then
If InStr(Udata, "KILL EXPLORER.EXE") > 0 Or InStr(Udata, "DEL EXPLORER.EXE") > 0 Or InStr(Udata, "ASSOC .EXE=") > 0 Or InStr(Udata, "TASKKILL") > 0 Then
CekByString = "Gen-BotVirut[Bsc]"
Exit Function
End If
End If

If isProperFile(sNameFile, "VBS BAT INI .DB VBE COM") = True Then
If InStr(Udata, "REGWRITE") > 0 And InStr(Udata, "WSCRIPT") > 0 Then
CekByString = "Gen-ScriptVirut[Scp]"
Exit Function
End If
End If

If isProperFile(sNameFile, "VBS BAT INI .DB VBE COM") = True Then
If InStr(Udata, "WSCRIPT.SCRIPTFULLNAME") > 0 Or InStr(Udata, "FORMAT C:\") > 0 Or InStr(Udata, "FORMAT D:\") > 0 Or InStr(Udata, "AUTOEXEC.BAT") > 0 Then
CekByString = "Gen-ScriptVirut[Scp]"
Exit Function
End If
End If

If isProperFile(sNameFile, "VBS BAT INI .DB VBE COM") = True Then
If InStr(Udata, "VERSION\RUN") > 0 And InStr(Udata, ".BAT") > 0 Then
CekByString = "Gen-ScriptVirut[BAT]"
Exit Function
End If
End If

If isProperFile(sNameFile, "VBS BAT INI .DB VBE COM") = True Then
If InStr(Udata, "VERSION\RUN") > 0 And InStr(Udata, ".CMD") > 0 Then
CekByString = "Gen-ScriptVirut[cmd]"
Exit Function
End If
End If

If isProperFile(sNameFile, "VBS BAT INI .DB VBE COM") = True Then
If InStr(Udata, "VERSION\RUN") > 0 And InStr(Udata, ".VBS") > 0 Then
CekByString = "Gen-ScriptVirut[Vbs]"
Exit Function
End If
End If

If isProperFile(sNameFile, "VBS BAT INI .DB VBE COM") = True Then
If InStr(Udata, "VERSION\RUN") > 0 And InStr(Udata, ".PIF") > 0 Then
CekByString = "Gen-ScriptVirut[Wrm]"
Exit Function
End If
End If

If InStr(xData, ",$YGøF#G(") > 0 Then
CekByString = "Win32/KSplood"
Exit Function
End If

If UCase(Right(sNameFile, 4)) = ".VBS" Then
   Pos_Akhir = Len(xData)
   For iCount = 1 To Pos_Akhir
       AscKar = Asc(Mid(xData, iCount, 1))
       If AscKar >= 32 And AscKar <= 57 Then
          JumNumer = JumNumer + 1
       Else
          JumKar = JumKar + 1
       End If
   DoEvents
   Next
   If JumNumer > JumKar Then CekByString = "Malware Script": Exit Function
End If

CekByString = vbNullString
End Function

Private Function ReadUniFileArc(xPackedData() As Byte, sStartIn As String, sLeght As String, ArcSize As Long) As String
sStartIn = sStartIn - 1
IsiDataArc = Right(IsiDataArc, ArcSize - sStartIn)
IsiDataArc = Left(IsiDataArc, sLeght)
End Function

Public Function EqualArc(xPackDat() As Byte, ArcSizeFile As Long, NamaArc As String, xPath As String)
Static sSum As String
Static sD As String
Static iCount As Long
Static RetVirus As String

RetVirus = CekByString(sD, NamaArc)
sD = StrConv(xPackDat, vbUnicode)
If RetVirus <> vbNullString Then
AddToLVMal frMain.lvMal, RetVirus, xPath & "|> " & NamaArc, CStr(ArcSizeFile), "Valid\ Not Deleted", 0, 18
Exit Function
End If
Rem Mohon Maaf
Rem Fungsi Checksum Belum Bisa
'sSum = GetChecksum(xPackDat, ArcSizeFile)

'   If sSum = String(Len(sSum), "0") Then
'      sSum = GetChecksumCadangan(xPackDat, ArcSizeFile)
'   End If

'iCount = 1
'For iCount = 1 To xJumChecksum
'    If xNumChecksum(iCount) = sSum Then
'    MsgBox sSum
'        AddToLVMal frMain.lvMal, xNamChecksum(iCount), xPath & "|> " & NamaArc, CStr(ArcSizeFile), "Valid\ Not Deleted", 0, 18
'        VirusStatus = True
'        Exit For
'    End If
'Next
End Function
