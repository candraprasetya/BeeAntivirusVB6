Attribute VB_Name = "basScript"
Dim IsiVBS As String
Public xDataVBS As String
Public xDataVBS2 As String
Public xDataVBS3 As String
Public xDataVBS4 As String

Public Function BacaFileText(sFile As String) As String
Static sTemp As String
Static sTmp As String
sTmp = ""
Open sFile For Input As #1
Do While Not (EOF(1))
    Input #1, sTemp
    sTmp = sTmp & sTemp
Loop
Close #1
BacaFileText = Mid(sTmp, 3)
End Function

Public Function MalScrCheck(sFile As String, hFile As Long, MySize As Long) As Boolean
'On Error GoTo Cek_E
Dim JumNumer    As Long
Dim iCount      As Long
Dim JumKar      As Long
Dim AscKar      As Byte
Dim Pos_Akhir   As Long

Dim OutData()   As Byte
Dim OutData2()  As Byte

If isProperFile(sFile, "VBS BAT INI LNK COM .DB VBE CMD") = False Then MalScrCheck = False: Exit Function
   If MySize > 9000 Then
      Call ReadUnicodeFile2(hFile, 1, 4500, OutData)  ' 4500 dari depan
      Call ReadUnicodeFile2(hFile, MySize - 4500, 4500, OutData2)
      IsiVBS = StrConv(OutData, vbUnicode)
      IsiVBS = IsiVBS & StrConv(OutData2, vbUnicode) ' 4500 dari belakang
      Erase OutData()
      Erase OutData2()
   Else
      Call ReadUnicodeFile2(hFile, 1, MySize, OutData)
      IsiVBS = StrConv(OutData, vbUnicode)
      Erase OutData()
   End If
   
   IsiVBS = UCase(Replace(IsiVBS, Chr(0), ""))   ' [pembufferan hilangkan char 0]

      If InStr((IsiVBS), "REGWRITE") > 0 And InStr((IsiVBS), "WSCRIPT") > 0 Then XnamScrT = "Malware Script": GoSub BENAR
      If InStr((IsiVBS), "WSCRIPT.SCRIPTFULLNAME") > 0 Then XnamScrT = "Malware Script": GoSub BENAR
      If InStr((IsiVBS), "FORMAT C:") > 0 Then XnamScrT = "Gen-ScriptVirut": GoSub BENAR
      If InStr((IsiVBS), "FORMAT D:\") > 0 Then XnamScrT = "Gen-ScriptVirut": GoSub BENAR
      If InStr((IsiVBS), "AUTOEXEC.BAT") > 0 Then XnamScrT = "Gen-ScriptVirut": GoSub BENAR
      If InStr((IsiVBS), "VERSION\RUN") > 0 And InStr((IsiVBS), ".BAT") > 0 Then XnamScrT = "Malware Script": GoSub BENAR
      If InStr((IsiVBS), "VERSION\RUN") > 0 And InStr((IsiVBS), ".VBS") > 0 Then XnamScrT = "Malware Script": GoSub BENAR
      If InStr((IsiVBS), "VERSION\RUN") > 0 And InStr((IsiVBS), ".CMD") > 0 Then XnamScrT = "Malware Script": GoSub BENAR
      If InStr((IsiVBS), "VERSION\RUN") > 0 And InStr((IsiVBS), ".PIF") > 0 Then XnamScrT = "Malware Script": GoSub BENAR
      If MalBAT(sFile) = True Then XnamScrT = "Gen-ScriptVirut": GoSub BENAR
      If IsArrs(sFile) = True Then XnamScrT = "Gen-ScriptVirut": GoSub BENAR
      If ThumbCek(sFile) = True Then XnamScrT = "Gen-ScriptVirut": GoSub BENAR
      
If UCase(Right(sFile, 4)) = ".VBS" Then
   Pos_Akhir = Len(IsiVBS)
   For iCount = 1 To Pos_Akhir
       AscKar = Asc(Mid(IsiVBS, iCount, 1))
       If AscKar >= 32 And AscKar <= 57 Then
          JumNumer = JumNumer + 1
       Else
          JumKar = JumKar + 1
       End If
   DoEvents
   Next
   If JumNumer > JumKar Then XnamScrT = "Gen-ScriptVirut[VBS]": GoSub BENAR
Else
   MalScrCheck = False
End If

Exit Function

BENAR:
    MalScrCheck = True
    TutupFile hFile
End Function

Private Function MalBAT(sFileS As String) As Boolean
If IsPE32EXE = True Then Exit Function
If isProperFile(sFileS, "BAT CMD") = False Then Exit Function
If InStr((IsiVBS), "KILL EXPLORER.EXE") > 0 Then GoSub BENAR
If InStr((IsiVBS), "DEL EXPLORER.EXE") > 0 Then GoSub BENAR
If InStr((IsiVBS), "ASSOC .EXE=") > 0 Then GoSub BENAR
If InStr((IsiVBS), "TASKKILL") > 0 Then GoSub BENAR
Exit Function
BENAR:
MalBAT = True
End Function

Private Function ThumbCek(sFileS As String) As Boolean
If InStr(UCase(Right(sFileS, Len(sFileS) - InStrRev(sFileS, "\"))), "THUMB") > 0 Then
If isProperFile(sFileS, ".DB") = False Then Exit Function

If InStr((IsiVBS), "WSCRIPT") > 0 Then GoSub BENAR
If InStr((IsiVBS), "AUTORUN") > 0 Then GoSub BENAR
End If
Exit Function
BENAR:
ThumbCek = True
End Function

Private Function IsArrs(PathFile As String) As Boolean
Dim nmFile As String
Dim strDrv As String
On Error GoTo KELUAR

   nmFile = Mid(PathFile, 4)
   strDrv = Left(PathFile, 3)
    If Dir(strDrv & "Autorun.Inf", vbNormal Or vbHidden Or vbSystem) <> "" Then
        If InStr(UCase(IsiVBS), UCase(nmFile)) > 0 Then
            If isProperFile(PathFile, "ICO") = True Then
                IsArrs = False
                Exit Function
            Else
                IsArrs = True
                Exit Function
            End If
        Else
            IsArrs = False
        End If
    Else
        Exit Function
    End If
        
Exit Function
KELUAR:
End Function

Public Function ScanLNK(sFile As String) As Boolean
Static Xtuju As String
Static xDatum As String
    Dim LinkShell As New wshShell
    Dim LinkShortCut As WshShortcut
    Set LinkShortCut = LinkShell.CreateShortcut(sFile)

If UCase(Right(sFile, 3)) = "LNK" Then
Xtuju = LinkShortCut.TargetPath
Xtuju = UCase(Xtuju)
xDatum = ReadUnicodeFile(sFile)
If Len(Xtuju) = 0 Then
    If InStr(xDatum, "ê:i") > 0 Then
        ScanLNK = True
        Exit Function
    End If
End If

If ValidFile3(Xtuju) = False Then
Xtuju = UCase$(Xtuju)
If Right(Xtuju, 4) = ".SCR" Or Right(Xtuju, 4) = ".VBS" Or Right(Xtuju, 3) = ".DB" Then
    ScanLNK = True
    Exit Function
End If
End If

If UCase(Left(LinkShortCut.Arguments, 12)) = "//E:VBSCRIPT" Then
    ScanLNK = True
    Exit Function
End If

End If
ScanLNK = False
End Function

Function Tujuan(strPath As String) As String
On Error GoTo rusak
Dim wshShell As Object
Dim wshLink As Object
Set wshShell = CreateObject("WScript.Shell")
Set wshLink = wshShell.CreateShortcut(strPath)
Tujuan = wshLink.TargetPath
Set wshLink = Nothing
Set wshShell = Nothing
Exit Function
rusak:
End Function

Private Function EICARVirChk(xFile As String) As Boolean
If isProperFile(xFile, "COM") = False Then Exit Function
If InStr((IsiVBS), "EICAR") > 0 Then GoSub BENAR

Exit Function
BENAR:
EICARVirChk = True
End Function

Public Function GetVBSonHTML(sPath As String, xhFile As Long, xSizeFile As Long) As Boolean
Static sTemp As String
Static sTmp() As String
Static sTmp2() As String
Static pisah As String
Static iCount As Integer
Static iTemp As Integer
If isProperFile2(sPath, ".HTM HTML") = False Then Exit Function
sTemp = ReadUnicodeFile(sPath)
sTemp = UCase(sTemp)
If InStr(sTemp, "<SCRIPT LANGUAGE=VBSCRIPT><!--") > 0 Then
    sTmp2() = Split(sTemp, "<SCRIPT LANGUAGE=VBSCRIPT><!--", , vbTextCompare)
    xDataVBS4 = sTmp2(1)
    sTmp() = Split(sTemp, "//--></SCRIPT>", , vbTextCompare)
    xDataVBS2 = sTmp(1)
    xDataVBS3 = Len(xDataVBS2) + 14
    xDataVBS = Left(xDataVBS4, Len(xDataVBS4) - CInt(xDataVBS3))

WriteFileUniSim App.Path & "\tmp\tmp.vbs", xDataVBS
If MalScrCheck(App.Path & "\tmp\tmp.vbs", xhFile, xSizeFile) = True Then
GetVBSonHTML = True
HapusFile App.Path & "\tmp\tmp.vbs"
Exit Function
Else
GetVBSonHTML = False
End If
HapusFile App.Path & "\tmp\tmp.vbs"

ElseIf InStr(sTemp, "<SCRIPT LANGUAGE=" & Chr(34) & "VBSCRIPT" & Chr(34) & ">") > 0 Then

    sTmp2() = Split(sTemp, "<SCRIPT LANGUAGE=" & Chr(34) & "VBSCRIPT" & Chr(34) & ">", , vbTextCompare)
    xDataVBS4 = sTmp2(1)
    sTmp() = Split(sTemp, "</SCRIPT>", , vbTextCompare)
    xDataVBS2 = sTmp(1)
    xDataVBS3 = Len(xDataVBS2) + 14
    xDataVBS = Left(xDataVBS4, Len(xDataVBS4) - CInt(xDataVBS3))

WriteFileUniSim App.Path & "\tmp\tmp.vbs", xDataVBS

If MalScrCheck(App.Path & "\tmp\tmp.vbs", xhFile, xSizeFile) = True Then
GetVBSonHTML = True
HapusFile App.Path & "\tmp\tmp.vbs"
Exit Function
Else
GetVBSonHTML = False
End If
HapusFile App.Path & "\tmp\tmp.vbs"
End If
If InStr(sTemp, "OPEN(" & Chr(34) & "README.EML") > 0 Then
    GetVBSonHTML = True
Else
    GetVBSonHTML = False
End If
End Function

Public Function MalAutorun(sPath As String, xSizeFile As Long) As Boolean
Static sTemp23 As String
Static sTmp() As String
Static sTmp2() As String
Static pisah As String
Static iCount As Integer
Static iTemp As Integer
Static xNamaFile As String
Static xPathCadangan As String

If InStr(LCase$(GetPath(sPath, FileName)), "autorun") <= 0 Then Exit Function
If GetFileAttributes(StrPtr(sPath)) = 4 Then
MalAutorun = True
Exit Function
End If
If xSizeFile > 61440 Or xSizeFile < 10 Then Exit Function
If IsPE32EXE = True Then Exit Function
sTemp23 = LCase(BacaFileText(sPath))
If InStr(sTemp23, "[autorun]") > 0 And InStr(sTemp23, "shell") > 0 Then
    MalAutorun = True
    Exit Function
End If
If InStr(sTem23p, "[autorun]") > 0 And InStr(sTemp23, ".pif") > 0 Then
    MalAutorun = True
    Exit Function
End If
If InStr(sTemp23, "[autorun]") > 0 And InStr(sTemp23, ".bat") > 0 Then
    MalAutorun = True
    Exit Function
End If
If InStr(sTemp23, "[autorun]") > 0 And InStr(sTemp23, ".cmd") > 0 Then
    MalAutorun = True
    Exit Function
End If
If InStr(sTemp23, "[autorun]") > 0 And InStr(sTemp23, "open\command") > 0 Then
    MalAutorun = True
    Exit Function
End If
If InStr(sTemp23, "[autorun]") > 0 And InStr(sTemp23, "explore\command") > 0 Then
    MalAutorun = True
    Exit Function
End If
If InStr(sTemp23, "wscript.exe") > 0 Then
    MalAutorun = True
    Exit Function
End If
'If InStr(sTemp, "autorun") > 0 Then

'MalAutorun = True
'Exit Function
'pisah = Chr(13)
'stmp() = Split(sTemp, pisah)
'    iTemp = UBound(stmp())
'    For iCount = 1 To iTemp
'    stmp(iCount) = LCase$(stmp(iCount))
'    If InStr(stmp(iCount), "shellexecute") > 0 Or InStr(stmp(iCount), "shell\open\command") > 0 Or InStr(stmp(iCount), "shell\explore\command") > 0 Or InStr(stmp(iCount), ";") > 0 Then
'    MalAutorun = True
'    Exit Function
'        sTmp2() = Split(stmp(iCount), "=")
'        xNamaFile = sTmp2(1)
'    xPathCadangan = Left(sPath, Len(sPath) - Len(Right(sPath, Len(sPath) - InStrRev(sPath, "\"))))
'    xPathCadangan = xPathCadangan & xNamaFile
'    If ValidFile(xPathCadangan) = True Then
'        AddToLVMal frMain.lvMal, "Suspect Arrs.mod", xPathCadangan, CStr(xSizeFile), "Valid\ Not Deleted", 0, 18
'        Exit For
'    End If
'    End If
'    Next
'End If
MalAutorun = False
End Function
