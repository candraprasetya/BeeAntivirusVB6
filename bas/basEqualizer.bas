Attribute VB_Name = "basEqualizer"
Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long

Public Function Equal(xPath As String)
Static TMPGlobal As Long
Static xSize As String
Static iCount As Integer
Static xChkSum As String
Static XPEsum As String
Static xPathPendek As String
Dim RetPE        As Long
Dim RetVirus     As String
On Error GoTo l_End
VirusStatus = False
If isProperFile(xPath, "TAR ZIP GZIP GZ JAR") = True Then
If frMain.ck(8).Value = 1 Then
    GetFileInArc xPath
End If
End If
TMPGlobal = GetHandleFile(xPath)
If TMPGlobal <= 0 Then GoTo l_End

xSize = GetSizeFile(TMPGlobal)
If frMain.ck(2).Value = 1 Then
    If xSize > 104857600 Then GoTo l_End        ' 100MB
End If

RetPE = IsValidPE32(TMPGlobal) ' fungsi balik IsValidPE32 adalah AddresOfNewHeader
If RetPE >= 64 Then
    RetVirus = GetDataEP(TMPGlobal, 40, RetPE)
    If RetVirus <> "" Then
          If Left(RetVirus, 3) = "PW:" Then
             AddToLVMal frMain.lvMal, Mid(RetVirus, 4), xPath, xSize, "Valid\ Not Deleted", 0, 18
             VirusStatus = True
          Else
             AddToLVMal frMain.lvMal, RetVirus, xPath, xSize, "Valid\ Not Deleted", 0, 18
             VirusStatus = True
          End If
          GoTo l_End
    End If
End If
RetVirus = ""
If frMain.ck(3).Value = 1 Then
    If MalScrCheck(xPath, TMPGlobal, CLng(xSize)) = True Then
        AddToLVMal frMain.lvMal, XnamScrT, xPath, xSize, "Valid\ Not Deleted", 0, 18
        VirusStatus = True
        GoTo l_End
    End If
    If CheckIcon(xPath, TMPGlobal) = True Then
        AddToLVMal frMain.lvMal, "Icon Suspect", xPath, xSize, "Valid\ Not Deleted", 0, 18
        VirusStatus = True
        GoTo l_End
    End If
    If GetVBSonHTML(xPath, TMPGlobal, CLng(xSize)) = True Then
        AddToLVMal frMain.lvMal, "Infected HTML", xPath, xSize, "Valid\ Not Deleted", 0, 18
        VirusStatus = True
        GoTo l_End
    End If
    xPathPendek = GetShortFileName(xPath)
    If MalAutorun(xPathPendek, CLng(xSize)) = True Then
        AddToLVMal frMain.lvMal, "Malware Autorun", xPath, xSize, "Valid\ Not Deleted", 0, 18
        VirusStatus = True
        GoTo l_End
    End If
    RetVirus = CekMalwareGen(xPath, TMPGlobal, CLng(xSize))
    If RetVirus <> "" Then
        AddToLVMal frMain.lvMal, RetVirus, xPath, xSize, "Valid\ Not Deleted", 0, 18
        GoTo l_End
    End If
    RetVirus = ""
    RetVirus = ChkNamVir(xPath, CLng(xSize))
    If RetVirus <> "" Then
        AddToLVMal frMain.lvMal, RetVirus, xPath, xSize, "Valid\ Not Deleted", 0, 18
        VirusStatus = True
        GoTo l_End
    End If
    If UCase$(Right$(xPath, 4)) = ".LNK" Then
    If ScanLNK(xPath) = True Then
        AddToLVMal frMain.lvMal, "Malware Shorcut", xPath, xSize, "Valid\ Not Deleted", 0, 18
        VirusStatus = True
        GoTo l_End
    End If
    End If
End If
xChkSum = MYCeksum(xPath, TMPGlobal)

   If xChkSum = String(Len(xChkSum), "0") Then
      xChkSum = MYCeksumCadangan(xPath, TMPGlobal)
   End If

iCount = 0
For iCount = 1 To xJumChecksum
    If xNumChecksum(iCount) = xChkSum Then
        AddToLVMal frMain.lvMal, xNamChecksum(iCount), xPath, xSize, "Valid\ Not Deleted", 0, 18
        VirusStatus = True
        Exit For
    End If
Next
l_End:
    TutupFile TMPGlobal
End Function

Public Function EqualRTP(xPath As String)
Static TMPGlobal As Long
Static xSize As String
Static iCount As Integer
Static xChkSum As String
Static XPEsum As String
Dim RetPE        As Long
Dim RetVirus     As String
Dim xPathPendek As String
On Error GoTo l_End

TMPGlobal = GetHandleFile(xPath)
If TMPGlobal <= 0 Then GoTo l_End

xSize = GetSizeFile(TMPGlobal)
If frMain.ck(2).Value = 1 Then
    If xSize > 104857600 Then GoTo l_End        ' 100MB
End If

RetPE = IsValidPE32(TMPGlobal) ' fungsi balik IsValidPE32 adalah AddresOfNewHeader

If RetPE >= 64 Then
    RetVirus = GetDataEP(TMPGlobal, 40, RetPE)
    If RetVirus <> "" Then
          If Left(RetVirus, 3) = "PW:" Then
             AddToLVMal frRTP.lvRTP, Mid(RetVirus, 4), xPath, xSize, "Valid\ Not Deleted", 0, 18
          Else
             AddToLVMal frRTP.lvRTP, RetVirus, xPath, xSize, "Valid\ Not Deleted", 0, 18
          End If
       GoTo l_End
    End If
End If

If frMain.ck(3).Value = 1 Then
    If MalScrCheck(xPath, TMPGlobal, CLng(xSize)) = True Then
        AddToLVMal frRTP.lvRTP, XnamScrT, xPath, xSize, "Valid\ Not Deleted", 0, 18
        GoTo l_End
    End If
    If CheckIcon(xPath, TMPGlobal) = True Then
        AddToLVMal frRTP.lvRTP, "Icon Suspect", xPath, xSize, "Valid\ Not Deleted", 0, 18
        GoTo l_End
    End If
    If GetVBSonHTML(xPath, TMPGlobal, CLng(xSize)) = True Then
        AddToLVMal frRTP.lvRTP, "Infected HTML", xPath, xSize, "Valid\ Not Deleted", 0, 18
        GoTo l_End
    xPathPendek = GetShortFileName(xPath)
    If MalAutorun(xPathPendek, CLng(xSize)) = True Then
        AddToLVMal frRTP.lvRTP, "Malware Autorun", xPath, xSize, "Valid\ Not Deleted", 0, 18
        GoTo l_End
    End If
    RetVirus = CekMalwareGen(xPath, TMPGlobal, CLng(xSize))
    If RetVirus <> "" Then
        AddToLVMal frRTP.lvRTP, RetVirus, xPath, xSize, "Valid\ Not Deleted", 0, 18
        GoTo l_End
    End If
    End If
End If
xChkSum = MYCeksum(xPath, TMPGlobal)

   If xChkSum = String(Len(xChkSum), "0") Then
      xChkSum = MYCeksumCadangan(xPath, TMPGlobal)
   End If
   
iCount = 0
For iCount = 1 To xJumChecksum
    If xNumChecksum(iCount) = xChkSum Then
        AddToLVMal frRTP.lvRTP, xNamChecksum(iCount), xPath, xSize, "Valid\ Not Deleted", 0, 18
        Exit For
    End If
Next
l_End:
    TutupFile TMPGlobal
End Function

Public Function EqualProcess(xPath As String)
Static TMPGlobal As Long
Static xSize As String
Static iCount As Integer
Static xChkSum As String
Static XPEsum As String
Dim RetPE        As Long
Dim RetVirus     As String
On Error GoTo l_End
VirusStatus = False
TMPGlobal = GetHandleFile(xPath)
If TMPGlobal <= 0 Then GoTo l_End

xSize = GetSizeFile(TMPGlobal)
If frMain.ck(2).Value = 1 Then
    If xSize > 104857600 Then GoTo l_End        ' 100MB
End If

RetPE = IsValidPE32(TMPGlobal) ' fungsi balik IsValidPE32 adalah AddresOfNewHeader
If RetPE >= 64 Then
    RetVirus = GetDataEP(TMPGlobal, 40, RetPE)
    If RetVirus <> "" Then
          If Left(RetVirus, 3) = "PW:" Then
             AddToLVMal frMain.lvMal, Mid(RetVirus, 4), LCase$(xPath), xSize, "Valid\ Not Deleted", 0, 18
             VirusStatus = True
          Else
             AddToLVMal frMain.lvMal, RetVirus, LCase$(xPath), xSize, "Valid\ Not Deleted", 0, 18
             VirusStatus = True
          End If
          GoTo l_End
    End If
End If
RetVirus = ""
If frMain.ck(3).Value = 1 Then
    If CheckIcon(xPath, TMPGlobal) = True Then
        AddToLVMal frMain.lvMal, "Icon Suspect", LCase$(xPath), xSize, "Valid\ Not Deleted", 0, 18
        VirusStatus = True
        GoTo l_End
    End If
    RetVirus = CekMalwareGen(xPath, TMPGlobal, CLng(xSize))
    If RetVirus <> "" Then
        AddToLVMal frMain.lvMal, RetVirus, LCase$(xPath), xSize, "Valid\ Not Deleted", 0, 18
        GoTo l_End
    End If
End If
xChkSum = MYCeksum(xPath, TMPGlobal)

   If xChkSum = String(Len(xChkSum), "0") Then
      xChkSum = MYCeksumCadangan(xPath, TMPGlobal)
   End If

iCount = 0
For iCount = 1 To xJumChecksum
    If xNumChecksum(iCount) = xChkSum Then
        AddToLVMal frMain.lvMal, xNamChecksum(iCount), LCase$(xPath), xSize, "Valid\ Not Deleted", 0, 18
        VirusStatus = True
        Exit For
    End If
Next
l_End:
    TutupFile TMPGlobal
End Function
