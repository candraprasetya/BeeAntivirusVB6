Attribute VB_Name = "basProcess"
' Untuk Enum proses dan akses proses lain dari module ProsesAkses

Dim stFileStart() As String ' untuk penampung file-file startup
Dim JumStart      As Long ' penampung jumlahnya

Dim cImgList  As gComCtl

Public Sub ENUM_PROSES(ByRef Lv As ucListView, pcBuffer As PictureBox)
On Error Resume Next
Dim LEAX            As Long
Dim CTurn           As Long
Dim ENPC()          As ENUMERATE_PROCESSES_OUTPUT
Dim sTmp(9)         As String
Dim BufPathUni      As String
Set cImgList = New gComCtl

    LEAX = PamzEnumerateProcesses(ENPC())
    Lv.ListItems.Clear '---hapus isi yg lama.
    If LEAX <= 0 Then
        GoTo LBL_TERAKHIR
    End If
    
    Set Lv.ImageList = cImgList.NewImageList(16, 16, imlColor32)
    
    EnumStatStart stFileStart()  ' di enum startupnya dulu
    
    For CTurn = 0 To (LEAX - 1)
        sTmp(0) = ENPC(CTurn).szNtExecutableNameW
        sTmp(1) = GetStatStart(PamzNtPathToUserFriendlyPathW(ENPC(CTurn).szNtExecutablePathW)) & " value(s)"
        sTmp(2) = ENPC(CTurn).nProcessID
        sTmp(3) = ENPC(CTurn).nParentProcessID
        sTmp(4) = ENPC(CTurn).bIsHiddenProcess
        sTmp(5) = ENPC(CTurn).bIsBeingDebugged
        sTmp(6) = ENPC(CTurn).bIsLockedProcess
        sTmp(7) = Format$(ENPC(CTurn).nSizeOfExecutableOpInMemory, "#,#")
        sTmp(8) = PamzNtPathToUserFriendlyPathW(ENPC(CTurn).szNtExecutablePathW)
        sTmp(9) = "Unknow"
        
        BufPathUni = sTmp(8)
        If ValidFile(sTmp(8)) = False Then
           MakeExeBuffer GetSpecFolder(USER_DOC) & "\00.exe"
           DrawIco GetSpecFolder(USER_DOC) & "\00.exe", frMain.picBuff, ricnSmall
           HapusFile GetSpecFolder(USER_DOC) & "\00.exe"
        Else
           DrawIco BufPathUni, frMain.picBuff, ricnSmall
        End If
        
        Lv.ImageList.AddFromDc frMain.picBuff.hdc, 16, 16
        
        Lv.ListItems.Add , sTmp(0), , (Lv.ImageList.IconCount - 1), , , , , Array(sTmp(1), sTmp(2), sTmp(3), sTmp(4), sTmp(5), sTmp(6), sTmp(7), sTmp(8), sTmp(9))

        If ENPC(CTurn).bIsHiddenProcess = True Or IsHiddenFilePros(sTmp(8)) = True Then
           Lv.ListItems.Item(CTurn + 1).Cut = True
        End If
    Next
    
    Erase ENPC()
LBL_TERAKHIR:
    Set cImgList = Nothing
End Sub

' Kill, Kunci, dan retstart
Public Function KillProses(PID As Long, spath As String, bRestart As Boolean, bKunci As Boolean) As Boolean
    If PamzTerminateProcess(PID) > 0 Then
       KillProses = True
       If bRestart = True Then ' mau direstart
          Shell spath, vbNormalNoFocus
       End If
       If bKunci = True Then
          KunciFile spath
       End If
    Else
       KillProses = False
    End If
End Function

Public Function SuspendProses(PID As Long, bPause As Boolean) As String
If bPause = True Then ' mau pause
   If PamzSuspendResumeProcessThreads(PID, False) > 0 Then
      SuspendProses = "Paused"
   Else
      SuspendProses = "Ps-Failed"
   End If
Else
   If PamzSuspendResumeProcessThreads(PID, True) > 0 Then
      SuspendProses = "Resumed"
   Else
      SuspendProses = "Rs-Failed"
   End If
End If
End Function

' ----------------------------- Fungsi-Funsgi Buffer

' Melakukan enumerisasi Startup pada titik-titik yang dimasukan saat enum RegStartUp
Private Sub EnumStatStart(ByRef strFile() As String) 'output pada strFile
Dim iNum     As Long
Dim iCount   As Long
Dim stFile() As String
    iNum = EnumRegStartup(stFile, True)
    ReDim strFile(iNum) As String ' indeknya ga kepakai satu gpp
    
    For iCount = 1 To iNum
        strFile(iCount - 1) = stFile(iCount - 1)
    Next
    JumStart = iNum
End Sub

' Untuk mencocokan berapa nilai startup suatu alamat file
Private Function GetStatStart(sFile As String) As Long
Dim iNum As Long
Dim nJum As Long
For iNum = 1 To JumStart
    If UCase(sFile) = UCase(stFileStart(iNum - 1)) Then
       nJum = nJum + 1
    End If
Next
GetStatStart = nJum
End Function

' Menggambar icon ke picture box untuk tampilan listview proses
Public Sub DrawIco(spath As String, oPic As PictureBox, nDimension As IconRetrieve)
    With oPic
        .Cls: .AutoRedraw = True
        RetrieveIcon spath, oPic, nDimension
    End With
End Sub

' Cadangan path file yang tak mampu di enum
Private Sub MakeExeBuffer(spath As String)
Dim sTmp(1) As Byte
WriteUnicodeFile spath, 1, sTmp
End Sub

Private Function IsHiddenFilePros(sFilePro As String) As Boolean
Dim NAT As Long

NAT = GetFileAttributes(StrPtr(sFilePro))

If (NAT = 2 Or NAT = 34 Or NAT = 3 Or NAT = 6 Or NAT = 22 Or NAT = 18 Or NAT = 50 Or NAT = 19 Or NAT = 35) Then
    IsHiddenFilePros = True
Else
    IsHiddenFilePros = False
End If

End Function

