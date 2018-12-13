Attribute VB_Name = "basScanSys"
Dim PID_ToTerminated(70)    As Long
Dim PID_ToRestarted(70)     As Long
Dim nTerminate              As Long
Dim nRestart                As Long
Dim nKunci                  As Long
Dim MYID                    As Long
Dim Path_Terminate(70)      As String
Dim Path_Restart(70)        As String
Dim Path_ToKunci(70)        As String

Dim CLFL As New classFile

Public Sub ScanProses(bModuleScan As Boolean, lbProses As Label)
On Error Resume Next
Dim LEAX            As Long
Dim CTurn           As Long
Dim PID             As Long
Dim nSize           As Long
Dim lngItem         As Long
Dim FTotal          As Long
Dim FTotal2         As Long
Dim ENPC()          As ENUMERATE_PROCESSES_OUTPUT
Dim ENMC()          As ENUMERATE_MODULES_OUTPUT
Dim sProsesPath     As String
Dim WScript         As String ' jika proses WS script dibunuh dulu
    ProcessScan = False
    LEAX = PamzEnumerateProcesses(ENPC())
    MYID = GetCurrentProcessId()
    VirStatus = False 'init
    If LEAX <= 0 Then
        GoTo LBL_TERAKHIR
    End If
    nTerminate = 0 ' init
    WScript = GetSpecFolder(WINDOWS_DIR) & "\System32\wscript.exe"
    FTotal = LEAX
    For CTurn = 0 To (LEAX - 1)
        PID = ENPC(CTurn).nProcessID
        FTotal2 = PamzEnumerateModules(PID, ENMC)
        FTotal = FTotal + FTotal2
    Next
    frMain.pScan.Max = FTotal
    FileRemain = FTotal
    frMain.pScan.value = 0
    frMain.pScan.Text = "0%"
    With frMain.lvMal
    For CTurn = 0 To (LEAX - 1)
        sProsesPath = PamzNtPathToUserFriendlyPathW(ENPC(CTurn).szNtExecutablePathW)
        PID = ENPC(CTurn).nProcessID
        xScanPath = "(" & PID & ") - " & sProsesPath
        nSize = ENPC(CTurn).nSizeOfExecutableOpInMemory
        If UCase(WScript) = UCase(sProsesPath) Then ' buat jaga-jaga kalo ada WScript
            PamzTerminateProcess (PID)
            GoTo LANJUT_FOR
        End If
        FileScan = FileScan + 1
        FileToScan = FileToScan + 1
        FileRemain = FileRemain - 1
        If ValidFile(sProsesPath) = True Then  ' yakinkan yang discan adalah file
            EqualProcess sProsesPath  ' cek proses apakah virus atau bukan dan ditambah heuristic (jika diset)
        End If
        If StopScan = True Then Exit For
        If VirStatus = True And PID <> MYID Then ' jika status virus true namun dengan catatan bukan proses sendiri
           lngItem = .ListItems.count
           PID_ToTerminated(nTerminate) = PID
           Path_Terminate(nTerminate) = sProsesPath
           PamzSuspendResumeProcessThreads PID, False ' di pause dulu
           nTerminate = nTerminate + 1
           ' ganti status
           .ListItems.Item(lngItem).SubItem(4).Text = "In Memory [Killed+Locked]" ' status diganti
            ProcessScan = True
        Else ' modulenya di scan klo bukan proses virus [tapi bModuleScan harus true]
            If bModuleScan = True Then ScanModules PID, lbProses, sProsesPath
        End If
        
LANJUT_FOR:
    Next
    End With
    
    Erase ENPC()
    
    CTurn = 0 'reset
        
    ' saat-nya beraksi secara serempak
    
    For CTurn = 1 To nRestart  ' restart proses2 yang terinfeksi module virus
        KillProses PID_ToRestarted(CTurn - 1), Path_Restart(CTurn - 1), True, False
    Next
    
    CTurn = 0 'reset

    For CTurn = 1 To nTerminate ' terminate lalu kunci proses virus
        KillProses PID_ToTerminated(CTurn - 1), Path_Terminate(CTurn - 1), False, True
    Next
    
    CTurn = 0 'reset
    For CTurn = 1 To nKunci ' kusus untuk module-module yang belum ke ke-kunci [kusus proses udah dikunci di atas]
        KunciFile Path_ToKunci(CTurn - 1) ' gak berhasil pake cdangan dulu smntara
    Next
    
LBL_TERAKHIR:
End Sub

Private Function ScanModules(ByVal TargetPID As Long, lbModule As Label, sProses As String) As Boolean ' akan TRUE jika salah satu module adalah virus, lalu proses akan dimatikan atau restart [karena semntara belum punya fungsi untuk unload dll]
On Error Resume Next
Dim LEAX            As Long
Dim CTurn           As Long
Dim pAddress        As String
Dim sModulePath     As String
Dim ENMC()          As ENUMERATE_MODULES_OUTPUT
    LEAX = PamzEnumerateModules(TargetPID, ENMC)
    If LEAX <= 0 Then ' gagal mendapatkan module
        GoTo LBL_TERAKHIR
    End If
    With frMain.lvMal
    For CTurn = 0 To (LEAX - 1)
        pAddress = Hex$(CLng(PamzNtPathToUserFriendlyPathW(CStr(ENMC(CTurn).pBaseAddress))))
        sModulePath = PamzNtPathToUserFriendlyPathW(ENMC(CTurn).szNtModulePathW)
        FileScan = FileScan + 1
        FileToScan = FileToScan + 1
        FileRemain = FileRemain - 1
        If ValidFile(sModulePath) = True Then ' yakinkan yang discan adalah file
            EqualProcess sModulePath ' cek module apakah virus atau bukan dan ditambah heuristic (jika diset)
        End If
        If StopScan = True Then Exit For
        If VirStatus = True And TargetPID <> MYID Then ' jika status virus true (module) & bukan proses sendiri --> di ganti kalo udah ada fungsi unload dll
           lngItem = .ListItems.count
           Path_Restart(nRestart) = sProses ' masukan alamat file proses-nya
           PID_ToRestarted(nRestart) = TargetPID
           Path_ToKunci(nKunci) = sModulePath ' tambahkan path module untuk dikunci
           PamzSuspendResumeProcessThreads TargetPID, False ' di pause dulu prosesnya
           nRestart = nRestart + 1 ' jumlah yang akan distart dinaikan
           nKunci = nKunci + 1 ' Jumlah yang mau dikunci (kunci module-nya biar gak bisa dijalankan lagi)
           .ListItems.Item(lngItem).SubItem(4).Text = "In Memory [Killed+Locked]" ' status diganti
           ProcessScan = True
        End If
    Next
    End With
    Erase ENMC()

LBL_TERAKHIR:
End Function

