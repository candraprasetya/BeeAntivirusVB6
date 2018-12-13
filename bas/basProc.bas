Attribute VB_Name = "basProc"

Option Explicit
'Fungsi API yang berhubungan Process
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32.dll" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32.dll" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hHandle As Long) As Long
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long

' API Module
Public Declare Function Module32First Lib "kernel32" (ByVal hSnapshot As Long, uProcess As MODULEENTRY32) As Long
Public Declare Function Module32Next Lib "kernel32" (ByVal hSnapshot As Long, uProcess As MODULEENTRY32) As Long

'Fungsi API yang berhubungan Thread
Public Declare Function Thread32First Lib "kernel32.dll" (ByVal hSnapshot As Long, ByRef lpte As THREADENTRY32) As Boolean
Public Declare Function Thread32Next Lib "kernel32.dll" (ByVal hSnapshot As Long, ByRef lpte As THREADENTRY32) As Boolean

Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long
Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function OpenThread Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwThreadId As Long) As Long
Declare Function TerminateThread Lib "kernel32.dll" (ByVal hThread As Long, ByVal dwExitCode As Long) As Boolean


Public Type PROCESSENTRY32
    dwSize                  As Long
    cntUsage                As Long
    th32ProcessID           As Long
    th32DefaultHeapID       As Long
    th32ModuleID            As Long
    cntThreads              As Long
    th32ParentProcessID     As Long
    pcPriClassBase          As Long
    dwFlags                 As Long
    szExeFile               As String * 260
End Type

Public Type THREADENTRY32
    dwSize As Long
    cntUsage As Long
    th32ThreadID As Long
    th32OwnerProcessID As Long
    tpBasePri As Long
    tpDeltaPri As Long
    dwFlags As Long
End Type

Public Type MODULEENTRY32
  dwSize As Long
  th32ModuleID As Long
  th32ProcessID As Long
  GlblcntUsage As Long
  ProccntUsage As Long
  modBaseAddr As Long
  modBaseSize As Long
  hModule As Long
  szModule As String * 256
  szExePath As String * 260
End Type

'Konstanta yang dibutuhkan untuk Process
Public Const PROCESS_QUERY_INFORMATION As Long = &H400
Public Const MAX_PATH As Long = 260

Public Const PROCESS_VM_READ = &H10

'Konstanta yang dibutuhkan untuk Thread
Public Const THREAD_SUSPEND_RESUME As Long = &H2
Public Const HIGH_PRIORITY_CLASS = &H80
Public Const IDLE_PRIORITY_CLASS = &H40
Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const REALTIME_PRIORITY_CLASS = &H100
Public Const PROCESS_SET_INFORMATION As Long = (&H200)
Public Const TH32CS_SNAPHEAPLIST = &H1
Public Const TH32CS_SNAPPROCESS = &H2
Public Const TH32CS_SNAPTHREAD = &H4
Public Const TH32CS_SNAPMODULE = &H8
Public Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Public Const TH32CS_INHERIT = &H80000000
Public Const MAX_PATH2 As Integer = 260

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Long, lpThreadAttributes As Any, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long

Public Function RemoteExitProcess(lProcessID As Long) As Boolean
    Dim lProcess As Long
    Dim lRemThread As Long
    Dim lExitProcess As Long
    
    On Error GoTo errHandle
    
    lProcess = OpenProcess((&HF0000 Or &H100000 Or &HFFF), False, lProcessID) 'PROCESS_ALL_ACCESS
        lExitProcess = GetProcAddress(GetModuleHandleA("kernel32"), "ExitProcess")
        lRemThread = CreateRemoteThread(lProcess, ByVal 0, 0, ByVal lExitProcess, 0, 0, 0)
    CloseHandle lProcess
    
    CloseHandle lRemThread
    RemoteExitProcess = True
    
    Exit Function
errHandle:
    RemoteExitProcess = False
End Function
Public Function Strip_Null(ByVal strOrig As String) As String
    Strip_Null = Left$(strOrig, InStr(strOrig, vbNullChar) - 1)
End Function

Public Function Dapatkan_Process(Lv As ucListView)

Dim pProcess  As PROCESSENTRY32
Dim sSnapShot As Long
Dim rReturn   As Integer
Dim nTmp      As Integer
Dim ProName   As String
Dim ProPID    As Long
Dim ProPath   As String
Dim ItemXS      As cListItem

On Error Resume Next
    sSnapShot = CreateToolhelp32Snapshot(15, 0)
    pProcess.dwSize = Len(pProcess) '
    Process32First sSnapShot, pProcess ' proses pertama = [System Process]
    Lv.ListItems.Clear
    Do
        ProName = Strip_Null(pProcess.szExeFile)
        ProPID = pProcess.th32ProcessID
        ProPath = PathByPID(ProPID)
        If ProName = "[System Process]" Or ProName = "System" Then GoTo NAA
        If ProPath = "SYSTEM" Then ProPath = Environ$("windir") & "\System32\" & ProName
        DrawIco ProPath, frMain.picBuff, ricnSmall
        Lv.ImageList.AddFromDc frMain.picBuff.hdc, 16, 16
        Set ItemXS = Lv.ListItems.Add(, ProName, , Lv.ImageList.IconCount - 1)
        ItemXS.SubItem(2).Text = ProPID
        ItemXS.SubItem(3).Text = ProPath
        ItemXS.SubItem(4).Text = Round(FileLen(ProPath) / 1024, 1)
        SuspenResumeThread ProPID, True
        ItemXS.SubItem(5).Text = "Running !"
NAA:
        rReturn = Process32Next(sSnapShot, pProcess)
        DoEvents
    Loop While rReturn <> 0
    Set ItemXS = Nothing
    CloseHandle sSnapShot
End Function

Public Sub Matikan_Process(ID_nya As Long)
On Error Resume Next
    TerminateProcess OpenProcess(2035711, 1, ID_nya), 0 'Ends the process selected
    DoEvents 'Let's the computer "do it's stuff"
End Sub
' Untuk mematikan berdasarkan Nama Process
Public Sub Matikan_ByPath(Lv As ucListView, spath As String)
Dim Num As Integer
On Error Resume Next
For Num = 1 To Lv.ListItems.count
    If UCase(Lv.ListItems(Num)) = UCase(spath) Then
        Matikan_Process Lv.ListItems(Num).SubItems(1)
        Exit For
    End If
Next
End Sub

Public Function PathByPID(PID As Long) As String
    Dim cbNeeded As Long
    Dim Modules(1 To 200) As Long
    Dim ret As Long
    Dim ModuleName As String
    Dim nSize As Long
    Dim hProcess As Long
    
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
        Or PROCESS_VM_READ, 0, PID)
    
    If hProcess <> 0 Then
        
        ret = EnumProcessModules(hProcess, Modules(1), _
            200, cbNeeded)
        
        If ret <> 0 Then
            ModuleName = Space(MAX_PATH)
            nSize = 500
            ret = GetModuleFileNameExA(hProcess, _
                Modules(1), ModuleName, nSize)
            PathByPID = Left(ModuleName, ret)
        End If
    End If
    
    ret = CloseHandle(hProcess)
    
    If PathByPID = "" Then
        PathByPID = "SYSTEM"
    End If
    
    If Left(PathByPID, 4) = "\??\" Then
        PathByPID = Mid(PathByPID, 5, Len(PathByPID))
        Exit Function
    End If
    
    
    If Left(PathByPID, 12) = "\SystemRoot\" Then
        PathByPID = WinDirX & "\" & Mid(PathByPID, 13, Len(PathByPID))
        Exit Function
    End If
End Function

Private Function WinDirX() As String
WinDirX = Environ$("windir")
End Function

Public Function System_Buffer(Path As String) As Boolean
On Error GoTo fux
If FileLen(Path) > 0 Then System_Buffer = True
Exit Function

fux:
System_Buffer = False
End Function

Public Sub ListModule(PID As Long, lstMod As ListBox)
Dim hSnap As Long
    hSnap = CreateToolhelp32Snapshot(8, PID)

Dim meModule As MODULEENTRY32
    meModule.dwSize = LenB(meModule)

Dim nModule As Long
    nModule = Module32First(hSnap, meModule)

    lstMod.Clear
    If nModule = 0 Then lstMod.AddItem "Caution : cannot emumerate modules from selected PID !"

    Do While nModule
        lstMod.AddItem meModule.szExePath
        nModule = Module32Next(hSnap, meModule)
    Loop

    CloseHandle hSnap

End Sub


'====================================================================
' Untuk Resume atau Pause Process
Public Function SuspenResumeThread(PID As Long, isResume As Boolean)
        Dim hThread As Long
        Dim lSuspendCount As Long
        Dim tEntry() As THREADENTRY32
        
        GetEnumThread32 tEntry, PID
        Dim i As Integer
        For i = 0 To UBound(tEntry)
            If tEntry(i).th32OwnerProcessID = PID Then
               If isResume Then
                  hThread = OpenThread(THREAD_SUSPEND_RESUME, False, tEntry(i).th32ThreadID)
                  lSuspendCount = ResumeThread(hThread)
               Else
                  hThread = OpenThread(THREAD_SUSPEND_RESUME, False, tEntry(i).th32ThreadID)
                  lSuspendCount = SuspendThread(hThread)
               End If
            End If
        Next i
End Function

Public Function GetEnumThread32(ByRef Thread() As THREADENTRY32, Optional ByVal lProcessID As Long) As Long
On Error GoTo VB_Error
    ReDim Thread(0)
    
    Dim THREADENTRY32 As THREADENTRY32
    Dim hSnapshot As Long
    Dim lThread As Long
    
    hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, lProcessID)  'If hSnapShot = INVALID_HANDLE_VALUE Then Call Err_Dll(Err.LastDllError, "CreateToolHelp32Snapshoot ::: INVALID_HANDLE_VALUE failed", sLocation, "Thread32_Enum")
    
    THREADENTRY32.dwSize = Len(THREADENTRY32)
    If Thread32First(hSnapshot, THREADENTRY32) = False Then
        GetEnumThread32 = -1
        Exit Function
    Else
        ReDim Thread(lThread)
        Thread(lThread) = THREADENTRY32
    End If
    
    Do
        If Thread32Next(hSnapshot, THREADENTRY32) = False Then
            Exit Do
        Else
            lThread = lThread + 1
            ReDim Preserve Thread(lThread)
            Thread(lThread) = THREADENTRY32
        End If
    Loop
    GetEnumThread32 = lThread
    
Exit Function
VB_Error:
Resume Next
End Function
