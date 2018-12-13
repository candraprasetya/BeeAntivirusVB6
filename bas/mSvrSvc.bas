Attribute VB_Name = "mSvrSvc"
Option Explicit
'Asli created by oj4nBL4NK
Private Declare Function CreateThread Lib "kernel32" (ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private ServiceStatus As SERVICE_STATUS
Private hServiceStatus As Long
Function FncPtr(ByVal fnp As Long) As Long
    FncPtr = fnp
End Function
Public Function StartAsService() As Long
    Dim ThreadId As Long
    StartAsService = CreateThread(0&, 0&, AddressOf ServiceThread, 0&, 0&, ThreadId)
End Function
Private Sub ServiceThread(ByVal dummy As Long)
    Dim ServiceTableEntry As SERVICE_TABLE
    ServiceTableEntry.lpServiceName = ServiceNamePtr
    ServiceTableEntry.lpServiceProc = FncPtr(AddressOf ServiceMain)
    StartServiceCtrlDispatcher ServiceTableEntry
End Sub
Private Sub ServiceMain(ByVal dwArgc As Long, ByVal lpszArgv As Long)
    ServiceStatus.dwServiceType = SERVICE_WIN32_OWN_PROCESS Or SERVICE_INTERACTIVE_PROCESS
    ServiceStatus.dwControlsAccepted = SERVICE_ACCEPT_STOP Or SERVICE_ACCEPT_SHUTDOWN
    ServiceStatus.dwWin32ExitCode = 32&
    ServiceStatus.dwServiceSpecificExitCode = 0&
    ServiceStatus.dwCheckPoint = 0&
    ServiceStatus.dwWaitHint = 0&
    hServiceStatus = RegisterServiceCtrlHandler(Service_Name, AddressOf handler)
    SetServiceState SERVICE_START_PENDING
    SetEvent hStartEvent
    WaitForSingleObject hStopEvent, INFINITE
End Sub
Private Sub handler(ByVal fdwControl As Long)
    Select Case fdwControl
        Case SERVICE_CONTROL_SHUTDOWN, SERVICE_CONTROL_STOP
            SetServiceState SERVICE_STOP_PENDING
            SetEvent hStopPendingEvent
        Case Else
            SetServiceState
    End Select
End Sub
Public Sub SetServiceState(Optional ByVal NewState As SERVICE_STATE = 0&)
    If NewState <> 0& Then ServiceStatus.dwCurrentState = NewState
    SetServiceStatus hServiceStatus, ServiceStatus
End Sub
