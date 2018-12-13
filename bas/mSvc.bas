Attribute VB_Name = "mSvc"
Option Explicit
'Asli created by oj4nBL4NK
Public AppPath As String

Private Type QUERY_SERVICE_CONFIG
    dwServiceType As Long
    dwStartType As Long
    dwErrorControl As Long
    lpBinaryPathName As Long
    lpLoadOrderGroup As Long
    dwTagId As Long
    lpDependencies As Long
    lpServiceStartName As Long
    lpDisplayName As Long
End Type

Private Declare Function OpenSCManager Lib "advapi32" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function OpenService Lib "advapi32" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function QueryServiceConfig Lib "advapi32" Alias "QueryServiceConfigA" (ByVal hService As Long, lpServiceConfig As QUERY_SERVICE_CONFIG, ByVal cbBufSize As Long, pcbBytesNeeded As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function CloseServiceHandle Lib "advapi32" (ByVal hSCObject As Long) As Long
Private Declare Function QueryServiceStatus Lib "advapi32" (ByVal hService As Long, lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function ControlService Lib "advapi32" (ByVal hService As Long, ByVal dwControl As SERVICE_CONTROL, lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function DeleteService Lib "advapi32" (ByVal hService As Long) As Long
Private Declare Function CreateService Lib "advapi32" Alias "CreateServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal lpDisplayName As String, ByVal dwDesiredAccess As Long, ByVal dwServiceType As Long, ByVal dwStartType As Long, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As String, ByVal lpLoadOrderGroup As String, ByVal lpdwTagId As String, ByVal lpDependencies As String, ByVal lp As String, ByVal lpPassword As String) As Long
Private Declare Function StartService Lib "advapi32" Alias "StartServiceA" (ByVal hService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long
        
Private Const SC_MANAGER_CONNECT = &H1&
Private Const SERVICE_QUERY_CONFIG = &H1&
Private Const ERROR_INSUFFICIENT_BUFFER = 122&
Private Const SERVICE_QUERY_STATUS = &H4&
Private Const SC_MANAGER_CREATE_SERVICE = &H2&
Private Const SERVICE_AUTO_START As Long = 2
Private Const SERVICE_ERROR_NORMAL As Long = 1

Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SERVICE_CHANGE_CONFIG = &H2&
Private Const SERVICE_ENUMERATE_DEPENDENTS = &H8&
Private Const SERVICE_START = &H10&
Private Const SERVICE_STOP = &H20&
Private Const SERVICE_PAUSE_CONTINUE = &H40&
Private Const SERVICE_INTERROGATE = &H80&
Private Const SERVICE_USER_DEFINED_CONTROL = &H100&
Private Const SERVICE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SERVICE_QUERY_CONFIG Or SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS Or SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_START Or SERVICE_STOP Or SERVICE_PAUSE_CONTINUE Or SERVICE_INTERROGATE Or SERVICE_USER_DEFINED_CONTROL)

Public Const Service_Name As String = "oj4nBL4NK_Guard"
Public Const Service_Display_Name As String = "oj4nBL4NK_Guard"
Public Const Service_File_Name As String = "guard.exe -svc"

Public Function GetServiceConfig() As Long
Dim hSCManager As Long, hService As Long
Dim r As Long, SCfg() As QUERY_SERVICE_CONFIG, r1 As Long, s As String

hSCManager = OpenSCManager(vbNullString, vbNullString, _
                       SC_MANAGER_CONNECT)
If hSCManager <> 0 Then
    hService = OpenService(hSCManager, Service_Name, SERVICE_QUERY_CONFIG)
    If hService <> 0 Then
        ReDim SCfg(1 To 1)
        If QueryServiceConfig(hService, SCfg(1), 36, r) = 0 Then
            If Err.LastDllError = ERROR_INSUFFICIENT_BUFFER Then
                r1 = r \ 36 + 1
                ReDim SCfg(1 To r1)
                If QueryServiceConfig(hService, SCfg(1), r1 * 36, r) <> 0 Then
                    s = Space$(255)
                    lstrcpy s, SCfg(1).lpServiceStartName
                    s = Left$(s, lstrlen(s))
                    'frmServiceControl.txtAccount = s
                Else
                    GetServiceConfig = Err.LastDllError
                End If
            Else
                GetServiceConfig = Err.LastDllError
            End If
        End If
        CloseServiceHandle hService
    Else
        GetServiceConfig = Err.LastDllError
    End If
    CloseServiceHandle hSCManager
Else
    GetServiceConfig = Err.LastDllError
End If
End Function

Public Function GetServiceStatus() As SERVICE_STATE
Dim hSCManager As Long, hService As Long, Status As SERVICE_STATUS
hSCManager = OpenSCManager(vbNullString, vbNullString, _
                       SC_MANAGER_CONNECT)
If hSCManager <> 0 Then
    hService = OpenService(hSCManager, Service_Name, SERVICE_QUERY_STATUS)
    If hService <> 0 Then
        If QueryServiceStatus(hService, Status) Then
            GetServiceStatus = Status.dwCurrentState
        End If
        CloseServiceHandle hService
    End If
    CloseServiceHandle hSCManager
End If
End Function

Public Function DeleteNTService() As Long
Dim hSCManager As Long
Dim hService As Long, Status As SERVICE_STATUS

hSCManager = OpenSCManager(vbNullString, vbNullString, _
                       SC_MANAGER_CONNECT)
If hSCManager <> 0 Then
    hService = OpenService(hSCManager, Service_Name, _
                       SERVICE_ALL_ACCESS)
    If hService <> 0 Then
' Stop service if it is running
        ControlService hService, SERVICE_CONTROL_STOP, Status
        If DeleteService(hService) = 0 Then
            DeleteNTService = Err.LastDllError
        End If
        CloseServiceHandle hService
    Else
        DeleteNTService = Err.LastDllError
    End If
    CloseServiceHandle hSCManager
Else
    DeleteNTService = Err.LastDllError
End If

End Function

Public Function SetNTService() As Long
Dim hSCManager As Long
Dim hService As Long, DomainName As String

hSCManager = OpenSCManager(vbNullString, vbNullString, _
                       SC_MANAGER_CREATE_SERVICE)
If hSCManager <> 0 Then
    hService = CreateService(hSCManager, Service_Name, _
                       Service_Display_Name, SERVICE_ALL_ACCESS, _
                       SERVICE_WIN32_OWN_PROCESS Or SERVICE_INTERACTIVE_PROCESS, _
                       SERVICE_AUTO_START, SERVICE_ERROR_NORMAL, _
                       AppPath & Service_File_Name, vbNullString, _
                       vbNullString, vbNullString, "LocalSystem", _
                       vbNullString)
    If hService <> 0 Then
        CloseServiceHandle hService
    Else
        SetNTService = Err.LastDllError
    End If
    CloseServiceHandle hSCManager
Else
    SetNTService = Err.LastDllError
End If
End Function

Public Function StopNTService() As Long
Dim hSCManager As Long, hService As Long, Status As SERVICE_STATUS
hSCManager = OpenSCManager(vbNullString, vbNullString, _
                       SC_MANAGER_CONNECT)
If hSCManager <> 0 Then
    hService = OpenService(hSCManager, Service_Name, SERVICE_STOP)
    If hService <> 0 Then
        If ControlService(hService, SERVICE_CONTROL_STOP, Status) = 0 Then
            StopNTService = Err.LastDllError
        End If
    CloseServiceHandle hService
    Else
        StopNTService = Err.LastDllError
    End If
CloseServiceHandle hSCManager
Else
    StopNTService = Err.LastDllError
End If
End Function

Public Function StartNTService() As Long
Dim hSCManager As Long, hService As Long
hSCManager = OpenSCManager(vbNullString, vbNullString, _
                       SC_MANAGER_CONNECT)
If hSCManager <> 0 Then
    hService = OpenService(hSCManager, Service_Name, SERVICE_START)
    If hService <> 0 Then
        If StartService(hService, 0, 0) = 0 Then
            StartNTService = Err.LastDllError
        End If
    CloseServiceHandle hService
    Else
        StartNTService = Err.LastDllError
    End If
CloseServiceHandle hSCManager
Else
    StartNTService = Err.LastDllError
End If
End Function




