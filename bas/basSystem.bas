Attribute VB_Name = "basSystem"
Option Explicit
'Asli created by oj4nBL4NK
Public Const INFINITE = -1&
Private Const WAIT_TIMEOUT = 258&
Private Const msgSETFG = 4160

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(1 To 128) As Byte
End Type

Public Const VER_PLATFORM_WIN32_NT = 2&
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Public hStopEvent As Long, hStartEvent As Long, hStopPendingEvent
Public IsNTService As Boolean
Public ServiceName() As Byte, ServiceNamePtr As Long
Private Sub MainSvc()
    Dim hnd As Long
    Dim h(0 To 1) As Long

    hStopEvent = CreateEvent(0, 1, 0, vbNullString)
    hStopPendingEvent = CreateEvent(0, 1, 0, vbNullString)
    hStartEvent = CreateEvent(0, 1, 0, vbNullString)
    ServiceName = StrConv(Service_Name, vbFromUnicode)
    ServiceNamePtr = VarPtr(ServiceName(LBound(ServiceName)))

        hnd = StartAsService
        h(0) = hnd
        h(1) = hStartEvent
        IsNTService = WaitForMultipleObjects(2&, h(0), 0&, INFINITE) = 1&
        If Not IsNTService Then
            CloseHandle hnd
            MessageBox 0&, "This program must be started as a service.", App.Title, msgSETFG
        End If
    
    If IsNTService Then
        SetServiceState SERVICE_RUNNING
        Do: DoEvents
        Loop While WaitForSingleObject(hStopPendingEvent, 10&) = WAIT_TIMEOUT
        
        SetServiceState SERVICE_STOPPED
        SetEvent hStopEvent
        WaitForSingleObject hnd, INFINITE
        CloseHandle hnd
    End If
    CloseHandle hStopEvent
    CloseHandle hStartEvent
    CloseHandle hStopPendingEvent
End Sub
Public Function CheckIsNT() As Boolean
    Dim OSVer As OSVERSIONINFO
    OSVer.dwOSVersionInfoSize = LenB(OSVer)
    GetVersionEx OSVer
    CheckIsNT = OSVer.dwPlatformId = VER_PLATFORM_WIN32_NT
End Function

Private Sub CheckService()
Dim ServState As SERVICE_STATE
If Not GetServiceConfig() = 0 Then
    Call SetNTService: GoTo nXt
Else
nXt:
    ServState = GetServiceStatus()
    Select Case ServState
        Case SERVICE_STOPPED
        StartNTService
    End Select
End If
End Sub

Public Sub endEX()
ExitProcess GetExitCodeProcess(GetCurrentProcess, 0)
End Sub

Public Sub Set2beSvc()
If Not CheckIsNT() Then
    Exit Sub
End If
AppPath = App.Path
If Right$(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
 If Trim$(Command$) = "-svc" Then
  Call MainSvc
 Else
  CheckService
 End If
End Sub

Public Sub RemFromSvc()
Dim ServState As SERVICE_STATE
ServState = GetServiceStatus()
    If GetServiceConfig() = 0 Then
            Select Case ServState
            Case SERVICE_RUNNING
                StopNTService
        End Select
    DeleteNTService
    End If
End Sub
