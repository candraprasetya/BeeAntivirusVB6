Attribute VB_Name = "basSysTray"
' ########################################################
' Module untuk penanganan systray dan balon tips
'
'



' API untuk Baloon Tips dan Systray Icon
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnID As NOTIFYICONDATA) As Boolean


Public Type NOTIFYICONDATA
    cbSize              As Long
    hWnd                As Long
    Uid                 As Long
    uFlags              As Long
    uCallbackMessage    As Long
    hIcon               As Long
    szTip               As String * 128
    dwState             As Long
    dwStateMask         As Long
    szInfo              As String * 256
    uTimeout            As Long
    szInfoTitle         As String * 64
    dwInfoFlags         As Long
End Type


Public Enum TypeBallon
        NIIF_NONE = &H0
        NIIF_WARNING = &H2
        NIIF_ERROR = &H3
        NIIF_INFO = &H1
        NIIF_GUID = &H4
End Enum



'Konstanta Icon Tray
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201     'Button down kiri.
Public Const WM_LBUTTONUP = &H202       'Button up kiri.
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click.
Public Const WM_RBUTTONDOWN = &H204     'Button down kanan.
Public Const WM_RBUTTONUP = &H205       'Button up kanan.
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click.
Public Const NIF_INFO                    As Long = &H10

Public nID As NOTIFYICONDATA
Public tJudul As String, tPesan As String


Public Sub UpdateIcon(IconApa As Long, sTips As String, Frm As Form)
    With nID
        .cbSize = Len(nID)
        .hWnd = Frm.hWnd
        .Uid = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = IconApa
        .szTip = sTips & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nID
End Sub


' Balons tips
' Untuk menampilkan Baloon Tips
Public Sub TampilkanBalon(Frm As Form, Message As String, Title As String, Optional balType As TypeBallon = NIIF_INFO)

    Call CabutBalon(Frm)

        With nID
            .cbSize = Len(nID)
            .hWnd = Frm.hWnd
            .Uid = vbNull
            .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIM_MODIFY 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
            .uCallbackMessage = WM_MOUSEMOVE
            .hIcon = Frm.Icon
            .dwState = 0
            .dwStateMask = 0
            .szInfo = Message & Chr(0)
            .szInfoTitle = Title & Chr(0)
            .dwInfoFlags = balType
        End With
    
    Call Shell_NotifyIcon(NIM_MODIFY, nID)

End Sub

' Mengakhiri tampilan Balon Tips
Public Sub CabutBalon(Frm As Form)
' Kill any current Balloon tip on screen for referenced form
  
        With nID
            .cbSize = Len(nID)
            .hWnd = Frm.hWnd
            .Uid = vbNull
            .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIM_MODIFY
            .uCallbackMessage = WM_MOUSEMOVE
            .hIcon = Frm.Icon
            .dwState = 0
            .dwStateMask = 0
            .szInfo = Chr(0)
            .szInfoTitle = Chr(0)
            .dwInfoFlags = NIIF_NONE
        End With
    
    Call Shell_NotifyIcon(NIM_MODIFY, nID)

End Sub



