Attribute VB_Name = "modHooking"
Option Explicit

Public hHook As Long

Private Const NC_XBUTTON1_MASK = &H10000
Private Const NC_XBUTTON2_MASK = &H20000

' Virtual Key constants.
Public Const VK_CONTROL = &H11
Public Const VK_SHIFT = &H10
Public Const VK_MENU = &H12
Public Const VK_XBUTTON1 = &H5&          '/* NOT contiguous with L & RBUTTON */
Public Const VK_XBUTTON2 = &H6&          '/* NOT contiguous with L & RBUTTON */


' Masks for detecting keyboard and mouse states from wParam.
'
' eg. IF (wParam and MK_LBUTTON) THEN LeftButtonPressed.
'
Public Const MK_LBUTTON = &H1
Public Const MK_RBUTTON = &H2
Public Const MK_SHIFT = &H4
Public Const MK_CONTROL = &H8
Public Const MK_MBUTTON = &H10
Public Const MK_XBUTTON1 = &H20
Public Const MK_XBUTTON2 = &H40

' Windows Constants used by Subclasser.
Public Const WH_KEYBOARD            As Long = &H2
Public Const WM_MOVE                As Long = &H3
Public Const WM_ACTIVATE            As Long = &H6
Public Const WH_MOUSE_LL            As Long = &HE
Public Const WM_CLOSE               As Long = &H10
Public Const WM_SETCURSOR           As Long = &H20
Public Const WM_MOUSEACTIVATE       As Long = &H21
Public Const WM_WINDOWPOSCHANGED    As Long = &H47
Public Const WM_NCLBUTTONDOWN       As Long = &HA1
Public Const WM_NCLBUTTONUP         As Long = &HA2
Public Const WM_NCRBUTTONDOWN       As Long = &HA4
Public Const WM_NCRBUTTONUP         As Long = &HA5
Public Const WM_NCMBUTTONDOWN       As Long = &HA7
Public Const WM_NCMBUTTONUP         As Long = &HA8
Public Const WM_NCXBUTTONDOWN       As Long = &HAB
Public Const WM_NCXBUTTONUP         As Long = &HAC
Public Const WM_NCXBUTTONDBLCLK     As Long = &HAD
Public Const WM_NCPAINT             As Long = &H85
Public Const WM_SYSCOMMAND          As Long = &H112
Public Const WM_MENUSELECT          As Long = &H11F
Public Const WM_MENUCOMMAND         As Long = &H126
Public Const WM_MOUSEMOVE           As Long = &H200
Public Const WM_LBUTTONDOWN         As Long = &H201
Public Const WM_LBUTTONUP           As Long = &H202
Public Const WM_LBUTTONDBLCLK       As Long = &H203
Public Const WM_RBUTTONDOWN         As Long = &H204
Public Const WM_RBUTTONUP           As Long = &H205
Public Const WM_RBUTTONDBLCLK       As Long = &H206
Public Const WM_MBUTTONDOWN         As Long = &H207
Public Const WM_MBUTTONUP           As Long = &H208
Public Const WM_MBUTTONDBLCLK       As Long = &H209
Public Const WM_MOUSEWHEEL          As Long = &H20A
Public Const WM_XBUTTONDOWN         As Long = &H20B
Public Const WM_XBUTTONUP           As Long = &H20C
Public Const WM_XBUTTONDBLCLK       As Long = &H20D
Public Const WM_MOVING              As Long = &H216

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type Msg
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public cMouseEvents As clsMouseEvents

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long

Public Function MouseHook(ByVal idHook As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim bButtonPressed As Boolean
    Dim bX1Pressed As Boolean, bX2Pressed As Boolean
    
    Dim amsg As Msg
    
    Static LastXbutton As Integer
    
    bX1Pressed = CBool(GetAsyncKeyState(VK_XBUTTON1))
    bX2Pressed = CBool(GetAsyncKeyState(VK_XBUTTON2))
    
    ' We only need to process items with positive idhooks.
    If idHook >= 0 Then
              
        bButtonPressed = (wParam <> 0)
        Select Case wParam
        
            Case Is = WM_LBUTTONDOWN:   Call cMouseEvents.MouseButton(True, 1)
            Case Is = WM_LBUTTONUP:     Call cMouseEvents.MouseButton(False, 1)
        
            Case Is = WM_RBUTTONDOWN:   Call cMouseEvents.MouseButton(True, 2)
            Case Is = WM_RBUTTONUP:     Call cMouseEvents.MouseButton(False, 2)
        
            Case Is = WM_MBUTTONDOWN:   Call cMouseEvents.MouseButton(True, 4)
            Case Is = WM_MBUTTONUP:     Call cMouseEvents.MouseButton(False, 4)
        
            Case Is = WM_XBUTTONDOWN
                If bX1Pressed Then Call cMouseEvents.MouseButton(True, 8): LastXbutton = 1
                If bX2Pressed Then Call cMouseEvents.MouseButton(True, 16): LastXbutton = 2
                        
            Case Is = WM_XBUTTONUP
                If LastXbutton = 1 Then Call cMouseEvents.MouseButton(False, 8)
                If LastXbutton = 2 Then Call cMouseEvents.MouseButton(False, 16)
                    
            Case Is = WM_MOUSEWHEEL         ' MouseWheel
                Call GetMessage(amsg, 0, 0, 0)
                Call cMouseEvents.MouseWheelUsed(amsg.wParam > 0)
        
            Case Is = WM_SYSCOMMAND         ' Form unloading
                ' if the form is unloading we will unhook to prevent crashes
                If wParam = 61536 Then cMouseEvents.UnhookMouseEvents

        End Select
        
    End If
    
    ' Call the next hook
    MouseHook = CallNextHookEx(hHook, idHook, wParam, ByVal lParam)

End Function

