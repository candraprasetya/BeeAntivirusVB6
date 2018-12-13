Attribute VB_Name = "basUSB"
' Untuk Detek Device yang masuk - dipakai untuk detek USB

' Help menggunakan code ini, :D


Private Type GUID
    Data1(1 To 4) As Byte
    Data2(1 To 2) As Byte
    Data3(1 To 2) As Byte
    Data4(1 To 8) As Byte
End Type

Private Type DEV_BROADCAST_DEVICEINTERFACE
  dbcc_size As Long
  dbcc_devicetype As Long
  dbcc_reserved As Long
  dbcc_classguid As GUID
  dbcc_name As Long
End Type

Private Type DEV_BROADCAST_DEVICEINTERFACE2
  dbcc_size As Long
  dbcc_devicetype As Long
  dbcc_reserved As Long
  dbcc_classguid As GUID
  dbcc_name As String * 1024
End Type

Private Declare Function RegisterDeviceNotification Lib "user32.dll" _
            Alias "RegisterDeviceNotificationA" ( _
            ByVal hRecipient As Long, _
            NotificationFilter As Any, _
            ByVal Flags As Long) As Long
            
Private Declare Function UnregisterDeviceNotification Lib "user32.dll" _
             (ByVal hRecipient As Long) As Long
            
Private Const DEVICE_NOTIFY_ALL_INTERFACE_CLASSES = &H4
Private Const DEVICE_NOTIFY_WINDOW_HANDLE = 0


 
Private Declare Sub CopyMemory Lib "kernel32" Alias _
    "RtlMoveMemory" (Destination As Any, Source As Any, _
    ByVal Length As Long)

Private Declare Function CallWindowProc Lib "user32" _
  Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                           ByVal hWnd As Long, _
                           ByVal Msg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long

Private Declare Function GetWindowLong Lib "user32" _
  Alias "GetWindowLongA" (ByVal hWnd As Long, _
                          ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" _
  Alias "SetWindowLongA" (ByVal hWnd As Long, _
                          ByVal nIndex As Long, _
                          ByVal dwNewLong As Long) As Long

Private Const GWL_WNDPROC = (-4)
Private Const WM_DEVICECHANGE = &H219
Private Const UNSAFE_REMOVE = &H1C

Private glngPrevWndProc As Long

Private Type DEV_BROADCAST_HDR
  dbch_size As Long
  dbch_devicetype As Long
  dbch_reserved As Long
End Type

Private Const DBT_CONFIGCHANGECANCELED As Long = 25
Private Const DBT_CONFIGCHANGED As Long = 24
Private Const DBT_CUSTOMEVENT As Long = 32744
Private Const DBT_DEVICEARRIVAL As Long = 32768
Private Const DBT_DEVICEQUERYREMOVE As Long = 32769
Private Const DBT_DEVICEQUERYREMOVEFAILED As Long = 32770
Private Const DBT_DEVICEREMOVECOMPLETE As Long = 32772
Private Const DBT_DEVICEREMOVEPENDING As Long = 32771
Private Const DBT_DEVICETYPESPECIFIC As Long = 32773
Private Const DBT_DEVNODES_CHANGED As Long = 7
Private Const DBT_QUERYCHANGECONFIG As Long = 23
Private Const DBT_USERDEFINED As Long = 65535

Private Const DBT_DEVTYP_OEM                 As Long = 0
Private Const DBT_DEVTYP_DEVNODE             As Long = 1
Private Const DBT_DEVTYP_VOLUME              As Long = 2
Private Const DBT_DEVTYP_PORT                As Long = 3
Private Const DBT_DEVTYP_NET                 As Long = 4
Private Const DBT_DEVTYP_DEVICEINTERFACE     As Long = 5
Private Const DBT_DEVTYP_HANDLE              As Long = 6

 
Dim hDevNotify As Long

Private Function MyWindowProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim dbHdr As DEV_BROADCAST_HDR
    
    Dim dbDdb As DEV_BROADCAST_DEVICEINTERFACE2
    
    Select Case Msg
        Case WM_DEVICECHANGE
        
            Debug.Print "WM_DEVICECHANGE " & Msg
   
            Select Case wParam
            
                Case DBT_CONFIGCHANGECANCELED
                    Debug.Print "wParam = DBT_CONFIGCHANGECANCELED"
                    
                Case DBT_CONFIGCHANGED
                    Debug.Print "wParam = DBT_CONFIGCHANGED"
                    
                Case DBT_CUSTOMEVENT
                    Debug.Print "wParam = DBT_CUSTOMEVENT"
                    
                Case DBT_DEVICEARRIVAL
                    Debug.Print "wParam = DBT_DEVICEARRIVAL"
                    
                     CopyMemory dbHdr, ByVal (lParam), Len(dbHdr)
                 
                        Select Case dbHdr.dbch_devicetype
                            Case DBT_DEVTYP_OEM
                                Debug.Print "Device Type: DBT_DEVTYP_OEM"
                                
                            Case DBT_DEVTYP_DEVNODE
                                Debug.Print "Device Type: DBT_DEVTYP_DEVNODE"
                                
                            Case DBT_DEVTYP_VOLUME
                                Debug.Print "Device Type: DBT_DEVTYP_VOLUME"
                                
                            Case DBT_DEVTYP_PORT
                                Debug.Print "Device Type: DBT_DEVTYP_PORT"
                                
                            Case DBT_DEVTYP_NET
                                Debug.Print "Device Type: DBT_DEVTYP_NET"
                                
                            Case DBT_DEVTYP_DEVICEINTERFACE
                                Debug.Print "Device Type: DBT_DEVTYP_DEVICEINTERFACE"
                                
                                CopyMemory dbDdb, ByVal (lParam), ByVal (dbHdr.dbch_size)
                                
                                Debug.Print "Device Name: " & Mid(dbDdb.dbcc_name, 1, InStr(dbDdb.dbcc_name, Chr(0)))
                            
                            Case DBT_DEVTYP_HANDLE
                                Debug.Print "Device Type: DBT_DEVTYP_HANDLE"
                                
                            Case Else
                                Debug.Print "Device Type unknown: " & CStr(dbHdr.dbch_devicetype)
                                
                        End Select
                        
                 
                Case DBT_DEVICEQUERYREMOVE
                    Debug.Print "wParam = DBT_DEVICEQUERYREMOVE"
                    
                Case DBT_DEVICEQUERYREMOVEFAILED
                    Debug.Print "wParam = DBT_DEVICEQUERYREMOVEFAILED"
                    
                Case DBT_DEVICEREMOVECOMPLETE
                    Debug.Print "wParam = DBT_DEVICEREMOVECOMPLETE"
                    
                Case DBT_DEVICEREMOVEPENDING
                    Debug.Print "wParam = DBT_DEVICEREMOVEPENDING"
                    
                Case DBT_DEVICETYPESPECIFIC
                    Debug.Print "wParam = DBT_DEVICETYPESPECIFIC"
                    
                Case DBT_DEVNODES_CHANGED
                    Debug.Print "wParam = DBT_DEVNODES_CHANGED"
                    
                Case DBT_QUERYCHANGECONFIG
                    Debug.Print "wParam = DBT_QUERYCHANGECONFIG"
                    
                Case DBT_USERDEFINED
                    Debug.Print "wParam = DBT_USERDEFINED"
                    
                Case Else
                    Debug.Print "wParam = unknown = " & wParam
            
            End Select
        
    End Select
  
 ' pass the rest messages onto VB's own Window Procedure
  MyWindowProc = CallWindowProc(glngPrevWndProc, hWnd, Msg, wParam, lParam)
End Function


Private Function DoRegisterDeviceInterface(hWnd As Long, ByRef hDevNotify As Long) As Boolean

    Dim NotificationFilter As DEV_BROADCAST_DEVICEINTERFACE
        
    NotificationFilter.dbcc_size = Len(NotificationFilter)
    NotificationFilter.dbcc_devicetype = DBT_DEVTYP_DEVICEINTERFACE
    
    hDevNotify = RegisterDeviceNotification(hWnd, NotificationFilter, DEVICE_NOTIFY_WINDOW_HANDLE Or DEVICE_NOTIFY_ALL_INTERFACE_CLASSES)
    
    If hDevNotify = 0 Then
        MsgBox "RegisterDeviceNotification failed: " & CStr(Err.LastDllError), vbOKOnly
        DoRegisterDeviceInterface = False
        Exit Function
    End If
    
    DoRegisterDeviceInterface = True
End Function


' Yang dipanggil........................

Public Sub RegisterDevice(Frm As Form) ' Panggil ini dulu
    Call DoRegisterDeviceInterface(Frm.hWnd, hDevNotify)
    glngPrevWndProc = GetWindowLong(Frm.hWnd, GWL_WNDPROC)
    SetWindowLong Frm.hWnd, GWL_WNDPROC, AddressOf MyWindowProc
End Sub

Public Sub UnRegisterDevice(Frm As Form) ' Panggil ini klo mau keluar
   If hDevNotify <> 0 Then
       Call UnregisterDeviceNotification(hDevNotify)
   End If
   'pass control back to previous windows
   SetWindowLong Frm.hWnd, GWL_WNDPROC, glngPrevWndProc
End Sub


' Pakai yang ini dulu (cuma detek flasdish baru dengan vol terendah)

Public Function AdakahFDBaru(LastFDVolume As Long) As Boolean
Dim MyCounter As Byte
Dim lgnVol    As Byte
For MyCounter = 67 To 85 ' dari C
    If GetDriveType(Chr(MyCounter) & ":\") = 2 Then
       If MyCounter > LastFDVolume Then
          AdakahFDBaru = True
          LastFlashVolume = MyCounter
          Exit Function
       ElseIf MyCounter = LastFDVolume Then
          AdakahFDBaru = False
          lgnVol = MyCounter
      ElseIf MyCounter < LastFDVolume Then
          AdakahFDBaru = False
          lgnVol = MyCounter
       End If
    End If
Next
LastFlashVolume = lgnVol
AdakahFDBaru = False
End Function

'panggil sekali
Public Sub GetLasFDVolume()
Dim MyCounter As Byte
For MyCounter = 67 To 85 ' dari C
    If GetDriveType(Chr(MyCounter) & ":\") = 2 Then
       LastFlashVolume = MyCounter
    End If
Next
End Sub


