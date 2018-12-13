Attribute VB_Name = "mDebug"
'==================================================================================================
'mDebug.bas      9/2/05
'
'           PURPOSE:
'               Override the necessary api functions to track all allocated resources. Print a debug
'               message if any are not released before the program terminates.
'
'==================================================================================================

Option Explicit

#If bDebug Then
    
Public Enum eDEBUG_Resource
DEBUG_hMod
DEBUG_hIcon
DEBUG_hImagelist
DEBUG_hGlobal
DEBUG_hMem
DEBUG_hFile
DEBUG_hCursor
DEBUG_hBitmap
DEBUG_WinProp
DEBUG_hWnd
DEBUG_hHook
DEBUG_hMenu
DEBUG_hDcWindow
DEBUG_hAccel
DEBUG_Timer
DEBUG_hFont
DEBUG_hBrush
DEBUG_hPen
DEBUG_hDc
DEBUG_hMemCoTask
DEBUG_pIdl
End Enum
    
Private msName(DEBUG_hMod To DEBUG_pIdl) As String
Private miAllocated(DEBUG_hMod To DEBUG_pIdl) As Long
Private miPeak(DEBUG_hMod To DEBUG_pIdl) As Long
Private miTotal(DEBUG_hMod To DEBUG_pIdl) As Long
    
Private moDebug As pcDebug
    
Private Sub pInit()
    If moDebug Is Nothing Then
        msName(DEBUG_hMod) = "hMod"
        msName(DEBUG_hIcon) = "hIcon"
        msName(DEBUG_hImagelist) = "hImagelist"
        msName(DEBUG_hGlobal) = "hGlobal"
        msName(DEBUG_hMem) = "hMem"
        msName(DEBUG_hFile) = "hFile"
        msName(DEBUG_hCursor) = "hCursor"
        msName(DEBUG_hBitmap) = "hBitmap"
        msName(DEBUG_WinProp) = "Window Prop"
        msName(DEBUG_hWnd) = "hWnd"
        msName(DEBUG_hHook) = "hHook"
        msName(DEBUG_hMenu) = "hMenu"
        msName(DEBUG_hDcWindow) = "hDc (Window)"
        msName(DEBUG_hAccel) = "hAccel"
        msName(DEBUG_Timer) = "Windows Timer"
        msName(DEBUG_hFont) = "hFont"
        msName(DEBUG_hBrush) = "hBrush"
        msName(DEBUG_hPen) = "hPen"
        msName(DEBUG_hDc) = "hDc (Other)"
        msName(DEBUG_hMemCoTask) = "hMemCoTask"
        msName(DEBUG_pIdl) = "pIdl"
        Set moDebug = New pcDebug
    End If
    End Sub
    
Public Sub DEBUG_Add(ByVal iType As eDEBUG_Resource, ByVal iHandle As Long)
    pInit
    moDebug.Add msName(iType), iHandle
    miAllocated(iType) = miAllocated(iType) + OneL
    If miPeak(iType) < miAllocated(iType) Then miPeak(iType) = miAllocated(iType)
    miTotal(iType) = miTotal(iType) + OneL
    End Sub
    
Public Sub DEBUG_Remove(ByVal iType As eDEBUG_Resource, ByVal iHandle As Long)
    'debug.assert Not moDebug Is Nothing
    If Not moDebug Is Nothing Then moDebug.Remove msName(iType), iHandle
    miAllocated(iType) = miAllocated(iType) - OneL
    End Sub
    
Public Property Get DEBUG_Grid(ByVal X As Long, ByVal Y As eDEBUG_Resource) As String
    If Y = ZeroL Then
        Select Case X
        Case 0: DEBUG_Grid = "Name"
        Case 1: DEBUG_Grid = "Allocated"
        Case 2: DEBUG_Grid = "Peak"
        Case 3: DEBUG_Grid = "Total"
        End Select
    Else
        Select Case X
        Case 0: DEBUG_Grid = msName(Y)
        Case 1: DEBUG_Grid = miAllocated(Y)
        Case 2: DEBUG_Grid = miPeak(Y)
        Case 3: DEBUG_Grid = miTotal(Y)
        End Select
    End If
    End Property
    
Public Property Get DEBUG_GridCountY() As Long
    DEBUG_GridCountY = DEBUG_pIdl + OneL
    End Property
    
Public Property Get DEBUG_GridCountX() As Long
    DEBUG_GridCountX = 4&
    End Property
    
    
'###############################################
'##  KERNEL32
'###############################################
Public Function LoadLibrary(ByVal lp As Long) As Long
    LoadLibrary = vbComCtlTlb.LoadLibrary(ByVal lp)
    If LoadLibrary Then DEBUG_Add DEBUG_hMod, LoadLibrary
    End Function
    
Public Function FreeLibrary(ByVal hMod As Long) As Long
    'debug.assert hMod
    FreeLibrary = vbComCtlTlb.FreeLibrary(hMod)
    If hMod Then
        DEBUG_Remove DEBUG_hMod, hMod
    End If
    End Function
    
Public Function GlobalAlloc(ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    GlobalAlloc = vbComCtlTlb.GlobalAlloc(wFlags, dwBytes)
    If GlobalAlloc Then
        DEBUG_Add DEBUG_hGlobal, GlobalAlloc
    End If
    End Function
    
Public Function GlobalFree(ByVal hGlobal As Long) As Long
    GlobalFree = vbComCtlTlb.GlobalFree(hGlobal)
    If GlobalFree Then
        DEBUG_Remove DEBUG_hGlobal, hGlobal
    End If
    End Function
    
Public Function HeapAlloc(ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
    HeapAlloc = vbComCtlTlb.HeapAlloc(hHeap, dwFlags, dwBytes)
    If HeapAlloc Then
        DEBUG_Add DEBUG_hMem, HeapAlloc
    End If
    End Function
    
Public Function HeapFree(ByVal hHeap As Long, ByVal dwFlags As Long, ByVal hMem As Long) As Long
    HeapFree = vbComCtlTlb.HeapFree(hHeap, dwFlags, hMem)
    If HeapFree Then
        DEBUG_Remove DEBUG_hMem, hMem
    End If
    End Function
    
'###############################################
'##  COMCTL32
'###############################################
Public Function ImageList_GetIcon(ByVal hImagelist As Long, ByVal ImgIndex As Long, ByVal fuFlags As Long) As Long
    'debug.assert hImagelist
    ImageList_GetIcon = vbComCtlTlb.ImageList_GetIcon(hImagelist, ImgIndex, fuFlags)
    If ImageList_GetIcon Then
        DEBUG_Add DEBUG_hIcon, ImageList_GetIcon
    End If
    End Function
    
Public Function ImageList_Create(ByVal cx As Long, ByVal cy As Long, ByVal Flags As Long, ByVal Initial As Long, ByVal Grow As Long) As Long
    ImageList_Create = vbComCtlTlb.ImageList_Create(cx, cy, Flags, Initial, Grow)
    If ImageList_Create Then
        DEBUG_Add DEBUG_hImagelist, ImageList_Create
    End If
    End Function
    
Public Function ImageList_Destroy(ByVal hIml As Long) As Long
    ImageList_Destroy = vbComCtlTlb.ImageList_Destroy(hIml)
    If ImageList_Destroy Then
        DEBUG_Remove DEBUG_hImagelist, hIml
    End If
    End Function
    
    
'###############################################
'##  USER32
'###############################################
    
Public Function DestroyIcon(ByVal hIcon As Long) As Long
    DestroyIcon = vbComCtlTlb.DestroyIcon(hIcon)
    If DestroyIcon Then
        DEBUG_Remove DEBUG_hIcon, hIcon
    End If
    End Function
    
Public Function LoadImage(ByVal hInstance As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
    LoadImage = vbComCtlTlb.LoadImage(hInstance, ByVal lpsz, un1, n1, n2, un2)
    If LoadImage Then
        If un1 = imlIcon Then
            DEBUG_Add DEBUG_hIcon, LoadImage
        ElseIf un1 = imlCursor Then
            DEBUG_Add DEBUG_hCursor, LoadImage
        Else
            'debug.assert GetObjectType(LoadImage) = OBJ_BITMAP
            DEBUG_Add DEBUG_hBitmap, LoadImage
        End If
    End If
    End Function
    
Public Function SetProp(ByVal hWnd As Long, ByVal lpsz As Long, ByVal hData As Long) As Long
    If vbComCtlTlb.GetProp(hWnd, lpsz) = ZeroL Then
        DEBUG_Add DEBUG_WinProp, Hash(lpsz, lstrlen(lpsz)) + hWnd
    End If
    SetProp = vbComCtlTlb.SetProp(hWnd, lpsz, hData)
    End Function
    
Public Function RemoveProp(ByVal hWnd As Long, ByVal lpsz As Long) As Long
    RemoveProp = vbComCtlTlb.RemoveProp(hWnd, ByVal lpsz)
    DEBUG_Remove DEBUG_WinProp, Hash(lpsz, lstrlen(lpsz)) + hWnd
    End Function
    
Public Function CreateWindowEx(ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lpParam As Long) As Long
    CreateWindowEx = vbComCtlTlb.CreateWindowEx(dwExStyle, lpClassName, lpWindowName, dwStyle, X, Y, nWidth, nHeight, hWndParent, hMenu, hInstance, ByVal lpParam)
    If CreateWindowEx Then
        DEBUG_Add DEBUG_hWnd, CreateWindowEx
    End If
    End Function
    
Public Function DestroyWindow(ByVal hWnd As Long) As Long
    DestroyWindow = vbComCtlTlb.DestroyWindow(hWnd)
    If DestroyWindow Then
        DEBUG_Remove DEBUG_hWnd, hWnd
    End If
    End Function
    
Public Function SetWindowsHookEx(ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
    SetWindowsHookEx = vbComCtlTlb.SetWindowsHookEx(idHook, lpfn, hMod, dwThreadId)
    If SetWindowsHookEx Then
        DEBUG_Add DEBUG_hHook, SetWindowsHookEx
    End If
    End Function
    
Public Function UnHookWindowsHookEx(ByVal hHook As Long) As Long
    UnHookWindowsHookEx = vbComCtlTlb.UnHookWindowsHookEx(hHook)
    If UnHookWindowsHookEx Then
        DEBUG_Remove DEBUG_hHook, hHook
    End If
    End Function
    
Public Function CreatePopupMenu()
    CreatePopupMenu = vbComCtlTlb.CreatePopupMenu()
    If CreatePopupMenu Then
        DEBUG_Add DEBUG_hMenu, CreatePopupMenu
    End If
    End Function
    
Public Function DestroyMenu(ByVal hMenu As Long) As Long
    DestroyMenu = vbComCtlTlb.DestroyMenu(hMenu)
    If DestroyMenu Then
        DEBUG_Remove DEBUG_hMenu, hMenu
    End If
    End Function
    
Public Function GetDC(ByVal hWnd As Long) As Long
    GetDC = vbComCtlTlb.GetDC(hWnd)
    If GetDC Then
        DEBUG_Add DEBUG_hDcWindow, GetDC
    End If
    End Function
    
Public Function GetWindowDC(ByVal hWnd As Long) As Long
    GetWindowDC = vbComCtlTlb.GetWindowDC(hWnd)
    If GetWindowDC Then
        DEBUG_Add DEBUG_hDcWindow, GetWindowDC
    End If
    End Function
    
Public Function ReleaseDC(ByVal hWnd As Long, ByVal hdc As Long) As Long
    ReleaseDC = vbComCtlTlb.ReleaseDC(hWnd, hdc)
    If ReleaseDC Then
        DEBUG_Remove DEBUG_hDcWindow, hdc
    End If
    End Function

Public Function CreateAcceleratorTable(ByRef tAccel As ACCEL, ByVal cEntries As Long) As Long
    CreateAcceleratorTable = vbComCtlTlb.CreateAcceleratorTable(tAccel, cEntries)
    If CreateAcceleratorTable Then
        DEBUG_Add DEBUG_hAccel, CreateAcceleratorTable
    End If
    End Function
    
Public Function DestroyAcceleratorTable(ByVal hAccel As Long) As Long
    DestroyAcceleratorTable = vbComCtlTlb.DestroyAcceleratorTable(hAccel)
    If DestroyAcceleratorTable Then
        DEBUG_Remove DEBUG_hAccel, hAccel
    End If
    End Function
    
Public Function SetTimer(ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    'If the given timer is running then SetTimer would do this for us, but we need to
    'stop tracking the handle if that is the case.
    If hWnd Then KillTimer hWnd, nIDEvent
        
    SetTimer = vbComCtlTlb.SetTimer(hWnd, nIDEvent, uElapse, lpTimerFunc)
        
    If SetTimer Then
        If hWnd Then
            DEBUG_Add DEBUG_Timer, hWnd + nIDEvent
        Else
            DEBUG_Add DEBUG_Timer, SetTimer
        End If
    End If
    End Function
    
Public Function KillTimer(ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
    KillTimer = vbComCtlTlb.KillTimer(hWnd, nIDEvent)
    If KillTimer Then
        If hWnd Then
            DEBUG_Remove DEBUG_Timer, hWnd + nIDEvent
        Else
            DEBUG_Remove DEBUG_Timer, nIDEvent
        End If
    End If
    End Function
    
Public Function DestroyCursor(ByVal hCursor As Long) As Long
    DestroyCursor = vbComCtlTlb.DestroyCursor(hCursor)
    If DestroyCursor Then
        DEBUG_Remove DEBUG_hCursor, hCursor
    End If
    End Function
    
    
'###############################################
'##  GDI32
'###############################################
    
Public Function CreateFontIndirect(ByVal lpLogFont As Long) As Long
    CreateFontIndirect = vbComCtlTlb.CreateFontIndirect(ByVal lpLogFont)
    If CreateFontIndirect Then
        DEBUG_Add DEBUG_hFont, CreateFontIndirect
    End If
    End Function
    
Public Function CreateBrushIndirect(ByVal lpLogBrush As Long) As Long
    CreateBrushIndirect = vbComCtlTlb.CreateBrushIndirect(ByVal lpLogBrush)
    If CreateBrushIndirect Then
        DEBUG_Add DEBUG_hBrush, CreateBrushIndirect
    End If
    End Function
    
Public Function CreatePenIndirect(ByVal lpLogPen As Long) As Long
    CreatePenIndirect = vbComCtlTlb.CreatePenIndirect(ByVal lpLogPen)
    If CreatePenIndirect Then
        DEBUG_Add DEBUG_hPen, CreatePenIndirect
    End If
    End Function
    
Public Function CreateDC(ByVal lpDriverName As Long, ByVal lpDeviceName As Long, ByVal lpOutput As Long, ByVal lpInitData As Long) As Long
    CreateDC = vbComCtlTlb.CreateDC(ByVal lpDriverName, ByVal lpDeviceName, ByVal lpOutput, ByVal lpInitData)
    If CreateDC Then
        DEBUG_Add DEBUG_hDc, CreateDC
    End If
    End Function
        
Public Function CreateCompatibleDC(ByVal hdc As Long) As Long
    CreateCompatibleDC = vbComCtlTlb.CreateCompatibleDC(hdc)
    If CreateCompatibleDC Then
        DEBUG_Add DEBUG_hDc, CreateCompatibleDC
    End If
    End Function
        
Public Function DeleteDC(ByVal hdc As Long) As Long
    DeleteDC = vbComCtlTlb.DeleteDC(hdc)
    If DeleteDC Then
        DEBUG_Remove DEBUG_hDc, hdc
    End If
    End Function
    
Public Function CreateDIBSection(ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
    CreateDIBSection = vbComCtlTlb.CreateDIBSection(hdc, pBitmapInfo, un, lplpVoid, handle, dw)
    If CreateDIBSection Then
        DEBUG_Add DEBUG_hBitmap, CreateDIBSection
    End If
    End Function
    
Public Function CreateCompatibleBitmap(ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
    CreateCompatibleBitmap = vbComCtlTlb.CreateCompatibleBitmap(hdc, nWidth, nHeight)
    If CreateCompatibleBitmap Then
        DEBUG_Add DEBUG_hBitmap, CreateCompatibleBitmap
    End If
    End Function
    
Public Function DeleteObject(ByVal hObject As Long) As Long
    Dim liObjectType      As Long
    liObjectType = GetObjectType(hObject)
    DeleteObject = vbComCtlTlb.DeleteObject(hObject)
    If DeleteObject Then
        Select Case liObjectType
        Case OBJ_FONT
            DEBUG_Remove DEBUG_hFont, hObject
        Case OBJ_BRUSH
            DEBUG_Remove DEBUG_hBrush, hObject
        Case OBJ_PEN
            DEBUG_Remove DEBUG_hPen, hObject
        Case OBJ_BITMAP
            DEBUG_Remove DEBUG_hBitmap, hObject
        Case OBJ_DC, OBJ_MEMDC
            DEBUG_Remove DEBUG_hDc, hObject
        Case Else
            'debug.assert False
        End Select
    End If
    End Function
    
'###############################################
'##  OLE32
'###############################################
    
Public Function CoTaskMemAlloc(ByVal iBytes As Long) As Long
    CoTaskMemAlloc = vbComCtlTlb.CoTaskMemAlloc(iBytes)
    If CoTaskMemAlloc Then
        DEBUG_Add DEBUG_hMemCoTask, CoTaskMemAlloc
    End If
    End Function
    
Public Function CoTaskMemFree(ByVal hMem As Long) As Long
    CoTaskMemFree = vbComCtlTlb.CoTaskMemFree(hMem)
    If CoTaskMemFree Then
        DEBUG_Remove DEBUG_hMemCoTask, hMem
    End If
End Function
#End If
