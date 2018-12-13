VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcImageDragAlpha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'pcImageDragAlpha.cls      4/20/05
'
'           PURPOSE:
'               Create a layered window and use alpha blending to create an image drag effect.
'
'==================================================================================================

Option Explicit

Implements pcImageDrag
Implements iHook

Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private mhWnd     As Long
Private mhDc      As Long
Private mhBmpOld  As Long
Private moDib     As pcDibSection

Private miWidth As Long
Private miHeight As Long

Private miHotSpotX As Long
Private miHotSpotY As Long

Private Sub Class_Terminate()
    pDestroy
End Sub

Private Sub iHook_After(lReturn As Long, ByVal iType As eHookType, ByVal nCode As eHookCode, ByVal wParam As Long, ByVal lParam As Long): End Sub
Private Sub iHook_Before(bHandled As Boolean, lReturn As Long, ByVal iType As eHookType, ByVal nCode As eHookCode, ByVal wParam As Long, ByVal lParam As Long)
    If (nCode = HC_ACTION) Then
        If (wParam = WM_MOUSEMOVE) Then MoveWindow mhWnd, MemOffset32(lParam, ZeroL) - miHotSpotX, MemOffset32(lParam, 4&) - miHotSpotY, miWidth, miHeight, OneL
    End If
End Sub

Private Sub pDestroy()
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Release all resources.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        DestroyWindow mhWnd
        mhWnd = ZeroL
        Hook_Remove Me, WH_MOUSE
    End If
    
    If mhDc Then
        If mhBmpOld Then SelectObject mhDc, mhBmpOld
        DeleteDC mhDc
    End If
    mhBmpOld = ZeroL
    mhDc = ZeroL
    
    Set moDib = Nothing
    
End Sub

Private Sub pcImageDrag_Show(ByVal bVal As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Not necessary.  The alpha blended window draws much more cleanly than
    '             the ImageList_Drag* functions.
    '---------------------------------------------------------------------------------------
End Sub

Private Sub pcImageDrag_StartDrag(ByVal oDib As pcDibSection, ByVal X As Long, ByVal Y As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Create the layered window and begin the drag operation.
    '---------------------------------------------------------------------------------------
    pDestroy
    
    Set moDib = oDib
    
    Dim lhDc      As Long
    lhDc = CreateDisplayDC()
    If lhDc Then
        mhDc = CreateCompatibleDC(lhDc)
        If mhDc Then mhBmpOld = SelectObject(mhDc, oDib.hBitmap)
        DeleteDC lhDc
    End If
    
    miHeight = oDib.Height
    miWidth = oDib.Width
    
    Dim ltCursor      As POINT
    GetCursorPos ltCursor
    
    Dim lsAnsi      As String
    lsAnsi = StrConv(WC_DIALOG & vbNullChar, vbFromUnicode)
    mhWnd = CreateWindowEx(WS_EX_TOOLWINDOW Or WS_EX_TRANSPARENT Or WS_EX_LAYERED, StrPtr(lsAnsi), ZeroL, WS_POPUP, ltCursor.X - X, ltCursor.Y - Y, miWidth, miHeight, ZeroL, ZeroL, App.hInstance, ByVal ZeroL)
    
    If mhWnd Then
        miHotSpotX = X
        miHotSpotY = Y
        
        Dim ltPoint      As POINT
        ltPoint.X = 0
        ltPoint.Y = 0
        
        Dim ltSize      As Size
        ltSize.cx = miWidth
        ltSize.cy = miHeight
         
        Dim tBlend      As BLENDFUNCTION
        With tBlend
            .AlphaFormat = AC_SRC_ALPHA
            .BlendFlags = 0
            .BlendOp = AC_SRC_OVER
            .SourceConstantAlpha = 255
        End With
        
        UpdateLayeredWindow mhWnd, ZeroL, ByVal ZeroL, ltSize, mhDc, ltPoint, ZeroL, tBlend, ULW_ALPHA
        SetWindowPos mhWnd, HWND_TOPMOST, ZeroL, ZeroL, ZeroL, ZeroL, SWP_SHOWWINDOW Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE
        
        Hook_Install Me, WH_MOUSE
        
    End If
    
End Sub

Private Sub pcImageDrag_StopDrag()
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Stop any current drag operation.
    '---------------------------------------------------------------------------------------
    pDestroy
End Sub
