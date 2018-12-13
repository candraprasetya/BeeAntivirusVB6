VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcImageDrag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'pcImageDrag.cls      4/20/05
'
'           PURPOSE:
'               Manage an image drag operation using the comctl32 ImageList_Drag* functions.
'
'           LINEAGE:
'               LVDrag.vbp by Brad Martinez http://www.mvps.org
'
'==================================================================================================

Option Explicit

Implements iHook

Private mhIml      As Long
Private mhWndLock  As Long

Private Sub Class_Terminate()
    pDestroy
End Sub

Private Sub iHook_After(lReturn As Long, ByVal iType As eHookType, ByVal nCode As eHookCode, ByVal wParam As Long, ByVal lParam As Long): End Sub
Private Sub iHook_Before(bHandled As Boolean, lReturn As Long, ByVal iType As eHookType, ByVal nCode As eHookCode, ByVal wParam As Long, ByVal lParam As Long)
    If (nCode = HC_ACTION) Then pMove MemOffset32(lParam, ZeroL), MemOffset32(lParam, 4&)
End Sub

Private Sub pMove(ByVal X As Long, ByVal Y As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Move the image drag to the new location.
    '---------------------------------------------------------------------------------------
    Dim lhWnd      As Long
    
    lhWnd = RootParent(WindowFromPoint(X, Y))
    'lhWnd = WindowFromPoint(x, y)
    
    Dim tR      As RECT
    If GetWindowRect(lhWnd, tR) Then
        X = X - tR.Left
        Y = Y - tR.Top
    End If
    
    If lhWnd <> mhWndLock Then
        ImageList_DragLeave mhWndLock
        ImageList_DragEnter lhWnd, X, Y
        mhWndLock = lhWnd
    End If
    
    ImageList_DragMove X, Y
    
End Sub

Private Sub pDestroy()
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : End the image drag operation.
    '---------------------------------------------------------------------------------------
    If mhIml Then
        ImageList_DragLeave mhWndLock
        ImageList_EndDrag
        ImageList_Destroy mhIml
   
        Hook_Remove Me, WH_MOUSE
    End If
End Sub

Public Sub StartDrag(ByVal oDib As pcDibSection, ByVal X As Long, ByVal Y As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Begin a new image drag operation.
    '---------------------------------------------------------------------------------------
    pDestroy
    
    mhIml = ImageList_Create(oDib.Width, oDib.Height, ILC_MASK, OneL, ZeroL)
    
    If mhIml Then
        
        ImageList_AddMasked mhIml, oDib.hBitmap, ImageDrag_TransColor
        ImageList_BeginDrag mhIml, ZeroL, X, Y
        Hook_Install Me, WH_MOUSE
        
    End If
    
End Sub

Public Sub StopDrag()
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : End the image drag operation.
    '---------------------------------------------------------------------------------------
    pDestroy
End Sub

Public Sub Show(ByVal bVal As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Show or hide the image.
    '---------------------------------------------------------------------------------------
    ImageList_DragShowNolock -CLng(bVal)
End Sub
