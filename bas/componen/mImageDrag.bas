Attribute VB_Name = "mImageDrag"
'==================================================================================================
'mImageDrag.bas      4/21/05
'
'           PURPOSE:
'               Facilitate image dragging using pcImageDrag.cls or pcImageDragAlpha.cls.
'
'           LINEAGE:
'               Scrolling inspired by Chapter 13 of Inside OLE 2nd Edition By Kraig Brockschmidt - MSPRESS
'
'==================================================================================================
Option Explicit

Public Const ImageDrag_TransColor As Long = &H12345

Private Type tScrollData
    iData(SB_HORZ To SB_VERT) As Long
End Type

Private moImageDrag  As pcImageDrag

Public Sub ImageDrag_Stop()
    '---------------------------------------------------------------------------------------
    ' Date      : 3/6/05
    ' Purpose   : Stop any current image drag operation.
    '---------------------------------------------------------------------------------------
    Set moImageDrag = Nothing
End Sub

Public Sub ImageDrag_Start(ByVal oDib As pcDibSection, ByVal X As Long, ByVal Y As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 3/6/05
    ' Purpose   : Start an image drag operation.
    '---------------------------------------------------------------------------------------
    If oDib.hBitmap Then
        If ImageDrag_Alpha Then
            Set moImageDrag = New pcImageDragAlpha
        Else
            Set moImageDrag = New pcImageDrag
        End If
        moImageDrag.StartDrag oDib, X, Y
    End If
End Sub

Public Sub ImageDrag_Show(ByVal bShow As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 3/6/05
    ' Purpose   : Show or hide the image drag to redraw correctly.
    '---------------------------------------------------------------------------------------
    If Not moImageDrag Is Nothing Then moImageDrag.Show bShow
End Sub

Public Sub ImageDrag_Move(ByVal hWndTarget As Long, ByVal X As Long, ByVal Y As Long, ByVal State As evbComCtlOleDragOverState)
    '---------------------------------------------------------------------------------------
    ' Date      : 3/6/05
    ' Purpose   : Scroll after the mouse enters an edge of the target window.
    '---------------------------------------------------------------------------------------
    
    Const DD_DEFSCROLLDELAY As Long = 50
    Const DD_DEFSCROLLINSET As Long = 11
    Const DD_DEFSCROLLINTERVAL As Long = 50
    
    Static stWParam As tScrollData
    Static stTick As tScrollData
    
    Dim i      As Long
    
    If State = vbccOleDragLeave Then
        For i = SB_HORZ To SB_VERT
            stWParam.iData(i) = ZeroL
            stTick.iData(i) = ZeroL
        Next
        Exit Sub
    End If
    
    Dim ltMessage         As tScrollData
    Dim ltPos             As tScrollData
    Dim ltMin             As tScrollData
    Dim ltMax             As tScrollData
    Dim ltSize            As tScrollData
    Dim ltWindowSize      As tScrollData
    Dim ltCursor          As tScrollData
    
    ltCursor.iData(SB_HORZ) = X
    ltCursor.iData(SB_VERT) = Y
    
    Dim ltSI      As SCROLLINFO
    ltSI.cbSize = Len(ltSI)
    ltSI.fMask = SIF_RANGE Or SIF_PAGE Or SIF_POS
    
    If hWndTarget Then
        Dim ltRect      As RECT
        
        GetClientRect hWndTarget, ltRect
        ltSize.iData(SB_HORZ) = ltRect.Right
        ltSize.iData(SB_VERT) = ltRect.Bottom
        
        GetWindowRect hWndTarget, ltRect
        ltWindowSize.iData(SB_HORZ) = ltRect.Right - ltRect.Left - GetSystemMetrics(SM_CYHSCROLL)
        ltWindowSize.iData(SB_VERT) = ltRect.Bottom - ltRect.Top - GetSystemMetrics(SM_CXVSCROLL)
        
        For i = SB_HORZ To SB_VERT
            If ltCursor.iData(i) > ltSize.iData(i) Then Exit Sub
        Next
        
        Dim lwParam      As Long
        
        For i = SB_HORZ To SB_VERT
            GetScrollInfo hWndTarget, i, ltSI
            
            ltPos.iData(i) = ltSI.nPos
            ltMin.iData(i) = ltSI.nMin
            ltMax.iData(i) = ltSI.nMax - ltSI.nPage
            
            If ltWindowSize.iData((Not i) And SB_VERT) >= ltSize.iData((Not i) And SB_VERT) Then
                
                lwParam = NegOneL
                
                If ltCursor.iData(i) < DD_DEFSCROLLINSET Then
                    If ltPos.iData(i) > ltMin.iData(i) Then
                        lwParam = SB_LINEUP
                    ElseIf ltPos.iData(i) = ltMax.iData(i) Then
                        lwParam = SB_TOP
                    End If
                ElseIf ltCursor.iData(i) > ltSize.iData(i) - DD_DEFSCROLLINSET Then
                    If ltPos.iData(i) < ltMax.iData(i) Then
                        lwParam = SB_LINEDOWN
                    ElseIf ltPos.iData(i) = ltMax.iData(i) Then
                        lwParam = SB_BOTTOM
                    End If
                End If
                
                If lwParam <> NegOneL Then
                    
                    If lwParam <> stWParam.iData(i) Then stTick.iData(i) = ZeroL
                    stWParam.iData(i) = lwParam
                    If stTick.iData(i) Then
                        If TickDiff(GetTickCount(), stTick.iData(i)) >= DD_DEFSCROLLINTERVAL Then
                            stTick.iData(i) = GetTickCount()
                            ltMessage.iData(i) = WM_HSCROLL + i
                        End If
                    Else
                        stTick.iData(i) = GetTickCount()
                        If stTick.iData(i) < &H7FFFFFFF - DD_DEFSCROLLDELAY Then stTick.iData(i) = stTick.iData(i) + DD_DEFSCROLLDELAY
                    End If
                Else
                    stWParam.iData(i) = ZeroL
                    stTick.iData(i) = ZeroL
                End If
            End If
        Next
        
        If ltMessage.iData(SB_HORZ) Or ltMessage.iData(SB_VERT) Then
            ImageDrag_Show False
            For i = SB_HORZ To SB_VERT
                If ltMessage.iData(i) Then SendMessage hWndTarget, ltMessage.iData(i), stWParam.iData(i), ByVal ZeroL
            Next
            ImageDrag_Show True
        End If
        
    End If
    
End Sub

Public Property Get ImageDrag_Alpha() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 3/6/05
    ' Purpose   : Return a value indicating whether the system supports an alpha blended
    '             image drag.  Otherwise the ImageList_Drag* functions are used.
    '---------------------------------------------------------------------------------------
    Static bInit As Boolean
    Static bIsAvailable As Boolean
    
    If Not bInit Then
        bIsAvailable = IsApiAvailable("user32.dll", "UpdateLayeredWindow")
        bInit = True
    End If
    
    ImageDrag_Alpha = bIsAvailable And (SystemColorDepth > 8&)
    
End Property
