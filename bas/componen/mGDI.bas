Attribute VB_Name = "mGDI"
'==================================================================================================
'mGDI.bas              2/13/05
'
'           PURPOSE:
'               Manage the creation/destruction of gdi objects.
'               Notify the programmer if gdi handles are leaked.
'               Identical brushes, pens and fonts are pooled using pcGDIObjectStore.cls.
'               Provide debugging and statistic functions that are enabled using compiler switches.
'
'           LINEAGE:
'               "GDI Font Management" by LaVolpe at www.pscode.com
'
'==================================================================================================

Option Explicit

Private moBrushes  As pcGDIObjectStore
Private moPens     As pcGDIObjectStore
Private moFonts    As pcGDIObjectStore

Public Function GdiMgr_CreateSolidBrush(ByVal iColor As OLE_COLOR) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Create a solid brush or return a cached brush handle.
    '---------------------------------------------------------------------------------------
    
    Dim ltBrush      As LOGBRUSH
    With ltBrush
        .lbColor = iColor
        .lbStyle = BS_SOLID
    End With
    
    GdiMgr_CreateSolidBrush = GdiMgr_CreateBrushIndirect(ltBrush)
    
End Function

Public Function GdiMgr_CreateBrushIndirect(ByRef tLogBrush As LOGBRUSH) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Create a solid brush or return a cached brush handle.
    '---------------------------------------------------------------------------------------
    
    If moBrushes Is Nothing Then
        Set moBrushes = New pcGDIObjectStore
        moBrushes.Init OBJ_BRUSH
    End If
    
    OleTranslateColor tLogBrush.lbColor, ZeroL, tLogBrush.lbColor
    GdiMgr_CreateBrushIndirect = moBrushes.AddRef(VarPtr(tLogBrush))
    'debug.assert GdiMgr_CreateBrushIndirect
    
End Function

Public Function GdiMgr_DeleteBrush(ByVal hBrush As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/13/05
    ' Purpose   : Decrement a brush reference count, and delete it when no longer used.
    '---------------------------------------------------------------------------------------
    'debug.assert GetObjectType(hBrush) = OBJ_BRUSH
    'debug.assert Not moBrushes Is Nothing
    
    GdiMgr_DeleteBrush = moBrushes.Release(hBrush)

End Function


Public Function GdiMgr_CreateFontIndirect(ByRef tLogFont As LOGFONT) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Create a font or return a cached font handle
    '---------------------------------------------------------------------------------------
    If moFonts Is Nothing Then
        Set moFonts = New pcGDIObjectStore
        moFonts.Init OBJ_FONT
    End If
    
    GdiMgr_CreateFontIndirect = moFonts.AddRef(VarPtr(tLogFont))
    'debug.assert GdiMgr_CreateFontIndirect
    
End Function

Public Function GdiMgr_DeleteFont(ByVal hFont As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/13/05
    ' Purpose   : Decrement a font reference count, and delete it when no longer used.
    '---------------------------------------------------------------------------------------
    'debug.assert GetObjectType(hFont) = OBJ_FONT
    'debug.assert Not moFonts Is Nothing
    
    GdiMgr_DeleteFont = moFonts.Release(hFont)
End Function


Public Function GdiMgr_CreatePen(ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As OLE_COLOR) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Create a pen or return a cached pen handle
    '---------------------------------------------------------------------------------------
    Dim ltPen      As LOGPEN
    
    ltPen.lopnColor = crColor
    ltPen.lopnStyle = nPenStyle
    ltPen.lopnWidth.X = nWidth
    
    GdiMgr_CreatePen = GdiMgr_CreatePenIndirect(ltPen)
    
End Function

Public Function GdiMgr_CreatePenIndirect(ByRef tLogPen As LOGPEN) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Create a pen or return a cached pen handle
    '---------------------------------------------------------------------------------------
    If moPens Is Nothing Then
        Set moPens = New pcGDIObjectStore
        moPens.Init OBJ_PEN
    End If
    
    OleTranslateColor tLogPen.lopnColor, ZeroL, tLogPen.lopnColor
    GdiMgr_CreatePenIndirect = moPens.AddRef(VarPtr(tLogPen))
    'debug.assert GdiMgr_CreatePenIndirect
    
End Function

Public Function GdiMgr_DeletePen(ByVal hPen As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/13/05
    ' Purpose   : Decrement a pen reference count, and delete it when no longer used.
    '---------------------------------------------------------------------------------------
    'debug.assert GetObjectType(hPen) = OBJ_PEN
    'debug.assert Not moPens Is Nothing
    
    GdiMgr_DeletePen = moPens.Release(hPen)

End Function


Public Function CreateDisplayDC() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/13/05
    ' Purpose   : Create a dc for the display driver.
    '---------------------------------------------------------------------------------------
    Dim lsAnsi      As String
    lsAnsi = StrConv(DisplayDriver & vbNullChar, vbFromUnicode)
    CreateDisplayDC = CreateDC(ByVal StrPtr(lsAnsi), ByVal ZeroL, ByVal ZeroL, ByVal ZeroL)
End Function

Public Sub DrawGradient(ByVal hdc As Long, _
ByVal iLeft As Long, ByVal iTop As Long, _
ByVal iWidth As Long, ByVal iHeight As Long, _
ByVal iColorFrom As OLE_COLOR, ByVal iColorTo As OLE_COLOR, _
Optional ByVal bVertical As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Draw a gradient using the StretchDIBits function.
    ' Lineage   : www.pscode.com submission "Let's talk about speed"
    '---------------------------------------------------------------------------------------

    Dim ltBits()      As RGBQUAD, ltBIH As BITMAPINFOHEADER

    Dim r  As Long, G  As Long, B  As Long
    Dim dR As Long, dG As Long, dB As Long
    Dim d  As Long, dEnd As Long

    If bVertical Then
        dEnd = iHeight
        'swap to/from colors
        iColorTo = iColorTo Xor iColorFrom
        iColorFrom = iColorTo Xor iColorFrom
        iColorTo = iColorFrom Xor iColorTo
    Else
        dEnd = iWidth
    End If

    If dEnd > ZeroL Then

        iColorTo = TranslateColor(iColorTo)
        iColorFrom = TranslateColor(iColorFrom)

        'split from color to R, G, B
        r = iColorFrom And &HFF&
        iColorFrom = iColorFrom \ &H100&
        G = iColorFrom And &HFF&
        iColorFrom = iColorFrom \ &H100&
        B = iColorFrom And &HFF&

        'get the relative changes in R, G, B
        dR = (iColorTo And &HFF&) - r
        iColorTo = iColorTo \ &H100&
        dG = (iColorTo And &HFF&) - G
        iColorTo = iColorTo \ &H100&
        dB = (iColorTo And &HFF&) - B

        ReDim ltBits(0 To dEnd - 1&)

        For d = ZeroL To dEnd - 1&
            With ltBits(d)
                .rgbRed = (r + dR * d \ dEnd)
                .rgbGreen = (G + dG * d \ dEnd)
                .rgbBlue = (B + dB * d \ dEnd)
            End With
        Next

        With ltBIH
            .biSize = Len(ltBIH)
            .biBitCount = 32&
            .biPlanes = 1&

            If bVertical Then
                .biWidth = 1&
                .biHeight = dEnd
            Else
                .biWidth = dEnd
                .biHeight = 1&
            End If

            StretchDIBits hdc, _
            iLeft, iTop, iLeft + iWidth, iTop + iHeight, _
            ZeroL, ZeroL, .biWidth, .biHeight, _
            ltBits(0), ltBIH, ZeroL, vbSrcCopy

        End With
    End If
End Sub

Public Sub ImageListDraw( _
ByVal hIml As Long, _
ByVal iIndex As Long, _
ByVal hdc As Long, _
ByVal X As Long, _
ByVal Y As Long, _
Optional ByVal iStyle As eImlDrawStyle, _
Optional ByVal iCutDitherColor As OLE_COLOR = NegOneL)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Draw the given icon in the style and position specified on the dc.
    ' Lineage   : www.vbaccelerator.com
    '---------------------------------------------------------------------------------------

    Dim hIcon         As Long
    Dim lFlags        As Long
    Dim liHeight      As Long
    Dim liWidth       As Long
    
    If iIndex > NegOneL Then
        If hIml Then
            lFlags = ILD_TRANSPARENT
            If iStyle = imlDrawCut Or iStyle = imlDrawSelected Then lFlags = lFlags Or ILD_SELECTED

            If iStyle = imlDrawCut Then
                iCutDitherColor = TranslateColor(iCutDitherColor)
                If (iCutDitherColor = NegOneL) Then iCutDitherColor = GetSysColor(COLOR_WINDOW)
                ImageList_DrawEx hIml, iIndex, hdc, X, Y, 0, 0, NegOneL, iCutDitherColor, lFlags
            ElseIf iStyle = imlDrawDisabled Then
                ImageList_GetIconSize hIml, liWidth, liHeight
                hIcon = ImageList_GetIcon(hIml, iIndex, 0)
                If hIcon Then
                    ' Draw it disabled at x,y:
                    DrawState hdc, 0, 0, hIcon, 0, X, Y, liWidth, liHeight, DST_ICON Or DSS_DISABLED
                    ' Clear up the icon:
                    DestroyIcon hIcon
                End If
            Else
                ' Standard draw:
                ImageList_Draw hIml, iIndex, hdc, X, Y, lFlags
            End If
        End If
    End If

End Sub
