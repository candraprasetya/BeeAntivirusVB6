VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cImageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'cImageList.cls                  2/25/05
'
'           GENERAL PURPOSE:
'               Manage an api image list.
'
'           LINEAGE:
'               cImageList.cls from www.vbaccelerator.com
'
'==================================================================================================

Option Explicit

Public Enum eImlDrawStyle
    imlDrawNormal
    imlDrawDisabled
    imlDrawCut
    imlDrawSelected
End Enum

Public Enum eImlImageType
    imlBitmap = 0
    imlIcon = 1
    imlCursor = 2
End Enum

Public Enum eImlColorDepth
    imlColor = &H0
    imlColor4 = &H4
    imlColor8 = &H8
    imlColor16 = &H10
    imlColor24 = &H18
    imlColor32 = &H20
End Enum

Event Changed()

Private mhIml         As Long
Private miColorDepth  As Long

Private Sub Class_Terminate()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/25/05
    ' Purpose   : Destroy the imagelist if it exists.
    '---------------------------------------------------------------------------------------
    If mhIml Then ImageList_Destroy mhIml
End Sub

Friend Function fCreate( _
ByVal iIconWidth As Long, _
ByVal iIconHeight As Long, _
ByVal iColorDepth As eImlColorDepth) _
As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/25/05
    ' Purpose   : Create the imagelist if it is not created already.
    '---------------------------------------------------------------------------------------
    If mhIml = ZeroL Then
        miColorDepth = iColorDepth
        mhIml = ImageList_Create(iIconWidth, iIconHeight, ILC_MASK Or miColorDepth, 4&, 4&)
        fCreate = CBool(mhIml)
    End If
End Function

Private Function pGetDefBackColor(ByVal hBmp As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/25/05
    ' Purpose   : Return the top-left pixel of the bitmap.
    '---------------------------------------------------------------------------------------
    pGetDefBackColor = CLR_NONE
    If hBmp Then
        Dim lhDCDisp      As Long, lhDc As Long, hBmpOld As Long
        lhDCDisp = CreateDisplayDC()
        If lhDCDisp Then
            lhDc = CreateCompatibleDC(lhDCDisp)
            If lhDc Then
                hBmpOld = SelectObject(lhDc, hBmp)
                If hBmpOld Then
                    pGetDefBackColor = GetPixel(lhDc, ZeroL, ZeroL)
                    SelectObject lhDc, hBmpOld
                End If
                DeleteDC lhDc
            End If
            DeleteDC lhDCDisp
        End If
    End If
End Function

Private Function pAdd( _
ByVal hImage As Long, _
ByVal iType As eImlImageType, _
ByVal iBackColor As OLE_COLOR) _
As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/25/05
    ' Purpose   : Add a bitmap or icon to the imagelist.
    '---------------------------------------------------------------------------------------
    If mhIml <> ZeroL And hImage <> ZeroL Then
        If iType = imlBitmap Then
            If iBackColor = NegOneL Then iBackColor = pGetDefBackColor(hImage) Else iBackColor = TranslateColor(iBackColor)
            pAdd = CBool(ImageList_AddMasked(mhIml, hImage, iBackColor) > NegOneL)
        ElseIf iType = imlIcon Or iType = imlCursor Then
            pAdd = CBool(ImageList_AddIcon(mhIml, hImage) > NegOneL)
        End If
        If pAdd Then RaiseEvent Changed
    End If
End Function

Public Property Get hIml() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/25/05
    ' Purpose   : Return the imagelist.
    '---------------------------------------------------------------------------------------
    hIml = mhIml
End Property

Public Property Get IconWidth() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/25/05
    ' Purpose   : Return the width of the icons in the imagelist.
    '---------------------------------------------------------------------------------------
    If mhIml Then ImageList_GetIconSize mhIml, IconWidth, ZeroL
End Property

Public Property Get IconHeight() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/25/05
    ' Purpose   : Return the height of the icons in the imagelist.
    '---------------------------------------------------------------------------------------
    If mhIml Then ImageList_GetIconSize mhIml, ZeroL, IconHeight
End Property

Public Property Get ColorDepth() As eImlColorDepth
    '---------------------------------------------------------------------------------------
    ' Date      : 2/25/05
    ' Purpose   : Return the color depth.
    '---------------------------------------------------------------------------------------
    ColorDepth = miColorDepth
End Property

Public Function AddFromFile( _
ByRef sFilename As String, _
ByVal iType As eImlImageType, _
Optional ByVal bMapSysColors As Boolean = False, _
Optional ByVal iBackColor As OLE_COLOR = NegOneL) _
As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/25/05
    ' Purpose   : Add an icon or set of icons from a file.
    '---------------------------------------------------------------------------------------
    iType = iType And 2&
            
    Dim hImage      As Long

    If mhIml Then
        Dim lsAnsi As String
        lsAnsi = StrConv(sFilename & vbNullChar, vbFromUnicode)
        hImage = LoadImage(App.hInstance, ByVal StrPtr(lsAnsi), iType, 0, 0, LR_LOADFROMFILE Or (-bMapSysColors * LR_LOADMAP3DCOLORS))
        If hImage Then
            AddFromFile = pAdd(hImage, iType, iBackColor)
            If iType = imlIcon Then
                DestroyIcon hImage
            ElseIf iType = imlCursor Then
                DestroyCursor hImage
            Else
                DeleteObject hImage
            End If
        End If
    End If
End Function

Public Function AddFromResource( _
ByVal iId As Long, _
ByVal iType As eImlImageType, _
Optional ByVal hInst As Long, _
Optional ByVal bMapSysColors As Boolean = False, _
Optional ByVal iBackColor As OLE_COLOR = NegOneL) _
As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/25/05
    ' Purpose   : Add a icon or set of icons from a resource.
    '             This will only work with a resource in a compiled exe.
    '---------------------------------------------------------------------------------------
    iType = iType And 2&
    
    Dim lhImage      As Long
    
    If mhIml Then
        lhImage = LoadImage(hInst, ByVal iId, iType, _
        -CBool(iType <> imlBitmap) * IconWidth, _
        -CBool(iType <> imlBitmap) * IconHeight, _
        (-bMapSysColors * LR_LOADMAP3DCOLORS) _
        Or _
        (-CBool(hInst = ZeroL) * LR_SHARED))
        
        If lhImage Then
            AddFromResource = pAdd(lhImage, iType, iBackColor)
            
            If hInst Then
                If iType = imlIcon Then
                    DestroyIcon lhImage
                ElseIf iType = imlCursor Then
                    DestroyCursor lhImage
                Else
                    DeleteObject lhImage
                End If
            End If
        End If
    End If
End Function

Public Function AddFromDc( _
ByVal hdc As Long, _
ByVal iWidth As Long, _
ByVal iHeight As Long, _
Optional ByVal iLeft As Long, _
Optional ByVal iTop As Long, _
Optional ByVal iBackColor As OLE_COLOR = NegOneL) _
As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/25/05
    ' Purpose   : Add a icon or set of icons from a device context.
    '---------------------------------------------------------------------------------------

    Dim lhDc          As Long
    Dim lhBmp         As Long
    Dim lhBmpOld      As Long
    
    If mhIml Then
        lhDc = CreateCompatibleDC(hdc)
        If lhDc Then
            lhBmp = CreateCompatibleBitmap(hdc, iWidth, iHeight)
            If lhBmp Then
                lhBmpOld = SelectObject(lhDc, lhBmp)
                If lhBmpOld Then
                    BitBlt lhDc, ZeroL, ZeroL, iWidth, iHeight, hdc, iLeft, iTop, vbSrcCopy
                    SelectObject lhDc, lhBmpOld
                    AddFromDc = pAdd(lhBmp, imlBitmap, IIf(iBackColor = NegOneL, GetPixel(hdc, iLeft, iTop), iBackColor))
                End If
                DeleteObject lhBmp
            End If
            DeleteDC lhDc
        End If
    End If
End Function

Public Function AddFromHandle(ByVal hImage As Long, ByVal iType As eImlImageType, Optional ByVal iBackColor As OLE_COLOR = NegOneL) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/25/05
    ' Purpose   : Add a icon or set of icons from a hImage.
    '---------------------------------------------------------------------------------------
    AddFromHandle = pAdd(hImage, iType And 2&, iBackColor)
End Function

Public Property Get IconCount() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/25/05
    ' Purpose   : Return the number of icons in the imagelist.
    '---------------------------------------------------------------------------------------
    If mhIml Then
        IconCount = ImageList_GetImageCount(mhIml)
    End If
End Property

'Public Sub Remove(ByVal iIndex As Long)
'    If mhIml Then
'        If ImageList_Remove(mhIml, iIndex) Then RaiseEvent Changed
'    End If
'End Sub
'
'Public Sub Clear()
'    If mhIml Then
'        ImageList_SetImageCount mhIml, ZeroL
'        RaiseEvent Changed
'    End If
'End Sub
'
'Public Property Get PictureStrip( _
'      Optional ByVal iStartIndex As Long = NegOneL, _
'      Optional ByVal iEndIndex As Long = NegOneL, _
'      Optional ByVal iBackColor As OLE_COLOR = vbButtonFace, _
'      Optional ByVal hCompatibleDc As Long) _
'            As IPictureDisp
'    Dim lbCreated As Boolean
'
'    If hCompatibleDc = ZeroL Then
'        lbCreated = True
'        hCompatibleDc = GetDC(ZeroL)
'    End If
'
'    Dim lhDc As Long
'    Dim hBrush As Long
'    Dim liSizeX As Long
'    Dim lhBmp As Long
'    Dim lhBmpOld As Long
'    Dim lhBrush As Long
'    Dim tR as RECT
'
'    If mhIml Then
'        If iStartIndex < ZeroL Then iStartIndex = ZeroL
'        If iStartIndex > ImageList_GetImageCount(mhIml) - OneL Then iStartIndex = ImageList_GetImageCount(mhIml) - OneL
'        If iEndIndex < ZeroL Then iEndIndex = ImageList_GetImageCount(mhIml) - OneL
'        If iEndIndex < iStartIndex Then iEndIndex = iStartIndex
'        If iEndIndex > ImageList_GetImageCount(mhIml) - OneL Then iEndIndex = ImageList_GetImageCount(mhIml) - OneL
'
'        lhDc = CreateCompatibleDC(hCompatibleDc)
'        If lhDc Then
'            liSizeX = (iEndIndex - iStartIndex + 1&) * IconWidth
'            lhBmp = CreateCompatibleBitmap(hCompatibleDc, liSizeX, IconHeight)
'            If lhBmp Then
'                lhBmpOld = SelectObject(lhDc, lhBmp)
'                If lhBmpOld Then
'                    lhBrush = gdimgr_createsolidbrush(TranslateColor(iBackColor))
'                    If lhBrush Then
'                        tR.Bottom = IconHeight
'                        tR.Right = liSizeX
'                        FillRect lhDc, tR, lhBrush
'                        gdimgr_deleteobject lhBrush
'                    End If
'
'                    For iEndIndex = iStartIndex To iEndIndex
'                       ImageList_Draw mhIml, iEndIndex, lhDc, (iEndIndex - iStartIndex) * IconWidth, ZeroL, ILD_TRANSPARENT
'                    Next
'                    SelectObject lhDc, lhBmpOld
'
'                    Set PictureStrip = BitmapToPicture(lhBmp)
'                Else
'                    gdimgr_deleteobject lhBmp
'                End If
'            End If
'            gdimgr_deletedc lhDc
'        End If
'    End If
'
'    If lbCreated Then ReleaseDC ZeroL, hCompatibleDc
'End Property

Public Sub Draw( _
ByVal iIndex As Long, _
ByVal hdc As Long, _
ByVal X As Long, _
ByVal Y As Long, _
Optional ByVal iStyle As eImlDrawStyle, _
Optional ByVal iCutDitherColor As OLE_COLOR = NegOneL)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/25/05
    ' Purpose   : Draw the given icon into the dc.
    '---------------------------------------------------------------------------------------
    ImageListDraw mhIml, iIndex, hdc, X, Y, iStyle, iCutDitherColor
End Sub

'Public Property Get ExtractIcon( _
'            ByVal iIndex As Long) _
'                As StdPicture
''---------------------------------------------------------------------------------------
'' Date      : 2/25/05
'' Purpose   : Extract an icon object from the given index.
''---------------------------------------------------------------------------------------
'    Set ExtractIcon = mComCtlShared.ExtractIcon(mhIml, iIndex)
'End Property

