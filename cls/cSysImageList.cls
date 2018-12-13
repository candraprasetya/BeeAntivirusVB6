VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSysImageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'cSysImageList.cls              8/25/04
'
'           GENERAL PURPOSE:
'               Manage a system image list.
'
'           LINEAGE:
'               cSysImageList.cls from www.vbaccelerator.com
'
'==================================================================================================

Option Explicit

Implements cImageList

Private mhIml    As Long
Private miFlags  As Long

Private Const ClassName As String = "cSysImageList"
Private Const ErrorDescription = "Invalid call to a system image list."

Private Function cImageList_AddFromDc(ByVal hdc As Long, ByVal iWidth As Long, ByVal iHeight As Long, Optional ByVal iLeft As Long, Optional ByVal iTop As Long, Optional ByVal iBackColor As stdole.OLE_COLOR = -1&) As Boolean
    gErr vbccInvalidProcedureCall, ClassName, ErrorDescription
End Function

Private Function cImageList_AddFromFile(sFilename As String, ByVal iType As eImlImageType, Optional ByVal bMapSysColors As Boolean = False, Optional ByVal iBackColor As stdole.OLE_COLOR = -1&) As Boolean
    gErr vbccInvalidProcedureCall, ClassName, ErrorDescription
End Function

Private Function cImageList_AddFromHandle(ByVal hImage As Long, ByVal iType As eImlImageType, Optional ByVal iBackColor As stdole.OLE_COLOR = -1&) As Boolean
    gErr vbccInvalidProcedureCall, ClassName, ErrorDescription
End Function

Private Function cImageList_AddFromResource(ByVal iId As Long, ByVal iType As eImlImageType, Optional ByVal hInst As Long, Optional ByVal bMapSysColors As Boolean = False, Optional ByVal iBackColor As stdole.OLE_COLOR = -1&) As Boolean
    gErr vbccInvalidProcedureCall, ClassName, ErrorDescription
End Function

Private Property Get cImageList_ColorDepth() As eImlColorDepth
    cImageList_ColorDepth = ColorDepth()
End Property

Private Sub cImageList_Draw(ByVal iIndex As Long, ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal iStyle As eImlDrawStyle, Optional ByVal iCutDitherColor As stdole.OLE_COLOR = -1&)
    ImageListDraw mhIml, iIndex, hdc, X, Y, iStyle, iCutDitherColor
End Sub

Private Property Get cImageList_hIml() As Long
    cImageList_hIml = mhIml
End Property

Private Property Get cImageList_IconCount() As Long
    cImageList_IconCount = ImageList_GetImageCount(mhIml)
End Property

Private Property Get cImageList_IconHeight() As Long
    If mhIml Then ImageList_GetIconSize mhIml, ZeroL, cImageList_IconHeight
End Property

Private Property Get cImageList_IconWidth() As Long
    If mhIml Then ImageList_GetIconSize mhIml, cImageList_IconWidth, ZeroL
End Property



Friend Function fCreate(ByVal bLarge As Boolean) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/25/05
    ' Purpose   : Create the imagelist if it is not created already.
    '---------------------------------------------------------------------------------------
    Dim FileInfo      As SHFILEINFO
     
    If mhIml = ZeroL Then
        miFlags = SHGFI_SYSICONINDEX Or ((bLarge + OneL) * SHGFI_SMALLICON)
        
        Dim lsText      As String
        lsText = StrConv(".txt" & vbNullChar, vbFromUnicode)
        mhIml = SHGetFileInfo(ByVal StrPtr(lsText), FILE_ATTRIBUTE_NORMAL, FileInfo, LenB(FileInfo), miFlags Or SHGFI_USEFILEATTRIBUTES)
        
        'debug.assert FileInfo.hIcon = ZeroL
        
        fCreate = CBool(mhIml <> ZeroL And mhIml <> NegOneL)
        If Not fCreate Then mhIml = ZeroL
    End If
End Function

Public Function ItemIndex( _
ByRef sFile As String, _
Optional ByVal bForceLoadFromDisk As Boolean) _
As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/25/05
    ' Purpose   : Return the index of the icon for the given type of file.
    '             If the bForceLoadFromDisk flag is not set, then any filespec
    '             may be used and the extension will be used to find the icon.
    '             The icon is added to the imagelist only if it exists in the system cache.
    '             If the bForceLoadFromDisk flag is set, then the file must exist.
    '             The correct icon is guaranteed, but performance suffers.
    '---------------------------------------------------------------------------------------
    
    Dim FileInfo      As SHFILEINFO
    If mhIml Then
        Dim lsFile As String
        lsFile = StrConv(sFile & vbNullChar, vbFromUnicode)
        If SHGetFileInfo(ByVal StrPtr(lsFile), FILE_ATTRIBUTE_NORMAL, FileInfo, LenB(FileInfo), miFlags Or (-bForceLoadFromDisk * SHGFI_USEFILEATTRIBUTES)) Then
            ItemIndex = FileInfo.iIcon
        Else
            ItemIndex = NegOneL
        End If
        
        'debug.assert FileInfo.hIcon = ZeroL
    End If
End Function

Public Property Get hIml() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/25/05
    ' Purpose   : Return a stored handle to the system imagelist.
    '---------------------------------------------------------------------------------------
    hIml = mhIml
End Property

Public Property Get IconCount() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/25/05
    ' Purpose   : Return the number of icons which have been exist to the list.  Items are
    '             added by the system as they are accessed.
    '---------------------------------------------------------------------------------------
    IconCount = ImageList_GetImageCount(mhIml)
End Property

Public Property Get ColorDepth() As eImlColorDepth
    '---------------------------------------------------------------------------------------
    ' Date      : 2/25/05
    ' Purpose   : Return the system color depth.
    '---------------------------------------------------------------------------------------
    ColorDepth = SystemColorDepth
End Property

Public Property Get IconHeight() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/25/05
    ' Purpose   : Return the height of the icons in the imagelist.
    '---------------------------------------------------------------------------------------
    If mhIml Then ImageList_GetIconSize mhIml, ZeroL, IconHeight
End Property

Public Property Get IconWidth() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/25/05
    ' Purpose   : Return the Width of the icons in the imagelist.
    '---------------------------------------------------------------------------------------
    If mhIml Then ImageList_GetIconSize mhIml, IconWidth, ZeroL
End Property

Public Sub Draw(ByRef sFile As String, ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal iStyle As eImlDrawStyle, Optional ByVal iCutDitherColor As stdole.OLE_COLOR = -1&, Optional ByVal bForceLoadFromDisk As Boolean)
    cImageList_Draw ItemIndex(sFile, bForceLoadFromDisk), hdc, X, Y, iStyle, iCutDitherColor
End Sub

'Public Property Get ExtractIcon( _
'               ByRef sFile As String, _
'      Optional ByVal bForceLoadFromDisk As Boolean) _
'        As StdPicture
'    Set ExtractIcon = mComCtlShared.ExtractIcon(mhIml, ItemIndex(sFile, bForceLoadFromDisk))
'End Property

