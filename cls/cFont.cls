VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'cFont.cls                  1/17/04
'
'           PURPOSE:
'               Wrap a LOGFONT structure and expose its members.
'               Convert LOGFONT <--> StdFont
'
'               Browse For Font.
'
'               Font Source can be any system, ambient or custom font.
'               When set to a system or ambient font, you will need to call
'               OnSettingChange or OnAmbientFontChanged in the WM_SETTINGCHANGE
'               handler or the usercontrol's AmbientChanged event, respectively,
'               for the font to update itself.
'
'               Persistable.
'
'==================================================================================================

Option Explicit

Public Enum eFontDialog
    dlgFontScreenFonts = CF_SCREENFONTS
    dlgFontPrinterFonts = CF_PRINTERFONTS
    dlgFontBoth = CF_BOTH
    dlgFontEffects = CF_EFFECTS
    dlgFontNoVectorFonts = CF_NOVECTORFONTS
    dlgFontNoSimulations = CF_NOSIMULATIONS
    dlgFontFixedPitchOnly = CF_FIXEDPITCHONLY
    dlgFontWysiwyg = CF_WYSIWYG ' Must also have ScreenFonts And PrinterFonts
    dlgFontForceExist = CF_FORCEFONTEXIST
    dlgFontScalableOnly = CF_SCALABLEONLY
    dlgFontTTOnly = CF_TTONLY
    dlgFontNoFaceSel = CF_NOFACESEL
    dlgFontNoStyleSel = CF_NOSTYLESEL
    dlgFontNoSizeSel = CF_NOSIZESEL
    
    dlgFontSelectScript = CF_SELECTSCRIPT
    dlgFontNoScriptSel = CF_NOSCRIPTSEL
    dlgFontNoVertFonts = CF_NOVERTFONTS
    
    dlgFontEnableHook = CF_ENABLEHOOK
End Enum


Public Enum eFontWeight
    fntWeightDefault = 0
    fntWeightThin = 100
    fntWeightExtraLight = 200
    fntWeightLight = 300
    fntWeightNormal = 400
    fntWeightMedium = 500
    fntWeightSemibold = 600
    fntWeightBold = 700
    fntWeightExtraBold = 800
    fntWeightHeavy = 900
End Enum

Public Enum eFontCharset
    fntCharsetANSI = 0
    fntCharsetDefault = 1
    fntCharsetSymbol = 2
    fntCharsetShiftJIS = 128
    fntCharsetOEM = 255
End Enum

Public Enum eFontOutputPrecision
    fntOutPrecCharacter = 2
    fntOutPrecDefault = 0
    fntOutPrecDevice = 5
    fntOutPrecOutline = 8
    fntOutPrecRaster = 6
    fntOutPrecString = 1
    fntOutPrecStroke = 3
    fntOutPrecTTOnly = 7
    fntOutPrecTT = 4
End Enum

Public Enum eFontClipPrecision
    fntClipPrecDefault = 0
    'fntClipPrecCharacter = 1
    'fntClipPrecStroke = 2
    'fntClipPrecMask = &HF&
    fntClipPrecEmbedded = 128
    fntClipPrecLHAngles = 16
End Enum

Public Enum eFontQuality
    fntQualityDefault = 0
    fntQualityDraft = 1
    fntQualityProof = 2
End Enum

Public Enum eFontPitchAndFamily
    fntPitchDefault = 0
    fntPitchFixed = 1
    fntPitchVariable = 2
    fntFamilyDecorative = 80
    fntFamilyDefault = 0
    fntFamilyModern = 48
    fntFamilyRoman = 16
    fntFamilyScript = 64
    fntFamilySwiss = 32
End Enum

Public Enum eFontSource
    fntSourceCustom = 0
    fntSourceAmbient
    fntSourceSysMenu
    fntSourceSysMessage
    fntSourceSysStatus
    fntSourceSysCaption
    fntSourceSysSmallCaption
End Enum

Public Enum eFontDataType
    fntDataTypeDefault
    fntDataTypeStdFont
    fntDataTypeCFont
End Enum

Public Event Changed()

Private Enum eSystemFonts
    sysFontMenu = 2
    sysFontMessage
    sysFontStatus
    sysFontCaption
    sysFontSmallCaption
End Enum

Private Const DEF_FaceName           As String = "MS Sans Serif"
Private Const DEF_Height             As Long = ZeroL
Private Const DEF_Width              As Long = ZeroL
Private Const DEF_Escapement         As Long = ZeroL
Private Const DEF_Orientation        As Long = ZeroL
Private Const DEF_Weight             As Long = ZeroL
Private Const DEF_Italic             As Byte = ZeroY
Private Const DEF_Underline          As Byte = ZeroY
Private Const DEF_Strikeout          As Byte = ZeroY
Private Const DEF_Charset            As Byte = ZeroY
Private Const DEF_OutPrecision       As Byte = ZeroY
Private Const DEF_ClipPrecision      As Byte = ZeroY
Private Const DEF_Quality            As Byte = ZeroY
Private Const DEF_PitchAndFamily     As Byte = ZeroY
Private Const DEF_Source             As Long = fntSourceAmbient

Private Const PROP_FaceName         As String = "Name"
Private Const PROP_Height           As String = "Height"
Private Const PROP_Width            As String = "Width"
Private Const PROP_Escapement       As String = "Esc"
Private Const PROP_Orientation      As String = "Orient"
Private Const PROP_Weight           As String = "Weight"
Private Const PROP_Italic           As String = "Italic"
Private Const PROP_Underline        As String = "Underline"
Private Const PROP_Strikeout        As String = "Strikeout"
Private Const PROP_Charset          As String = "Charset"
Private Const PROP_OutPrecision     As String = "OutPrec"
Private Const PROP_ClipPrecision    As String = "ClipPrec"
Private Const PROP_Quality          As String = "Quality"
Private Const PROP_PitchAndFamily   As String = "PitchFam"
Private Const PROP_Source           As String = "Source"

Private Const SPI_SETNONCLIENTMETRICS As Long = 42

Private mtLogFont                   As LOGFONT
Private miFontSource                As eFontSource

Private Property Let pFaceName(ByRef sName As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Set the byte array used to store the face name from a unicode string.
    '---------------------------------------------------------------------------------------
    
    Dim ls        As String
    Dim iLen      As Long
    
    ls = StrConv(sName, vbFromUnicode)
    
    iLen = LenB(ls)
    If iLen > LF_FACESIZE Then iLen = LF_FACESIZE
    
    If iLen > ZeroL Then CopyMemory mtLogFont.lfFaceName(0), ByVal StrPtr(ls), iLen
    
    If iLen < LF_FACESIZE _
        Then ZeroMemory mtLogFont.lfFaceName(iLen), (LF_FACESIZE - iLen) _
    Else mtLogFont.lfFaceName(LF_FACESIZE - 1&) = ZeroY
    
End Property

Private Property Get pFaceName() As String
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Return a unicode string from a byte array.
    '---------------------------------------------------------------------------------------
    pFaceName = StrConv(mtLogFont.lfFaceName, vbUnicode)
    Dim i      As Long
    i = InStr(1&, pFaceName, vbNullChar)
    If i Then pFaceName = Left$(pFaceName, i - 1&)
End Property

Private Sub pPutStdFont(ByVal oFont As StdFont)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : "Put" data into our structure from a StdFont object.
    '---------------------------------------------------------------------------------------
    On Error Resume Next
    miFontSource = fntSourceCustom
    pFaceName = oFont.Name
    With mtLogFont
        .lfHeight = -MulDiv(oFont.Size, 1440& / Screen.TwipsPerPixelY, 72&)
        .lfWeight = IIf(oFont.Bold, fntWeightBold, fntWeightNormal)
        .lfItalic = Abs(oFont.Italic)
        .lfUnderline = Abs(oFont.Underline)
        .lfStrikeOut = Abs(oFont.Strikethrough)
        .lfCharSet = oFont.Charset And &HFF
        .lfEscapement = ZeroL
        .lfOrientation = ZeroL
        .lfWidth = ZeroL
        .lfOutPrecision = 0
        .lfClipPrecision = 0
        .lfQuality = 0
        .lfPitchAndFamily = 0
    End With
    On Error GoTo 0
End Sub

Private Sub pGetStdFont(ByVal oFont As StdFont)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : "Get" data from our structure into a StdFont object.
    '---------------------------------------------------------------------------------------
    On Error Resume Next
    With oFont
        .Name = pFaceName()
        If mtLogFont.lfHeight Then
            .Size = MulDiv(72&, Abs(mtLogFont.lfHeight), (1440& / Screen.TwipsPerPixelY))
        End If
        .Charset = mtLogFont.lfCharSet
        .Italic = CBool(mtLogFont.lfItalic)
        .Underline = CBool(mtLogFont.lfUnderline)
        .Strikethrough = CBool(mtLogFont.lfStrikeOut)
        .Bold = CBool(mtLogFont.lfWeight > fntWeightNormal)
    End With
    On Error GoTo 0
End Sub

Private Sub pPutFont(ByVal oFont As cFont)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : "Put" data into our structure from another cFont object.
    '---------------------------------------------------------------------------------------
    oFont.fGetLogFont mtLogFont, miFontSource
End Sub

Private Sub pGetFont(ByVal oFont As cFont)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : "Get" data from our structure into another cFont object.
    '---------------------------------------------------------------------------------------
    oFont.fPutLogFont mtLogFont, miFontSource
End Sub

Private Sub pGetSystemFont(ByVal iFontType As eSystemFonts, ByRef tLogFont As LOGFONT)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Get the font attributes for a given system font.
    '---------------------------------------------------------------------------------------
    Dim lR         As Long
    Dim liLen      As Long
    Dim ltNCM      As NONCLIENTMETRICS
    
    liLen = Len(ltNCM)
    
    ltNCM.cbSize = liLen - 16&
    lR = SystemParametersInfo(SPI_GETNONCLIENTMETRICS, liLen, ltNCM, ZeroL)
    
    If iFontType = sysFontMessage Then
        LSet tLogFont = ltNCM.lfMessageFont
    ElseIf iFontType = sysFontCaption Then
        LSet tLogFont = ltNCM.lfCaptionFont
    ElseIf iFontType = sysFontMenu Then
        LSet tLogFont = ltNCM.lfMenuFont
    ElseIf iFontType = sysFontSmallCaption Then
        LSet tLogFont = ltNCM.lfSmCaptionFont
    ElseIf iFontType = sysFontStatus Then
        LSet tLogFont = ltNCM.lfStatusFont
    End If
    
End Sub

Friend Sub fPutLogFont(ByRef tLogFont As LOGFONT, ByVal iFontSource As eFontSource)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : "Put" data into our structure from a LOGFONT structure.
    '---------------------------------------------------------------------------------------
    LSet mtLogFont = tLogFont
    miFontSource = iFontSource
    pRaiseEvent
End Sub

Friend Sub fGetLogFont(ByRef tLogFont As LOGFONT, ByRef iFontSource As eFontSource)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : "Get" data from our structure into a LOGFONT structure.
    '---------------------------------------------------------------------------------------
    LSet tLogFont = mtLogFont
    iFontSource = miFontSource
End Sub

Friend Sub fPutLogFontLong(ByVal lptr As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Put data into our structure from a given memory location.
    '---------------------------------------------------------------------------------------
    CopyMemory mtLogFont, ByVal lptr, Len(mtLogFont)
    miFontSource = fntSourceCustom
    pRaiseEvent
End Sub

Friend Sub fGetLogFontLong(ByVal lptr As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Get data from our structure into a given memory location.
    '---------------------------------------------------------------------------------------
    CopyMemory ByVal lptr, mtLogFont, Len(mtLogFont)
End Sub


Public Property Get Height() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Return the height of the font stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    Height = mtLogFont.lfHeight
End Property
Public Property Let Height(ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Set the height of the font stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    If CBool(iNew Xor mtLogFont.lfHeight) Then
        mtLogFont.lfHeight = iNew
        miFontSource = fntSourceCustom
        pRaiseEvent
    End If
End Property

Public Property Get Width() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Return the width of the font stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    Width = mtLogFont.lfWidth
End Property
Public Property Let Width(ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Set the width of the font stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    If CBool(iNew Xor mtLogFont.lfWidth) Then
        mtLogFont.lfWidth = iNew
        miFontSource = fntSourceCustom
        pRaiseEvent
    End If
End Property

Public Property Get Escapement() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Return the escapement stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    Escapement = mtLogFont.lfEscapement
End Property
Public Property Let Escapement(ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Set the escapement stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    If CBool(iNew Xor mtLogFont.lfEscapement) Then
        mtLogFont.lfEscapement = iNew
        miFontSource = fntSourceCustom
        pRaiseEvent
    End If
End Property

Public Property Get Orientation() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Return the orientation stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    Orientation = mtLogFont.lfOrientation
End Property
Public Property Let Orientation(ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Set the orientation stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    If CBool(iNew Xor mtLogFont.lfOrientation) Then
        mtLogFont.lfOrientation = iNew
        miFontSource = fntSourceCustom
        pRaiseEvent
    End If
End Property

Public Property Get Weight() As eFontWeight
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Return the weight stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    Weight = mtLogFont.lfWeight
End Property
Public Property Let Weight(ByVal iNew As eFontWeight)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Set the weight stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    If CBool(iNew Xor mtLogFont.lfWeight) Then
        mtLogFont.lfWeight = iNew
        miFontSource = fntSourceCustom
        pRaiseEvent
    End If
End Property

Public Property Get Italic() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Return the italic indicator stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    Italic = CBool(mtLogFont.lfItalic)
End Property
Public Property Let Italic(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Set the italic indicator stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    If CBool(bNew Xor CBool(mtLogFont.lfItalic)) Then
        mtLogFont.lfItalic = Abs(bNew)
        miFontSource = fntSourceCustom
        pRaiseEvent
    End If
End Property

Public Property Get Underline() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Return the underline indicator stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    Underline = CBool(mtLogFont.lfUnderline)
End Property
Public Property Let Underline(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Set the underline indicator stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    If CBool(bNew Xor CBool(mtLogFont.lfUnderline)) Then
        mtLogFont.lfUnderline = Abs(bNew)
        miFontSource = fntSourceCustom
        pRaiseEvent
    End If
End Property

Public Property Get Strikeout() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Return the strikeout indicator stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    Strikeout = mtLogFont.lfStrikeOut
End Property
Public Property Let Strikeout(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Set the strikeout indicator stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    If CBool(bNew Xor CBool(mtLogFont.lfStrikeOut)) Then
        mtLogFont.lfStrikeOut = Abs(bNew)
        miFontSource = fntSourceCustom
        pRaiseEvent
    End If
End Property

Public Property Get Charset() As eFontCharset
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Return the charset stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    Charset = mtLogFont.lfCharSet
End Property
Public Property Let Charset(ByVal iNew As eFontCharset)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Set the charset stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    If CBool(iNew Xor CLng(mtLogFont.lfCharSet)) Then
        mtLogFont.lfCharSet = iNew And &HFF
        miFontSource = fntSourceCustom
        pRaiseEvent
    End If
End Property

Public Property Get OutPrecision() As eFontOutputPrecision
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Return the output precision stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    OutPrecision = mtLogFont.lfOutPrecision
End Property
Public Property Let OutPrecision(ByVal iNew As eFontOutputPrecision)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Set the output precision stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    If CBool(iNew Xor CLng(mtLogFont.lfOutPrecision)) Then
        mtLogFont.lfOutPrecision = iNew And &HFF
        miFontSource = fntSourceCustom
        pRaiseEvent
    End If
End Property

Public Property Get ClipPrecision() As eFontClipPrecision
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Return the clipping precision stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    ClipPrecision = mtLogFont.lfClipPrecision
End Property
Public Property Let ClipPrecision(ByVal iNew As eFontClipPrecision)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Set the clipping precision stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    If CBool(iNew Xor CLng(mtLogFont.lfClipPrecision)) Then
        mtLogFont.lfClipPrecision = iNew And &HFF
        miFontSource = fntSourceCustom
        pRaiseEvent
    End If
End Property

Public Property Get Quality() As eFontQuality
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Return the quality stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    Quality = mtLogFont.lfQuality
End Property
Public Property Let Quality(ByVal iNew As eFontQuality)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Set the quality stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    If CBool(iNew Xor CLng(mtLogFont.lfQuality)) Then
        mtLogFont.lfQuality = iNew And &HFF
        miFontSource = fntSourceCustom
        pRaiseEvent
    End If
End Property

Public Property Get PitchAndFamily() As eFontPitchAndFamily
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Return the pitch and family stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    PitchAndFamily = mtLogFont.lfPitchAndFamily
End Property
Public Property Let PitchAndFamily(ByVal iNew As eFontPitchAndFamily)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Set the pitch and family stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    If CBool(iNew Xor CLng(mtLogFont.lfPitchAndFamily)) Then
        mtLogFont.lfPitchAndFamily = iNew And &HFF
        miFontSource = fntSourceCustom
        pRaiseEvent
    End If
End Property

Public Property Get FaceName() As String
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Return the face name stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    FaceName = pFaceName()
End Property

Public Property Let FaceName(ByRef sNew As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Set the face name stored by the LOGFONT structure.
    '---------------------------------------------------------------------------------------
    pFaceName = sNew
    miFontSource = fntSourceCustom
    pRaiseEvent
End Property

Public Property Get Source() As eFontSource
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Return the font source.
    '---------------------------------------------------------------------------------------
    Source = miFontSource
End Property

Public Property Let Source(ByVal iNew As eFontSource)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Set the font source. If set to a system font, the values will be retrieved.
    '             If set to the ambient font, no values will be retreived until OnAmbientFontChanged is called.
    '---------------------------------------------------------------------------------------
    If iNew >= fntSourceCustom And iNew <= fntSourceSysSmallCaption Then
        If iNew > fntSourceAmbient Then
            pGetSystemFont iNew, mtLogFont
            miFontSource = iNew
            pRaiseEvent
        Else
            If iNew Xor miFontSource Then
                miFontSource = iNew
                pRaiseEvent
            End If
        End If
    Else
        miFontSource = fntSourceCustom
    End If
End Property

Public Sub OnAmbientFontChanged(ByVal oFont As Object)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Update the font if we are using the ambient font as the source.
    '---------------------------------------------------------------------------------------
    If miFontSource = fntSourceAmbient Then
        If Not oFont Is Nothing Then
            If TypeOf oFont Is cFont Then
                pPutFont oFont
            ElseIf TypeOf oFont Is StdFont Then
                pPutStdFont oFont
            End If
            miFontSource = fntSourceAmbient
            pRaiseEvent
        End If
    End If
End Sub

Public Sub OnSettingChange(Optional ByVal wParam As Long = SPI_SETNONCLIENTMETRICS)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Update the font if we are using a system font as the source.
    '---------------------------------------------------------------------------------------
    If wParam = SPI_SETNONCLIENTMETRICS Then
        If miFontSource >= fntSourceSysMenu And miFontSource <= fntSourceSysSmallCaption Then
            pGetSystemFont miFontSource, mtLogFont
            pRaiseEvent
        End If
    End If
End Sub

Friend Function GetHandle() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Return a new or cached handle for the current font.
    '---------------------------------------------------------------------------------------
    GetHandle = GdiMgr_CreateFontIndirect(mtLogFont)
End Function
Friend Function ReleaseHandle(ByVal hFont As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Release a reference to the font handle.
    '---------------------------------------------------------------------------------------
    ReleaseHandle = GdiMgr_DeleteFont(hFont) <> ZeroL
End Function

Public Property Get FontData(Optional ByVal iType As eFontDataType = fntDataTypeDefault) As Object
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Return data in another cFont object or a StdFont object.
    '---------------------------------------------------------------------------------------
    Select Case iType
    Case fntDataTypeStdFont
        Set FontData = New StdFont
        pGetStdFont FontData
    Case Else
        Set FontData = New cFont
        pGetFont FontData
    End Select
End Property

Public Property Set FontData(Optional ByVal iType As eFontDataType = fntDataTypeDefault, ByVal oNew As Object)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Set data from another cFont object or a StdFont object.
    '---------------------------------------------------------------------------------------
    Select Case iType
    Case fntDataTypeStdFont
        pPutStdFont oNew
    Case fntDataTypeCFont
        pPutFont oNew
    Case Else
        If TypeOf oNew Is cFont Then
            pPutFont oNew
        ElseIf TypeOf oNew Is StdFont Then
            pPutStdFont oNew
        Else
            gErr vbccTypeMismatch, "cFont"
        End If
    End Select
    pRaiseEvent
End Property

Public Function Browse( _
Optional ByVal hWndOwner As Long, _
Optional ByVal iFlags As eFontDialog = dlgFontScreenFonts, _
Optional ByVal hdc As Long, _
Optional ByVal iMinSize As Long = 6, _
Optional ByVal iMaxSize As Long = 72) _
As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Prompt the user for new font attributes.
    '---------------------------------------------------------------------------------------
    
    Const CF_INITTOLOGFONTSTRUCT As Long = &H40&
    Const CF_LIMITSIZE As Long = &H2000&
    
    Dim ltLogFont      As LOGFONT
   
    LSet ltLogFont = mtLogFont
   
    Dim ltChooseFont      As CHOOSEFONT
    With ltChooseFont
        .lStructSize = LenB(ltChooseFont)
        .Flags = (iFlags Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE)
        .hdc = hdc
        .hInstance = ZeroL
        .hWndOwner = RootParent(hWndOwner)
        .iPointSize = ZeroL
        .lCustData = ZeroL
        .lpfnHook = ZeroL
        .lpLogFont = VarPtr(ltLogFont)
        .lpszStyle = ZeroL
        .lpTemplateName = ZeroL
        .nFontType = ZeroL
        .nSizeMax = iMaxSize
        .nSizeMin = iMinSize
        .rgbColors = ZeroL
    End With
    
    If CHOOSEFONT(ltChooseFont) = OneL Then
        Browse = True
        LSet mtLogFont = ltLogFont
        miFontSource = fntSourceCustom
        pRaiseEvent
    End If
    
End Function

Public Property Get TextWidth(ByRef sText As String, Optional ByVal hdc As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Return the width of the given text in the current font.
    '---------------------------------------------------------------------------------------
    
    Dim lbCreated      As Boolean
    If hdc = ZeroL Then
        lbCreated = True
        hdc = GetDC(ZeroL)
    End If
    If hdc Then
        Dim lhFont      As Long
        lhFont = GdiMgr_CreateFontIndirect(mtLogFont)
        If lhFont Then
            Dim ltSize         As Size
            Dim lhFontOld      As Long
            lhFontOld = SelectObject(hdc, lhFont)
            If lhFontOld Then
                If GetTextExtentPoint32W(hdc, sText, Len(sText), ltSize) Then
                    TextWidth = ltSize.cx
                End If
                SelectObject hdc, lhFontOld
            End If
            GdiMgr_DeleteFont lhFont
        End If
        If lbCreated Then ReleaseDC ZeroL, hdc
    End If
    
End Property

Public Property Get TextHeight(ByRef sText As String, Optional ByVal hdc As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Return the height of the given text in the current font.
    '---------------------------------------------------------------------------------------
    Dim lbCreated      As Boolean
    If hdc = ZeroL Then
        lbCreated = True
        hdc = GetDC(ZeroL)
    End If
    If hdc Then
        Dim lhFont      As Long
        lhFont = GdiMgr_CreateFontIndirect(mtLogFont)
        If lhFont Then
            Dim lhFontOld      As Long
            lhFontOld = SelectObject(hdc, lhFont)
            If lhFontOld Then
                Dim ltSize      As Size
                If GetTextExtentPoint32W(hdc, sText, Len(sText), ltSize) Then
                    TextHeight = ltSize.cy
                End If
                SelectObject hdc, lhFontOld
            End If
            GdiMgr_DeleteFont lhFont
        End If
        If lbCreated Then ReleaseDC ZeroL, hdc
    End If
End Property

Friend Function fCompare(ByRef tLF As LOGFONT) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Compare the given LOGFONT with our LOGFONT structure.
    '---------------------------------------------------------------------------------------
    fCompare = MemCmp(VarPtr(tLF), VarPtr(mtLogFont), LenB(mtLogFont))
End Function

Public Function Compare(ByVal oFont As cFont) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Compare the given font to our properties.
    '---------------------------------------------------------------------------------------
    If miFontSource = fntSourceCustom _
        Then Compare = oFont.fCompare(mtLogFont) _
    Else Compare = (oFont.Source = miFontSource)
End Function

Private Sub Class_InitProperties()
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Initialize property values to their defaults.
    '---------------------------------------------------------------------------------------
    miFontSource = DEF_Source
    pFaceName = DEF_FaceName
    With mtLogFont
        .lfCharSet = DEF_Charset
        .lfClipPrecision = DEF_ClipPrecision
        .lfEscapement = DEF_Escapement
        .lfHeight = DEF_Height
        .lfItalic = DEF_Italic
        .lfOrientation = DEF_Orientation
        .lfOutPrecision = DEF_OutPrecision
        .lfPitchAndFamily = DEF_PitchAndFamily
        .lfQuality = DEF_Quality
        .lfStrikeOut = DEF_Strikeout
        .lfUnderline = DEF_Underline
        .lfWeight = DEF_Weight
        .lfWidth = DEF_Width
    End With
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Read properties from a previously saved instance.
    '---------------------------------------------------------------------------------------
    miFontSource = PropBag.ReadProperty(PROP_Source, DEF_Source)
    If miFontSource >= fntSourceSysMenu And miFontSource <= fntSourceSysSmallCaption Then
        pGetSystemFont miFontSource, mtLogFont
    ElseIf miFontSource <> fntSourceAmbient Then
        miFontSource = fntSourceCustom
        pFaceName = PropBag.ReadProperty(PROP_FaceName, DEF_FaceName)
        With mtLogFont
            .lfHeight = PropBag.ReadProperty(PROP_Height, DEF_Height)
            .lfWidth = PropBag.ReadProperty(PROP_Width, DEF_Width)
            .lfEscapement = PropBag.ReadProperty(PROP_Escapement, DEF_Escapement)
            .lfOrientation = PropBag.ReadProperty(PROP_Orientation, DEF_Orientation)
            .lfWeight = PropBag.ReadProperty(PROP_Weight, DEF_Weight)
            .lfItalic = PropBag.ReadProperty(PROP_Italic, DEF_Italic)
            .lfUnderline = PropBag.ReadProperty(PROP_Underline, DEF_Underline)
            .lfStrikeOut = PropBag.ReadProperty(PROP_Strikeout, DEF_Strikeout)
            .lfCharSet = PropBag.ReadProperty(PROP_Charset, DEF_Charset)
            .lfOutPrecision = PropBag.ReadProperty(PROP_OutPrecision, DEF_OutPrecision)
            .lfClipPrecision = PropBag.ReadProperty(PROP_ClipPrecision, DEF_ClipPrecision)
            .lfQuality = PropBag.ReadProperty(PROP_Quality, DEF_Quality)
            .lfPitchAndFamily = PropBag.ReadProperty(PROP_PitchAndFamily, DEF_PitchAndFamily)
        End With
    End If
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Store property values for future use.
    '---------------------------------------------------------------------------------------
    PropBag.WriteProperty PROP_Source, miFontSource, DEF_Source
    If miFontSource = fntSourceCustom Then
        PropBag.WriteProperty PROP_FaceName, pFaceName, DEF_FaceName
        With mtLogFont
            PropBag.WriteProperty PROP_Height, .lfHeight, DEF_Height
            PropBag.WriteProperty PROP_Width, .lfWidth, DEF_Width
            PropBag.WriteProperty PROP_Escapement, .lfEscapement, DEF_Escapement
            PropBag.WriteProperty PROP_Orientation, .lfOrientation, DEF_Orientation
            PropBag.WriteProperty PROP_Weight, .lfWeight, DEF_Weight
            PropBag.WriteProperty PROP_Italic, .lfItalic, DEF_Italic
            PropBag.WriteProperty PROP_Underline, .lfUnderline, DEF_Underline
            PropBag.WriteProperty PROP_Strikeout, .lfStrikeOut, DEF_Strikeout
            PropBag.WriteProperty PROP_Charset, .lfCharSet, DEF_Charset
            PropBag.WriteProperty PROP_OutPrecision, .lfOutPrecision, DEF_OutPrecision
            PropBag.WriteProperty PROP_ClipPrecision, .lfClipPrecision, DEF_ClipPrecision
            PropBag.WriteProperty PROP_Quality, .lfQuality, DEF_Quality
            PropBag.WriteProperty PROP_PitchAndFamily, .lfPitchAndFamily, DEF_PitchAndFamily
        End With
    End If
End Sub

Private Sub pRaiseEvent()
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/04
    ' Purpose   : Raise a changed event, but ignore any changed events triggered while
    '             processing the first event.
    '---------------------------------------------------------------------------------------
    Static bInHere As Boolean
    If Not bInHere Then
        bInHere = True
        RaiseEvent Changed
        bInHere = False
    End If
End Sub
