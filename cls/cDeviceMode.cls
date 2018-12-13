VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDeviceMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'cDeviceMode.cls                    12/15/04
'
'           PURPOSE:
'               Wrap a DEVMODE structure.
'
'==================================================================================================

Option Explicit

Public Enum eDeviceModeQuality
    dmDraft = vbPRPQDraft
    dmLow = vbPRPQLow
    dmMedium = vbPRPQMedium
    dmHigh = vbPRPQHigh
End Enum

Public Enum eDeviceModeOrientation
    dmPortrait = 1
    dmLandscape
End Enum

Public Enum eDeviceModePaperSize
    dmPaperSize10x11 = 45
    dmPaperSize10x14 = 16
    dmPaperSize11x17 = 17
    dmPaperSize12x11 = 90
    dmPaperSize15x11 = 46
    dmPaperSize9x11 = 44
    dmPaperSizeA_Plus = 57
    dmPaperSizeA2 = 66
    dmPaperSizeA3 = 8                                 ' A3, 297 x 420 mm
    dmPaperSizeA3Extra = 63
    dmPaperSizeA3ExtraTransverse = 68
    dmPaperSizeA3Rotated = 76
    dmPaperSizeA3Transverse = 67
    dmPaperSizeA4 = 9                                 ' A4, 210 x 297 mm
    dmPaperSizeA4Extra = 53
    dmPaperSizeA4Plus = 60
    dmPaperSizeA4Rotated = 77
    dmPaperSizeA4Transverse = 55
    dmPaperSizeA4Small = 10                           ' A4 Small, 210 x 297 mm
    dmPaperSizeA5 = 11                                ' A5, 148 x 210 mm
    dmPaperSizeA5Extra = 64
    dmPaperSizeA5Rotated = 78
    dmPaperSizeA5Transverse = 61
    dmPaperSizeA6 = 70
    dmPaperSizeA6Rotated = 83
    dmPaperSizeBPlus = 58
    dmPaperSizeB4 = 12                                ' B4, 250 x 354 mm
    dmPaperSizeB4JisRotated = 79
    dmPaperSizeB5 = 13                                ' B5, 182 x 257 mm
    dmPaperSizeB5Extra = 65
    dmPaperSizeB5JisRotated = 80
    dmPaperSizeB5Transverse = 62
    dmPaperSizeB6Jis = 88
    dmPaperSizeB6JisRotated = 89
    dmPaperSizeCSheet = 24
    dmPaperSizeDblJapanesePostcard = 69
    dmPaperSizeDblJapanesePostcartRotated = 82
    dmPaperSizeDSheet = 25
    dmPaperSizeEnv10 = 20                             ' Envelope #10, 4 1/8 x 9 1/2 in.
    dmPaperSizeEnv11 = 21                             ' Envelope #11, 4 1/2 x 10 3/8 in.
    dmPaperSizeEnv12 = 22                             ' Envelope #12, 4 1/2 x 11 in.
    dmPaperSizeEnv14 = 23                             ' Envelope #14, 5 x 11 1/2 in.
    dmPaperSizeEnv9 = 19                              ' Envelope #9, 3 7/8 x 8 7/8 in.
    dmPaperSizeEnvB4 = 33                             ' Envelope B4, 250 x 353 mm
    dmPaperSizeEnvB5 = 34                             ' Envelope B5, 176 x 250 mm
    dmPaperSizeEnvB6 = 35                             ' Envelope B6, 176 x 125 mm
    dmPaperSizeEnvC3 = 29                             ' Envelope C3, 324 x 458 mm
    dmPaperSizeEnvC4 = 30                             ' Envelope C4, 229 x 324 mm
    dmPaperSizeEnvC5 = 28                             ' Envelope C5, 162 x 229 mm
    dmPaperSizeEnvC6 = 31                             ' Envelope C6, 114 x 162 mm
    dmPaperSizeEnvC65 = 32                            ' Envelope C65, 114 x 229 mm
    dmPaperSizeEnvDL = 27                             ' Envelope DL, 110 x 220 mm
    dmPaperSizeEnvInvite = 47
    dmPaperSizeEnvItaly = 36                          ' Envelope, 110 x 230 mm
    dmPaperSizeEnvMonarch = 37                        ' Envelope Monarch, 3 7/8 x 7 1/2 in.
    dmPaperSizeEnvPersonal = 38                       ' Envelope, 3 5/8 x 6 1/2 in.
    dmPaperSizeESheet = 26
    dmPaperSizeExecutive = 7                          ' Executive, 7 1/2 x 10 1/2 in.
    dmPaperSizeFanfoldLglGerman = 41                  ' German Legal Fanfold, 8 1/2 x 13 in.
    dmPaperSizeFanfoldStdGerman = 40                  ' German Standard Fanfold, 8 1/2 x 12 in.
    dmPaperSizeFanfoldUS = 39                         ' U.S. Standard Fanfold, 14 7/8 x 11 in.
    dmPaperSizeLetter = 1                             ' Letter, 8 1/2 x 11 in.
    dmPaperSizeFolio = 14                             ' Folio, 8 1/2 x 13 in.
    dmPaperSizeIso_B4 = 42
    dmPaperSizeJapanesePostcard = 43
    dmPaperSizeJapanesePostcardRotated = 81
    dmPaperSizeJEnvChou3 = 73
    dmPaperSizeJEnvChou3Rotated = 86
    dmPaperSizeJEnvChou4 = 74
    dmPaperSizeJEnvChou4_Rotated = 87
    dmPaperSizeJEnvKaku2 = 71
    dmPaperSizeJEnvKaku2_Rotated = 84
    dmPaperSizeJEnvKaku3 = 72
    dmPaperSizeJEnvKaku3_Rotated = 85
    dmPaperSizeJEnvYou4 = 91
    dmPaperSizeJEnvYou4_Rotated = 92
    dmPaperSizeLedger = 4                             ' Ledger, 17 x 11 in.
    dmPaperSizeLegal = 5                              ' Legal, 8 1/2 x 14 in.
    dmPaperSizeLegalExtra = 51
    dmPaperSizeLetterExtra = 50
    dmPaperSizeLetterExtraTransverse = 56
    dmPaperSizeLetterPlus = 59
    dmPaperSizeLetterRotated = 75
    dmPaperSizeLetterTransverse = 54
    dmPaperSizeLetterSmall = 2                        ' Letter Small, 8 1/2 x 11 in.
    dmPaperSizeNote = 18                              ' Note, 8 1/2 x 11 in.
    dmPaperSizeP16K = 93
    dmPaperSizeP16KRotated = 106
    dmPaperSizeP32K = 94
    dmPaperSizeP32KRotated = 107
    dmPaperSizeP32KBig = 95
    dmPaperSizeP32KBigRotated = 108
    dmPaperSizePEnv1 = 96
    dmPaperSizePEnv1Rotated = 109
    dmPaperSizePEnv10 = 105
    dmPaperSizePEnv10Rotated = 118
    dmPaperSizePEnv2 = 97
    dmPaperSizePEnv2Rotated = 110
    dmPaperSizePEnv3 = 98
    dmPaperSizePEnv3Rotated = 111
    dmPaperSizePEnv4 = 99
    dmPaperSizePEnv4Rotated = 112
    dmPaperSizePEnv5 = 100
    dmPaperSizePEnv5Rotated = 113
    dmPaperSizePEnv6 = 101
    dmPaperSizePEnv6Rotated = 114
    dmPaperSizePEnv7 = 102
    dmPaperSizePEnv7Rotated = 115
    dmPaperSizePEnv8 = 103
    dmPaperSizePEnv8Rotated = 116
    dmPaperSizePEnv9 = 104
    dmPaperSizePEnv9Rotated = 117
    dmPaperSizeQuarto = 15                            ' Quarto, 215 x 275 mm
    dmPaperSizeReserved48 = 48
    dmPaperSizeReserved49 = 49
    dmPaperSizeStatement = 6                          ' Statement, 5 1/2 x 8 1/2 in.
    dmPaperSizeTabloid = 3                            ' Tabloid, 11 x 17 in.
    dmPaperSizeTabloid_Extra = 52
    dmPaperSizeUserDefined = 256
End Enum

Public Enum eDeviceModeFields
    dmFieldOrientation = DM_ORIENTATION
    dmFieldPaperSize = DM_PAPERSIZE
    dmFieldPaperLength = DM_PAPERLENGTH
    dmFieldPaperWidth = DM_PAPERWIDTH
    dmFieldScale = DM_SCALE
    dmFieldCopies = DM_COPIES
    dmFieldDefaultSource = DM_DEFAULTSOURCE
    dmFieldPrintQuality = DM_PRINTQUALITY
    dmFieldColor = DM_COLOR
    dmFieldDuplex = DM_DUPLEX
    dmFieldYResolution = DM_YRESOLUTION
    dmFieldTTOption = DM_TTOPTION
    dmFieldCollate = DM_COLLATE
    dmFieldFormName = DM_FORMNAME
End Enum

Public Enum eDeviceModeDuplex
    dmDuplexNone = DMDUP_SIMPLEX
    dmDuplexHorizontal = DMDUP_HORIZONTAL
    dmDuplexVertical = DMDUP_VERTICAL
End Enum

Public Enum eDeviceModeTrueTypeOption
    dmTTBitmap = DMTT_BITMAP
    dmTTDownload = DMTT_DOWNLOAD
    dmTTDownloadOutline = DMTT_DOWNLOAD_OUTLINE
    dmTTSubstituteDevice = DMTT_SUBDEV
End Enum

Public Enum eDeviceModePaperSource
    dmPaperSourceAuto = DMBIN_AUTO
    dmPaperSourceCassette = DMBIN_CASSETTE
    dmPaperSourceEnvelope = DMBIN_ENVELOPE
    dmPaperSourceEnvelopeManual = DMBIN_ENVMANUAL
    dmPaperSourceFormSource = DMBIN_FORMSOURCE
    dmPaperSourceLargeCapacity = DMBIN_LARGECAPACITY
    dmPaperSourceLargeFmt = DMBIN_LARGEFMT
    dmPaperSourceLower = DMBIN_LOWER
    dmPaperSourceManual = DMBIN_MANUAL
    dmPaperSourceMiddle = DMBIN_MIDDLE
    dmPaperSourceOnlyOne = DMBIN_ONLYONE
    dmPaperSourceSmallFmt = DMBIN_SMALLFMT
    dmPaperSourceTractor = DMBIN_TRACTOR
    dmPaperSourceUpper = DMBIN_UPPER
    dmPaperSourceUser = DMBIN_USER
End Enum

Event Changed()

Private mtDevMode  As DEVMODE

Private Const DEF_Collate As Boolean = False
Private Const DEF_Color As Boolean = True
Private Const DEF_Copies As Integer = 0
Private Const DEF_DefaultSource As Integer = DMBIN_AUTO
Private Const DEF_DeviceName As String = vbNullString
Private Const DEF_DriverVersion As Integer = 0
Private Const DEF_Duplex As Integer = DMDUP_SIMPLEX
Private Const DEF_Fields As Long = 0
Private Const DEF_FormName As String = vbNullString
Private Const DEF_Orientation As Integer = dmPortrait
Private Const DEF_Scale As Integer = 0
Private Const DEF_PaperLength As Integer = 0
Private Const DEF_PaperSize As Integer = 0
Private Const DEF_PaperWidth As Integer = 0
Private Const DEF_PrintQuality As Integer = 0
Private Const DEF_TTOption As Integer = 0
Private Const DEF_YResolution As Integer = 0

Private Const PROP_Collate = "Collate"
Private Const PROP_Color = "Color"
Private Const PROP_Copies = "Copies"
Private Const PROP_DefaultSource = "DefSrc"
Private Const PROP_DeviceName = "DevName"
Private Const PROP_DriverVersion = "Ver"
Private Const PROP_Duplex = "Duplx"
Private Const PROP_Fields = "Fields"
Private Const PROP_FormName = "Form"
Private Const PROP_Orientation = "Orient"
Private Const PROP_Scale = "Scale"
Private Const PROP_PaperLength = "Length"
Private Const PROP_PaperSize = "Size"
Private Const PROP_PaperWidth = "Width"
Private Const PROP_PrintQuality = "Quality"
Private Const PROP_TTOption = "TTOption"
Private Const PROP_YResolution = "YRes"

Private Property Get pByteString(ByRef Y() As Byte, ByVal iByteLen As Long) As String
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Return a unicode string from an ANSI byte array.
    '---------------------------------------------------------------------------------------
    pByteString = StrConv(Y, vbUnicode)
    iByteLen = InStr(1, pByteString, vbNullChar)
    If iByteLen Then pByteString = Left$(pByteString, iByteLen - 1)
End Property
Private Property Let pByteString(ByRef Y() As Byte, ByVal iByteLen As Long, ByRef sNew As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Set an ANSI byte array from a unicode string.
    '---------------------------------------------------------------------------------------
    ZeroMemory Y(0), iByteLen
    If Len(sNew) < iByteLen Then iByteLen = Len(sNew)
    
    Dim lsAnsi      As String
    lsAnsi = StrConv(sNew & vbNullChar, vbFromUnicode)
    
    If iByteLen Then CopyMemory Y(0), ByVal StrPtr(lsAnsi), iByteLen
End Property

Private Property Get pInt(ByVal i As Long) As Integer
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Return the closest integer value from a long.
    '---------------------------------------------------------------------------------------
    If i > &H7FFF Then
        pInt = &H7FFF
    ElseIf i < &H8000 Then
        pInt = &H8000
    Else
        pInt = CInt(i)
    End If
End Property

Public Property Get DeviceName() As String
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Return the device name stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    DeviceName = pByteString(mtDevMode.dmDeviceName, CCHDEVICENAME)
End Property
Public Property Let DeviceName(ByRef sNew As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Set the device name stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    If Len(sNew) > CCHDEVICENAME Then sNew = Left$(sNew, CCHDEVICENAME)
    If StrComp(sNew, pByteString(mtDevMode.dmDeviceName, CCHDEVICENAME)) Then
        pByteString(mtDevMode.dmDeviceName, CCHDEVICENAME) = sNew
        RaiseEvent Changed
    End If
End Property

Public Property Get DriverVersion() As Integer
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Return the driver version stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    DriverVersion = mtDevMode.dmDriverVersion
End Property
Public Property Let DriverVersion(ByVal iNew As Integer)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Set the driver version stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    If iNew Xor mtDevMode.dmDriverVersion Then
        mtDevMode.dmDriverVersion = iNew
        RaiseEvent Changed
    End If
End Property

Public Property Get Fields() As eDeviceModeFields
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Return a bitmask indicating which fields contain valid data.
    '---------------------------------------------------------------------------------------
    Fields = mtDevMode.dmFields
End Property
Public Property Let Fields(ByVal iNew As eDeviceModeFields)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Set a bitmask indicating which fields contain valid data.
    '---------------------------------------------------------------------------------------
    If iNew Xor mtDevMode.dmFields Then
        mtDevMode.dmFields = iNew
        RaiseEvent Changed
    End If
End Property

Public Property Get Orientation() As eDeviceModeOrientation
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Return the orientation stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    Orientation = mtDevMode.dmOrientation
End Property
Public Property Let Orientation(ByVal iNew As eDeviceModeOrientation)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Set the orientation stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    If pInt(iNew) Xor mtDevMode.dmOrientation Then
        mtDevMode.dmOrientation = pInt(iNew)
        RaiseEvent Changed
    End If
End Property

Public Property Get PaperSize() As eDeviceModePaperSize
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Return the paper size by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    PaperSize = mtDevMode.dmPaperSize
End Property
Public Property Let PaperSize(ByVal iNew As eDeviceModePaperSize)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Set the paper size stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    If pInt(iNew) Xor mtDevMode.dmPaperSize Then
        mtDevMode.dmPaperSize = pInt(iNew)
        RaiseEvent Changed
    End If
End Property

Public Property Get PaperLength() As Integer
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Return the paper length stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    PaperLength = mtDevMode.dmPaperLength
End Property
Public Property Let PaperLength(ByVal iNew As Integer)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Set the paper length stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    If iNew Xor mtDevMode.dmPaperLength Then
        mtDevMode.dmPaperLength = iNew
        RaiseEvent Changed
    End If
End Property

Public Property Get PaperWidth() As Integer
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Return the paper width stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    PaperWidth = mtDevMode.dmPaperWidth
End Property
Public Property Let PaperWidth(ByVal iNew As Integer)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Set the paper width stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    If iNew Xor mtDevMode.dmPaperWidth Then
        mtDevMode.dmPaperWidth = iNew
        RaiseEvent Changed
    End If
End Property

Public Property Get OutputScaleFactor() As Integer
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Return the output scale factor stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    OutputScaleFactor = mtDevMode.dmScale
End Property
Public Property Let OutputScaleFactor(ByVal iNew As Integer)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Set the output scale factor stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    If iNew Xor mtDevMode.dmScale Then
        mtDevMode.dmScale = iNew
        RaiseEvent Changed
    End If
End Property

Public Property Get Copies() As Integer
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Return the number of copies stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    Copies = mtDevMode.dmCopies
End Property
Public Property Let Copies(ByVal iNew As Integer)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Set the number of copies stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    If iNew Xor mtDevMode.dmCopies Then
        mtDevMode.dmCopies = iNew
        RaiseEvent Changed
    End If
End Property

Public Property Get DefaultSource() As eDeviceModePaperSource
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Return the default source stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    DefaultSource = mtDevMode.dmDefaultSource
End Property
Public Property Let DefaultSource(ByVal iNew As eDeviceModePaperSource)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Set the default source stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    If pInt(iNew) Xor mtDevMode.dmDefaultSource Then
        mtDevMode.dmDefaultSource = pInt(iNew)
        RaiseEvent Changed
    End If
End Property

Public Property Get PrintQuality() As eDeviceModeQuality
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Return the print quality stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    PrintQuality = mtDevMode.dmPrintQuality
End Property
Public Property Let PrintQuality(ByVal iNew As eDeviceModeQuality)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Set the print quality stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    If pInt(iNew) Xor mtDevMode.dmPrintQuality Then
        mtDevMode.dmPrintQuality = pInt(iNew)
        RaiseEvent Changed
    End If
End Property

Public Property Get Color() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Return the color indicator stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    Color = (mtDevMode.dmColor = DMCOLOR_COLOR)
End Property
Public Property Let Color(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Set the color indicator stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    Dim iNew      As Integer: iNew = IIf(bNew, DMCOLOR_COLOR, DMCOLOR_MONOCHROME)
    If iNew Xor mtDevMode.dmColor Then
        mtDevMode.dmColor = iNew
        RaiseEvent Changed
    End If
End Property

Public Property Get Duplex() As eDeviceModeDuplex
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Return the duplex mode stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    Duplex = mtDevMode.dmDuplex
End Property
Public Property Let Duplex(ByVal iNew As eDeviceModeDuplex)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Set the duplex mode stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    If pInt(iNew) Xor mtDevMode.dmDuplex Then
        mtDevMode.dmDuplex = pInt(iNew)
        RaiseEvent Changed
    End If
End Property

Public Property Get YResolution() As Integer
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Return the y resolution stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    YResolution = mtDevMode.dmYResolution
End Property
Public Property Let YResolution(ByVal iNew As Integer)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Set the y resolution stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    If iNew Xor mtDevMode.dmYResolution Then
        mtDevMode.dmYResolution = iNew
        RaiseEvent Changed
    End If
End Property

Public Property Get TTOption() As eDeviceModeTrueTypeOption
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Return the truetype option stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    TTOption = mtDevMode.dmTTOption
End Property
Public Property Let TTOption(ByVal iNew As eDeviceModeTrueTypeOption)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Set the truetype option stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    If pInt(iNew) Xor mtDevMode.dmTTOption Then
        mtDevMode.dmTTOption = pInt(iNew)
        RaiseEvent Changed
    End If
End Property

Public Property Get Collate() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Return the collate indicator stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    Collate = mtDevMode.dmCollate
End Property
Public Property Let Collate(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Set the collate indicator stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    If CInt(Abs(bNew)) Xor mtDevMode.dmCollate Then
        mtDevMode.dmCollate = Abs(bNew)
        RaiseEvent Changed
    End If
End Property

Public Property Get FormName() As String
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Return the form name stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    FormName = pByteString(mtDevMode.dmFormName, CCHFORMNAME)
End Property
Public Property Let FormName(ByRef sNew As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Set the form name stored by the DEVMODE structure.
    '---------------------------------------------------------------------------------------
    If Len(sNew) > CCHFORMNAME Then sNew = Left$(sNew, CCHFORMNAME)
    If StrComp(sNew, pByteString(mtDevMode.dmFormName, CCHFORMNAME)) Then
        pByteString(mtDevMode.dmFormName, CCHFORMNAME) = sNew
        RaiseEvent Changed
    End If
End Property

Friend Property Get lpDevMode() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Return a pointer the the DEVMODE structure in memory.
    '---------------------------------------------------------------------------------------
    lpDevMode = VarPtr(mtDevMode)
End Property

Friend Property Get fSizeOf() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Return the length of the devmode structure in memory.
    '---------------------------------------------------------------------------------------
    fSizeOf = LenB(mtDevMode)
End Property

Friend Sub fChanged()
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Notify any clients that the data has changed.
    '---------------------------------------------------------------------------------------
    RaiseEvent Changed
End Sub

Private Sub Class_InitProperties()
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Initialize property values to their defaults.
    '---------------------------------------------------------------------------------------
    With mtDevMode
        .dmSize = Len(mtDevMode)
        .dmSpecVersion = DM_SPECVERSION
        
        .dmCollate = DEF_Collate
        .dmColor = DEF_Color
        .dmCopies = DEF_Copies
        .dmDefaultSource = DEF_DefaultSource
        pByteString(.dmDeviceName, CCHDEVICENAME) = DEF_DeviceName
        .dmDriverVersion = DEF_DriverVersion
        .dmDuplex = DEF_Duplex
        .dmFields = DEF_Fields
        pByteString(.dmFormName, CCHFORMNAME) = DEF_FormName
        .dmOrientation = DEF_Orientation
        .dmScale = DEF_Scale
        .dmPaperLength = DEF_PaperLength
        .dmPaperSize = DEF_PaperSize
        .dmPaperWidth = DEF_PaperWidth
        .dmPrintQuality = DEF_PrintQuality
        .dmTTOption = DEF_TTOption
        .dmYResolution = DEF_YResolution
    End With
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Read property values saved from a previous instance.
    '---------------------------------------------------------------------------------------
    With mtDevMode
        .dmSize = Len(mtDevMode)
        .dmSpecVersion = DM_SPECVERSION
        
        .dmCollate = PropBag.ReadProperty(PROP_Collate, DEF_Collate)
        .dmColor = PropBag.ReadProperty(PROP_Color, DEF_Color)
        .dmCopies = PropBag.ReadProperty(PROP_Copies, DEF_Copies)
        .dmDefaultSource = PropBag.ReadProperty(PROP_DefaultSource, DEF_DefaultSource)
        pByteString(.dmDeviceName, CCHDEVICENAME) = PropBag.ReadProperty(PROP_DeviceName, DEF_DeviceName)
        .dmDriverVersion = PropBag.ReadProperty(PROP_DriverVersion, DEF_DriverVersion)
        .dmDuplex = PropBag.ReadProperty(PROP_Duplex, DEF_Duplex)
        .dmFields = PropBag.ReadProperty(PROP_Fields, DEF_Fields)
        pByteString(.dmFormName, CCHFORMNAME) = PropBag.ReadProperty(PROP_FormName, DEF_FormName)
        .dmOrientation = PropBag.ReadProperty(PROP_Orientation, DEF_Orientation)
        .dmScale = PropBag.ReadProperty(PROP_Scale, DEF_Scale)
        .dmPaperLength = PropBag.ReadProperty(PROP_PaperLength, DEF_PaperLength)
        .dmPaperSize = PropBag.ReadProperty(PROP_PaperSize, DEF_PaperSize)
        .dmPaperWidth = PropBag.ReadProperty(PROP_PaperWidth, DEF_PaperWidth)
        .dmPrintQuality = PropBag.ReadProperty(PROP_PrintQuality, DEF_PrintQuality)
        .dmTTOption = PropBag.ReadProperty(PROP_TTOption, DEF_TTOption)
        .dmYResolution = PropBag.ReadProperty(PROP_YResolution, DEF_YResolution)
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    '---------------------------------------------------------------------------------------
    ' Date      : 12/15/04
    ' Purpose   : Store properties for future use.
    '---------------------------------------------------------------------------------------
    With mtDevMode
        PropBag.WriteProperty PROP_Collate, .dmCollate, DEF_Collate
        PropBag.WriteProperty PROP_Color, .dmColor, DEF_Color
        PropBag.WriteProperty PROP_Copies, .dmCopies, DEF_Copies
        PropBag.WriteProperty PROP_DefaultSource, .dmDefaultSource, DEF_DefaultSource
        PropBag.WriteProperty PROP_DeviceName, pByteString(.dmDeviceName, CCHDEVICENAME), DEF_DeviceName
        PropBag.WriteProperty PROP_DriverVersion, .dmDriverVersion, DEF_DriverVersion
        PropBag.WriteProperty PROP_Duplex, .dmDuplex, DEF_Duplex
        PropBag.WriteProperty PROP_Fields, .dmFields, DEF_Fields
        PropBag.WriteProperty PROP_FormName, pByteString(.dmFormName, CCHFORMNAME), DEF_FormName
        PropBag.WriteProperty PROP_Orientation, .dmOrientation, DEF_Orientation
        PropBag.WriteProperty PROP_Scale, .dmScale, DEF_Scale
        PropBag.WriteProperty PROP_PaperLength, .dmPaperLength, DEF_PaperLength
        PropBag.WriteProperty PROP_PaperSize, .dmPaperSize, DEF_PaperSize
        PropBag.WriteProperty PROP_PaperWidth, .dmPaperWidth, DEF_PaperWidth
        PropBag.WriteProperty PROP_PrintQuality, .dmPrintQuality, DEF_PrintQuality
        PropBag.WriteProperty PROP_TTOption, .dmTTOption, DEF_TTOption
        PropBag.WriteProperty PROP_YResolution, .dmYResolution, DEF_YResolution
    End With
End Sub
