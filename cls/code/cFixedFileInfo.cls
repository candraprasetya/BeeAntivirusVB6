VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFixedFileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************
'Module:        cFixedFileInfo - Class Module
'Filename:      cFixedFileInfo.cls
'Author:        Jim Kahl
'Purpose:       to extract information from a file at run time that will fill the
'               API Type structure of VS_FIXEDFILEINFO and convert it to a readable
'               class object
'NOTE:          code based on the original from Karl Peterson's One Stop Source Shop
'               http://www.mvps.org/vb
'               FileInfo sample
'****************************************************************************************
Option Explicit

'****************************************************************************************
'API CONSTANTS
'****************************************************************************************
Private Const MAX_PATH As Long = 260

Public Enum ffiFileFlags
    VS_FFI_SIGNATURE = &HFEEF04BD
    VS_FFI_STRUCVERSION = &H10000
    VS_FFI_FILEFLAGSMASK = &H3F&
    VS_FF_DEBUG = &H1               'contains debug information
    VS_FF_PRERELEASE = &H2          'not a production release
    VS_FF_PATCHED = &H4             'file has been modified and is not identical to the
                                    'original file
                                    
    VS_FF_PRIVATEBUILD = &H8        'was not built using standard release procedures
    VS_FF_INFOINFERRED = &H10       'version info was created dynamically data may be
                                    'missing
                                    
    VS_FF_SPECIALBUILD = &H20       'built using standard release procedures but is not
                                    'identical to the normal file with same version
    #If False Then
        VS_FFI_SIGNATURE
        VS_FFI_STRUCVERSION
        VS_FFI_FILEFLAGSMASK
        VS_FF_DEBUG
        VS_FF_PRERELEASE
        VS_FF_PATCHED
        VS_FF_PRIVATEBUILD
        VS_FF_INFOINFERRED
        VS_FF_SPECIALBUILD
    #End If
End Enum

Public Enum ffiOperatingSystems
    VOS__PM16 = &H2&                '16 bit Presentation Manager
    VOS__PM32 = &H3&                '32 bit Presentation Manager
    VOS__WINDOWS16 = &H1&           '16 bit Windows
    VOS__WINDOWS32 = &H4&           '32 bit Windows
    VOS_DOS = &H10000               'Standard DOS
    VOS_DOS_WINDOWS16 = &H10001     '16 bit Windows running on DOS
    VOS_DOS_WINDOWS32 = &H10004     '32 bit Windows running on DOS
    VOS_NT = &H40000                'Windows NT
    VOS_NT_WINDOWS32 = &H40004      '32 bit Windows running on Windows NT
    VOS_OS216 = &H20000             '16 bit OS/2
    VOS_OS216_PM16 = &H20002        '16 bit Presentation Manager running on 16 bit OS/2
    VOS_OS232 = &H30000             '32 bit OS/2
    VOS_OS232_PM32 = &H30003        '32 bit Presentation Manager running on 32 bit OS/2
    VOS_UNKNOWN = &H0&              'OS Unknown
    #If False Then
        VOS__PM16
        VOS__PM32
        VOS__WINDOWS16
        VOS__WINDOWS32
        VOS_DOS
        VOS_DOS_WINDOWS16
        VOS_DOS_WINDOWS32
        VOS_NT
        VOS_NT_WINDOWS32
        VOS_OS216
        VOS_OS216_PM16
        VOS_OS232
        VOS_OS232_PM32
        VOS_UNKNOWN
    #End If
End Enum

Public Enum ffiFileTypes
    VFT_UNKNOWN = &H0               'file is an undetermined type
    VFT_APP = &H1                   'an application
    VFT_DLL = &H2                   'an ActiveX or COM object file
    VFT_DRV = &H3                   'device driver
    VFT_FONT = &H4                  'font
    VFT_VXD = &H5                   'virtual device
    VFT_STATIC_LIB = &H7            'static library
    #If False Then
        VFT_UNKNOWN
        VFT_APP
        VFT_DLL
        VFT_DRV
        VFT_FONT
        VFT_VXD
        VFT_STATIC_LIB
    #End If
End Enum

Public Enum ffiFileSubtypes
    VFT2_FONT_RASTER = &H1&         'raster font
    VFT2_FONT_VECTOR = &H2&         'vector font
    VFT2_FONT_TRUETYPE = &H3&       'true type font
    VFT2_UNKNOWN = &H0              'unknown sub type
    VFT2_DRV_PRINTER = &H1          'printer driver
    VFT2_DRV_KEYBOARD = &H2         'keyboard driver
    VFT2_DRV_LANGUAGE = &H3         'language driver
    VFT2_DRV_DISPLAY = &H4          'display driver
    VFT2_DRV_MOUSE = &H5            'mouse driver
    VFT2_DRV_NETWORK = &H6          'network driver
    VFT2_DRV_SYSTEM = &H7           'system driver
    VFT2_DRV_INSTALLABLE = &H8      'installable driver
    VFT2_DRV_SOUND = &H9            'sound driver
    VFT2_DRV_COMM = &HA             'communication driver
    #If False Then
        VFT2_FONT_RASTER
        VFT2_FONT_VECTOR
        VFT2_FONT_TRUETYPE
        VFT2_UNKNOWN
        VFT2_DRV_PRINTER
        VFT2_DRV_KEYBOARD
        VFT2_DRV_LANGUAGE
        VFT2_DRV_DISPLAY
        VFT2_DRV_MOUSE
        VFT2_DRV_NETWORK
        VFT2_DRV_SYSTEM
        VFT2_DRV_INSTALLABLE
        VFT2_DRV_SOUND
        VFT2_DRV_COMM
    #End If
End Enum

'****************************************************************************************
'API TYPES
'****************************************************************************************
Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersion As Long              'e.g. 0x00000042 = "0.42"
    dwFileVersionMS As Long             'e.g. 0x00030075 = "3.75"
    dwFileVersionLS As Long             'e.g. 0x00000031 = "0.31"
    dwProductVersionMS As Long          'e.g. 0x00030010 = "3.10"
    dwProductVersionLS As Long          'e.g. 0x00000031 = "0.31"
    dwFileFlagsMask As Long             '= 0x3F for version "0.42"
    dwFileFlags As Long                 'e.g. VFF_DEBUG Or VFF_PRERELEASE
    dwFileOS As Long                    'e.g. VOS_DOS_WINDOWS16
    dwFileType As Long                  'e.g. VFT_DRIVER
    dwFileSubtype As Long               'e.g. VFT2_DRV_KEYBOARD
    dwFileDateMS As Long                'e.g. 0
    dwFileDateLS As Long                'e.g. 0
End Type


'****************************************************************************************
'API FUNCTIONS
'****************************************************************************************
Private Declare Sub CopyMemory Lib "KERNEL32" _
        Alias "RtlMoveMemory" ( _
                Destination As Any, _
                Source As Any, _
                ByVal Length As Long)
Private Declare Function GetFileVersionInfo Lib "Version.dll" _
        Alias "GetFileVersionInfoA" ( _
                ByVal lptstrFilename As String, _
                ByVal dwhandle As Long, _
                ByVal dwlen As Long, _
                lpData As Any) _
                As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" _
        Alias "GetFileVersionInfoSizeA" ( _
                ByVal lptstrFilename As String, _
                lpdwHandle As Long) _
                As Long
Private Declare Function GetFullPathName Lib "KERNEL32" _
        Alias "GetFullPathNameA" ( _
                ByVal lpFileName As String, _
                ByVal nBufferLength As Long, _
                ByVal lpBuff As String, _
                lpFilePart As Long) _
                As Long
Private Declare Function lstrlenA Lib "KERNEL32" (ByVal lpString As Long) As Long
Private Declare Function VerLanguageName Lib "KERNEL32" _
        Alias "VerLanguageNameA" ( _
                ByVal wLang As Long, _
                ByVal szLang As String, _
                ByVal nSize As Long) _
                As Long
Private Declare Function VerQueryValue Lib "Version.dll" _
        Alias "VerQueryValueA" ( _
                pBlock As Any, _
                ByVal lpSubBlock As String, _
                lpBuffer As Any, _
                lLen As Long) _
                As Long

'****************************************************************************************
'ENUMERATED CONSTANTS
'****************************************************************************************
Public Enum VersionInfoStrings
    viPredefinedFirst = 0
    viLanguage = 0
    viComments = 1
    viCompanyName = 2
    viFileDescription = 3
    viFileVersion = 4
    viInternalName = 5
    viLegalCopyright = 6
    viLegalTrademarks = 7
    viOriginalFilename = 8
    viProductName = 9
    viProductVersion = 10
    viPrivateBuild = 11
    viSpecialBuild = 12
    'these two are used by MSOffice apps only?
    viLegalTrademarks1 = 13
    viLegalTrademarks2 = 14
    viPredefinedLast = 14
    'trick to maintain proper case of enums in the IDE
    #If False Then
        viPredefinedFirst
        viLanguage
        viComments
        viCompanyName
        viFileDescription
        viFileVersion
        viInternalName
        viLegalCopyright
        viLegalTrademarks
        viOriginalFilename
        viProductName
        viProductVersion
        viPrivateBuild
        viSpecialBuild
        viLegalTrademarks1
        viLegalTrademarks2
        viPredefinedLast
    #End If
End Enum

'****************************************************************************************
'VARIABLES - PRIVATE
'****************************************************************************************
Private msPath As String
Private mbAvailable As Boolean
Private msStructVer As String               'structure Version - NOT USED
Private msFileVer As String                 'File Version
Private msProdVer As String                 'Product Version
Private meFileFlags As ffiFileFlags         'Boolean attributes of file
Private meFileOS As ffiOperatingSystems     'OS file is designed for
Private meFileType As ffiFileTypes          'Type of file
Private meFileSubType As ffiFileSubtypes    'Sub-type of file
Private msLanguage As String
Private msComments As String
Private msCompany As String
Private msDescription As String
Private msVerFileVer As String
Private msInternalName As String
Private msCopyright As String
Private msTrademark As String
Private msOrigFilename As String
Private msProductName As String
Private msVerProdVer As String
Private msPrivateBuild As String
Private msSpecialBuild As String
Private msTrademarks1 As String
Private msTrademarks2 As String

'****************************************************************************************
'PROPERTIES - PUBLIC READ/WRITE
'****************************************************************************************
'FullPathName - sets/returns the path and file name of the file
Public Property Let FullPathName(ByVal Filespec As String)
    Dim sBuff As String
    Dim lFilePart As Long
    Dim lRet As Long
    
    sBuff = Space$(MAX_PATH)
    lRet = GetFullPathName(Filespec, Len(sBuff), sBuff, lFilePart)
    If lRet Then
        msPath = Left$(sBuff, lRet)
        Initialize
    End If
End Property

Public Property Get FullPathName() As String
    FullPathName = msPath
End Property

'****************************************************************************************
'PROPERTIES - PUBLIC READ ONLY
'****************************************************************************************
Public Property Get FileName() As String
    'Returns:       filename and extension of the file
    Dim lPos As Long
    
    lPos = InStrRev(msPath, "\")
    
    FileName = Right$(msPath, Len(msPath) - lPos)
End Property

Public Property Get FileExtension() As String
    Dim lPos As Long
    
    lPos = InStrRev(msPath, ".")
    
    FileExtension = Right(msPath, Len(msPath) - lPos)
End Property

Public Property Get Available() As Boolean
    'Returns:       True - version information is available
    '               False - version information is not available
    Available = mbAvailable
End Property

Public Property Get FileFlags() As ffiFileFlags
    FileFlags = meFileFlags
End Property

Public Property Get FileOS() As ffiOperatingSystems
    FileOS = meFileOS
End Property

Public Property Get FileType() As ffiFileTypes
    FileType = meFileType
End Property

Public Property Get FileSubType() As ffiFileSubtypes
    FileSubType = meFileSubType
End Property

Public Property Get VerFile() As String
    VerFile = msFileVer
End Property

Public Property Get VerProduct() As String
    VerProduct = msProdVer
End Property

Public Property Get VerStructure() As String
    VerStructure = msStructVer
End Property

Public Property Get Comments() As String
    Comments = msComments
End Property

Public Property Get CompanyName() As String
    CompanyName = msCompany
End Property

Public Property Get FileDescription() As String
    FileDescription = msDescription
End Property

Public Property Get FileVersion() As String
    FileVersion = msVerFileVer
End Property

Public Property Get InternalName() As String
    InternalName = msInternalName
End Property

Public Property Get Language() As String
    Language = msLanguage
End Property

Public Property Get LegalCopyright() As String
    LegalCopyright = msCopyright
End Property

Public Property Get LegalTrademarks() As String
    LegalTrademarks = msTrademark
End Property

Public Property Get OriginalFilename() As String
    OriginalFilename = msOrigFilename
End Property

Public Property Get ProductName() As String
    ProductName = msProductName
End Property

Public Property Get ProductVersion() As String
    ProductVersion = msVerProdVer
End Property

Public Property Get PrivateBuild() As String
    PrivateBuild = msPrivateBuild
End Property

Public Property Get SpecialBuild() As String
    SpecialBuild = msSpecialBuild
End Property

'****************************************************************************************
'METHODS - PUBLIC
'****************************************************************************************
Public Sub Initialize()
    'Purpose:       initialize the type structure and fill the elements of it
    Dim lTemp As Long
    Dim lRet As Long
    Dim sBuff() As Byte
    Dim lBuffLen As Long
    Dim lpBuff As Long
    Dim uFFI As VS_FIXEDFILEINFO
    Dim lLen As Long
    Dim lLang As Integer
    Dim lCodePg As Integer
    Dim sSubBlock As String
    Dim sTemp As String
    
    'Get size of udt
    lBuffLen = GetFileVersionInfoSize(msPath, lTemp)
    If lBuffLen Then
        mbAvailable = True
    Else
        mbAvailable = False
        Exit Sub
    End If

    'store info to our udt
    ReDim sBuff(lBuffLen) As Byte
    Call GetFileVersionInfo(msPath, 0&, lBuffLen, sBuff(0))
    Call VerQueryValue(sBuff(0), "\", lpBuff, lLen)
    Call CopyMemory(uFFI, ByVal lpBuff, Len(uFFI))
    
    'The StructureVersion we can format it the way we want
    'determine structure version - handle as dword
    sTemp = hiWord(uFFI.dwStrucVersion)
    msStructVer = sTemp & "." & Format$(loWord(uFFI.dwStrucVersion), "00")
    
    'VerFile we can format the way we want
    'determine the file version - handle as dword
    sTemp = hiWord(uFFI.dwFileVersionMS)
    msFileVer = sTemp & "." & loWord(uFFI.dwFileVersionMS)
    sTemp = hiWord(uFFI.dwFileVersionLS)
    If sTemp = "0" Then
        sTemp = vbNullString
    End If
    sTemp = sTemp & loWord(uFFI.dwFileVersionLS)
    msFileVer = msFileVer & "." & sTemp
    
    'VerProduct we can format the way we want
    'determine product version - handle as dword
    sTemp = hiWord(uFFI.dwProductVersionMS)
    msProdVer = sTemp & "." & loWord(uFFI.dwProductVersionMS)
    sTemp = hiWord(uFFI.dwProductVersionLS)
    If sTemp = "0" Then
        sTemp = vbNullString
    End If
    sTemp = sTemp & loWord(uFFI.dwProductVersionLS)
    msProdVer = msProdVer & "." & sTemp
    
    'get the file flags
    meFileFlags = uFFI.dwFileFlags
    
    'determine OS for which file was designed
    meFileOS = uFFI.dwFileOS

    'determine type of file
    meFileType = uFFI.dwFileType
    meFileSubType = uFFI.dwFileSubtype
    
    'get language translations
    If VerQueryValue(sBuff(0), "\VarFileInfo\Translation", lpBuff, lLen) Then
        If lLen Then
            'handle the language/codepage as a dword.
            lRet = pointerToDWord(lpBuff)
            lLang = loWord(lRet)
            lCodePg = hiWord(lRet)
            
            'determine language
            msLanguage = Space$(256)
            lRet = VerLanguageName(CLng(lLang), msLanguage, Len(msLanguage))
            If lRet Then
                msLanguage = Left$(msLanguage, lRet)
            Else
                msLanguage = vbNullString
            End If
            
            'COMMENT FROM KARL PETERSON'S ORIGINAL SOURCE
            'Microsoft really screwed up in a number of their
            'applications, including Office 95 and Office 97,
            'by storing the resources under the U.S. English
            'codepage after identifying the language as neutral.
            'Test to see if no company name shows up in these
            'cases, then try English to see if that works better.
            
            sSubBlock = "\StringFileInfo\" & _
            formatHex(lLang, 4) & _
            formatHex(lCodePg, 4) & "\"
            msCompany = getStdValue(VarPtr(sBuff(0)), sSubBlock & "CompanyName")
            If Len(msCompany) = 0 Then
                'Try U.S. English...?
                sTemp = "\StringFileInfo\" & formatHex(&H409, 4) & _
                        formatHex(lCodePg, 4) & "\"
                msCompany = getStdValue(VarPtr(sBuff(0)), sTemp & "CompanyName")
                If Len(msCompany) > 0 Then
                    'We probably found the MS version bug.
                    sSubBlock = sTemp
                End If
            End If
            
            'Get remaining predefined version resources
            msComments = getStdValue(VarPtr(sBuff(0)), sSubBlock & "Comments")
            msDescription = getStdValue(VarPtr(sBuff(0)), sSubBlock & "FileDescription")
            msVerFileVer = getStdValue(VarPtr(sBuff(0)), sSubBlock & "FileVersion")
            msInternalName = getStdValue(VarPtr(sBuff(0)), sSubBlock & "InternalName")
            msCopyright = getStdValue(VarPtr(sBuff(0)), sSubBlock & "LegalCopyright")
            msTrademark = getStdValue(VarPtr(sBuff(0)), sSubBlock & "LegalTrademarks")
            msOrigFilename = getStdValue(VarPtr(sBuff(0)), sSubBlock & "OriginalFilename")
            msProductName = getStdValue(VarPtr(sBuff(0)), sSubBlock & "ProductName")
            msVerProdVer = getStdValue(VarPtr(sBuff(0)), sSubBlock & "ProductVersion")
            msPrivateBuild = getStdValue(VarPtr(sBuff(0)), sSubBlock & "PrivateBuild")
            msSpecialBuild = getStdValue(VarPtr(sBuff(0)), sSubBlock & "SpecialBuild")
            msTrademarks1 = getStdValue(VarPtr(sBuff(0)), sSubBlock & "LegalTrademarks1")
            msTrademarks2 = getStdValue(VarPtr(sBuff(0)), sSubBlock & "LegalTrademarks2")
        End If
    End If
End Sub

'****************************************************************************************
'METHODS - PRIVATE
'****************************************************************************************
Private Function formatHex(ByVal InVal As Long, ByVal OutLen As Integer) As String
    'Left pad with zeros to OutLen.
    formatHex = Right$(String$(OutLen, "0") & Hex$(InVal), OutLen)
End Function

Private Function getStdValue(ByVal lpBlock As Long, ByVal Value As String) As String
    'Purpose:       to retrieve version infromation that does not need to be decoded
    Dim lpBuff As Long
    Dim lLen As Long
    
    If VerQueryValue(ByVal lpBlock, Value, lpBuff, lLen) Then
        If lLen Then
            getStdValue = pointerToString(lpBuff)
        End If
    End If
End Function

Private Function loWord(ByVal LongIn As Long) As Long
    'Purpose:       to treat the lowest 2 bytes of a long integer as a single number
    Call CopyMemory(loWord, LongIn, 2)
End Function

Private Function hiWord(ByVal LongIn As Long) As Long
    'Purpose:       to treat the highest 2 bytes of a long integer as a single number
    Call CopyMemory(hiWord, ByVal (VarPtr(LongIn) + 2), 2)
End Function

Private Function pointerToDWord(ByVal lpDWord As Long) As Long
    'pointer to a DWORD
    Dim lRet As Long
    
    If lpDWord Then
        CopyMemory lRet, ByVal lpDWord, 4
        pointerToDWord = lRet
    End If
End Function

Private Function pointerToString(ByRef lpString As Long) As String
    'Pointer to an ANSI string
    Dim byBuff As String
    Dim lLen As Long
   
    If lpString Then
        lLen = lstrlenA(lpString)
        If lLen Then
            byBuff = Space$(lLen)
            CopyMemory ByVal byBuff, ByVal lpString, lLen
            pointerToString = byBuff
        End If
    End If
End Function


