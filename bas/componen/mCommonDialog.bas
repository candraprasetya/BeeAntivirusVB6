Attribute VB_Name = "mCommonDialog"
'==================================================================================================
'mCommonDialog.bas              8/25/04
'
'           GENERAL PURPOSE:
'               Implement all modal COMDLG32 common dialogs, shell browse for folder and html help.
'
'           LINEAGE:
'               CommonDialogDirect6 from www.vbaccelerator.com
'
'==================================================================================================

Option Explicit

Public Enum eDialogType
    dlgTypeFile
    dlgTypeFont
    dlgTypeColor
    dlgTypePrint
    dlgTypePageSetup
    dlgTypeFolder
End Enum

Public Type tColorDialog
    'in:
    hWndOwner               As Long
    iFlags                  As Long
    oHookCallback           As iComDlgHook
    
    'out
    iReturnExtendedError    As Long
    
    'in/out:
    iColor                  As OLE_COLOR
    iColors(0 To 15)        As Long
End Type

Public Type tFileDialog
    'in:
    iFlags                  As Long
    sFilter                 As String
    iFilterIndex            As Long
    sDefaultExt             As String
    sInitPath               As String
    sInitFile               As String
    sTitle                  As String
    hWndOwner               As Long
    vTemplate               As Variant
    hInstance               As Long
    oHookCallback           As iComDlgHook
    
    'out:
    sReturnFileName         As String
    iReturnFlags            As Long
    iReturnExtendedError    As Long
    iReturnFilterIndex      As Long
End Type
   
Public Type tFolderDialog
    'in:
    hWndOwner               As Long
    sTitle                  As String
    sInitialPath            As String
    sRootPath               As String
    iFlags                  As Long
    oHookCallback           As iComDlgHook
    
    'out:
    sReturnPath             As String
End Type
    
Public Type tFontDialog
    'in:
    iFlags                  As Long
    hdc                     As Long
    hWndOwner               As Long
    iMinSize                As Long
    iMaxSize                As Long
    oHookCallback           As iComDlgHook
    
    'out
    iReturnFlags            As Long
    iReturnExtendedError    As Long
    
    'in/out:
    iColor                  As OLE_COLOR
    oFont                   As Object
End Type

Public Type tHelpDialog
    'in:
    hWnd                    As Long
    iCommand                As Long
    sFile                   As String
    vTopic                  As Variant
End Type
   
Public Type tPageSetupDialog
    'in:
    iFlags                  As Long
    hWndOwner               As Long
    iUnits                  As Long
    fMinLeftMargin          As Single
    fMinRightMargin         As Single
    fMinTopMargin           As Single
    fMinBottomMargin        As Single
    oHookCallback           As iComDlgHook
    
    'out:
    iReturnExtendedError    As Long
    
    'in/out:
    fLeftMargin             As Single
    fRightMargin            As Single
    fTopMargin              As Single
    fBottomMargin           As Single
    oDeviceMode             As cDeviceMode
    oDeviceNames            As cDeviceNames
End Type
    
Public Type tPrintDialog
    'in:
    hWndOwner               As Long
    iFlags                  As Long
    iMinPage                As Long
    iMaxPage                As Long
    oHookCallback           As iComDlgHook
    
    'out:
    hdc                     As Long
    iReturnFlags            As Long
    iReturnExtendedError    As Long
    
    'in/out:
    iRange                  As Long
    iFromPage               As Long
    iToPage                 As Long
    oDeviceMode             As cDeviceMode
    oDeviceNames            As cDeviceNames
    
End Type

Private Function pGetHook(ByRef iFlags As Long, ByVal iHookFlag As Long, ByVal iDialogType As eDialogType, ByVal oObject As iComDlgHook) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Allocate assembly code to use as a common dialog hook callback function.
    '---------------------------------------------------------------------------------------
    
    If CBool(iFlags And iHookFlag) And Not oObject Is Nothing Then
        
        Const PATCH_Ide As Long = 6
        Const PATCH_EbMode As Long = 9
        Const PATCH_DialogType As Long = 38
        Const PATCH_ObjPtr As Long = 44
        
        pGetHook = Thunk_Alloc(tnkCommonDialogProc, PATCH_Ide)                  'allocate the assembly code
        If pGetHook Then
            Thunk_Patch pGetHook, PATCH_DialogType, iDialogType                 'patch the dialog type
            Thunk_Patch pGetHook, PATCH_ObjPtr, ObjPtr(oObject)                 'patch the owner object
            Thunk_PatchFuncAddr pGetHook, PATCH_EbMode, "vba6.dll", "EbMode"    'patch the relative address to vba6.EbMode
        Else
            iFlags = (iFlags And Not iHookFlag)                                 'if the allocation failed, remove the hook flag
        End If
        
    End If
    
End Function

Private Sub pReleaseHook(ByVal lpHookProc As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Release memory allocated for a common dialog hook callback function.
    '---------------------------------------------------------------------------------------
    If lpHookProc Then MemFree lpHookProc
End Sub
    
    
Public Function File_ShowOpenIndirect( _
ByRef tFileDialog As tFileDialog) _
As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Show the file open dialog from a udt.
    '---------------------------------------------------------------------------------------
    File_ShowOpenIndirect = pFile_Show(tFileDialog, False)
End Function

Public Function File_ShowSaveIndirect( _
ByRef tFileDialog As tFileDialog) _
As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Show the file save dialog from a udt.
    '---------------------------------------------------------------------------------------
    File_ShowSaveIndirect = pFile_Show(tFileDialog, True)
End Function

Public Function File_GetMultiFileNames(ByRef sReturnFileName As String, ByRef sReturnPath As String, ByRef sFileNames() As String) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Transfer the filename(s) from a delimited string to a string array.
    '---------------------------------------------------------------------------------------
    
    Erase sFileNames
    sReturnPath = vbNullString
    
    Dim liTemp      As Long
    liTemp = InStr(sReturnFileName, vbNullChar)
    
    If liTemp > ZeroL Then
        sReturnPath = Left$(sReturnFileName, liTemp - 1)
        If Right$(sReturnPath, 1) <> "\" Then sReturnPath = sReturnPath & "\"
        sFileNames = Split(Mid$(sReturnFileName, liTemp + 1), vbNullChar)
        File_GetMultiFileNames = UBound(sFileNames) + 1&
    Else
        liTemp = InStrRev(sReturnFileName, "\")
        If liTemp > ZeroL Then
            sReturnPath = Left$(sReturnFileName, liTemp)
            ReDim sFileNames(0 To 0)
            sFileNames(0) = Mid$(sReturnFileName, liTemp + 1&)
            File_GetMultiFileNames = 1&
        End If
    End If
End Function
    
Private Function pFile_Show( _
ByRef tFileDialog As tFileDialog, _
ByVal bSave As Boolean) _
As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Show a file open or save dialog.
    '---------------------------------------------------------------------------------------

    Const ValidFlags As Long = OFN_EXPLORER Or _
    OFN_FILEMUSTEXIST Or _
    OFN_PATHMUSTEXIST Or _
    OFN_ALLOWMULTISELECT Or _
    OFN_CREATEPROMPT Or _
    OFN_ENABLESIZING Or _
    OFN_NODEREFERENCELINKS Or _
    OFN_NONETWORKBUTTON Or _
    OFN_HIDEREADONLY Or _
    OFN_NOREADONLYRETURN Or _
    OFN_NOTESTFILECREATE Or _
    OFN_OVERWRITEPROMPT Or _
    OFN_READONLY Or _
    OFN_SHOWHELP Or _
    OFN_ENABLEHOOK Or _
    OFN_ENABLETEMPLATE Or _
    OFN_DONTADDTORECENT Or _
    OFN_EXTENSIONDIFFERENT
    
    Dim ltOFN      As OPENFILENAME
    
    Const lpstrFile_Len_MultiSelect As Long = 8192&
    Const lpstrFile_Len_SingleSelect As Long = MAX_PATH
    
    'Allocate ANSI strings to pass to the api
    Dim lpstrInitialDir      As String:  lpstrInitialDir = StrConv(tFileDialog.sInitPath & vbNullChar, vbFromUnicode)
    Dim lpstrDefExt          As String:      lpstrDefExt = StrConv(tFileDialog.sDefaultExt & vbNullChar, vbFromUnicode)
    Dim lpstrTitle           As String:       lpstrTitle = StrConv(tFileDialog.sTitle & vbNullChar, vbFromUnicode)
    Dim lpstrFilter          As String:      lpstrFilter = StrConv(Replace$(tFileDialog.sFilter, OFN_FilterDelim, vbNullChar) & vbNullChar & vbNullChar, vbFromUnicode)
    Dim lpstrInitFile        As String:    lpstrInitFile = StrConv(tFileDialog.sInitFile & vbNullChar, vbFromUnicode)
    Dim lpstrTemplate        As String
    
    With tFileDialog
        .sReturnFileName = vbNullString
        .iReturnExtendedError = ZeroL
        .iReturnFlags = ZeroL
        .iReturnFilterIndex = NegOneL
    End With
    
    With ltOFN
        .lStructSize = LenB(ltOFN)                                      'store the structure size
        .Flags = tFileDialog.iFlags And ValidFlags                      'mask out invalid flags
        .hInstance = tFileDialog.hInstance                              'store hInstance
        .hWndOwner = tFileDialog.hWndOwner                              'store hwndOwner
        .lCustData = ZeroL                                              'not used
        'allocate the hook procedure, if any
        .lpfnHook = pGetHook(.Flags, OFN_ENABLEHOOK, dlgTypeFile, tFileDialog.oHookCallback)
        
        .lpstrCustomFilter = ZeroL                                      'not used
        .nMaxCustFilter = ZeroL                                         'not used
        
        .lpstrDefExt = StrPtr(lpstrDefExt)
        
        If CBool(.Flags And OFN_ALLOWMULTISELECT) Then                  'if multiselect, allocate an extra large buffer
            .lpstrFile = MemAllocFromString(StrPtr(lpstrInitFile), lpstrFile_Len_MultiSelect)
            .nMaxFile = lpstrFile_Len_MultiSelect
        Else                                                            'if single select, allocate a normal size buffer
            .lpstrFile = MemAllocFromString(StrPtr(lpstrInitFile), lpstrFile_Len_SingleSelect)
            .nMaxFile = lpstrFile_Len_SingleSelect
        End If
        
        If .lpstrFile = ZeroL Then .nMaxFile = ZeroL                    'just in case allocation failed
        
            .lpstrFileTitle = ZeroL                                         'not used
            .nMaxFileTitle = ZeroL                                          'not used
        
            .lpstrFilter = StrPtr(lpstrFilter)                              'pointer to our local ANSI string
            .lpstrInitialDir = StrPtr(lpstrInitialDir)                      'pointer to our local ANSI string
            .lpstrTitle = StrPtr(lpstrTitle)                                'pointer to our local ANSI string
        
            If .Flags And OFN_ENABLETEMPLATE Then
                If VarType(tFileDialog.vTemplate) = vbString Then           'if template is string id
                    lpstrTemplate = StrConv(CStr(tFileDialog.vTemplate) & vbNullChar, vbFromUnicode)
                    .lpTemplateName = StrPtr(lpstrTemplate)
                Else
                    On Error Resume Next                                    'otherwise, template is numeric id
                    .lpTemplateName = CLng(tFileDialog.vTemplate)
                    On Error GoTo 0
                End If
            
                If .lpTemplateName = ZeroL Or .hInstance = ZeroL Then       'if template id or hinstance is invalid, cancel
                    .lpTemplateName = ZeroL
                    .hInstance = ZeroL
                    .Flags = .Flags And Not OFN_ENABLETEMPLATE
                End If
            Else
                .hInstance = ZeroL                                          'not used
                .lpTemplateName = ZeroL                                     'not used
            End If
        
            .nFileExtension = ZeroL                                         'not used
            .nFileOffset = ZeroL                                            'not used
            .nFilterIndex = tFileDialog.iFilterIndex                        'store initial filter index
        
            Dim liReturn      As Long
        
            If bSave _
                Then liReturn = GetSaveFileName(ltOFN) _
            Else liReturn = GetOpenFileName(ltOFN)                      'call the api
        
                pReleaseHook .lpfnHook
        
                If liReturn Then                                                'if success

                    pFile_Show = True                                           'return true
                    tFileDialog.iReturnFlags = .Flags                           'return the flags
                    tFileDialog.iReturnFilterIndex = .nFilterIndex              'return the filter index
                    'get the return file name(s)
                    lstrToStringA .lpstrFile, tFileDialog.sReturnFileName, .nMaxFile
            
                    If CBool(tFileDialog.iFlags And OFN_ALLOWMULTISELECT) Then  'if multiselect
                        If CBool(tFileDialog.iFlags And OFN_EXPLORER) Then      'if explorer style
                            'return the file names (already null separated)
                            liReturn = InStr(1, tFileDialog.sReturnFileName, vbNullChar & vbNullChar)
                            If liReturn > ZeroL Then
                                tFileDialog.sReturnFileName = Left$(tFileDialog.sReturnFileName, liReturn - OneL)
                            Else
                                'should never happen!
                                ''debug.assert False
                            End If
                        Else                                                    'if multiselect and not explorerstyle (space separated)
                            liReturn = InStr(1, tFileDialog.sReturnFileName, vbNullChar)
                            If liReturn > ZeroL Then
                                tFileDialog.sReturnFileName = Left$(tFileDialog.sReturnFileName, liReturn - OneL)
                            Else
                                'should never happen!
                                ''debug.assert False
                            End If
                            'return the file name(s), (space separated -> null separated)
                            tFileDialog.sReturnFileName = Replace$(tFileDialog.sReturnFileName, " ", vbNullChar)
                        End If
                
                    Else                                                        'if not multiselect
                        liReturn = InStr(1, tFileDialog.sReturnFileName, vbNullChar)
                        If liReturn > ZeroL Then
                            tFileDialog.sReturnFileName = Left$(tFileDialog.sReturnFileName, liReturn - OneL)
                        Else
                            'should never happen!
                            ''debug.assert False
                        End If
                    End If
        
                Else                                                            'if api failed (cancel or error)
            
                    tFileDialog.iReturnExtendedError = CommDlgExtendedError()
            
                End If
        
                If .lpstrFile Then MemFree .lpstrFile                     'free the string buffer
        
                End With
    
End Function

Private Function pFolder_PathToPidl(ByRef spath As String) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Get a pidl from a path.
    '---------------------------------------------------------------------------------------
    Dim pIdlMain      As Long, cParsed As Long, afItem As Long
    If pFolder_GetDesktopFolder.ParseDisplayName(ZeroL, ZeroL, StrConv(spath, vbUnicode), cParsed, pIdlMain, afItem) = ZeroL Then
        pFolder_PathToPidl = pIdlMain
        #If bDebug Then
        mDebug.DEBUG_Add DEBUG_pIdl, pIdlMain
        #End If
    End If
End Function

Public Function Folder_PathFromPidl(ByVal pidl As Long) As String
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Get a path from a pidl.
    '---------------------------------------------------------------------------------------
    Dim lpPath      As Long
    lpPath = MemAllocFromString(ZeroL, MAX_PATH)
    If lpPath Then
        If SHGetPathFromIDList(pidl, ByVal lpPath) Then
            lstrToStringA lpPath, Folder_PathFromPidl
        End If
        MemFree lpPath
    End If
End Function

Private Sub pFolder_Free(ByVal pidl As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Return the shell IMalloc object
    '---------------------------------------------------------------------------------------
    Static oAlloc As IMalloc
    If oAlloc Is Nothing Then SHGetMalloc oAlloc
    oAlloc.Free ByVal pidl
    #If bDebug Then
    DEBUG_Remove DEBUG_pIdl, pidl
    #End If
End Sub

Private Function pFolder_GetDesktopFolder() As IShellFolder
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Return the desktop IShellFolder interface
    '---------------------------------------------------------------------------------------
    SHGetDesktopFolder pFolder_GetDesktopFolder
End Function

Public Function Color_ShowIndirect(ByRef tColorDialog As tColorDialog) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Show a choose color dialog.
    '---------------------------------------------------------------------------------------

    Const ValidFlags As Long = CC_FULLOPEN Or _
    CC_PREVENTFULLOPEN Or _
    CC_SOLIDCOLOR Or _
    CC_ANYCOLOR Or _
    CC_ENABLEHOOK

    Dim ltChooseColor      As CHOOSECOLOR
    
    tColorDialog.iReturnExtendedError = ZeroL
    
    With ltChooseColor
        .lStructSize = Len(ltChooseColor)                                           'initialize the api structure
        .hWndOwner = tColorDialog.hWndOwner
        If OleTranslateColor(tColorDialog.iColor, ZeroL, .rgbResult) Then .rgbResult = NegOneL
        .Flags = (tColorDialog.iFlags And ValidFlags) Or Abs(CBool(.rgbResult <> NegOneL))
        .lpCustColors = VarPtr(tColorDialog.iColors(0))
        
        .lpfnHook = pGetHook(.Flags, CC_ENABLEHOOK, _
        dlgTypeColor, tColorDialog.oHookCallback)              'store the address of the hook procedure
        
        .hInstance = ZeroL                                                          'make it obvious we don't use these members
        .lCustData = ZeroL
        .lpTemplateName = ZeroL
        
        Dim liReturn      As Long
        
        liReturn = CHOOSECOLOR(ltChooseColor)                                       'show the dialog
        
        pReleaseHook .lpfnHook
        
        If liReturn Then                                                            'if succeeded
            Color_ShowIndirect = True
            tColorDialog.iColor = .rgbResult                                        'return the color
        Else
            tColorDialog.iColor = NegOneL
            tColorDialog.iReturnExtendedError = CommDlgExtendedError()              'store the error value
        End If
    End With

End Function

Public Property Get Color_OKMsg() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Return the message value the color OK confirmation.
    '---------------------------------------------------------------------------------------
    Static B As Boolean
    Static iMsg As Long
    Const COLOROKSTRING As String = "commdlg_ColorOK"

    If Not B Then
        Dim lsAnsi      As String
        lsAnsi = StrConv(COLOROKSTRING & vbNullChar, vbFromUnicode)
        iMsg = RegisterWindowMessage(ByVal StrPtr(lsAnsi))
        B = True
    End If
    Color_OKMsg = iMsg
End Property
    
    
Public Function Font_ShowIndirect(ByRef tFontDialog As tFontDialog) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Show the choose font dialog.
    '---------------------------------------------------------------------------------------

    Const ValidFlags As Long = _
    CF_SCREENFONTS Or _
    CF_PRINTERFONTS Or _
    CF_USESTYLE Or _
    CF_EFFECTS Or _
    CF_ANSIONLY Or _
    CF_NOVECTORFONTS Or _
    CF_NOSIMULATIONS Or _
    CF_FIXEDPITCHONLY Or _
    CF_WYSIWYG Or _
    CF_FORCEFONTEXIST Or _
    CF_SCALABLEONLY Or _
    CF_TTONLY Or _
    CF_NOFACESEL Or _
    CF_NOSTYLESEL Or _
    CF_NOSIZESEL Or _
    CF_SELECTSCRIPT Or _
    CF_NOSCRIPTSEL Or _
    CF_NOVERTFONTS Or _
    CF_APPLY Or _
    CF_ENABLEHOOK
    
    Dim ltChooseFont      As CHOOSEFONT
    Dim ltLF              As LOGFONT
    
    With tFontDialog
        If .oFont Is Nothing Then Set .oFont = New StdFont
        
        If TypeOf .oFont Is StdFont Then
            pFont_PutStdFont ltLF, .oFont
        ElseIf TypeOf .oFont Is cFont Then
            pFont_PutFont ltLF, .oFont
        End If
        
        .iReturnFlags = ZeroL                                       'initialize the return values
        .iReturnExtendedError = ZeroL
    End With

    With ltChooseFont
        .lStructSize = LenB(ltChooseFont)                           'initialize the api structure
        .Flags = (tFontDialog.iFlags And ValidFlags) _
        Or CF_INITTOLOGFONTSTRUCT _
        Or CF_LIMITSIZE * -CBool(tFontDialog.iMinSize > ZeroL Or tFontDialog.iMaxSize > ZeroL)
        .hdc = tFontDialog.hdc
        .hInstance = ZeroL
        .hWndOwner = tFontDialog.hWndOwner
        .iPointSize = ZeroL
        .lCustData = ZeroL
        .lpfnHook = pGetHook(.Flags, CF_ENABLEHOOK, dlgTypeFont, tFontDialog.oHookCallback)
        .lpLogFont = VarPtr(ltLF)
        .lpszStyle = ZeroL
        .lpTemplateName = ZeroL
        .nFontType = ZeroL
        .nSizeMin = tFontDialog.iMinSize
        .nSizeMax = tFontDialog.iMaxSize
        .rgbColors = tFontDialog.iColor
    End With
    
    Dim liReturn      As Long
    liReturn = CHOOSEFONT(ltChooseFont)
    
    pReleaseHook ltChooseFont.lpfnHook                              'show the dialog

    If liReturn Then                                                'if we succeeded
        Font_ShowIndirect = True
        
        With tFontDialog
            If TypeOf .oFont Is StdFont Then                        'return the font data
                pFont_GetStdFont ltLF, .oFont
                
            ElseIf TypeOf .oFont Is cFont Then
                pFont_GetFont ltLF, .oFont
                
            Else
                'unknown font object, can't return the font chosen!
                ''debug.assert False
            End If
            
            .iReturnFlags = ltChooseFont.Flags                      'return the flags
            .iColor = ltChooseFont.rgbColors                        'return the color
        End With
    Else                                                            'if failed
        tFontDialog.iReturnExtendedError = CommDlgExtendedError()   'return the extended error
        
    End If
    
End Function

Private Sub pFont_PutStdFont(ByRef tLogFont As LOGFONT, ByVal oFont As StdFont)
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Put data from a StdFont object into a LOGFONT structure.
    '---------------------------------------------------------------------------------------
    On Error Resume Next
    With tLogFont
        BytesFromString .lfFaceName, oFont.name
        .lfHeight = -MulDiv(oFont.SIZE, 1440& / Screen.TwipsPerPixelY, 72&)
        .lfWeight = IIf(oFont.Bold, fntWeightBold, fntWeightNormal)
        .lfItalic = Abs(oFont.Italic)
        .lfUnderline = Abs(oFont.Underline)
        .lfStrikeOut = Abs(oFont.Strikethrough)
        .lfCharSet = oFont.Charset And &HFF
        .lfEscapement = ZeroL
        .lfOrientation = ZeroL
        .lfOutPrecision = 0
        .lfClipPrecision = 0
        .lfQuality = 0
        .lfPitchAndFamily = 0
    End With
    On Error GoTo 0
End Sub

Private Sub pFont_GetStdFont(ByRef tLogFont As LOGFONT, ByVal oFont As StdFont)
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Get data from a LOGFONT structure into a StdFont object.
    '---------------------------------------------------------------------------------------
    On Error Resume Next
    With oFont
        Dim lsName      As String
        StringFromBytes tLogFont.lfFaceName, lsName
        .name = lsName
        If tLogFont.lfHeight Then
            .SIZE = MulDiv(72&, Abs(tLogFont.lfHeight), (1440& / Screen.TwipsPerPixelY))
        End If
        .Charset = tLogFont.lfCharSet
        .Italic = CBool(tLogFont.lfItalic)
        .Underline = CBool(tLogFont.lfUnderline)
        .Strikethrough = CBool(tLogFont.lfStrikeOut)
        .Bold = CBool(tLogFont.lfWeight > fntWeightNormal)
    End With
    On Error GoTo 0
End Sub
    
Private Sub pFont_GetFont(ByRef tLogFont As LOGFONT, ByVal oFont As cFont)
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Get data from a LOGFONT structure into a cFont object.
    '---------------------------------------------------------------------------------------
    oFont.fPutLogFontLong VarPtr(tLogFont)
End Sub

Private Sub pFont_PutFont(ByRef tLogFont As LOGFONT, ByVal oFont As cFont)
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Put data from a cFont object into a LOGFONT structure.
    '---------------------------------------------------------------------------------------
    oFont.fGetLogFontLong VarPtr(tLogFont)
End Sub


Public Function Print_ShowIndirect(ByRef tDialog As tPrintDialog) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Show the Print Dialog.
    '---------------------------------------------------------------------------------------
    
    Const ValidRange As Long = PD_SELECTION Or PD_PAGENUMS
    
    Const ValidFlags As Long = _
    PD_ALLPAGES Or _
    PD_SELECTION Or _
    PD_PAGENUMS Or _
    PD_NOSELECTION Or _
    PD_NOPAGENUMS Or _
    PD_COLLATE Or _
    PD_PRINTTOFILE Or _
    PD_PRINTSETUP Or _
    PD_NOWARNING Or _
    PD_RETURNDC Or _
    PD_RETURNIC Or _
    PD_RETURNDEFAULT Or _
    PD_SHOWHELP Or _
    PD_ENABLEPRINTHOOK Or _
    PD_ENABLESETUPHOOK Or _
    PD_DISABLEPRINTTOFILE Or _
    PD_HIDEPRINTTOFILE Or _
    PD_NONETWORKBUTTON
    
    Dim hDevMode        As Long
    Dim hDevNames       As Long
    Dim liHookProc      As Long

    tDialog.iReturnExtendedError = ZeroL                                    'initialize return values
    tDialog.iReturnFlags = ZeroL

    ' Fill in PRINTDLG structure
    Dim ltPrintDialog      As PRINTDLG
    With ltPrintDialog
        .lStructSize = Len(ltPrintDialog)                                   'init the api structure
        .hWndOwner = tDialog.hWndOwner
        .Flags = (tDialog.iFlags And ValidFlags) Or (tDialog.iRange And ValidRange) Or PD_USEDEVMODECOPIESANDCOLLATE
        .nFromPage = tDialog.iFromPage
        .nToPage = tDialog.iToPage
        .nMinPage = tDialog.iMinPage
        .nMaxPage = tDialog.iMaxPage
        
        If Not (tDialog.oDeviceMode Is Nothing Or CBool(tDialog.iFlags And PD_RETURNDEFAULT)) Then
            hDevMode = pGetDevMode(tDialog.oDeviceMode)
            .hDevMode = hDevMode
        End If
        
        If Not (tDialog.oDeviceNames Is Nothing Or CBool(tDialog.iFlags And PD_RETURNDEFAULT)) Then
            hDevNames = pGetDevNames(tDialog.oDeviceNames)
            .hDevNames = hDevNames
        End If
        
        .hInstance = ZeroL
        .lCustData = ZeroL
        .lpfnPrintHook = ZeroL
        .lpfnSetupHook = ZeroL
        .lpPrintTemplateName = ZeroL
        .lpSetupTemplateName = ZeroL
        .hPrintTemplate = ZeroL
        .hSetupTemplate = ZeroL
        
        .hdc = ZeroL
        
        .lpfnPrintHook = ZeroL
        .lpfnSetupHook = ZeroL
        
        liHookProc = pGetHook(.Flags, PD_ENABLEPRINTHOOK Or PD_ENABLESETUPHOOK, dlgTypePrint, tDialog.oHookCallback)
        
        If CBool(.Flags And PD_ENABLEPRINTHOOK) Then .lpfnPrintHook = liHookProc
        If CBool(.Flags And PD_ENABLESETUPHOOK) Then .lpfnSetupHook = liHookProc
                
    End With

    'Show the dialog
    If PRINTDLG(ltPrintDialog) Then                                         'if success
        Print_ShowIndirect = True
        With ltPrintDialog                                                  'Return dialog values in parameters
            tDialog.hdc = .hdc
            tDialog.iReturnFlags = .Flags
            If CBool(.Flags And PD_PAGENUMS) Then
                tDialog.iRange = PD_PAGENUMS
            ElseIf CBool(.Flags And PD_SELECTION) Then
                tDialog.iRange = PD_SELECTION
            Else
                tDialog.iRange = ZeroL
            End If
            tDialog.iFromPage = .nFromPage
            tDialog.iToPage = .nToPage
            
            If tDialog.oDeviceMode Is Nothing Then Set tDialog.oDeviceMode = New cDeviceMode
            pSetDevMode tDialog.oDeviceMode, .hDevMode
            
            If tDialog.oDeviceNames Is Nothing Then Set tDialog.oDeviceNames = New cDeviceNames
            pSetDevNames tDialog.oDeviceNames, .hDevNames
            
        End With
    Else
        
        tDialog.iReturnExtendedError = CommDlgExtendedError()                               'return the extended error
        
    End If
    
    With ltPrintDialog
        If .hDevNames <> hDevNames And .hDevNames <> ZeroL Then GlobalFree .hDevMode        'free the return devnames
            If hDevNames Then GlobalFree hDevNames                                              'free the allocated devnames
                If .hDevMode <> hDevMode And .hDevMode <> ZeroL Then GlobalFree .hDevMode           'free the return devmode
                    If hDevMode Then GlobalFree hDevMode                                                'free the allocated devmode
                    End With
    
                    pReleaseHook liHookProc                                                                 'free the hook procedure, if any
    
End Function


Public Function PageSetup_ShowIndirect(ByRef tDialog As tPageSetupDialog) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Show the page setup dialog.
    '---------------------------------------------------------------------------------------

    Const ValidFlags As Long = _
    PSD_DEFAULTMINMARGINS Or _
    PSD_MINMARGINS Or _
    PSD_MARGINS Or _
    PSD_DISABLEMARGINS Or _
    PSD_DISABLEPRINTER Or _
    PSD_NOWARNING Or _
    PSD_DISABLEORIENTATION Or _
    PSD_RETURNDEFAULT Or _
    PSD_DISABLEPAPER Or _
    PSD_SHOWHELP Or _
    PSD_ENABLEPAGESETUPHOOK Or _
    PSD_DISABLEPAGEPAINTING

    Dim ltPageSetup      As PAGESETUPDLG
    Dim hDevMode         As Long
    Dim hDevNames        As Long
    
    tDialog.iReturnExtendedError = ZeroL                                                        'initialize return extended error

    ' Fill in PRINTDLG structure
    With ltPageSetup
        .lStructSize = Len(ltPageSetup)                                                         'initialize the api structure
        .Flags = (tDialog.iFlags And ValidFlags)
        .hWndOwner = tDialog.hWndOwner

        Dim liUnits      As Long
        
        If tDialog.iUnits = PSD_UNITS_Millimeters Then                                          'initialize the scale factor
            liUnits = 100
            .Flags = .Flags Or PSD_INHUNDREDTHSOFMILLIMETERS
        Else
            liUnits = 1000
            .Flags = .Flags Or PSD_INTHOUSANDTHSOFINCHES
        End If
        
        With .rtMargin
            .Top = tDialog.fTopMargin * liUnits                                                  'set the default margins
            .Left = tDialog.fLeftMargin * liUnits
            .Bottom = tDialog.fBottomMargin * liUnits
            .Right = tDialog.fRightMargin * liUnits
        End With

        With .rtMinMargin
            .Top = tDialog.fMinTopMargin * liUnits                                               'set the default min margins
            .Left = tDialog.fMinLeftMargin * liUnits
            .Bottom = tDialog.fMinBottomMargin * liUnits
            .Right = tDialog.fMinRightMargin * liUnits
        End With
        
        .lpfnPageSetupHook = pGetHook(.Flags, PSD_ENABLEPAGESETUPHOOK, dlgTypePageSetup, tDialog.oHookCallback)
        
        If Not (tDialog.oDeviceMode Is Nothing Or CBool(.Flags And PSD_RETURNDEFAULT)) Then
            hDevMode = pGetDevMode(tDialog.oDeviceMode)                                         'init the devmode
            .hDevMode = hDevMode
        End If
        
        If Not (tDialog.oDeviceNames Is Nothing Or CBool(.Flags And PSD_RETURNDEFAULT)) Then    'init the devnames
            hDevNames = pGetDevNames(tDialog.oDeviceNames)
            .hDevNames = hDevNames
        End If
        
        ZeroMemory .ptPaperSize, LenB(.ptPaperSize)                                              'make it obvious that we don't use these members
        .hInstance = ZeroL
        .lCustData = ZeroL
        .lpfnPagePaintHook = ZeroL
        .lpPageSetupTemplateName = ZeroL
        .hPageSetupTemplate = ZeroL
        
    End With
    
    ' Show Print dialog
    If PAGESETUPDLG(ltPageSetup) Then                                                           'if dialog succeeded
        PageSetup_ShowIndirect = True
        ' Return dialog values in parameters
        With ltPageSetup
            With .rtMargin                                                                      'return the selected margins
                tDialog.fTopMargin = .Top / liUnits
                tDialog.fLeftMargin = .Left / liUnits
                tDialog.fBottomMargin = .Bottom / liUnits
                tDialog.fRightMargin = .Right / liUnits
            End With
            
            If tDialog.oDeviceMode Is Nothing Then Set tDialog.oDeviceMode = New cDeviceMode
            pSetDevMode tDialog.oDeviceMode, .hDevMode                                          'return the devmode
            
            If tDialog.oDeviceNames Is Nothing Then Set tDialog.oDeviceNames = New cDeviceNames
            pSetDevNames tDialog.oDeviceNames, .hDevNames                                       'return the devnames
            
        End With
        
    Else                                                                                        'if dialog failed
        tDialog.iReturnExtendedError = CommDlgExtendedError()                                   'return the extended error
    End If
    
    If ltPageSetup.hDevMode Then GlobalFree ltPageSetup.hDevMode                                'free the allocated memory
        If ltPageSetup.hDevNames Then GlobalFree ltPageSetup.hDevNames
        If hDevMode <> ZeroL And hDevMode <> ltPageSetup.hDevMode Then GlobalFree hDevMode
        If hDevNames <> ZeroL And hDevNames <> ltPageSetup.hDevNames Then GlobalFree hDevNames
    
        pReleaseHook ltPageSetup.lpfnPageSetupHook
    
End Function
    

Private Function pGetDevMode(ByVal oDevMode As cDeviceMode) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Allocate a DEVMODE structure and copy the data from the cDeviceMode object into it.
    '---------------------------------------------------------------------------------------
    
    pGetDevMode = GlobalAlloc(GMEM_MOVEABLE, oDevMode.fSizeOf)
    If pGetDevMode Then
        Dim lhMem      As Long
        lhMem = GlobalLock(pGetDevMode)
        If lhMem Then
            CopyMemory ByVal lhMem, ByVal oDevMode.lpDevMode, oDevMode.fSizeOf
            GlobalUnlock pGetDevMode
        End If
    End If
End Function

Private Sub pSetDevMode(ByVal oDevMode As cDeviceMode, ByVal hDevMode As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Copy the data from an allocated DEVMODE structure into a cDeviceMode Class.
    '---------------------------------------------------------------------------------------
    If hDevMode Then
        Dim lhMem      As Long
        lhMem = GlobalLock(hDevMode)
        If lhMem Then
            CopyMemory ByVal oDevMode.lpDevMode, ByVal lhMem, oDevMode.fSizeOf
            GlobalUnlock hDevMode
            oDevMode.fChanged
        End If
    End If
End Sub

Private Function pGetDevNames(ByVal oDeviceNames As cDeviceNames) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Allocate a DEVNAMES structure and copy the data from the cDeviceNames class into it.
    '---------------------------------------------------------------------------------------

    Dim lsDriver      As String: lsDriver = oDeviceNames.DriverName
    Dim lsDevice      As String: lsDevice = oDeviceNames.DeviceName
    Dim lsOutput      As String: lsOutput = oDeviceNames.OutputPort
    
    If LenB(lsDriver) Or LenB(lsDevice) Or LenB(lsOutput) Then
        
        Dim ltDevNames As DEVNAMES
        pGetDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(lsDriver) + Len(lsDevice) + Len(lsOutput) + Len(ltDevNames) + 6 + 16)
        If pGetDevNames Then
            Dim lhMem      As Long
            lhMem = GlobalLock(pGetDevNames)
            If lhMem Then
                With ltDevNames
                    .wDriverOffset = Len(ltDevNames)
                    .wDeviceOffset = .wDriverOffset + Len(lsDriver) + 2
                    .wOutputOffset = .wDeviceOffset + Len(lsDevice) + 2
                End With
                CopyMemory ByVal lhMem, ltDevNames, Len(ltDevNames)
                
                Dim lsAnsi      As String
                
                lsAnsi = StrConv(lsDriver & vbNullChar, vbFromUnicode)
                CopyMemory ByVal UnsignedAdd(lhMem, ltDevNames.wDriverOffset), ByVal StrPtr(lsAnsi), LenB(lsAnsi)
                
                lsAnsi = StrConv(lsDevice & vbNullChar, vbFromUnicode)
                CopyMemory ByVal UnsignedAdd(lhMem, ltDevNames.wDriverOffset), ByVal StrPtr(lsAnsi), LenB(lsAnsi)
                
                lsAnsi = StrConv(lsOutput & vbNullChar, vbFromUnicode)
                CopyMemory ByVal UnsignedAdd(lhMem, ltDevNames.wDriverOffset), ByVal StrPtr(lsAnsi), LenB(lsAnsi)
                
                GlobalUnlock pGetDevNames
            End If
        End If
    End If
End Function

Private Sub pSetDevNames(ByVal oDevNames As cDeviceNames, ByVal hDevNames As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Copy the data from an allocated DEVNAMES structure into a cDeviceNames class.
    '---------------------------------------------------------------------------------------
    If hDevNames Then
        Dim lhMem           As Long
        Dim ltDevNames      As DEVNAMES
        lhMem = GlobalLock(hDevNames)
        If lhMem Then
            CopyMemory ltDevNames, ByVal lhMem, LenB(ltDevNames)
            oDevNames.fInit ltDevNames.wDefault And DN_DEFAULTPRN, _
            lstrToStringAFunc(UnsignedAdd(lhMem, ltDevNames.wDriverOffset)), _
            lstrToStringAFunc(UnsignedAdd(lhMem, ltDevNames.wDeviceOffset)), _
            lstrToStringAFunc(UnsignedAdd(lhMem, ltDevNames.wOutputOffset))
            GlobalUnlock lhMem
        End If
    End If
End Sub

Public Function Help_Show( _
ByRef sFile As String, _
ByVal iCommand As Long, _
Optional ByVal vTopicNameOrId As Variant, _
Optional ByVal hWnd As Long) _
As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Show the help dialog using arguments instead of a udt.
    '---------------------------------------------------------------------------------------
    Dim ltDialog      As tHelpDialog
    With ltDialog
        .sFile = sFile
        .vTopic = vTopicNameOrId
        .iCommand = iCommand
        .hWnd = hWnd
    End With
    Help_Show = Help_ShowIndirect(ltDialog)
End Function
    
Public Function Help_ShowIndirect(ByRef tDialog As tHelpDialog) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Show the help dialog.
    '---------------------------------------------------------------------------------------
    With tDialog
        
        Dim lsFile      As String:   lsFile = StrConv(tDialog.sFile & vbNullChar, vbFromUnicode)
        Dim lpFile      As Long:     lpFile = StrPtr(lsFile)
        
        Dim lsAnsi As String
        
        Select Case .iCommand
        Case HH_HELP_CONTEXT
            If VarType(.vTopic) = vbLong Or VarType(.vTopic) = vbInteger Then
                Help_ShowIndirect = CBool(HtmlHelp(.hWnd, ByVal lpFile, .iCommand, ByVal CLng(tDialog.vTopic)))
            Else
                lsAnsi = StrConv(tDialog.vTopic & vbNullChar, vbFromUnicode)
                Help_ShowIndirect = CBool(HtmlHelp(.hWnd, ByVal lpFile, .iCommand, ByVal StrPtr(lsAnsi)))
            End If
        Case HH_DISPLAY_SEARCH
            Dim ltSearch      As HH_FTS_QUERY
            ltSearch.cbStruct = LenB(ltSearch)
            Help_ShowIndirect = CBool(HtmlHelp(.hWnd, ByVal lpFile, .iCommand, ltSearch))
        Case HH_DISPLAY_TOC
            Help_ShowIndirect = CBool(HtmlHelp(.hWnd, ByVal lpFile, .iCommand, ByVal ZeroL))
        Case HH_DISPLAY_INDEX
            If Len(tDialog.vTopic) Then
                lsAnsi = StrConv(.vTopic & vbNullChar, vbFromUnicode)
                Help_ShowIndirect = CBool(HtmlHelp(.hWnd, ByVal lpFile, .iCommand, ByVal StrPtr(lsAnsi)))
            Else
                Help_ShowIndirect = CBool(HtmlHelp(.hWnd, ByVal lpFile, .iCommand, ByVal ZeroL))
            End If
        Case HH_CLOSE_ALL
            Help_ShowIndirect = CBool(HtmlHelp(.hWnd, ByVal ZeroL, .iCommand, ByVal ZeroL))
        End Select
    End With
End Function


Public Sub Dialog_CenterWindow(ByVal hWnd As Long, ByRef vCenterTo As Variant)
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Center a dialog window relative to a form, hWnd or a Screen object.
    '---------------------------------------------------------------------------------------
    Dim hWndCenter          As Long
    Dim ltRectDialog        As RECT
    Dim ltRectCenterTo      As RECT
    Dim ltRectWorkArea      As RECT
    
    If (SystemParametersInfo(SPI_GETWORKAREA, ZeroL, ltRectWorkArea, ZeroL) = ZeroL) Then
        ' Call failed - just use standard screen:
        With Screen
            ltRectWorkArea.Right = .Width \ .TwipsPerPixelX
            ltRectWorkArea.Bottom = .Height \ .TwipsPerPixelY
        End With
    End If

    If GetWindowRect(hWnd, ltRectDialog) Then
        If VarType(vCenterTo) = vbObject Then
            If TypeOf vCenterTo Is Screen Then
                LSet ltRectCenterTo = ltRectWorkArea
            Else
                On Error Resume Next
                hWndCenter = vCenterTo.hWnd
                On Error GoTo 0
                If GetWindowRect(hWndCenter, ltRectCenterTo) = ZeroL Then LSet ltRectCenterTo = ltRectWorkArea
            End If
        ElseIf VarType(vCenterTo) = vbLong Then
            hWndCenter = vCenterTo
            If GetWindowRect(hWndCenter, ltRectCenterTo) = ZeroL Then LSet ltRectCenterTo = ltRectWorkArea
        Else
            LSet ltRectCenterTo = ltRectWorkArea
        End If
        
        Dim X       As Long
        Dim Y       As Long
        Dim cx      As Long
        Dim cy      As Long
        
        With ltRectCenterTo
            X = .Left + ((.Right - .Left) \ TwoL)
            Y = .Top + ((.Bottom - .Top) \ TwoL)
        End With
        
        With ltRectDialog
            cx = .Right - .Left
            cy = .Bottom - .Top
        End With
        
        X = X - (cx \ TwoL)
        Y = Y - (cy \ TwoL)
        
        With ltRectWorkArea
            If X + cx > .Right Then X = .Right - cx
            If Y + cy > .Bottom Then Y = .Bottom - cy
            If X < .Left Then X = .Left
            If Y < .Top Then Y = .Top
        End With
        
        MoveWindow hWnd, X, Y, cx, cy, OneL
    End If
End Sub


Private Sub StringFromBytes(ByRef yBytes() As Byte, ByRef sString As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Translate ANSI bytes to a unicode string.
    '---------------------------------------------------------------------------------------
    sString = StrConv(yBytes, vbUnicode)
    Dim i      As Long
    i = InStr(OneL, sString, vbNullChar)
    If i Then sString = Left$(sString, i - OneL)
End Sub

Private Sub BytesFromString(ByRef yBytes() As Byte, ByRef sString As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 8/31/05
    ' Purpose   : Translate ANSI bytes to a unicode string.
    '---------------------------------------------------------------------------------------
    Dim liLBound      As Long
    Dim liLen         As Long
    liLBound = LBound(yBytes)
    liLen = UBound(yBytes) - liLBound + OneL
    
    Dim lsAnsi      As String
    lsAnsi = StrConv(sString, vbFromUnicode)
    
    Dim liStringLen      As Long
    liStringLen = LenB(lsAnsi)
    
    If liStringLen > liLen Then liStringLen = liLen
    
    If liStringLen > ZeroL Then CopyMemory yBytes(liLBound), ByVal StrPtr(lsAnsi), liStringLen
    If liStringLen < liLen _
        Then ZeroMemory yBytes(liStringLen + liLBound), liLen - liStringLen _
    Else yBytes(liLen - liLBound - OneL) = ZeroY
End Sub






'Public Function File_ShowSave( _
'   Optional ByRef sReturnFileName As String, _
'   Optional ByVal iFlags As Long = OFN_EXPLORER, _
'   Optional ByVal sFilter As String = "All Files (*.*)|*.*", _
'   Optional ByVal iFilterIndex As Long = 1, _
'   Optional ByVal sDefaultExt As String, _
'   Optional ByVal sInitPath As String, _
'   Optional ByVal sInitFile As String, _
'   Optional ByVal sTitle As String, _
'   Optional ByVal hWndOwner As Long, _
'   Optional ByVal vTemplate As Variant, _
'   Optional ByVal hInstance As Long, _
'   Optional ByRef iReturnFlags As Long, _
'   Optional ByRef iReturnExtendedError As Long, _
'   Optional ByRef iReturnFilterIndex As Long, _
'   Optional ByVal oHookCallback As iComDlgHook) _
'                As Boolean
''---------------------------------------------------------------------------------------
'' Date      : 8/31/05
'' Purpose   : Show the file save dialog using arguments instead of a udt.
''---------------------------------------------------------------------------------------
'
'    'get the udt
'    Dim ltDialog As tFileDialog
'    pFile_GetUDT ltDialog, sTitle, iFlags, sFilter, iFilterIndex, sDefaultExt, sInitPath, sInitFile, hWndOwner, vTemplate, hInstance, oHookCallback
'
'    'show the dialog
'    File_ShowSave = pFile_Show(ltDialog, True)
'
'    'return the out parameters
'    With ltDialog
'        iReturnExtendedError = .iReturnExtendedError
'        iReturnFilterIndex = .iReturnFilterIndex
'        iReturnFlags = .iReturnFlags
'        sReturnFileName = .sReturnFileName
'    End With
'
'End Function
'
'Private Sub pFile_GetUDT( _
'            ByRef tDialog As tFileDialog, _
'            ByRef sTitle As String, _
'            ByVal iFlags As Long, _
'            ByRef sFilter As String, _
'            ByVal iFilterIndex As Long, _
'            ByRef sDefaultExt As String, _
'            ByRef sInitPath As String, _
'            ByRef sInitFile As String, _
'            ByVal hWndOwner As Long, _
'            ByRef vTemplate As Variant, _
'            ByVal hInstance As Long, _
'            ByVal oHookCallback As iComDlgHook)
''---------------------------------------------------------------------------------------
'' Date      : 8/31/05
'' Purpose   : Fill a file dialog udt with the given arguments.
''---------------------------------------------------------------------------------------
'
'    With tDialog
'        .iFlags = iFlags
'        .hWndOwner = hWndOwner
'        .sTitle = CStr(sTitle)
'        .sInitFile = CStr(sInitFile)
'        .sInitPath = CStr(sInitPath)
'        .sDefaultExt = CStr(sDefaultExt)
'        .hInstance = hInstance
'        .vTemplate = vTemplate
'        .sFilter = CStr(sFilter)
'        .iFilterIndex = iFilterIndex
'
'        Set .oHookCallback = oHookCallback
'
'    End With
'
'End Sub
'
'Public Function File_ShowOpen( _
'   Optional ByRef sReturnFileName As String, _
'   Optional ByVal iFlags As Long = OFN_EXPLORER Or OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY, _
'   Optional ByVal sFilter As String = "All Files (*.*)|*.*", _
'   Optional ByVal iFilterIndex As Long = 1, _
'   Optional ByVal sDefaultExt As String, _
'   Optional ByVal sInitPath As String, _
'   Optional ByVal sInitFile As String, _
'   Optional ByVal sTitle As String, _
'   Optional ByVal hWndOwner As Long, _
'   Optional ByVal vTemplate As Variant, _
'   Optional ByVal hInstance As Long, _
'   Optional ByRef iReturnFlags As Long, _
'   Optional ByRef iReturnExtendedError As Long, _
'   Optional ByRef iReturnFilterIndex As Long, _
'   Optional ByVal oHookCallback As iComDlgHook) _
'                As Boolean
''---------------------------------------------------------------------------------------
'' Date      : 8/31/05
'' Purpose   : Show the file open dialog using arguments instead of a udt.
''---------------------------------------------------------------------------------------
'
'    'get the udt
'    Dim ltDialog As tFileDialog
'    pFile_GetUDT ltDialog, sTitle, iFlags, sFilter, iFilterIndex, sDefaultExt, sInitPath, sInitFile, hWndOwner, vTemplate, hInstance, oHookCallback
'
'    'show the dialog
'    File_ShowOpen = pFile_Show(ltDialog, False)
'
'    'return the out parameters
'    With ltDialog
'        iReturnExtendedError = .iReturnExtendedError
'        iReturnFilterIndex = .iReturnFilterIndex
'        iReturnFlags = .iReturnFlags
'        sReturnFileName = .sReturnFileName
'    End With
'
'End Function
'
'Public Function Folder_Show( _
'   Optional ByRef sReturnPath As String, _
'   Optional ByVal iFlags As Long = BIF_USENEWUI Or BIF_STATUSTEXT, _
'   Optional ByVal sTitle As String, _
'   Optional ByVal sInitialPath As String, _
'   Optional ByVal sRootPath As String, _
'   Optional ByVal hWndOwner As Long, _
'   Optional ByVal oHookCallback As iComDlgHook) _
'        As Boolean
''---------------------------------------------------------------------------------------
'' Date      : 8/31/05
'' Purpose   : Show the folder dialog using arguments instead of a udt.
''---------------------------------------------------------------------------------------
'
'    Dim tFolder As tFolderDialog
'    With tFolder
'        .iFlags = iFlags
'        .sTitle = sTitle
'        .sInitialPath = sInitialPath
'        .sRootPath = sRootPath
'        .hWndOwner = hWndOwner
'        Set .oHookCallback = oHookCallback
'    End With
'
'    Folder_Show = Folder_ShowIndirect(tFolder)
'
'    sReturnPath = tFolder.sReturnPath
'
'End Function
'
'Public Function Color_Show( _
'                ByRef iColor As Long, _
'       Optional ByVal iFlags As Long = CC_ANYCOLOR, _
'       Optional ByVal hWndOwner As Long, _
'       Optional ByRef vColors As Variant, _
'       Optional ByRef iReturnExtendedError As Long, _
'       Optional ByVal oHookCallback As iComDlgHook) _
'                    As Boolean
''---------------------------------------------------------------------------------------
'' Date      : 8/31/05
'' Purpose   : Show the color dialog using arguments instead of a udt.
''---------------------------------------------------------------------------------------
'
'    Dim ltDialog As tColorDialog
'    With ltDialog
'        .iColor = iColor                            'fill the udt
'        .iFlags = iFlags
'        .hWndOwner = hWndOwner
'        Set .oHookCallback = oHookCallback
'
'        Dim i As Long
'        If IsArray(vColors) Then
'            On Error Resume Next
'            For i = 0 To 15
'                .iColors(i) = vColors(i)
'            Next
'            On Error GoTo 0
'        Else
'            For i = 0 To 15
'                .iColors(i) = QBColor(i)
'            Next
'        End If
'
'        Color_Show = Color_ShowIndirect(ltDialog)   'show the dialog
'
'        If Color_Show Then
'            iColor = .iColor                        'return the info from the dialog
'
'            If IsArray(vColors) Then
'                On Error Resume Next
'                For i = 0 To 15
'                    vColors(i) = .iColors(i)
'                Next
'                On Error GoTo 0
'            Else
'                vColors = .iColors
'            End If
'
'        End If
'
'        iReturnExtendedError = .iReturnExtendedError
'
'    End With
'
'End Function
'
'Public Function Font_Show( _
'            Optional ByVal oFont As Object, _
'            Optional ByVal iFlags As Long = CF_SCREENFONTS, _
'            Optional ByVal hDc As Long, _
'            Optional ByVal hWndOwner As Long, _
'            Optional ByVal iMinSize As Long = 6, _
'            Optional ByVal iMaxSize As Long = 72, _
'            Optional ByRef iColor As Long, _
'            Optional ByRef iReturnFlags As Long, _
'            Optional ByRef iReturnExtendedError As Long, _
'            Optional ByVal oHookCallback As iComDlgHook) _
'                As Boolean
''---------------------------------------------------------------------------------------
'' Date      : 8/31/05
'' Purpose   : Show the font dialog using arguments instead of a udt.
''---------------------------------------------------------------------------------------
'
'    Dim ltDialog As tFontDialog
'    With ltDialog
'        Set .oHookCallback = oHookCallback
'        Set .oFont = oFont
'        .iFlags = iFlags
'        .hDc = hDc
'        .hWndOwner = hWndOwner
'        .iMinSize = iMinSize
'        .iMaxSize = iMaxSize
'        .iColor = iColor
'    End With
'
'    Font_Show = Font_ShowIndirect(ltDialog)
'
'    If Font_Show Then
'        iColor = ltDialog.iColor
'        iReturnFlags = ltDialog.iReturnFlags
'    End If
'
'    Set oFont = ltDialog.oFont
'
'    iReturnExtendedError = ltDialog.iReturnExtendedError
'
'End Function
'
'Public Function Print_Show( _
'            Optional ByRef hDc As Long, _
'            Optional ByVal iFlags As Long, _
'            Optional ByVal hWndOwner As Long, _
'            Optional ByRef iRange As Long, _
'            Optional ByRef iFromPage As Long = 1, _
'            Optional ByRef iToPage As Long = 1, _
'            Optional ByVal iMinPage As Long = 1, _
'            Optional ByVal iMaxPage As Long = 1, _
'            Optional ByRef oDeviceMode As cDeviceMode, _
'            Optional ByRef oDeviceNames As cDeviceNames, _
'            Optional ByRef iReturnFlags As Long, _
'            Optional ByRef iReturnExtendedError As Long, _
'            Optional ByVal oHookCallback As iComDlgHook) _
'                As Boolean
''---------------------------------------------------------------------------------------
'' Date      : 8/31/05
'' Purpose   : Show the print dialog using arguments instead of a udt.
''---------------------------------------------------------------------------------------
'
'    Dim ltDialog As tPrintDialog
'
'    With ltDialog
'        .hWndOwner = hWndOwner
'        .iFlags = iFlags
'        .iRange = iRange
'        .iMinPage = iMinPage
'        .iMaxPage = iMaxPage
'
'        Set .oDeviceMode = oDeviceMode
'        Set .oDeviceNames = oDeviceNames
'
'        .iFromPage = iFromPage
'        .iToPage = iToPage
'        Set .oHookCallback = oHookCallback
'
'        Print_Show = Print_ShowIndirect(ltDialog)
'
'        If Print_Show Then
'            hDc = .hDc
'            iRange = .iRange
'            iFromPage = .iFromPage
'            iToPage = .iToPage
'            iReturnFlags = .iReturnFlags
'            iReturnExtendedError = .iReturnExtendedError
'
'            Set oDeviceMode = .oDeviceMode
'            Set oDeviceNames = .oDeviceNames
'
'        End If
'
'    End With
'
'End Function
'
'Public Function PageSetup_Show( _
'    Optional ByVal iFlags As Long, _
'    Optional ByVal iUnits As Long = PSD_UNITS_Inches, _
'    Optional ByRef fLeftMargin As Single = 1, _
'    Optional ByVal fMinLeftMargin As Single = 1, _
'    Optional ByRef fRightMargin As Single = 1, _
'    Optional ByVal fMinRightMargin As Single = 1, _
'    Optional ByRef fTopMargin As Single = 0.5, _
'    Optional ByVal fMinTopMargin As Single = 0.5, _
'    Optional ByRef fBottomMargin As Single = 0.5, _
'    Optional ByVal fMinBottomMargin As Single = 0.5, _
'    Optional ByRef oDeviceMode As cDeviceMode, _
'    Optional ByRef oDeviceNames As cDeviceNames, _
'    Optional ByVal hWndOwner As Long, _
'    Optional ByRef iReturnExtendedError As Long, _
'    Optional ByVal oHookCallback As iComDlgHook) _
'        As Boolean
''---------------------------------------------------------------------------------------
'' Date      : 8/31/05
'' Purpose   : Show the page setup dialog using arguments instead of a udt.
''---------------------------------------------------------------------------------------
'
'    Dim ltDialog As tPageSetupDialog
'
'    With ltDialog
'        .iFlags = iFlags
'        .hWndOwner = hWndOwner
'        .iUnits = iUnits
'        .fLeftMargin = fLeftMargin
'        .fMinLeftMargin = fMinLeftMargin
'        .fRightMargin = fRightMargin
'        .fMinRightMargin = fMinRightMargin
'        .fTopMargin = fTopMargin
'        .fMinTopMargin = fMinTopMargin
'        .fBottomMargin = fBottomMargin
'        .fMinBottomMargin = fMinBottomMargin
'
'        Set .oDeviceMode = oDeviceMode
'        Set .oDeviceNames = oDeviceNames
'        Set .oHookCallback = oHookCallback
'
'        PageSetup_Show = PageSetup_ShowIndirect(ltDialog)
'
'        If PageSetup_Show Then
'            fLeftMargin = .fLeftMargin
'            fRightMargin = .fRightMargin
'            fBottomMargin = .fBottomMargin
'            fTopMargin = .fTopMargin
'            iReturnExtendedError = .iReturnExtendedError
'
'            Set oDeviceMode = .oDeviceMode
'            Set oDeviceNames = .oDeviceNames
'
'        End If
'
'    End With
'
'End Function


