Attribute VB_Name = "mComCtlShared"
'==================================================================================================
'mComCtlShared.bas        9/2/05
'
'           PURPOSE:
'               Shared declarations and procedures used throughout the project.
'
'==================================================================================================

Option Explicit

Public Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Public Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As Any, pClipRect As Any) As Long
Public Declare Function DrawThemeText Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As Long, ByVal iCharCount As Long, ByVal dwTextFlag As Long, ByVal dwTextFlags2 As Long, pRect As RECT) As Long
Public Declare Function GetThemeTextExtent Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As Long, ByVal iCharCount As Long, ByVal dwTextFlags As Long, pBoundingRect As RECT, pExtentRect As RECT) As Long
Public Declare Function IsAppThemedApi Lib "uxtheme.dll" Alias "IsAppThemed" () As Long
Public Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Public Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszSubAppName As Long, ByVal pszSubIdList As Long) As Long

Public Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal hdcDst As Long, ByRef pptDst As Any, ByRef psize As Any, ByVal hdcSrc As Long, ByRef pptSrc As Any, ByVal crKey As Long, ByRef pblend As Any, ByVal dwFlags As Long) As Long

Public Const UM_SIZEBAND As Long = WM_USER + &H66BA&

Public Function TranslateColor(ByVal clr As OLE_COLOR, _
Optional hPal As Long = 0) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : Get rgb color from ole color
    '---------------------------------------------------------------------------------------
    If OleTranslateColor(clr, hPal, TranslateColor) Then TranslateColor = NegOneL
End Function

Public Function RoundToInterval(ByVal iNumber As Long, Optional ByVal iInterval As Long = 8) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : Round a number to the nearest multiple of a given interval.
    '---------------------------------------------------------------------------------------
    
    'if iNumber < 0 then iNumber = 0
    If CBool(iNumber And &H80000000) Then iNumber = ZeroL
    
    Dim liMod      As Long
    liMod = iNumber Mod iInterval
    
    If Not (liMod = 0) Then
        'If the number is not an even multiple, then round it up
        RoundToInterval = iNumber + iInterval - liMod
    Else
        'If it is an even multiple then keep it the same,
        'unless it's zero then make it equal to iInterval
        If Not (iNumber = ZeroL) _
            Then RoundToInterval = iNumber _
        Else RoundToInterval = iInterval
        End If
End Function

Public Sub gErr(ByVal iNum As evbComCtlError, ByRef sSource As String, Optional ByRef sDesc As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : Raise the specified error.
    '---------------------------------------------------------------------------------------
    On Error Resume Next
    Dim lsDesc      As String
    If LenB(sDesc) = ZeroL Then
        Select Case iNum
        Case vbccLetSetNoRunTime
            lsDesc = "Property Let/Set not supported at run time."
        Case vbccLetSetNoDesignTime
            lsDesc = "Property Let/Set not supported at design time."
        Case vbccGetNoRunTime
            lsDesc = "Property Get not supported at run time."
        Case vbccGetNoDesignTime
            lsDesc = "Property Get not supported at design time."
        Case vbccKeyOrIndexNotFound
            lsDesc = "Specified key was not found."
        Case vbccKeyAlreadyExists
            lsDesc = "Specified key already exists."
        Case vbccItemDetached
            lsDesc = "Item has been detached from the collection."
        Case vbccCollectionChangedDuringEnum
            lsDesc = "Collection changed during enumeration."
        Case vbccInvalidProcedureCall
            lsDesc = "Invalid Procedure call."
        Case vbccOutOfMemory
            lsDesc = "Out of memory."
        Case vbccUnsupported
            lsDesc = "Functionality is not supported by the libraries available."
        Case vbccUserCanceled
            lsDesc = "User canceled the operation."
        Case vbccComDlgExtendedError
            lsDesc = "Extended common dialog error."
        Case vbccTypeMismatch
            lsDesc = "Type Mismatch."
        End Select
    Else
        lsDesc = sDesc
    End If
    Err.Raise iNum, sSource, lsDesc
End Sub
    
Public Sub lstrToStringA(ByVal iPtr As Long, ByRef sOut As String, Optional ByVal iLength As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : get a vb unicode string from an ansi lstr.
    '---------------------------------------------------------------------------------------
    
    'debug.assert iPtr
    
    If iPtr Then
        If iLength = ZeroL Then iLength = lstrlen(iPtr)
        
        sOut = Space$(MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, ByVal iPtr, iLength, ByVal ZeroL, ByVal ZeroL))
        MultiByteToWideChar CP_ACP, MB_PRECOMPOSED, ByVal iPtr, iLength, ByVal StrPtr(sOut), ByVal LenB(sOut)
        
    End If
    
End Sub

Public Sub lstrFromStringA(ByVal iPtr As Long, ByVal iPtrLen As Long, ByRef sIn As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : get an ansi lstr from a vb unicode string.
    '---------------------------------------------------------------------------------------

    Dim liLen       As Long
    Dim lsAnsi      As String
    
    'debug.assert iPtr
    
    If iPtr Then
        
        lsAnsi = StrConv(sIn, vbFromUnicode)
        liLen = LenB(lsAnsi)
        
        If liLen > iPtrLen - 1 Then liLen = iPtrLen - 1
        
        If liLen > ZeroL Then
            CopyMemory ByVal iPtr, ByVal StrPtr(lsAnsi), liLen
        End If
        MemByte(ByVal UnsignedAdd(iPtr, liLen)) = ZeroY
    End If
End Sub

Public Sub lstrFromStringW(ByVal iPtr As Long, ByVal iPtrLen As Long, ByRef sIn As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : get a unicode lstr from a vb unicode string.
    '---------------------------------------------------------------------------------------
    
    Dim liLen      As Long
    
    'debug.assert iPtr
    
    If iPtr Then
    
        liLen = LenB(sIn)
        iPtrLen = iPtrLen - TwoL
        
        If liLen > iPtrLen - 2 Then liLen = iPtrLen - 2
        
        If liLen > ZeroL Then
            CopyMemory ByVal iPtr, ByVal StrPtr(sIn), liLen
        End If
        MemWord(ByVal UnsignedAdd(iPtr, liLen)) = 0
        
    End If
End Sub

Public Sub lstrToStringW(ByVal iPtr As Long, ByRef sOut As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : get a vb string from a unicode lstr.
    '---------------------------------------------------------------------------------------
    Dim liLen      As Long
    
    'debug.assert iPtr
    
    If iPtr Then
        liLen = lstrlenW(iPtr)
        sOut = Space$(liLen)
        If liLen Then CopyMemory ByVal StrPtr(sOut), ByVal iPtr, liLen + liLen
    End If
    
End Sub

Public Function lstrToStringAFunc(ByVal lpString As Long) As String
    lstrToStringA lpString, lstrToStringAFunc
End Function

'Public Function lstrToStringWFunc(ByVal lpString As Long) As String
'    lstrToStringW lpString, lstrToStringWFunc
'End Function

Public Sub DateToSysTime(ByRef dDate As Date, ByRef tSysTime As SYSTEMTIME)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : Convert VB's date to a SYSTIME
    '---------------------------------------------------------------------------------------
    With tSysTime
        .wDay = Day(dDate)
        .wMonth = Month(dDate)
        .wYear = Year(dDate)
        .wHour = Hour(dDate)
        .wSecond = Second(dDate)
        .wMinute = Minute(dDate)
    End With
End Sub

Public Sub SysTimeToDate(ByRef dDate As Date, ByRef tSysTime As SYSTEMTIME)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : Convert SYSTIME to VB's date
    '---------------------------------------------------------------------------------------
    dDate = DateSerial(tSysTime.wYear, tSysTime.wMonth, tSysTime.wDay) + TimeSerial(tSysTime.wHour, tSysTime.wMinute, tSysTime.wSecond)
End Sub

Public Function KBState() As evbComCtlKeyboardState
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : Get a keyboard state value with no input
    '---------------------------------------------------------------------------------------
    If KeyIsDown(VK_SHIFT, False) Then KBState = KBState Or vbccShiftMask
    If KeyIsDown(VK_MENU, False) Then KBState = KBState Or vbccAltMask
    If KeyIsDown(VK_CONTROL, False) Then KBState = KBState Or vbccControlMask
End Function

Public Function TranslateContextMenuCoords(ByVal hWnd As Long, ByVal lParam As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : Extract client mouse coordinates from a WM_CONTEXTMENU wParam.
    '---------------------------------------------------------------------------------------
    If lParam = NegOneL Then
        Dim tR      As RECT
        If GetWindowRect(hWnd, tR) Then
            With tR
                lParam = (.Left + ((.Right - .Left) \ TwoL)) Or _
                (.Top + ((.Bottom - .Top) \ TwoL)) * &H10000
            End With
        End If
    End If
    
    tR.Left = loword(lParam)
    tR.Top = hiword(lParam)
    ScreenToClient hWnd, tR
    TranslateContextMenuCoords = (loword(tR.Top) * &H10000) Or loword(tR.Left)
End Function


Public Property Get NextItemId() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : Return an incrementing 32 bit value.
    '---------------------------------------------------------------------------------------
    Static iId As Long
  
    iId = iId + OneL
NextItemId = iId
If iId = &H7FFFFFFF Then
    iId = &H80000000
ElseIf iId = NegOneL Then
    iId = OneL
End If
End Property

Public Property Get NextItemIdShort() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : Return a positive incrementing 16 bit value.
    '---------------------------------------------------------------------------------------
    Static iId As Long
    iId = iId + OneL
NextItemIdShort = iId
If iId = &H7FFF& Then iId = ZeroL
End Property

Public Function RootParent(ByVal hWnd As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : Get the last ancestor of a given window.
    '---------------------------------------------------------------------------------------
    
    Dim lhWndTest      As Long
    
    RootParent = hWnd
    lhWndTest = GetParent(RootParent)
    
    Do Until lhWndTest = ZeroL
        RootParent = lhWndTest
        lhWndTest = GetParent(RootParent)
    Loop
    
End Function

Public Function AccelChar(ByRef sCaption As String) As Integer
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : Find the first character with a "&" before it.  Exlcude "&&" occurances.
    '---------------------------------------------------------------------------------------
    
    Dim liPos      As Long
    
    liPos = OneL
    Do
        liPos = InStr(liPos, sCaption, "&")
        If liPos > ZeroL Then
            AccelChar = Asc(UCase$(Mid$(sCaption, liPos + OneL, OneL)))
            If AccelChar <> vbKeyUp Then Exit Function
            liPos = liPos + TwoL
        End If
    Loop Until liPos = ZeroL
    
    AccelChar = ZeroL
    
End Function

Public Sub ForceWindowToShowAllUIStates(ByVal hWnd As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : force a vb window to show focus rectangle, alt key accelerators, etc.
    '---------------------------------------------------------------------------------------
    
    Const UIS_SET As Long = 1
    Const UIS_CLEAR As Long = 2
    
    Const UISF_HIDEACCEL As Long = &H2
    Const UISF_HIDEFOCUS As Long = &H1
    
    Const CLEAR_IT_ALL As Long = ((UISF_HIDEACCEL Or UISF_HIDEFOCUS) * &H10000) Or UIS_CLEAR
    
    SendMessage hWnd, WM_CHANGEUISTATE, CLEAR_IT_ALL, ZeroL
    SendMessage hWnd, WM_CHANGEUISTATE, UIS_SET, ZeroL
    
End Sub

Public Sub EnableWindowTheme(ByVal hWnd As Long, ByVal bEnable As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : Enable/Disable support for the default theme if available.
    '---------------------------------------------------------------------------------------
    If IsAppThemed() And hWnd <> ZeroL Then
        Dim lsString      As String * 1
        Dim lp            As Long
        If Not bEnable Then lp = StrPtr(lsString)
        SetWindowTheme hWnd, lp, lp
    End If
End Sub

Public Property Get SystemColorDepth() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : Return the default color depth for the system.
    '---------------------------------------------------------------------------------------
    Const BITSPIXEL = 12&
    
    Dim lhDc      As Long
    lhDc = CreateDisplayDC()
    If lhDc Then
        SystemColorDepth = GetDeviceCaps(lhDc, BITSPIXEL)
        DeleteDC lhDc
    End If
End Property

Public Function GetVirtKey(ByVal iCharCode As Integer) As Integer
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : Translate a character code into a VK_* value.
    '---------------------------------------------------------------------------------------
    If (GetVersion() And &H80000000) = ZeroL _
        Then GetVirtKey = VkKeyScanW(iCharCode) And &HFF& _
    Else GetVirtKey = VkKeyScan(iCharCode And &HFF&) And &HFF&
End Function

Public Property Get NextCommandId() As Integer
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : Return an incrementing value moving in multiples of 100.
    '---------------------------------------------------------------------------------------
    Static iId As Long
    If iId = ZeroL Then iId = 1000& Else iId = iId + 100&
NextCommandId = iId
If iId >= &H7FFFFFD0 Then iId = ZeroL
End Property

Public Function IsApiAvailable(ByRef sMod As String, ByRef sProc As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : Return a value indicating whether a function exists in a dll library.
    '---------------------------------------------------------------------------------------
    Dim lsAnsi      As String
    lsAnsi = StrConv(sMod & vbNullChar, vbFromUnicode)
    
    Dim hLib      As Long
    hLib = GetModuleHandle(ByVal StrPtr(lsAnsi))
    
    If hLib Then
        lsAnsi = StrConv(sProc & vbNullChar, vbFromUnicode)
        IsApiAvailable = CBool(GetProcAddress(hLib, ByVal StrPtr(lsAnsi)))
    End If
End Function

Public Property Get IsAppThemed() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : Return a value indicating whether a windows theme is enabled.
    '---------------------------------------------------------------------------------------
    Static bInit As Boolean, bUxThemeExists As Boolean
    If Not bInit Then
        bInit = True
        bUxThemeExists = IsApiAvailable("uxtheme.dll", "IsAppThemed")
    End If
    If bUxThemeExists Then IsAppThemed = CBool(IsAppThemedApi())
End Property

Public Function TickDiff(ByVal iTickCount As Long, ByVal iTickCountStored As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : Return the difference between the two tick counts, allowing for the unlikely
    '             possibility that the count has passed unsigned territory or wrappped back
    '             to zero.
    '---------------------------------------------------------------------------------------
    If iTickCount > iTickCountStored Then TickDiff = iTickCount - iTickCountStored Else TickDiff = iTickCountStored - iTickCount
End Function


Public Function GetTextExtentPoint32W(ByVal lhDc As Long, ByRef sText As String, ByVal iTextLen As Long, ByRef tSize As SIZE) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : Wrap the api call.
    '---------------------------------------------------------------------------------------
    Dim lsAnsi      As String
    lsAnsi = StrConv(sText & vbNullChar, vbFromUnicode)
    GetTextExtentPoint32W = GetTextExtentPoint32(lhDc, ByVal StrPtr(lsAnsi), iTextLen, tSize)
End Function

Public Function FindWindowExW(ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByRef sText1 As String, ByRef sText2 As String) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : Wrap the api call.
    '---------------------------------------------------------------------------------------
    Dim lsAnsi1      As String: If LenB(sText1) Then lsAnsi1 = StrConv(sText1 & vbNullChar, vbFromUnicode)
    Dim lsAnsi2      As String: If LenB(sText2) Then lsAnsi2 = StrConv(sText2 & vbNullChar, vbFromUnicode)
        
    FindWindowExW = FindWindowEx(hWnd1, hWnd2, ByVal StrPtr(lsAnsi1), ByVal StrPtr(lsAnsi2))
End Function

Public Sub SetWindowStyle(ByVal hWnd As Long, ByVal iStyleOr As Long, ByVal iStyleAndNot As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : Wrap the api call.
    '---------------------------------------------------------------------------------------
    Dim liStyle         As Long
    Dim liStyleNew      As Long
    
    liStyle = GetWindowLong(hWnd, GWL_STYLE)
    liStyleNew = ((liStyle And Not iStyleAndNot) Or iStyleOr)
    
    If liStyle Xor liStyleNew Then SetWindowLong hWnd, GWL_STYLE, liStyleNew
    
End Sub

Public Sub SetWindowStyleEx(ByVal hWnd As Long, ByVal iStyleOr As Long, ByVal iStyleAndNot As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : Wrap the api call.
    '---------------------------------------------------------------------------------------
    Dim liStyle         As Long
    Dim liStyleNew      As Long
    
    liStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
    liStyleNew = ((liStyle Or iStyleOr) And Not iStyleAndNot)
    
    If liStyle Xor liStyleNew Then SetWindowLong hWnd, GWL_EXSTYLE, liStyleNew
    
End Sub

Public Function GetDispId(ByVal oObject As Object, ByRef sMethod As String) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : Return a DISPID given a method name.
    '---------------------------------------------------------------------------------------
    Dim oIDispatch      As Interfaces.IDispatch
    Dim IID_Null        As vbComCtlTlb.Guid

    'get ref to OLE IDispatch interface
    Set oIDispatch = oObject

    'get DispatchID for method from IDispatch interface
    '(VB will throw an 'Object Doesn't Support Property Or Method' error on failure)
    oIDispatch.GetIDsOfNames IID_Null, StrConv(sMethod & vbNullChar, vbUnicode), 1, 0&, GetDispId
    
End Function

Public Function KeyIsDown( _
ByVal iVirtKey As Long, _
Optional ByVal bAsync As Boolean = True) _
As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : Return a value indicating whether the given key is pressed.
    '---------------------------------------------------------------------------------------
    If bAsync _
        Then KeyIsDown = CBool(GetAsyncKeyState(iVirtKey) And &H8000) _
    Else KeyIsDown = CBool(GetKeyState(iVirtKey) And &H8000)
        
End Function


'Public Function ExtractIcon(ByVal hIml As Long, ByVal iIndex As Long) As IPicture
''---------------------------------------------------------------------------------------
'' Date      : 9/2/05
'' Purpose   : Return a new picture object that contains a copy of the given icon.
''---------------------------------------------------------------------------------------
'    If hIml Then
'        Dim hIcon As Long
'        hIcon = ImageList_GetIcon(hIml, iIndex, ILD_TRANSPARENT)
'        If hIcon Then
'            Set ExtractIcon = IconToPicture(hIcon)
'        Else
'            'debug.assert False
'            Set ExtractIcon = New StdPicture
'        End If
'    End If
'End Function
'
'Private Function IconToPicture(ByVal hIcon As Long) As IPicture
''---------------------------------------------------------------------------------------
'' Date      : 9/2/05
'' Purpose   : Return a new picture object that contains a copy of the given icon.
''---------------------------------------------------------------------------------------
'
'    If hIcon = ZeroL Then Exit Function
'
'    Dim NewPic As Picture, PicConv As PICTDESC, IGuid As vbComCtlTlb.GUID
'
'    PicConv.cbSizeOfStruct = Len(PicConv)
'    PicConv.picType = vbPicTypeIcon
'    PicConv.hImage = hIcon
'
'    ' Fill in IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
'    With IGuid
'        .Data1 = &H7BF80980
'        .Data2 = &HBF32
'        .Data3 = &H101A
'        .Data4(0) = &H8B
'        .Data4(1) = &HBB
'        .Data4(2) = &H0
'        .Data4(3) = &HAA
'        .Data4(4) = &H0
'        .Data4(5) = &H30
'        .Data4(6) = &HC
'        .Data4(7) = &HAB
'    End With
'    OleCreatePictureIndirect PicConv, IGuid, True, NewPic
'
'    Set IconToPicture = NewPic
'
'End Function
