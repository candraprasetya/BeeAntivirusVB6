VERSION 5.00
Begin VB.UserControl ucTabStrip 
   Alignable       =   -1  'True
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2685
   HasDC           =   0   'False
   PropertyPages   =   "ucTabStrip.ctx":0000
   ScaleHeight     =   495
   ScaleWidth      =   2685
   ToolboxBitmap   =   "ucTabStrip.ctx":000D
End
Attribute VB_Name = "ucTabStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'ucTabStrip.ctl        12/15/04
'
'           PURPOSE:
'               Implement the comctl32 tabstrip control.
'
'           LINEAGE:
'               http://www.vbaccelerator.com/home/VB/Code/Controls/Tab_Controls/vbAccelerator_ComCtl32_Tab_Control/VB6_TabStrip_Control_Full_Source.asp
'               cTabCtrl.ctl
'
'==================================================================================================

Option Explicit

Implements iOleInPlaceActiveObjectVB
Implements iOleControlVB
Implements iSubclass

Const PROP_Font = "Font"
Const PROP_HotTrack = "HotTrack"
Const PROP_Buttons = "Buttons"
Const PROP_MultiLine = "MultiLine"
Const PROP_RightJustify = "RightJustify"
Const PROP_FlatSeparators = "FlatSeparators"
Const PROP_FlatButtons = "FlatButtons"
Const PROP_Themeable = "Themeable"

Const DEF_HotTrack = False
Const DEF_Buttons = False
Const DEF_MultiLine = False
Const DEF_RightJustify = False
Const DEF_FlatSeparators = False
Const DEF_FlatButtons = False
Const DEF_Themeable = True

Const NMHDR_code As Long = 8
Const NMHDR_hwndFrom As Long = 0

Private mhWnd  As Long

Private moImageList         As cImageList
Private WithEvents moImageListEvent As cImageList
Attribute moImageListEvent.VB_VarHelpID = -1

Private WithEvents moFont   As cFont
Attribute moFont.VB_VarHelpID = -1
Private WithEvents moFontPage As pcSupportFontPropPage
Attribute moFontPage.VB_VarHelpID = -1

Private WithEvents moMnemonics As pcMnemonics
Attribute moMnemonics.VB_VarHelpID = -1

Private mhFont              As Long

Private mbHotTrack          As Boolean
Private mbButtons           As Boolean
Private mbMultiLine         As Boolean
Private mbRightJustify      As Boolean
Private mbFlatSeparators    As Boolean
Private mbFlatButtons       As Boolean
Private mbThemeable         As Boolean

Private miTabControl        As Long
Private miWheelDelta        As Long

Private msTextBuffer        As String * 130

Private mtItem              As TCITEM

Public Event BeforeClick(ByRef bCancel As OLE_CANCELBOOL)
Public Event Click(ByVal oTab As cTab)
Public Event RightClick(ByVal oTab As cTab)

Private Const cTabs = "cTabs"
Private Const cTab = "cTab"
'Private Const ucTabStrip = "ucTabStrip"

Private Const UM_MouseDown As Long = WM_USER + &H66BF&

Private Sub iOleControlVB_GetControlInfo(bHandled As Boolean, iAccelCount As Long, hAccelTable As Long, iFlags As Long)
    bHandled = True
    iAccelCount = moMnemonics.Count
    hAccelTable = moMnemonics.hAccel
    iFlags = ZeroL
End Sub
Private Sub iOleControlVB_OnMnemonic(bHandled As Boolean, ByVal iShift As evbComCtlKeyboardState, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    Dim liKey        As Long
    Dim liAccel      As Long
    
    liKey = wParam And &HFFFF&
    If mhWnd Then
        Dim liIndex      As Long
        For liIndex = ZeroL To SendMessage(mhWnd, TCM_GETITEMCOUNT, ZeroL, ZeroL) - OneL
            liAccel = AccelChar(pTab_Text(liIndex))
            If liAccel Then
                If GetVirtKey(liAccel) = liKey Then
                    If liIndex <> SendMessage(mhWnd, TCM_GETCURSEL, ZeroL, ZeroL) Then SetSelectedTab liIndex + OneL
                    Exit For
                End If
            End If
        Next
    End If
    bHandled = True
End Sub

Private Sub iOleInPlaceActiveObjectVB_TranslateAccelerator(bHandled As Boolean, lReturn As Long, ByVal iShift As evbComCtlKeyboardState, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Forward arrow and other keys to the tabstrip.
    '---------------------------------------------------------------------------------------
    If uMsg = WM_KEYDOWN Then
        Select Case wParam And &HFFFF&
        Case vbKeyPageUp To vbKeyDown
            If mhWnd Then
                SendMessage mhWnd, uMsg, wParam, lParam
                bHandled = True
            End If
        End Select
    End If
End Sub

Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    Select Case uMsg
    Case WM_SETFOCUS
        vbComCtlTlb.SetFocus mhWnd
    Case WM_KILLFOCUS
        DeActivateIPAO Me
    End Select
End Sub

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Respond to notifications from the tabstrip and handle focus.
    '---------------------------------------------------------------------------------------
    Dim loItem       As cTab
    Dim bCancel      As OLE_CANCELBOOL
    
    Select Case uMsg
    Case WM_NOTIFY
        bHandled = True
        If mhWnd = MemOffset32(lParam, NMHDR_hwndFrom) Then
            Select Case MemOffset32(lParam, NMHDR_code)
            Case TCN_SELCHANGING
                RaiseEvent BeforeClick(bCancel)
                lReturn = -bCancel
                
            Case TCN_SELCHANGE
                Set loItem = pItem(SendMessage(mhWnd, TCM_GETCURSEL, ZeroL, ZeroL))
                'debug.assert Not loItem Is Nothing
                If Not loItem Is Nothing Then
                    RaiseEvent Click(loItem)
                End If
                
            End Select
        
        End If
   
    Case WM_SETFOCUS
        ActivateIPAO Me

    Case WM_MOUSEACTIVATE
        If GetFocus() <> mhWnd Then
            vbComCtlTlb.SetFocus UserControl.hWnd
            lReturn = MA_NOACTIVATE
            bHandled = True
        End If

    Case WM_MOUSEWHEEL
        miWheelDelta = miWheelDelta - hiword(wParam)
        If Abs(miWheelDelta) >= 120& Then
            If CBool(miWheelDelta And &H80000000) Then
                If SendMessage(mhWnd, TCM_GETCURSEL, ZeroL, ZeroL) > ZeroL _
                    Then SetSelectedTab SendMessage(mhWnd, TCM_GETCURSEL, ZeroL, ZeroL)
                Else
                    If SendMessage(mhWnd, TCM_GETCURSEL, ZeroL, ZeroL) < (SendMessage(mhWnd, TCM_GETITEMCOUNT, ZeroL, ZeroL) - OneL) _
                        Then SetSelectedTab SendMessage(mhWnd, TCM_GETCURSEL, ZeroL, ZeroL) + TwoL
                    End If
                    miWheelDelta = ZeroL
                End If
                lReturn = ZeroL
                bHandled = True
            Case WM_PARENTNOTIFY
                Select Case wParam And &HFFFF&
                Case WM_RBUTTONDOWN:    PostMessage hWnd, UM_MouseDown, wParam, lParam
                End Select
                bHandled = True
            Case UM_MouseDown
                Select Case wParam And &HFFFF&
                Case WM_RBUTTONDOWN:    pRightClick loword(lParam), hiword(lParam)
                End Select
                bHandled = True
            End Select

End Sub

Private Sub moImageListEvent_Changed()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Update the imagelist.
    '---------------------------------------------------------------------------------------
    Set ImageList = moImageListEvent
End Sub

Private Sub moMnemonics_Changed()
    OnControlInfoChanged Me
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Update the font if appropriate.
    '---------------------------------------------------------------------------------------
    If StrComp("Font", PropertyName) = ZeroL Then moFont.OnAmbientFontChanged Ambient.Font
End Sub

Private Sub UserControl_Initialize()
    ForceWindowToShowAllUIStates hWnd
    LoadShellMod
    InitCC ICC_TAB_CLASSES
    Set moFontPage = New pcSupportFontPropPage
    Set moMnemonics = New pcMnemonics
End Sub

Private Sub UserControl_InitProperties()
    Set moFont = Font_CreateDefault(Ambient.Font)
    mbHotTrack = DEF_HotTrack
    mbButtons = DEF_Buttons
    mbMultiLine = DEF_MultiLine
    mbRightJustify = DEF_RightJustify
    mbFlatSeparators = DEF_FlatSeparators
    mbFlatButtons = DEF_FlatButtons
    mbThemeable = DEF_Themeable
    pCreate
End Sub

Private Sub UserControl_Resize()
    If mhWnd Then
        MoveWindow mhWnd, ZeroL, ZeroL, ScaleX(Width, vbTwips, vbPixels), ScaleY(Height, vbTwips, vbPixels), OneL
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set moFont = Font_Read(PropBag, PROP_Font, Ambient.Font)
    mbHotTrack = PropBag.ReadProperty(PROP_HotTrack, DEF_HotTrack)
    mbButtons = PropBag.ReadProperty(PROP_Buttons, DEF_Buttons)
    mbMultiLine = PropBag.ReadProperty(PROP_MultiLine, DEF_MultiLine)
    mbRightJustify = PropBag.ReadProperty(PROP_RightJustify, DEF_RightJustify)
    mbFlatSeparators = PropBag.ReadProperty(PROP_FlatSeparators, DEF_FlatSeparators)
    mbFlatButtons = PropBag.ReadProperty(PROP_FlatButtons, DEF_FlatButtons)
    mbThemeable = PropBag.ReadProperty(PROP_Themeable, DEF_Themeable)
    pCreate
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Font_Write moFont, PropBag, PROP_Font
    PropBag.WriteProperty PROP_HotTrack, mbHotTrack, DEF_HotTrack
    PropBag.WriteProperty PROP_Buttons, mbButtons, DEF_Buttons
    PropBag.WriteProperty PROP_MultiLine, mbMultiLine, DEF_MultiLine
    PropBag.WriteProperty PROP_RightJustify, mbRightJustify, DEF_RightJustify
    PropBag.WriteProperty PROP_FlatSeparators, mbFlatSeparators, DEF_FlatSeparators
    PropBag.WriteProperty PROP_FlatButtons, mbFlatButtons, DEF_FlatButtons
    PropBag.WriteProperty PROP_Themeable, mbThemeable, DEF_Themeable
End Sub

Private Sub UserControl_Terminate()
    pDestroy
    If mhFont Then moFont.ReleaseHandle mhFont
    ReleaseShellMod
    Set moFontPage = Nothing
End Sub

Friend Property Get fSupportFontPropPage() As pcSupportFontPropPage
    Set fSupportFontPropPage = moFontPage
End Property

Private Sub moFontPage_AddFonts(ByVal o As ppFont)
    o.ShowProps PROP_Font
End Sub

Private Sub moFontPage_GetAmbientFont(o As stdole.StdFont)
    Set o = Ambient.Font
End Sub

Private Sub pRightClick(ByVal X As Long, ByVal Y As Long)
    Dim tHitTest      As TCHITTESTINFO
    tHitTest.pt.X = X
    tHitTest.pt.Y = Y
    If mhWnd Then RaiseEvent RightClick(pItem(SendMessage(mhWnd, TCM_HITTEST, ZeroL, VarPtr(tHitTest))))
End Sub



Private Sub moFont_Changed()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Update the font used by the control.
    '---------------------------------------------------------------------------------------
    moFont.OnAmbientFontChanged Ambient.Font
    pSetFont
    pPropChanged PROP_Font
End Sub

Private Sub pSetFont()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Update the font handle used by the statusbar.
    '---------------------------------------------------------------------------------------
    On Error GoTo handler
    Dim hFont      As Long
    hFont = moFont.GetHandle
    SendMessage mhWnd, WM_SETFONT, hFont, OneL
    If mhFont Then moFont.ReleaseHandle mhFont
    mhFont = hFont
    On Error GoTo 0
    Exit Sub
handler:
    Resume Next
End Sub


Private Sub pPropChanged(ByRef s As String)
    If Ambient.UserMode = False Then PropertyChanged s
End Sub

Private Sub pCreate()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Create the tabstrip and install the subclasses.
    '---------------------------------------------------------------------------------------
    pDestroy
    
    Dim lsAnsi      As String
    lsAnsi = StrConv(WC_TABCONTROL & vbNullChar, vbFromUnicode)
    mhWnd = CreateWindowEx(ZeroL, StrPtr(lsAnsi), ZeroL, pStyle(), ZeroL, ZeroL, ScaleX(Width, vbTwips, vbPixels), ScaleY(Height, vbTwips, vbPixels), UserControl.hWnd, ZeroL, App.hInstance, ByVal ZeroL)
    
    If mhWnd Then
        If Not mbMultiLine Then EnableWindowTheme FindWindowExW(mhWnd, ZeroL, "msctls_updown32", vbNullString), mbThemeable
        
        SendMessage mhWnd, TCM_SETITEMEXTRA, 8&, ZeroL
        
        pSetFont
        If Not moImageList Is Nothing Then SendMessage mhWnd, TCM_SETIMAGELIST, ZeroL, moImageList.hIml
        SendMessage mhWnd, TCM_SETEXTENDEDSTYLE, TCS_EX_FLATSEPARATORS, -mbFlatSeparators
        UserControl_Resize
        
        If Ambient.UserMode Then
            
            VTableSubclass_OleControl_Install Me
            VTableSubclass_IPAO_Install Me
            
            Subclass_Install Me, UserControl.hWnd, Array(WM_NOTIFY, WM_MOUSEWHEEL, WM_PARENTNOTIFY, UM_MouseDown), WM_SETFOCUS
            Subclass_Install Me, mhWnd, Array(WM_SETFOCUS, WM_MOUSEACTIVATE), WM_KILLFOCUS
            
        Else
            fTabs_Add "Sample Tab", NegOneL, vbNullString, pGetMissing()
            
        End If
        EnableWindowTheme mhWnd, mbThemeable
    End If

End Sub

Private Function pGetMissing(Optional ByVal v As Variant)
    pGetMissing = v
End Function

Private Sub pDestroy()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Destroy the tabstrip and subclasses.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        fTabs_Clear 'release memory allocated for storing key strings
        VTableSubclass_OleControl_Remove
        VTableSubclass_IPAO_Remove
        Subclass_Remove Me, UserControl.hWnd
        Subclass_Remove Me, mhWnd
        DestroyWindow mhWnd
        mhWnd = ZeroL
    End If
End Sub

Private Function pItem(ByVal iIndex As Long) As cTab
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return a cTab object representing the given index.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        If iIndex > NegOneL And iIndex < SendMessage(mhWnd, TCM_GETITEMCOUNT, ZeroL, ZeroL) Then
            Set pItem = New cTab
            pItem.fInit Me, pTab_Info(iIndex, TCIF_PARAM), iIndex
        End If
    End If
End Function

Private Function pStyle() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the window style for the tabstrip based on our properties.
    '---------------------------------------------------------------------------------------
    If (mbHotTrack) Then pStyle = pStyle Or TCS_HOTTRACK
    If (mbButtons) Then pStyle = pStyle Or TCS_BUTTONS
    If (mbFlatButtons) Then pStyle = pStyle Or TCS_FLATBUTTONS
    If (mbMultiLine) Then pStyle = pStyle Or TCS_MULTILINE
    If (mbRightJustify) Then pStyle = pStyle Or TCS_RIGHTJUSTIFY
    pStyle = pStyle Or WS_VISIBLE Or WS_CHILD Or WS_CLIPSIBLINGS
End Function



Friend Sub fTabs_Enum_NextItem(tEnum As tEnum, vNextItem As Variant, bNoMoreItems As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the next cTab object in an enumeration.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        If tEnum.iControl <> miTabControl Then gErr vbccCollectionChangedDuringEnum, cTabs
        tEnum.iIndex = tEnum.iIndex + OneL
        
        bNoMoreItems = (tEnum.iIndex < ZeroL) Or (tEnum.iIndex >= SendMessage(mhWnd, TCM_GETITEMCOUNT, ZeroL, ZeroL))
        
        If Not bNoMoreItems Then
            Dim loTab      As cTab
            Set loTab = New cTab
            
            loTab.fInit Me, pTab_Info(tEnum.iIndex, TCIF_PARAM), tEnum.iIndex
            Set vNextItem = loTab
        End If
    Else
        bNoMoreItems = True
        
    End If
End Sub

Friend Sub fTabs_Enum_Skip(tEnum As tEnum, ByVal iSkipCount As Long, bSkippedAll As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Skip a number of tabs in an enumeration.
    '---------------------------------------------------------------------------------------
    If tEnum.iControl <> miTabControl Then gErr vbccCollectionChangedDuringEnum, cTabs
    tEnum.iIndex = tEnum.iIndex + iSkipCount
    bSkippedAll = (tEnum.iIndex <= SendMessage(mhWnd, TCM_GETITEMCOUNT, ZeroL, ZeroL))
End Sub

Friend Property Get fTabs_Enum_Control() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return an identity value for the current tabs collection.
    '---------------------------------------------------------------------------------------
    fTabs_Enum_Control = miTabControl
End Property


Friend Function fTabs_Add(ByRef sText As String, ByVal iIconIndex As Long, ByRef sKey As String, ByRef vTabInsertBefore As Variant) As cTab
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Add a tab to the collection.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim ls            As String
        Dim lpKey         As Long
        Dim liInsert      As Long
    
        If IsMissing(vTabInsertBefore) Then
            liInsert = SendMessage(mhWnd, TCM_GETITEMCOUNT, ZeroL, ZeroL)
        Else
            liInsert = pTabs_GetIndex(vTabInsertBefore)
        End If
    
        ls = StrConv(sText & vbNullChar, vbFromUnicode)
        MidB$(msTextBuffer, OneL, LenB(ls)) = ls
        ls = vbNullString
        'just in case sText is longer than the text buffer, ensure its still null-terminated
        MidB$(msTextBuffer, LenB(msTextBuffer), OneL) = vbNullChar
    
        If LenB(sKey) Then
            ls = StrConv(sKey & vbNullString, vbFromUnicode)
            If pTabs_FindString(StrPtr(ls)) > NegOneL Then gErr vbccKeyAlreadyExists, cTabs
            lpKey = MemAllocFromString(StrPtr(ls), LenB(ls))
        End If
    
        With mtItem
            .mask = TCIF_TEXT Or TCIF_PARAM
            If iIconIndex > NegOneL Then .mask = .mask Or TCIF_IMAGE
            .pszText = StrPtr(msTextBuffer)
            .cchTextMax = LenB(msTextBuffer)
            .iImage = iIconIndex
            .lpKey = lpKey
            .lParam = NextItemId()
        
            If SendMessage(mhWnd, TCM_INSERTITEMA, liInsert, VarPtr(mtItem)) = liInsert Then
                Set fTabs_Add = New cTab
                fTabs_Add.fInit Me, .lParam, liInsert
                Incr miTabControl
                moMnemonics.Add sText
                'If sendmessage(mhwnd,TCM_GETITEMCOUNT, ZeroL, ZeroL) = OneL Then SetSelectedTab OneL
            Else
                MemFree lpKey
                'debug.assert False
            End If
        End With
    End If
    
End Function

Friend Sub fTabs_Remove(ByRef vTab As Variant)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Remove a tab from the collection.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        mtItem.mask = TCIF_PARAM
        pTabs_Delete pTabs_GetIndex(vTab)
    End If
End Sub

Friend Property Get fTabs_Exists(ByRef vTab As Variant) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return a value indicating whether a tab exists in the collection.
    '---------------------------------------------------------------------------------------
    On Error GoTo handler
    pTabs_GetIndex vTab
    fTabs_Exists = True
handler:
    On Error GoTo 0
End Property

Friend Property Get fTabs_Item(ByRef vTab As Variant) As cTab
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the specified cTab object.
    '---------------------------------------------------------------------------------------
    Dim liIndex      As Long
    liIndex = pTabs_GetIndex(vTab)
    Set fTabs_Item = New cTab
    fTabs_Item.fInit Me, pTab_Info(liIndex, TCIF_PARAM), liIndex
End Property

Friend Sub fTabs_Clear()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Remove all tabs.
    '---------------------------------------------------------------------------------------
    Dim liIndex      As Long
    
    If mhWnd Then
        For liIndex = SendMessage(mhWnd, TCM_GETITEMCOUNT, ZeroL, ZeroL) - OneL To ZeroL Step NegOneL
            pTabs_Delete liIndex
        Next
    End If
End Sub

Friend Property Get fTabs_Count() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the number of tabs in the collection.
    '---------------------------------------------------------------------------------------
    If mhWnd Then fTabs_Count = SendMessage(mhWnd, TCM_GETITEMCOUNT, ZeroL, ZeroL)
End Property

Private Function pTabs_GetIndex(ByRef vItem As Variant) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return a tab index given its key, object or index.
    '---------------------------------------------------------------------------------------
    On Error GoTo handler
    If VarType(vItem) = vbObject Then
        Dim loTab      As cTab
        Set loTab = vItem
        If loTab.fIsOwner(Me) Then
            pTabs_GetIndex = loTab.Index - OneL
        End If
    ElseIf VarType(vItem) = vbString Then
        Dim ls      As String
        ls = StrConv(vItem & vbNullChar, vbFromUnicode)
        If LenB(ls) = OneL Then GoTo handler
        pTabs_GetIndex = pTabs_FindString(StrPtr(ls))
        If pTabs_GetIndex = NegOneL Then GoTo handler
        
    Else
        pTabs_GetIndex = (CLng(vItem) - OneL)
        If pTabs_GetIndex < ZeroL Or pTabs_GetIndex >= SendMessage(mhWnd, TCM_GETITEMCOUNT, ZeroL, ZeroL) Then GoTo handler
        
    End If
    
    Exit Function
handler:
    On Error GoTo 0
    gErr vbccKeyOrIndexNotFound, cTabs
End Function

Private Function pTabs_FindString(ByVal lpsz As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Find a tab given its key.
    '---------------------------------------------------------------------------------------
    'debug.assert lpsz
    
    If lpsz Then
    
        Dim lpStrTab      As Long
        
        If mhWnd Then
        
            For pTabs_FindString = ZeroL To SendMessage(mhWnd, TCM_GETITEMCOUNT, ZeroL, ZeroL) - OneL
                lpStrTab = pTab_Info(pTabs_FindString, TCIF_lpKey)
                If lpStrTab Then
                    If lstrcmp(lpStrTab, lpsz) = ZeroL Then Exit Function
                End If
            Next
        End If
    End If
    
    pTabs_FindString = NegOneL
    
End Function


Friend Property Get fTab_Text(ByVal hTab As Long, ByRef iIndex As Long) As String
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the text of a tab.
    '---------------------------------------------------------------------------------------
    If pTab_Verify(hTab, iIndex) Then
        fTab_Text = pTab_Text(iIndex)
    End If
End Property
Friend Property Let fTab_Text(ByVal hTab As Long, ByRef iIndex As Long, ByRef sNew As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the text of a tab.
    '---------------------------------------------------------------------------------------
    If pTab_Verify(hTab, iIndex) Then
        pTab_Text(iIndex) = sNew
    End If
End Property

Friend Property Get fTab_IconIndex(ByVal hTab As Long, ByRef iIndex As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the iconindex of a tab.
    '---------------------------------------------------------------------------------------
    If pTab_Verify(hTab, iIndex) Then
        fTab_IconIndex = pTab_Info(iIndex, TCIF_IMAGE)
    End If
End Property
Friend Property Let fTab_IconIndex(ByVal hTab As Long, ByRef iIndex As Long, ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the iconindex of a tab.
    '---------------------------------------------------------------------------------------
    If pTab_Verify(hTab, iIndex) Then
        pTab_Info(iIndex, TCIF_IMAGE) = iNew
    End If
End Property

Friend Property Get fTab_Key(ByVal hTab As Long, ByRef iIndex As Long) As String
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the key of a tab.
    '---------------------------------------------------------------------------------------
    If pTab_Verify(hTab, iIndex) Then
        Dim lpKey      As Long
        lpKey = pTab_Info(iIndex, TCIF_lpKey)
        lstrToStringA lpKey, fTab_Key
    End If
End Property
Friend Property Let fTab_Key(ByVal hTab As Long, ByRef iIndex As Long, ByRef sNew As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the key of an item.
    '---------------------------------------------------------------------------------------
    If pTab_Verify(hTab, iIndex) Then
        Dim ls      As String
        
        If LenB(sNew) Then
            ls = StrConv(sNew & vbNullChar, vbFromUnicode)
            If pTabs_FindString(StrPtr(ls)) <> NegOneL Then gErr vbccKeyAlreadyExists, cTab
        End If
        
        mtItem.lpKey = pTab_Info(iIndex, TCIF_lpKey)
        If mtItem.lpKey Then MemFree mtItem.lpKey
        
        With mtItem
            'lParam is already set because of the "pTab_Info(iIndex, TCIF_lpKey)" call above
            '.lParam = pTab_Info(iIndex, TCIF_PARAM)
            '.mask = TCIF_PARAM
            
            If mhWnd Then
                .lpKey = MemAllocFromString(StrPtr(ls), LenB(ls))
                SendMessage mhWnd, TCM_SETITEMA, iIndex, VarPtr(mtItem)
            End If

        End With
    End If
End Property

Friend Property Get fTab_Index(ByVal hTab As Long, ByRef iIndex As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the one-based index in the collection.
    '---------------------------------------------------------------------------------------
    If pTab_Verify(hTab, iIndex) Then
        fTab_Index = iIndex + OneL
    End If
End Property

Private Function pTab_Verify(ByVal hTab As Long, ByRef iIndex As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Verify whether a tab still exists in the collection.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        pTab_Verify = CBool(hTab = pTab_Info(iIndex, TCIF_PARAM))
        If Not pTab_Verify Then
            pTab_Verify = True
            For iIndex = ZeroL To SendMessage(mhWnd, TCM_GETITEMCOUNT, ZeroL, ZeroL)
                If pTab_Info(iIndex, TCIF_PARAM) = hTab Then Exit Function
            Next
            iIndex = NegOneL
            pTab_Verify = False
        End If
    End If
    If Not pTab_Verify Then gErr vbccItemDetached, cTab
End Function

Private Sub pTabs_Delete(ByVal iIndex As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Tell the tab control to delete a tab.
    '---------------------------------------------------------------------------------------
    Dim lpKey      As Long
    lpKey = pTab_Info(iIndex, TCIF_lpKey)
    moMnemonics.Remove pTab_Text(iIndex)
    
    If SendMessage(mhWnd, TCM_DELETEITEM, iIndex, ZeroL) Then
        If lpKey Then MemFree lpKey
        Incr miTabControl
    End If
End Sub

Private Property Get pTab_Info(ByVal iIndex As Long, ByVal iMask As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return a value from the TCITEM structure.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        With mtItem
            .mask = iMask And TCIF_ComCtlMask
            If SendMessage(mhWnd, TCM_GETITEMA, iIndex, VarPtr(mtItem)) Then
                If iMask = TCIF_lpKey Then
                    pTab_Info = .lpKey
                ElseIf iMask = TCIF_PARAM Then
                    pTab_Info = .lParam
                ElseIf iMask = TCIF_IMAGE Then
                    pTab_Info = .iImage
                Else
                    'debug.assert False
                    
                End If
                
            Else
                'debug.assert False
                
            End If
        End With
    End If
End Property

Private Property Let pTab_Info(ByVal iIndex As Long, ByVal iMask As Long, ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set a value in the TCITEM structure.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        With mtItem
            .mask = iMask And TCIF_ComCtlMask
            
            If iMask = TCIF_IMAGE Then
                .iImage = iNew
            Else
                'debug.assert False
            End If
            
            If SendMessage(mhWnd, TCM_SETITEMA, iIndex, VarPtr(mtItem)) = ZeroL Then
                'debug.assert False
                
            End If
        End With
    End If
End Property

Private Property Get pTab_Text(ByVal iIndex As Long) As String
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return an item's text.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        With mtItem
            .mask = TCIF_TEXT
            .pszText = StrPtr(msTextBuffer)
            .cchTextMax = LenB(msTextBuffer)
            
            SendMessage mhWnd, TCM_GETITEMA, iIndex, VarPtr(mtItem)
        
            lstrToStringA .pszText, pTab_Text
        End With
    End If
End Property

Private Property Let pTab_Text(ByVal iIndex As Long, ByVal sNew As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set an item's text.
    '---------------------------------------------------------------------------------------
    With mtItem
        .mask = TCIF_TEXT
        
        Dim ls      As String
        ls = StrConv(sNew & vbNullChar, vbFromUnicode)
        .cchTextMax = LenB(ls)
        
        MidB$(msTextBuffer, OneL, .cchTextMax) = ls
        MidB$(msTextBuffer, LenB(msTextBuffer), OneL) = vbNullChar
        .pszText = StrPtr(msTextBuffer)
        
        If mhWnd Then SendMessage mhWnd, TCM_SETITEMA, iIndex, VarPtr(mtItem)
        
    End With
End Property



Public Property Get Tabs() As cTabs
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return a collection of tabs
    '---------------------------------------------------------------------------------------
    Set Tabs = New cTabs
    Tabs.fInit Me
End Property


Public Property Get FlatSeparators() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return a value indicating whether separators are shown when flat buttons are
    '             enabled.
    '---------------------------------------------------------------------------------------
    FlatSeparators = mbFlatSeparators
End Property
Public Property Let FlatSeparators(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set whether separators are shown when flat buttons are enabled.
    '---------------------------------------------------------------------------------------
    mbFlatSeparators = bNew
    If mhWnd Then SendMessage mhWnd, TCM_SETEXTENDEDSTYLE, TCS_EX_FLATSEPARATORS, -bNew
    pPropChanged PROP_FlatSeparators
End Property
Public Property Get HotTrack() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return a value indicating whether hot tracking is enabled.
    '---------------------------------------------------------------------------------------
    HotTrack = mbHotTrack
End Property
Public Property Let HotTrack(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set whether hot tracking is enabled.
    '---------------------------------------------------------------------------------------
    mbHotTrack = bNew
    If mhWnd Then SetWindowLong mhWnd, GWL_STYLE, pStyle()
    pPropChanged PROP_HotTrack
End Property
Public Property Get Buttons() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return a value indicating if buttons are used rather than tabs.
    '---------------------------------------------------------------------------------------
    Buttons = mbButtons
End Property
Public Property Let Buttons(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set a value indicating if buttons are used rather than tabs.
    '---------------------------------------------------------------------------------------
    mbButtons = bNew
    If mhWnd Then SetWindowLong mhWnd, GWL_STYLE, pStyle()
    pPropChanged PROP_Buttons
End Property
Public Property Get FlatButtons() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return a value indicating whether to paint 3D borders around the buttons.
    '---------------------------------------------------------------------------------------
    FlatButtons = mbFlatButtons
End Property
Public Property Let FlatButtons(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set whether to paint 3D borders around the buttons.
    '---------------------------------------------------------------------------------------
    mbFlatButtons = bNew
    If mhWnd Then SetWindowLong mhWnd, GWL_STYLE, pStyle()
    pPropChanged PROP_FlatButtons
End Property
Public Property Get MultiLine() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return a value indicating whether tabs are allowed to wrap to multiple lines.
    '---------------------------------------------------------------------------------------
    MultiLine = mbMultiLine
End Property
Public Property Let MultiLine(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set whether tabs are allowed to wrap to multiple lines.
    '---------------------------------------------------------------------------------------
    mbMultiLine = bNew
    If mhWnd Then SetWindowLong mhWnd, GWL_STYLE, pStyle()
    pPropChanged PROP_MultiLine
    If bNew = False Then EnableWindowTheme FindWindowExW(mhWnd, ZeroL, "msctls_updown32", vbNullString), mbThemeable
End Property
Public Property Get RightJustify() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return a value indicating whether tabs are aligned to the right of the control.
    '---------------------------------------------------------------------------------------
    RightJustify = mbRightJustify
End Property
Public Property Let RightJustify(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set whether tabs are aligned to the right of the control.
    '---------------------------------------------------------------------------------------
    mbRightJustify = bNew
    If mhWnd Then SetWindowLong mhWnd, GWL_STYLE, pStyle()
    pPropChanged PROP_RightJustify
End Property
Public Sub SetPadding(ByVal fx As Single, ByVal fY As Single)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the padding of the tabs.
    '---------------------------------------------------------------------------------------
    If mhWnd _
        Then SendMessage mhWnd, TCM_SETPADDING, ZeroL, _
        (ScaleX(fx, vbContainerSize, vbPixels) And &H7FFF&) _
        Or _
        ((ScaleY(fY, vbContainerSize, vbPixels) And &H7FFF&) * &H10000)
End Sub
Public Property Get Font() As cFont
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the font used by the tab control.
    '---------------------------------------------------------------------------------------
    Set Font = moFont
End Property
Public Property Set Font(ByVal oNew As cFont)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the font used by the tab control.
    '---------------------------------------------------------------------------------------
    If oNew Is Nothing _
        Then Set moFont = Font_CreateDefault(Ambient.Font) _
    Else Set moFont = oNew
        pSetFont
        pPropChanged PROP_Font
End Property


Public Property Get ImageList() As cImageList
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the imagelist used by the tab control.
    '---------------------------------------------------------------------------------------
    Set ImageList = moImageList
End Property
Public Property Set ImageList(ByVal oNew As cImageList)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the imagelist used by the tab control.
    '---------------------------------------------------------------------------------------
    On Error Resume Next
    Set moImageList = oNew
    Set moImageListEvent = oNew
    On Error GoTo 0
    If mhWnd Then
        If Not moImageList Is Nothing Then _
        SendMessage mhWnd, TCM_SETIMAGELIST, ZeroL, moImageList.hIml Else _
        SendMessage mhWnd, TCM_SETIMAGELIST, ZeroL, ZeroL
    End If
End Property


Public Property Get SelectedTab() As cTab
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the selected tab.
    '---------------------------------------------------------------------------------------
    Dim iIndex      As Long
    If mhWnd Then
        iIndex = SendMessage(mhWnd, TCM_GETCURSEL, ZeroL, ZeroL)
        If iIndex > NegOneL Then
            Set SelectedTab = New cTab
            SelectedTab.fInit Me, pTab_Info(iIndex, TCIF_PARAM), iIndex
        End If
    End If
End Property

Public Property Set SelectedTab(ByVal oNew As cTab)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the selected tab.
    '---------------------------------------------------------------------------------------
    SetSelectedTab oNew
End Property

Public Sub SetSelectedTab(ByVal vTab As Variant)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the selected tab.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim liIndex      As Long
        liIndex = pTabs_GetIndex(vTab)
        
        Dim loTab      As cTab
        Set loTab = New cTab
        Dim bCancel      As OLE_CANCELBOOL
        
        loTab.fInit Me, pTab_Info(liIndex, TCIF_PARAM), liIndex
        
        RaiseEvent BeforeClick(bCancel)
        If Not bCancel Then
            If SendMessage(mhWnd, TCM_SETCURSEL, liIndex, ZeroL) > NegOneL Then RaiseEvent Click(loTab)
        End If
    End If
End Sub

Public Property Get ClientLeft() As Single
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the left of the client area.
    '---------------------------------------------------------------------------------------
    Dim rc      As RECT
    pGetClientRect rc
    ClientLeft = ScaleX(rc.Left, vbPixels, vbContainerPosition)
End Property
Public Property Get ClientTop() As Single
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the top of the client area.
    '---------------------------------------------------------------------------------------
    Dim rc      As RECT
    pGetClientRect rc
    ClientTop = ScaleY(rc.Top, vbPixels, vbContainerPosition)
End Property
Public Property Get ClientWidth() As Single
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the width of the client area.
    '---------------------------------------------------------------------------------------
    Dim rc      As RECT
    pGetClientRect rc
    ClientWidth = ScaleX((rc.Right - rc.Left), vbPixels, vbContainerPosition)
End Property
Public Property Get ClientHeight() As Single
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the height of the client area.
    '---------------------------------------------------------------------------------------
    Dim rc      As RECT
    pGetClientRect rc
    ClientHeight = ScaleY((rc.bottom - rc.Top), vbPixels, vbContainerPosition)
End Property

Public Sub MoveToClient(ByVal oControl As Object)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Move a control to the client area.
    '---------------------------------------------------------------------------------------
    Dim o2      As VBControlExtender
    Dim rc      As RECT
    
    pGetClientRect rc
    On Error GoTo handler
    
    Set o2 = oControl
    o2.Move ScaleX(rc.Left, vbPixels, vbContainerPosition), _
    ScaleY(rc.Top, vbPixels, vbContainerPosition), _
    ScaleX((rc.Right - rc.Left), vbPixels, vbContainerSize), _
    ScaleY((rc.bottom - rc.Top), vbPixels, vbContainerSize)
    o2.Visible = CBool(o2.Width > ScaleX(2, vbPixels, vbTwips)) And CBool(o2.Height > ScaleY(2, vbPixels, vbTwips))
    If False Then
handler:
        On Error Resume Next
        oControl.Move ScaleX(rc.Left, vbPixels, vbContainerPosition), _
        ScaleY(rc.Top, vbPixels, vbContainerPosition), _
        ScaleX((rc.Right - rc.Left), vbPixels, vbContainerSize), _
        ScaleY((rc.bottom - rc.Top), vbPixels, vbContainerSize)
        oControl.Visible = CBool(oControl.Width > ScaleX(2, vbPixels, vbTwips)) And CBool(oControl.Height > ScaleY(2, vbPixels, vbTwips))
    End If
    On Error GoTo 0
End Sub

Private Sub pGetClientRect(rc As RECT)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Get the dimensions of the client area.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
    
        GetWindowRect mhWnd, rc
        ScreenToClient UserControl.ContainerHwnd, rc.Left
        ScreenToClient UserControl.ContainerHwnd, rc.Right
        SendMessage mhWnd, TCM_ADJUSTRECT, ZeroL, VarPtr(rc)
        If rc.Top > rc.bottom Then rc.bottom = rc.Top
        If rc.Left > rc.Right Then rc.Right = rc.Left
    End If
End Sub

Public Property Get hWnd() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the hwnd of the usercontrol.
    '---------------------------------------------------------------------------------------
    hWnd = UserControl.hWnd
End Property
Public Property Get hWndTabControl() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the hwnd of the tab control.
    '---------------------------------------------------------------------------------------
    If mhWnd Then hWndTabControl = mhWnd
End Property

Public Property Get Themeable() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return a value indicating whether the default theme is to be used if available.
    '---------------------------------------------------------------------------------------
    Themeable = mbThemeable
End Property
Public Property Let Themeable(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set a value indicating whether the default theme is to be used if available.
    '---------------------------------------------------------------------------------------
    If bNew Xor mbThemeable Then
        mbThemeable = bNew
        pPropChanged PROP_Themeable
        If mhWnd Then
            If mbButtons And bNew Then
                
                SendMessage mhWnd, WM_SETREDRAW, ZeroL, ZeroL

                mbButtons = False
                SetWindowLong mhWnd, GWL_STYLE, pStyle()
                
                EnableWindowTheme mhWnd, mbThemeable
                EnableWindowTheme FindWindowExW(mhWnd, ZeroL, "msctls_updown32", vbNullString), mbThemeable
                
                mbButtons = True
                SetWindowLong mhWnd, GWL_STYLE, pStyle()
                
                SendMessage mhWnd, WM_SETREDRAW, OneL, ZeroL
                InvalidateRect mhWnd, ByVal ZeroL, OneL
            
            Else
                EnableWindowTheme mhWnd, mbThemeable
                EnableWindowTheme FindWindowExW(mhWnd, ZeroL, "msctls_updown32", vbNullString), mbThemeable
            End If
        End If
    End If
End Property
