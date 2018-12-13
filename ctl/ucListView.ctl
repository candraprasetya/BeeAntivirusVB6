VERSION 5.00
Begin VB.UserControl ucListView 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   HasDC           =   0   'False
   PropertyPages   =   "ucListView.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ucListView.ctx":000D
End
Attribute VB_Name = "ucListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'ucListView.ctl              12/15/04
'
'           PURPOSE:
'               Implement the SysListView control.
'
'           LINEAGE:
'               http://www.vbaccelerator.com/home/VB/Code/Controls/ListView/VB6_ListView_Full_Source.asp
'               vbalListView.ctl
'
'               Image drag created from LVDrag.vbp by Brad Martinez http://www.mvps.org
'
'==================================================================================================

Option Explicit

Public Enum eListViewColumnAutoSize
    lvwColumnSizeToItemText = LVSCW_AUTOSIZE
    lvwColumnSizeToColumnText = LVSCW_AUTOSIZE_USEHEADER
End Enum

Public Enum eListViewAutoArrange
    lvwArrangeNone
    lvwArrangeLeft
    lvwArrangeTop
End Enum

Public Enum eListViewArrange
    lvwDefault = LVA_DEFAULT
    lvwLeft = LVA_ALIGNLEFT
    lvwTop = LVA_ALIGNTOP
    lvwSnapToGrid = LVA_SNAPTOGRID
End Enum

Public Enum eListViewImageType
    lvwImageLargeIcon = LVSIL_NORMAL
    lvwImageSmallIcon = LVSIL_SMALL
    'lvwImageStateImages = LVSIL_STATE
    lvwImageHeaderImages = &H8000&  ' Not part of ComCtl32.DLL
End Enum

Public Enum eListViewStyle
    lvwIcon = LVS_ICON
    lvwDetails = LVS_REPORT
    lvwSmallIcon = LVS_SMALLICON
    lvwList = LVS_LIST
    lvwTile = LV_VIEW_TILE
End Enum

Public Enum eListViewColumnAlign
    lvwAlignLeft
    lvwAlignRight
    lvwAlignCenter
End Enum

Public Enum eListViewSortOrder
    lvwSortNone
    lvwSortAscending
    lvwSortDescending
End Enum

Public Enum eListViewSortType
    lvwSortString
    lvwSortStringNoCase
    lvwSortNumeric
    lvwSortCurrency
    lvwSortDate
    lvwSortIndent
    lvwSortSelected
End Enum

Public Enum eListViewGetNextItem
    lvwFindAll = LVNI_ALL
    lvwFindSelected = LVNI_SELECTED
    lvwFindCut = LVNI_CUT
    
    lvwFindDirAbove = LVNI_ABOVE
    lvwFindDirBelow = LVNI_BELOW
    lvwFindDirLeft = LVNI_TOLEFT
    lvwFindDirRight = LVNI_TORIGHT
End Enum

Public Enum eListViewOleImageDrag
    lvwOleImageDragNone
    lvwOleImageDragFocused
    lvwOleImageDragSelected
End Enum

Implements iSubclass
Implements iOleInPlaceActiveObjectVB
Implements iOleControlVB
Implements iLVCompare

Public Event ColumnClick(ByVal oColumn As cColumn)
Public Event ColumnAfterDrag(ByVal oColumn As cColumn, ByRef iNewPosition As Long, ByRef bCancel As OLE_CANCELBOOL)
Public Event ColumnAfterSize(ByVal oColumn As cColumn)
Public Event ColumnBeforeDrag(ByVal oColumn As cColumn, ByRef bCancel As OLE_CANCELBOOL)
Public Event ColumnBeforeSize(ByVal oColumn As cColumn, ByRef bCancel As OLE_CANCELBOOL)
Public Event ContextMenu(ByVal X As Single, ByVal Y As Single)
Public Event Click(ByVal iButton As evbComCtlMouseButton)
Public Event ItemActivate(ByVal oItem As cListItem)
Public Event ItemAfterEdit(ByVal oItem As cListItem, ByRef bCancel As OLE_CANCELBOOL, ByRef sNew As String)
Public Event ItemBeforeEdit(ByVal oItem As cListItem, ByRef bCancel As OLE_CANCELBOOL)
Public Event ItemClick(ByVal oItem As cListItem, ByVal iButton As evbComCtlMouseButton)
Public Event ItemCheck(ByVal oItem As cListItem, ByVal bCheck As Boolean)
Public Event ItemDrag(ByVal oItem As cListItem, ByVal iButton As evbComCtlMouseButton)
Public Event ItemFocus(ByVal oItem As cListItem)
Public Event ItemSelect(ByVal oItem As cListItem, ByVal bSelect As Boolean)
Public Event KeyDown(iKeyCode As Integer, ByVal iState As evbComCtlKeyboardState, ByVal bRepeat As Boolean)
Public Event OLECompleteDrag(Effect As evbComCtlOleDropEffect)
Public Event OLEDragDrop(Data As DataObject, Effect As evbComCtlOleDropEffect, Button As evbComCtlMouseButton, Shift As evbComCtlKeyboardState, X As Single, Y As Single)
Public Event OLEDragOver(Data As DataObject, Effect As evbComCtlOleDropEffect, Button As evbComCtlMouseButton, Shift As evbComCtlKeyboardState, X As Single, Y As Single, State As evbComCtlOleDragOverState)
Public Event OLEGiveFeedback(Effect As evbComCtlOleDropEffect, DefaultCursors As Boolean)
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As evbComCtlOleDropEffect)

Private Type tColHeaderInfo
    iId                  As Long
    iSortOrder           As eListViewSortOrder
    iSortType            As eListViewSortType
    sKey                 As String
    sFormat              As String
End Type

Private Type tItemGroupInfo
    iId                 As Long
    sKey                As String
End Type

Private Const DEF_View           As Long = lvwDetails
Private Const DEF_BorderStyle    As Long = vbccBorderThin
Private Const DEF_Style          As Long = ZeroL
Private Const DEF_StyleEx        As Long = ZeroL
Private Const DEF_HeaderStyle    As Long = HDS_BUTTONS
Private Const DEF_IconSpace      As Long = NegOneL
Private Const DEF_Backcolor      As Long = vbWindowBackground
Private Const DEF_ForeColor      As Long = vbWindowText
Private Const DEF_GroupsEnabled As Boolean = True
Private Const DEF_Enabled        As Boolean = True
Private Const DEF_BackURL        As String = vbNullString
Private Const DEF_BackTile       As Boolean = True
Private Const DEF_BackX          As Long = ZeroL
Private Const DEF_BackY          As Long = ZeroL
Private Const DEF_ShowSort       As Boolean = False
Private Const DEF_Themeable      As Boolean = True
Private Const DEF_OleDrop        As Boolean = False
Private Const DEF_TileLines      As Long = 1

Private Const PROP_View         As String = "View"
Private Const PROP_BorderStyle  As String = "Border"
Private Const PROP_Style        As String = "Style"
Private Const PROP_StyleEx      As String = "StyleEx"
Private Const PROP_HeaderStyle  As String = "HdrStyle"
Private Const PROP_IconSpaceX   As String = "IconSpaceX"
Private Const PROP_IconSpaceY   As String = "IconSpaceY"
Private Const PROP_BackColor    As String = "BackColor"
Private Const PROP_ForeColor    As String = "ForeColor"
Private Const PROP_Enabled      As String = "Enabled"
Private Const PROP_Font         As String = "Font"
Private Const PROP_BackURL      As String = "BackURL"
Private Const PROP_BackTile     As String = "BackTile"
Private Const PROP_BackX        As String = "BackX"
Private Const PROP_BackY        As String = "BackY"
Private Const PROP_ShowSort     As String = "ShowSort"
Private Const PROP_Themeable    As String = "Themeable"
Private Const PROP_OleDrop      As String = "OleDrop"
Private Const PROP_TileLines    As String = "TileLines"

Private WithEvents moFont       As cFont
Attribute moFont.VB_VarHelpID = -1
Private WithEvents moFontPage   As pcSupportFontPropPage
Attribute moFontPage.VB_VarHelpID = -1

Private mhWnd                   As Long
Private moKeyMap                As pcStringMap
Private moItemDataMap           As pcIntegerMap

Private msBackgroundURL         As String

Private miBorderStyle           As evbComCtlBorderStyle
Private miViewStyle             As eListViewStyle
Private miStyle                 As Long
Private miStyleEx               As Long
Private miHeaderStyle           As Long
Private miIconSpaceX            As Long
Private miIconSpaceY            As Long
Private miStrPtr                As Long
Private miStrPtrCmp             As Long
Private miFocusIndex            As Long
Private mhFont                  As Long

Private moImageListLarge        As cImageList
Private WithEvents moImageListLargeEvent As cImageList
Attribute moImageListLargeEvent.VB_VarHelpID = -1

Private moImageListSmall        As cImageList
Private WithEvents moImageListSmallEvent As cImageList
Attribute moImageListSmallEvent.VB_VarHelpID = -1

Private moImageListHeader       As cImageList
Private WithEvents moImageListHeaderEvent As cImageList
Attribute moImageListHeaderEvent.VB_VarHelpID = -1

Private mtItem                  As LVITEM
Private mtTile                  As LVTILEINFO
Private mtCol                   As LVCOLUMN
Private mtGroup                 As LVGROUP
Private mtFind                  As LVFINDINFO
Private mtBack                  As LVBKIMAGE

Private mbInEdit                As Boolean
Private mbNoPropChange          As Boolean
Private mbGroupsEnabled         As Boolean
Private mbShowSortArrow         As Boolean
Private mbCCVer_GE_4_71         As Boolean
Private mbThemeable             As Boolean

Private mtColumns()             As tColHeaderInfo
Private miColumnCount           As Long
Private miColumnUbound          As Long
Private miColumnControl         As Long

Private miSortOrder             As eListViewSortOrder

Private mtItemGroups()          As tItemGroupInfo
Private miItemGroupCount        As Long
Private miItemGroupUbound       As Long
Private miItemGroupControl      As Long

Private miItemCount             As Long
Private miItemControl           As Long

Private miSortMsg               As Long

Private Const MAX_TEXT As Long = 512 '260 --> Maximal textnya m,enjadi 256
Private Const BufferLen As Long = MAX_TEXT \ 2
Private msTextBuffer            As String
Private msCompareBuffer         As String

Private miLastXMouseDown        As Long
Private miLastYMouseDown        As Long
Private mbRedraw                As Boolean

Private miTileViewItemLines     As Long

Private Const cColumns          As String = "cColumns"
Private Const cColumn           As String = "cColumn"
Private Const cItemGroup        As String = "cItemGroup"
Private Const cItemGroups       As String = "cItemGroups"
Private Const cListItem         As String = "cListItem"
Private Const cListItems        As String = "cListItems"
Private Const cListSubItem      As String = "cListSubItem"
'Private Const cSubItems         As String = "cSubItems"
Private Const ucListView        As String = "ucListView"

'structure offsets
Private Const NMHDR_idfrom = 4&
Private Const NMHDR_code = 8&
Private Const NMLISTVIEW_uNewState = 20&
Private Const NMLISTVIEW_uOldState = 24&
Private Const NMLISTVIEW_iItem = 12&
Private Const NMLISTVIEW_iSubItem = 16&
Private Const NMLVDISPINFO_LVITEM_LT_iItem = 16&
Private Const NMLVDISPINFO_LVITEM_LT_lpszText = 32&
Private Const NMLVDISPINFO_LVITEM_LT_cchTextMax = 36&
Private Const NMLVKEYDOWN_wVKey = 12&
Private Const NMLVGETINFOTIP_pszText = 16&
Private Const NMLVGETINFOTIP_cchTextMax = 20&
Private Const NMLVGETINFOTIP_iItem = 24&
Private Const NMITEMACTIVATE_iItem = 12&
Private Const NMHEADER_iItem = 12&
Private Const NMHEADER_pItem = 20&
Private Const HDITEM_iOrder = 32&
Private Const HDITEM_iMask = ZeroL
'Private Const HDITEM_cxy = 4&

Private Const HDI_WIDTH = 1&

Private Const LISTITEM_lpKey            As Long = ZeroL
Private Const LISTITEM_lpKeyNext        As Long = 4&
Private Const LISTITEM_iItemData        As Long = 8&
Private Const LISTITEM_iItemDataNext    As Long = 12&
Private Const LISTITEM_lpToolTip        As Long = 16&
Private Const LISTITEM_Len              As Long = 20&

Private Sub iOleControlVB_OnMnemonic(bHandled As Boolean, ByVal iShift As evbComCtlKeyboardState, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long): End Sub
Private Sub iOleControlVB_GetControlInfo(bHandled As Boolean, iAccelCount As Long, hAccelTable As Long, iFlags As Long)
    bHandled = True
    iFlags = vbccEatsReturn Or (-mbInEdit * (vbccEatsEscape))
End Sub

Private Sub iOleInPlaceActiveObjectVB_TranslateAccelerator(bHandled As Boolean, lReturn As Long, ByVal iShift As evbComCtlKeyboardState, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Eat the arrow keys, pageup/pagedown, home, end and forward the key to either the
    '             listview or its edit control.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        If uMsg = WM_KEYDOWN Or uMsg = WM_KEYUP Then
            Select Case wParam And &HFFFF&
            Case vbKeyPageUp To vbKeyDown, vbKeyReturn, vbKeyEscape
                If mbInEdit Then
                    SendMessage SendMessage(mhWnd, LVM_GETEDITCONTROL, ZeroL, ZeroL), uMsg, wParam, lParam
                Else
                    If (wParam And &HFFFF&) = vbKeyEscape Then Exit Sub
                    SendMessage mhWnd, uMsg, wParam, lParam
                End If
                bHandled = True
            End Select
        End If
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
    Dim liPtr         As Long
    Dim liTemp        As Long
    Dim loObject      As Object
    Dim lbCancel      As Boolean
    Dim lsNew         As String
    
    Select Case uMsg
    Case WM_NOTIFY
        bHandled = (hWnd = UserControl.hWnd)

        Select Case MemOffset32(lParam, NMHDR_code)
        Case LVN_ITEMCHANGED
            ' Edited the state of an item
            pItemChanged MemOffset32(lParam, NMLISTVIEW_iItem), _
            MemOffset32(lParam, NMLISTVIEW_uOldState), _
            MemOffset32(lParam, NMLISTVIEW_uNewState)
        
            'Case LVN_INSERTITEM
        
        Case LVN_DELETEITEM
            pDeleteItem MemOffset32(lParam, NMLISTVIEW_iItem)
            
        Case LVN_DELETEALLITEMS
            pDeleteAllItems
        
        Case LVN_BEGINLABELEDITA
            Set loObject = pItem(MemOffset32(lParam, NMLVDISPINFO_LVITEM_LT_iItem))
            
            ''debug.assert Not (loObject Is Nothing)
            
            If Not loObject Is Nothing Then
                RaiseEvent ItemBeforeEdit(loObject, lbCancel)
                mbInEdit = Not lbCancel
                lReturn = Abs(lbCancel)
                If mbInEdit Then OnControlInfoChanged Me
            End If
            
        Case LVN_ENDLABELEDITA
            Set loObject = pItem(MemOffset32(lParam, NMLVDISPINFO_LVITEM_LT_iItem))
            
            ''debug.assert Not (loObject Is Nothing)
            
            If Not loObject Is Nothing Then
                
                liPtr = MemOffset32(lParam, NMLVDISPINFO_LVITEM_LT_lpszText)
                
                If liPtr Then
                    lstrToStringA liPtr, lsNew
                    
                    RaiseEvent ItemAfterEdit(loObject, lbCancel, lsNew)
                    
                    lstrFromStringA liPtr, MemOffset32(lParam, NMLVDISPINFO_LVITEM_LT_cchTextMax), lsNew
                    
                    lReturn = Abs(Not lbCancel)
                End If
            End If
            
            mbInEdit = False
            OnControlInfoChanged Me
            
        Case LVN_COLUMNCLICK
            Set loObject = pColumn(MemOffset32(lParam, NMLISTVIEW_iSubItem))
            
            ''debug.assert Not (loObject Is Nothing)
            
            If Not loObject Is Nothing Then
                RaiseEvent ColumnClick(loObject)
            End If
            
        Case LVN_KEYDOWN
            liTemp = MemOffset32(lParam, NMLVKEYDOWN_wVKey)
            RaiseEvent KeyDown(liTemp And &HFFFF&, KBState(), CBool(liTemp And &H40000000))
            
        Case LVN_GETINFOTIPA
            pInfoTip MemOffset32(lParam, NMLVGETINFOTIP_iItem), _
            MemOffset32(lParam, NMLVGETINFOTIP_pszText), _
            MemOffset32(lParam, NMLVGETINFOTIP_cchTextMax)
            
        Case LVN_ITEMACTIVATE
            
            If mbCCVer_GE_4_71 Then
                Set loObject = pItem(MemOffset32(lParam, NMITEMACTIVATE_iItem))
                
            Else
                Set loObject = pItem(MemOffset32(lParam, NMHDR_idfrom))
                
            End If
            
            ''debug.assert Not (loObject Is Nothing)
            
            If Not loObject Is Nothing Then
                RaiseEvent ItemActivate(loObject)
            End If
            
        Case NM_CLICK
            If mbCCVer_GE_4_71 Then
                Set loObject = pItem(MemOffset32(lParam, NMITEMACTIVATE_iItem))
                
            Else
                Set loObject = pItem(MemOffset32(lParam, NMHDR_idfrom))
                
            End If
            
            If Not loObject Is Nothing Then
                RaiseEvent ItemClick(loObject, vbLeftButton)
            Else
                RaiseEvent Click(vbLeftButton)
            End If
            
        Case NM_RCLICK
            If mbCCVer_GE_4_71 Then
                Set loObject = pItem(MemOffset32(lParam, NMITEMACTIVATE_iItem))
                
            Else
                Set loObject = pItem(MemOffset32(lParam, NMHDR_idfrom))
                
            End If
            
            If Not loObject Is Nothing Then
                RaiseEvent ItemClick(loObject, vbRightButton)
            Else
                RaiseEvent Click(vbRightButton)
            End If
            
        Case NM_CUSTOMDRAW
            lReturn = CDRF_DODEFAULT
            
        Case LVN_BEGINDRAG
            Set loObject = pItem(MemOffset32(lParam, NMLISTVIEW_iItem))
            ''debug.assert Not loObject Is Nothing
            
            If Not loObject Is Nothing Then
                RaiseEvent ItemDrag(loObject, vbLeftButton)
            End If
            
        Case LVN_BEGINRDRAG
            Set loObject = pItem(MemOffset32(lParam, NMLISTVIEW_iItem))
            ''debug.assert Not loObject Is Nothing
            
            If Not loObject Is Nothing Then
                RaiseEvent ItemDrag(loObject, vbRightButton)
            End If
            
        Case HDN_BEGINTRACKA
            Set loObject = pColumn(MemOffset32(lParam, NMHEADER_iItem))
            
            ''debug.assert Not loObject Is Nothing
            
            If Not loObject Is Nothing Then
            
                RaiseEvent ColumnBeforeSize(loObject, lbCancel)
                bHandled = lbCancel
                lReturn = Abs(lbCancel)
            End If
        
        Case HDN_ENDTRACKA
            Set loObject = pColumn(MemOffset32(lParam, NMHEADER_iItem))
            
            ''debug.assert Not loObject Is Nothing
            
            If Not loObject Is Nothing Then
                RaiseEvent ColumnAfterSize(loObject)
            End If
            
        Case HDN_BEGINDRAG
            Set loObject = pColumn(MemOffset32(lParam, NMHEADER_iItem))
            
            ''debug.assert Not loObject Is Nothing
            
            If Not loObject Is Nothing Then
                RaiseEvent ColumnBeforeDrag(loObject, lbCancel)
                bHandled = lbCancel
                lReturn = Abs(lbCancel)
                If lbCancel Then ReleaseCapture
            End If

        Case HDN_ENDDRAG
            Set loObject = pColumn(MemOffset32(lParam, NMHEADER_iItem))
            
            ''debug.assert Not loObject Is Nothing
            
            If Not loObject Is Nothing Then
                liPtr = MemOffset32(lParam, NMHEADER_pItem)
                If liPtr Then
                    liTemp = MemOffset32(liPtr, HDITEM_iOrder) + OneL
                    RaiseEvent ColumnAfterDrag(loObject, liTemp, lbCancel)
                    If liTemp < OneL Then
                        liTemp = ZeroL
                    ElseIf liTemp > miColumnCount Then
                        liTemp = miColumnCount - OneL
                    Else
                        liTemp = liTemp - OneL
                    End If
                    
                    MemOffset32(liPtr, HDITEM_iOrder) = liTemp
                    
                Else
                    ''debug.assert False
                    RaiseEvent ColumnAfterDrag(loObject, NegOneL, lbCancel)
                    
                End If
                
                bHandled = lbCancel
                lReturn = Abs(lbCancel)
                
            End If
        
        Case HDN_ITEMCHANGEDA
            
            pHeaderItemChanged lParam
            
        End Select
        
    Case WM_SETFOCUS
        ActivateIPAO Me
        
    Case WM_MOUSEACTIVATE
        liTemp = GetFocus()
        If Not (liTemp = mhWnd Or liTemp = SendMessage(mhWnd, LVM_GETEDITCONTROL, ZeroL, ZeroL)) Then
            vbComCtlTlb.SetFocus UserControl.hWnd
            lReturn = MA_NOACTIVATE
            bHandled = True
        End If
        
    Case WM_CONTEXTMENU
        bHandled = True
        lReturn = ZeroL
        lParam = TranslateContextMenuCoords(mhWnd, lParam)
        RaiseEvent ContextMenu(UserControl.ScaleX(loword(lParam), vbPixels, vbContainerPosition), UserControl.ScaleY(hiword(lParam), vbPixels, vbContainerPosition))
    
    Case WM_PARENTNOTIFY
        bHandled = True
        If (wParam And &HFFFF&) = WM_RBUTTONDOWN Or (wParam And &HFFFF&) = WM_LBUTTONDOWN Then
            miLastXMouseDown = loword(lParam)
            miLastYMouseDown = hiword(lParam)
        End If
    End Select
End Sub

Private Function iLVCompare_String(ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Compare two items using a binary comparison of their text.
    '---------------------------------------------------------------------------------------
    pComp_GetText lParam1, lParam2, False
       
    iLVCompare_String = lstrcmp(miStrPtr, miStrPtrCmp)
    
    If miSortOrder = lvwSortDescending Then iLVCompare_String = -iLVCompare_String
    
End Function

Private Function iLVCompare_StringNoCase(ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Compare two items using a textual comparison of their text.
    '---------------------------------------------------------------------------------------
    pComp_GetText lParam1, lParam2, False
    
    iLVCompare_StringNoCase = lstrcmpi(miStrPtr, miStrPtrCmp)
    
    If miSortOrder = lvwSortDescending Then iLVCompare_StringNoCase = -iLVCompare_StringNoCase
    
End Function


Private Function iLVCompare_Date(ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Compare two items using a date comparison of their text.
    '---------------------------------------------------------------------------------------
    
    pComp_GetText lParam1, lParam2, True
    
    Dim ld1      As Date, ld2 As Date
    
    On Error Resume Next
    
    ld1 = CDate(msTextBuffer)
    ld2 = CDate(msCompareBuffer)
    
    On Error GoTo 0
    
    If ld1 > ld2 Then
        iLVCompare_Date = OneL
    ElseIf ld1 <> ld2 Then
        iLVCompare_Date = NegOneL
    Else
        iLVCompare_Date = ZeroL
    End If
    
    If miSortOrder = lvwSortDescending Then iLVCompare_Date = -iLVCompare_Date
    
End Function

Private Function iLVCompare_Selected(ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Compare two items based on their selected status.
    '---------------------------------------------------------------------------------------
    
    Dim lb1      As Boolean
    Dim lb2      As Boolean
    
    If miSortMsg <> LVM_SORTITEMSEX Then
        lb1 = pItem_State(pItem_IndexFromlParam(lParam1), LVIS_SELECTED)
        lb2 = pItem_State(pItem_IndexFromlParam(lParam2), LVIS_SELECTED)
    Else
        lb1 = pItem_State(lParam1, LVIS_SELECTED)
        lb2 = pItem_State(lParam2, LVIS_SELECTED)
    End If
    
    If lb1 Xor lb2 Then
        If lb2 _
            Then iLVCompare_Selected = OneL _
        Else iLVCompare_Selected = NegOneL
        End If
        If miSortOrder = lvwSortDescending Then iLVCompare_Selected = -iLVCompare_Selected
    
End Function

Private Function iLVCompare_Numeric(ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Compare two items based on the numeric value of their text.
    '---------------------------------------------------------------------------------------
    
    pComp_GetText lParam1, lParam2, True
    
    Dim lr1      As Double, lr2 As Double
    
    On Error Resume Next
    
    lr1 = CDbl(msTextBuffer)
    lr2 = CDbl(msCompareBuffer)
    
    On Error GoTo 0
    
    If lr1 > lr2 Then
        iLVCompare_Numeric = OneL
    ElseIf lr1 <> lr2 Then
        iLVCompare_Numeric = NegOneL
    Else
        iLVCompare_Numeric = ZeroL
    End If
    
    If miSortOrder = lvwSortDescending Then iLVCompare_Numeric = -iLVCompare_Numeric
    
End Function

Private Function iLVCompare_Indent(ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Compare two items based on their indentation level.
    '---------------------------------------------------------------------------------------
    Dim li1      As Long
    Dim li2      As Long
    
    If miSortMsg <> LVM_SORTITEMSEX Then
        li1 = pItem_Info(pItem_IndexFromlParam(lParam1), ZeroL, LVIF_INDENT)
        li2 = pItem_Info(pItem_IndexFromlParam(lParam2), ZeroL, LVIF_INDENT)
    Else
        li1 = pItem_Info(lParam1, ZeroL, LVIF_INDENT)
        li2 = pItem_Info(lParam2, ZeroL, LVIF_INDENT)
    End If
    
    If li1 > li2 Then
        iLVCompare_Indent = OneL
    ElseIf li1 <> li2 Then
        iLVCompare_Indent = NegOneL
    Else
        iLVCompare_Indent = ZeroL
    End If
    
    If miSortOrder = lvwSortDescending Then iLVCompare_Indent = -iLVCompare_Indent
    
End Function

Private Function iLVCompare_Currency(ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Compare two items based on the currency value of their text.
    '---------------------------------------------------------------------------------------
    
    pComp_GetText lParam1, lParam2, True
    
    Dim lc1      As Currency, lc2 As Currency
    
    On Error Resume Next
    
    lc1 = CCur(msTextBuffer)
    lc2 = CCur(msCompareBuffer)
    
    On Error GoTo 0
    
    If lc1 > lc2 Then
        iLVCompare_Currency = OneL
    ElseIf lc1 <> lc2 Then
        iLVCompare_Currency = NegOneL
    Else
        iLVCompare_Currency = ZeroL
    End If
    
    If miSortOrder = lvwSortDescending Then iLVCompare_Currency = -iLVCompare_Currency

End Function

Private Property Get pcListItem_lpKey(ByVal lpItem As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Return the strptr of the item key given its memory handle.
    '---------------------------------------------------------------------------------------
    ''debug.assert lpItem
    If lpItem Then
        pcListItem_lpKey = MemOffset32(lpItem, LISTITEM_lpKey)
    End If
End Property
'private Property Let pcListItem_lpKey(ByVal lpItem As Long, ByRef iNew As Long)
'    ''debug.assert lpItem
'    If lpItem Then
'        MemOffset32(lpItem, LISTITEM_lpKey) = iNew
'    End If
'End Property

Private Sub pcListItem_GetKey(ByVal lpItem As Long, ByRef sOut As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Return the item key given its memory handle.
    '---------------------------------------------------------------------------------------
    ''debug.assert lpItem
    
    If lpItem Then
        Dim lpKey      As Long
        lpKey = MemOffset32(lpItem, LISTITEM_lpKey)
        If lpKey Then lstrToStringA lpKey, sOut
    End If
End Sub

Private Function pcListItem_SetKey(ByVal lpItem As Long, ByRef sIn As String) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Allocate a string with the new key and store its pointer, maintain
    '             the keys collection.
    '---------------------------------------------------------------------------------------
    ''debug.assert lpItem
    
    If lpItem Then
        
        Dim lpKeyOld      As Long:   lpKeyOld = MemOffset32(lpItem, LISTITEM_lpKey)
        
        If LenB(sIn) Then
            
            If moKeyMap Is Nothing Then Set moKeyMap = New pcStringMap
            
            Dim lsAnsi As String:   lsAnsi = StrConv(sIn & vbNullChar, vbFromUnicode)
            Dim lpKeyNew As Long:   lpKeyNew = StrPtr(lsAnsi)
            Dim liHashNew As Long:  liHashNew = Hash(lpKeyNew, lstrlen(lpKeyNew))
            
            If moKeyMap.Find(lpKeyNew, liHashNew) = ZeroL Then
                
                lpKeyNew = MemAllocFromString(lpKeyNew, LenB(lsAnsi))
                
                If CBool(lpKeyNew) Then
                    pcListItem_SetKey = OneL
                    MemOffset32(lpItem, LISTITEM_lpKey) = lpKeyNew
                    moKeyMap.Add lpItem, liHashNew
                Else
                    pcListItem_SetKey = ZeroL
                End If
                
            Else
                pcListItem_SetKey = NegOneL
            End If
            
        Else
            
            pcListItem_SetKey = OneL
            MemOffset32(lpItem, LISTITEM_lpKey) = ZeroL
            
        End If
        
        If (pcListItem_SetKey = OneL) And CBool(lpKeyOld) Then
            moKeyMap.Remove lpItem, Hash(lpKeyOld, lstrlen(lpKeyOld))
            MemFree lpKeyOld
        End If
        
    End If
    
End Function

Private Function pcListItem_SetItemData(ByVal lpItem As Long, ByVal iItemData As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Store the itemdata as a unique identifier.
    '---------------------------------------------------------------------------------------
    ''debug.assert lpItem
    
    If lpItem Then
        
        Dim liItemDataOld      As Long:  liItemDataOld = MemOffset32(lpItem, LISTITEM_iItemData)
        Dim liHash             As Long
        
        If iItemData <> ZeroL Then
            If moItemDataMap Is Nothing Then Set moItemDataMap = New pcIntegerMap
            
            liHash = HashLong(iItemData)
            
            If moItemDataMap.Find(iItemData, liHash) = ZeroL _
                Then pcListItem_SetItemData = OneL _
            Else pcListItem_SetItemData = NegOneL
            
            Else
            
                pcListItem_SetItemData = OneL
            
            End If
        
            If (pcListItem_SetItemData = OneL) Then
                If liItemDataOld Then moItemDataMap.Remove lpItem, HashLong(liItemDataOld)
                MemOffset32(lpItem, LISTITEM_iItemData) = iItemData
                If iItemData Then moItemDataMap.Add lpItem, liHash
            End If
        
        End If
    
End Function

Private Property Get pcListItem_iItemData(ByVal lpItem As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Return the itemdata given the memory handle
    '---------------------------------------------------------------------------------------
    ''debug.assert lpItem
    If lpItem Then
        pcListItem_iItemData = MemOffset32(lpItem, LISTITEM_iItemData)
    End If
End Property


Private Property Get pcListItem_lpToolTip(ByVal lpItem As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Return the strptr of the item tooltip given the memory handle
    '---------------------------------------------------------------------------------------
    ''debug.assert lpItem
    If lpItem Then
        pcListItem_lpToolTip = MemOffset32(lpItem, LISTITEM_lpToolTip)
    End If
End Property
'private Property Let pcListItem_lpToolTip(ByVal lpItem As Long, ByRef iNew As Long)
'    ''debug.assert lpItem
'    If lpItem Then
'        MemOffset32(lpItem, LISTITEM_lpToolTip) = iNew
'    End If
'End Property

Private Sub pcListItem_GetToolTip(ByVal lpItem As Long, ByRef sOut As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Get a vb string with the tooltip of an item given the memory handle.
    '---------------------------------------------------------------------------------------
    ''debug.assert lpItem
    
    If lpItem Then
        Dim lpToolTip      As Long
        lpToolTip = MemOffset32(lpItem, LISTITEM_lpToolTip)
        If lpToolTip Then lstrToStringA lpToolTip, sOut
    End If
End Sub

Private Function pcListItem_SetToolTip(ByVal lpItem As Long, ByRef sIn As String) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Allocate a string and store the strptr.
    '---------------------------------------------------------------------------------------
    Dim lpToolTipOld      As Long
    Dim lpToolTip         As Long
    
    ''debug.assert lpItem
    
    If lpItem Then
        lpToolTipOld = MemOffset32(lpItem, LISTITEM_lpToolTip)
        
        If LenB(sIn) Then
            
            Dim lsAnsi      As String
            Dim liPtr       As Long
            
            lsAnsi = StrConv(sIn & vbNullChar, vbFromUnicode)
            liPtr = StrPtr(lsAnsi)
                    
            lpToolTip = MemAllocFromString(liPtr, LenB(lsAnsi))
            If lpToolTip Then pcListItem_SetToolTip = OneL
            
        Else
            pcListItem_SetToolTip = OneL
            
        End If
        
        If (pcListItem_SetToolTip = OneL) Then
            If CBool(lpToolTipOld) Then MemFree lpToolTipOld
            MemOffset32(lpItem, LISTITEM_lpToolTip) = lpToolTip
        End If
        
    End If
   
End Function


Private Function pcListItem_Alloc( _
ByRef sKey As String, _
ByRef sToolTipText As String, _
ByVal iItemData As Long) _
As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Allocate memory to store extra item data.
    '---------------------------------------------------------------------------------------
    
    pcListItem_Alloc = MemAlloc(LISTITEM_Len)
    
    If pcListItem_Alloc Then
        
        MemOffset32(pcListItem_Alloc, LISTITEM_lpKey) = ZeroL
        MemOffset32(pcListItem_Alloc, LISTITEM_lpToolTip) = ZeroL
        MemOffset32(pcListItem_Alloc, LISTITEM_iItemData) = ZeroL
        
        Dim liResult      As Long
        liResult = pcListItem_SetKey(pcListItem_Alloc, sKey)
        If liResult = OneL Then liResult = pcListItem_SetItemData(pcListItem_Alloc, iItemData)
        If liResult = OneL Then liResult = pcListItem_SetToolTip(pcListItem_Alloc, sToolTipText)
        
        If liResult <> OneL Then
            pcListItem_Free pcListItem_Alloc
            If liResult = ZeroL Then gErr vbccOutOfMemory, ucListView
            If liResult = NegOneL Then gErr vbccKeyAlreadyExists, ucListView
            ''debug.assert False
        End If
    End If
End Function

Private Sub pcListItem_Free(ByVal lpItem As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Free the memory associated with an item.
    '---------------------------------------------------------------------------------------
    ''debug.assert lpItem
    If lpItem Then
        Dim lp      As Long
        lp = MemOffset32(lpItem, LISTITEM_lpKey)
        
        If lp Then
            moKeyMap.Remove lpItem, Hash(lp, lstrlen(lp))
            MemFree lp
        End If
        
        lp = MemOffset32(lpItem, LISTITEM_lpToolTip)
        If lp Then MemFree lp
        
        lp = MemOffset32(lpItem, LISTITEM_iItemData)
        If lp Then moItemDataMap.Remove lpItem, HashLong(lp)
        
        MemFree lpItem
        
    End If
End Sub

Private Sub pHeaderItemChanged(ByVal lpHDNMHeader As Long)
    If Not CheckCCVersion(6&) And CBool(miHeaderStyle And HDS_FULLDRAG) Then
        Dim lpItem      As Long: lpItem = MemOffset32(lpHDNMHeader, NMHEADER_pItem)
        Debug.Print pColumn_Info(MemOffset32(lpHDNMHeader, NMHEADER_iItem), LVCF_FMT) And LVCFMT_IMAGE
        If lpItem Then
            If MemOffset32(lpItem, HDITEM_iMask) And HDI_WIDTH Then
                Dim lhWndHeader      As Long: lhWndHeader = SendMessage(mhWnd, LVM_GETHEADER, ZeroL, ZeroL)
                Dim ltRect           As RECT
                If lhWndHeader Then
                    If SendMessage(lhWndHeader, HDM_GETITEMRECT, MemOffset32(lpHDNMHeader, NMHEADER_iItem), VarPtr(ltRect)) Then
                        InvalidateRect lhWndHeader, ltRect, OneL
                        UpdateWindow lhWndHeader
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub pItemChanged(ByVal iItem As Long, ByVal iOld As Long, ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : If an item status has changed, check it for flags that we watch for:
    '             selected, focus and checked.
    '---------------------------------------------------------------------------------------
    
    If CBool(iNew And LVIS_SELECTED) Xor CBool(iOld And LVIS_SELECTED) Then RaiseEvent ItemSelect(pItem(iItem), CBool(iNew And LVIS_SELECTED))
    If CBool(iNew And LVIS_FOCUSED) Xor CBool(iOld And LVIS_FOCUSED) Then If CBool(iNew And LVIS_FOCUSED) Then pInFocus iItem
    If CBool(iNew And &H2000&) Xor CBool(iOld And &H2000&) Then RaiseEvent ItemCheck(pItem(iItem), CBool(iNew And &H2000&))
    
End Sub

Private Sub pInFocus(ByVal iIndex As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Raise a focuschanged event if the item in focus is new.
    '---------------------------------------------------------------------------------------
        
    Dim lbParamValid        As Boolean
    Dim lbModularValid      As Boolean
    
    lbParamValid = CBool(iIndex > NegOneL And iIndex < miItemCount)
    lbModularValid = CBool(miFocusIndex > NegOneL)
    
    If (lbParamValid And lbModularValid And (iIndex <> miFocusIndex)) _
        Or _
        (lbParamValid Xor lbModularValid) Then
        
        RaiseEvent ItemFocus(pItem(iIndex))
        miFocusIndex = iIndex
        
    End If
    
End Sub

Private Sub pDeleteItem(ByVal iIndex As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Free all resources associated with the given item.
    '---------------------------------------------------------------------------------------
    '''debug.assert (iIndex > NegOneL And iIndex < miItemCount)
    If iIndex > NegOneL And iIndex < miItemCount Then
        miItemCount = miItemCount - OneL
       
        pcListItem_Free pItem_Info(iIndex, ZeroL, LVIF_PARAM)
        
        If miItemCount = ZeroL Then
            pInFocus NegOneL
        ElseIf miFocusIndex > iIndex Then
            miFocusIndex = miFocusIndex - OneL
        End If
        
        Incr miItemControl
        
    End If
End Sub

Private Sub pDeleteAllItems()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Free all resources associated with all items.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim lhWnd      As Long
        lhWnd = mhWnd
        With mtItem
            .mask = LVIF_PARAM
            For .iItem = ZeroL To miItemCount - OneL
                If SendMessageAny(lhWnd, LVM_GETITEMA, ZeroL, .mask) Then
                    pcListItem_Free .lParam
                End If
            Next
        End With
        
    End If
    
    miItemCount = ZeroL
    pInFocus NegOneL
    Incr miItemControl
    
End Sub

Private Function pItem(ByVal iIndex As Long) As cListItem
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return a cListItem object representing a given item.
    '---------------------------------------------------------------------------------------
    '''debug.assert iIndex > NegOneL And iIndex < miItemCount
    Dim liPtr      As Long
    
    liPtr = pItem_Info(iIndex, ZeroL, LVIF_PARAM)
    
    If liPtr <> ZeroL Then
        Set pItem = New cListItem
        pItem.fInit Me, liPtr, iIndex
        
    End If
    
End Function


Private Function pColumn(ByVal iIndex As Long) As cColumn
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return a cColumn object representing a given item.
    '---------------------------------------------------------------------------------------
    '''debug.assert iIndex > NegOneL And iIndex < miColumnCount
    If iIndex > NegOneL And iIndex < miColumnCount Then
        Set pColumn = New cColumn
        pColumn.fInit Me, iIndex, mtColumns(iIndex).iId
    End If
End Function

Private Sub pInfoTip(ByVal iIndex As Long, ByVal iPtr As Long, ByVal iLen As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Store the infotip data in the NMLVGETINFOTIP structure.
    '---------------------------------------------------------------------------------------
    ''debug.assert iIndex > NegOneL And iIndex < miItemCount
    If iIndex > NegOneL And iIndex < miItemCount Then
        Dim lpToolTip      As Long
        Dim liLen          As Long
        
        lpToolTip = pcListItem_lpToolTip(pItem_Info(iIndex, ZeroL, LVIF_PARAM))
                
        If iPtr Then
            If lpToolTip Then
                liLen = lstrlen(lpToolTip) + 1
                If iLen < liLen Then liLen = iLen
                If liLen > ZeroL Then
                    CopyMemory ByVal iPtr, ByVal lpToolTip, liLen
                End If
            Else
                CopyMemory ByVal iPtr, 0, 2&
            End If
        End If
    
    End If
End Sub




Friend Function fColumns_Add( _
ByRef sKey As String, _
ByRef sText As String, _
ByVal iIcon As Long, _
ByVal iSortType As eListViewSortType, _
ByVal iAlignment As eListViewColumnAlign, _
ByVal fWidth As Single, _
ByRef sFormat As String, _
ByRef vColumnBefore As Variant) _
As cColumn
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Add a column to the listview.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
   
        Dim liIndex      As Long
        Dim i            As Long
    
        Dim iWidth As Long
        If fWidth > NegOneF _
            Then iWidth = ScaleX(fWidth, vbContainerSize, vbPixels) _
        Else iWidth = 96
    
            If LenB(sKey) <> ZeroL Then
                If pColumns_FindKey(sKey) <> NegOneL Then gErr vbccKeyAlreadyExists, cColumns
            End If
    
            If Not IsMissing(vColumnBefore) Then
                liIndex = pColumns_GetIndex(vColumnBefore)
                If liIndex = NegOneL Then gErr vbccKeyOrIndexNotFound, cColumns
        
            Else
                liIndex = miColumnCount
        
            End If
    
            With mtCol
        
                .mask = LVCF_FMT
                .fmt = ZeroL
        
                If iWidth > ZeroL Then
                    .mask = .mask Or LVCF_WIDTH
                    .cx = iWidth
                End If
        
                If iIcon > NegOneL Then
                    .fmt = .fmt Or LVCFMT_IMAGE
                    .mask = .mask Or LVCF_IMAGE
                    .iImage = iIcon
                End If
        
                If iAlignment <> lvwAlignLeft Then
                    .fmt = .fmt Or (iAlignment And LVCFMT_JUSTIFYMASK)
                End If
        
                If LenB(sText) Then
                    lstrFromStringW .pszText, .cchTextMax, sText
                    .mask = .mask Or LVCF_TEXT
                End If
        
            End With
    
            i = SendMessage(mhWnd, LVM_INSERTCOLUMNW, liIndex, VarPtr(mtCol.mask))
            ''debug.assert i = liIndex
            liIndex = i
            ''debug.assert liIndex > NegOneL And liIndex <= miColumnCount
    
            If liIndex > NegOneL And liIndex <= miColumnCount Then
        
                Incr miColumnControl
        
                i = RoundToInterval(miColumnCount)
        
                If i > miColumnUbound Then
                    ReDim Preserve mtColumns(ZeroL To i)
                    miColumnUbound = i
                End If
        
                If liIndex < miColumnCount Then
                    With mtColumns(miColumnCount)
                        .sKey = vbNullString
                        .sFormat = vbNullString
                    End With
                    CopyMemory mtColumns(liIndex + OneL).iId, mtColumns(liIndex).iId, (Len(mtColumns(ZeroL)) * (miColumnCount - liIndex))
                    ZeroMemory mtColumns(liIndex).iId, Len(mtColumns(ZeroL))
                End If
        
                miColumnCount = miColumnCount + OneL
        
                With mtColumns(liIndex)
                    .sKey = sKey
                    .iId = NextItemId()
                    .sFormat = sFormat
                    .iSortType = iSortType
            
                    Set fColumns_Add = New cColumn
                    fColumns_Add.fInit Me, liIndex, .iId
        
                End With
            End If
        End If
End Function

Friend Sub fColumns_Clear()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Remove all columns from the listview.
    '---------------------------------------------------------------------------------------
    On Error Resume Next
    Dim i      As Long
    For i = fColumns_Count - 1& To ZeroL Step NegOneL
        fColumns_Remove i
    Next
    On Error GoTo 0
End Sub

Friend Property Get fColumns_Count() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the current number of columns.
    '---------------------------------------------------------------------------------------
    fColumns_Count = miColumnCount
    ''debug.assert fColumns_Count = SendMessage(SendMessage(mhWnd, LVM_GETHEADER, ZeroL, ZeroL), HDM_GETITEMCOUNT, ZeroL, ZeroL)
End Property

Friend Property Get fColumns_Item(ByRef vColumn As Variant) As cColumn
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return a cColumn object representing the given column.
    '---------------------------------------------------------------------------------------
    Dim i      As Long
    
    i = pColumns_GetIndex(vColumn)
     
    If i <> NegOneL Then
        Set fColumns_Item = New cColumn
        fColumns_Item.fInit Me, i, mtColumns(i).iId
        
    Else
        gErr vbccKeyOrIndexNotFound, cColumns
        
    End If
     
End Property

Friend Property Get fColumns_Exists(ByRef vColumn As Variant) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return a value indicating whether a column exists in the collection.
    '---------------------------------------------------------------------------------------
    fColumns_Exists = pColumns_GetIndex(vColumn) <> NegOneL
End Property

Friend Sub fColumns_Remove(ByRef vColumn As Variant)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Remove a column from the listview.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim i      As Long
    
        i = pColumns_GetIndex(vColumn)
    
        If i <> NegOneL Then
            SendMessage mhWnd, LVM_DELETECOLUMN, i, ZeroL
            Incr miColumnControl
        
            miColumnCount = miColumnCount - OneL
            If i < miColumnCount Then
                With mtColumns(i)
                    .sKey = vbNullString
                    .sFormat = vbNullString
                End With
                CopyMemory mtColumns(i).iId, mtColumns(i + OneL).iId, LenB(mtColumns(0)) * (miColumnCount - i)
                ZeroMemory mtColumns(miColumnCount).iId, Len(mtColumns(0))
            End If
        
        Else
            gErr vbccKeyOrIndexNotFound, cColumns
        
        End If
    End If
End Sub

Friend Property Get fColumns_Control() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return an identifier for the current column collection.
    '---------------------------------------------------------------------------------------
    fColumns_Control = miColumnControl
End Property

Friend Sub fColumns_NextItem(ByRef tEnum As tEnum, ByRef vNextItem As Variant, ByRef bNoMore As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the next cColumn object in an enumeration.
    '---------------------------------------------------------------------------------------
    If tEnum.iControl <> miColumnControl Then gErr vbccCollectionChangedDuringEnum, cColumns
    tEnum.iIndex = tEnum.iIndex + OneL
    If tEnum.iIndex > NegOneL And tEnum.iIndex < miColumnCount Then
        Set vNextItem = pColumn(tEnum.iIndex)
    Else
        bNoMore = True
    End If
End Sub

Private Function pColumns_GetIndex(ByRef vColumn As Variant) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return a column index given its key or index.
    '---------------------------------------------------------------------------------------
    If VarType(vColumn) = vbString Then
        pColumns_GetIndex = pColumns_FindKey(CStr(vColumn))
        
    ElseIf VarType(vColumn) = vbObject Then
        Dim loCol      As cColumn
        On Error Resume Next
        Set loCol = vColumn
        If loCol.fIsOwner(Me) Then
            pColumns_GetIndex = loCol.Index
        End If
        pColumns_GetIndex = pColumns_GetIndex - OneL
        On Error GoTo 0
        
    Else
        On Error Resume Next
        pColumns_GetIndex = CLng(vColumn)
        pColumns_GetIndex = pColumns_GetIndex - OneL
        On Error GoTo 0
        
        If pColumns_GetIndex < ZeroL Or pColumns_GetIndex >= miColumnCount Then pColumns_GetIndex = NegOneL
        
    End If
End Function

Private Function pColumns_FindKey(ByRef sKey As String) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Find a key in the column collection.
    '---------------------------------------------------------------------------------------
    If LenB(sKey) Then
        For pColumns_FindKey = ZeroL To miColumnCount - OneL
            If StrComp(mtColumns(pColumns_FindKey).sKey, sKey) = ZeroL Then Exit Function
        Next
    End If
    pColumns_FindKey = NegOneL
End Function

Friend Property Get fColumn_Width(ByRef iIndex As Long, ByVal iId As Long) As Single
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the width of the column.
    '---------------------------------------------------------------------------------------
    If pColumn_Verify(iIndex, iId) Then
        fColumn_Width = ScaleX(pColumn_Info(iIndex, LVCF_WIDTH), vbPixels, vbContainerSize)
    End If
End Property

Friend Property Let fColumn_Width(ByRef iIndex As Long, ByVal iId As Long, ByVal fNew As Single)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the width of the column.
    '---------------------------------------------------------------------------------------
    If pColumn_Verify(iIndex, iId) Then
        pColumn_Info(iIndex, LVCF_WIDTH) = ScaleX(fNew, vbContainerSize, vbPixels)
    End If
End Property

Friend Property Get fColumn_Text(ByRef iIndex As Long, ByVal iId As Long) As String
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the text of the column.
    '---------------------------------------------------------------------------------------
    If pColumn_Verify(iIndex, iId) Then
        pColumn_GetText iIndex, fColumn_Text
    End If
End Property
Friend Property Let fColumn_Text(ByRef iIndex As Long, ByVal iId As Long, ByRef sNew As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the text of the column.
    '---------------------------------------------------------------------------------------
    If pColumn_Verify(iIndex, iId) Then
        pColumn_PutText iIndex, sNew
    End If
End Property

Friend Property Get fColumn_Key(ByRef iIndex As Long, ByVal iId As Long) As String
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the column key.
    '---------------------------------------------------------------------------------------
    If pColumn_Verify(iIndex, iId) Then
        fColumn_Key = mtColumns(iIndex).sKey
    End If
End Property
Friend Property Let fColumn_Key(ByRef iIndex As Long, ByVal iId As Long, ByRef sNew As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the column key.
    '---------------------------------------------------------------------------------------
    If pColumn_Verify(iIndex, iId) Then
        If pColumns_FindKey(sNew) = NegOneL Then
            mtColumns(iIndex).sKey = sNew
        Else
            gErr vbccKeyAlreadyExists, cColumn
        End If
    End If
End Property

Friend Property Get fColumn_IconIndex(ByRef iIndex As Long, ByVal iId As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the column icon index.
    '---------------------------------------------------------------------------------------
    If pColumn_Verify(iIndex, iId) Then
        fColumn_IconIndex = pColumn_Info(iIndex, LVCF_IMAGE)
    End If
End Property

Friend Property Let fColumn_IconIndex(ByRef iIndex As Long, ByVal iId As Long, ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the column icon index.
    '---------------------------------------------------------------------------------------
    If pColumn_Verify(iIndex, iId) Then
        pColumn_Info(iIndex, LVCF_IMAGE) = iNew
    End If
End Property

Friend Property Get fColumn_ImageOnRight(ByRef iIndex As Long, ByVal iId As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a value indicating whether the column displays its icon on the right.
    '---------------------------------------------------------------------------------------
    If pColumn_Verify(iIndex, iId) Then
        fColumn_ImageOnRight = pColumn_Format(iIndex, LVCFMT_BITMAP_ON_RIGHT)
    End If
End Property
Friend Property Let fColumn_ImageOnRight(ByRef iIndex As Long, ByVal iId As Long, ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether the column displays its image on the right.
    '---------------------------------------------------------------------------------------
    If pColumn_Verify(iIndex, iId) Then
        pColumn_Format(iIndex, LVCFMT_BITMAP_ON_RIGHT) = bNew
    End If
End Property

Friend Property Get fColumn_Alignment(ByRef iIndex As Long, ByVal iId As Long) As eListViewColumnAlign
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the column alignment.
    '---------------------------------------------------------------------------------------
    If pColumn_Verify(iIndex, iId) Then
        fColumn_Alignment = (pColumn_GetFormat(iIndex) And LVCFMT_JUSTIFYMASK)
    End If
End Property
Friend Property Let fColumn_Alignment(ByRef iIndex As Long, ByVal iId As Long, ByVal iNew As eListViewColumnAlign)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the column alignment.
    '---------------------------------------------------------------------------------------
    If pColumn_Verify(iIndex, iId) Then
        pColumn_SetFormat iIndex, LVCFMT_JUSTIFYMASK, (iNew And LVCFMT_JUSTIFYMASK)
        Refresh
    End If
End Property

Friend Property Get fColumn_SortOrder(ByRef iIndex As Long, ByVal iId As Long) As eListViewSortOrder
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the default sort order for the given client.
    '---------------------------------------------------------------------------------------
    If pColumn_Verify(iIndex, iId) Then
        fColumn_SortOrder = mtColumns(iIndex).iSortOrder
    End If
End Property
Friend Property Let fColumn_SortOrder(ByRef iIndex As Long, ByVal iId As Long, ByVal iSortOrder As eListViewSortOrder)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the default sort order for the column.
    '---------------------------------------------------------------------------------------
    If pColumn_Verify(iIndex, iId) Then
        mtColumns(iIndex).iSortOrder = iSortOrder
    End If
End Property

Friend Property Get fColumn_SortType(ByRef iIndex As Long, ByVal iId As Long) As eListViewSortType
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the default sort type of the column.
    '---------------------------------------------------------------------------------------
    If pColumn_Verify(iIndex, iId) Then
        fColumn_SortType = mtColumns(iIndex).iSortType
    End If
End Property
Friend Property Let fColumn_SortType(ByRef iIndex As Long, ByVal iId As Long, ByVal iNew As eListViewSortType)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the default sort type of the column.
    '---------------------------------------------------------------------------------------
    If pColumn_Verify(iIndex, iId) Then
        mtColumns(iIndex).iSortType = iNew
    End If
End Property

Friend Sub fColumn_AutoSize(ByRef iIndex As Long, ByVal iId As Long, ByVal iSize As eListViewColumnAutoSize)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Size the column to fit the widest text item or the header text.
    '---------------------------------------------------------------------------------------
    If pColumn_Verify(iIndex, iId) Then
        If iSize <> lvwColumnSizeToColumnText Then iSize = lvwColumnSizeToItemText
        If mhWnd Then
            SendMessage mhWnd, LVM_SETCOLUMNWIDTH, iIndex, iSize
        End If
    End If
End Sub

Friend Sub fColumn_Sort(ByRef iIndex As Long, ByVal iId As Long, ByVal iType As eListViewSortType, ByVal iOrder As eListViewSortOrder)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Sort the listview by this column.
    '---------------------------------------------------------------------------------------
    If pColumn_Verify(iIndex, iId) Then
        If mhWnd Then
            mtItem.mask = LVIF_TEXT
            mtItem.iSubItem = iIndex
            
            Dim loComp      As iLVCompare: Set loComp = Me
            Dim liPtr       As Long:        liPtr = ObjPtr(loComp)
            
            msCompareBuffer = String$(BufferLen, ZeroL)
            miStrPtrCmp = StrPtr(msCompareBuffer)
            
            ''debug.assert CBool(miStrPtrCmp) And CBool(miStrPtr)
            ''debug.assert miStrPtr = StrPtr(msTextBuffer) And miStrPtrCmp = StrPtr(msCompareBuffer)
            
            If CBool(miStrPtrCmp) And CBool(miStrPtr) Then
                With mtColumns(iIndex)
                    
                    If iType = NegOneL Then iType = .iSortType
                    
                    If iOrder = NegOneL Then
                        iOrder = (.iSortOrder Mod TwoL) + OneL
                        .iSortOrder = iOrder
                    End If
                    
                    pUpdateSortArrow iIndex, IIf(iOrder = lvwSortDescending, HDF_SORTUP, HDF_SORTDOWN)
                    
                    miSortOrder = iOrder
                    
                    If CheckCCVersion(5&, 8&) _
                        Then miSortMsg = LVM_SORTITEMSEX _
                    Else miSortMsg = LVM_SORTITEMS
                    
                        Dim lpCompareProc      As Long
                        lpCompareProc = Thunk_Alloc(tnkLVCompareProc)
                    
                        ''debug.assert lpCompareProc
                    
                        If lpCompareProc Then
                            If iType < lvwSortString Or iType > lvwSortSelected Then iType = lvwSortString: ''debug.assert False
                                Const PATCH_VTableOffset As Long = &H18
                                MemOffset8(lpCompareProc, PATCH_VTableOffset) = MemOffset8(lpCompareProc, PATCH_VTableOffset) + (iType * 4)
                                SendMessage mhWnd, miSortMsg, liPtr, lpCompareProc
                        
                                MemFree lpCompareProc
                            End If
                    
                        End With
                
                        msCompareBuffer = vbNullString
                        miStrPtrCmp = ZeroL
                        mtItem.pszText = miStrPtr
                
                        For liPtr = ZeroL To miColumnCount - OneL
                            If liPtr <> iIndex Then mtColumns(liPtr).iSortOrder = lvwSortNone
                        Next
                
                    End If
                End If
            End If
End Sub

Friend Property Get fColumn_Position(ByRef iIndex As Long, ByVal iId As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the display position of the column.
    '---------------------------------------------------------------------------------------
    If pColumn_Verify(iIndex, iId) Then
        fColumn_Position = pColumn_Info(iIndex, LVCF_ORDER)
    End If
End Property
Friend Property Let fColumn_Position(ByRef iIndex As Long, ByVal iId As Long, ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the display position of the column.
    '---------------------------------------------------------------------------------------
    If pColumn_Verify(iIndex, iId) Then
        If iNew < OneL Or iNew >= miColumnCount Then gErr vbccInvalidProcedureCall, cColumn
        pColumn_Info(iIndex, LVCF_ORDER) = iNew - OneL
        Refresh
    End If
End Property

Friend Property Get fColumn_Format(ByRef iIndex As Long, ByVal iId As Long) As String
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the default format used for the column text.
    '---------------------------------------------------------------------------------------
    If pColumn_Verify(iIndex, iId) Then
        fColumn_Format = mtColumns(iIndex).sFormat
    End If
End Property
Friend Property Let fColumn_Format(ByRef iIndex As Long, ByVal iId As Long, ByRef sNew As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the default format used for the column text.
    '---------------------------------------------------------------------------------------
    If pColumn_Verify(iIndex, iId) Then
        mtColumns(iIndex).sFormat = sNew
    End If
End Property

Friend Property Get fColumn_Index(ByRef iIndex As Long, ByVal iId As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the index of the column in the collection.
    '---------------------------------------------------------------------------------------
    If pColumn_Verify(iIndex, iId) Then
        fColumn_Index = iIndex + OneL
    End If
End Property

Private Function pColumn_Verify(ByRef iIndex As Long, ByVal iId As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Verify that a given column is still part of the collection.
    '---------------------------------------------------------------------------------------

    If iIndex > NegOneL And iIndex < miColumnCount _
        Then pColumn_Verify = CBool(mtColumns(iIndex).iId = iId)
    
        If Not pColumn_Verify Then
        
            For iIndex = ZeroL To miColumnCount - OneL
                pColumn_Verify = CBool(mtColumns(iIndex).iId = iId)
                If pColumn_Verify Then Exit For
            Next
        
            If Not pColumn_Verify Then
                iIndex = NegOneL
                gErr vbccItemDetached, cColumn
            End If
        
        End If
    
End Function

Private Property Get pColumn_Info(ByVal iIndex As Long, ByVal iInfo As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return information from the LVCOLUMN structure.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        With mtCol
            .mask = iInfo
            SendMessage mhWnd, LVM_GETCOLUMNA, iIndex, VarPtr(.mask)
            If iInfo = LVCF_IMAGE Then
                pColumn_Info = .iImage
            ElseIf iInfo = LVCF_ORDER Then
                pColumn_Info = .iOrder
            ElseIf iInfo = LVCF_WIDTH Then
                pColumn_Info = .cx
            ElseIf iInfo = LVCF_FMT Then
                pColumn_Info = .fmt
            End If
        End With
    End If
End Property
Private Property Let pColumn_Info(ByVal iIndex As Long, ByVal iInfo As Long, ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set information in the LVCOLUMN structure.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        With mtCol
            .mask = iInfo
            If iInfo = LVCF_IMAGE Then
                .iImage = iNew
            ElseIf iInfo = LVCF_ORDER Then
                .iOrder = iNew
            ElseIf iInfo = LVCF_WIDTH Then
                .cx = iNew
            ElseIf iInfo = LVCF_FMT Then
                .fmt = iNew
            End If
            SendMessage mhWnd, LVM_SETCOLUMNA, iIndex, VarPtr(.mask)
        End With
    End If
End Property

Private Property Get pColumn_Format(ByVal iIndex As Long, ByVal iFormat As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return whether a bit is set in the fmt member of the LVCOLUMN structure.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        With mtCol
            .mask = LVCF_FMT
            SendMessage mhWnd, LVM_GETCOLUMNA, iIndex, VarPtr(.mask)
            pColumn_Format = CBool(.fmt And iFormat)
        End With
    End If
End Property
Private Property Let pColumn_Format(ByVal iIndex As Long, ByVal iFormat As Long, ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set a bit in the fmt member of the LVCOLUMN structure.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        With mtCol
            .mask = LVCF_FMT
            SendMessage mhWnd, LVM_GETCOLUMNA, iIndex, VarPtr(.mask)
            If bNew Then
                .fmt = .fmt Or iFormat
            Else
                .fmt = .fmt And Not iFormat
            End If
            SendMessage mhWnd, LVM_SETCOLUMNA, iIndex, VarPtr(.mask)
        End With
    End If
End Property

Private Function pColumn_GetFormat(ByVal iIndex As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the fmt member of the LVCOLUMN structure for the given column.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        With mtCol
            .mask = LVCF_FMT
            SendMessage mhWnd, LVM_GETCOLUMNA, iIndex, VarPtr(.mask)
            pColumn_GetFormat = .fmt
        End With
    End If
End Function
Private Sub pColumn_SetFormat(ByVal iIndex As Long, ByVal iMask As Long, ByVal iFormat As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the fmt member of the LVCOLUMN structure for the given column.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        With mtCol
            .mask = LVCF_FMT
            SendMessage mhWnd, LVM_GETCOLUMNA, iIndex, VarPtr(.mask)
            .fmt = (.fmt And Not iMask) Or iFormat
            SendMessage mhWnd, LVM_SETCOLUMNA, iIndex, VarPtr(.mask)
        End With
    End If
End Sub


Private Sub pColumn_PutText(ByVal iIndex As Long, ByRef sIn As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the column text.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
    
        With mtCol
            .mask = LVCF_TEXT
            lstrFromStringW .pszText, .cchTextMax, sIn
            SendMessage mhWnd, LVM_SETCOLUMNW, iIndex, VarPtr(.mask)
        End With
    
    End If
End Sub

Private Sub pColumn_GetText(ByVal iIndex As Long, ByRef sOut As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the column text.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
    
        With mtCol
            .mask = LVCF_TEXT
            SendMessage mhWnd, LVM_GETCOLUMNW, iIndex, VarPtr(.mask)
            lstrToStringW .pszText, sOut
        End With

    End If
End Sub



Friend Property Get fItemGroup_Text(ByRef iIndex As Long, ByVal iId As Long) As String
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the text of an item group.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        If pItemGroup_Verify(iIndex, iId) Then
            With mtGroup
                .mask = LVGF_HEADER
                SendMessage mhWnd, LVM_GETGROUPINFO, iId, VarPtr(mtGroup)
                lstrToStringW .pszHeader, fItemGroup_Text
            End With
        End If
    End If
End Property

'Friend Property Let fItemGroup_Text(ByRef iIndex as long, ByVal iId as long, ByRef sNew As String)
'if mhWnd then
'    If pItemGroup_Verify(iIndex, iId) Then
'        With mtGroup
'            .cbSize = Len(mtGroup)
'            .cchHeader = Len(sNew)
'            .pszHeader = StrPtr(sNew)
'            .mask = LVGF_HEADER
'        End With
'
'        SendMessage mhWnd, LVM_SETGROUPINFO, tPointer.iID, VarPtr(mtGroup)
'
'    End If
'End If
'End Property


Friend Property Get fItemGroup_Key(ByRef iIndex As Long, ByVal iId As Long) As String
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the key of an item group.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        If pItemGroup_Verify(iIndex, iId) Then
            fItemGroup_Key = mtItemGroups(iIndex).sKey
        End If
    End If
End Property
Friend Property Let fItemGroup_Key(ByRef iIndex As Long, ByVal iId As Long, ByVal sNew As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the key of an item group.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        If pItemGroup_Verify(iIndex, iId) Then
            If pItemGroups_FindKey(sNew) <> NegOneL Then gErr vbccKeyAlreadyExists, cItemGroup
            mtItemGroups(iIndex).sKey = sNew
        End If
    End If
End Property

Friend Property Get fItemGroup_Index(ByRef iIndex As Long, ByVal iId As Long) As String
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the Index of an item group.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        If pItemGroup_Verify(iIndex, iId) Then
            fItemGroup_Index = iIndex + OneL
        End If
    End If
End Property


Private Function pItemGroup_Verify(ByRef iIndex As Long, ByVal iId As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Verify that an item group is still part of the collection.
    '---------------------------------------------------------------------------------------
    Dim i      As Long
    i = iIndex
    
    If i > NegOneL And i < miItemGroupCount Then
        pItemGroup_Verify = CBool(mtItemGroups(i).iId = iId)
        
    End If
    
    If Not pItemGroup_Verify Then
        
        For i = ZeroL To miItemGroupCount - OneL
            pItemGroup_Verify = CBool(mtItemGroups(i).iId = iId)
            If pItemGroup_Verify Then Exit For
        Next
        
        If pItemGroup_Verify _
            Then iIndex = i _
        Else iIndex = NegOneL
        
        End If
    
        If Not pItemGroup_Verify Then gErr vbccItemDetached, cItemGroup
    
End Function






Friend Function fItemGroups_Add( _
ByRef sKey As String, _
ByRef sText As String, _
ByRef vInsertBefore As Variant) _
As cItemGroup
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Add an item group to the collection.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim liId             As Long
        Dim liNewUbound      As Long
        Dim liIndex          As Long
    
        liId = NextItemId()
    
        If Len(sKey) Then
            If pItemGroups_FindKey(sKey) <> NegOneL Then gErr vbccKeyAlreadyExists, cItemGroups
        End If
    
        If IsMissing(vInsertBefore) Then
            liIndex = miItemGroupCount
        Else
            liIndex = pItemGroups_GetIndex(vInsertBefore)
            If liIndex = NegOneL Then gErr vbccKeyOrIndexNotFound, cItemGroups
        End If
    
        With mtGroup
            lstrFromStringW .pszHeader, .cchHeader, sText
        
            .iGroupId = liId
            .mask = LVGF_HEADER Or LVGF_GROUPID
        End With
    
        liIndex = SendMessage(mhWnd, LVM_INSERTGROUP, liIndex, VarPtr(mtGroup.cbSize))
    
        ''debug.assert liIndex > NegOneL And liIndex <= miItemGroupCount
    
        If liIndex > NegOneL And liIndex <= miItemGroupCount Then
            Incr miItemGroupControl
        
            liNewUbound = RoundToInterval(miItemGroupCount)
        
            If liNewUbound > miItemGroupUbound Then
                ReDim Preserve mtItemGroups(0 To liNewUbound)
                miItemGroupUbound = liNewUbound
            End If
        
            If liIndex < miItemGroupCount Then
                mtItemGroups(miItemGroupCount).sKey = vbNullString
                CopyMemory mtItemGroups(liIndex + OneL).iId, mtItemGroups(liIndex).iId, (Len(mtItemGroups(0)) * (miItemGroupCount - liIndex))
                ZeroMemory mtItemGroups(liIndex).iId, Len(mtItemGroups(0))
            
            End If
        
            With mtItemGroups(liIndex)
                .iId = liId
                .sKey = sKey
            End With
        
            Set fItemGroups_Add = New cItemGroup
            fItemGroups_Add.fInit Me, liId, miItemGroupCount
        
            miItemGroupCount = miItemGroupCount + OneL
    
        End If
    End If
End Function
   
Friend Property Get fItemGroups_Count() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the number of groups in the list.
    '---------------------------------------------------------------------------------------
    fItemGroups_Count = miItemGroupCount
End Property

Friend Sub fItemGroups_Clear()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Remove all groups in the list.
    '---------------------------------------------------------------------------------------
    ' Clear all the data associated with existing groups:
    If mhWnd Then
        SendMessage mhWnd, LVM_REMOVEALLGROUPS, 0, 0
    End If

    miItemGroupCount = ZeroL
    
End Sub

Friend Sub fItemGroups_Remove(ByRef vGroup As Variant)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Remove a group from the list.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim liIndex      As Long
    
        liIndex = pItemGroups_GetIndex(vGroup)
        If liIndex = NegOneL Then gErr vbccKeyOrIndexNotFound, cItemGroups
    
        SendMessage mhWnd, LVM_REMOVEGROUP, mtItemGroups(liIndex).iId, 0
    
        Incr miItemGroupControl
    
        If liIndex < miItemGroupCount Then
            mtItemGroups(liIndex).sKey = vbNullString
            CopyMemory mtItemGroups(liIndex).iId, mtItemGroups(liIndex + OneL).iId, Len(mtItemGroups(0)) * (miItemGroupCount - liIndex)
            ZeroMemory mtItemGroups(liIndex).iId, Len(mtItemGroups(0))
        End If
    
        miItemGroupCount = miItemGroupCount - OneL
    
    End If
End Sub
Friend Property Get fItemGroups_Enabled() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return whether item groups are enabled.
    '---------------------------------------------------------------------------------------
    fItemGroups_Enabled = mbGroupsEnabled
End Property
Friend Property Let fItemGroups_Enabled(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether item groups are enabled.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        SendMessage mhWnd, LVM_ENABLEGROUPVIEW, Abs(bNew), 0
        mbGroupsEnabled = bNew
    End If
End Property

Friend Property Get fItemGroups_Item(ByRef vGroup As Variant) As cItemGroup
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return a cItemGroup object representing the given group.
    '---------------------------------------------------------------------------------------
    Dim liIndex      As Long
    liIndex = pItemGroups_GetIndex(vGroup)
    If liIndex = NegOneL Then gErr vbccKeyOrIndexNotFound, cItemGroups
    Set fItemGroups_Item = New cItemGroup
    fItemGroups_Item.fInit Me, mtItemGroups(liIndex).iId, liIndex
End Property

Friend Property Get fItemGroups_Exists(ByRef vGroup As Variant) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return a value indicating whether the given group exists in the collection.
    '---------------------------------------------------------------------------------------
    fItemGroups_Exists = pItemGroups_GetIndex(vGroup) <> NegOneL
End Property

Friend Property Get fItemGroups_Control() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return an identity value for the current itemgroups collection.
    '---------------------------------------------------------------------------------------
    fItemGroups_Control = miItemGroupControl
End Property

Friend Sub fItemGroups_NextItem(ByRef tEnum As tEnum, ByRef vNextItem As Variant, ByRef bNoMore As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the next cItemGroup in an enumeration.
    '---------------------------------------------------------------------------------------
    If tEnum.iControl <> miItemGroupControl Then gErr vbccCollectionChangedDuringEnum, cItemGroups
    
    tEnum.iIndex = tEnum.iIndex + OneL
    If tEnum.iIndex > NegOneL And tEnum.iIndex < miItemGroupCount Then
        Dim loGroup      As cItemGroup
        Set loGroup = New cItemGroup
        
        loGroup.fInit Me, mtItemGroups(tEnum.iIndex).iId, tEnum.iIndex
        Set vNextItem = loGroup
    Else
        bNoMore = True
    End If
End Sub

Private Function pItemGroups_GetIndex(ByRef vGroup As Variant) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the index of a group given its key or index.
    '---------------------------------------------------------------------------------------
    If VarType(vGroup) = vbString Then
        pItemGroups_GetIndex = pItemGroups_FindKey(vGroup)
    ElseIf VarType(vGroup) = vbObject Then
        Dim loGroup      As cItemGroup
        On Error Resume Next
        Set loGroup = vGroup
        If loGroup.fIsOwner(Me) Then
            pItemGroups_GetIndex = loGroup.Index
        End If
        pItemGroups_GetIndex = pItemGroups_GetIndex - OneL
        On Error GoTo 0
        
    Else
        On Error Resume Next
        pItemGroups_GetIndex = CLng(vGroup)
        pItemGroups_GetIndex = pItemGroups_GetIndex - OneL
        On Error GoTo 0
        If pItemGroups_GetIndex < ZeroL Or pItemGroups_GetIndex >= miItemGroupCount Then pItemGroups_GetIndex = NegOneL
        
    End If
End Function

Private Function pItemGroups_FindKey(ByRef vGroup As Variant) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Find a key in the itemgroups collection.
    '---------------------------------------------------------------------------------------
    Dim sKey      As String
    sKey = CStr(vGroup)
    If LenB(sKey) Then
        For pItemGroups_FindKey = ZeroL To miItemGroupCount - OneL
            If StrComp(mtItemGroups(pItemGroups_FindKey).sKey, sKey) = ZeroL Then Exit Function
        Next
    End If
    pItemGroups_FindKey = NegOneL
End Function


Friend Property Get fItem_ToolTipText(ByVal lpItem As Long, ByRef iIndex As Long) As String
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the tooltiptext of an item.
    '---------------------------------------------------------------------------------------
    If pItem_Verify(lpItem, iIndex) Then
        pcListItem_GetToolTip lpItem, fItem_ToolTipText
    End If
End Property
Friend Property Let fItem_ToolTipText(ByVal lpItem As Long, ByRef iIndex As Long, ByRef sNew As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the tooltiptext of an item.
    '---------------------------------------------------------------------------------------
    If pItem_Verify(lpItem, iIndex) Then
        If pcListItem_SetToolTip(lpItem, sNew) = ZeroL Then gErr vbccOutOfMemory, cListItem
    End If
End Property

Friend Property Get fItem_ItemData(ByVal lpItem As Long, ByRef iIndex As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the ItemData of an item.
    '---------------------------------------------------------------------------------------
    If pItem_Verify(lpItem, iIndex) Then
        fItem_ItemData = pcListItem_iItemData(lpItem)
    End If
End Property
Friend Property Let fItem_ItemData(ByVal lpItem As Long, ByRef iIndex As Long, ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the ItemData of an item.
    '---------------------------------------------------------------------------------------
    If pItem_Verify(lpItem, iIndex) Then
        If pcListItem_SetItemData(lpItem, iNew) = NegOneL Then gErr vbccKeyAlreadyExists, cListItem
    End If
End Property

Friend Property Get fItem_Text(ByVal lpItem As Long, ByRef iIndex As Long) As String
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the text of an item.
    '---------------------------------------------------------------------------------------
    If pItem_Verify(lpItem, iIndex) Then
        pItem_GetText iIndex, ZeroL, fItem_Text
    End If
End Property
Friend Property Let fItem_Text(ByVal lpItem As Long, ByRef iIndex As Long, ByRef sNew As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the text of an item.
    '---------------------------------------------------------------------------------------
    If pItem_Verify(lpItem, iIndex) Then
        pItem_PutText iIndex, ZeroL, sNew
    End If
End Property

Friend Property Get fItem_Key(ByVal lpItem As Long, ByRef iIndex As Long) As String
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the key of an item.
    '---------------------------------------------------------------------------------------
    If pItem_Verify(lpItem, iIndex) Then
        pcListItem_GetKey lpItem, fItem_Key
    End If
End Property

Friend Property Let fItem_Key(ByVal lpItem As Long, ByRef iIndex As Long, ByRef sNew As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the key of an item.
    '---------------------------------------------------------------------------------------
    If pItem_Verify(lpItem, iIndex) Then
        Dim liResult      As Long
        liResult = pcListItem_SetKey(lpItem, sNew)
        If liResult <> OneL Then
            If liResult = ZeroL Then gErr vbccOutOfMemory, cListItem
            If liResult = NegOneL Then gErr vbccKeyAlreadyExists, cListItem
            ''debug.assert False
        End If
    End If
End Property

Friend Property Get fItem_Index(ByVal lpItem As Long, ByRef iIndex As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the index of an item in the collection.
    '---------------------------------------------------------------------------------------
    If pItem_Verify(lpItem, iIndex) Then
        fItem_Index = iIndex + OneL
    End If
End Property

Friend Property Get fItem_IconIndex(ByVal lpItem As Long, ByRef iIndex As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the iconindex of an item.
    '---------------------------------------------------------------------------------------
    If pItem_Verify(lpItem, iIndex) Then
        fItem_IconIndex = pItem_Info(iIndex, ZeroL, LVIF_IMAGE)
    End If
End Property
Friend Property Let fItem_IconIndex(ByVal lpItem As Long, ByRef iIndex As Long, ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the iconindex of an item.
    '---------------------------------------------------------------------------------------
    If pItem_Verify(lpItem, iIndex) Then
        pItem_Info(iIndex, ZeroL, LVIF_IMAGE) = iNew
    End If
End Property

Friend Property Get fItem_Checked(ByVal lpItem As Long, ByRef iIndex As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return whether an item is checked.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        If pItem_Verify(lpItem, iIndex) Then
            fItem_Checked = CBool(SendMessage(mhWnd, LVM_GETITEMSTATE, iIndex, LVIS_STATEIMAGEMASK) And &H2000&)
        End If
    End If
End Property

Friend Property Let fItem_Checked(ByVal lpItem As Long, ByRef iIndex As Long, ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the checked state of an item.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        If pItem_Verify(lpItem, iIndex) Then
            With mtItem
                .iItem = iIndex
                .mask = LVIF_STATE
                .stateMask = &H3000&
                If bNew Then
                    ' check
                    .State = &H2000&
                
                Else
                    ' uncheck
                    .State = &H1000&
                
                End If
                SendMessage mhWnd, LVM_SETITEMSTATE, iIndex, VarPtr(.mask)
            End With
        End If
    End If
End Property

Friend Property Get fItem_Cut(ByVal lpItem As Long, ByRef iIndex As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return a value indicating whether the item is displayed ghosted.
    '---------------------------------------------------------------------------------------
    
    If pItem_Verify(lpItem, iIndex) Then
        fItem_Cut = pItem_State(iIndex, LVIS_CUT)
    End If
End Property
Friend Property Let fItem_Cut(ByVal lpItem As Long, ByRef iIndex As Long, ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether the item is displayed ghosted.
    '---------------------------------------------------------------------------------------
    If pItem_Verify(lpItem, iIndex) Then
        pItem_State(iIndex, LVIS_CUT) = bNew
    End If
End Property

Friend Property Get fItem_Indent(ByVal lpItem As Long, ByRef iIndex As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the indentation level of an item.
    '---------------------------------------------------------------------------------------
    If pItem_Verify(lpItem, iIndex) Then
        fItem_Indent = pItem_Info(iIndex, ZeroL, LVIF_INDENT)
    End If
End Property
Friend Property Let fItem_Indent(ByVal lpItem As Long, ByRef iIndex As Long, ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the indentation level of an item.
    '---------------------------------------------------------------------------------------
    If pItem_Verify(lpItem, iIndex) Then
        pItem_Info(iIndex, ZeroL, LVIF_INDENT) = iNew
    End If
End Property

Friend Property Get fItem_Selected(ByVal lpItem As Long, ByRef iIndex As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the selected state of an item.
    '---------------------------------------------------------------------------------------
    If pItem_Verify(lpItem, iIndex) Then
        fItem_Selected = pItem_State(iIndex, LVIS_SELECTED)
    End If
End Property

Friend Property Let fItem_Selected(ByVal lpItem As Long, ByRef iIndex As Long, ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the selected state of an item.
    '---------------------------------------------------------------------------------------
    If pItem_Verify(lpItem, iIndex) Then
        pItem_State(iIndex, LVIS_SELECTED) = bNew
    End If
End Property

Friend Property Get fItem_Top(ByVal lpItem As Long, ByRef iIndex As Long) As Single
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the y coordinate of the top of the item.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        If pItem_Verify(lpItem, iIndex) Then
            Dim tp      As POINT
            SendMessage mhWnd, LVM_GETITEMPOSITION, iIndex, VarPtr(tp)
            fItem_Top = ScaleY(tp.Y, vbPixels, vbContainerPosition)
        End If
    End If
End Property
Friend Property Let fItem_Top(ByVal lpItem As Long, ByRef iIndex As Long, ByVal fNew As Single)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the y coordinate of the top of the item.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        If pItem_Verify(lpItem, iIndex) Then
            Dim tp      As POINT
            SendMessage mhWnd, LVM_GETITEMPOSITION, iIndex, VarPtr(tp)
            tp.Y = ScaleY(fNew, vbContainerPosition, vbPixels)
            SendMessage mhWnd, LVM_SETITEMPOSITION32, iIndex, VarPtr(tp)
        End If
    End If
End Property

Friend Property Get fItem_Left(ByVal lpItem As Long, ByRef iIndex As Long) As Single
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the x coordinate of the left of the item.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        If pItem_Verify(lpItem, iIndex) Then
            Dim tp      As POINT
            SendMessage mhWnd, LVM_GETITEMPOSITION, iIndex, VarPtr(tp)
            fItem_Left = ScaleX(tp.X, vbPixels, vbContainerPosition)
        End If
    End If
End Property
Friend Property Let fItem_Left(ByVal lpItem As Long, ByRef iIndex As Long, ByVal fNew As Single)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the x coordinate of the left of the item.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        If pItem_Verify(lpItem, iIndex) Then
            Dim tp      As POINT
            SendMessage mhWnd, LVM_GETITEMPOSITION, iIndex, VarPtr(tp)
            tp.X = ScaleX(fNew, vbContainerPosition, vbPixels)
            SendMessage mhWnd, LVM_SETITEMPOSITION32, iIndex, VarPtr(tp)
        End If
    End If
End Property

Friend Property Get fItem_SubItem(ByVal lpItem As Long, ByRef iIndex As Long, ByRef vColumn As Variant) As cListSubItem
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a cListSubItem object representing the given subitem.
    '---------------------------------------------------------------------------------------
    If pItem_Verify(lpItem, iIndex) Then
        Dim liIndex      As Long
        liIndex = pColumns_GetIndex(vColumn)
        If liIndex < ZeroL Or liIndex > (miColumnCount - OneL) Then gErr vbccKeyOrIndexNotFound, cListItem
        Set fItem_SubItem = New cListSubItem
        fItem_SubItem.fInit Me, lpItem, iIndex, liIndex, mtColumns(liIndex).iId
    End If
End Property

Friend Property Get fItem_Group(ByVal lpItem As Long, ByRef iIndex As Long) As cItemGroup
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a cItemGroup object represeting the group that the item is in.  Requires CC 6.0.
    '---------------------------------------------------------------------------------------
    If pItem_Verify(lpItem, iIndex) Then
        Dim liId      As Long
        liId = pItem_Info(iIndex, ZeroL, LVIF_GROUPID)

        If liId <> ZeroL Then
            Set fItem_Group = New cItemGroup
            fItem_Group.fInit Me, liId, pItem_GroupIndexFromId(liId)
        End If
    End If
End Property
Friend Property Set fItem_Group(ByVal lpItem As Long, ByRef iIndex As Long, ByVal cG As cItemGroup)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the group that an item is in.  Requires CC 6.0.
    '---------------------------------------------------------------------------------------
    If pItem_Verify(lpItem, iIndex) Then
        pItem_Info(iIndex, ZeroL, LVIF_GROUPID) = cG.fId
    End If
End Property

Friend Sub fItem_SetGroup(ByVal lpItem As Long, ByRef iIndex As Long, ByRef vGroup As Variant)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the group that an item is in.  Requires CC 6.0.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        If pItem_Verify(lpItem, iIndex) Then
            Dim liIndex      As Long
            liIndex = pItemGroups_GetIndex(vGroup)

            If liIndex = NegOneL Then
                SendMessage mhWnd, LVM_MOVEITEMTOGROUP, iIndex, ZeroL
            Else
                SendMessage mhWnd, LVM_MOVEITEMTOGROUP, iIndex, mtItemGroups(liIndex).iId
            End If
        End If
    End If
End Sub


Friend Sub fItem_SetTileViewItems(ByVal lpItem As Long, ByRef iIndex As Long, ByRef iItems() As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the subitems that are shown in tile view.  Requires CC 6.0.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        If pItem_Verify(lpItem, iIndex) Then
            Dim i      As Long
            With mtTile
                .cbSize = Len(mtTile)
                .cColumns = UBound(iItems) - LBound(iItems) + OneL
                .iItem = iIndex
                .puColumns = VarPtr(iItems(LBound(iItems)))
            End With
            SendMessage mhWnd, LVM_SETTILEINFO, ZeroL, VarPtr(mtTile)
        End If
    End If
End Sub

Private Function pItem_Verify(ByVal lpItem As Long, ByRef iIndex As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Determine whether an item is still a member of the collection.
    '---------------------------------------------------------------------------------------
    ''debug.assert lpItem
    
    If lpItem Then
    
        pItem_Verify = CBool(lpItem = pItem_Info(iIndex, ZeroL, LVIF_PARAM))
        
        If Not pItem_Verify Then
            If mhWnd Then
                iIndex = pItem_IndexFromlParam(lpItem)
                pItem_Verify = (iIndex > NegOneL)
            End If
        End If
        
        If Not pItem_Verify Then gErr vbccItemDetached, cListItem
    
    End If
    
End Function

Private Function pItem_GroupIndexFromId(ByVal iId As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a group index from its id.
    '---------------------------------------------------------------------------------------
    For pItem_GroupIndexFromId = ZeroL To miColumnCount - 1
        If mtItemGroups(pItem_GroupIndexFromId).iId = iId Then Exit Function
    Next
    pItem_GroupIndexFromId = NegOneL
End Function

Private Property Get pItem_Info(ByVal iIndex As Long, ByVal iSubItem As Long, ByVal iInfo As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return a 32 bit value from the LVITEM structure of an item.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        With mtItem
            .mask = iInfo
            .iItem = iIndex
            .iSubItem = iSubItem
        
            If SendMessage(mhWnd, LVM_GETITEMA, ZeroL, VarPtr(.mask)) Then
                If iInfo = LVIF_PARAM Then
                    pItem_Info = .lParam
                ElseIf iInfo = LVIF_COLUMNS Then
                    pItem_Info = .cColumns
                ElseIf iInfo = LVIF_GROUPID Then
                    pItem_Info = .iGroupId
                ElseIf iInfo = LVIF_IMAGE Then
                    pItem_Info = .iImage
                ElseIf iInfo = LVIF_INDENT Then
                    pItem_Info = .iIndent
                End If
            End If
        End With
    End If
End Property
Private Property Let pItem_Info(ByVal iIndex As Long, ByVal iSubItem As Long, ByVal iInfo As Long, ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set a 32 bit value from the LVITEM structure of an item.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        With mtItem
            .mask = iInfo
            .iItem = iIndex
            .iSubItem = iSubItem
        
            If iInfo = LVIF_COLUMNS Then
                .cColumns = iNew
            ElseIf iInfo = LVIF_GROUPID Then
                .iGroupId = iNew
            ElseIf iInfo = LVIF_IMAGE Then
                .iImage = iNew
            ElseIf iInfo = LVIF_INDENT Then
                .iIndent = iNew
            ElseIf iInfo = LVIF_PARAM Then
                .lParam = iNew
            End If
        
            SendMessage mhWnd, LVM_SETITEMA, ZeroL, VarPtr(.mask)
        
        End With
    End If
End Property

Private Property Get pItem_State(ByVal iIndex As Long, ByVal iState As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return a value indicating if an item has the given state.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        pItem_State = CBool(SendMessage(mhWnd, LVM_GETITEMSTATE, iIndex, iState) And iState)
    End If
End Property
Private Property Let pItem_State(ByVal iIndex As Long, ByVal iState As Long, ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Add or remove an item state.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        With mtItem
            .mask = LVIF_STATE
            .stateMask = iState
            '.iItem = iIndex
            .iSubItem = ZeroL
        
            If bNew _
                Then .State = iState _
            Else .State = ZeroL
        
                SendMessage mhWnd, LVM_SETITEMSTATE, iIndex, VarPtr(.mask)
            End With
        End If
End Property

Private Sub pItem_PutText(ByVal iIndex As Long, ByVal iSubItem As Long, ByRef sIn As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the text of an item.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
    
        With mtItem
            .iSubItem = iSubItem
            lstrFromStringW .pszText, .cchTextMax, sIn
        End With
    
        SendMessage mhWnd, LVM_SETITEMTEXTW, iIndex, VarPtr(mtItem.mask)
    
    End If
End Sub

Private Sub pItem_GetText(ByVal iIndex As Long, ByVal iSubItem As Long, ByRef sOut As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the text of an item.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
    
        Dim liLen      As Long
    
        mtItem.iSubItem = iSubItem
    
        liLen = SendMessage(mhWnd, LVM_GETITEMTEXTW, iIndex, VarPtr(mtItem.mask))
    
        If liLen > ZeroL Then
            lstrToStringW mtItem.pszText, sOut
        
        Else
            sOut = ""
        
        End If

    End If
End Sub

Private Function pItem_IndexFromlParam(ByVal lParam As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return an item index from the item lparam (memory handle to private storage).
    '---------------------------------------------------------------------------------------
    With mtFind
        .Flags = LVFI_PARAM
        .lParam = lParam
        pItem_IndexFromlParam = SendMessage(mhWnd, LVM_FINDITEMA, NegOneL, VarPtr(.Flags))
    End With
End Function




Friend Function fItems_Add( _
ByRef sKey As String, _
ByRef sText As String, _
ByRef sToolTipText As String, _
ByVal iIconIndex As Long, _
ByVal iItemData As Long, _
ByVal iIndent As Long, _
ByRef vItemBefore As Variant, _
ByRef vItemGroup As Variant, _
ByRef vSubItems As Variant) _
As cListItem
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Add an item to the collection.
    '---------------------------------------------------------------------------------------
    ' On Error Resume Next
            
    If mhWnd Then
    
        Dim liInsertBefore      As Long
        liInsertBefore = pItems_GetIndex(vItemBefore)
        If liInsertBefore = NegOneL Then
            If Not IsMissing(vItemBefore) Then gErr vbccKeyOrIndexNotFound, cListItems
            liInsertBefore = miItemCount
        End If
    
        Dim liGroupId      As Long
        liGroupId = pItemGroups_GetIndex(vItemGroup)
    
        If liGroupId = NegOneL Then
            If Not IsMissing(vItemGroup) Then gErr vbccKeyOrIndexNotFound, cListItems
            liGroupId = ZeroL
        Else
            liGroupId = mtItemGroups(liGroupId).iId
        End If
    
        Dim lpItem      As Long
        lpItem = pcListItem_Alloc(sKey, sToolTipText, iItemData)
    
        With mtItem
            lstrFromStringW .pszText, .cchTextMax, sText
            .iImage = iIconIndex
            .iIndent = iIndent
            .lParam = lpItem
            .iItem = liInsertBefore
            .iGroupId = liGroupId
            .iSubItem = ZeroL
            If .iItem = NegOneL Then .iItem = miItemCount
            .mask = LVIF_TEXT Or LVIF_IMAGE Or LVIF_PARAM Or LVIF_INDENT
            If liGroupId Then .mask = .mask Or LVIF_GROUPID
        End With
            
        liInsertBefore = SendMessage(mhWnd, LVM_INSERTITEMW, ZeroL, VarPtr(mtItem.mask))
    
        ''debug.assert liInsertBefore > NegOneL
    
        If liInsertBefore > NegOneL Then
            Incr miItemControl
        
            If IsArray(vSubItems) Then
                Dim i             As Long
                Dim liLBound      As Long
                liLBound = LBound(vSubItems)
                For i = liLBound To UBound(vSubItems)
                    If Not IsMissing(vSubItems(i)) Then pItem_PutText liInsertBefore, i - liLBound + OneL, Format$(vSubItems(i), mtColumns(i - liLBound + OneL).sFormat)
                Next
            End If
        
            Set fItems_Add = New cListItem
            fItems_Add.fInit Me, lpItem, liInsertBefore
        
            miItemCount = miItemCount + OneL
        Else
        
            ''debug.assert False
            pcListItem_Free lpItem
        
        End If
    
    End If
End Function

Friend Function fItems_Clear() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Clear all items from the collection.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        'additional work is done in the LVN_DELETEALLITEMS notification
        fItems_Clear = SendMessage(mhWnd, LVM_DELETEALLITEMS, ZeroL, ZeroL)
        pUpdateSortArrow NegOneL, ZeroL
    End If
End Function

Friend Function fItems_Remove(ByRef vItem As Variant) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Remove an item from the collection.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        
        Dim liIndex      As Long
        
        liIndex = pItems_GetIndex(vItem)
        
        If liIndex = NegOneL Then gErr vbccKeyOrIndexNotFound, cListItems
        
        If liIndex = miFocusIndex Then miFocusIndex = NegOneL
        
        'additional work is done in the LVN_DELETEITEM notification
        SendMessage mhWnd, LVM_DELETEITEM, liIndex, ZeroL
            
    End If
End Function

Friend Property Get fItems_Count() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the number of items in the collection.
    '---------------------------------------------------------------------------------------
    fItems_Count = miItemCount
    ''debug.assert fItems_Count = SendMessage(mhWnd, LVM_GETITEMCOUNT, 0, 0)
End Property

Friend Property Get fItems_Exists(ByRef vItem As Variant) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return a value indicating whether a given item exists in the collection.
    '---------------------------------------------------------------------------------------
    fItems_Exists = pItems_GetIndex(vItem) <> NegOneL
End Property

Friend Property Get fItems_Item(ByRef vItem As Variant) As cListItem
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return a cListItem object representing a given item.
    '---------------------------------------------------------------------------------------
    Dim liIndex      As Long
    liIndex = pItems_GetIndex(vItem)
    If liIndex = NegOneL Then gErr vbccKeyOrIndexNotFound, cListItems
    
    Dim liPtr      As Long
    
    liPtr = pItem_Info(liIndex, ZeroL, LVIF_PARAM)
    If liPtr Then
        Set fItems_Item = New cListItem
        fItems_Item.fInit Me, liPtr, liIndex
    Else
        'debug.assert False
    End If
    
    
End Property

Friend Sub fItems_InitStorage(ByVal iAdditionalItems As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Tell the listview to allocate memory to store this number of items.
    '---------------------------------------------------------------------------------------
    Dim liNewUbound      As Long
    liNewUbound = RoundToInterval(miItemCount + iAdditionalItems, 128&)
    
    If mhWnd Then
        SendMessage mhWnd, LVM_SETITEMCOUNT, miItemCount + iAdditionalItems, ZeroL
    End If
    
End Sub

Friend Function fItems_Control() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return an identity value for the current item collection.
    '---------------------------------------------------------------------------------------
    fItems_Control = miItemControl
End Function

Friend Sub fItems_NextItem(ByRef tEnum As tEnum, ByRef vNextItem As Variant, ByRef bNoMore As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the next cListItem in an enumeration.
    '---------------------------------------------------------------------------------------
    If tEnum.iControl <> miItemControl Then gErr vbccCollectionChangedDuringEnum, cListItems
    
    tEnum.iIndex = tEnum.iIndex + OneL
    bNoMore = tEnum.iIndex < ZeroL Or tEnum.iIndex >= miItemCount
    If bNoMore = False Then Set vNextItem = pItem(tEnum.iIndex)
End Sub



Private Function pItems_GetIndex(ByRef vItem As Variant) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the item index from its key or index.
    '---------------------------------------------------------------------------------------
    If VarType(vItem) = vbString Then
        Dim sKey      As String
        sKey = StrConv(vItem & vbNullChar, vbFromUnicode)
        If LenB(sKey) > OneL _
            Then pItems_GetIndex = pItems_FindKey(StrPtr(sKey)) _
        Else pItems_GetIndex = NegOneL
        ElseIf VarType(vItem) = vbObject Then
            On Error Resume Next
            Dim loItem      As cListItem
            Set loItem = vItem
            If loItem.fIsOwner(Me) Then pItems_GetIndex = loItem.Index
            pItems_GetIndex = pItems_GetIndex - OneL
            On Error GoTo 0
        Else
            On Error Resume Next
            If Not IsMissing(vItem) Then
                pItems_GetIndex = CLng(vItem)
            End If
            If pItems_GetIndex < ZeroL Or pItems_GetIndex > miItemCount _
                Then pItems_GetIndex = NegOneL _
            Else pItems_GetIndex = pItems_GetIndex - OneL
                On Error GoTo 0
            End If
End Function


Private Function pItems_FindKey(ByVal lpsz As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Find an item by its key.
    '---------------------------------------------------------------------------------------
    pItems_FindKey = NegOneL
    
    If mhWnd And Not moKeyMap Is Nothing Then
        With mtFind
            .Flags = LVFI_PARAM
            .lParam = moKeyMap.Find(lpsz, Hash(lpsz, lstrlen(lpsz)))
            If .lParam Then pItems_FindKey = SendMessage(mhWnd, LVM_FINDITEMA, NegOneL, VarPtr(.Flags))
        End With
    End If
End Function

Friend Property Get fSubItem_Text(ByVal lpItem As Long, ByRef iIndex As Long, ByRef iColumnIndex As Long, ByVal iColumnId As Long) As String
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the text of a subitem.
    '---------------------------------------------------------------------------------------
    If pSubItem_Verify(lpItem, iIndex, iColumnIndex, iColumnId) Then
        pItem_GetText iIndex, iColumnIndex, fSubItem_Text
    End If
End Property

Friend Property Let fSubItem_Text(ByVal lpItem As Long, ByRef iIndex As Long, ByRef iColumnIndex As Long, ByVal iColumnId As Long, ByRef sNew As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the text of a subitem.
    '---------------------------------------------------------------------------------------
    If pSubItem_Verify(lpItem, iIndex, iColumnIndex, iColumnId) Then
        pItem_PutText iIndex, iColumnIndex, sNew
    End If
End Property

Friend Property Get fSubItem_IconIndex(ByVal lpItem As Long, ByRef iIndex As Long, ByRef iColumnIndex As Long, ByVal iColumnId As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the iconindex of a subitem.
    '---------------------------------------------------------------------------------------
    If miStyleEx And LVS_EX_SUBITEMIMAGES Then
        If pSubItem_Verify(lpItem, iIndex, iColumnIndex, iColumnId) _
            Then fSubItem_IconIndex = pItem_Info(iIndex, iColumnIndex, LVIF_IMAGE)
        Else
            fSubItem_IconIndex = NegOneL
        End If
End Property
Friend Property Let fSubItem_IconIndex(ByVal lpItem As Long, ByRef iIndex As Long, ByRef iColumnIndex As Long, ByVal iColumnId As Long, ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the iconindex of a subitem.
    '---------------------------------------------------------------------------------------
    If pSubItem_Verify(lpItem, iIndex, iColumnIndex, iColumnId) Then
        pItem_Info(iIndex, iColumnIndex, LVIF_IMAGE) = iNew
    End If
End Property

Friend Sub fSubItem_SetFormattedText(ByVal lpItem As Long, ByRef iIndex As Long, ByRef iColumnIndex As Long, ByVal iColumnId As Long, ByRef vData As Variant)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the text of an item formatted according to the column format.
    '---------------------------------------------------------------------------------------
    If pSubItem_Verify(lpItem, iIndex, iColumnIndex, iColumnId) Then
        pItem_PutText iIndex, iColumnIndex, Format$(vData, mtColumns(iColumnIndex + OneL).sFormat)
    End If
End Sub

Friend Property Get fSubItem_ShowInTileView(ByVal lpItem As Long, ByRef iIndex As Long, ByRef iColumnIndex As Long, ByVal iColumnId As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a value indicating whether the given subitem is visible in tile view.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        If pSubItem_Verify(lpItem, iIndex, iColumnIndex, iColumnId) Then
            If miColumnCount = ZeroL Then Exit Property
        
            Dim liIndex      As Long
            liIndex = iColumnIndex + OneL
        
            Dim liCols()      As Long
            Dim i             As Long
        
            If miColumnCount > ZeroL Then
                ReDim liCols(0 To miColumnCount - OneL)
                With mtTile
                    .cbSize = Len(mtTile)
                    .iItem = iIndex
                    .puColumns = VarPtr(liCols(0))
                    .cColumns = miColumnCount
                    SendMessage mhWnd, LVM_GETTILEINFO, ZeroL, VarPtr(.cbSize)
                End With
            
                For i = ZeroL To mtTile.cColumns - OneL
                    fSubItem_ShowInTileView = (liCols(i) = liIndex)
                    If fSubItem_ShowInTileView Then Exit For
                Next
            End If
        End If
    End If
End Property

Friend Property Let fSubItem_ShowInTileView(ByVal lpItem As Long, ByRef iIndex As Long, ByRef iColumnIndex As Long, ByVal iColumnId As Long, ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether the given subitem is visible in tile view.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        If pSubItem_Verify(lpItem, iIndex, iColumnIndex, iColumnId) Then
            If miColumnCount = ZeroL Then Exit Property
        
            Dim liCols()            As Long
            Dim liIndex             As Long
            Dim liSubItemIndex      As Long
            liSubItemIndex = iColumnIndex + OneL

            ReDim liCols(0 To miColumnCount - OneL)
        
            With mtTile
                .iItem = iIndex
                .puColumns = VarPtr(liCols(0))
                .cColumns = miColumnCount
                SendMessage mhWnd, LVM_GETTILEINFO, ZeroL, VarPtr(.cbSize)
            
                For liIndex = ZeroL To .cColumns - OneL
                    If liCols(liIndex) = liSubItemIndex Then Exit For
                Next
            
                If bNew Then
                    'debug.assert .cColumns < miColumnCount
                    If liIndex = .cColumns And .cColumns < miColumnCount Then
                        For liIndex = ZeroL To .cColumns - OneL
                            If liCols(liIndex) > liSubItemIndex Then Exit For
                        Next
                        If liIndex < .cColumns Then
                            CopyMemory liCols(liIndex + OneL), liCols(liIndex), (.cColumns - liIndex) * 4&
                        End If
                        liCols(liIndex) = liSubItemIndex
                        .cColumns = .cColumns + OneL
                        SendMessage mhWnd, LVM_SETTILEINFO, ZeroL, VarPtr(mtTile.cbSize)
                    End If
                Else
                    If liIndex < .cColumns Then
                        .cColumns = .cColumns - OneL
                        If .cColumns = ZeroL Then
                            .puColumns = ZeroL
                        Else
                            If liIndex < .cColumns Then
                                CopyMemory liCols(liIndex), liCols(liIndex + OneL), (.cColumns - liIndex) * 4&
                            End If
                        End If
                        SendMessage mhWnd, LVM_SETTILEINFO, ZeroL, VarPtr(mtTile.cbSize)
                    End If
                End If
            End With
        End If
    End If
End Property

Private Function pSubItem_Verify(lpItem As Long, ByRef iIndex As Long, ByRef iColumnIndex As Long, ByVal iColumnId As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Verify whether a subitem is still part of the collection.
    '---------------------------------------------------------------------------------------
    On Error GoTo handler
    pSubItem_Verify = pItem_Verify(lpItem, iIndex)
    If pSubItem_Verify Then pSubItem_Verify = pColumn_Verify(iColumnIndex, iColumnId)
    Exit Function
handler:
    gErr vbccItemDetached, cListSubItem
End Function

Private Sub pComp_GetText(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal bUnicode As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the text of the items being compared.
    '---------------------------------------------------------------------------------------
    Dim liMsg      As Long
    
    If bUnicode Then liMsg = LVM_GETITEMTEXTW Else liMsg = LVM_GETITEMTEXTA
    
    If miSortMsg <> LVM_SORTITEMSEX Then
        lParam1 = pItem_IndexFromlParam(lParam1)
        lParam2 = pItem_IndexFromlParam(lParam2)
    End If
    
    mtItem.pszText = miStrPtr
    SendMessage mhWnd, liMsg, lParam1, VarPtr(mtItem.mask)
    mtItem.pszText = miStrPtrCmp
    SendMessage mhWnd, liMsg, lParam2, VarPtr(mtItem.mask)
    
End Sub

Public Property Get BorderStyle() As evbComCtlBorderStyle
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the border style used by the control.
    '---------------------------------------------------------------------------------------
    BorderStyle = miBorderStyle
End Property
Public Property Let BorderStyle(ByVal iNew As evbComCtlBorderStyle)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the border style used by the control.
    '---------------------------------------------------------------------------------------
    miBorderStyle = iNew
    If Not Ambient.UserMode Then pPropChanged PROP_BorderStyle
    pSetBorder
End Property

Private Sub pSetBorder()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the border style in the usercontrol and listview.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        Select Case miBorderStyle
        Case vbccBorderSunken
            If Not Ambient.UserMode Then UserControl.BackColor = Me.ColorBack
            UserControl.BorderStyle = vbFixedSingle
            UserControl.Appearance = OneL
            SetWindowStyleEx mhWnd, ZeroL, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE
        Case vbccBorderSingle
            UserControl.BorderStyle = vbFixedSingle
            UserControl.Appearance = ZeroL
            SetWindowStyleEx mhWnd, ZeroL, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE
        Case vbccBorderNone
            UserControl.BorderStyle = vbBSNone
            UserControl.Appearance = ZeroL
            SetWindowStyleEx mhWnd, ZeroL, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE
        Case vbccBorderThin
            UserControl.BorderStyle = vbBSNone
            UserControl.Appearance = ZeroL
            SetWindowStyleEx mhWnd, WS_EX_STATICEDGE, WS_EX_CLIENTEDGE
        End Select
        SetWindowPos mhWnd, ZeroL, ZeroL, ZeroL, ZeroL, ZeroL, SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOOWNERZORDER Or SWP_NOMOVE
        Refresh
    End If

End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the enabled status of the control.
    '---------------------------------------------------------------------------------------
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal bState As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the enabled status of the control.
    '---------------------------------------------------------------------------------------
    If Not Ambient.UserMode Then pPropChanged PROP_Enabled
    UserControl.Enabled = bState
    If mhWnd Then
        SetWindowStyle mhWnd, WS_DISABLED * -(Not bState), WS_DISABLED
        EnableWindow mhWnd, -CLng(bState)
        InvalidateRect mhWnd, ByVal ZeroL, OneL
    End If
End Property

Public Property Get IconSpaceX() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the horizontal space for auto arrange items in small or large icon view.
    '---------------------------------------------------------------------------------------
    IconSpaceX = miIconSpaceX
End Property
Public Property Let IconSpaceX(ByVal X As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the horizontal space for auto arrange items in small or large icon view.
    '---------------------------------------------------------------------------------------
    If Not Ambient.UserMode Then pPropChanged PROP_IconSpaceX
    miIconSpaceX = X
    pSetIconSpacing
End Property
Public Property Get IconSpaceY() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the vertical space for auto arrange items in small or large icon view.
    '---------------------------------------------------------------------------------------
    IconSpaceY = miIconSpaceY
End Property
Public Property Let IconSpaceY(ByVal Y As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the vertical space for auto arrange items in small or large icon view.
    '---------------------------------------------------------------------------------------
    If Not Ambient.UserMode Then pPropChanged PROP_IconSpaceY
    miIconSpaceY = Y
    pSetIconSpacing
End Property
Private Sub pSetIconSpacing()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Inform the listview of our icon spacing settings.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim lXY       As Long
        Dim lXYD      As Long
        Dim lXYC      As Long
        Dim cx        As Long, cy As Long
    
        ' Set cx=-1, cy=-1 to reset to default and return current settings:
        lXYD = &HFFFFFFFF
        lXYC = SendMessage(mhWnd, LVM_SETICONSPACING, 0, lXYD)
        ' cX is loword:
        cx = (lXYC And &HFFFF&)
        ' cY is hiword:
        cy = (lXYC \ &H10000)
        If miIconSpaceX > 0 Then cx = miIconSpaceX
        If miIconSpaceY > 0 Then cy = miIconSpaceY
    
        lXY = cx And &H7FFF
        lXY = lXY Or ((cy And &H7FFF) * &H10000)
        SendMessage mhWnd, LVM_SETICONSPACING, 0, lXY
    End If
End Sub

Public Property Get ColorBack() As OLE_COLOR
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the backcolor of the listview.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        ColorBack = SendMessage(mhWnd, LVM_GETBKCOLOR, ZeroL, ZeroL)
    End If
End Property
Public Property Let ColorBack(ByVal iColor As OLE_COLOR)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Change the backcolor of the listview.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        If Not Ambient.UserMode Then pPropChanged PROP_BackColor
        If (iColor = NegOneL) Then
            SendMessage mhWnd, LVM_SETBKCOLOR, 0, NegOneL
            SendMessage mhWnd, LVM_SETTEXTBKCOLOR, 0, NegOneL
        Else
            SendMessage mhWnd, LVM_SETBKCOLOR, 0, TranslateColor(iColor)
            SendMessage mhWnd, LVM_SETTEXTBKCOLOR, 0, TranslateColor(iColor)
        End If
        UserControl.Refresh
        Me.Refresh
    End If
End Property

Public Property Get ColorFore() As OLE_COLOR
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the text color of the listview.
    '---------------------------------------------------------------------------------------
    If mhWnd Then ColorFore = SendMessage(mhWnd, LVM_GETTEXTCOLOR, ZeroL, ZeroL)
End Property
Public Property Let ColorFore(ByVal iColor As OLE_COLOR)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the text color of the listview.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        If Not Ambient.UserMode Then pPropChanged PROP_ForeColor
        If iColor = NegOneL Then
            SendMessage mhWnd, LVM_SETTEXTCOLOR, 0, NegOneL
        Else
            SendMessage mhWnd, LVM_SETTEXTCOLOR, 0, TranslateColor(iColor)
        End If
    End If
End Property

Public Property Get View() As eListViewStyle
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the current view style.
    '---------------------------------------------------------------------------------------
    View = miViewStyle
End Property
Public Property Let View(ByVal iNew As eListViewStyle)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the current view style.
    '---------------------------------------------------------------------------------------
    miViewStyle = iNew
    pSetView
    If Not Ambient.UserMode Then pPropChanged PROP_View
End Property

Private Sub pSetView()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Inform the listview of our view setting.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        If miViewStyle = lvwSmallIcon Then
            miViewStyle = lvwIcon
            pSetView
            miViewStyle = lvwSmallIcon
        End If
        If CheckCCVersion(6&) Then
            SendMessage mhWnd, LVM_SETVIEW, miViewStyle, ZeroL
        Else
            If miViewStyle = lvwTile Then miViewStyle = lvwIcon
            SetWindowStyle mhWnd, miViewStyle, (LVS_ICON Or LVS_SMALLICON Or LVS_REPORT Or LVS_LIST)
        End If
    End If
End Sub

Public Property Get DoubleBuffer() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a value indicating whether a double buffer is used for drawing operations.
    '---------------------------------------------------------------------------------------
    DoubleBuffer = CBool(miStyleEx And LVS_EX_DOUBLEBUFFER)
End Property
Public Property Let DoubleBuffer(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether a double buffer is used for drawing operations.
    '---------------------------------------------------------------------------------------
    pSetExtendedStyle LVS_EX_DOUBLEBUFFER, bNew
End Property

Public Property Get SubItemImages() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a value indicating whether subitem images are enabled.
    '---------------------------------------------------------------------------------------
    SubItemImages = CBool(miStyleEx And LVS_EX_SUBITEMIMAGES)
End Property
Public Property Let SubItemImages(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether subitem images are enabled.
    '---------------------------------------------------------------------------------------
    pSetExtendedStyle LVS_EX_SUBITEMIMAGES, bNew
End Property

Public Property Get FlatScrollBar() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a value indicating whether a flat scrollbar is used instead of the normal one.
    '---------------------------------------------------------------------------------------
    FlatScrollBar = CBool(miStyleEx And LVS_EX_FLATSB)
End Property
Public Property Let FlatScrollBar(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether a flat scrollbar is used instead of the normal one.
    '---------------------------------------------------------------------------------------
    pSetExtendedStyle LVS_EX_FLATSB, bNew
End Property

Public Property Get GridLines() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get whether gridlines are shown.
    '---------------------------------------------------------------------------------------
    GridLines = CBool(miStyleEx And LVS_EX_GRIDLINES)
End Property
Public Property Let GridLines(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether gridlines are shown.
    '---------------------------------------------------------------------------------------
    pSetExtendedStyle LVS_EX_GRIDLINES, bNew
End Property

Public Property Get BorderSelect() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get whether a border is painted around the selected item.
    '---------------------------------------------------------------------------------------
    BorderSelect = CBool(miStyleEx And LVS_EX_BORDERSELECT)
End Property
Public Property Let BorderSelect(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether a border is painted around the selected item.  This only has the
    '             desired effect when the list is in Tile view.
    '---------------------------------------------------------------------------------------
    If Not CheckCCVersion(6&) And bNew Then gErr vbccUnsupported, ucListView
    pSetExtendedStyle LVS_EX_BORDERSELECT, bNew
End Property

Public Property Get InfoTips() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a value indicating whether infotips are enabled.
    '---------------------------------------------------------------------------------------
    InfoTips = CBool(miStyleEx And LVS_EX_INFOTIP)
End Property
Public Property Let InfoTips(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether infotips are enabled.
    '---------------------------------------------------------------------------------------
    pSetExtendedStyle LVS_EX_INFOTIP, bNew
End Property

Public Property Get LabelTips() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a value indicating whether labeltips are enabled.
    '---------------------------------------------------------------------------------------
    LabelTips = CBool(miStyleEx And LVS_EX_LABELTIP)
End Property
Public Property Let LabelTips(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether a tooltip is displayed with the item text
    '             when the item is clipped and the mouse hovers over it.
    '---------------------------------------------------------------------------------------
    pSetExtendedStyle LVS_EX_LABELTIP, bNew
End Property

Public Property Get CheckBoxes() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a value indicating whether checkboxes appear next to the item icon.
    '---------------------------------------------------------------------------------------
    CheckBoxes = CBool(miStyleEx And LVS_EX_CHECKBOXES)
End Property
Public Property Let CheckBoxes(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set a value indicating whether checkboxes appear next to the item icon.
    '---------------------------------------------------------------------------------------
    pSetExtendedStyle LVS_EX_CHECKBOXES, bNew
End Property

Public Property Get TrackSelect() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a value indicating whether the items are selected as the mouse hovers over them.
    '---------------------------------------------------------------------------------------
    TrackSelect = CBool(miStyleEx And LVS_EX_TRACKSELECT)
End Property
Public Property Let TrackSelect(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether the items are selected as the mouse hovers over them.
    '---------------------------------------------------------------------------------------
    pSetExtendedStyle LVS_EX_TRACKSELECT, bNew
End Property

Public Property Get HeaderDragDrop() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a value indicating whether dragging the headers is enabled.
    '---------------------------------------------------------------------------------------
    HeaderDragDrop = CBool(miStyleEx And LVS_EX_HEADERDRAGDROP)
End Property
Public Property Let HeaderDragDrop(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether dragging the headers is enabled.
    '---------------------------------------------------------------------------------------
    pSetExtendedStyle LVS_EX_HEADERDRAGDROP, bNew
End Property

Public Property Get FullRowSelect() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a value indicating whether the full row is displayed as selected in report view.
    '---------------------------------------------------------------------------------------
    FullRowSelect = CBool(miStyleEx And LVS_EX_FULLROWSELECT)
End Property
Public Property Let FullRowSelect(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set a value indicating whether the full row is displayed as selected in report view.
    '---------------------------------------------------------------------------------------
    pSetExtendedStyle LVS_EX_FULLROWSELECT, bNew
    Refresh
End Property

Public Property Get OneClickActivate() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a value indicating whether the activate event is raised on a single click.
    '---------------------------------------------------------------------------------------
    OneClickActivate = CBool(miStyleEx And LVS_EX_ONECLICKACTIVATE)
End Property

Public Property Let OneClickActivate(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether the activate event is raised on a single click.
    '---------------------------------------------------------------------------------------
    pSetExtendedStyle LVS_EX_ONECLICKACTIVATE, bNew
End Property

Public Property Get UnderlineHot() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a value indicating whether the hot item is underlined, like a link in IE.
    '---------------------------------------------------------------------------------------
    UnderlineHot = CBool(miStyleEx And LVS_EX_UNDERLINEHOT)
End Property

Public Property Let UnderlineHot(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether the hot item is underlined, like a link in IE.
    '---------------------------------------------------------------------------------------
    pSetExtendedStyle LVS_EX_UNDERLINEHOT, bNew
End Property

Public Property Get HideColumnHeaders() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a value indicating whether the column headers are hidden in report view.
    '---------------------------------------------------------------------------------------
    HideColumnHeaders = CBool(miStyle And LVS_NOCOLUMNHEADER)
End Property
Public Property Let HideColumnHeaders(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether the column headers are hidden in report view.
    '---------------------------------------------------------------------------------------
    pSetStyle LVS_NOCOLUMNHEADER, bNew
End Property

Public Property Get AutoArrange() As eListViewAutoArrange
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the autoarrange state.
    '---------------------------------------------------------------------------------------
    AutoArrange = lvwArrangeNone
    If CBool(miStyle And LVS_AUTOARRANGE) Then
        If CBool(miStyle And LVS_ALIGNLEFT) Then
            AutoArrange = lvwArrangeLeft
        Else
            AutoArrange = lvwArrangeTop
        End If
    End If
End Property
Public Property Let AutoArrange(ByVal iNew As eListViewAutoArrange)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the autoarrange state.
    '---------------------------------------------------------------------------------------
    pSetStyle LVS_ALIGNTOP Or LVS_AUTOARRANGE Or LVS_ALIGNLEFT, False
    If iNew = lvwArrangeLeft Then
        pSetStyle LVS_AUTOARRANGE Or LVS_ALIGNLEFT, True
    ElseIf iNew = lvwArrangeTop Then
        pSetStyle LVS_AUTOARRANGE Or LVS_ALIGNTOP, True
    End If
End Property

Public Property Get AutoSort() As eListViewSortOrder
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a value indicating how items are sorted as they are added.
    '---------------------------------------------------------------------------------------
    If CBool(miStyle And LVS_SORTASCENDING) Then
        AutoSort = lvwSortAscending
    ElseIf CBool(miStyle And LVS_SORTDESCENDING) Then
        AutoSort = lvwSortDescending
    Else
        AutoSort = lvwSortNone
    End If
End Property

Public Property Let AutoSort(ByVal iNew As eListViewSortOrder)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set how items are sorted as they are added.
    '---------------------------------------------------------------------------------------
    pSetStyle LVS_SORTASCENDING Or LVS_SORTDESCENDING, False
    If iNew = lvwSortAscending Then
        pSetStyle LVS_SORTASCENDING, True
    ElseIf iNew = lvwSortDescending Then
        pSetStyle LVS_SORTDESCENDING, True
    End If
End Property

Public Property Get LabelEdit() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a value indicating whether label editing is started automatically.
    '---------------------------------------------------------------------------------------
    LabelEdit = CBool(miStyle And LVS_EDITLABELS)
End Property
Public Property Let LabelEdit(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether label editing is started automatically.
    '---------------------------------------------------------------------------------------
    pSetStyle LVS_EDITLABELS, bNew
End Property

Public Property Get LabelWrap() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a value indicating whether item text can be wrapped in icon view.
    '---------------------------------------------------------------------------------------
    LabelWrap = Not CBool(miStyle And LVS_NOLABELWRAP)
End Property
Public Property Let LabelWrap(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether item text can be wrapped in icon view.
    '---------------------------------------------------------------------------------------
    pSetStyle LVS_NOLABELWRAP, Not bNew
End Property

Public Property Get MultiSelect() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a value indicating whether multiple items can be selected.
    '---------------------------------------------------------------------------------------
    MultiSelect = Not CBool(miStyle And LVS_SINGLESEL)
End Property
Public Property Let MultiSelect(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether multiple items can be selected.
    '---------------------------------------------------------------------------------------
    pSetStyle LVS_SINGLESEL, Not bNew
End Property
Public Property Get HideSelection() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a value indicating whether the selection is hidden when the control is not in focus.
    '---------------------------------------------------------------------------------------
    HideSelection = Not CBool(miStyle And LVS_SHOWSELALWAYS)
End Property
Public Property Let HideSelection(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether the selection is hidden when the control is not in focus.
    '---------------------------------------------------------------------------------------
    pSetStyle LVS_SHOWSELALWAYS, Not bNew
    If mhWnd Then
        InvalidateRect mhWnd, ByVal ZeroL, OneL
        SendMessage mhWnd, LVM_UPDATE, ZeroL, ZeroL
    End If
End Property

Public Property Get HeaderButtons() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a value indicating whether the column headers can be clicked.
    '---------------------------------------------------------------------------------------
    HeaderButtons = CBool(miHeaderStyle And HDS_BUTTONS)
End Property
Public Property Let HeaderButtons(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether the column headers can be clicked.
    '---------------------------------------------------------------------------------------
    pSetHeaderStyle HDS_BUTTONS, bNew
End Property

Public Property Get HeaderHotTrack() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a value indicating whether the header text changes color when the mouse is over it.
    '---------------------------------------------------------------------------------------
    HeaderHotTrack = CBool(miHeaderStyle And HDS_HOTTRACK)
End Property
Public Property Let HeaderHotTrack(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether the header text changes color when the mouse is over it.
    '---------------------------------------------------------------------------------------
    pSetHeaderStyle HDS_HOTTRACK, bNew
End Property

Public Property Get HeaderTrackSize() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a value indicating whether the columns resize while the mouse moves or only
    '             when the button is released.
    '---------------------------------------------------------------------------------------
    HeaderTrackSize = CBool(miHeaderStyle And HDS_FULLDRAG)
End Property
Public Property Let HeaderTrackSize(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether the columns resize while the mouse moves or only
    '             when the button is released.
    '---------------------------------------------------------------------------------------
    pSetHeaderStyle HDS_FULLDRAG, bNew
End Property

Private Sub pSetStyle(ByVal iStyle As Long, ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set or remove the given mask in the window style.
    '---------------------------------------------------------------------------------------
    If Not Ambient.UserMode Then pPropChanged PROP_Style
    If bNew Then miStyle = miStyle Or iStyle Else miStyle = miStyle And Not iStyle
    If mhWnd Then SetWindowStyle mhWnd, iStyle * -bNew, iStyle
End Sub

Private Sub pSetExtendedStyle(ByVal iStyle As Long, ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set or remove the given mask in the extended window style.
    '---------------------------------------------------------------------------------------
    If Not Ambient.UserMode Then pPropChanged PROP_StyleEx
    If bNew Then miStyleEx = miStyleEx Or iStyle Else miStyleEx = miStyleEx And Not iStyle
    If mhWnd Then
        If bNew Then
            SendMessage mhWnd, LVM_SETEXTENDEDSTYLE, ZeroL, SendMessage(mhWnd, LVM_GETEXTENDEDSTYLE, 0, 0) Or iStyle
        Else
            SendMessage mhWnd, LVM_SETEXTENDEDSTYLE, ZeroL, SendMessage(mhWnd, LVM_GETEXTENDEDSTYLE, 0, 0) And Not iStyle
        End If
    End If
End Sub

Private Sub pSetHeaderStyle(ByVal iStyle As Long, ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set or remove the given mask in the header window style.
    '---------------------------------------------------------------------------------------
    If Not Ambient.UserMode Then pPropChanged PROP_HeaderStyle
    
    If bNew _
        Then miHeaderStyle = miHeaderStyle Or iStyle _
    Else miHeaderStyle = miHeaderStyle And Not iStyle
    
        If mhWnd Then
            ' Set the Buttons mode of the ListView's header control:
            Dim lhWnd      As Long
            lhWnd = SendMessage(mhWnd, LVM_GETHEADER, ZeroL, ZeroL)
            If Not (lhWnd = 0) Then
                Dim ls      As Long
                ls = GetWindowLong(lhWnd, GWL_STYLE)
                If bNew Then
                    ls = ls Or iStyle
                Else
                    ls = ls And Not iStyle
                End If
                SetWindowLong lhWnd, GWL_STYLE, ls
            End If
        End If
End Sub


Public Property Get ShowSortArrow() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a value indicating whether a sort arrow is shown automatically when columns
    '             are sorted.  CC 6.0 only.
    '---------------------------------------------------------------------------------------
    ShowSortArrow = mbShowSortArrow
End Property
Public Property Let ShowSortArrow(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether a sort arrow is shown automatically when columns
    '             are sorted.  CC 6.0 only.
    '---------------------------------------------------------------------------------------
    mbShowSortArrow = bNew
    pUpdateSortArrow NegOneL, ZeroL
End Property

Private Sub pUpdateSortArrow(ByVal iNewIndex As Long, ByVal iNewFormat As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Update the sort arrow according to the ShowSortArrow Property.
    '---------------------------------------------------------------------------------------
    If mbShowSortArrow Or iNewIndex = NegOneL Then
        If CheckCCVersion(6&) Then
            Static iIndex As Long
            Static iFormat As Long
            Static bInit As Boolean
            
            If Not bInit Then
                bInit = True
                iIndex = NegOneL
            End If
            
            If (iNewIndex <> iIndex) Or (iFormat <> iNewFormat) Then
                If iIndex <> NegOneL And iFormat <> ZeroL Then
                    pColumn_Format(iIndex, iFormat) = False
                End If
                
                iIndex = iNewIndex
                iFormat = iNewFormat
                
                If iIndex <> NegOneL And iFormat <> ZeroL Then
                    pColumn_Format(iIndex, iFormat) = True
                End If
                
            End If
        End If
    End If
End Sub

Public Property Get ListItems() As cListItems
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the collection of list items.
    '---------------------------------------------------------------------------------------
    Set ListItems = New cListItems
    ListItems.fInit Me
End Property

Public Property Get Columns() As cColumns
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the collection of columns.
    '---------------------------------------------------------------------------------------
    Set Columns = New cColumns
    Columns.fInit Me
End Property
Public Property Get ItemGroups() As cItemGroups
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the collection of item groups.  CC 6.0 only.
    '---------------------------------------------------------------------------------------
    If CheckCCVersion(6&) Then
        Set ItemGroups = New cItemGroups
        ItemGroups.fInit Me
    End If
End Property

Public Property Get Font() As cFont
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the font used by this control.
    '---------------------------------------------------------------------------------------
    Set Font = moFont
End Property

Public Property Set Font(ByVal oNew As cFont)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the font used by this control.
    '---------------------------------------------------------------------------------------
    If oNew Is Nothing _
        Then Set moFont = Font_CreateDefault(Ambient.Font) _
    Else Set moFont = oNew
        moFont_Changed
End Property

Public Property Get NextItem(ByVal iState As eListViewGetNextItem, Optional ByVal vStartItem As Variant) As cListItem
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the next item.  Can be filtered by selection or ghosted state, or by direction
    '             in icon view.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        Set NextItem = pItem(SendMessage(mhWnd, LVM_GETNEXTITEM, pItems_GetIndex(vStartItem), iState And &HFFFF&))
    End If
End Property

Public Property Get FindItem(ByRef sText As String, Optional ByVal vStartItem As Variant, Optional ByVal bPartial As Boolean, Optional ByVal bWrap As Boolean) As cListItem
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the first item that matches the given text.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        lstrFromStringW mtFind.psz, MAX_TEXT, sText
        mtFind.Flags = LVFI_STRING Or (LVFI_PARTIAL * -bPartial) Or (LVFI_WRAP * -bWrap)
        Set FindItem = pItem(SendMessage(mhWnd, LVM_FINDITEMW, pItems_GetIndex(vStartItem), VarPtr(mtFind.Flags)))
    End If
End Property

Public Property Get FindItemData(ByVal iItemData As Long) As cListItem
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the item that matches the given itemdata.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        If Not moItemDataMap Is Nothing Then
            With mtFind
                .Flags = LVFI_PARAM
                .lParam = moItemDataMap.Find(iItemData, HashLong(iItemData))
                If .lParam Then
                    Set FindItemData = New cListItem
                    FindItemData.fInit Me, .lParam, SendMessage(mhWnd, LVM_FINDITEMA, NegOneL, VarPtr(.Flags))
                End If
            End With
        End If
    End If
End Property

Public Property Get FocusedItem() As cListItem
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the item that is currently in focus.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        Set FocusedItem = pItem(SendMessage(mhWnd, LVM_GETNEXTITEM, NegOneL, LVNI_FOCUSED))
    End If
End Property
Public Property Set FocusedItem(ByVal oNew As cListItem)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the item that is currently in focus.
    '---------------------------------------------------------------------------------------
    SetFocusedItem oNew
End Property

Public Sub SetFocusedItem(ByVal vItem As Variant)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the item that is currently in focus.
    '---------------------------------------------------------------------------------------
    Dim liIndex      As Long
    liIndex = pItems_GetIndex(vItem)
    pItem_State(NegOneL, LVIS_SELECTED Or LVIS_FOCUSED) = False
    If liIndex <> NegOneL Then
        pItem_State(liIndex, LVIS_FOCUSED Or LVIS_SELECTED) = True
    End If
End Sub

Public Property Get DropHighlightedItem() As cListItem
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the item that is displayed as a drop highlight.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        Set DropHighlightedItem = pItem(SendMessage(mhWnd, LVM_GETNEXTITEM, NegOneL, LVNI_DROPHILITED))
    End If
End Property

Public Property Set DropHighlightedItem(ByVal oNew As cListItem)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the item that is displayed as a drop highlight.
    '---------------------------------------------------------------------------------------
    SetDropHighlightedItem oNew
End Property

Public Sub SetDropHighlightedItem(ByVal vItem As Variant)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the item that is displayed as a drop highlight.
    '---------------------------------------------------------------------------------------
    Dim liIndex      As Long
    liIndex = pItems_GetIndex(vItem)
    
    If mhWnd Then
        If liIndex <> SendMessage(mhWnd, LVM_GETNEXTITEM, NegOneL, LVNI_DROPHILITED) Then
            ImageDrag_Show False
            pItem_State(NegOneL, LVIS_DROPHILITED) = False
            If liIndex <> NegOneL Then
                pItem_State(liIndex, LVIS_DROPHILITED) = True
            End If
            UpdateWindow mhWnd
            ImageDrag_Show True
        End If
    End If
End Sub

Public Property Get HotItem() As cListItem
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the item that is displayed as though the mouse is over it.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        Set HotItem = pItem(SendMessage(mhWnd, LVM_GETHOTITEM, ZeroL, ZeroL))
    End If
End Property

Public Property Set HotItem(ByVal oNew As cListItem)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the item that is displayed as though the mouse is over it.
    '---------------------------------------------------------------------------------------
    SetHotItem oNew
End Property

Public Sub SetHotItem(ByVal vItem As Variant)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the item that is displayed as though the mouse is over it.
    '---------------------------------------------------------------------------------------
    SendMessage mhWnd, LVM_SETHOTITEM, pItems_GetIndex(vItem), ZeroL
End Sub

Public Property Get SelectionMark() As cListItem
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the item that is used as the axis for selection operations
    '             while the shift key is held.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        Set SelectionMark = pItem(SendMessage(mhWnd, LVM_GETSELECTIONMARK, ZeroL, ZeroL))
    End If
End Property

Public Property Set SelectionMark(ByVal oNew As cListItem)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the item that is used as the axis for selection keyboard operations
    '             while the shift key is held.
    '---------------------------------------------------------------------------------------
    SetSelectionMark oNew
End Property

Public Sub SetSelectionMark(ByVal vItem As Variant)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the item that is used as the axis for selection keyboard operations
    '             while the shift key is held.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        SendMessage mhWnd, LVM_SETSELECTIONMARK, ZeroL, pItems_GetIndex(vItem)
    End If
End Sub

Public Property Get SelectedColumn() As cColumn
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the column that is displayed with a darker backcolor. CC 6.0 only.
    '---------------------------------------------------------------------------------------
    If CheckCCVersion(6&) Then
        If mhWnd Then
            Set SelectedColumn = pColumn(SendMessage(mhWnd, LVM_GETSELECTEDCOLUMN, ZeroL, ZeroL))
        End If
    End If
End Property

Public Property Set SelectedColumn(ByVal oNew As cColumn)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the column that is displayed with a darker backcolor. CC 6.0 only.
    '---------------------------------------------------------------------------------------
    SetSelectedColumn oNew
End Property

Public Sub SetSelectedColumn(ByVal vColumn As Variant)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the column that is displayed with a darker backcolor. CC 6.0 only.
    '---------------------------------------------------------------------------------------
    If CheckCCVersion(6&) Then
        If mhWnd Then
            SendMessage mhWnd, LVM_SETSELECTEDCOLUMN, pColumns_GetIndex(vColumn), ZeroL
        End If
    End If
End Sub

Public Property Get SelectedCount() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the number of items that are selected.
    '---------------------------------------------------------------------------------------
    If mhWnd Then SelectedCount = SendMessage(mhWnd, LVM_GETSELECTEDCOUNT, ZeroL, ZeroL)
End Property

Public Property Get ItemsPerPage() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the approximate number of items that fit in a display page of the listview.
    '---------------------------------------------------------------------------------------
    If mhWnd Then ItemsPerPage = SendMessage(mhWnd, LVM_GETCOUNTPERPAGE, ZeroL, ZeroL)
End Property


Public Sub StartLabelEdit(ByVal vItem As Variant)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Start a labeledit operation on the given item.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim liIndex      As Long
        liIndex = pItems_GetIndex(vItem)
        If liIndex = NegOneL Then gErr vbccKeyOrIndexNotFound, ucListView
        vbComCtlTlb.SetFocus mhWnd
        SendMessage mhWnd, LVM_EDITLABELA, NegOneL, ZeroL
        liIndex = SendMessage(mhWnd, LVM_EDITLABELA, liIndex, ZeroL)
    End If
End Sub

Public Sub StopLabelEdit()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : End the label edit.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        SendMessage mhWnd, LVM_EDITLABELA, NegOneL, ZeroL
    End If
End Sub

Public Sub EnsureVisible(ByVal vItem As Variant)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Scroll the listview as necessary to display a given item.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim liIndex      As Long
        liIndex = pItems_GetIndex(vItem)
        If liIndex = NegOneL Then gErr vbccKeyOrIndexNotFound, ucListView
        SendMessage mhWnd, LVM_ENSUREVISIBLE, liIndex, ZeroL
    End If
End Sub

Public Property Get TileViewItemLines() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the number of lines that are displayed in tile view.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim tLVI      As LVTILEVIEWINFO
        tLVI.cbSize = Len(tLVI)
        tLVI.dwMask = LVTVIM_COLUMNS
        SendMessage mhWnd, LVM_GETTILEVIEWINFO, ZeroL, VarPtr(tLVI)
        ''debug.assert tLVI.cLines = miTileViewItemLines
        TileViewItemLines = miTileViewItemLines
    End If
End Property

Public Property Let TileViewItemLines(ByVal lLines As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the number of lines that are displayed in tile view.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        miTileViewItemLines = lLines
        Dim tLVI      As LVTILEVIEWINFO
        tLVI.cbSize = Len(tLVI)
        tLVI.dwMask = LVTVIM_COLUMNS
        SendMessage mhWnd, LVM_GETTILEVIEWINFO, 0, VarPtr(tLVI)
        tLVI.cLines = lLines
        SendMessage mhWnd, LVM_SETTILEVIEWINFO, 0, VarPtr(tLVI)
        pPropChanged PROP_TileLines
    End If
End Property

Public Property Get ImageList(Optional ByVal iType As eListViewImageType = lvwImageSmallIcon) As cImageList
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the imagelist used by the listview.
    '---------------------------------------------------------------------------------------
    Select Case iType
    Case lvwImageSmallIcon: Set ImageList = moImageListSmall
    Case lvwImageLargeIcon: Set ImageList = moImageListLarge
    Case lvwImageHeaderImages: Set ImageList = moImageListHeader
    End Select
End Property

Public Property Set ImageList(Optional ByVal iType As eListViewImageType = lvwImageSmallIcon, ByVal oNew As cImageList)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the imagelist used by the listview.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        On Error Resume Next
        Select Case iType
        Case lvwImageSmallIcon
            Set moImageListSmall = Nothing
            Set moImageListSmallEvent = Nothing
            Set moImageListSmall = oNew
            Set moImageListSmallEvent = oNew
        Case lvwImageLargeIcon
            Set moImageListLarge = Nothing
            Set moImageListLargeEvent = Nothing
            Set moImageListLarge = oNew
            Set moImageListLargeEvent = oNew
        Case lvwImageHeaderImages
            Set moImageListHeader = Nothing
            Set moImageListHeaderEvent = Nothing
            Set moImageListHeader = oNew
            Set moImageListHeaderEvent = oNew
        End Select
        On Error GoTo 0
        pSetImageLists
    End If
End Property

Public Property Get HitTest(ByVal X As Single, ByVal Y As Single) As cListItem
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the item at the given coordinates.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim lIdx       As Long
        Dim tLVHI      As LVHITTESTINFO
    
        tLVHI.pt.X = ScaleX(X, vbContainerPosition, vbPixels)
        tLVHI.pt.Y = ScaleY(Y, vbContainerPosition, vbPixels)
        lIdx = SendMessage(mhWnd, LVM_HITTEST, ZeroL, VarPtr(tLVHI))
    
        If CBool(tLVHI.Flags And LVHT_ONITEM) Then
            Set HitTest = pItem(tLVHI.iItem)
            'debug.assert Not (HitTest Is Nothing)
        End If
    End If
End Property

Public Sub Arrange(ByVal iArrange As eListViewArrange)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Arrange the items while in icon view.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        SendMessage mhWnd, LVM_ARRANGE, iArrange, ZeroL
    End If
End Sub

Public Sub Scroll(ByVal dx As Long, ByVal dy As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Scroll the listview.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        SendMessage mhWnd, LVM_SCROLL, dx, dy
    End If
End Sub

Public Property Get TopItem() As cListItem
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the topmost visible item.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        Set TopItem = pItem(SendMessage(mhWnd, LVM_GETTOPINDEX, ZeroL, ZeroL))
    End If
End Property

Public Property Get OriginX() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a value indicating how for a list is scrolled to the right.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim tp      As POINT
        SendMessage mhWnd, LVM_GETORIGIN, ZeroL, VarPtr(tp)
        OriginX = tp.X
    End If
End Property
Public Property Get OriginY() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a value indicating how for a list is scrolled down.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim tp      As POINT
        SendMessage mhWnd, LVM_GETORIGIN, ZeroL, VarPtr(tp)
        OriginY = tp.Y
    End If
End Property

Public Sub Refresh()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Redraw the listview.
    '---------------------------------------------------------------------------------------
    If mhWnd Then SendMessage mhWnd, LVM_UPDATE, ZeroL, ZeroL
End Sub

Public Property Get Redraw() As Boolean
Attribute Redraw.VB_MemberFlags = "400"
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return a value indicating whether the control redraws when changing the list items.
    '---------------------------------------------------------------------------------------
    Redraw = mbRedraw
End Property

Public Property Let Redraw(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether the control redraws when changing the list items.
    '---------------------------------------------------------------------------------------
    mbRedraw = bNew
    If mhWnd Then SendMessage mhWnd, WM_SETREDRAW, Abs(mbRedraw), ZeroL
End Property

Public Property Get hWnd() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the hwnd of the usercontrol.
    '---------------------------------------------------------------------------------------
    hWnd = UserControl.hWnd
End Property
Public Property Get hWndListView() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the hwnd of the listview.
    '---------------------------------------------------------------------------------------
    If mhWnd Then hWndListView = mhWnd
End Property
Public Property Get hWndEdit() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the hwnd of the edit control, if any.
    '---------------------------------------------------------------------------------------
    If mhWnd Then hWndEdit = SendMessage(mhWnd, LVM_GETEDITCONTROL, ZeroL, ZeroL)
End Property

Public Function TextWidth(ByRef sText As String) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return the textwidth of the given string using the current font.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        lstrFromStringW miStrPtr, MAX_TEXT, sText
        TextWidth = SendMessage(mhWnd, LVM_GETSTRINGWIDTHW, ZeroL, miStrPtr)
    End If
End Function

Public Property Get BackPictureURL() As String
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get the path to the background picture.
    '---------------------------------------------------------------------------------------
    BackPictureURL = msBackgroundURL
End Property
Public Property Let BackPictureURL(ByVal sNew As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set the path to the background picture.
    '---------------------------------------------------------------------------------------
    msBackgroundURL = sNew
    
    Dim lsNew      As String
    
    With mtBack
        If LenB(sNew) Then
            If Asc(sNew) = vbKeyDelete Then
                lsNew = App.Path
                If Right$(lsNew, 1) = "\" Then
                    lsNew = lsNew & sNew & vbNullChar
                Else
                    lsNew = lsNew & "\" & sNew & vbNullChar
                End If
            Else
                lsNew = sNew & vbNullChar
            End If
        Else
            lsNew = sNew & vbNullChar
        End If
        
        lsNew = StrConv(lsNew, vbFromUnicode)
        .cchImageMax = LenB(lsNew)
        If .pszImage Then MemFree .pszImage
        .pszImage = MemAllocFromString(StrPtr(lsNew), .cchImageMax)
        
        If LenB(sNew) = ZeroL Then
            .ulFlags = (.ulFlags And Not LVBKIF_SOURCE_URL)
        Else
            .ulFlags = (.ulFlags Or LVBKIF_SOURCE_URL)
        End If
    End With
    pPropChanged PROP_BackURL
    pSetBackground
End Property
Public Property Let BackPictureTile(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether the backpicture is tiled over the background of the control.
    '---------------------------------------------------------------------------------------
    If bNew Then
        mtBack.ulFlags = mtBack.ulFlags Or LVBKIF_STYLE_TILE
    Else
        mtBack.ulFlags = mtBack.ulFlags And Not LVBKIF_STYLE_TILE
    End If
    pPropChanged PROP_BackTile
    pSetBackground
End Property
Public Property Get BackPictureTile() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Get a value indicating whether the backpicture is tiled over the background.
    '---------------------------------------------------------------------------------------
    BackPictureTile = CBool(mtBack.ulFlags And LVBKIF_STYLE_TILE)
End Property

Public Property Let BackPictureXOffset(ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : If the background is not tiled, it can use an offset.  This is in a percentage
    '             if the width of the listbox except on older versions of comctl32, where it is an
    '             absolute value.
    '---------------------------------------------------------------------------------------
    mtBack.xOffsetPercent = iNew
    pSetBackground
    pPropChanged PROP_BackX
End Property
Public Property Get BackPictureXOffset() As Long
    BackPictureXOffset = mtBack.xOffsetPercent
End Property

Public Property Let BackPictureYOffset(ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : If the background is not tiled, it can use an offset.  This is in a percentage
    '             if the height of the listbox except on older versions of comctl32, where it is an
    '             absolute value.
    '---------------------------------------------------------------------------------------
    mtBack.yOffsetPercent = iNew
    pSetBackground
    pPropChanged PROP_BackY
End Property
Public Property Get BackPictureYOffset() As Long
    BackPictureYOffset = mtBack.yOffsetPercent
End Property

Private Sub pSetBackground()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Inform the listview of our backpicture settings
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        SendMessage mhWnd, LVM_SETBKIMAGEA, ZeroL, VarPtr(mtBack)
        SendMessage mhWnd, LVM_SETTEXTBKCOLOR, ZeroL, NegOneL
        Refresh
    End If
End Sub

Private Sub pSetTheme()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Enable or disable the default window theme for the listview and columnheaders.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        EnableWindowTheme mhWnd, mbThemeable
        EnableWindowTheme SendMessage(mhWnd, LVM_GETHEADER, ZeroL, ZeroL), mbThemeable
    End If
End Sub

Private Sub pSetImageLists()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Inform the listview of which imagelists it is to use.
    '---------------------------------------------------------------------------------------
    
    If Not moImageListSmall Is Nothing _
        Then SendMessage mhWnd, LVM_SETIMAGELIST, LVSIL_SMALL, moImageListSmall.hIml _
    Else SendMessage mhWnd, LVM_SETIMAGELIST, LVSIL_SMALL, ZeroL
    
        If Not moImageListLarge Is Nothing _
            Then SendMessage mhWnd, LVM_SETIMAGELIST, LVSIL_NORMAL, moImageListLarge.hIml _
        Else SendMessage mhWnd, LVM_SETIMAGELIST, LVSIL_NORMAL, ZeroL
        
            Dim lhWndHeader      As Long
            lhWndHeader = SendMessage(mhWnd, LVM_GETHEADER, ZeroL, ZeroL)
    
            If lhWndHeader And Not moImageListHeader Is Nothing Then SendMessage lhWndHeader, HDM_SETIMAGELIST, ZeroL, moImageListHeader.hIml
    
End Sub

Private Sub pCreate()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Create the listview and install the needed subclasses.
    '---------------------------------------------------------------------------------------
    On Error GoTo handler
    
    pDestroy
    
    If Not CheckCCVersion(6&) Then miStyleEx = miStyleEx And Not LVS_EX_BORDERSELECT
    
    Dim lsAnsi      As String
    lsAnsi = StrConv(WC_LISTVIEW & vbNullChar, vbFromUnicode)
    
    mhWnd = CreateWindowEx(ZeroL, StrPtr(lsAnsi), ZeroL, WS_VISIBLE Or WS_TABSTOP Or WS_CHILD Or LVS_SHAREIMAGELISTS Or miStyle Or (miViewStyle And Not LV_VIEW_TILE), ZeroL, ZeroL, ScaleWidth, ScaleHeight, UserControl.hWnd, ZeroL, App.hInstance, ByVal ZeroL)
    
    If mhWnd Then
        pSetTheme
        
        If Ambient.UserMode Then
        
            VTableSubclass_OleControl_Install Me
            VTableSubclass_IPAO_Install Me
            
            Subclass_Install Me, UserControl.hWnd, Array(WM_NOTIFY, WM_CONTEXTMENU, WM_PARENTNOTIFY), WM_SETFOCUS
            Subclass_Install Me, mhWnd, Array(WM_SETFOCUS, WM_MOUSEACTIVATE), WM_KILLFOCUS
        End If
        
        If CheckCCVersion(5&, 8&) Then
            ' ...you will get better results if you send a CCM_SETVERSION message with the wParam value
            ' set to 5 before adding any items to the control
            Dim liMajor      As Long
            GetCCVersion liMajor
            SendMessage mhWnd, CCM_SETVERSION, 5, 0
        End If
        
        SendMessage mhWnd, LVM_SETEXTENDEDSTYLE, ZeroL, miStyleEx
        
        Dim lhWndHeader      As Long
        lhWndHeader = SendMessage(mhWnd, LVM_GETHEADER, ZeroL, ZeroL)
        If lhWndHeader Then SetWindowStyle lhWndHeader, miHeaderStyle, HDS_BUTTONS Or HDS_HOTTRACK Or HDS_FULLDRAG

        pSetBorder
        pSetIconSpacing
        pSetView
        
    End If
    
    Exit Sub
handler:
    Debug.Print "Create ListView Error: " & Err.Number, Err.Description
    'debug.assert False
    Resume Next
End Sub

Private Sub pDestroy()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Destroy the listview and subclasses.
    '---------------------------------------------------------------------------------------
    On Error Resume Next
    If mhWnd Then
        VTableSubclass_OleControl_Remove
        VTableSubclass_IPAO_Remove
        Subclass_Remove Me, mhWnd
        DestroyWindow mhWnd
        mhWnd = ZeroL
        Subclass_Remove Me, UserControl.hWnd
    End If
    On Error GoTo 0
End Sub

Private Sub pPropChanged(ByRef sProp As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Inform the usercontrol that a property changed.
    '---------------------------------------------------------------------------------------
    If Not mbNoPropChange Then PropertyChanged sProp
End Sub

Friend Property Get fSupportFontPropPage() As pcSupportFontPropPage
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return a proxy object to receive notifications from the font property page.
    '---------------------------------------------------------------------------------------
    Set fSupportFontPropPage = moFontPage
End Property

Private Sub moFontPage_AddFonts(ByVal o As ppFont)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Tell the property page which font properties we support.
    '---------------------------------------------------------------------------------------
    o.ShowProps PROP_Font
End Sub

Private Sub moFontPage_GetAmbientFont(o As stdole.StdFont)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Provide the property page with the ambient font.
    '---------------------------------------------------------------------------------------
    Set o = Ambient.Font
End Sub

Private Sub moFont_Changed()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Update the font in the control.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        moFont.OnAmbientFontChanged Ambient.Font
        Dim lhFont      As Long
        lhFont = mhFont
        
        mhFont = moFont.GetHandle()
        SendMessage mhWnd, WM_SETFONT, mhFont, 1
        If lhFont Then moFont.ReleaseHandle lhFont
        
        If Not Ambient.UserMode Then pPropChanged PROP_Font
    End If
End Sub

Private Sub moImageListHeaderEvent_Changed()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Update the imagelists.
    '---------------------------------------------------------------------------------------
    pSetImageLists
End Sub

Private Sub moImageListLargeEvent_Changed()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Update the imagelists.
    '---------------------------------------------------------------------------------------
    pSetImageLists
End Sub

Private Sub moImageListSmallEvent_Changed()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Update the imagelists.
    '---------------------------------------------------------------------------------------
    pSetImageLists
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Update font if we are listening to the ambient font.
    '---------------------------------------------------------------------------------------
    If StrComp(PropertyName, "Font") = ZeroL Then moFont.OnAmbientFontChanged Ambient.Font
End Sub

Private Sub UserControl_Initialize()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Initialize the control.
    '---------------------------------------------------------------------------------------
    miFocusIndex = NegOneL
    ' For XP Visual Styles:
    LoadShellMod
    InitCC ICC_LISTVIEW_CLASSES
    
    Dim liLen      As Long
    msTextBuffer = Space$(BufferLen)
    Mid$(msTextBuffer, OneL, OneL) = vbNullChar
    miStrPtr = StrPtr(msTextBuffer)
    liLen = LenB(msTextBuffer)
    
    mtItem.pszText = miStrPtr
    mtItem.cchTextMax = liLen
    
    mtCol.pszText = miStrPtr
    mtCol.cchTextMax = liLen
    
    mtGroup.pszHeader = miStrPtr
    mtGroup.cchHeader = liLen
    mtGroup.cbSize = Len(mtGroup)
    
    mtFind.psz = miStrPtr
    
    
    mbCCVer_GE_4_71 = CheckCCVersion(4&, 71&)
    
    Set moFontPage = New pcSupportFontPropPage
End Sub

Private Sub UserControl_InitProperties()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Initialize the control properties.
    '---------------------------------------------------------------------------------------
    On Error Resume Next
    mbNoPropChange = True
    
    Set moFont = Font_CreateDefault(Ambient.Font)
    
    miViewStyle = DEF_View
    miBorderStyle = DEF_BorderStyle
    miStyle = DEF_Style
    miStyleEx = DEF_StyleEx
    miHeaderStyle = DEF_HeaderStyle
    miIconSpaceX = DEF_IconSpace
    miIconSpaceY = DEF_IconSpace
    mbGroupsEnabled = DEF_GroupsEnabled
    UserControl.Enabled = DEF_Enabled
    UserControl.OLEDropMode = -DEF_OleDrop
    mbThemeable = DEF_Themeable
    
    pCreate
    moFont_Changed
    
    BackColor = DEF_Backcolor
    ForeColor = DEF_ForeColor
    
    Me.TileViewItemLines = DEF_TileLines
    Me.BackPictureTile = DEF_BackTile
    Me.BackPictureXOffset = DEF_BackX
    Me.BackPictureYOffset = DEF_BackY
    Me.BackPictureURL = DEF_BackURL
    
    mbShowSortArrow = DEF_ShowSort
    
    mbNoPropChange = False
    mbRedraw = True
    On Error GoTo 0

End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Stop the image drag and delegate the event.
    '---------------------------------------------------------------------------------------
    ImageDrag_Stop
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Stop the image drag and delegate the event.
    '---------------------------------------------------------------------------------------
    ImageDrag_Stop
    RaiseEvent OLEDragDrop(Data, Effect, CLng(Button), CLng(Shift), ScaleX(X, ScaleMode, vbContainerPosition), ScaleY(Y, ScaleMode, vbContainerPosition))
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Update the image drag and delegate the event.
    '---------------------------------------------------------------------------------------
    ImageDrag_Move mhWnd, X, Y, State
    RaiseEvent OLEDragOver(Data, Effect, CLng(Button), CLng(Shift), ScaleX(X, ScaleMode, vbContainerPosition), ScaleY(Y, ScaleMode, vbContainerPosition), CLng(State))
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Delegate the event.
    '---------------------------------------------------------------------------------------
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Delegate the event.
    '---------------------------------------------------------------------------------------
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Delegate the event.
    '---------------------------------------------------------------------------------------
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Restore properties from a previously saved instance.
    '---------------------------------------------------------------------------------------
    On Error Resume Next
    mbNoPropChange = True
    
    With PropBag
        
        Set moFont = Font_Read(PropBag, PROP_Font, Ambient.Font)
                
        miViewStyle = .ReadProperty(PROP_View, DEF_View)
        miBorderStyle = .ReadProperty(PROP_BorderStyle, DEF_BorderStyle)
        miStyle = .ReadProperty(PROP_Style, DEF_Style)
        miStyleEx = .ReadProperty(PROP_StyleEx, DEF_StyleEx)
        
        miHeaderStyle = .ReadProperty(PROP_HeaderStyle, DEF_HeaderStyle)
        miIconSpaceX = .ReadProperty(PROP_IconSpaceX, DEF_IconSpace)
        miIconSpaceY = .ReadProperty(PROP_IconSpaceY, DEF_IconSpace)
        
        mbGroupsEnabled = DEF_GroupsEnabled
        UserControl.Enabled = .ReadProperty(PROP_Enabled, DEF_Enabled)
        UserControl.OLEDropMode = -.ReadProperty(PROP_OleDrop, DEF_OleDrop)
        
        mbThemeable = .ReadProperty(PROP_Themeable, DEF_Themeable)
        pCreate
        moFont_Changed
        
        BackColor = .ReadProperty(PROP_BackColor, DEF_Backcolor)
        ForeColor = .ReadProperty(PROP_ForeColor, DEF_ForeColor)
        
        Me.TileViewItemLines = .ReadProperty(PROP_TileLines, DEF_TileLines)
        Me.BackPictureTile = .ReadProperty(PROP_BackTile, DEF_BackTile)
        Me.BackPictureXOffset = .ReadProperty(PROP_BackX, DEF_BackX)
        Me.BackPictureYOffset = .ReadProperty(PROP_BackY, DEF_BackY)
        Me.BackPictureURL = .ReadProperty(PROP_BackURL, DEF_BackURL)
        
        mbShowSortArrow = .ReadProperty(PROP_ShowSort, DEF_ShowSort)
        
    End With
    
    mbRedraw = True
    mbNoPropChange = False
    On Error GoTo 0
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Store the property values.
    '---------------------------------------------------------------------------------------
    On Error Resume Next
    With PropBag
        
        Font_Write moFont, PropBag, PROP_Font
        
        .WriteProperty PROP_View, miViewStyle, DEF_View
        .WriteProperty PROP_BorderStyle, miBorderStyle, DEF_BorderStyle
        .WriteProperty PROP_Style, miStyle, DEF_Style
        .WriteProperty PROP_StyleEx, miStyleEx, DEF_StyleEx
        .WriteProperty PROP_HeaderStyle, miHeaderStyle, DEF_HeaderStyle
        .WriteProperty PROP_IconSpaceX, miIconSpaceX, DEF_IconSpace
        .WriteProperty PROP_IconSpaceY, miIconSpaceY, DEF_IconSpace
        .WriteProperty PROP_Enabled, UserControl.Enabled, DEF_Enabled
        .WriteProperty PROP_BackColor, BackColor, DEF_Backcolor
        .WriteProperty PROP_ForeColor, ForeColor, DEF_ForeColor
        
        .WriteProperty PROP_BackTile, Me.BackPictureTile, DEF_BackTile
        .WriteProperty PROP_BackX, Me.BackPictureXOffset, DEF_BackX
        .WriteProperty PROP_BackY, Me.BackPictureYOffset, DEF_BackY
        .WriteProperty PROP_BackURL, Me.BackPictureURL, DEF_BackURL
        
        .WriteProperty PROP_ShowSort, mbShowSortArrow, DEF_ShowSort
        .WriteProperty PROP_Themeable, mbThemeable, DEF_Themeable
        .WriteProperty PROP_OleDrop, CBool(-UserControl.OLEDropMode), DEF_OleDrop
        .WriteProperty PROP_TileLines, Me.TileViewItemLines, DEF_TileLines
    End With
    On Error GoTo 0
End Sub

Private Sub UserControl_Terminate()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Release resources.
    '---------------------------------------------------------------------------------------
    If mtBack.pszImage Then MemFree mtBack.pszImage
    pDestroy
    ReleaseShellMod
    If mhFont Then GdiMgr_DeleteFont mhFont
    Set moFontPage = Nothing
End Sub

Private Sub UserControl_Resize()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Size the listview to fit the usercontrol.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        MoveWindow mhWnd, ZeroL, ZeroL, UserControl.ScaleWidth, UserControl.ScaleHeight, OneL
    End If
End Sub


Public Property Get Themeable() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return a value indicating whether the default window theme is to be used if available.
    '---------------------------------------------------------------------------------------
    Themeable = mbThemeable
End Property

Public Property Let Themeable(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether the default window theme is to be used if available.
    '---------------------------------------------------------------------------------------
    If bNew Xor mbThemeable Then
        mbThemeable = bNew
        pPropChanged PROP_Themeable
        pSetTheme
        If mhWnd Then
            SetWindowPos mhWnd, ZeroL, ZeroL, ZeroL, ZeroL, ZeroL, SWP_FRAMECHANGED Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_NOSIZE
            RedrawWindow mhWnd, ByVal ZeroL, ZeroL, RDW_INVALIDATE
        End If
        UserControl.Refresh
    End If
End Property

Public Sub OLEDrag(Optional ByVal iImageListDrag As eListViewOleImageDrag = lvwOleImageDragSelected, Optional ByVal bHideSelection As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Initiate a drag operation through OLE.  The OLE Drag events are raised for further interaction.
    '---------------------------------------------------------------------------------------
    Dim liSelIndexes()      As Long
    
    If mhWnd Then

        If iImageListDrag = lvwOleImageDragSelected Then
            
            Dim liSelCount As Long
            liSelCount = SendMessage(mhWnd, LVM_GETSELECTEDCOUNT, ZeroL, ZeroL)
            If liSelCount Then
                
                ReDim liSelIndexes(ZeroL To liSelCount - OneL)
                
                Dim liIndex      As Long
                liIndex = SendMessage(mhWnd, LVM_GETNEXTITEM, NegOneL, LVNI_SELECTED)
                liSelCount = ZeroL
                Do While liIndex > NegOneL
                    liSelIndexes(liSelCount) = liIndex
                    liSelCount = liSelCount + OneL
                    liIndex = SendMessage(mhWnd, LVM_GETNEXTITEM, liIndex, LVNI_SELECTED)
                Loop
                
            End If
            
        ElseIf iImageListDrag = lvwOleImageDragFocused Then
            
            ReDim liSelIndexes(ZeroL To ZeroL)
            liSelIndexes(ZeroL) = SendMessage(mhWnd, LVM_GETNEXTITEM, NegOneL, LVNI_FOCUSED)
            liSelCount = IIf(liSelIndexes(ZeroL) > NegOneL, OneL, ZeroL)
            
        End If
        
        If GetActiveWindow <> RootParent(UserControl.ContainerHwnd) Then SetActiveWindow RootParent(UserControl.ContainerHwnd)
        If Not ImageDrag_Alpha Then
            SendMessage mhWnd, LVM_ENSUREVISIBLE, SendMessage(mhWnd, LVM_GETNEXTITEM, NegOneL, LVNI_FOCUSED), ZeroL
            UpdateWindow mhWnd
        End If
        
        Dim X      As Long, Y As Long
        ImageDrag_Start pDragDib(bHideSelection, liSelIndexes, liSelCount, X, Y), X, Y
        
        UserControl.OLEDrag
        
    End If
End Sub

Private Function pDragDib(ByVal bHideSelection As Boolean, ByRef iIndexes() As Long, ByVal iIndexCount As Long, ByRef iHotSpotX As Long, ByRef iHotSpotY As Long) As pcDibSection
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Return a bitmap representing the given item indexes and the relative
    '             offset to the current cursor location.
    '---------------------------------------------------------------------------------------
    
    Set pDragDib = New pcDibSection
    
    Dim liIndex      As Long
    
    Dim ltOrigin As POINT
    If SendMessage(mhWnd, LVM_GETORIGIN, ZeroL, VarPtr(ltOrigin)) = ZeroL Then
        ltOrigin.X = ZeroL
        ltOrigin.Y = ZeroL
    End If
    
    Dim ltCursor      As POINT
    ltCursor.X = miLastXMouseDown
    ltCursor.Y = miLastYMouseDown
    
    Dim ltRectBounds      As RECT
    Dim ltRectBitmap      As RECT
    Dim ltRectClient      As RECT
    
    GetClientRect mhWnd, ltRectClient
    ltRectClient.Top = pDragDib_HeaderHeight()
    
    Dim ltRectItem       As RECT
    Dim ltRectDummy      As RECT
    Dim lbInView         As Boolean
    
    '       Disregard items which are outside of the client area.
    '       Record the bounding rectangles of all items and the items
    '       which intersect the client area.
    
    For liIndex = ZeroL To iIndexCount - OneL
        lbInView = False
        
        ltRectItem.Left = LVIR_SELECTBOUNDS
        If SendMessage(mhWnd, LVM_GETITEMRECT, iIndexes(liIndex), VarPtr(ltRectItem)) Then
            lbInView = CBool(IntersectRect(ltRectDummy, ltRectItem, ltRectClient))
        End If
        
        If lbInView _
            Then UnionRect ltRectBitmap, ltRectBitmap, ltRectItem _
        Else iIndexes(liIndex) = NegOneL
        
            UnionRect ltRectBounds, ltRectBounds, ltRectItem
        Next
    
        If IsRectEmpty(ltRectBitmap) = ZeroL Then
        
            '       If the bitmap would be too large, adjust it while
            '       keeping the cursor position as near to the center as possible.
    
            Const MAXWIDTH As Long = 300
            Const MAXHEIGHT As Long = 300
            Const MAXHALFWIDTH As Long = MAXWIDTH \ TwoL
            Const MAXHALFHEIGHT As Long = MAXHEIGHT \ TwoL
        
            With ltRectBitmap
                If (.Right - .Left) > MAXWIDTH Then
                    If (ltCursor.X - .Left) > MAXHALFWIDTH Then
                        If (.Right - ltCursor.X) > MAXHALFWIDTH Then
                            .Left = ltCursor.X - MAXHALFWIDTH
                            .Right = ltCursor.X + MAXHALFWIDTH
                        Else
                            .Left = .Right - MAXWIDTH
                        End If
                    Else
                        .Right = .Left + MAXWIDTH
                    End If
                End If
            
                If (.Bottom - .Top) > MAXHEIGHT Then
                    If (ltCursor.Y - .Top) > MAXHALFHEIGHT Then
                        If (.Bottom - ltCursor.Y) > MAXHALFHEIGHT Then
                            .Top = ltCursor.Y - MAXHALFHEIGHT
                            .Bottom = ltCursor.Y + MAXHALFHEIGHT
                        Else
                            .Top = .Bottom - MAXHEIGHT
                        End If
                    Else
                        .Bottom = .Top + MAXHEIGHT
                    End If
                End If
            End With
            '       For each item, draw the icon and text.
        
            Dim lhDc      As Long
            lhDc = CreateCompatibleDC(ZeroL)
            If lhDc Then
            
                Dim liXOffset      As Long:  liXOffset = ltRectBitmap.Left
                Dim liYOffset      As Long:  liYOffset = ltRectBitmap.Top
                Dim liWidth        As Long:    liWidth = ltRectBitmap.Right - liXOffset
                Dim liHeight       As Long:   liHeight = ltRectBitmap.Bottom - liYOffset
                Dim lhFont         As Long:     lhFont = SendMessage(mhWnd, WM_GETFONT, ZeroL, ZeroL)
                Dim lhIml          As Long:      lhIml = SendMessage(mhWnd, LVM_GETIMAGELIST, Sgn(miViewStyle And LVS_TYPEMASK), ZeroL)
            
                Dim liIconWidth As Long
                Dim liIconHeight As Long
                If lhIml Then ImageList_GetIconSize lhIml, liIconWidth, liIconHeight
            
                Dim liDTFlags As Long
                liDTFlags = DT_END_ELLIPSIS Or _
                (DT_SINGLELINE Or DT_EDITCONTROL) * -CBool(miViewStyle <> LVS_ICON) Or _
                (DT_WORDBREAK Or DT_CENTER) * -CBool(miViewStyle = LVS_ICON) Or _
                DT_VCENTER * -CBool(miViewStyle = LVS_REPORT Or miViewStyle = LVS_LIST Or miViewStyle = LVS_ICON)
            
                Dim liTextXPadding      As Long: liTextXPadding = 4& - 4& * CBool(miViewStyle <> lvwList)
                Dim liTextXMargin       As Long:  liTextXMargin = -TwoL - CBool(miViewStyle = lvwTile)
                Dim liTextYMargin       As Long:  liTextYMargin = TwoL * CBool(miViewStyle <> lvwIcon)
                Dim liTextColor         As Long, liTextFillColor As Long, liImlDrawFlags As Long
            
                If bHideSelection Then
                    SetBkColor lhDc, ImageDrag_TransColor
                    liTextColor = SendMessage(mhWnd, LVM_GETTEXTCOLOR, ZeroL, ZeroL)
                    liImlDrawFlags = ILD_NORMAL Or ILD_TRANSPARENT
                    liTextFillColor = ZeroL
                Else
                    SetBkColor lhDc, GetSysColor(COLOR_HIGHLIGHT)
                    liTextColor = GetSysColor(COLOR_HIGHLIGHTTEXT)
                    liImlDrawFlags = ILD_SELECTED Or ILD_TRANSPARENT
                    liTextFillColor = COLOR_HIGHLIGHT + OneL
                End If
            
                SetTextColor lhDc, liTextColor
            
                OffsetWindowOrgEx lhDc, liXOffset, liYOffset, ltRectDummy
            
                Dim lhFontOld      As Long
                If lhFont Then lhFontOld = SelectObject(lhDc, lhFont)
            
                If pDragDib.Create(liWidth, liHeight, lhDc, IIf(ImageDrag_Alpha, 32, 24)) Then
                
                    Dim lhBmpOld As Long: lhBmpOld = SelectObject(lhDc, pDragDib.hBitmap)
                    If lhBmpOld Then
                    
                        Dim lhBrush As Long: lhBrush = GdiMgr_CreateSolidBrush(ImageDrag_TransColor)
                        If lhBrush Then
                            FillRect lhDc, ltRectBitmap, lhBrush
                            GdiMgr_DeleteBrush lhBrush
                        End If
                    
                        Dim liItem      As Long, liColIndex As Long, liLen As Long, liCols() As Long
                    
                        If miViewStyle <> lvwTile And (miViewStyle <> lvwDetails Or ((miStyleEx And LVS_EX_FULLROWSELECT) = ZeroL)) Then
                            For liIndex = ZeroL To iIndexCount - OneL
                                liItem = iIndexes(liIndex)
                                If liItem > NegOneL Then
                                    If pDragDib_GetItem(liItem, ZeroL) Then
                                        pDragDib_DrawText lhDc, liItem, ZeroL, mtItem.pszText, liDTFlags, liTextFillColor, liTextXPadding, liTextXMargin, liTextYMargin, ZeroL
                                        pDragDib_DrawIcon lhDc, liItem, ZeroL, lhIml, mtItem.iImage, liIconWidth, liIconHeight, liImlDrawFlags
                                    End If
                                End If
                            Next
                    
                        ElseIf miViewStyle = lvwDetails Then 'Full row select
                            If miColumnCount Then
                                ReDim liCols(ZeroL To miColumnCount - OneL)
                            
                                For liColIndex = ZeroL To miColumnCount - OneL
                                    Select Case pColumn_GetFormat(liColIndex) And LVCFMT_JUSTIFYMASK
                                    Case LVCFMT_LEFT:   liCols(liColIndex) = liDTFlags
                                    Case LVCFMT_RIGHT:  liCols(liColIndex) = liDTFlags Or DT_RIGHT
                                    Case LVCFMT_CENTER: liCols(liColIndex) = liDTFlags Or DT_CENTER
                                    End Select
                                Next
                            End If
                        
                            For liIndex = ZeroL To iIndexCount - OneL
                                liItem = iIndexes(liIndex)
                                If liItem > NegOneL Then
                                    For liColIndex = ZeroL To miColumnCount - OneL
                                        If pDragDib_GetItem(liItem, liColIndex) Then
                                            pDragDib_DrawText lhDc, liItem, liColIndex, mtItem.pszText, liCols(liColIndex), liTextFillColor, ZeroL, liTextXMargin + CBool(liColIndex) * 4&, liTextYMargin, (CBool(liColIndex) And CBool(mtItem.iImage > NegOneL)) * -liIconWidth
                                            pDragDib_DrawIcon lhDc, liItem, liColIndex, lhIml, mtItem.iImage, liIconWidth, liIconHeight, liImlDrawFlags
                                        End If
                                    Next
                                End If
                            Next
                        
                        Else
                            'debug.assert miViewStyle = lvwTile
                        
                            ReDim liCols(0 To miColumnCount - OneL)
                        
                            With mtTile
                                .cbSize = Len(mtTile)
                                .puColumns = VarPtr(liCols(0))
                                .cColumns = miColumnCount
                            End With
                        
                            Dim liTextHeight      As Long
                            Dim ltSize            As SIZE
                        
                            GetTextExtentPoint32W lhDc, "A", OneL, ltSize
                            liTextHeight = ltSize.cy
                        
                            For liIndex = ZeroL To iIndexCount - OneL
                                liItem = iIndexes(liIndex)
                                If liItem > NegOneL Then
                                    If pDragDib_GetItem(liItem, ZeroL) Then
                                    
                                        mtTile.iItem = liItem
                                        mtTile.cColumns = miColumnCount
                                        SendMessage mhWnd, LVM_GETTILEINFO, ZeroL, VarPtr(mtTile)
                                    
                                        If bHideSelection Then SetTextColor lhDc, liTextColor
                                    
                                        ltRectItem.Left = LVIR_LABEL
                                        ltRectItem.Top = ZeroL
                                        SendMessage mhWnd, LVM_GETSUBITEMRECT, liItem, VarPtr(ltRectItem)
                                        InflateRect ltRectItem, ZeroL, -TwoL
                                    
                                        If Not (CBool(miStyleEx And LVS_EX_FULLROWSELECT) Or bHideSelection) Then
                                            ltRectItem.Right = ltRectItem.Left
                                            ltRectItem.Right = pDragDib_Max(ltRectItem.Right, ltRectItem.Left + pDragDib_TextWidth(lhDc, mtItem.pszText, lstrlen(mtItem.pszText)) + TwoL)
                                            For liColIndex = ZeroL To mtTile.cColumns - OneL
                                                mtItem.iSubItem = liCols(liColIndex)
                                                If SendMessage(mhWnd, LVM_GETITEMTEXTA, liItem, VarPtr(mtItem)) Then
                                                    ltRectItem.Right = pDragDib_Max(ltRectItem.Right, ltRectItem.Left + pDragDib_TextWidth(lhDc, mtItem.pszText, lstrlen(mtItem.pszText)) + TwoL)
                                                End If
                                            Next
                                            FillRect lhDc, ltRectItem, liTextFillColor
                                        
                                            mtItem.iSubItem = ZeroL
                                            If SendMessage(mhWnd, LVM_GETITEMTEXTA, liItem, VarPtr(mtItem)) Then
                                                pDragDib_DrawText lhDc, liItem, ZeroL, mtItem.pszText, liDTFlags, ZeroL, ZeroL, NegOneL, liTextYMargin, ZeroL
                                            End If
                                        
                                        Else
                                            If Not bHideSelection Then FillRect lhDc, ltRectItem, liTextFillColor
                                            pDragDib_DrawText lhDc, liItem, ZeroL, mtItem.pszText, liDTFlags, ZeroL, ZeroL, NegOneL, liTextYMargin, ZeroL
                                        End If
                                    
                                        pDragDib_DrawIcon lhDc, liItem, ZeroL, lhIml, mtItem.iImage, liIconWidth, liIconHeight, liImlDrawFlags
                                    
                                        If bHideSelection Then SetTextColor lhDc, &HD0C19
                                    
                                        ltRectItem.Left = ltRectItem.Left + OneL
                                        'ltRectItem.Top = ltRectItem.Top + TwoL
                                        For liColIndex = ZeroL To mtTile.cColumns - OneL
                                            mtItem.iSubItem = liCols(liColIndex)
                                            If SendMessage(mhWnd, LVM_GETITEMTEXTA, liItem, VarPtr(mtItem)) Then
                                                ltRectItem.Top = ltRectItem.Top + liTextHeight
                                                DrawText lhDc, ByVal mtItem.pszText, lstrlen(mtItem.pszText), ltRectItem, ZeroL
                                                pDragDib_DrawText lhDc, liItem, liCols(liColIndex), mtItem.pszText, liDTFlags, ZeroL, liTextXPadding, liTextXMargin, liTextYMargin, ZeroL
                                            End If
                                        Next
                                    End If
                                End If
                            Next
                        
                        End If
                    
                        SelectObject lhDc, lhBmpOld
                    
                        If ImageDrag_Alpha Then
                            OffsetRect ltRectBounds, -liXOffset, -liYOffset
                            pDragDib_FadeAlpha pDragDib, ltRectBounds
                        End If
                    
                    End If
                End If
            
                If lhFontOld Then SelectObject lhDc, lhFontOld
                DeleteDC lhDc
            
            End If
        
            'Return the offset to the cursor hotspot on the dib
            iHotSpotX = ltCursor.X - liXOffset
            iHotSpotY = ltCursor.Y - liYOffset
        End If
    
End Function

Private Function pDragDib_GetItem(ByVal iItem As Long, ByVal iSubItem As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Get the specified item information and return a success code.
    '---------------------------------------------------------------------------------------
    
    With mtItem
        .mask = LVIF_TEXT Or LVIF_IMAGE
        .iItem = iItem
        .iSubItem = iSubItem
        .iImage = NegOneL
    End With
    
    pDragDib_GetItem = CBool(SendMessage(mhWnd, LVM_GETITEMA, ZeroL, VarPtr(mtItem.mask)))
End Function

Private Sub pDragDib_DrawIcon(ByVal lhDc As Long, ByVal iItem As Long, ByVal iSubItemIndex As Long, ByVal lhIml As Long, ByVal iIndex As Long, ByVal iIconWidth As Long, ByVal iIconHeight As Long, ByVal iDrawFlags As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Draw the item or subitem icon.
    '---------------------------------------------------------------------------------------
    Dim ltRectIcon      As RECT
    ltRectIcon.Left = LVIR_ICON
    ltRectIcon.Top = iSubItemIndex
    If SendMessage(mhWnd, LVM_GETSUBITEMRECT, iItem, VarPtr(ltRectIcon)) Then
        With ltRectIcon
            ImageList_Draw lhIml, iIndex, lhDc, .Left + (.Right - .Left) \ TwoL - iIconWidth \ TwoL, .Top + (.Bottom - .Top) \ TwoL - iIconHeight \ TwoL, iDrawFlags
        End With
    End If
End Sub

Private Sub pDragDib_DrawText(ByVal lhDc As Long, ByVal iItem As Long, ByVal iSubItemIndex As Long, ByVal lpText As Long, ByVal iDrawFlags As Long, ByVal iFillColor As Long, ByVal iXPadding As Long, ByVal iXMargin As Long, ByVal iYMargin As Long, ByVal iXOffset As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Draw the item or subitem text.
    '---------------------------------------------------------------------------------------
    Dim ltRectItem      As RECT
    Dim liLen           As Long
    
    ltRectItem.Left = LVIR_LABEL
    ltRectItem.Top = iSubItemIndex
    If SendMessage(mhWnd, LVM_GETSUBITEMRECT, iItem, VarPtr(ltRectItem)) Then
        liLen = lstrlen(lpText)
        If iXPadding Then ltRectItem.Right = pDragDib_Min(ltRectItem.Right, ltRectItem.Left + pDragDib_TextWidth(lhDc, lpText, liLen) + iXPadding)
        If iFillColor Then FillRect lhDc, ltRectItem, iFillColor
        ltRectItem.Left = ltRectItem.Left + iXOffset
        InflateRect ltRectItem, iXMargin, iYMargin
        DrawText lhDc, ByVal lpText, liLen, ltRectItem, iDrawFlags
    End If
End Sub

Private Function pDragDib_TextWidth(ByVal lhDc As Long, ByVal lpText As Long, ByVal iLen As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Return the width of the text in the given hdc.
    '---------------------------------------------------------------------------------------
    Dim ltSize      As SIZE
    'debug.assert lpText
    If lpText Then
        If GetTextExtentPoint32(lhDc, ByVal lpText, iLen, ltSize) Then
            pDragDib_TextWidth = ltSize.cx
        Else
            'debug.assert False
        End If
    End If
End Function

Private Function pDragDib_HeaderHeight() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Return the height taken by the header control if in details view.
    '---------------------------------------------------------------------------------------
    Dim lhWndHeader      As Long
    lhWndHeader = SendMessage(mhWnd, LVM_GETHEADER, ZeroL, ZeroL)
    If lhWndHeader Then
        If GetWindowLong(lhWndHeader, GWL_STYLE) And WS_VISIBLE Then
            Dim ltR      As RECT
            GetWindowRect lhWndHeader, ltR
            pDragDib_HeaderHeight = ltR.Bottom - ltR.Top
        End If
    End If
End Function
Private Sub pDragDib_FadeAlpha(ByVal oDib As pcDibSection, ByRef tBounds As RECT)
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Fade the edges of an alpha dib to indicate whether the selected items
    '             exceed the viewable image.
    '---------------------------------------------------------------------------------------
    
    Dim ltRectInside      As RECT
    Dim liWidth           As Long
    Dim liHeight          As Long
    
    Dim liMaxFadeDist As Long
    Dim liMaxFadeDivisor As Long
    
    liMaxFadeDist = 80&
    liMaxFadeDivisor = 4&
    
    With ltRectInside
        liWidth = oDib.Width
        liHeight = oDib.Height
        .Right = liWidth
        .Bottom = liHeight
        
        If tBounds.Left < .Left Then .Left = pDragDib_Min(liWidth \ liMaxFadeDivisor, liMaxFadeDist)
        If tBounds.Right > liWidth Then .Right = pDragDib_Max(liWidth - (liWidth \ liMaxFadeDivisor), liWidth - liMaxFadeDist)
        
        If tBounds.Bottom > liHeight Then
            If tBounds.Top < .Top Then .Bottom = pDragDib_Max(liHeight - (liHeight \ liMaxFadeDivisor), liHeight - liMaxFadeDist)
            .Top = pDragDib_Min(liHeight \ liMaxFadeDivisor, liMaxFadeDist)
        Else
            If tBounds.Top < .Top Then .Bottom = pDragDib_Max(liHeight - (liHeight \ liMaxFadeDivisor), liHeight - liMaxFadeDist)
        End If
        
    End With
    
    Const Bkgnd As Long = ((ImageDrag_TransColor And &HFF0000) \ &H10000) Or _
    (ImageDrag_TransColor And &HFF00&) Or _
    ((ImageDrag_TransColor And &HFF&) * &H10000)
    Dim liBkgnd      As Long
    liBkgnd = Bkgnd
    
    Dim X      As Long
    Dim Y      As Long
    
    Dim liAlpha As Long
    
    Dim lyBits() As Byte
    SAPtr(lyBits) = oDib.ArrPtr(1)
    
    For X = LBound(lyBits, 1) To UBound(lyBits, 1) Step 4
        For Y = LBound(lyBits, 2) To UBound(lyBits, 2)
            If (MemOffset32(VarPtr(lyBits(X, Y)), ZeroL) And &HFFFFFF) = liBkgnd Then
                lyBits(X, Y) = ZeroY
                lyBits(X + 1, Y) = ZeroY
                lyBits(X + 2, Y) = ZeroY
                lyBits(X + 3, Y) = ZeroY
            Else
                liAlpha = pDragDib_GetAlpha(liWidth, liHeight, ltRectInside, X \ 4, Y)
                
                lyBits(X, Y) = liAlpha * lyBits(X, Y) \ &HFF&
                lyBits(X + 1, Y) = liAlpha * lyBits(X + 1, Y) \ &HFF&
                lyBits(X + 2, Y) = liAlpha * lyBits(X + 2, Y) \ &HFF&
                lyBits(X + 3, Y) = liAlpha
            End If
        Next
    Next

    SAPtr(lyBits) = ZeroL
    
End Sub

Private Function pDragDib_GetAlpha(ByVal iWidth As Long, ByVal iHeight As Long, ByRef tRectInside As RECT, ByVal X As Long, ByVal Y As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Return the new alpha value for a given point in the image to acheive
    '             a round rectangle fade effect.  This is done by determining which
    '             region the point lies in, and scaling ALPHA_FullStrength depending
    '             on the distance to the nearest boundary point on region 5.
    '
    '             Image Regions:
    '
    '                  1  |  2  | 3
    '                _____|_____|_____
    '                     |     |
    '                  4  |  5  |  6
    '                _____|_____|_____
    '                     |     |
    '                  7  |  8  |  9
    '                     |     |
    '
    '
    '             The argument tRectInside specifies the rectangle of region 5. Other regions
    '             are calculated using the iWidth and iHeight arguments.
    '
    '---------------------------------------------------------------------------------------
    
    Const ALPHA_FullStrength As Long = 150
    pDragDib_GetAlpha = ALPHA_FullStrength
    
    Const RGN_TopLeft As Long = 1
    Const RGN_TopCenter As Long = 2
    Const RGN_TopRight As Long = 3
    Const RGN_MiddleLeft As Long = 4
    Const RGN_MiddleCenter As Long = 5
    Const RGN_MiddleRight As Long = 6
    Const RGN_BottomLeft As Long = 7
    Const RGN_BottomCenter As Long = 8
    Const RGN_BottomRight As Long = 9
    
    'if the point is Not in RGN_MiddleCenter
    If Not pDragDib_PtInRect(tRectInside, X, Y) Then
        Dim lX         As Long
        Dim lY         As Long
        Dim liRgn      As Long
        
        With tRectInside
            'FIRST: Determine which region contains the point and the closest
            '       point to the region on the boundary line of region 5.
            If X < .Left Then
                liRgn = RGN_TopLeft
                lX = .Left
            ElseIf X > .Right Then
                liRgn = RGN_TopRight
                lX = .Right
            Else
                liRgn = RGN_TopCenter
                lX = X
            End If
            
            If Y >= .Top Then
                If Y > .Bottom Then
                    liRgn = pDragDib_Choose(liRgn, RGN_BottomLeft, RGN_BottomCenter, RGN_BottomRight)
                    lY = .Bottom
                Else
                    liRgn = pDragDib_Choose(liRgn, RGN_MiddleLeft, RGN_MiddleCenter, RGN_MiddleRight)
                    lY = Y
                End If
            Else
                lY = .Top
            End If
            
            'debug.assert liRgn <> RGN_MiddleCenter
            
            'NEXT: Determine the distance from the nearest point on the region 5 boundary
            '      line until the end of the fade into the background.
            
            Dim liMaxDist      As Long
            Dim liDist         As Long
            
            If liRgn And OneL Then
                'Odd numbered regions
                If liRgn = RGN_TopLeft Then
                    liMaxDist = pDragDib_GetMaxFadeDist(X + 3, Y, .Left, .Top)
                ElseIf liRgn = RGN_TopRight Then
                    liMaxDist = pDragDib_GetMaxFadeDist(X + 3 - .Right, .Top - Y, iWidth - .Right, .Top)
                ElseIf liRgn = RGN_BottomLeft Then
                    liMaxDist = pDragDib_GetMaxFadeDist(X + 3, Y - .Bottom, .Left, iHeight - .Bottom)
                ElseIf liRgn = RGN_BottomRight Then
                    liMaxDist = pDragDib_GetMaxFadeDist(X + 3 - .Right, Y - .Bottom, iWidth - .Right, iHeight - .Bottom)
                End If
            Else
                'even numbered regions
                If liRgn = RGN_TopCenter Then
                    liMaxDist = .Top
                ElseIf liRgn = RGN_MiddleLeft Then
                    liMaxDist = .Left
                ElseIf liRgn = RGN_MiddleRight Then
                    liMaxDist = iWidth - .Right
                ElseIf liRgn = RGN_BottomCenter Then
                    liMaxDist = iHeight - .Bottom
                End If
            End If
            
            'debug.assert liMaxDist
            
            pDragDib_GetAlpha = pDragDib_ScaleAlpha(pDragDib_GetAlpha, pDragDib_Dist(X, Y, lX, lY), liMaxDist)
            
        End With
    End If
    
End Function

Private Function pDragDib_ScaleAlpha(ByVal iAlpha As Long, ByVal c As Long, ByVal d As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Return an alpha value scaled from 0 to d based on the ratio of c to d.
    '---------------------------------------------------------------------------------------
    If c > d Then c = d
    If d Then pDragDib_ScaleAlpha = iAlpha - (iAlpha * c \ d)
End Function

Private Function pDragDib_GetMaxFadeDist(ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Return a value scaled from cx to cy based on the ratio of x to y.
    '---------------------------------------------------------------------------------------
    If (X Or Y) = ZeroL Then Y = OneL
    pDragDib_GetMaxFadeDist = cx + ((cy - cx) * Y \ (X + Y))
End Function

Private Function pDragDib_PtInRect(ByRef tR As RECT, ByVal X As Long, ByVal Y As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : PtInRect implementation that is inclusive of the right and bottom edges.
    '---------------------------------------------------------------------------------------
    pDragDib_PtInRect = CBool(Y >= tR.Top And Y <= tR.Bottom And X <= tR.Right And X >= tR.Left)
End Function

Private Function pDragDib_Choose(ByVal iIndex As Long, ByVal i1 As Long, ByVal i2 As Long, ByVal i3 As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Return one of the three values passed in.
    '---------------------------------------------------------------------------------------
    If iIndex = OneL Then
        pDragDib_Choose = i1
    ElseIf iIndex = TwoL Then
        pDragDib_Choose = i2
    ElseIf iIndex = 3& Then
        pDragDib_Choose = i3
    End If
End Function

Private Function pDragDib_Min(ByVal i1 As Long, ByVal i2 As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Return the higher number.
    '---------------------------------------------------------------------------------------
    If i1 < i2 Then
        pDragDib_Min = i1
    Else
        pDragDib_Min = i2
    End If
End Function

Private Function pDragDib_Max(ByVal i1 As Long, ByVal i2 As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Return the lower number.
    '---------------------------------------------------------------------------------------
    If i1 < i2 Then
        pDragDib_Max = i2
    Else
        pDragDib_Max = i1
    End If
End Function

Private Function pDragDib_Dist(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Return the distance between two points.
    '---------------------------------------------------------------------------------------
    If X1 = X2 Then
        pDragDib_Dist = Abs(Y1 - Y2)
    ElseIf Y1 = Y2 Then
        pDragDib_Dist = Abs(X1 - X2)
    Else
        '(a^2) + (b^2) = (c^2)
        Dim a      As Long: a = X2 - X1
        Dim B      As Long: B = Y2 - Y1
        pDragDib_Dist = Sqr(a * a + B * B)
    End If
End Function

Public Property Get OLERegisterDrop() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 3/5/05
    ' Purpose   : Return a value indicating whether the control registers itself as an ole
    '             drag-drop target.
    '---------------------------------------------------------------------------------------
    OLERegisterDrop = -UserControl.OLEDropMode
End Property

Public Property Let OLERegisterDrop(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 3/5/05
    ' Purpose   : Set whether the control registers itself as an ole drag-drop target.
    '---------------------------------------------------------------------------------------
    UserControl.OLEDropMode = -bNew
    pPropChanged PROP_OleDrop
End Property
