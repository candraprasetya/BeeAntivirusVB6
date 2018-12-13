Attribute VB_Name = "mPopup"
'==================================================================================================
'mPopup.bas                      7/15/05
'
'           PURPOSE:
'               Manage memory structures used for popup menus, items and hierarchies.
'               Manage globally unique popup item ids using a hash table.
'
'==================================================================================================

Option Explicit

Private Const ITEM_Style                As Long = 0
Private Const ITEM_Id                   As Long = 4
Private Const ITEM_ItemData             As Long = 8
Private Const ITEM_ShortcutKey          As Long = 12
Private Const ITEM_ShortcutMask         As Long = 16
Private Const ITEM_Accelerator          As Long = 20
Private Const ITEM_IconIndex            As Long = 24
Private Const ITEM_lpCaption            As Long = 28
Private Const ITEM_lpHelp               As Long = 32
Private Const ITEM_lpKey                As Long = 36
Private Const ITEM_lpShortcutDisplay    As Long = 40
Private Const ITEM_pItemNext            As Long = 44
Private Const ITEM_pItemPrev            As Long = 48
Private Const ITEM_pMenuParent          As Long = 52
Private Const ITEM_pMenuChild           As Long = 56
Private Const ITEM_RefCount             As Long = 60
Private Const ITEM_Len                  As Long = 64

Private Const MENU_hMenu                As Long = 0
Private Const MENU_pItemFirst           As Long = 4
Private Const MENU_pItemLast            As Long = 8
Private Const MENU_Control              As Long = 12
Private Const MENU_Style                As Long = 16
Private Const MENU_pItemParent          As Long = 20
Private Const MENU_SidebarHdc           As Long = 24
Private Const MENU_SidebarHbmp          As Long = 28
Private Const MENU_SidebarHbmpOld       As Long = 32
Private Const MENU_SidebarWidth         As Long = 36
Private Const MENU_SidebarHeight        As Long = 40
Private Const MENU_pMenuNext            As Long = 44
Private Const MENU_pHierarchy           As Long = 48
Private Const MENU_RefCount             As Long = 52
Private Const MENU_Len                  As Long = 56

Private Const HARCHY_iID                As Long = 0
Private Const HARCHY_pHierarchyNext     As Long = 4
Private Const HARCHY_Len                As Long = 8

Private Const ID_Id                     As Long = 0
Private Const ID_pNextId                As Long = 4
Private Const ID_Len                    As Long = 8

Private mpIds(0 To 255) As Long     'hash table/linked lists

Public Property Get PopupItem_Initialize(ByVal iStyle As Long, ByVal iId As Long, ByVal iItemData As Long, ByVal iShortcutKey As Long, ByVal iShortcutMask As Long, ByVal iAccelerator As Long, ByVal iIconIndex As Long, ByVal lpCaption As Long, ByVal lpHelp As Long, ByVal lpKey As Long, ByVal lpShortcutDisplay As Long, ByVal pItemNext As Long, ByVal pItemPrev As Long, ByVal pMenuParent As Long) As Long
    
    PopupItem_Initialize = MemAlloc(ITEM_Len)
    
    Debug.Assert PopupItem_Initialize
    
    If PopupItem_Initialize Then
        MemOffset32(PopupItem_Initialize, ITEM_Style) = iStyle
        MemOffset32(PopupItem_Initialize, ITEM_Id) = iId
        MemOffset32(PopupItem_Initialize, ITEM_ItemData) = iItemData
        MemOffset32(PopupItem_Initialize, ITEM_ShortcutKey) = iShortcutKey
        MemOffset32(PopupItem_Initialize, ITEM_ShortcutMask) = iShortcutMask
        MemOffset32(PopupItem_Initialize, ITEM_Accelerator) = iAccelerator
        MemOffset32(PopupItem_Initialize, ITEM_IconIndex) = iIconIndex
        MemOffset32(PopupItem_Initialize, ITEM_lpCaption) = lpCaption
        MemOffset32(PopupItem_Initialize, ITEM_lpHelp) = lpHelp
        MemOffset32(PopupItem_Initialize, ITEM_lpKey) = lpKey
        MemOffset32(PopupItem_Initialize, ITEM_lpShortcutDisplay) = lpShortcutDisplay
        MemOffset32(PopupItem_Initialize, ITEM_pItemNext) = pItemNext
        MemOffset32(PopupItem_Initialize, ITEM_pItemPrev) = pItemPrev
        MemOffset32(PopupItem_Initialize, ITEM_pMenuParent) = pMenuParent
        MemOffset32(PopupItem_Initialize, ITEM_pMenuChild) = ZeroL
        MemOffset32(PopupItem_Initialize, ITEM_RefCount) = ZeroL
    End If
    
End Property

Public Function PopupItem_Terminate(ByVal pItem As Long) As Long
    Debug.Assert pItem
    If pItem Then PopupItem_Terminate = MemFree(pItem)
End Function

Public Property Get PopupMenu_Initialize(Optional ByVal pItemParent As Long) As Long
    
    PopupMenu_Initialize = MemAlloc(MENU_Len)
    
    Debug.Assert PopupMenu_Initialize
    
    If PopupMenu_Initialize Then
        MemOffset32(PopupMenu_Initialize, MENU_hMenu) = ZeroL
        MemOffset32(PopupMenu_Initialize, MENU_pItemFirst) = ZeroL
        MemOffset32(PopupMenu_Initialize, MENU_pItemLast) = ZeroL
        MemOffset32(PopupMenu_Initialize, MENU_Control) = ZeroL
        MemOffset32(PopupMenu_Initialize, MENU_Style) = ZeroL
        MemOffset32(PopupMenu_Initialize, MENU_pItemParent) = pItemParent
        MemOffset32(PopupMenu_Initialize, MENU_SidebarHdc) = ZeroL
        MemOffset32(PopupMenu_Initialize, MENU_SidebarHbmp) = ZeroL
        MemOffset32(PopupMenu_Initialize, MENU_SidebarHbmpOld) = ZeroL
        MemOffset32(PopupMenu_Initialize, MENU_SidebarWidth) = ZeroL
        MemOffset32(PopupMenu_Initialize, MENU_SidebarHeight) = ZeroL
        MemOffset32(PopupMenu_Initialize, MENU_pMenuNext) = ZeroL
        MemOffset32(PopupMenu_Initialize, MENU_pHierarchy) = ZeroL
        MemOffset32(PopupMenu_Initialize, MENU_RefCount) = ZeroL
    End If
    
End Property

Public Function PopupMenu_Terminate(ByVal pMenu As Long) As Long
    Debug.Assert pMenu
    If pMenu Then PopupMenu_Terminate = MemFree(pMenu)
End Function

Public Property Get Hierarchy_Initialize(ByVal iId As Long, ByVal pHierarchyNext As Long) As Long
    Hierarchy_Initialize = MemAlloc(HARCHY_Len)
    
    Debug.Assert Hierarchy_Initialize
    
    If Hierarchy_Initialize Then
        MemOffset32(Hierarchy_Initialize, HARCHY_iID) = iId
        MemOffset32(Hierarchy_Initialize, HARCHY_pHierarchyNext) = pHierarchyNext
    End If
    
End Property

Public Function Hierarchy_Terminate(ByVal pHierarchy As Long) As Long
    Debug.Assert pHierarchy
    If pHierarchy Then Hierarchy_Terminate = MemFree(pHierarchy)

End Function

Public Property Get Hierarchy_iId(ByVal pHierarchy As Long) As Long
    Debug.Assert pHierarchy
    If pHierarchy Then Hierarchy_iId = MemOffset32(pHierarchy, HARCHY_iID)
End Property
Public Property Let Hierarchy_iId(ByVal pHierarchy As Long, ByVal iNew As Long)
    Debug.Assert pHierarchy
    If pHierarchy Then MemOffset32(pHierarchy, HARCHY_iID) = iNew
End Property

Public Property Get Hierarchy_pHierarchyNext(ByVal pHierarchy As Long) As Long
    Debug.Assert pHierarchy
    If pHierarchy Then Hierarchy_pHierarchyNext = MemOffset32(pHierarchy, HARCHY_pHierarchyNext)
End Property
Public Property Let Hierarchy_pHierarchyNext(ByVal pHierarchy As Long, ByVal iNew As Long)
    Debug.Assert pHierarchy
    If pHierarchy Then MemOffset32(pHierarchy, HARCHY_pHierarchyNext) = iNew
End Property




Public Property Get PopupItem_Style(ByVal pItem As Long) As Long
    Debug.Assert pItem
    If pItem Then PopupItem_Style = MemOffset32(pItem, ITEM_Style)
End Property
Public Property Let PopupItem_Style(ByVal pItem As Long, ByVal iNew As Long)
    Debug.Assert pItem
    If pItem Then MemOffset32(pItem, ITEM_Style) = iNew
End Property

Public Property Get PopupItem_RefCount(ByVal pItem As Long) As Long
    Debug.Assert pItem
    If pItem Then PopupItem_RefCount = MemOffset32(pItem, ITEM_RefCount)
End Property
Public Property Let PopupItem_RefCount(ByVal pItem As Long, ByVal iNew As Long)
    Debug.Assert pItem
    If pItem Then MemOffset32(pItem, ITEM_RefCount) = iNew
End Property

Public Property Get PopupItem_Id(ByVal pItem As Long) As Long
    Debug.Assert pItem
    If pItem Then PopupItem_Id = MemOffset32(pItem, ITEM_Id)
End Property
Public Property Let PopupItem_Id(ByVal pItem As Long, ByVal iNew As Long)
    Debug.Assert pItem
    If pItem Then MemOffset32(pItem, ITEM_Id) = iNew
End Property

Public Property Get PopupItem_ItemData(ByVal pItem As Long) As Long
    Debug.Assert pItem
    If pItem Then PopupItem_ItemData = MemOffset32(pItem, ITEM_ItemData)
End Property
Public Property Let PopupItem_ItemData(ByVal pItem As Long, ByVal iNew As Long)
    Debug.Assert pItem
    If pItem Then MemOffset32(pItem, ITEM_ItemData) = iNew
End Property

Public Property Get PopupItem_ShortcutKey(ByVal pItem As Long) As Long
    Debug.Assert pItem
    If pItem Then PopupItem_ShortcutKey = MemOffset32(pItem, ITEM_ShortcutKey)
End Property
Public Property Let PopupItem_ShortcutKey(ByVal pItem As Long, ByVal iNew As Long)
    Debug.Assert pItem
    If pItem Then MemOffset32(pItem, ITEM_ShortcutKey) = iNew
End Property

Public Property Get PopupItem_ShortcutMask(ByVal pItem As Long) As Long
    Debug.Assert pItem
    If pItem Then PopupItem_ShortcutMask = MemOffset32(pItem, ITEM_ShortcutMask)
End Property
Public Property Let PopupItem_ShortcutMask(ByVal pItem As Long, ByVal iNew As Long)
    Debug.Assert pItem
    If pItem Then MemOffset32(pItem, ITEM_ShortcutMask) = iNew
End Property

Public Property Get PopupItem_Accelerator(ByVal pItem As Long) As Long
    Debug.Assert pItem
    If pItem Then PopupItem_Accelerator = MemOffset32(pItem, ITEM_Accelerator)
End Property
Public Property Let PopupItem_Accelerator(ByVal pItem As Long, ByVal iNew As Long)
    Debug.Assert pItem
    If pItem Then MemOffset32(pItem, ITEM_Accelerator) = iNew
End Property

Public Property Get PopupItem_IconIndex(ByVal pItem As Long) As Long
    Debug.Assert pItem
    If pItem Then PopupItem_IconIndex = MemOffset32(pItem, ITEM_IconIndex)
End Property
Public Property Let PopupItem_IconIndex(ByVal pItem As Long, ByVal iNew As Long)
    Debug.Assert pItem
    If pItem Then MemOffset32(pItem, ITEM_IconIndex) = iNew
End Property

Public Property Get PopupItem_lpCaption(ByVal pItem As Long) As Long
    Debug.Assert pItem
    If pItem Then PopupItem_lpCaption = MemOffset32(pItem, ITEM_lpCaption)
End Property
Public Property Let PopupItem_lpCaption(ByVal pItem As Long, ByVal iNew As Long)
    Debug.Assert pItem
    If pItem Then MemOffset32(pItem, ITEM_lpCaption) = iNew
End Property

Public Property Get PopupItem_lpHelp(ByVal pItem As Long) As Long
    Debug.Assert pItem
    If pItem Then PopupItem_lpHelp = MemOffset32(pItem, ITEM_lpHelp)
End Property
Public Property Let PopupItem_lpHelp(ByVal pItem As Long, ByVal iNew As Long)
    Debug.Assert pItem
    If pItem Then MemOffset32(pItem, ITEM_lpHelp) = iNew
End Property

Public Property Get PopupItem_lpKey(ByVal pItem As Long) As Long
    Debug.Assert pItem
    If pItem Then PopupItem_lpKey = MemOffset32(pItem, ITEM_lpKey)
End Property
Public Property Let PopupItem_lpKey(ByVal pItem As Long, ByVal iNew As Long)
    Debug.Assert pItem
    If pItem Then MemOffset32(pItem, ITEM_lpKey) = iNew
End Property

Public Property Get PopupItem_lpShortcutDisplay(ByVal pItem As Long) As Long
    Debug.Assert pItem
    If pItem Then PopupItem_lpShortcutDisplay = MemOffset32(pItem, ITEM_lpShortcutDisplay)
End Property
Public Property Let PopupItem_lpShortcutDisplay(ByVal pItem As Long, ByVal iNew As Long)
    Debug.Assert pItem
    If pItem Then MemOffset32(pItem, ITEM_lpShortcutDisplay) = iNew
End Property

Public Property Get PopupItem_pItemNext(ByVal pItem As Long) As Long
    Debug.Assert pItem
    If pItem Then PopupItem_pItemNext = MemOffset32(pItem, ITEM_pItemNext)
End Property
Public Property Let PopupItem_pItemNext(ByVal pItem As Long, ByVal iNew As Long)
    Debug.Assert pItem
    If pItem Then MemOffset32(pItem, ITEM_pItemNext) = iNew
End Property

Public Property Get PopupItem_pItemPrev(ByVal pItem As Long) As Long
    Debug.Assert pItem
    If pItem Then PopupItem_pItemPrev = MemOffset32(pItem, ITEM_pItemPrev)
End Property
Public Property Let PopupItem_pItemPrev(ByVal pItem As Long, ByVal iNew As Long)
    Debug.Assert pItem
    If pItem Then MemOffset32(pItem, ITEM_pItemPrev) = iNew
End Property

Public Property Get PopupItem_pMenuParent(ByVal pItem As Long) As Long
    Debug.Assert pItem
    If pItem Then PopupItem_pMenuParent = MemOffset32(pItem, ITEM_pMenuParent)
End Property
Public Property Let PopupItem_pMenuParent(ByVal pItem As Long, ByVal iNew As Long)
    Debug.Assert pItem
    If pItem Then MemOffset32(pItem, ITEM_pMenuParent) = iNew
End Property

Public Property Get PopupItem_pMenuChild(ByVal pItem As Long) As Long
    Debug.Assert pItem
    If pItem Then PopupItem_pMenuChild = MemOffset32(pItem, ITEM_pMenuChild)
End Property
Public Property Let PopupItem_pMenuChild(ByVal pItem As Long, ByVal iNew As Long)
    Debug.Assert pItem
    If pItem Then MemOffset32(pItem, ITEM_pMenuChild) = iNew
End Property


Public Property Get PopupMenu_hMenu(ByVal pMenu As Long) As Long
    Debug.Assert pMenu
    If pMenu Then PopupMenu_hMenu = MemOffset32(pMenu, MENU_hMenu)
End Property
Public Property Let PopupMenu_hMenu(ByVal pMenu As Long, ByVal iNew As Long)
    Debug.Assert pMenu
    If pMenu Then MemOffset32(pMenu, MENU_hMenu) = iNew
End Property

Public Property Get PopupMenu_pItemFirst(ByVal pMenu As Long) As Long
    Debug.Assert pMenu
    If pMenu Then PopupMenu_pItemFirst = MemOffset32(pMenu, MENU_pItemFirst)
End Property
Public Property Let PopupMenu_pItemFirst(ByVal pMenu As Long, ByVal iNew As Long)
    Debug.Assert pMenu
    If pMenu Then MemOffset32(pMenu, MENU_pItemFirst) = iNew
End Property

Public Property Get PopupMenu_pItemLast(ByVal pMenu As Long) As Long
    Debug.Assert pMenu
    If pMenu Then PopupMenu_pItemLast = MemOffset32(pMenu, MENU_pItemLast)
End Property
Public Property Let PopupMenu_pItemLast(ByVal pMenu As Long, ByVal iNew As Long)
    Debug.Assert pMenu
    If pMenu Then MemOffset32(pMenu, MENU_pItemLast) = iNew
End Property

Public Property Get PopupMenu_Control(ByVal pMenu As Long) As Long
    Debug.Assert pMenu
    If pMenu Then PopupMenu_Control = MemOffset32(pMenu, MENU_Control)
End Property
Public Property Let PopupMenu_Control(ByVal pMenu As Long, ByVal iNew As Long)
    Debug.Assert pMenu
    If pMenu Then MemOffset32(pMenu, MENU_Control) = iNew
End Property

Public Property Get PopupMenu_SidebarHdc(ByVal pMenu As Long) As Long
    Debug.Assert pMenu
    If pMenu Then PopupMenu_SidebarHdc = MemOffset32(pMenu, MENU_SidebarHdc)
End Property
Public Property Let PopupMenu_SidebarHdc(ByVal pMenu As Long, ByVal iNew As Long)
    Debug.Assert pMenu
    If pMenu Then MemOffset32(pMenu, MENU_SidebarHdc) = iNew
End Property

Public Property Get PopupMenu_SidebarHbmp(ByVal pMenu As Long) As Long
    Debug.Assert pMenu
    If pMenu Then PopupMenu_SidebarHbmp = MemOffset32(pMenu, MENU_SidebarHbmp)
End Property
Public Property Let PopupMenu_SidebarHbmp(ByVal pMenu As Long, ByVal iNew As Long)
    Debug.Assert pMenu
    If pMenu Then MemOffset32(pMenu, MENU_SidebarHbmp) = iNew
End Property

Public Property Get PopupMenu_SidebarHbmpOld(ByVal pMenu As Long) As Long
    Debug.Assert pMenu
    If pMenu Then PopupMenu_SidebarHbmpOld = MemOffset32(pMenu, MENU_SidebarHbmpOld)
End Property
Public Property Let PopupMenu_SidebarHbmpOld(ByVal pMenu As Long, ByVal iNew As Long)
    Debug.Assert pMenu
    If pMenu Then MemOffset32(pMenu, MENU_SidebarHbmpOld) = iNew
End Property

Public Property Get PopupMenu_SidebarWidth(ByVal pMenu As Long) As Long
    Debug.Assert pMenu
    If pMenu Then PopupMenu_SidebarWidth = MemOffset32(pMenu, MENU_SidebarWidth)
End Property
Public Property Let PopupMenu_SidebarWidth(ByVal pMenu As Long, ByVal iNew As Long)
    Debug.Assert pMenu
    If pMenu Then MemOffset32(pMenu, MENU_SidebarWidth) = iNew
End Property

Public Property Get PopupMenu_SidebarHeight(ByVal pMenu As Long) As Long
    Debug.Assert pMenu
    If pMenu Then PopupMenu_SidebarHeight = MemOffset32(pMenu, MENU_SidebarHeight)
End Property
Public Property Let PopupMenu_SidebarHeight(ByVal pMenu As Long, ByVal iNew As Long)
    Debug.Assert pMenu
    If pMenu Then MemOffset32(pMenu, MENU_SidebarHeight) = iNew
End Property

Public Property Get PopupMenu_Style(ByVal pMenu As Long) As Long
    Debug.Assert pMenu
    If pMenu Then PopupMenu_Style = MemOffset32(pMenu, MENU_Style)
End Property
Public Property Let PopupMenu_Style(ByVal pMenu As Long, ByVal iNew As Long)
    Debug.Assert pMenu
    If pMenu Then MemOffset32(pMenu, MENU_Style) = iNew
End Property

Public Property Get PopupMenu_pItemParent(ByVal pMenu As Long) As Long
    Debug.Assert pMenu
    If pMenu Then PopupMenu_pItemParent = MemOffset32(pMenu, MENU_pItemParent)
End Property
Public Property Let PopupMenu_pItemParent(ByVal pMenu As Long, ByVal iNew As Long)
    Debug.Assert pMenu
    If pMenu Then MemOffset32(pMenu, MENU_pItemParent) = iNew
End Property

Public Property Get PopupMenu_pMenuNext(ByVal pMenu As Long) As Long
    Debug.Assert pMenu
    If pMenu Then PopupMenu_pMenuNext = MemOffset32(pMenu, MENU_pMenuNext)
End Property
Public Property Let PopupMenu_pMenuNext(ByVal pMenu As Long, ByVal iNew As Long)
    Debug.Assert pMenu
    If pMenu Then MemOffset32(pMenu, MENU_pMenuNext) = iNew
End Property

Public Property Get PopupMenu_pHierarchy(ByVal pMenu As Long) As Long
    Debug.Assert pMenu
    If pMenu Then PopupMenu_pHierarchy = MemOffset32(pMenu, MENU_pHierarchy)
End Property
Public Property Let PopupMenu_pHierarchy(ByVal pMenu As Long, ByVal iNew As Long)
    Debug.Assert pMenu
    If pMenu Then MemOffset32(pMenu, MENU_pHierarchy) = iNew
End Property

Public Property Get PopupMenu_RefCount(ByVal pMenu As Long) As Long
    Debug.Assert pMenu
    If pMenu Then PopupMenu_RefCount = MemOffset32(pMenu, MENU_RefCount)
End Property
Public Property Let PopupMenu_RefCount(ByVal pMenu As Long, ByVal iNew As Long)
    Debug.Assert pMenu
    If pMenu Then MemOffset32(pMenu, MENU_RefCount) = iNew
End Property



Public Function PopupMenus_GetID() As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Get a unique menu id. This is outside the range 0x0000 to 0x0800 to avoid
'             possible conflicts with any vb menus and also must not conflict with submenu
'             handles which are also used as menu ids.
'---------------------------------------------------------------------------------------
    
    Static id As Long
    Dim liIndex As Long

    liIndex = NegOneL

    Do

        If id = &H7FFFFFFF Then
            id = &H80000000
        ElseIf id >= NegOneL And id < &H800& Then
            id = &H801&
        Else
            id = id + OneL
        End If

        liIndex = PopupMenus_AddID(id)
        
    Loop While liIndex = NegOneL

    PopupMenus_GetID = id

End Function

Public Function PopupMenus_AddID(ByVal iId As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Try to add a specific id.  If the id is being used, return 0.  Otherwise
'             store the id as used and return the memory handle.
'---------------------------------------------------------------------------------------
    
    Dim liHashTableIndex As Long
    liHashTableIndex = HashLong(iId)
    
    Dim lpId As Long
    Dim lpIdPrev As Long
    
    lpId = mpIds(liHashTableIndex)
    
    Do While lpId
        If MemOffset32(lpId, ID_Id) = iId Then Exit Function
        lpIdPrev = lpId
        lpId = MemOffset32(lpIdPrev, ID_pNextId)
    Loop
    
    PopupMenus_AddID = MemAlloc(ID_Len)
    Debug.Assert PopupMenus_AddID
        
    If PopupMenus_AddID Then
        
        MemOffset32(PopupMenus_AddID, ID_Id) = iId
        MemOffset32(PopupMenus_AddID, ID_pNextId) = mpIds(liHashTableIndex)
        
        mpIds(liHashTableIndex) = PopupMenus_AddID
        
    End If
    
End Function

Public Function PopupMenus_ReleaseId(ByVal iId As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Release the memory allocated to store a given id.
'---------------------------------------------------------------------------------------
    Dim liHashTableIndex As Long
    liHashTableIndex = HashLong(iId)
    
    Dim lpId As Long
    Dim lpIdPrev As Long
    
    lpId = mpIds(liHashTableIndex)
    
    Do While lpId
        If MemOffset32(lpId, ID_Id) = iId Then Exit Do
        lpIdPrev = lpId
        lpId = MemOffset32(lpIdPrev, ID_pNextId)
    Loop
    
    Debug.Assert lpId
    
    If lpId Then
        
        If lpIdPrev _
            Then MemOffset32(lpIdPrev, ID_pNextId) = MemOffset32(lpId, ID_pNextId) _
            Else mpIds(liHashTableIndex) = MemOffset32(lpId, ID_pNextId)
        
        PopupMenus_ReleaseId = MemFree(lpId)
        
    End If
    
End Function

