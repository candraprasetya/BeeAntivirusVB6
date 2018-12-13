Attribute VB_Name = "mSubclass"
'---------------------------------------------------------------------------------------
'mSubclass.bas           8/24/05
'
'            PURPOSE:
'                IDE safe subclassing that allows multiple objects to subclass the same
'                window in any order.
'
'            LINEAGE:
'                http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=51403&lngWId=1
'
'---------------------------------------------------------------------------------------

Option Explicit

#Const bUseASM = True                                       'If True, ASM is loaded from the resource file and provides fast
'callbacks and IDE protection.  If False, the callback function is
'a regular VB function with no IDE protection. Best to set to True
'unless debugging the assembler.

Public Const ALL_MESSAGES                 As Long = -1        'a special value that indicates all messages call back to the iSubclass interface

'Dynamically allocated memory structures                    'Private Type tSubclass
Private Const SUBCLASS_pWndProcPrev     As Long = 0          '    pWndProcPrev As Long
Private Const SUBCLASS_pSubclassNext    As Long = 4          '    pSubclassNext As Long
Private Const SUBCLASS_pObject          As Long = 8          '    pObject As Long
Private Const SUBCLASS_pMsgTableA       As Long = 12         '    pMsgTableA As Long '0 for no messages, ALL_MESSAGES for all messages, otherwise ptr to a tMsgTable
Private Const SUBCLASS_pMsgTableB       As Long = 16         '    pMsgTableB As Long '0 for no messages, ALL_MESSAGES for all messages, otherwise ptr to a tMsgTable
Private Const SUBCLASS_Len              As Long = 20         '    tMsgTableA as tMsgTable
'    tMsgTableB as tMsgTable
'End Type
                                                            
'Private Type tMsgTable
'    cMessages As Long
'    iMessages(0 to this.cMessages - 1) As Long
'End Type

Private Const sPropName                  As String = "Sub$%" 'Window property string identifier

Public Function Subclass_Install(ByVal oOwner As iSubclass, ByVal hWnd As Long, Optional ByVal vMessagesBefore As Variant, Optional ByVal vMessagesAfter As Variant) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 8/24/05
    ' Purpose   : Subclass a window and set up the data structures to call back on the given iSubclass object.
    '
    ' Arguments : oOwner            iSubclass interface on which callbacks are to be made
    '             hWnd              the window handle that is subclassed
    '             vMessagesBefore   messages called back before the original procedure
    '             vMessagesAfter    messages called back after the original procedure
    '
    '             vMessagesBefore/vMessagesAfter arguments can be:
    '                   ALL_MESSAGES
    '                   a single message value, i.e. WM_MOUSEMOVE
    '                   an array of message values, i.e. Array(WM_MOUSEMOVE, WM_LBUTTONDOWN, WM_RBUTTONDOWN)
    '---------------------------------------------------------------------------------------
    
    Dim lpSubclassProc      As Long: lpSubclassProc = pSubclassProc()                                            'get a pointer to the subclass procedure
    If lpSubclassProc Then                                                                                  'proceed only if the pointer was obtained
        Dim lpObject As Long: lpObject = ObjPtr(oOwner)                                                     'get a pointer to the iSubclass interface
        If lpObject Then                                                                                    'proceed only if the iSubclass object is not Nothing
            Dim lpSubclass As Long: lpSubclass = pWindowProp(hWnd)                                          'get the first data structure in the current list
            If pFindSubclass(lpSubclass, lpObject) = 0& Then                                                'proceed only if the object has not already subclassed this window
                Dim lpWndProcPrev As Long: If lpSubclass Then lpWndProcPrev = MemOffset32(lpSubclass, SUBCLASS_pWndProcPrev) 'store the previous wndproc address
                lpSubclass = pAllocSubclass(lpObject, vMessagesBefore, vMessagesAfter, lpSubclass)          'allocate a new subclass data structure
                If lpSubclass Then                                                                          'proceed only if the structure was successfully allocated
                    If lpWndProcPrev = 0& Then lpWndProcPrev = SetWindowLong(hWnd, GWL_WNDPROC, lpSubclassProc) 'Subclass the window if not subclassed already
                        If lpWndProcPrev Then                                                                   'proceed only if the window has been subclassed
                            Subclass_Install = True                                                             'indicate success
                            MemOffset32(lpSubclass, SUBCLASS_pWndProcPrev) = lpWndProcPrev                      'store the previous wndproc
                            pWindowProp(hWnd) = lpSubclass                                                      'store the subclass data structure as the first in the list
                        Else
                            MemFree lpSubclass                                                                  'on failure, free the memory we would have used
                        End If
                    End If
                End If
            End If
        End If
End Function

Public Function Subclass_Remove(ByVal oOwner As iSubclass, ByVal hWnd As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 8/24/05
    ' Purpose   : Cease callbacks on the given interface and remove the subclass if no other
    '             objects are receiving callbacks.
    '
    ' Arguments : oOwner            iSubclass interface on which callbacks were being made
    '             hWnd              the window handle that is subclassed
    '---------------------------------------------------------------------------------------
    Dim lpObject      As Long: lpObject = ObjPtr(oOwner)                                                         'get a pointer to the iSubclass interface
    If lpObject Then                                                                                        'proceed only if the iSubclass object is not Nothing
        Dim lpSubclass As Long, lpSubclassPrev As Long
        lpSubclass = pFindSubclass(pWindowProp(hWnd), lpObject, lpSubclassPrev)                             'allocate a subclass data structure
        If lpSubclass Then                                                                                  'proceed only if the structure was allocated
            Subclass_Remove = True                                                                          'indicate success
            Dim lpSubclassNext      As Long: lpSubclassNext = MemOffset32(lpSubclass, SUBCLASS_pSubclassNext)    'get a pointer to the next subclass data structure in the linked list
            If (lpSubclassPrev Or lpSubclassNext) = 0& Then                                                 'if there is neither a previous or next data structure, then
                SetWindowLong hWnd, GWL_WNDPROC, MemOffset32(lpSubclass, SUBCLASS_pWndProcPrev)             '    this is the last one, and we may remove the subclass entirely.
                pWindowProp(hWnd) = 0&                                                                      '    remove the window property
            Else                                                                                            'if there is a previous or next data structure, then
                If lpSubclassPrev _
                    Then MemOffset32(lpSubclassPrev, SUBCLASS_pSubclassNext) = lpSubclassNext _
                Else pWindowProp(hWnd) = lpSubclassNext                                                 '    leave the subclass in place and adjust the linked list to remove this structure.
                End If
                MemFree lpSubclass                                                                              'free resources for this data structure
            End If
        End If
End Function

'Public Function Subclass_CallOldWindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
''---------------------------------------------------------------------------------------
'' Date      : 8/24/05
'' Purpose   : Call the original window procedure of a subclassed window.
''
'' Arguments : hWnd              window handle that is subclassed
''             uMsg              message to pass to the original procedure
''             wParam            message specific data
''             lParam            message specific data
''---------------------------------------------------------------------------------------
'    Dim lpSubclass As Long: lpSubclass = pWindowProp(hWnd)
'    If lpSubclass Then Subclass_CallOldWindowProc = CallWindowProc(MemOffset32(lpSubclass, SUBCLASS_pWndProcPrev), hWnd, uMsg, wParam, lParam)
'End Function

Private Function pAllocSubclass(ByVal lpOwner As Long, ByRef vMessagesBefore As Variant, ByRef vMessagesAfter As Variant, ByVal pSubClassNext As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 8/24/05
    ' Purpose   : Allocate a memory structure to hold data for an object that is requesting a subclass of a window.
    '---------------------------------------------------------------------------------------
    Dim liAfterOffset      As Long:  liAfterOffset = SUBCLASS_Len + pGetMsgTableLength(vMessagesBefore)          'Store the offset to the second message table in the structure
    pAllocSubclass = MemAlloc(liAfterOffset + pGetMsgTableLength(vMessagesAfter))                           'Allocate the structure
    If pAllocSubclass Then                                                                                  'proceed only if the structure was allocated
        MemOffset32(pAllocSubclass, SUBCLASS_pObject) = lpOwner                                             'Store the object pointer
        If pSubClassNext _
            Then MemOffset32(pAllocSubclass, SUBCLASS_pWndProcPrev) = MemOffset32(pSubClassNext, SUBCLASS_pWndProcPrev) _
        Else MemOffset32(pAllocSubclass, SUBCLASS_pWndProcPrev) = 0&                                    'Store the previous wndproc, if any
            MemOffset32(pAllocSubclass, SUBCLASS_pSubclassNext) = pSubClassNext                                 'Store the next node in the linked list
            MemOffset32(pAllocSubclass, SUBCLASS_pMsgTableB) = pGetMsgTable(pAllocSubclass, SUBCLASS_Len, vMessagesBefore) 'Store the message tables and their pointers
            MemOffset32(pAllocSubclass, SUBCLASS_pMsgTableA) = pGetMsgTable(pAllocSubclass, liAfterOffset, vMessagesAfter)
        End If
End Function

Private Function pGetMsgTable(ByVal pSubclass As Long, ByVal iOffset As Long, ByRef vMessages As Variant) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 8/24/05
    ' Purpose   : Return a pointer to a message table, 0 for no messages, or ALL_MESSAGES.
    '             Store the message values, if any, at the given offset of the data structure.
    '---------------------------------------------------------------------------------------
    Dim i      As Long
    If IsArray(vMessages) Then                                                                              'if the vMessages argument is an array
        pGetMsgTable = UnsignedAdd(pSubclass, iOffset)                                                      'get the offset the the message table
        i = UBound(vMessages) - LBound(vMessages)                                                           'get the length of the table (minus one)
        MemLong(pGetMsgTable) = i + 1&                                                                      'store the length of the table
        For i = 0& To i                                                                                     'for each element in the message table
            pGetMsgTable = UnsignedAdd(pGetMsgTable, 4&)                                                    'get a pointer to the current element
            MemLong(pGetMsgTable) = Val(vMessages(i))                                                       'store the current element
        Next
        pGetMsgTable = UnsignedAdd(pSubclass, iOffset)                                                      'return the offset to the message table
    ElseIf VarType(vMessages) = vbLong Then                                                                 'if the vMessages argument is a long value
        i = CLng(vMessages)                                                                                 'store the value
        If i = ALL_MESSAGES Then                                                                            'if it is ALL_MESSAGES
            pGetMsgTable = i                                                                                '    return ALL_MESSAGES
        Else                                                                                                'if it is not ALL_MESSAGES
            pGetMsgTable = UnsignedAdd(pSubclass, iOffset)                                                  '    return the offset to the message table
            MemLong(pGetMsgTable) = 1&                                                                      '    store the length of the table, one
            MemLong(UnsignedAdd(pGetMsgTable, 4&)) = i                                                      '    store the message value
        End If
    Else                                                                                                    'otherwise vMessages is unrecognized, hopefully Missing
        'debug.assert IsMissing(vMessages)
        pGetMsgTable = 0&
    End If
    
End Function

Private Function pGetMsgTableLength(ByRef vMessages As Variant) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 8/24/05
    ' Purpose   : Return the needed length to accomodate the given messages.
    '---------------------------------------------------------------------------------------
    If IsArray(vMessages) Then                                                                              'if the vMessages argument is an array
        pGetMsgTableLength = (UBound(vMessages) - LBound(vMessages) + 2&) * 4&                              '    allocate enough room to store the array plus four bytes for the message count.
    ElseIf VarType(vMessages) = vbLong Then                                                                 'if the vMessages argument is an long value
        pGetMsgTableLength = 8 * -CBool(CLng(vMessages) <> ALL_MESSAGES)                                    '    allocate enough room to store the value plus four bytes for the message count.
    Else                                                                                                    'otherwise vMessages is unrecognized, hopefully Missing
        'debug.assert IsMissing(vMessages)
        pGetMsgTableLength = 0&
    End If
End Function

Private Function pFindSubclass(ByVal pSubclass As Long, ByVal pObject As Long, Optional ByRef pSubclassPrev As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 8/24/05
    ' Purpose   : Search the linked list for a given object.
    '---------------------------------------------------------------------------------------
    pFindSubclass = pSubclass                                                                               'get the first subclass data structure
    pSubclassPrev = 0&                                                                                      'indicate that there is no previous data structure
    Do While pFindSubclass                                                                                  'loop through each data structure
        If MemOffset32(pFindSubclass, SUBCLASS_pObject) = pObject Then Exit Do                              'if the object matches, bail
            pSubclassPrev = pFindSubclass                                                                       'store this data structure as the previous
            pFindSubclass = MemOffset32(pFindSubclass, SUBCLASS_pSubclassNext)                                  'get a handle to the next data structure
        Loop
End Function

Private Property Get pWindowProp(ByVal hWnd As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 8/24/05
    ' Purpose   : Return a pointer to subclass data, if any.
    '---------------------------------------------------------------------------------------
    Dim lsAnsi      As String: lsAnsi = StrConv(sPropName & vbNullChar, vbFromUnicode)                           'store an ansi string
    pWindowProp = GetProp(hWnd, ByVal StrPtr(lsAnsi))                                                       'return the window property
End Property

Private Property Let pWindowProp(ByVal hWnd As Long, ByVal lpNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 8/24/05
    ' Purpose   : Set or remove the pointer to subclass data for the given window.
    '---------------------------------------------------------------------------------------
    Dim lsAnsi      As String: lsAnsi = StrConv(sPropName & vbNullChar, vbFromUnicode)                           'store an ansi string
    If lpNew _
        Then SetProp hWnd, ByVal StrPtr(lsAnsi), lpNew _
    Else RemoveProp hWnd, ByVal StrPtr(lsAnsi)                                                          'Set or remove the window property
End Property

#If bUseASM Then

Private Function pSubclassProc() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 8/24/05
    ' Purpose   : Allocate the linked list and patch the runtime values.
    '---------------------------------------------------------------------------------------
    Static sProp As String
    Static pAsm As Long
        
    Const PATCH_pPropString     As Long = 86
    Const PATCH_GetProp         As Long = 94
    Const PATCH_CallWindowProc  As Long = 156
    Const PATCH_SetWindowLong   As Long = 176
    Const PATCH_EbMode          As Long = 199
    Const PATCH_InIde           As Long = 196
        
    If pAsm = 0& Then                                                                                   'if not already initialized
        pAsm = Thunk_Alloc(tnkSubclassProc, PATCH_InIde)                                                '    allocate the assembly code.
        If pAsm Then                                                                                    '    if successfully allocated, patch the runtime values
            Thunk_PatchFuncAddr pAsm, PATCH_EbMode, "vba6.dll", "EbMode"                                '        asm will call EbMode to determine whether it is safe to make a callback
            Thunk_PatchFuncAddr pAsm, PATCH_CallWindowProc, "user32.dll", "CallWindowProcA"             '        asm will call CallWindowProc on the previous procedure if a message is not handled
            Thunk_PatchFuncAddr pAsm, PATCH_SetWindowLong, "user32.dll", "SetWindowLongA" '             '        asm will call SetWindowLong if the ide stops to unsubclass
            Thunk_PatchFuncAddr pAsm, PATCH_GetProp, "user32.dll", "GetPropA"                           '        asm will call GetProp to get a pointer to the first structure in the linked list
            sProp = StrConv(sPropName & vbNullChar, vbFromUnicode)                                      '        store an ANSI copy of the property identifier
            MemOffset32(pAsm, PATCH_pPropString) = StrPtr(sProp)                                        '        asm will pass this strptr to User32.GetProp
        End If
        #If bDebug Then
        DEBUG_Remove DEBUG_hMem, pAsm                                                               'this will never be deallocated but is not actually being leaked, so remove debugging
        #End If
    End If
    pSubclassProc = pAsm                                                                                'return the address of the allocated assembly procedure
        
    End Function

#Else

    Private Function pSubclassProc() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 8/24/05
    ' Purpose   : return a pointer to our callback function.
    '---------------------------------------------------------------------------------------
        pSubclassProc = pAddrFunc(AddressOf Subclass_Proc)
    End Function
    
    Private Function pAddrFunc(ByVal i As Long) As Long
        pAddrFunc = i
    End Function
    
    Private Function Subclass_Proc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 8/24/05
    ' Purpose   : Make the callbacks on the iSubclass interface.  This is the same methodology used
    '             by the ASM but as one might imagine, the ASM is much more efficient. (And IDE safe)
    '---------------------------------------------------------------------------------------
        Dim lpSubclass As Long: lpSubclass = pWindowProp(hWnd)                                              'get a pointer to the subclass data structure
        
        If lpSubclass Then                                                                                  'proceed only if the pointer was obtained
            Dim lpWndProcPrev As Long: lpWndProcPrev = MemOffset32(lpSubclass, SUBCLASS_pWndProcPrev)       'get a pointer to the previous window proc - it may not be accessible after the callbacks are made
            Dim lpSubclassNext As Long, lbHandled As Boolean
            Do While lpSubclass                                                                             'loop through each subclass data structure
                lpSubclassNext = MemOffset32(lpSubclass, SUBCLASS_pSubclassNext)                            'get a pointer to the next subclass data structure - it may not be accessible after the callback is made
                If pMsgInTable(MemOffset32(lpSubclass, SUBCLASS_pMsgTableB), uMsg) Then                     'if the message is requested
                    pObjectFromPtr(MemOffset32(lpSubclass, SUBCLASS_pObject)).Before _
                        lbHandled, Subclass_Proc, hWnd, uMsg, wParam, lParam                                '    make the before callback
                    If lbHandled Then Exit Do                                                               '    if handled, bail
                End If
                lpSubclass = lpSubclassNext                                                                 'start over with the next data structure
            Loop
            
            If Not lbHandled Then
                Subclass_Proc = CallWindowProc(lpWndProcPrev, hWnd, uMsg, wParam, lParam)                   'if not handled by one of the objects, call the original procedure
                lpSubclass = pWindowProp(hWnd)                                                              'try again to get a pointer to the subclass data structure - it may have been removed during one of the callbacks
                If lpSubclass Then                                                                          'proceed only if the pointer was obtained
                    Do While lpSubclass                                                                     'loop through each subclass data structure
                        lpSubclassNext = MemOffset32(lpSubclass, SUBCLASS_pSubclassNext)                    'get a pointer to the next subclass data structure - it may not be accessible after the callback is made
                        If pMsgInTable(MemOffset32(lpSubclass, SUBCLASS_pMsgTableA), uMsg) Then             'if the message is requested
                            pObjectFromPtr(MemOffset32(lpSubclass, SUBCLASS_pObject)).After _
                                Subclass_Proc, hWnd, uMsg, wParam, lParam                                   '    make the after callback
                        End If
                        lpSubclass = lpSubclassNext                                                         'start over with the next data structure
                    Loop
                End If
            End If
        End If
    End Function
    
    Private Function pObjectFromPtr(ByVal pObject As Long) As iSubclass
    '---------------------------------------------------------------------------------------
    ' Date      : 8/24/05
    ' Purpose   : Return a object reference from the interface pointer.
    '---------------------------------------------------------------------------------------
        Dim loSubclass As iSubclass
        MemLong(VarPtr(loSubclass)) = pObject                                                               'copy the object pointer
        Set pObjectFromPtr = loSubclass                                                                     'copy the iSubclass object and addref
        MemLong(VarPtr(loSubclass)) = 0&                                                                    'clear the uncounted reference
    End Function
    
    Private Function pMsgInTable(ByVal pMsgTable As Long, ByVal uMsg As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 8/24/05
    ' Purpose   : Return a value indicating whether callbacks should be made on the message.
    '---------------------------------------------------------------------------------------
        If pMsgTable = ALL_MESSAGES Then                                                                    'if ALL_MESSAGES, all messages are found
            pMsgInTable = True
        ElseIf pMsgTable Then                                                                               'if <> 0 then scan the message table for the given message.
            Dim i As Long
            For i = 1 To MemLong(pMsgTable)                                                                 '    loop through each element
                pMsgTable = UnsignedAdd(pMsgTable, 4&)                                                      '    get the pointer to the current element
                If uMsg = MemLong(pMsgTable) Then                                                           '    if the message matches
                    pMsgInTable = True                                                                      '        return true
                    Exit For                                                                                '        bail
                End If
            Next
        End If                                                                                              'otherwise, return false
    End Function
    
#End If
