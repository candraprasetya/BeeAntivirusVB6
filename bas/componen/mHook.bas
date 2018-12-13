Attribute VB_Name = "mHook"
'---------------------------------------------------------------------------------------
'mHook.bas           8/24/05
'
'            PURPOSE:
'                IDE safe windows hooks that route callbacks to object methods.
'
'            IMPLEMENTATION:
'                One large callback function is allocated for all hooks, and small callback
'                functions are allocated for each hook that pushes hook-specific data on the
'                stack to pass to the large function which performs the work.  One linked list
'                is maintained for each hook type which stores the small hook functions.
'
'            LINEAGE:
'                http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=51403&lngWId=1
'
'---------------------------------------------------------------------------------------
Option Explicit

#Const bUseASM = True                   'If True, ASM is loaded from the resource file and provides fast
'callbacks and IDE protection.  If False, the callback function is
'a regular VB function with no IDE protection.  Best to set to True
'unless debugging the assembler.

Public Enum eHookType
    WH_MSGFILTER = -1       'local
    WH_KEYBOARD = 2         'local
    WH_GETMESSAGE = 3       'local
    WH_CALLWNDPROC = 4      'local
    WH_CBT = 5              'local
    WH_MOUSE = 7            'local
    WH_DEBUG = 9            'local
    WH_SHELL = 10           'local
    WH_FOREGROUNDIDLE = 11  'local
    WH_CALLWNDPROCRET = 12  'local
    WH_KEYBOARD_LL = 13     'global - NT/XP only
    WH_MOUSE_LL = 14        'global - NT/XP only
End Enum

Public Enum eHookCode
    HC_ACTION = 0
    HC_GETNEXT = 1
    HC_SKIP = 2
    HC_NOREMOVE = 3
    HCBT_MOVESIZE = 0
    HCBT_MINMAX = 1
    HCBT_QS = 2
    HCBT_CREATEWND = 3
    HCBT_DESTROYWND = 4
    HCBT_ACTIVATE = 5
    HCBT_CLICKSKIPPED = 6
    HCBT_KEYSKIPPED = 7
    HCBT_SYSCOMMAND = 8
    HCBT_SETFOCUS = 9
    HSHELL_WINDOWCREATED = 1
    HSHELL_WINDOWDESTROYED = 2
    HSHELL_ACTIVATESHELLWINDOW = 3
    HSHELL_WINDOWACTIVATED = 4
    HSHELL_GETMINRECT = 5
    HSHELL_REDRAW = 6
    HSHELL_TASKMAN = 7
    HSHELL_LANGUAGE = 8
    MSGF_DIALOGBOX = 0
    MSGF_MESSAGEBOX = 1
    MSGF_MENU = 2
    MSGF_SCROLLBAR = 5
    MSGF_NEXTWINDOW = 6
    MSGF_MAX = 8
    MSGF_USER = 4096
    MSGF_DDEMGR = 32769
End Enum

#If False Then
'Preserve case in the ide
Private WH_MSGFILTER, WH_KEYBOARD, WH_GETMESSAGE, WH_CALLWNDPROC, WH_CBT, WH_MOUSE, WH_DEBUG, WH_SHELL, WH_FOREGROUNDIDLE, WH_CALLWNDPROCRET, WH_KEYBOARD_LL, WH_MOUSE_LL, HCBT_MOVESIZE, HCBT_MINMAX, HCBT_QS, HCBT_CREATEWND, HCBT_DESTROYWND, HCBT_ACTIVATE, HCBT_CLICKSKIPPED, HCBT_KEYSKIPPED, HCBT_SYSCOMMAND, HCBT_SETFOCUS, PM_NOREMOVE, PM_REMOVE, PM_NOYIELD, HC_ACTION, HC_GETNEXT, HC_SKIP, HC_NOREMOVE, HC_NOREM, HC_SYSMODALON, HC_SYSMODALOFF, MSGF_DIALOGBOX, MSGF_MESSAGEBOX, MSGF_MENU, MSGF_SCROLLBAR, MSGF_NEXTWINDOW, MSGF_MAX, MSGF_USER, MSGF_DDEMGR, HSHELL_WINDOWCREATED, HSHELL_WINDOWDESTROYED, HSHELL_ACTIVATESHELLWINDOW, HSHELL_WINDOWACTIVATED, HSHELL_GETMINRECT, HSHELL_REDRAW, HSHELL_TASKMAN, HSHELL_LANGUAGE
#End If

Private Const PATCH_iHookType            As Long = 2
Private Const PATCH_hHook                As Long = 7
Private Const PATCH_oOwner               As Long = 12
Private Const PATCH_Callback             As Long = 18
Private Const PATCH_pClientHookProcNext As Long = 24

Private mpHooks(0 To 11)                As Long         'stores the first node in a linked list for each hook type

Public Function Hook_Install(ByVal oOwner As iHook, ByVal iType As eHookType) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 8/24/05
    ' Purpose   : Install a hook and set up a callback function to the given object.
    '
    ' Arguments : oOwner            iHook interface on which callbacks are to be made
    '             iType             type of hook to be installed
    '---------------------------------------------------------------------------------------
    
    'get an array index from the hook type
    Dim liArrayIndex      As Long
    liArrayIndex = pArrayIndex(iType)
    
    'proceed only if the hook type was valid
    If liArrayIndex > -1& Then
        
        'get a pointer to the iHook Interface
        Dim lpObject      As Long
        lpObject = ObjPtr(oOwner)
        
        'proceed only if the object is not Nothing
        If lpObject Then
            
            'proceed only if this client does not already receive callbacks on this type of hook
            If pFindClient(lpObject, liArrayIndex) = ZeroL Then
            
                'get a pointer to the main hook procedure
                Dim lpHookProc      As Long
                lpHookProc = pHookProc()
                
                'proceed only if the main hook procedure is present
                If lpHookProc Then
                    
                    'allocate a new client hook procedure
                    Dim lpHook      As Long
                    lpHook = Thunk_Alloc(tnkHookClientProc)
                    
                    'proceed only if the procedure was allocated
                    If lpHook Then
                        
                        'install the requested hook
                        Dim lhHook      As Long
                        Select Case iType
                        Case WH_KEYBOARD_LL, WH_MOUSE_LL
                            lhHook = SetWindowsHookEx(iType, lpHook, App.hInstance, ZeroL)
                        Case Else
                            lhHook = SetWindowsHookEx(iType, lpHook, ZeroL, App.ThreadID)
                        End Select
                        
                        Hook_Install = CBool(lhHook)
                        
                    End If
                End If
            End If
        End If
    End If
    
    If Hook_Install Then
        'if the hook installation was successful, patch the runtime values and store it in the linked list
        MemOffset32(lpHook, PATCH_iHookType) = iType
        MemOffset32(lpHook, PATCH_hHook) = lhHook
        MemOffset32(lpHook, PATCH_oOwner) = lpObject
        MemOffset32(lpHook, PATCH_Callback) = lpHookProc
        MemOffset32(lpHook, PATCH_pClientHookProcNext) = mpHooks(liArrayIndex)
        mpHooks(liArrayIndex) = lpHook
    Else
        'if the hook installation was not successful, free the client procedure we were going to use
        If lhHook Then MemFree lhHook
    End If
    
End Function

Public Function Hook_Remove(ByVal oOwner As iHook, ByVal iType As eHookType) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 8/24/05
    ' Purpose   : Remove a hook and its associated callback function.
    '
    ' Arguments : oOwner            iHook interface on which to cease callbacks
    '             iType             type of hook to be removed
    '---------------------------------------------------------------------------------------
    
    'get a pointer to the iHook interface
    Dim lpObject      As Long
    lpObject = ObjPtr(oOwner)
    
    'proceed only if the object is not Nothing
    If lpObject Then
    
        'get an array index from the hook type
        Dim liArrayIndex      As Long
        liArrayIndex = pArrayIndex(iType)
        
        'proceed only if the hook type was valid
        If liArrayIndex > -1& Then
            
            'get a pointer to the client hook procedure
            Dim lpHook          As Long
            Dim lpHookPrev      As Long
            lpHook = pFindClient(lpObject, liArrayIndex, lpHookPrev)
            
            'proceed only if the client hook procedure was found
            If lpHook Then
                
                Hook_Remove = True
                'remove the hook
                UnHookWindowsHookEx MemOffset32(lpHook, PATCH_hHook)
                
                'adjust the linked list to remove this node
                If lpHookPrev _
                    Then MemOffset32(lpHookPrev, PATCH_pClientHookProcNext) = MemOffset32(lpHook, PATCH_pClientHookProcNext) _
                Else mpHooks(liArrayIndex) = MemOffset32(lpHook, PATCH_pClientHookProcNext)
                
                    'free the client hook procedure
                    MemFree lpHook
                End If
            End If
        End If
End Function

Public Function Hook_CallNext(ByVal oOwner As iHook, ByVal iType As eHookType, ByVal nCode As eHookCode, ByVal wParam As Long, ByVal lParam As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 8/24/05
    ' Purpose   : Delegate the the CallNextHookEx function.
    '---------------------------------------------------------------------------------------
    
    'get a pointer to the iHook interface
    Dim lpObject      As Long
    lpObject = ObjPtr(oOwner)
    
    'proceed only if the object is not Nothing
    If lpObject Then
        
        'get an array index from the hook type
        Dim liArrayIndex      As Long
        liArrayIndex = pArrayIndex(iType)
        
        'proceed only if the hook type was valid
        If liArrayIndex > -1& Then
            
            'find the client hook procedure
            Dim lpHook      As Long
            lpHook = pFindClient(lpObject, liArrayIndex)
            
            'proceed only if the client hook procedure was found
            If lpHook Then
                'call the next hook and return its value
                Hook_CallNext = CallNextHookEx(MemOffset32(lpHook, PATCH_hHook), nCode, wParam, lParam)
            End If
        End If
    End If
    
End Function

Private Function pFindClient(ByVal lpObject As Long, ByVal iArrayIndex As Long, Optional ByRef pClientHookProcPrev As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 8/24/05
    ' Purpose   : Search a linked list for the given client.
    '---------------------------------------------------------------------------------------
    
    'indicate no previous procedure
    pClientHookProcPrev = ZeroL
    'start with the first procedure in the list
    pFindClient = mpHooks(iArrayIndex)
    
    'loop through each procedure
    Do While pFindClient
        'if the client matches, exit
        If MemOffset32(pFindClient, PATCH_oOwner) = lpObject Then Exit Do
        'store the previous procedure
        pClientHookProcPrev = pFindClient
        'store the next procedure
        pFindClient = MemOffset32(pFindClient, PATCH_pClientHookProcNext)
    Loop
    
End Function

Private Function pArrayIndex(ByVal iHookType As eHookType) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 8/24/05
    ' Purpose   : Adjust the hook values to get an contiguous array index.
    '---------------------------------------------------------------------------------------
    Select Case iHookType
    Case WH_MSGFILTER
        pArrayIndex = ZeroL
    Case WH_KEYBOARD, WH_GETMESSAGE, WH_CALLWNDPROC, WH_CBT
        pArrayIndex = iHookType - 1&
    Case WH_MOUSE
        pArrayIndex = iHookType - 2&
    Case WH_DEBUG, WH_SHELL, WH_FOREGROUNDIDLE, WH_CALLWNDPROCRET, WH_KEYBOARD_LL, WH_MOUSE_LL
        pArrayIndex = iHookType - 3&
    Case Else
        'debug.assert False
        pArrayIndex = -1&
    End Select
End Function


#If bUseASM Then

Private Function pHookProc() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 8/24/05
    ' Purpose   : Allocate the main hook procedure and patch the runtime values.
    '---------------------------------------------------------------------------------------
    Static pAsm As Long
    If pAsm = ZeroL Then
            
        Const PATCH_InIde As Long = 8
        Const PATCH_EbMode As Long = 11
        Const PATCH_CallNextHookEx As Long = 135
        Const PATCH_UnhookWindowsHookEx As Long = 147
            
        pAsm = Thunk_Alloc(tnkHookProc, PATCH_InIde)
            
        Thunk_PatchFuncAddr pAsm, PATCH_EbMode, "vba6.dll", "EbMode"
        Thunk_PatchFuncAddr pAsm, PATCH_CallNextHookEx, "user32.dll", "CallNextHookEx"
        Thunk_PatchFuncAddr pAsm, PATCH_UnhookWindowsHookEx, "user32.dll", "UnhookWindowsHookEx"
            
        #If bDebug Then
        DEBUG_Remove DEBUG_hMem, pAsm
        #End If
            
    End If
        
    pHookProc = pAsm
    End Function

#Else

    Private Function pHookProc() As Long
        pHookProc = pAddrFunc(AddressOf Hook_Proc)
    End Function

    Private Function pAddrFunc(ByVal i As Long) As Long
        pAddrFunc = i
    End Function
    
    Private Function Hook_Proc(ByVal oObject As iHook, ByVal hHook As Long, ByVal iType As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 8/24/05
    ' Purpose   : Make the callbacks on the iHook interface.
    '---------------------------------------------------------------------------------------
        Dim lbHandled As Boolean
        oObject.Before lbHandled, Hook_Proc, iType, nCode, wParam, lParam
        If Not lbHandled Then
            Hook_Proc = CallNextHookEx(hHook, nCode, wParam, lParam)
            oObject.After Hook_Proc, iType, nCode, wParam, lParam
        End If
    End Function
    
#End If
