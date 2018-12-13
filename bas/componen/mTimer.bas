Attribute VB_Name = "mTimer"
'---------------------------------------------------------------------------------------
'mTimer.bas           8/24/05
'
'            PURPOSE:
'                IDE safe win32 timer callbacks on an object interface.
'
'            IMPLEMENTATION:
'                One large callback procedure is allocated for all timers and small procedures
'                are allocated for each timer.  The small procedures are stored in a linked list.
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

Const PATCH_InitialTick As Long = 4
Const PATCH_iOwnerID As Long = 12
Const PATCH_oOwner As Long = 20
Const PATCH_Callback As Long = 25
Const PATCH_iTimerID As Long = 31
Const PATCH_pClientTimerProcNext As Long = 35

Private mpLinkedList  As Long

Public Function Timer_Install(ByVal oOwner As iTimer, ByVal iId As Long, ByVal iInterval As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 8/24/05
    ' Purpose   : Install a timer and commence callbacks to the given object.
    '---------------------------------------------------------------------------------------
    
    'get a pointer to the iTimer interface
    Dim lpObject      As Long
    lpObject = ObjPtr(oOwner)
    
    'proceed only if the object is not Nothing
    If lpObject Then
        
        'proceed only if the object does not already have a timer with this id
        If pFindClient(mpLinkedList, lpObject, iId) = ZeroL Then
            
            'get a pointer to the main timer procedure
            Dim lpTimerProc      As Long
            lpTimerProc = pTimerProc
            
            'proceed only if the main timer procedure is valid
            If lpTimerProc Then
                
                'allocate a client timer procedure
                Dim lpClientTimerProc      As Long
                lpClientTimerProc = Thunk_Alloc(tnkTimerClientProc)
                
                'proceed only if the client timer prodedure is valid
                If lpClientTimerProc Then
                    
                    'set the timer
                    Dim liId      As Long
                    liId = SetTimer(ZeroL, ZeroL, iInterval, lpClientTimerProc)
                    
                    If liId Then
                        'if the timer was set, patch the runtime values and insert the
                        'procedure into the linked list
                        Timer_Install = True
                        MemOffset32(lpClientTimerProc, PATCH_InitialTick) = GetTickCount()
                        MemOffset32(lpClientTimerProc, PATCH_iOwnerID) = iId
                        MemOffset32(lpClientTimerProc, PATCH_oOwner) = lpObject
                        MemOffset32(lpClientTimerProc, PATCH_Callback) = lpTimerProc
                        MemOffset32(lpClientTimerProc, PATCH_iTimerID) = liId
                        MemOffset32(lpClientTimerProc, PATCH_pClientTimerProcNext) = mpLinkedList
                        mpLinkedList = lpClientTimerProc
                    Else
                        'if the timer was not set, free the procedure we were going to use
                        MemFree lpClientTimerProc
                    End If
                End If
            End If
        End If
    End If
End Function

Public Function Timer_Remove(ByVal oOwner As iTimer, ByVal iId As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 8/24/05
    ' Purpose   : Kill the specified timer and free the procedure allocated for it.
    '---------------------------------------------------------------------------------------
    
    'get a pointer to the iTimer interface
    Dim lpObject      As Long
    lpObject = ObjPtr(oOwner)
    
    'proceed only if the object is not Nothing
    If lpObject Then
        
        'find the client timer procedure for this client/id pair
        Dim lpClientTimerProc          As Long
        Dim lpClientTimerProcPrev      As Long
        lpClientTimerProc = pFindClient(mpLinkedList, lpObject, iId, lpClientTimerProcPrev)
        
        'proceed only if the procedure was found
        If lpClientTimerProc Then
            
            'fix up the linked list to remove this procedure
            If lpClientTimerProcPrev _
                Then MemOffset32(lpClientTimerProcPrev, PATCH_pClientTimerProcNext) = MemOffset32(lpClientTimerProc, PATCH_pClientTimerProcNext) _
            Else mpLinkedList = MemOffset32(lpClientTimerProc, PATCH_pClientTimerProcNext)
            
                'kill the timer to ensure no further callbacks
                KillTimer ZeroL, MemOffset32(lpClientTimerProc, PATCH_iTimerID)
            
                'free the timer procedure
                MemFree lpClientTimerProc
            
            End If
        End If
End Function

Private Function pFindClient(ByVal pLinkedList As Long, ByVal pObject As Long, ByVal iId As Long, Optional ByRef pClientTimerProcPrev As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 8/24/05
    ' Purpose   : Search the linked list for a given object/id pair.
    '---------------------------------------------------------------------------------------
    
    'indicate the the procedure is the first
    pClientTimerProcPrev = ZeroL
    'start with the first procedure in the list
    pFindClient = pLinkedList
    
    'loop through each procedure
    Do While pFindClient
        'if the object and id match, exit
        If MemOffset32(pLinkedList, PATCH_oOwner) = pObject Then
            If MemOffset32(pLinkedList, PATCH_iOwnerID) = iId Then Exit Do
        End If
        'store the previous procedure handle
        pClientTimerProcPrev = pFindClient
        'get the next procedure handle
        pFindClient = MemOffset32(pFindClient, PATCH_pClientTimerProcNext)
    Loop
    
End Function

#If bUseASM Then

Private Function pTimerProc() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 8/24/05
    ' Purpose   : Allocate a single timer procedure and return its handled for all subsequent calls.
    '---------------------------------------------------------------------------------------
    Static pAsm As Long
        
    If pAsm = ZeroL Then
        'if not already done, allocate the procedure and patch the runtime values.
        Const PATCH_InIde As Long = 0
        Const PATCH_EbMode As Long = 3
        Const PATCH_KillTimer As Long = 53
            
        pAsm = Thunk_Alloc(tnkTimerProc, PATCH_InIde)
            
        Thunk_PatchFuncAddr pAsm, PATCH_EbMode, "vba6.dll", "EbMode"
        Thunk_PatchFuncAddr pAsm, PATCH_KillTimer, "user32.dll", "KillTimer"
            
        #If bDebug Then
        DEBUG_Remove DEBUG_hMem, pAsm
        #End If
            
    End If
        
    pTimerProc = pAsm
        
    End Function

#Else
    
    Private Function pTimerProc() As Long
        pTimerProc = pAddrFunc(AddressOf Timer_Proc)
    End Function
    
    Private Function pAddrFunc(ByVal i As Long)
        pAddrFunc = i
    End Function
    
    Private Function Timer_Proc(ByVal oObject As iTimer, ByVal iId As Long, ByVal iTimerId As Long, ByVal iElapsed As Long) As Long
        oObject.Proc iElapsed, iId
    End Function
    
#End If
