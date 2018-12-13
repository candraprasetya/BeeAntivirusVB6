VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcSubclassVTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'pcSubclassVTable.cls                   9/9/05
'
'           PURPOSE:
'               Maintain pointers to old and new vtable functions, subclassing and unsubclassing each
'               as requested.
'
'           LINEAGE:
'               Based on vbACOM.dll from vbaccelerator.com written by Paul Wilde, which was in turn
'               based on code from Matt Curland's book about something or other.
'
'==================================================================================================

Option Explicit

Private miVTable()     As Long         'Pointers to new VTable entries
Private miOldVTable()  As Long         'Pointers to original VTable entries

Private miPtrVTable     As Long         'pointer to the head of the VTable
Private miRefCount      As Long         'count the number of objects requiring subclass services
Private miVTableCount   As Long         'number of entries in the VTable

Public Sub Subclass(ByVal iObjPtr As Long, ByVal iVTableCount As Long, ByVal iFirstEntry As Long, ParamArray iNewEntries() As Variant)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/09/05
    ' Purpose   : On the first call, subclass the vtable and store the old pointers.
    '             Thereon, ignore.
    '---------------------------------------------------------------------------------------
    
    If miRefCount = ZeroL Then
        
        iFirstEntry = iFirstEntry - OneL      'adjust from 1-based to 0-based index
        
        'If you get this assertion, you're either passing some wacky parameters
        'or you are already subclassed!
        'debug.assert Not ( _
        iFirstEntry < ZeroL Or _
        iFirstEntry >= iVTableCount Or _
        iVTableCount < ZeroL Or _
        miPtrVTable <> ZeroL Or _
        iObjPtr = ZeroL)
        
        miPtrVTable = MemLong(ByVal iObjPtr)
        
        miVTableCount = iVTableCount                    'store the VTable count
        
        ReDim miVTable(0 To miVTableCount - 1&)         'redim the arrays
        ReDim miOldVTable(0 To miVTableCount - 1&)
        
        Dim liEntry         As Long
        Dim liPtrEntry      As Long
        
        liEntry = UBound(iNewEntries) + iFirstEntry                             'get the number of entries that we will subclass
        If liEntry >= miVTableCount Then liEntry = miVTableCount - 1&           'make sure it's not bigger than the count
            liPtrEntry = UnsignedAdd(miPtrVTable, iFirstEntry * 4&)                 'get the pointer to the first entry we care about
        
            For liEntry = iFirstEntry To liEntry
                miVTable(liEntry) = iNewEntries(liEntry - iFirstEntry)              'store the new vtable address
                If miVTable(liEntry) <> ZeroL Then                                     'if we are subclassing this entry
                    pSubclass liPtrEntry, miVTable(liEntry), miOldVTable(liEntry)   'set the subclass
                End If
                liPtrEntry = UnsignedAdd(liPtrEntry, 4&)                            'increment the VTable entry pointer
            Next
        
        End If
    
        miRefCount = miRefCount + OneL
    
End Sub

Public Property Let SubclassEntry(ByVal iEntry As Long, ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/09/05
    ' Purpose   : Subclass/Unsubclass specific functions in the vtable.
    '---------------------------------------------------------------------------------------
    
    iEntry = iEntry - 1&                                                    'change from 1-based to 0-based
    
    
    'if you get this assertion, you're either passing a bad parameter or have not yet set the subclass
    'debug.assert iEntry > -1& And iEntry < miVTableCount And miPtrVTable <> ZeroL
    'if you get this assertion, you are providing a function that wasn't subclassed in the first place
    'debug.assert miVTable(iEntry)
    'if you get this assertion, the subclass state is the same as you are requesting
    'did you hit the end button?  better restart the ide....
    'debug.assert SubclassEntry(iEntry + 1&) Xor bNew
    
    If CBool(miVTable(iEntry)) And CBool(SubclassEntry(iEntry + 1&) Xor bNew) Then
        
        Dim liPtrEntry      As Long
        liPtrEntry = UnsignedAdd(miPtrVTable, iEntry * 4&)                  'Calculate pointer to specified entry
        
        If bNew Then
            pSubclass liPtrEntry, miVTable(iEntry), miOldVTable(iEntry)     'set the new entry
            
        Else
            pSubclass liPtrEntry, miOldVTable(iEntry), ZeroL                'restore the old entry
            miOldVTable(iEntry) = ZeroL                                     'release the pointer to the old entry
            
        End If
        
    End If
    
End Property

Public Property Get SubclassEntry(ByVal iEntry As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 9/09/05
    ' Purpose   : Return a value indicating whether the given function is subclassed.
    '---------------------------------------------------------------------------------------
    
    iEntry = iEntry - 1&        '1-based to 0-based
    
    'bad parameter
    'debug.assert iEntry > -1& And iEntry < miVTableCount And miPtrVTable <> ZeroL
    SubclassEntry = CBool(miOldVTable(iEntry))  'we are subclassed if we have a pointer to the old entry
    
End Property

Public Sub UnSubclass()
    '---------------------------------------------------------------------------------------
    ' Date      : 9/09/05
    ' Purpose   : Decrement the subclass reference count. When 0, unsubclass.
    '---------------------------------------------------------------------------------------
    
    'you can't unsubclass if you're not subclassed!
    'debug.assert miPtrVTable
    
    miRefCount = miRefCount - OneL
   
    If miRefCount = ZeroL Then
        
        If miPtrVTable Then                             'if we are subclassed
            Dim i      As Long
            
            For i = ZeroL To miVTableCount - OneL          'loop through each entry
                If miOldVTable(i) Then                  'If we have an old entry
                    pSubclass UnsignedAdd(miPtrVTable, i * 4&), miOldVTable(i), ZeroL  'restore the old entry
                    miOldVTable(i) = ZeroL              'clear the pointer
                End If
            Next
            
            miPtrVTable = ZeroL                         'clear the pointer to the VTable
            
        End If
        
    End If
    
End Sub

Private Sub pSubclass(ByVal iPtrEntry As Long, ByVal iPtrNew As Long, ByRef iPtrOld As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/09/05
    ' Purpose   : Copy the pointer to the new function into the vtable and store the
    '             pointer to the old function.
    '---------------------------------------------------------------------------------------
    
    iPtrOld = MemLong(ByVal iPtrEntry)                          'Save address of previous method for return
    'did you hit the end button?  better restart the ide....
    'debug.assert iPtrOld <> iPtrNew
    VirtualProtect iPtrEntry, 4&, PAGE_EXECUTE_READWRITE, ZeroL    'Need to change page protection to write to code
    MemLong(ByVal iPtrEntry) = iPtrNew                          ' Write the new function address into the vtable
   
End Sub

Private Sub Class_Terminate()
    '---------------------------------------------------------------------------------------
    ' Date      : 9/09/05
    ' Purpose   : You'd better call unsubclass for every subclass ..... or else!
    '---------------------------------------------------------------------------------------
    'debug.assert miPtrVTable = ZeroL
End Sub
