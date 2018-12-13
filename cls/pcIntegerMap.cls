VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcIntegerMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'pcIntegerMap.cls                    12/15/04
'
'           PURPOSE:
'               Manage integer values for a keyed collection.  Uses C style null-terminated
'               strings and a hash table for quick compare.
'
'               The hash table node node and string pointers are maintained by the client, and
'               must not be released between adding to and removing from this collection.
'
'==================================================================================================
Option Explicit

Private Const HashTableCount As Long = 256
Private miHashTable(0 To HashTableCount - 1) As Long

Private Const NODE_iInt     As Long = 8&
Private Const NODE_NextPtr  As Long = 12&

Public Sub Add(ByVal pNode As Long, ByVal iHash As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Add an integer to the collection.
    '---------------------------------------------------------------------------------------
    'debug.assert pNode
    
    MemOffset32(pNode, NODE_NextPtr) = miHashTable(iHash)
    miHashTable(iHash) = pNode
    
End Sub

Public Sub Remove(ByVal pNode As Long, ByVal iHash As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Remove an integer from the collection.
    '---------------------------------------------------------------------------------------
    'debug.assert pNode
    
    Dim lpNode          As Long
    Dim lpNodePrev      As Long
    
    lpNode = miHashTable(iHash)
    
    Do While lpNode
        
        If lpNode = pNode Then
            'debug.assert CBool(lpNodePrev) Or miHashTable(iHash) = lpNode
            pNode = MemOffset32(lpNode, NODE_NextPtr)
            If lpNodePrev _
                Then MemOffset32(lpNodePrev, NODE_NextPtr) = pNode _
            Else miHashTable(iHash) = pNode
                Exit Do
            End If
        
            lpNodePrev = lpNode
            lpNode = MemOffset32(lpNode, NODE_NextPtr)
        
        Loop
    
        'debug.assert lpNode
    
End Sub

Public Function Find(ByVal iInt As Long, ByVal iHash As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Get the value associated with this integer.
    '---------------------------------------------------------------------------------------
    'debug.assert iInt
    
    Find = miHashTable(iHash)
    
    Do While Find
        If iInt = MemOffset32(Find, NODE_iInt) Then Exit Do
        Find = MemOffset32(Find, NODE_NextPtr)
    Loop
    
End Function
