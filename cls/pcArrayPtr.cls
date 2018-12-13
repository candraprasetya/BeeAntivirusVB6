VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcArrayPtr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'pcArrayPtr.bas                      1/19/05
'
'           PURPOSE:
'               Replace the SAFEARRAY Pointer of a VB array with our own,
'               forcing it to access the memory we point it to.
'
'           LINEAGE:
'               www.vbaccelerator.com
'
'==================================================================================================

Option Explicit

Private mtHeader     As SAFEARRAY1D
Private miOldHeader  As Long
Private miArrPtr     As Long

Public Sub SetArrayByte(ByRef yArray() As Byte)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Initialize the custom array pointer to a byte array.
    '---------------------------------------------------------------------------------------
    Destroy
    
    Const SAFEARRAY_AUTO = &H1
    Const SAFEARRAY_FIXEDSIZE = &H10
    
    With mtHeader
        .cDims = OneL                                           'initialize the structure
        .fFeatures = SAFEARRAY_AUTO Or SAFEARRAY_FIXEDSIZE
        .cbElements = OneL
        .cLocks = ZeroL
        '.DataPointer = ZeroL
        '.Elements = ZeroL
        .Bounds(0).lLbound = ZeroL
    End With
    
    miArrPtr = ArrPtr(yArray)
    miOldHeader = MemLong(ByVal miArrPtr)
    'debug.assert miOldHeader = ZeroL
    MemLong(ByVal miArrPtr) = VarPtr(mtHeader)
End Sub


Public Sub SetArrayLong(ByRef iArray() As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Initialize the custom array pointer to a 32 bit array.
    '---------------------------------------------------------------------------------------
    Destroy
    
    Const SAFEARRAY_AUTO = &H1
    Const SAFEARRAY_FIXEDSIZE = &H10
    
    With mtHeader
        .cDims = OneL                                           'initialize the structure
        .fFeatures = SAFEARRAY_AUTO Or SAFEARRAY_FIXEDSIZE
        .cbElements = 4&
        .cLocks = ZeroL
        '.DataPointer = ZeroL
        '.Elements = ZeroL
        .Bounds(0).lLbound = ZeroL
    End With
    
    miArrPtr = ArrPtr(iArray)
    miOldHeader = MemLong(ByVal miArrPtr)
    'debug.assert miOldHeader = ZeroL
    MemLong(ByVal miArrPtr) = VarPtr(mtHeader)
    
End Sub

Public Sub SetArrayInt(ByRef iArray() As Integer)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Initialize the custom array pointer to a 16 bit array.
    '---------------------------------------------------------------------------------------
    Destroy
    
    Const SAFEARRAY_AUTO = &H1
    Const SAFEARRAY_FIXEDSIZE = &H10
    
    With mtHeader
        .cDims = OneL                                           'initialize the structure
        .fFeatures = SAFEARRAY_AUTO Or SAFEARRAY_FIXEDSIZE
        .cbElements = 4&
        .cLocks = ZeroL
        '.DataPointer = ZeroL
        '.Elements = ZeroL
        .Bounds(0).lLbound = ZeroL
    End With
    
    miArrPtr = ArrPtr(iArray)
    miOldHeader = MemLong(ByVal miArrPtr)
    'debug.assert miOldHeader = ZeroL
    MemLong(ByVal miArrPtr) = VarPtr(mtHeader)
    
End Sub

Public Sub Destroy()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Destroy the custom array pointer.
    '---------------------------------------------------------------------------------------
    If CBool(miArrPtr) Then
        MemLong(ByVal miArrPtr) = miOldHeader
        miArrPtr = ZeroL
        miOldHeader = ZeroL
    End If
End Sub

Public Function PointToString(ByRef s As String) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Point the array to a string.
    '---------------------------------------------------------------------------------------
    mtHeader.pvData = StrPtr(s)
    PointToString = LenB(s)
    mtHeader.Bounds(0).cElements = PointToString
End Function

Public Sub PointToLong(ByRef i As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Point the array to a 32 bit value.
    '---------------------------------------------------------------------------------------
    mtHeader.pvData = VarPtr(i)
    mtHeader.Bounds(0).cElements = 4&
End Sub

Public Sub POINT(ByVal iPtr As Long, ByVal iLen As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Point the array to any data.
    '---------------------------------------------------------------------------------------
    mtHeader.pvData = iPtr
    mtHeader.Bounds(0).cElements = iLen
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub
