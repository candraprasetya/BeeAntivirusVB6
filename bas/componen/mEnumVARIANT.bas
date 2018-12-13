Attribute VB_Name = "mEnumVariant"
'==================================================================================================
'mEnumVARIANT.bas                                8/25/04
'
'           LINEAGE:
'               Paul Wilde's vbACOM.dll from www.vbaccelerator.com
'
'           PURPOSE:
'               Provides VTable subclassing for the IEnumVARIANT interface.
'
'==================================================================================================

Option Explicit

Private mtSAHeader  As SAFEARRAY1D
Private mvArray()   As Variant 'never dimensioned, accesses memory already allocated

Private moSubclass      As pcSubclassVTable

Private Enum eVTable
    '      Ignore item 1: QueryInterface
    '      Ignore item 2: AddRef
    '      Ignore item 3: Release
    vtblNext = 4
    vtblSkip
    vtblReset
    vtblClone
    vtblCount
End Enum

Public Sub ReplaceIEnumVARIANT(ByVal oObject As Interfaces.IEnumVARIANT)
    '---------------------------------------------------------------------------------------
    ' Date      : 8/25/04
    ' Purpose   : Replace vtable methods for the IEnumVARIANT interface.
    '---------------------------------------------------------------------------------------
    If moSubclass Is Nothing Then Set moSubclass = New pcSubclassVTable
    
    moSubclass.Subclass ObjPtr(oObject), vtblCount, vtblNext, _
    AddressOf IEnumVARIANT_Next, _
    AddressOf IEnumVARIANT_Skip, _
    AddressOf IEnumVARIANT_Reset, _
    AddressOf IEnumVARIANT_Clone
    
End Sub
Public Sub RestoreIEnumVARIANT(ByVal oObject As Interfaces.IEnumVARIANT)
    '---------------------------------------------------------------------------------------
    ' Date      : 8/25/04
    ' Purpose   : Restore vtable methods for the IEnumVARIANT interface.
    '---------------------------------------------------------------------------------------
    If Not moSubclass Is Nothing Then moSubclass.UnSubclass

End Sub
Private Function IEnumVARIANT_Next(ByVal oThis As Object, ByVal lngVntCount As Long, vntArray As Variant, ByVal pcvFetched As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 8/25/04
    ' Purpose   : New vtable method for IEnumVARIANT::Next.
    '---------------------------------------------------------------------------------------

    On Error GoTo CATCH_EXCEPTION

    Dim oEnumVARIANT      As pcEnumeration
    Dim liFetched         As Long, lbNoMore As Boolean
    Dim i                 As Integer
    
    pInitArray VarPtr(vntArray), lngVntCount
    
    'cast method to source interface
    Set oEnumVARIANT = oThis
    
    'loop through each requested variant
    For i = 0 To lngVntCount - 1&
        'call the class method
        oEnumVARIANT.GetNextItem mvArray(i), lbNoMore
        
        'if nothing fetched, we're done
        If lbNoMore Then Exit For
        
        ' Count the item fetched
        liFetched = liFetched + 1&
    Next
    
    'Return success if we got all requested items
    If liFetched = lngVntCount Then
        IEnumVARIANT_Next = S_OK
        
    Else
        IEnumVARIANT_Next = S_FALSE
        
    End If
        
    'copy the actual number fetched to the pointer to fetched count
    If pcvFetched Then
        MemLong(ByVal pcvFetched) = liFetched
    End If
    
    pInitArray ZeroL, ZeroL
    
    Exit Function
    
CATCH_EXCEPTION:
        
    'convert error to COM format
    IEnumVARIANT_Next = MapCOMErr(Err.Number)
    
    'iterate back, emptying the invalid fetched variants
    For i = i To 0& Step -1&
        mvArray(i) = Empty
    Next

    'return 0 as the number fetched after error
    If pcvFetched Then
        MemLong(ByVal pcvFetched) = 0&
    End If
    
End Function
Private Function IEnumVARIANT_Skip(ByVal oThis As Object, ByVal cV As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 8/25/04
    ' Purpose   : New vtable method for IEnumVARIANT::Skip.
    '---------------------------------------------------------------------------------------
    Dim oEnumVARIANT      As pcEnumeration
    Dim bSkippedAll       As Boolean
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'cast method to source interface
    Set oEnumVARIANT = oThis
    
    'call the class method
    oEnumVARIANT.skip cV, bSkippedAll
   
    If bSkippedAll _
        Then IEnumVARIANT_Skip = S_OK _
    Else IEnumVARIANT_Skip = S_FALSE
    
        Exit Function
    
CATCH_EXCEPTION:
        Debug.Print "Error!"
        IEnumVARIANT_Skip = MapCOMErr(Err.Number)
    
End Function

Private Function IEnumVARIANT_Reset(ByVal oThis As Object) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 8/25/04
    ' Purpose   : New vtable method for IEnumVARIANT::Reset.
    '---------------------------------------------------------------------------------------
    Dim oEnumVARIANT      As pcEnumeration
    
    On Error GoTo CATCH_EXCEPTION
    
    Set oEnumVARIANT = oThis
    oEnumVARIANT.Reset
    IEnumVARIANT_Reset = S_OK
    
    Exit Function
    
CATCH_EXCEPTION:
    
    IEnumVARIANT_Reset = MapCOMErr(Err.Number)
        
End Function

Private Function IEnumVARIANT_Clone(ByVal oThis As Object, ByRef ppEnum As Interfaces.IEnumVARIANT) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 8/25/04
    ' Purpose   : Forget it.
    '---------------------------------------------------------------------------------------
    IEnumVARIANT_Clone = E_NOTIMPL
End Function


Private Sub pInitArray(ByVal iAddr As Long, icEl As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 8/25/04
    ' Purpose   : Point the modular array to the given elements.
    '---------------------------------------------------------------------------------------
    Const FADF_STATIC = &H2&      '// Array is statically allocated.
    Const FADF_FIXEDSIZE = &H10&  '// Array may not be resized or reallocated.
    Const FADF_VARIANT = &H800&   '// An array of VARIANTs.
    
    Const FADF_Flags = FADF_STATIC Or FADF_FIXEDSIZE Or FADF_VARIANT
    
    With mtSAHeader
        If .cDims = 0& Then
            .cbElements = 16
            .cDims = 1
            .fFeatures = FADF_Flags
            CopyMemory ByVal ArrPtr(mvArray), VarPtr(mtSAHeader), 4&
        End If
        .Bounds(0).cElements = icEl + 1&
        .pvData = iAddr
    End With

End Sub

Private Function MapCOMErr(ByVal ErrNumber As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 8/25/04
    ' Purpose   : Map vb error to COM error.
    '---------------------------------------------------------------------------------------
    If ErrNumber <> 0& Then
        If (ErrNumber And &H80000000) Or (ErrNumber = 1&) Then
            'Error HRESULT already set
            MapCOMErr = ErrNumber
            
        Else
            'Map back to a basic error number
            MapCOMErr = &H800A0000 Or ErrNumber
            
        End If
        
    End If
End Function
