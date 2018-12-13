Attribute VB_Name = "mPerPropertyBrowsing"
'==================================================================================================
'mPerPropertyBrowsing.bas               8/25/04
'
'           PURPOSE:
'               Subclassed implementation of IPerPropertyBrowsing
'
'           LINEAGE:
'               Based on modIPerPropertyBrowsing.bas in vbACOM.dll from vbaccelerator.com
'
'==================================================================================================

Option Explicit

Private moSubclass        As pcSubclassVTable
Private mbStringsNotImpl  As Boolean

Private Enum eVTable
    ' Ignore item 1: QueryInterface
    ' Ignore item 2: AddRef
    ' Ignore item 3: Release
    ' Ignore item 4: GetDisplayString
    ' Ignore item 5: MapPropertyToPage
    vtblGetPredefinedStrings = 6
    vtblGetPredefinedValue
    vtblCount
End Enum

Public Sub VTableSubclass_PPB_Install(ByVal pObject As Interfaces.IPerPropertyBrowsing)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/08/05
    ' Purpose   : Install vtable subclassing on the IPerPropertyBrowsing interface.
    '---------------------------------------------------------------------------------------
    'replace vtable for IPerPropertyBrowsing interface

    If moSubclass Is Nothing Then Set moSubclass = New pcSubclassVTable
    
    moSubclass.Subclass ObjPtr(pObject), vtblCount, vtblGetPredefinedStrings, _
    AddressOf IPerPropertyBrowsing_GetPredefinedStrings, _
    AddressOf IPerPropertyBrowsing_GetPredefinedValue

End Sub
Public Sub VTableSubclass_PPB_Remove()
    '---------------------------------------------------------------------------------------
    ' Date      : 9/08/05
    ' Purpose   : Remove vtable subclassing from the IPerPropertyBrowsing interface.
    '---------------------------------------------------------------------------------------
    If Not moSubclass Is Nothing Then moSubclass.UnSubclass
End Sub

Private Function IPerPropertyBrowsing_GetPredefinedStrings(ByVal oThis As Object, ByVal DispID As Long, pCaStringsOut As CALPOLESTR, pCaCookiesOut As CADWORD) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 9/08/05
    ' Purpose   : New vtable method for IPerPropertyBrowsing::GetPredefinedStrings.
    '---------------------------------------------------------------------------------------
    Dim oIPerPropertyBrowsingVB      As iPerPropertyBrowsingVB
    Dim bNoDefault                   As Boolean
    
    Dim cElems As Long
    Dim pElems As Long
    Dim lpString As Long
    
    'Debug.Print "GetPredefinedStrings"
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    mbStringsNotImpl = False
    
    'validate passed pointers
    If VarPtr(pCaStringsOut) = 0 Or VarPtr(pCaCookiesOut) = 0 Then
        IPerPropertyBrowsing_GetPredefinedStrings = E_POINTER
        Exit Function
        
    End If
    
    'create & initialise cPropertyListItems collection
    Dim loProps      As pcPropertyListItems
    Set loProps = New pcPropertyListItems
    
    'cast method to source interface
    Set oIPerPropertyBrowsingVB = oThis
    oIPerPropertyBrowsingVB.GetPredefinedStrings bNoDefault, DispID, loProps
    
    'if no param set by user
    If bNoDefault And loProps.Count > 0& Then
        'initialise CALPOLESTR struct
        cElems = loProps.Count
        pElems = CoTaskMemAlloc(cElems * 4)
        #If bDebug Then
        DEBUG_Remove DEBUG_hMemCoTask, pElems
        #End If
        
        pCaStringsOut.cElems = cElems
        pCaStringsOut.pElems = pElems
        
        Dim lsTemp      As String
        Dim loProp      As pcPropertyListItem
        
        For Each loProp In loProps
            lpString = loProp.lpDisplayName
            CopyMemory ByVal pElems, lpString, 4&
            'incr the element count
            pElems = UnsignedAdd(pElems, 4&)
        Next
        
        'initialise CADWORD struct
        pElems = CoTaskMemAlloc(cElems * 4)
        #If bDebug Then
        DEBUG_Remove DEBUG_hMemCoTask, pElems
        #End If
        pCaCookiesOut.cElems = cElems
        pCaCookiesOut.pElems = pElems
        
        'copy dwords to CADWORD struct
        For Each loProp In loProps
            CopyMemory ByVal pElems, loProp.Cookie, 4
            pElems = UnsignedAdd(pElems, 4&)
        Next
        
    Else

CATCH_EXCEPTION:
        
        IPerPropertyBrowsing_GetPredefinedStrings = E_NOTIMPL
        mbStringsNotImpl = True
    End If

    
End Function
Private Function IPerPropertyBrowsing_GetPredefinedValue(ByVal oThis As Object, ByVal DispID As Long, ByVal dwCookie As Long, pVarOut As Variant) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 9/08/05
    ' Purpose   : New vtable method for IPerPropertyBrowsing::GetPredefinedValue.
    '---------------------------------------------------------------------------------------
    
    If mbStringsNotImpl Then
        Debug.Print "Strings Not Implemented"
        IPerPropertyBrowsing_GetPredefinedValue = E_NOTIMPL
        Exit Function
    End If
    
    Dim oIPerPropertyBrowsingVB      As iPerPropertyBrowsingVB
    Dim bNoDefault                   As Boolean
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION

    'validate passed pointers
    If VarPtr(dwCookie) = 0 Or VarPtr(pVarOut) = 0 Then
        IPerPropertyBrowsing_GetPredefinedValue = E_POINTER
        Exit Function
        
    End If
    
    'cast method to source interface
    Set oIPerPropertyBrowsingVB = oThis
    oIPerPropertyBrowsingVB.GetPredefinedValue bNoDefault, DispID, dwCookie, pVarOut
    
    'if no param set by user
    If bNoDefault Then
        
        IPerPropertyBrowsing_GetPredefinedValue = S_OK
        
    Else
    
CATCH_EXCEPTION:
        
        IPerPropertyBrowsing_GetPredefinedValue = E_NOTIMPL
        
    End If
    
End Function

'Private Function Original_IPerPropertyBrowsing_GetPredefinedStrings(ByVal oThis As IPerPropertyBrowsing, ByVal DispID As Long, pCaStringsOut As CALPOLESTR, pCaCookiesOut As CADWORD) As Long
'    moSubclass.SubclassEntry(vtblGetPredefinedStrings) = False
'    Original_IPerPropertyBrowsing_GetPredefinedStrings = oThis.GetPredefinedStrings(DispID, pCaStringsOut, pCaCookiesOut)
'    moSubclass.SubclassEntry(vtblGetPredefinedStrings) = True
'End Function
'Private Function Original_IPerPropertyBrowsing_GetPredefinedValue(ByVal oThis As IPerPropertyBrowsing, ByVal DispID As Long, ByVal dwCookie As Long, pVarOut As Variant) As Long
'    moSubclass.SubclassEntry(vtblGetPredefinedValue) = False
'    Original_IPerPropertyBrowsing_GetPredefinedValue = oThis.GetPredefinedValue(DispID, dwCookie, pVarOut)
'    moSubclass.SubclassEntry(vtblGetPredefinedValue) = True
'End Function
