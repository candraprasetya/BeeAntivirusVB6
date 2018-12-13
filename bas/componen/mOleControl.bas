Attribute VB_Name = "mOleControl"
'==================================================================================================
'mOleControl.bas                                9/10/05
'
'           PURPOSE:
'               Provides VTable subclassing for the IOleControl interface.
'
'           LINEAGE:
'               Paul Wilde's vbACOM.dll from www.vbaccelerator.com
'
'==================================================================================================

Option Explicit

Private Enum eVTable
    
    ' Ignore item 1: QueryInterface
    ' Ignore item 2: AddRef
    ' Ignore item 3: Release
    vtblGetControlInfo = 4
    vtblOnMnemonic
    ' Ignore item 6: OnAmbientPropertyChange
    ' Ignore item 7: FreezeEvents
    vtblCount
    
End Enum

Private moSubclass  As pcSubclassVTable

Public Sub VTableSubclass_OleControl_Install(ByVal pObject As Interfaces.IOleControl)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Install vtable subclassing on the OleControl interface.
    '---------------------------------------------------------------------------------------
    If moSubclass Is Nothing Then Set moSubclass = New pcSubclassVTable
    moSubclass.Subclass ObjPtr(pObject), vtblCount, vtblGetControlInfo, _
    AddressOf IOleControl_GetControlInfo, _
    AddressOf IOleControl_OnMnemonic
End Sub
Public Sub VTableSubclass_OleControl_Remove()
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Remove vtable subclassing from the OleControl interface.
    '---------------------------------------------------------------------------------------
    If Not moSubclass Is Nothing Then moSubclass.UnSubclass
End Sub

Private Function IOleControl_OnMnemonic(ByVal oObject As Object, pMsg As Msg) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : New vtable method for IOleControl::OnMnemonic.
    '---------------------------------------------------------------------------------------

    If VarPtr(pMsg) = ZeroL Then
        IOleControl_OnMnemonic = E_POINTER
        Exit Function
    ElseIf Not TypeOf oObject Is iOleControlVB Then
        IOleControl_OnMnemonic = Original_IOleControl_OnMnemonic(oObject, pMsg)
        Exit Function
    End If
    
    Dim oIOleControlVB      As iOleControlVB
    Dim bHandled            As Boolean

    Set oIOleControlVB = oObject
    oIOleControlVB.OnMnemonic bHandled, KBState(), pMsg.Message, pMsg.wParam, pMsg.lParam
    
    If bHandled _
        Then IOleControl_OnMnemonic = S_OK _
    Else IOleControl_OnMnemonic = Original_IOleControl_OnMnemonic(oObject, pMsg)
    
End Function

Private Function IOleControl_GetControlInfo(ByVal oObject As Object, pCI As CONTROLINFO) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : New vtable method for IOleControl::GetControlInfo.
    '---------------------------------------------------------------------------------------
    
    If VarPtr(pCI) = ZeroL Then
        IOleControl_GetControlInfo = E_POINTER
        Exit Function
    ElseIf Not TypeOf oObject Is iOleControlVB Then
        IOleControl_GetControlInfo = Original_IOleControl_GetControlInfo(oObject, pCI)
        Exit Function
        
    End If
    
    Dim oIOleControlVB      As iOleControlVB
    Dim bHandled            As Boolean
    
    Set oIOleControlVB = oObject
    
    pCI.cb = LenB(pCI)
    oIOleControlVB.GetControlInfo bHandled, pCI.cAccel, pCI.hAccel, pCI.dwFlags
    
    If bHandled Then
        If CBool(pCI.cAccel) And pCI.hAccel = ZeroL _
            Then IOleControl_GetControlInfo = E_OUTOFMEMORY _
        Else IOleControl_GetControlInfo = S_OK
            
        Else
        
            IOleControl_GetControlInfo = Original_IOleControl_GetControlInfo(oObject, pCI)
        
        End If
    
End Function

Private Function Original_IOleControl_GetControlInfo(ByVal oObject As Interfaces.IOleControl, pCI As CONTROLINFO) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Call original IOleControl::GetControlInfo method.
    '---------------------------------------------------------------------------------------
    moSubclass.SubclassEntry(vtblGetControlInfo) = False
    Original_IOleControl_GetControlInfo = oObject.GetControlInfo(pCI)
    moSubclass.SubclassEntry(vtblGetControlInfo) = True
End Function
Private Function Original_IOleControl_OnMnemonic(ByVal oObject As Interfaces.IOleControl, pMsg As Msg) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Call original IOleControl::OnMnemonic method.
    '---------------------------------------------------------------------------------------
    moSubclass.SubclassEntry(vtblOnMnemonic) = False
    Original_IOleControl_OnMnemonic = oObject.OnMnemonic(ByVal VarPtr(pMsg))
    moSubclass.SubclassEntry(vtblOnMnemonic) = True
End Function

Public Function OnControlInfoChanged(ByVal oControl As Object) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : Notify the container that a control has changed its mnemonic information.
    '---------------------------------------------------------------------------------------
    Dim loOleObject        As IOleObject
    Dim loClientSite       As IOleClientSite
    Dim loUnknown          As stdole.IUnknown
    Dim loControlSite      As IOleControlSite
    
    On Error Resume Next
    
    Set loOleObject = oControl
    loOleObject.GetClientSite loClientSite
    Set loUnknown = loClientSite
    Set loControlSite = loUnknown
    'notify the control site that info has changed
    loControlSite.OnControlInfoChanged
    'force the control site to update default/cancel buttons
    loControlSite.OnFocus 1&
    OnControlInfoChanged = Not CBool(Err.Number)
    On Error GoTo 0
End Function
