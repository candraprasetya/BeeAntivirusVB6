Attribute VB_Name = "mOleInPlaceActiveObject"
'==================================================================================================
'mOleInPlaceActiveObject.bas            9/8/05
'
'           PURPOSE:
'               Subclassed implementation of IOleInPlaceActiveObject.
'
'           LINEAGE:
'               Based on modIOleInPlaceActiveObject.bas in vbACOM.dll from vbaccelerator.com
'
'==================================================================================================

Option Explicit

Private Enum eVTable
    ' Ignore item 1: QueryInterface
    ' Ignore item 2: AddRef
    ' Ignore item 3: Release
    ' Ignore item 4: GetWindow
    ' Ignore item 5: ContextSensitiveHelp
    vtblTranslateAccelerator = 6
    ' Ignore item 7: OnFrameWindowActivate
    ' Ignore item 8: OnDocWindowActivate
    ' Ignore item 9: ResizeBorder
    ' Ignore item 10: EnableModeless
    vtblCount
End Enum

Private moSubclass  As pcSubclassVTable

Private ActiveObject As iOleInPlaceActiveObjectVB

Public Sub VTableSubclass_IPAO_Install(ByVal oObject As Interfaces.IOleInPlaceActiveObject)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/08/05
    ' Purpose   : Install vtable subclassing on the IPAO interface.
    '---------------------------------------------------------------------------------------
    
    If moSubclass Is Nothing Then Set moSubclass = New pcSubclassVTable
    moSubclass.Subclass ObjPtr(oObject), vtblCount, vtblTranslateAccelerator, AddressOf IOleInPlaceActiveObject_TranslateAccelerator
    
End Sub

Public Sub VTableSubclass_IPAO_Remove()
    '---------------------------------------------------------------------------------------
    ' Date      : 9/08/05
    ' Purpose   : Remove vtable subclassing from the IPAO interface.
    '---------------------------------------------------------------------------------------
    If Not moSubclass Is Nothing Then moSubclass.UnSubclass
End Sub

Private Function IOleInPlaceActiveObject_TranslateAccelerator(ByVal oObject As Object, pMsg As MSG) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 9/08/05
    ' Purpose   : New vtable method for IOleInPlaceActiveObject::TranslateAccelerator.
    '---------------------------------------------------------------------------------------
    
    If VarPtr(pMsg) = ZeroL Then                                    'validate param
        IOleInPlaceActiveObject_TranslateAccelerator = E_POINTER
        
    ElseIf ActiveObject Is Nothing Then                             'if we've got nothing to do
        IOleInPlaceActiveObject_TranslateAccelerator = Original_IOleInPlaceActiveObject_TranslateAccelerator(oObject, pMsg)
        
    Else
        Dim lbHandled      As Boolean
        ActiveObject.TranslateAccelerator lbHandled, IOleInPlaceActiveObject_TranslateAccelerator, KBState(), pMsg.Message, pMsg.wParam, pMsg.lParam
        
        If Not lbHandled Then _
        IOleInPlaceActiveObject_TranslateAccelerator = Original_IOleInPlaceActiveObject_TranslateAccelerator(oObject, pMsg)
        'if control is not overriding default method, call original
        
    End If
    
End Function

Private Function Original_IOleInPlaceActiveObject_TranslateAccelerator(ByVal oThis As Interfaces.IOleInPlaceActiveObject, pMsg As MSG) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 9/08/05
    ' Purpose   : Call the original IOleInPlaceActiveObject method.
    '---------------------------------------------------------------------------------------
    On Error Resume Next
    
    moSubclass.SubclassEntry(vtblTranslateAccelerator) = False  'temporarily unhook method so we can call the original
    Original_IOleInPlaceActiveObject_TranslateAccelerator = _
    oThis.TranslateAccelerator(ByVal VarPtr(pMsg))  'call the original method
    moSubclass.SubclassEntry(vtblTranslateAccelerator) = True   're-hook the method
    
End Function


Public Sub ActivateIPAO(ByVal oObject As Object)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/2/05
    ' Purpose   : Ask the container to activate a given control.
    '---------------------------------------------------------------------------------------
    On Error Resume Next
      
    Dim pOleObject                   As Interfaces.IOleObject
    Dim pOleInPlaceSite              As Interfaces.IOleInPlaceSite
    Dim pOleInPlaceFrame             As Interfaces.IOleInPlaceFrame
    Dim pOleInPlaceUIWindow          As Interfaces.IOleInPlaceUIWindow
    Dim pOleInPlaceActiveObject      As Interfaces.IOleInPlaceActiveObject
    Dim PosRect                      As RECT
    Dim ClipRect                     As RECT
    Dim FrameInfo                    As OLEINPLACEFRAMEINFO
    
    Set pOleObject = oObject
    Set pOleInPlaceActiveObject = oObject
    
    pOleObject.GetClientSite pOleInPlaceSite
    pOleInPlaceSite.GetWindowContext pOleInPlaceFrame, pOleInPlaceUIWindow, PosRect, ClipRect, FrameInfo
    pOleInPlaceFrame.SetActiveObject pOleInPlaceActiveObject, vbNullString
    If Not pOleInPlaceUIWindow Is Nothing _
        Then pOleInPlaceUIWindow.SetActiveObject pOleInPlaceActiveObject, vbNullString
        
        ''debug.assert ActiveObject Is Nothing Or oObject Is ActiveObject
        If TypeOf oObject Is iOleInPlaceActiveObjectVB Then
            Set ActiveObject = oObject
            'Debug.Print "Activating: " & TypeName(ActiveObject)
        Else
            Set ActiveObject = Nothing
        End If
    
End Sub

Public Sub DeActivateIPAO(ByVal oObject As iOleInPlaceActiveObjectVB)
    'If Not ActiveObject Is Nothing Then Debug.Print "Deactivating: " & TypeName(ActiveObject)
    Set ActiveObject = Nothing
End Sub


