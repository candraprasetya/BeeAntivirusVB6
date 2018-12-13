VERSION 5.00
Begin VB.UserControl ucUpDown 
   CanGetFocus     =   0   'False
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   300
   HasDC           =   0   'False
   ScaleHeight     =   19
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   20
   ToolboxBitmap   =   "ucUpDown.ctx":0000
End
Attribute VB_Name = "ucUpDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'ucUpDown.ctl        8/31/04
'
'           PURPOSE:
'               Implement the win32 updown control.
'
'           LINEAGE:
'               http://www.vbaccelerator.com/codelib/comctl/updown.htm
'               cUpDown6.ctl
'
'==================================================================================================

Option Explicit

Public Enum eUpDownOrientation
    udVertical = 0
    udHorizontal = UDS_HORZ
End Enum

Private Const UM_SyncBuddy As Long = WM_USER + 40  ' + &H66BC&

Public Event BeforeChange(ByVal iValue As Long, ByRef iDelta As Long)
Public Event Change(ByVal iValue As Long)

Implements iSubclass
Implements iPerPropertyBrowsingVB

Private Enum eBooleanProps
    'bpAlignLeft = UDS_ALIGNLEFT
    'bpAlignRight = UDS_ALIGNRIGHT
    bpHorizontal = UDS_HORZ
    
    bpNoThousands = UDS_NOTHOUSANDS
    'bpChangeBuddyText = UDS_SETBUDDYINT
    'bpArrowKeysChange = UDS_ARROWKEYS
    bpWrap = UDS_WRAP
    
    bpHexadecimal = &H20000
    bpVisible = &H40000
    
    bpValidUDStyles = bpHorizontal Or bpNoThousands Or bpWrap
End Enum

Private Const PROP_BooleanProps      As String = "BProps"
Private Const PROP_BuddyControl      As String = "Buddy"
Private Const PROP_Max               As String = "Max"
Private Const PROP_Min               As String = "Min"
Private Const PROP_Value             As String = "Value"
Private Const PROP_LargeChange       As String = "Large"
Private Const PROP_SmallChange       As String = "Small"
Private Const PROP_Delay             As String = "Delay"
Private Const PROP_Enabled           As String = "Enabled"
Private Const PROP_Themeable         As String = "Themeable"
Private Const PROP_BuddyAlignment    As String = "BuddyAlign"
Private Const PROP_BuddyProperty     As String = "BuddyProp"

Private Const DEF_BooleanProps      As Long = bpVisible
Private Const DEF_BuddyControl      As String = vbNullString
Private Const DEF_Max               As Long = 1000
Private Const DEF_Min               As Long = 0
Private Const DEF_Value             As Long = 0
Private Const DEF_LargeChange       As Long = 10
Private Const DEF_SmallChange       As Long = 1
Private Const DEF_Delay             As Long = 1
Private Const DEF_Enabled           As Boolean = True
Private Const DEF_Themeable         As Boolean = True
Private Const DEF_BuddyAlignment    As Long = vbccAlignRight
Private Const DEF_BuddyProperty     As String = vbNullString

Private mhWnd                       As Long

Private msBuddy                     As String
Private miBuddyAlignment            As evbComCtlAlignment
Private msBuddyProperty             As String

Private miBooleanProps              As eBooleanProps

Private mbNoPropChange              As Boolean
Private mb32Bits                    As Boolean
Private mbThemeable                 As Boolean

Private miUpper                     As Long
Private miLower                     As Long
Private miValue                     As Long
Private miSmallChange               As Long
Private miLargeChange               As Long
Private miDelay                     As Long
Private miDispIdBuddy               As Long
Private miDispIdBuddyProp           As Long

Private mbCustomAcceleration        As Boolean
Private moBuddyNames                As pcPropertyListItems
Private moBuddyProps                As pcPropertyListItems

Private mbUserMode                  As Boolean

Private Sub pPropChanged(ByRef sProp As String)
    If Not mbNoPropChange Then PropertyChanged sProp
End Sub

Public Property Get Wrap() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return a value indicating whether the value wraps to the opposite boundary
    '             when it is incremented below the minimum or above the maximum.
    '---------------------------------------------------------------------------------------
    Wrap = CBool(miBooleanProps And bpWrap)
End Property
Public Property Let Wrap(ByVal bWrap As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set whether the value wraps to the opposite boundary
    '             when it is incremented below the minimum or above the maximum.
    '---------------------------------------------------------------------------------------
    If bWrap Xor CBool(miBooleanProps And bpWrap) Then
        If bWrap _
            Then miBooleanProps = miBooleanProps Or bpWrap _
        Else miBooleanProps = miBooleanProps And Not bpWrap
            If Not Ambient.UserMode Then pPropChanged PROP_BooleanProps
            pCreate
        End If
End Property

Public Sub SetCustomAcceleration( _
iChange() As Long, _
iDelay() As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set custom acceleration increments and delays.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        mbCustomAcceleration = True
        
        Dim tA()      As UDACCEL
        Dim i         As Long
        
        If (UBound(iChange) = UBound(iDelay)) Then
            ReDim tA(LBound(iChange) To UBound(iChange)) As UDACCEL
            For i = LBound(iChange) To UBound(iChange)
                tA(i).nInc = iChange(i)
                tA(i).nSec = iDelay(i)
            Next i
            
            SendMessage mhWnd, UDM_SETACCEL, (UBound(iChange) - LBound(iChange) + 1), VarPtr(tA(LBound(iChange)))
        End If
        
    End If
End Sub

Public Property Get SmallChange() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the small change when an arrow is clicked.
    '---------------------------------------------------------------------------------------
    SmallChange = miSmallChange
End Property
Public Property Get LargeChange() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the large change when an arrow is clicked and held down.
    '---------------------------------------------------------------------------------------
    LargeChange = miLargeChange
End Property
Public Property Let SmallChange(ByVal lSmallChange As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the small change when an arrow is clicked.
    '---------------------------------------------------------------------------------------
    If (miSmallChange <> lSmallChange) Or (mbCustomAcceleration) Then
        mbCustomAcceleration = False
        miSmallChange = lSmallChange
        pSetAccel
        If Not Ambient.UserMode Then pPropChanged PROP_SmallChange
    End If
End Property
Public Property Let LargeChange(ByVal lLargeChange As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the large change when an arrow is clicked and held down.
    '---------------------------------------------------------------------------------------
    If (miLargeChange <> lLargeChange) Or (mbCustomAcceleration) Then
        mbCustomAcceleration = False
        miLargeChange = lLargeChange
        pSetAccel
        If Not Ambient.UserMode Then pPropChanged PROP_LargeChange
    End If
End Property
Public Property Get LargeChangeDelay() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the time an arrow must be held to activate the large change.
    '---------------------------------------------------------------------------------------
    LargeChangeDelay = miDelay
End Property
Public Property Let LargeChangeDelay(ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the time an arrow must be held to activate the large change.
    '---------------------------------------------------------------------------------------
    If (iNew <> miDelay) Or mbCustomAcceleration Then
        mbCustomAcceleration = False
        miDelay = iNew
        pSetAccel
        If Not Ambient.UserMode Then pPropChanged PROP_Delay
    End If
End Property
Private Sub pSetAccel()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the standard acceleration.
    '---------------------------------------------------------------------------------------
    Dim tA(0 To 1) As UDACCEL
    If mhWnd Then
        tA(0).nInc = miSmallChange
        tA(0).nSec = 0
        tA(1).nInc = miLargeChange
        tA(1).nSec = miDelay
        SendMessage mhWnd, UDM_SETACCEL, 2&, VarPtr(tA(0))
    End If
End Sub
Public Property Get ShowThousandsSeparator() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return a value indicating whether a thousands separator is shown.
    '---------------------------------------------------------------------------------------
    ShowThousandsSeparator = Not CBool(miBooleanProps And bpNoThousands)
End Property
Public Property Let ShowThousandsSeparator(ByVal bShow As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set whether a thousands separator is shown.
    '---------------------------------------------------------------------------------------
    If bShow Xor (Not CBool(miBooleanProps And bpNoThousands)) Then
        If bShow _
            Then miBooleanProps = miBooleanProps And Not bpNoThousands _
        Else miBooleanProps = miBooleanProps Or bpNoThousands
            If Not Ambient.UserMode Then pPropChanged PROP_BooleanProps
            pCreate
        End If
End Property

Public Property Get BuddyProperty() As String
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the name of the property that is called to sync with the buddy control.
    '---------------------------------------------------------------------------------------
    BuddyProperty = msBuddyProperty
End Property
Public Property Let BuddyProperty(ByRef sNew As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the name of the property that is called to sync with the buddy control.
    '---------------------------------------------------------------------------------------
    msBuddyProperty = sNew
    pPropChanged PROP_BuddyProperty
End Property

Public Property Get Orientation() As eUpDownOrientation
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return a value indicating whether the updown control is vertical or horizontal.
    '---------------------------------------------------------------------------------------
    Orientation = (miBooleanProps And bpHorizontal)
End Property
Public Property Let Orientation(ByVal iNew As eUpDownOrientation)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set whether the updown control is vertical or horizontal.
    '---------------------------------------------------------------------------------------
    iNew = iNew And bpHorizontal
    If (iNew <> (miBooleanProps And bpHorizontal)) Then
        miBooleanProps = (miBooleanProps And Not bpHorizontal) Or iNew
        If Not Ambient.UserMode Then pPropChanged PROP_BooleanProps
        pCreate
    End If
End Property

Public Property Get Hexadecimal() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return a value indicating whether the text is displayed in hexadecimal.
    '---------------------------------------------------------------------------------------
    Hexadecimal = CBool(miBooleanProps And bpHexadecimal)
End Property
Public Property Let Hexadecimal(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set whether the text is displayed in hexadecimal.
    '---------------------------------------------------------------------------------------
    If (bNew Xor CBool(miBooleanProps And bpHexadecimal)) Then
        If bNew _
            Then miBooleanProps = miBooleanProps Or bpHexadecimal _
        Else miBooleanProps = miBooleanProps And Not bpHexadecimal
            pBuddy_SyncValue True
            'if mhwnd then
            'If bNew _
            Then sendmessage mhwnd, UDM_SETBASE, 16&, ZeroL _
        Else sendmessage mhwnd, UDM_SETBASE, 10&, ZeroL
            'End If
            pBuddy_SyncValue
        End If
End Property

Public Property Get Max() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the highest valid position.
    '---------------------------------------------------------------------------------------
    Max = miUpper
End Property
Public Property Let Max(ByVal lUpper As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the highest valid position.
    '---------------------------------------------------------------------------------------
    If (miUpper <> lUpper) Then
        If mb32Bits Then miUpper = lUpper Else miUpper = pInt(lUpper)
        pSetRange
        If Not Ambient.UserMode Then pPropChanged PROP_Max
    End If
End Property
Public Property Get Min() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Get the lowest valid position.
    '---------------------------------------------------------------------------------------
    Min = miLower
End Property
Public Property Let Min(ByVal lLower As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the lowest valid position.
    '---------------------------------------------------------------------------------------
    If (miLower <> lLower) Then
        miLower = lLower
        If mb32Bits Then miLower = lLower Else miLower = pInt(lLower)
        pSetRange
        If Not Ambient.UserMode Then pPropChanged PROP_Min
    End If
End Property

Public Property Get Value() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Get the position of the control.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        pBuddy_SyncValue True
        If mb32Bits Then
            Value = SendMessage(mhWnd, UDM_GETPOS32, ZeroL, ZeroL)
        Else
            'Must handle negative integer values!!!
            'hiword is an error value
            Value = CLng(loword(SendMessage(mhWnd, UDM_GETPOS, ZeroL, ZeroL)))
        End If
        miValue = Value
    Else
        Value = miValue
    End If
End Property
Public Property Let Value(ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the position of the control.
    '---------------------------------------------------------------------------------------
    pAdjustValue iNew
    
    If (iNew <> miValue) Then
    
        miValue = iNew
        
        If mhWnd Then
            If mb32Bits _
                Then SendMessage mhWnd, UDM_SETPOS32, 0, miValue _
            Else SendMessage mhWnd, UDM_SETPOS, 0, MakeLong(pInt(miValue), 0)
            End If
        
            If Not Ambient.UserMode Then pPropChanged PROP_Value
            pBuddy_SyncValue
            RaiseEvent Change(miValue)
        
        End If

    
End Property

Private Sub pAdjustValue(ByRef iVal As Long)
    If (iVal < miLower And iVal < miUpper) Or (iVal > miLower And iVal > miUpper) Then
        If miUpper > miLower Then
            If iVal < miLower Then iVal = miLower Else iVal = miUpper
        Else
            If iVal > miLower Then iVal = miLower Else iVal = miUpper
        End If
    End If
End Sub

Private Sub pSetRange()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the minimum and maximum bound for the updown control.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        If mb32Bits _
            Then SendMessage mhWnd, UDM_SETRANGE32, miLower, miUpper _
        Else SendMessage mhWnd, UDM_SETRANGE, 0, MakeLong(pInt(miUpper), pInt(miLower))
        End If
End Sub

Public Property Get hWnd() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the hwnd of the usercontrol.
    '---------------------------------------------------------------------------------------
    hWnd = UserControl.hWnd
End Property

Public Property Get hWndUpDown() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the hwnd of the updown control.
    '---------------------------------------------------------------------------------------
    If mhWnd Then hWndUpDown = mhWnd
End Property

Private Sub pCreate()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Create the updown and install the needed subclass.
    '---------------------------------------------------------------------------------------
    If mhWnd Then pBuddy_SyncValue True
    
    pDestroy
    
    If Ambient.UserMode Then
        
        Dim lbXPStyle      As Boolean
        lbXPStyle = IsAppThemed() And mbThemeable
        
        Dim lsAnsi      As String
        lsAnsi = StrConv(WC_UPDOWN & vbNullChar, vbFromUnicode)
        mhWnd = CreateWindowEx(((lbXPStyle + OneL) * WS_EX_CLIENTEDGE), StrPtr(lsAnsi), ZeroL, (-lbXPStyle * WS_BORDER) Or WS_CHILD Or (miBooleanProps And bpValidUDStyles) Or ((UserControl.Enabled + 1) * WS_DISABLED), ZeroL, ZeroL, ScaleWidth, ScaleHeight, UserControl.hWnd, ZeroL, App.hInstance, ByVal ZeroL)
        
        If mhWnd Then
        
            EnableWindowTheme mhWnd, mbThemeable
            
            'sendmessage mhwnd, UDM_SETBASE, IIf(CBool(miBooleanProps And bpHexadecimal), 16, 10), 0
            pSetAccel
            pSetRange
            
            'If TypeOf Parent Is ppFont Then vbBase.DisableIDEProtection
            
            Subclass_Install Me, UserControl.hWnd, Array(WM_NOTIFY, WM_VSCROLL, WM_HSCROLL, UM_SyncBuddy)
            
            If lbXPStyle Then
                'If TypeOf Parent Is ppFont Then vbBase.DisableIDEProtection
                Subclass_Install Me, mhWnd, , WM_NCPAINT
            End If
            
            If mb32Bits _
                Then SendMessage mhWnd, UDM_SETPOS32, 0, miValue _
            Else SendMessage mhWnd, UDM_SETPOS, 0, MakeLong(pInt(miValue), 0)
            
                ShowWindow mhWnd, SW_SHOWNORMAL
                PostMessage UserControl.hWnd, UM_SyncBuddy, ZeroL, ZeroL
            End If
        End If
    
End Sub

Private Sub pDestroy()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Destroy the updown control and subclasses.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        Subclass_Remove Me, UserControl.hWnd
        Subclass_Remove Me, mhWnd
        DestroyWindow mhWnd
        mhWnd = ZeroL
    End If
End Sub


Private Sub iPerPropertyBrowsingVB_GetPredefinedStrings(bHandled As Boolean, ByVal iDispID As Long, oProperties As Object)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Display names of the textboxes on the container control in the
    '             drop down list of properties.
    '---------------------------------------------------------------------------------------
    
    On Error Resume Next
    
    If iDispID = miDispIdBuddy Then
        
        Set moBuddyNames = oProperties
        
        Dim oControl      As Object
        For Each oControl In UserControl.ParentControls
            If TypeOf oControl Is TextBox Then
                moBuddyNames.Add pBuddy_FormatName(oControl)
            End If
        Next
        
        On Error GoTo 0
        
        
        bHandled = True
        
    ElseIf iDispID = miDispIdBuddyProp Then
        Set moBuddyProps = oProperties
        
        Dim loDispatch           As Interfaces.IDispatch
        Dim liTypeInfoIndex      As Long
        
        Set loDispatch = pBuddy_GetObject(msBuddy)
        
        If Not loDispatch Is Nothing Then
            loDispatch.GetTypeInfoCount liTypeInfoIndex
            ''debug.assert liTypeInfoIndex = 1
            If liTypeInfoIndex Then
                
                Dim loTypeInfo      As ITypeInfo
                Dim lpTypeAttr      As Long
                Dim ltTypeAttr      As TYPEATTR
                
                loDispatch.GetTypeInfo ZeroL, ZeroL, loTypeInfo
                loTypeInfo.GetTypeAttr VarPtr(lpTypeAttr)
                
                If lpTypeAttr Then
                    CopyMemory ltTypeAttr, ByVal lpTypeAttr, LenB(ltTypeAttr)
                    
                    Dim liIndex         As Long
                    Dim lpFuncDesc      As Long
                    Dim ltFuncDesc      As FUNCDESC
                    Dim ltElemDesc      As ELEMDESC
                    Dim lsName          As String
                    
                    For liIndex = ZeroL To ltTypeAttr.cFuncs - 1&
                        loTypeInfo.GetFuncDesc liIndex, VarPtr(lpFuncDesc)
                        If lpFuncDesc Then
                            CopyMemory ltFuncDesc, ByVal lpFuncDesc, LenB(ltFuncDesc)
                            
                            If Not CBool(ltFuncDesc.wFuncFlags And (FUNCFLAG_FHIDDEN Or FUNCFLAG_FRESTRICTED)) And CBool(ltFuncDesc.InvKind = INVOKE_PROPERTYPUT) Then
                                If ltFuncDesc.cParams = OneL Then
                                    CopyMemory ltElemDesc, ByVal ltFuncDesc.lprgElemDescParam, LenB(ltElemDesc)
                                    If ltElemDesc.tdesc.vt = VT_BSTR Then
                                        loTypeInfo.GetDocumentation ltFuncDesc.MemID, lsName, vbNullString, ZeroL, vbNullString
                                        moBuddyProps.Add lsName
                                    End If
                                End If
                            End If
                            loTypeInfo.ReleaseFuncDesc lpFuncDesc
                        End If
                    Next
                    loTypeInfo.ReleaseTypeAttr lpTypeAttr
                End If
            End If
            
            Set loTypeInfo = Nothing
            bHandled = True
        End If
    End If
    
    On Error GoTo 0
End Sub

Private Sub iPerPropertyBrowsingVB_GetPredefinedValue(bHandled As Boolean, ByVal iDispID As Long, ByVal iCookie As Long, vValue As Variant)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the control name from the cookie value that was provided in
    '             IPerPropertyBrowsingVB_GetPredefinedStrings.
    '---------------------------------------------------------------------------------------
    Dim loProperties      As pcPropertyListItems
    
    If iDispID = miDispIdBuddy Then
        Set loProperties = moBuddyNames
        
    ElseIf iDispID = miDispIdBuddyProp Then
        Set loProperties = moBuddyProps
        
    End If
    
    bHandled = Not loProperties Is Nothing
    If bHandled Then
        If loProperties.Exists(iCookie) Then
            vValue = loProperties.Item(iCookie).DisplayName
        Else
            vValue = ""
        End If
    End If
    
End Sub

Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    Select Case uMsg
    Case WM_NCPAINT
        pPaintXPBorder
        lReturn = OneL
    End Select
End Sub
Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Respond to notifications from the updown control.
    '---------------------------------------------------------------------------------------
    Dim tNMH       As NMHDR
    Dim tNMUD      As NMUPDOWN
    'Process messages here:
    Select Case uMsg
    Case WM_NOTIFY
        'debug.print "Got WM_NOTIFY"
        CopyMemory tNMH, ByVal lParam, Len(tNMH)
        If tNMH.hwndFrom = mhWnd Then
            If tNMH.code = UDN_DELTAPOS Then
                CopyMemory tNMUD, ByVal lParam, Len(tNMUD)
                On Error Resume Next
                pBuddy_SyncValue True
                pBuddy_GetObject(msBuddy).SetFocus
                On Error GoTo 0
                RaiseEvent BeforeChange(tNMUD.iPos, tNMUD.iDelta)
                lReturn = Abs(tNMUD.iDelta = ZeroL)
            End If
            bHandled = True
        End If
    Case WM_VSCROLL, WM_HSCROLL
        If (lParam = mhWnd) Then
            If mb32Bits Then
                miValue = SendMessage(mhWnd, UDM_GETPOS32, ZeroL, ZeroL)
            Else
                'Must handle negative integer values!!!
                'hiword is an error value
                miValue = CLng(loword(SendMessage(mhWnd, UDM_GETPOS, ZeroL, ZeroL)))
            End If
            
            lReturn = ZeroL
            bHandled = True
            pBuddy_SyncValue
            RaiseEvent Change(miValue)
        End If
    Case UM_SyncBuddy
        pBuddy_SyncAlignment
        pBuddy_SyncValue
    End Select

End Sub

Private Sub pPaintXPBorder()
    
    Dim tR        As RECT
    Dim lhDc      As Long
    lhDc = GetWindowDC(mhWnd)
    
    If lhDc Then
        tR.Left = ZeroL
        tR.Top = ZeroL
        tR.Right = ScaleWidth
        tR.bottom = ScaleHeight
        Select Case miBuddyAlignment
        Case vbccAlignLeft
            tR.Right = tR.Right + 50
        Case vbccAlignTop
            tR.bottom = tR.bottom + 50
        Case vbccAlignRight
            tR.Left = tR.Left - 50
        Case vbccAlignBottom
            tR.Top = tR.Top - 50
        End Select
        
        Dim lhTheme      As Long
        lhTheme = OpenThemeData(mhWnd, StrPtr("Edit"))
        If lhTheme Then
            DrawThemeBackground lhTheme, lhDc, EP_EDITTEXT, ETS_NORMAL, tR, tR
            CloseThemeData lhTheme
        End If
        
        ReleaseDC mhWnd, lhDc
        InvalidateRect mhWnd, ByVal ZeroL, OneL
        UpdateWindow mhWnd
    End If
End Sub

Private Sub UserControl_Initialize()
    LoadShellMod
    InitCC ICC_UPDOWN_CLASS
    mb32Bits = CheckCCVersion(5&, 8&)
    miDispIdBuddy = GetDispId(Me, "BuddyControl")
    miDispIdBuddyProp = GetDispId(Me, "BuddyProperty")
End Sub

Private Sub UserControl_Paint()
    If Not Ambient.UserMode Then
        Const DFCS_SCROLLDOWN = &H1&
        Const DFCS_BUTTON3STATE = &H10&
        Const DFC_SCROLL = 3&
        Dim lR      As RECT
        
        lR.Right = ScaleWidth
        lR.bottom = ScaleHeight \ 2
        DrawFrameControl UserControl.hdc, lR, DFC_SCROLL, 0
        lR.Top = lR.bottom
        lR.bottom = lR.Top + lR.bottom
        DrawFrameControl UserControl.hdc, lR, DFC_SCROLL, DFCS_SCROLLDOWN
    End If
End Sub

Private Sub UserControl_InitProperties()
    miUpper = DEF_Max
    miLower = DEF_Min
    miValue = DEF_Value
    miSmallChange = DEF_SmallChange
    miLargeChange = DEF_LargeChange
    miDelay = DEF_Delay
    miBooleanProps = DEF_BooleanProps
    msBuddy = DEF_BuddyControl
    UserControl.Enabled = DEF_Enabled
    mbThemeable = DEF_Themeable
    miBuddyAlignment = DEF_BuddyAlignment
    msBuddyProperty = DEF_BuddyProperty
    pCreate
    mbUserMode = Ambient.UserMode
    If Not mbUserMode Then VTableSubclass_PPB_Install Me
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    miUpper = PropBag.ReadProperty(PROP_Max, DEF_Max)
    miLower = PropBag.ReadProperty(PROP_Min, DEF_Min)
    miValue = PropBag.ReadProperty(PROP_Value, DEF_Value)
    miSmallChange = PropBag.ReadProperty(PROP_SmallChange, DEF_SmallChange)
    miLargeChange = PropBag.ReadProperty(PROP_LargeChange, DEF_LargeChange)
    miDelay = PropBag.ReadProperty(PROP_Delay, DEF_Delay)
    miBooleanProps = PropBag.ReadProperty(PROP_BooleanProps, DEF_BooleanProps)
    msBuddy = PropBag.ReadProperty(PROP_BuddyControl, DEF_BuddyControl)
    UserControl.Enabled = PropBag.ReadProperty(PROP_Enabled, DEF_Enabled)
    mbThemeable = PropBag.ReadProperty(PROP_Themeable, DEF_Themeable)
    miBuddyAlignment = PropBag.ReadProperty(PROP_BuddyAlignment, DEF_BuddyAlignment)
    msBuddyProperty = PropBag.ReadProperty(PROP_BuddyProperty, DEF_BuddyProperty)
    pCreate
    mbUserMode = Ambient.UserMode
    If Not mbUserMode Then VTableSubclass_PPB_Install Me
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty PROP_Min, miLower, DEF_Min
    PropBag.WriteProperty PROP_Max, miUpper, DEF_Max
    PropBag.WriteProperty PROP_Value, miValue, DEF_Value
    PropBag.WriteProperty PROP_BooleanProps, miBooleanProps, DEF_BooleanProps
    PropBag.WriteProperty PROP_SmallChange, miSmallChange, DEF_SmallChange
    PropBag.WriteProperty PROP_LargeChange, miLargeChange, DEF_LargeChange
    PropBag.WriteProperty PROP_BuddyControl, msBuddy, DEF_BuddyControl
    PropBag.WriteProperty PROP_Enabled, UserControl.Enabled, DEF_Enabled
    PropBag.WriteProperty PROP_Themeable, mbThemeable, DEF_Themeable
    PropBag.WriteProperty PROP_BuddyAlignment, miBuddyAlignment, DEF_BuddyAlignment
    PropBag.WriteProperty PROP_BuddyProperty, msBuddyProperty, DEF_BuddyProperty
End Sub

Private Sub UserControl_Terminate()
    pDestroy
    ReleaseShellMod
    If Not mbUserMode Then VTableSubclass_PPB_Remove
End Sub

Public Property Get BuddyControl() As Variant
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : In design mode, return the buddy control name to allow the property
    '             browser to display it.  In user mode, return the object so the developer
    '             can make method calls.
    '---------------------------------------------------------------------------------------
    If Ambient.UserMode _
        Then Set BuddyControl = pBuddy_GetObject(msBuddy) _
    Else BuddyControl = msBuddy
End Property

Public Property Let BuddyControl(ByVal vVal As Variant)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the buddy control by name or object.
    '---------------------------------------------------------------------------------------
    If IsObject(vVal) Then msBuddy = pBuddy_FormatName(vVal) Else msBuddy = CStr(vVal)
    If pBuddy_GetObject(msBuddy) Is Nothing Then msBuddy = vbNullString
    pBuddy_SyncAlignment
    If Not Ambient.UserMode Then pPropChanged PROP_BuddyControl
End Property

Public Property Set BuddyControl(ByVal oVal As Object)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the buddy control by object.
    '---------------------------------------------------------------------------------------
    If Not oVal Is Nothing Then
        msBuddy = pBuddy_FormatName(oVal)
        If pBuddy_GetObject(msBuddy) Is Nothing Then msBuddy = vbNullString
        pBuddy_SyncAlignment
        If Not Ambient.UserMode Then pPropChanged PROP_BuddyControl
    End If
End Property

Public Property Get BuddyAlignment() As evbComCtlAlignment
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the alignment mode that the updown control uses to move itself
    '             next to the buddy control.
    '---------------------------------------------------------------------------------------
    BuddyAlignment = miBuddyAlignment
End Property
Public Property Let BuddyAlignment(ByVal iNew As evbComCtlAlignment)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the alignment mode that the updown control uses to move itself
    '             next to the buddy control.
    '---------------------------------------------------------------------------------------
    miBuddyAlignment = iNew
    pPropChanged PROP_BuddyAlignment
    pBuddy_SyncAlignment
    If mhWnd Then SetWindowPos mhWnd, ZeroL, ZeroL, ZeroL, ZeroL, ZeroL, SWP_FRAMECHANGED Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOOWNERZORDER Or SWP_NOSIZE Or SWP_NOZORDER
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return whether the control is enabled.
    '---------------------------------------------------------------------------------------
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal bVal As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set whether the control is enabled.
    '---------------------------------------------------------------------------------------
    UserControl.Enabled = bVal
    If mhWnd Then
        EnableWindow mhWnd, -CLng(bVal)
        If bVal _
            Then SetWindowStyle mhWnd, ZeroL, WS_DISABLED _
        Else SetWindowStyle mhWnd, WS_DISABLED, ZeroL
            InvalidateRect mhWnd, ByVal ZeroL, OneL
        End If
        If Not Ambient.UserMode Then pPropChanged PROP_Enabled
End Property

Private Function pInt(ByVal i As Long) As Integer
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Convert the 32 bit to a 16 bit integer.
    '---------------------------------------------------------------------------------------
    On Error GoTo handler
    pInt = CInt(i)
    Exit Function
handler:
    pInt = Sgn(i) * 32767
End Function

Public Property Get Themeable() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Return a value indicating whether the default window theme is to be used if available.
    '---------------------------------------------------------------------------------------
    Themeable = mbThemeable
End Property

Public Property Let Themeable(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/22/05
    ' Purpose   : Set whether the default window theme is to be used if available.
    '---------------------------------------------------------------------------------------
    If bNew Xor mbThemeable Then
        pPropChanged PROP_Themeable
        mbThemeable = bNew
        If mhWnd Then EnableWindowTheme mhWnd, mbThemeable
        pCreate
    End If
End Property


Private Sub pBuddy_SyncValue(Optional ByVal bFromBuddy As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : If bFromBuddy = false
    '               If the position of the updown control has changed, update the buddy control.
    '             If bFromBuddy = true
    '               If the buddy control has changed, update the position of the updown control.
    '---------------------------------------------------------------------------------------
    
    If LenB(msBuddyProperty) And Ambient.UserMode Then
        On Error Resume Next
        
        Dim loBuddy           As Object
        Dim lsText            As String
        Dim liBuddyValue      As Long
        
        Set loBuddy = pBuddy_GetObject(msBuddy)
        
        If Not loBuddy Is Nothing Then
            
            If bFromBuddy Then
                
                lsText = Replace$(CallByName(loBuddy, msBuddyProperty, VbGet), ",", "")
                If Err.Number = ZeroL Then
                    If Left$(lsText, TwoL) = "0x" _
                        Then liBuddyValue = Val("&H" & Mid$(lsText, 3&) & "&") _
                    Else liBuddyValue = Val(lsText)
                        pAdjustValue liBuddyValue
                        If liBuddyValue <> miValue Then
                            miValue = liBuddyValue
                            If mb32Bits _
                                Then SendMessage mhWnd, UDM_SETPOS32, ZeroL, miValue _
                            Else SendMessage mhWnd, UDM_SETPOS, ZeroL, MakeLong(pInt(miValue), 0)
                                RaiseEvent Change(miValue)
                            End If
                        End If
                    Else
                
                        CallByName loBuddy, msBuddyProperty, VbLet, _
                        Switch( _
                        CBool(miBooleanProps And bpHexadecimal), _
                        "0x" & String$(IIf(Len(Hex$(miValue)) < 5, 4, 8) - Len(Hex$(miValue)), "0") & Hex$(miValue), _
                        CBool(miBooleanProps And bpNoThousands), _
                        CStr(miValue), _
                        True, _
                        Format$(miValue, "#,###,###,##0"))
                    End If
                End If
        
                On Error GoTo 0
            End If
End Sub

Private Sub pBuddy_SyncAlignment()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Move the updown control next to the buddy control.
    '---------------------------------------------------------------------------------------
    On Error GoTo handler
    
    Const IdealHeight As Long = 20
    Const IdealWidth As Long = 20
    Const BorderWidth As Long = 2
    
    If miBuddyAlignment Then
        
        Dim loControl      As Object
        Set loControl = pBuddy_GetObject(msBuddy)
        Dim ltRect      As RECT
        
        If Not loControl Is Nothing Then
            With ltRect
                .Left = ScaleX(loControl.Left, vbContainerPosition, vbPixels)
                .Top = ScaleY(loControl.Top, vbContainerPosition, vbPixels)
                .Right = ScaleX(loControl.Width, vbContainerSize, vbPixels)
                .bottom = ScaleY(loControl.Height, vbContainerSize, vbPixels)
                Select Case miBuddyAlignment
                Case vbccAlignLeft
                    .Left = .Left - IdealWidth + BorderWidth
                    .Right = IdealWidth
                Case vbccAlignTop
                    .Top = .Top - IdealHeight + BorderWidth
                    .bottom = IdealHeight
                Case vbccAlignRight
                    .Left = .Left + .Right - BorderWidth
                    .Right = IdealWidth
                Case vbccAlignBottom
                    .Top = .Top + .bottom - BorderWidth
                    .bottom = IdealHeight
                End Select
                UserControl.Extender.Move ScaleX(.Left, vbPixels, vbContainerPosition), _
                ScaleY(.Top, vbPixels, vbContainerPosition), _
                ScaleX(.Right, vbPixels, vbContainerSize), _
                ScaleY(.bottom, vbPixels, vbContainerSize)
                If mhWnd Then
                    If IsAppThemed() And mbThemeable Then
                        MoveWindow mhWnd, ZeroL, ZeroL, ScaleWidth, ScaleHeight, OneL
                    Else
                        Select Case miBuddyAlignment
                        Case vbccAlignLeft
                            MoveWindow mhWnd, ZeroL, ZeroL, ScaleWidth + BorderWidth, ScaleHeight, OneL
                        Case vbccAlignTop
                            MoveWindow mhWnd, ZeroL, ZeroL, ScaleWidth, ScaleHeight + BorderWidth, OneL
                        Case vbccAlignRight
                            MoveWindow mhWnd, -BorderWidth, ZeroL, ScaleWidth + BorderWidth, ScaleHeight, OneL
                        Case vbccAlignBottom
                            MoveWindow mhWnd, ZeroL, -BorderWidth, ScaleWidth, ScaleHeight + BorderWidth, OneL
                        End Select
                    End If
                End If
                Extender.ZOrder
            End With
        End If
    Else
        If mhWnd Then MoveWindow mhWnd, ZeroL, ZeroL, ScaleWidth, ScaleHeight, OneL
        
    End If
    
    Exit Sub
handler:
    Debug.Print "UpDown MoveBuddy Error", Err.Number, Err.Description
    ''debug.assert False
    Resume Next
End Sub


Private Sub pBuddy_ParseName(ByRef sVal As String, ByRef sName As String, ByRef iIndex As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the index and control name from a Name(Index) format.
    '---------------------------------------------------------------------------------------
    Dim liStart      As Long
    On Error Resume Next
    liStart = InStrRev(sVal, "(")
    If liStart > ZeroL Then
        sName = Left$(sVal, liStart - 1)
        iIndex = Val(Mid$(sVal, liStart + 1))
    Else
        iIndex = -1&
        sName = sVal
    End If
    On Error GoTo 0
End Sub

Private Function pBuddy_GetObject(Optional ByRef sVal As String) As Object
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Get a control from its name.
    '---------------------------------------------------------------------------------------
    If Len(sVal) = ZeroL Then sVal = msBuddy
    
    On Error Resume Next
    
    Dim liIndex      As Long
    Dim lsName       As String
    Dim loc          As Object
    
    pBuddy_ParseName sVal, lsName, liIndex
    
    Dim i      As Long
    
    For Each loc In UserControl.ParentControls
        If Not loc Is Parent Then
            Err.Clear
            If StrComp(loc.Name, lsName, vbTextCompare) = ZeroL Then
                If Err.Number = ZeroL Then
                    If liIndex = NegOneL Then
                        Set pBuddy_GetObject = loc
                        Exit For
                    ElseIf liIndex = loc.Index Then
                        If Err.Number = ZeroL Then
                            Set pBuddy_GetObject = loc
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    On Error GoTo 0
End Function

Private Function pBuddy_FormatName(ByVal oControl As Object) As String
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return a string containing the name and index of the control.
    '---------------------------------------------------------------------------------------
    On Error Resume Next
    pBuddy_FormatName = oControl.Name & "(" & oControl.Index & ")"
    If Err.Number <> ZeroL Or oControl.Index = -1& Then pBuddy_FormatName = oControl.Name
    On Error GoTo 0
End Function

Public Sub BuddyKeyDown(ByRef iKeyCode As Integer, ByRef iShift As Integer)
    If iShift = 0 And mhWnd <> ZeroL Then
        If iKeyCode = vbKeyDown Then
            iKeyCode = 0
            
            If Value = miLower Then
                If Wrap Then Value = miUpper
            Else
                Value = Value - (Value Mod miSmallChange) - miSmallChange
            End If
            
        ElseIf iKeyCode = vbKeyUp Then
            iKeyCode = 0
            If Value = miUpper Then
                If Wrap Then Value = miLower
            Else
                Value = Value + miSmallChange - (Value Mod miSmallChange)
            End If
            
        End If
    End If
End Sub

Public Sub SyncValueFromBuddy()
    pBuddy_SyncValue True
End Sub
