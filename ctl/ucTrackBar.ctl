VERSION 5.00
Begin VB.UserControl ucTrackbar 
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3615
   HasDC           =   0   'False
   ScaleHeight     =   37
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   241
   ToolboxBitmap   =   "ucTrackBar.ctx":0000
End
Attribute VB_Name = "ucTrackbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'ucTrackBar.ctl        12/15/04
'
'           PURPOSE:
'               Implement the comctl32 trackbar control.
'
'           LINEAGE:
'               N/A
'
'==================================================================================================

Option Explicit

Public Enum eTrackbarTicStyle
    trkBottomOrRight
    trkTopOrLeft
    trkBoth
End Enum

Public Event Change()

Implements iSubclass
Implements iOleInPlaceActiveObjectVB

Private mhWnd  As Long

Private Const PROP_Max = "Max"
Private Const PROP_Min = "Min"
Private Const PROP_Pos = "Pos"
Private Const PROP_Vert = "Vert"
Private Const PROP_Style = "Style"
Private Const PROP_Freq = "Freq"
Private Const PROP_ToolTips = "ToolTips"
Private Const PROP_LineSize = "Line"
Private Const PROP_PageSize = "Page"
Private Const PROP_BackColor = "Back"
Private Const PROP_Themeable = "Themeable"

Private Const DEF_Max As Long = 10
Private Const DEF_Min As Long = 0
Private Const DEF_Pos As Long = 0
Private Const DEF_Vert As Boolean = False
Private Const DEF_Style As Long = ZeroL
Private Const DEF_Freq As Long = OneL
Private Const DEF_ToolTips As Boolean = True
Private Const DEF_LineSize As Long = OneL
Private Const DEF_PageSize As Long = ZeroL
Private Const DEF_Backcolor = vbWindowBackground
Private Const DEF_Themeable = True

Private miMax As Long
Private miMin As Long
Private miPos As Long
Private mbVert As Boolean
Private miTicStyle As eTrackbarTicStyle
Private miTicFreq As Long
Private mbToolTips As Boolean
Private miLineSize As Long
Private miPageSize As Long
Private mbThemeable As Boolean

Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    Select Case uMsg
    Case WM_SETFOCUS
        vbComCtlTlb.SetFocus mhWnd
    Case WM_KILLFOCUS
        DeActivateIPAO Me
    End Select
End Sub

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Respond to the notifications from the trackbar.
    '---------------------------------------------------------------------------------------
    Select Case uMsg
    Case WM_HSCROLL, WM_VSCROLL
        miPos = SendMessage(mhWnd, TBM_GETPOS, ZeroL, ZeroL)
        RaiseEvent Change
        bHandled = True
    Case WM_SETFOCUS
        ActivateIPAO Me
    Case WM_MOUSEACTIVATE
        If (GetFocus() <> mhWnd) Then
            vbComCtlTlb.SetFocus UserControl.hWnd
            lReturn = MA_NOACTIVATE
            bHandled = True
        End If
    End Select

End Sub

Private Sub UserControl_Initialize()
    LoadShellMod
    ForceWindowToShowAllUIStates hWnd
End Sub

Private Sub UserControl_InitProperties()
    miMax = DEF_Max
    miMin = DEF_Min
    miPos = DEF_Pos
    mbVert = DEF_Vert
    miTicStyle = DEF_Style
    miTicFreq = DEF_Freq
    mbToolTips = DEF_ToolTips
    miLineSize = DEF_LineSize
    miPageSize = DEF_PageSize
    UserControl.BackColor = DEF_Backcolor
    mbThemeable = DEF_Themeable
    pCreate
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    miMax = PropBag.ReadProperty(PROP_Max, DEF_Max)
    miMin = PropBag.ReadProperty(PROP_Min, DEF_Min)
    miPos = PropBag.ReadProperty(PROP_Pos, DEF_Pos)
    mbVert = PropBag.ReadProperty(PROP_Vert, DEF_Vert)
    miTicStyle = PropBag.ReadProperty(PROP_Style, DEF_Style)
    miTicFreq = PropBag.ReadProperty(PROP_Freq, DEF_Freq)
    mbToolTips = PropBag.ReadProperty(PROP_ToolTips, DEF_ToolTips)
    miLineSize = PropBag.ReadProperty(PROP_LineSize, DEF_LineSize)
    miPageSize = PropBag.ReadProperty(PROP_PageSize, DEF_PageSize)
    UserControl.BackColor = PropBag.ReadProperty(PROP_BackColor, DEF_Backcolor)
    mbThemeable = PropBag.ReadProperty(PROP_Themeable, DEF_Themeable)
    pCreate
End Sub

Private Sub UserControl_Resize()
    If mhWnd Then MoveWindow mhWnd, ZeroL, ZeroL, ScaleWidth, ScaleHeight, OneL
End Sub

Private Sub UserControl_Terminate()
    pDestroy
    ReleaseShellMod
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty PROP_Max, miMax, DEF_Max
    PropBag.WriteProperty PROP_Min, miMin, DEF_Min
    PropBag.WriteProperty PROP_Pos, miPos, DEF_Pos
    PropBag.WriteProperty PROP_Vert, mbVert, DEF_Vert
    PropBag.WriteProperty PROP_Style, miTicStyle, DEF_Style
    PropBag.WriteProperty PROP_Freq, miTicFreq, DEF_Freq
    PropBag.WriteProperty PROP_ToolTips, mbToolTips, DEF_ToolTips
    PropBag.WriteProperty PROP_LineSize, miLineSize, DEF_LineSize
    PropBag.WriteProperty PROP_PageSize, miPageSize, DEF_PageSize
    PropBag.WriteProperty PROP_BackColor, UserControl.BackColor, DEF_Backcolor
    PropBag.WriteProperty PROP_Themeable, mbThemeable, DEF_Themeable
End Sub

Private Sub pCreate()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Create the trackbar and install the subclasses.
    '---------------------------------------------------------------------------------------
    pDestroy
    
    Dim lsAnsi      As String
    lsAnsi = StrConv(WC_TRACKBAR & vbNullChar, vbFromUnicode)
    
    mhWnd = CreateWindowEx(ZeroL, StrPtr(lsAnsi), ZeroL, pStyle(), ZeroL, ZeroL, ScaleWidth, ScaleHeight, UserControl.hWnd, ZeroL, App.hInstance, ByVal ZeroL)
    
    If mhWnd Then
    
        EnableWindowTheme mhWnd, mbThemeable
        SendMessage mhWnd, TBM_SETRANGEMIN, ZeroL, miMin
        SendMessage mhWnd, TBM_SETRANGEMAX, ZeroL, miMax
        SendMessage mhWnd, TBM_SETPOS, OneL, miPos
        SendMessage mhWnd, TBM_SETTICFREQ, miTicFreq, ZeroL
        SendMessage mhWnd, TBM_SETLINESIZE, ZeroL, miLineSize
        pSetPageSize
        
        If Ambient.UserMode Then
        
            VTableSubclass_IPAO_Install Me
            
            Subclass_Install Me, UserControl.hWnd, Array(WM_HSCROLL, WM_VSCROLL), WM_SETFOCUS
            Subclass_Install Me, mhWnd, Array(WM_SETFOCUS, WM_MOUSEACTIVATE), WM_KILLFOCUS
            
        End If
    End If
End Sub

Private Sub pDestroy()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Destroy the trackbar and the subclasses.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        VTableSubclass_IPAO_Remove
        Subclass_Remove Me, mhWnd
        Subclass_Remove Me, UserControl.hWnd
        DestroyWindow mhWnd
        mhWnd = ZeroL
    End If
End Sub

Private Sub pSetPageSize()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the size moved with the pgup/pgdown keys and by clicking on the trackbar.
    '---------------------------------------------------------------------------------------
    If miPageSize > ZeroL Then
        SendMessage mhWnd, TBM_SETPAGESIZE, ZeroL, miPageSize
    Else
        SendMessage mhWnd, TBM_SETPAGESIZE, ZeroL, (miMax - miMin) \ 5
    End If
End Sub

Private Function pStyle() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the window style of the trackbar.
    '---------------------------------------------------------------------------------------
    pStyle = WS_CHILD Or WS_VISIBLE Or TBS_AUTOTICKS Or (-mbVert * TBS_VERT) Or (-mbToolTips * TBS_TOOLTIPS)
    If miTicStyle = trkBoth Then
        pStyle = pStyle Or TBS_BOTH
    ElseIf miTicStyle = trkTopOrLeft Then
        If mbVert Then pStyle = pStyle Or TBS_LEFT Else pStyle = pStyle Or TBS_TOP
    End If
End Function

Private Sub pPropChanged(ByRef s As String)
    If Ambient.UserMode = False Then PropertyChanged s
End Sub

Private Sub iOleInPlaceActiveObjectVB_TranslateAccelerator(bHandled As Boolean, lReturn As Long, ByVal iShift As evbComCtlKeyboardState, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Intercept the keys we want to forward to the trackbar.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        If uMsg = WM_KEYDOWN Or uMsg = WM_KEYUP Then
            Select Case wParam And &HFFFF&
            Case vbKeyPageUp To vbKeyDown, vbKeyReturn
                SendMessage mhWnd, uMsg, wParam, lParam
                bHandled = True
            End Select
        End If
    End If
End Sub

Public Property Get Max() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the maximum position of the trackbar.
    '---------------------------------------------------------------------------------------
    Max = miMax
    If mhWnd Then
        'debug.assert SendMessage(mhWnd, TBM_GETRANGEMAX, ZeroL, ZeroL) = miMax
    End If
End Property
Public Property Let Max(ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the maximum position of the trackbar.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        SendMessage mhWnd, TBM_SETRANGEMAX, OneL, iNew
        miMax = SendMessage(mhWnd, TBM_GETRANGEMAX, ZeroL, ZeroL)
        miMin = SendMessage(mhWnd, TBM_GETRANGEMIN, ZeroL, ZeroL)
        miPos = SendMessage(mhWnd, TBM_GETPOS, ZeroL, ZeroL)
    End If
    pPropChanged PROP_Max
End Property

Public Property Get Min() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the minimum position of the trackbar.
    '---------------------------------------------------------------------------------------
    Min = miMin
    If mhWnd Then
        'debug.assert SendMessage(mhWnd, TBM_GETRANGEMIN, ZeroL, ZeroL) = miMin
    End If
End Property
Public Property Let Min(ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the minimum position of the trackbar.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        SendMessage mhWnd, TBM_SETRANGEMIN, OneL, iNew
        miMax = SendMessage(mhWnd, TBM_GETRANGEMAX, ZeroL, ZeroL)
        miMin = SendMessage(mhWnd, TBM_GETRANGEMIN, ZeroL, ZeroL)
        miPos = SendMessage(mhWnd, TBM_GETPOS, ZeroL, ZeroL)
    End If
    pPropChanged PROP_Min
End Property

Public Property Get pos() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return the trackbar position.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        pos = miPos
        'debug.assert miPos = SendMessage(mhWnd, TBM_GETPOS, ZeroL, ZeroL)
    End If
End Property
Public Property Let pos(ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the trackbar position.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        SendMessage mhWnd, TBM_SETPOS, OneL, iNew
        miPos = SendMessage(mhWnd, TBM_GETPOS, ZeroL, ZeroL)
    End If
    pPropChanged PROP_Pos
End Property

Public Property Get Vertical() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Return whether the trackbar is in vertical or horizontal mode.
    '---------------------------------------------------------------------------------------
    Vertical = mbVert
End Property
Public Property Let Vertical(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set whether the trackbar is in vertical or horizontal mode.
    '---------------------------------------------------------------------------------------
    mbVert = bNew
    pCreate
    pPropChanged PROP_Vert
End Property

Public Property Get TicStyle() As eTrackbarTicStyle
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Get the position of the tics on the trackbar.
    '---------------------------------------------------------------------------------------
    TicStyle = miTicStyle
End Property
Public Property Let TicStyle(ByVal iNew As eTrackbarTicStyle)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the position of the tics on the trackbar.
    '---------------------------------------------------------------------------------------
    miTicStyle = (iNew And (trkBoth Or trkTopOrLeft))
    pCreate
    pPropChanged PROP_Style
End Property

Public Property Get TicFreq() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Get the frequency of the tics on the trackbar.
    '---------------------------------------------------------------------------------------
    TicFreq = miTicFreq
End Property
Public Property Let TicFreq(ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the frequency of the tics on the trackbar.
    '---------------------------------------------------------------------------------------
    miTicFreq = iNew And Not &H80000000
    If mhWnd Then SendMessage mhWnd, TBM_SETTICFREQ, iNew, ZeroL
    pPropChanged PROP_Freq
End Property

Public Property Get ToolTips() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Get a value indicating whether tooltips are enabled while dragging
    '             the position of the trackbar.
    '---------------------------------------------------------------------------------------
    ToolTips = mbToolTips
End Property
Public Property Let ToolTips(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set whether tooltips are enabled while dragging
    '             the position of the trackbar.
    '---------------------------------------------------------------------------------------
    mbToolTips = bNew
    pCreate
    pPropChanged PROP_ToolTips
End Property

Public Property Get LineSize() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Get the number of tics moved in response to arrow keys.
    '---------------------------------------------------------------------------------------
    LineSize = miLineSize
End Property
Public Property Let LineSize(ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the number of tics moved in response to arrow keys.
    '---------------------------------------------------------------------------------------
    miLineSize = iNew
    If mhWnd Then SendMessage mhWnd, TBM_SETLINESIZE, ZeroL, iNew
    pPropChanged PROP_LineSize
End Property

Public Property Get PageSize() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Get the number of tics moved in response to pgup/pgdown keys or clicking on the trackbar.
    '---------------------------------------------------------------------------------------
    PageSize = miPageSize
End Property
Public Property Let PageSize(ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the number of tics moved in response to pgup/pgdown keys or clicking on the trackbar.
    '---------------------------------------------------------------------------------------
    miPageSize = iNew
    pSetPageSize
    pPropChanged PROP_PageSize
End Property

Public Property Get ColorBack() As OLE_COLOR
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Get the backcolor of the usercontrol.
    '---------------------------------------------------------------------------------------
    ColorBack = UserControl.BackColor
End Property

Public Property Let ColorBack(ByVal iNew As OLE_COLOR)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/26/05
    ' Purpose   : Set the backcolor of the usercontrol.
    '---------------------------------------------------------------------------------------
    UserControl.BackColor = iNew
    pPropChanged PROP_BackColor
End Property

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
    End If
End Property

