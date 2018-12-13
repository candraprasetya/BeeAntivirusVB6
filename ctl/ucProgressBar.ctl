VERSION 5.00
Begin VB.UserControl ucProgressBar 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   HasDC           =   0   'False
   ScaleHeight     =   30
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ucProgressBar.ctx":0000
End
Attribute VB_Name = "ucProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'ucProgressBar.ctl        9/10/05
'
'           PURPOSE:
'               Implement the PROGRESS_CLASS from comctl32.dll.
'
'           LINEAGE:
'               N/A
'
'==================================================================================================

Option Explicit

Private Const PROP_Value = "Value"
Private Const PROP_Min = "Min"
Private Const PROP_Max = "Max"
Private Const PROP_Smooth = "Smooth"
Private Const PROP_Vertical = "Vertical"
Private Const PROP_Themeable = "Themeable"

Private Const DEF_Value = 0
Private Const DEF_Min = 0
Private Const DEF_Max = 100
Private Const DEF_Smooth = False
Private Const DEF_Vertical = False
Private Const DEF_Themeable = True

Private mhWnd As Long

Private miValue As Long
Private miMin As Long
Private miMax As Long
Private mbSmooth As Boolean
Private mbVertical As Boolean
Private mbThemeable As Boolean

Private Sub UserControl_Initialize()
    LoadShellMod
    InitCC ICC_PROGRESS_CLASS
End Sub

Private Sub UserControl_InitProperties()
    miValue = DEF_Value
    miMin = DEF_Min
    miMax = DEF_Max
    mbSmooth = DEF_Smooth
    mbVertical = DEF_Vertical
    mbThemeable = DEF_Themeable
    pCreate
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    miValue = PropBag.ReadProperty(PROP_Value, DEF_Value)
    miMin = PropBag.ReadProperty(PROP_Min, DEF_Min)
    miMax = PropBag.ReadProperty(PROP_Max, DEF_Max)
    mbSmooth = PropBag.ReadProperty(PROP_Smooth, DEF_Smooth)
    mbVertical = PropBag.ReadProperty(PROP_Vertical, DEF_Vertical)
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
    PropBag.WriteProperty PROP_Value, miValue, DEF_Value
    PropBag.WriteProperty PROP_Min, miMin, DEF_Min
    PropBag.WriteProperty PROP_Max, miMax, DEF_Max
    PropBag.WriteProperty PROP_Smooth, mbSmooth, DEF_Smooth
    PropBag.WriteProperty PROP_Vertical, mbVertical, DEF_Vertical
    PropBag.WriteProperty PROP_Themeable, mbThemeable, DEF_Themeable
End Sub

Private Sub pCreate()
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Create the progressbar window.
    '---------------------------------------------------------------------------------------
    pDestroy
    
    Dim lsAnsi      As String
    lsAnsi = StrConv(WC_PROGRESSBAR & vbNullChar, vbFromUnicode)
    mhWnd = CreateWindowEx(ZeroL, StrPtr(lsAnsi), ZeroL, WS_CHILD Or WS_VISIBLE Or (-mbVertical * PBS_VERTICAL) Or (-mbSmooth * PBS_SMOOTH), ZeroL, ZeroL, ScaleWidth, ScaleHeight, UserControl.hWnd, ZeroL, App.hInstance, ByVal ZeroL)
    
    If mhWnd Then
        EnableWindowTheme mhWnd, mbThemeable
        SendMessage mhWnd, PBM_SETRANGE32, miMin, miMax
        SendMessage mhWnd, PBM_SETPOS, miValue, ZeroL
        miValue = SendMessage(mhWnd, PBM_GETPOS, ZeroL, ZeroL)
    End If
    
End Sub

Private Sub pDestroy()
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Destroy the progressbar window.
    '---------------------------------------------------------------------------------------
    If mhWnd Then
        DestroyWindow mhWnd
        mhWnd = ZeroL
    End If
End Sub

Public Property Get Value() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Return the progress value.
    '---------------------------------------------------------------------------------------
    miValue = SendMessage(mhWnd, PBM_GETPOS, ZeroL, ZeroL)
    Value = miValue
End Property
Public Property Let Value(ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Set the progress value.
    '---------------------------------------------------------------------------------------
    miValue = iNew
    SendMessage mhWnd, PBM_SETPOS, miValue, ZeroL
End Property

Public Property Get Max() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Return the maximum progress value.
    '---------------------------------------------------------------------------------------
    Max = miMax
    'debug.assert SendMessage(mhWnd, PBM_GETRANGE, ZeroL, ZeroL) = miMax
End Property
Public Property Let Max(ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Set the maximum progress value.
    '---------------------------------------------------------------------------------------
    miMax = iNew
    SendMessage mhWnd, PBM_SETRANGE32, miMin, miMax
End Property

Public Property Get Min() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Return the minimum progress value.
    '---------------------------------------------------------------------------------------
    Min = miMin
    'debug.assert SendMessage(mhWnd, PBM_GETRANGE, OneL, ZeroL) = miMin
End Property
Public Property Let Min(ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Return the minimum progress value.
    '---------------------------------------------------------------------------------------
    miMin = iNew
    SendMessage mhWnd, PBM_SETRANGE32, miMin, miMax
End Property

Public Property Get Vertical() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Return whether the progressbar is drawn vertically.
    '---------------------------------------------------------------------------------------
    Vertical = mbVertical
End Property
Public Property Let Vertical(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Set whether the progressbar is drawn vertically.
    '---------------------------------------------------------------------------------------
    mbVertical = bNew
    pCreate
End Property

Public Property Get Smooth() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Return whether the progressbar is drawn smoothly or in chunks.
    '---------------------------------------------------------------------------------------
    Smooth = mbSmooth
End Property
Public Property Let Smooth(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Set whether the progressbar is drawn smoothly or in chunks.
    '             Comctl version 6 always draws progressbars in chunks.
    '---------------------------------------------------------------------------------------
    mbSmooth = bNew
    pCreate
End Property

Public Sub Step(ByVal iInc As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Increment the progress value by the given amount.
    '---------------------------------------------------------------------------------------
    If mhWnd Then SendMessage mhWnd, PBM_DELTAPOS, iInc, ZeroL
End Sub

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
        PropertyChanged PROP_Themeable
        mbThemeable = bNew
        If mhWnd Then EnableWindowTheme mhWnd, mbThemeable
    End If
End Property

