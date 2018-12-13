VERSION 5.00
Begin VB.UserControl ucFrame 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   HasDC           =   0   'False
   PropertyPages   =   "ucFrame.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ucFrame.ctx":000D
End
Attribute VB_Name = "ucFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'ucFrame.ctl        9/10/05
'
'             PURPOSE:
'               Emulate the intrinsic VB Frame, but provide better support for CC 6.0.
'
'---------------------------------------------------------------------------------------

Option Explicit

Public Event Resize()

Const PROP_Themeable    As String = "Themeable"
Const PROP_Font         As String = "Font"
Const PROP_Caption      As String = "Caption"
Const PROP_BackColor    As String = "BColor"
Const PROP_Border       As String = "Border"

Const DEF_Border        As Boolean = True
Const DEF_Themeable     As Boolean = True
Const DEF_Caption       As String = vbNullString
Const DEF_Backcolor     As Long = vbButtonFace

Private WithEvents moFont As cFont
Attribute moFont.VB_VarHelpID = -1
Private WithEvents moSupportFontPropPage As pcSupportFontPropPage
Attribute moSupportFontPropPage.VB_VarHelpID = -1

Private msCaption       As String
Private mbThemeable     As Boolean
Private mhFont          As Long
Private mbBorder        As Boolean

Private Sub UserControl_AmbientChanged(PropertyName As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Update the font if it is set to use the ambient font.
    '---------------------------------------------------------------------------------------
    If StrComp(PropertyName, "Font") = ZeroL Then moFont.OnAmbientFontChanged Ambient.Font
End Sub

Private Sub UserControl_Initialize()
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Load the shell module to prevent crashes when linked to CC 6.0.
    '---------------------------------------------------------------------------------------
    LoadShellMod
    Set moSupportFontPropPage = New pcSupportFontPropPage
End Sub

Private Sub UserControl_InitProperties()
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Initialize property values.
    '---------------------------------------------------------------------------------------
    Set moFont = Font_CreateDefault(Ambient.Font)
    mbThemeable = DEF_Themeable
    msCaption = DEF_Caption
    UserControl.BackColor = DEF_Backcolor
    mbBorder = DEF_Border
    pSetFont
End Sub

Private Sub UserControl_Paint()
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Paint the frame using uxtheme or the DrawEdge function.
    '             Paint the caption using uxtheme or the TextOut function.
    '---------------------------------------------------------------------------------------

    Dim lhDc         As Long: lhDc = hdc
    Dim lhTheme      As Long
    
    On Error Resume Next
    If IsAppThemed() And mbThemeable Then lhTheme = OpenThemeData(hWnd, StrPtr("Button"))
    On Error GoTo 0
    
    Dim tR      As RECT
    With tR
        .Right = ScaleWidth
        .bottom = ScaleHeight
        .Top = moFont.TextHeight("A", lhDc) \ TwoL
    End With
    
    If mbBorder Then
        If lhTheme _
            Then DrawThemeBackground lhTheme, lhDc, BP_GROUPBOX, GBS_NORMAL, tR, tR _
        Else DrawEdge lhDc, tR, EDGE_ETCHED, BF_RECT
        End If
    
        If CBool(mhFont) And LenB(msCaption) Then
    
            Const xOffset As Long = 9&
            Dim lsCaption      As String:    lsCaption = msCaption & vbNullChar
            Dim lhFontOld      As Long:      lhFontOld = SelectObject(lhDc, mhFont)
        
            If lhFontOld Then
                If lhTheme Then
                    With tR
                        .Left = xOffset
                        .Right = ScaleWidth - tR.Left
                        .bottom = ScaleHeight
                        .Top = ZeroL
                    End With
                    If GetThemeTextExtent(lhTheme, lhDc, BP_GROUPBOX, GBS_NORMAL, StrPtr(lsCaption), NegOneL, ZeroL, tR, tR) = ZeroL Then
                        Dim lhBrush      As Long: lhBrush = GdiMgr_CreateSolidBrush(TranslateColor(UserControl.BackColor))
                        If lhBrush Then
                            FillRect lhDc, tR, lhBrush
                            GdiMgr_DeleteBrush lhBrush
                        End If
                    End If
                    DrawThemeText lhTheme, lhDc, BP_GROUPBOX, GBS_NORMAL, StrPtr(lsCaption), NegOneL, ZeroL, ZeroL, tR
                Else
                    Dim liOldBkMode      As Long: liOldBkMode = SetBkMode(lhDc, OPAQUE)
                    lsCaption = StrConv(lsCaption, vbFromUnicode) 'already null terminated
                    TextOut lhDc, xOffset, ZeroL, ByVal StrPtr(lsCaption), LenB(lsCaption) - OneL
                    If liOldBkMode Then SetBkMode lhDc, liOldBkMode
                End If
                SelectObject lhDc, lhFontOld
            End If
        End If
    
        If lhTheme Then CloseThemeData lhTheme
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Save property values between instances.
    '---------------------------------------------------------------------------------------
    Set moFont = Font_Read(PropBag, PROP_Font, Ambient.Font)
    mbThemeable = PropBag.ReadProperty(PROP_Themeable, DEF_Themeable)
    msCaption = PropBag.ReadProperty(PROP_Caption, DEF_Caption)
    UserControl.BackColor = PropBag.ReadProperty(PROP_BackColor, DEF_Backcolor)
    mbBorder = PropBag.ReadProperty(PROP_Border, DEF_Border)
    pSetFont
End Sub

Private Sub UserControl_Resize()
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Redraw the control and forward the event.
    '---------------------------------------------------------------------------------------
    Refresh
    RaiseEvent Resize
End Sub

Private Sub UserControl_Terminate()
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Release the shell module and destroy our font handle.
    '---------------------------------------------------------------------------------------
    ReleaseShellMod
    If mhFont Then moFont.ReleaseHandle mhFont
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Save property values between instances.
    '---------------------------------------------------------------------------------------
    Font_Write moFont, PropBag, PROP_Font
    PropBag.WriteProperty PROP_Caption, msCaption, DEF_Caption
    PropBag.WriteProperty PROP_Themeable, mbThemeable, DEF_Themeable
    PropBag.WriteProperty PROP_BackColor, UserControl.BackColor, DEF_Backcolor
    PropBag.WriteProperty PROP_Border, mbBorder, DEF_Border
End Sub

Private Sub moFont_Changed()
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Update the font displayed by the control.
    '---------------------------------------------------------------------------------------
    moFont.OnAmbientFontChanged Ambient.Font
    pSetFont
End Sub

Private Sub moSupportFontPropPage_AddFonts(ByVal o As ppFont)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Tell the property page which properties we implement.
    '---------------------------------------------------------------------------------------
    o.ShowProps PROP_Font
End Sub

Private Sub pSetFont()
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Update the font displayed by the control.
    '---------------------------------------------------------------------------------------
    If mhFont Then moFont.ReleaseHandle mhFont
    mhFont = moFont.GetHandle()
    Refresh
    PropertyChanged PROP_Font
End Sub

Friend Property Get fSupportFontPropPage() As pcSupportFontPropPage
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Return a proxy object to receive notifications from the font property page.
    '---------------------------------------------------------------------------------------
    Set fSupportFontPropPage = moSupportFontPropPage
End Property

Public Property Get Font() As cFont
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Return the font used by the control.
    '---------------------------------------------------------------------------------------
    Set Font = moFont
End Property
Public Property Set Font(ByVal oNew As cFont)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Set the font used by the control.
    '---------------------------------------------------------------------------------------
    If oNew Is Nothing Then Set oNew = Font_CreateDefault(Ambient.Font)
    Set moFont = oNew
    pSetFont
End Property

Public Property Get Themeable() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Return whether the control uses the default xp theme if present.
    '---------------------------------------------------------------------------------------
    Themeable = mbThemeable
End Property
Public Property Let Themeable(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Set whether the control uses the default xp theme if present.
    '---------------------------------------------------------------------------------------
    mbThemeable = bNew
    PropertyChanged PROP_Themeable
    Refresh
End Property

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Return the caption displayed by the control.
    '---------------------------------------------------------------------------------------
    Caption = msCaption
End Property
Public Property Let Caption(ByRef sNew As String)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Set the caption displayed by the control.
    '---------------------------------------------------------------------------------------
    msCaption = sNew
    PropertyChanged PROP_Caption
    Refresh
End Property

Public Property Get ColorBack() As OLE_COLOR
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Return the control backcolor.
    '---------------------------------------------------------------------------------------
    ColorBack = UserControl.BackColor
End Property
Public Property Let ColorBack(ByVal iNew As OLE_COLOR)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Set the control backcolor.
    '---------------------------------------------------------------------------------------
    UserControl.BackColor = iNew
    Refresh
    PropertyChanged PROP_BackColor
End Property

Public Property Get Border() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Return whether the frame border is displayed.
    '---------------------------------------------------------------------------------------
    Border = mbBorder
End Property
Public Property Let Border(ByVal bNew As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Set whether the frame border is displayed.
    '---------------------------------------------------------------------------------------
    mbBorder = bNew
    Refresh
    PropertyChanged PROP_Border
End Property

Public Sub MoveToClient(ByVal o As Object)
    '---------------------------------------------------------------------------------------
    ' Date      : 9/10/05
    ' Purpose   : Move a control to just inside the frame border.
    '---------------------------------------------------------------------------------------
    On Error Resume Next
    Set o.Container = Extender
    
    Dim liHeight      As Long: liHeight = ScaleX(Font.TextHeight("A"), vbPixels, vbTwips) + 15
    If mbBorder _
        Then o.Move 75, liHeight, Width - 145, Height - liHeight - 75 _
    Else o.Move 0, liHeight, Width, Height - liHeight
    
        On Error GoTo 0
End Sub
