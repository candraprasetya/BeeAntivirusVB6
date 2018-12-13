VERSION 5.00
Begin VB.UserControl XpBs 
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4290
   DefaultCancel   =   -1  'True
   EditAtDesignTime=   -1  'True
   MousePointer    =   99  'Custom
   ScaleHeight     =   111
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   286
   Tag             =   "030102-15"
   Begin VB.Timer OverTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1080
      Top             =   0
   End
   Begin VB.Image imgHAND 
      Height          =   480
      Left            =   1800
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "XpBs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetPixel Lib "gdi32" Alias "SetPixelV" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private AreaOriginal As Long
Dim Gen As Long
Dim Yuk As Long


Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long



Private Type RGB
    Red As Double
    Green As Double
    blue As Double
End Type


Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As textparametreleri) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function ShellExecute _
   Lib "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long


Private Type textparametreleri
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type

Public Enum XBPicturePosition
    gbTOP = 0
    gbLEFT = 1
    gbRIGHT = 2
    gbBOTTOM = 3
End Enum

Public Enum XBButtonStyle
    gbStandard = 0
    gbFlat = 1
    gbOfficeXP = 2
    gbWinXP = 3
    gbNoBorder = 4
End Enum

Public Enum XBPictureSize
    size16x16 = 0
    size32x32 = 1
    sizeDefault = 2
    sizeCustom = 3
End Enum

Private mvarClientRect As RECT
Private mvarPictureRect As RECT
Private mvarCaptionRect As RECT
Dim mvarOrgRect As RECT
Dim g_FocusRect As RECT
Dim alan As RECT

Dim m_OriginalPicSizeW  As Long
Dim m_OriginalPicSizeH  As Long


Dim m_PictureOriginal As Picture
Dim m_PictureHover As Picture
Dim m_Caption As String
Dim m_PicturePosition As XBPicturePosition
Dim m_ButtonStyle As XBButtonStyle
Dim m_Picture As Picture
Dim m_PictureWidth As Long
Dim m_PictureHeight As Long
Dim m_PictureSize As XBPictureSize
Dim mvarDrawTextParams As textparametreleri
Dim g_HasFocus As Byte
Dim g_MouseDown As Byte, g_MouseIn As Byte
Dim g_Button As Integer, g_Shift As Integer, g_X As Single, g_Y As Single
Dim g_KeyPressed As Byte
Dim m_URL As String
Dim m_ShowFocusRect As Boolean


Dim WithEvents g_Font As StdFont
Attribute g_Font.VB_VarHelpID = -1

Const mvarPadding As Byte = 4

Event Click()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseIn(Shift As Integer)
Event MouseOut(Shift As Integer)

Dim m_BEVEL As Integer
Dim m_BEVELDEPTH As Integer
Dim m_TransparentBG As Boolean
Dim m_MaskColor As OLE_COLOR
Dim m_XPShowBorderAlways As Boolean
Dim m_DefCurHand As Boolean
Dim m_SoundOver As String
Dim m_SoundClick As String
Dim m_ForeColor As OLE_COLOR
Dim m_BackColor As OLE_COLOR
Dim m_XPDefaultColors As Boolean
Dim m_XPColor_Pressed As OLE_COLOR
Dim m_XPColor_Hover As OLE_COLOR

Private Sub UserControl_InitProperties()
    m_BackColor = &H8000000F
    m_ForeColor = &H80000012
    m_ShowFocusRect = 1
    Set UserControl.Font = Ambient.Font
    Set g_Font = Ambient.Font
    m_Caption = Ambient.DisplayName
    m_PicturePosition = 1
    m_ButtonStyle = 2
    m_PictureWidth = 32
    m_PictureHeight = 32
    m_PictureSize = 1
    Set m_PictureHover = LoadPicture("")
    Set m_PictureOriginal = LoadPicture("")
    m_URL = ""
    m_XPColor_Pressed = &H80000014
    m_XPColor_Hover = &H80000016
    m_XPDefaultColors = 1
    
    m_SoundOver = ".\resources\over.wav"
    m_SoundClick = ".\resources\click.wav"
    m_DefCurHand = 0
    m_XPShowBorderAlways = 0
    m_MaskColor = 0
    m_TransparentBG = 0
    m_BEVEL = 1
    m_BEVELDEPTH = 8
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BackColor = m_BackColor
    m_ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.ForeColor = m_ForeColor
    
    m_ShowFocusRect = PropBag.ReadProperty("ShowFocusRect", 1)
    m_Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    m_PicturePosition = PropBag.ReadProperty("PicturePosition", 1)
    m_ButtonStyle = PropBag.ReadProperty("ButtonStyle", 2)
    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
    m_PictureWidth = PropBag.ReadProperty("PictureWidth", 32)
    m_PictureHeight = PropBag.ReadProperty("PictureHeight", 32)
    m_PictureSize = PropBag.ReadProperty("PictureSize", 1)
    m_OriginalPicSizeW = PropBag.ReadProperty("OriginalPicSizeW", 32)
    m_OriginalPicSizeH = PropBag.ReadProperty("OriginalPicSizeH", 32)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set g_Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set m_PictureHover = PropBag.ReadProperty("PictureHover", Nothing)
    Set m_PictureOriginal = PropBag.ReadProperty("Picture", Nothing)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)

    m_URL = PropBag.ReadProperty("URL", "")
    m_XPColor_Pressed = PropBag.ReadProperty("XPColor_Pressed", &H80000014)
    m_XPColor_Hover = PropBag.ReadProperty("XPColor_Hover", &H80000016)
    m_XPDefaultColors = PropBag.ReadProperty("XPDefaultColors", 1)
    
    m_SoundOver = PropBag.ReadProperty("SoundOver", ".\resources\over.wav")
    m_SoundClick = PropBag.ReadProperty("SoundClick", ".\resources\click.wav")
    m_DefCurHand = PropBag.ReadProperty("DefCurHand", 0)
    m_XPShowBorderAlways = PropBag.ReadProperty("XPShowBorderAlways", 0)
    m_MaskColor = PropBag.ReadProperty("MaskColor", 0)
    m_TransparentBG = PropBag.ReadProperty("TransparentBG", 0)
    m_BEVEL = PropBag.ReadProperty("BEVEL", 1)
    m_BEVELDEPTH = PropBag.ReadProperty("BEVELDEPTH", 8)
    SetAccessKeys
    
UserControl_Resize
End Sub

Private Sub UserControl_Terminate()
    DeleteObject AreaOriginal
    Set g_Font = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", m_Caption, Ambient.DisplayName)
    Call PropBag.WriteProperty("PicturePosition", m_PicturePosition, 1)
    Call PropBag.WriteProperty("ButtonStyle", m_ButtonStyle, 2)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("PictureWidth", m_PictureWidth, 32)
    Call PropBag.WriteProperty("PictureHeight", m_PictureHeight, 32)
    Call PropBag.WriteProperty("PictureSize", m_PictureSize, 1)
    Call PropBag.WriteProperty("OriginalPicSizeW", m_OriginalPicSizeW, 32)
    Call PropBag.WriteProperty("OriginalPicSizeH", m_OriginalPicSizeH, 32)
    
    Call PropBag.WriteProperty("PictureHover", m_PictureHover, Nothing)
    
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("ShowFocusRect", m_ShowFocusRect, 1)
 
    Call PropBag.WriteProperty("URL", m_URL, "")
    Call PropBag.WriteProperty("XPColor_Pressed", m_XPColor_Pressed, &H80000014)
    Call PropBag.WriteProperty("XPColor_Hover", m_XPColor_Hover, &H80000016)
    Call PropBag.WriteProperty("XPDefaultColors", m_XPDefaultColors, 1)
    Call PropBag.WriteProperty("BackColor", m_BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, &H80000012)
    
    Call PropBag.WriteProperty("SoundOver", m_SoundOver, ".\resources\over.wav")
    Call PropBag.WriteProperty("SoundClick", m_SoundClick, ".\resources\click.wav")
    Call PropBag.WriteProperty("DefCurHand", m_DefCurHand, 0)
    Call PropBag.WriteProperty("XPShowBorderAlways", m_XPShowBorderAlways, 0)
    Call PropBag.WriteProperty("MaskColor", m_MaskColor, 0)
    Call PropBag.WriteProperty("TransparentBG", m_TransparentBG, 0)
    Call PropBag.WriteProperty("BEVEL", m_BEVEL, 1)
    Call PropBag.WriteProperty("BEVELDEPTH", m_BEVELDEPTH, 8)
 End Sub
Private Sub CalcRECTs()
    Dim picWidth, picHeight, capWidth, capHeight As Long
    With alan
        .Left = 0
        .Top = 0
        .Right = ScaleWidth - 1
        .Bottom = ScaleHeight - 1
    End With
    
    With mvarClientRect
     .Left = alan.Left + mvarPadding
     .Top = alan.Top + mvarPadding
     .Right = alan.Right - mvarPadding + 1
     .Bottom = alan.Bottom - mvarPadding + 1
    End With
    
    If m_Picture Is Nothing Then
        With mvarCaptionRect
           .Left = mvarClientRect.Left
           .Top = mvarClientRect.Top
           .Right = mvarClientRect.Right
           .Bottom = mvarClientRect.Bottom
        End With
        CalculateCaptionRect
    Else
        If m_Caption = "" Then
         With mvarPictureRect
            .Left = (((mvarClientRect.Right - mvarClientRect.Left) - m_PictureWidth) \ 2) + mvarClientRect.Left
            .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - m_PictureHeight) \ 2) + mvarClientRect.Top
            .Right = mvarPictureRect.Left + m_PictureWidth
            .Bottom = mvarPictureRect.Top + m_PictureHeight
         End With
            Exit Sub
        End If
        
        With mvarCaptionRect
        .Left = mvarClientRect.Left
        .Top = mvarClientRect.Top
        .Right = mvarClientRect.Right
        .Bottom = mvarClientRect.Bottom
        End With
        CalculateCaptionRect
        
        picWidth = m_PictureWidth
        picHeight = m_PictureHeight
        capWidth = mvarCaptionRect.Right - mvarCaptionRect.Left
        capHeight = mvarCaptionRect.Bottom - mvarCaptionRect.Top
        
        
        If m_PicturePosition = gbLEFT Then
            With mvarPictureRect
                .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - picHeight) \ 2) + mvarClientRect.Top
                .Left = (((mvarClientRect.Right - mvarClientRect.Left) - (picWidth + mvarPadding + capWidth)) \ 2) + mvarClientRect.Left
                .Bottom = mvarPictureRect.Top + picHeight
                .Right = mvarPictureRect.Left + picWidth
            End With
            With mvarCaptionRect
                .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - capHeight) \ 2) + mvarClientRect.Top
                .Left = mvarPictureRect.Right + mvarPadding
                .Bottom = mvarCaptionRect.Top + capHeight
                .Right = mvarCaptionRect.Left + capWidth
            End With
        
        ElseIf m_PicturePosition = gbRIGHT Then
            With mvarCaptionRect
                .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - capHeight) \ 2) + mvarClientRect.Top
                .Left = (((mvarClientRect.Right - mvarClientRect.Left) - (picWidth + mvarPadding + capWidth)) \ 2) + mvarClientRect.Left
                .Bottom = mvarCaptionRect.Top + capHeight
                .Right = mvarCaptionRect.Left + capWidth
            End With
            With mvarPictureRect
                .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - picHeight) \ 2) + mvarClientRect.Top
                .Left = mvarCaptionRect.Right + mvarPadding
                .Bottom = mvarPictureRect.Top + picHeight
                .Right = mvarPictureRect.Left + picWidth
            End With
        ElseIf m_PicturePosition = gbTOP Then
            With mvarPictureRect
                .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - (picHeight + mvarPadding + capHeight)) \ 2) + mvarClientRect.Top
                .Left = (((mvarClientRect.Right - mvarClientRect.Left) - picWidth) \ 2) + mvarClientRect.Left
                .Bottom = mvarPictureRect.Top + picHeight
                .Right = mvarPictureRect.Left + picWidth
            End With
            With mvarCaptionRect
                .Top = mvarPictureRect.Bottom + mvarPadding
                .Left = (((mvarClientRect.Right - mvarClientRect.Left) - capWidth) \ 2) + mvarClientRect.Left
                .Bottom = mvarCaptionRect.Top + capHeight
                .Right = mvarCaptionRect.Left + capWidth
            End With
        ElseIf m_PicturePosition = gbBOTTOM Then
            With mvarCaptionRect
                .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - (picHeight + mvarPadding + capHeight)) \ 2) + mvarClientRect.Top
                .Left = (((mvarClientRect.Right - mvarClientRect.Left) - capWidth) \ 2) + mvarClientRect.Left
                .Bottom = mvarCaptionRect.Top + capHeight
                .Right = mvarCaptionRect.Left + capWidth
            End With
            With mvarPictureRect
                .Top = mvarCaptionRect.Bottom + mvarPadding
                .Left = (((mvarClientRect.Right - mvarClientRect.Left) - picWidth) \ 2) + mvarClientRect.Left
                .Bottom = mvarPictureRect.Top + picHeight
                .Right = mvarPictureRect.Left + picWidth
            End With
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    Set g_Font = New StdFont
    
    ScaleMode = 3
    PaletteMode = 3
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    If Not Me.Enabled Then Exit Sub
        RaiseEvent Click
        GoToURL
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    Refresh
End Sub

Private Sub UserControl_EnterFocus()
    g_HasFocus = 1
    Refresh
End Sub

Private Sub UserControl_ExitFocus()
    g_HasFocus = 0
    g_MouseDown = 0
    Refresh
End Sub

Private Sub UserControl_Resize()
    
    If ScaleWidth < 10 Then UserControl.Width = 150
    If ScaleHeight < 10 Then UserControl.Height = 150
    
Gen = ScaleWidth
Yuk = ScaleHeight

    
    g_FocusRect.Left = 4
    g_FocusRect.Right = ScaleWidth - 4
    g_FocusRect.Top = 4
    g_FocusRect.Bottom = ScaleHeight - 4
    
    DeleteObject AreaOriginal
    If m_ButtonStyle = gbWinXP Then
        RoundCorners
    End If
    Refresh
End Sub
Public Sub Refresh()
    AutoRedraw = True
                      
    
    UserControl.Cls
    
    
    XPAdjustColorScheme
    If m_ButtonStyle <> gbNoBorder Then Draw3DEffect
    CalcRECTs
    DrawPicture
    If g_HasFocus = 1 And m_ShowFocusRect And m_ButtonStyle <> gbWinXP Then DrawFocusRect hdc, g_FocusRect
    DrawCaption
    AutoRedraw = False
End Sub

Private Sub UserControl_DblClick()
    SetCapture hwnd
    UserControl_MouseDown g_Button, g_Shift, g_X, g_Y
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If g_KeyPressed = 0 Then
                             
                             
            If KeyCode = 32 Then
                g_MouseDown = 1
                g_MouseIn = 1
                Refresh
            End If
        g_KeyPressed = 1
    End If
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 Then
        g_MouseDown = 0
        g_MouseIn = 0
        GoToURL
        Refresh

        UserControl_MouseUp 1, Shift, 0, 0
    End If
    g_KeyPressed = 0
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    g_Button = Button: g_Shift = Shift: g_X = x: g_Y = y
    If Button <> 2 Then
        g_MouseDown = 1
        Refresh
    End If
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (x >= 0 And y >= 0) And (x < ScaleWidth And y < ScaleHeight) Then
        If g_MouseIn = 0 Then
            OverTimer.Enabled = True
            g_MouseIn = 1
            If Not m_PictureHover Is Nothing Then
                Set m_Picture = m_PictureHover
            End If
            RaiseEvent MouseIn(Shift)
            Refresh
            DoEvents
            
        End If
    End If
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    g_MouseDown = 0
    If Button <> 2 Then
        Refresh
        If (x >= 0 And y >= 0) And (x < ScaleWidth And y < ScaleHeight) Then
            Call PlayASound(SoundClick)
            RaiseEvent Click
            GoToURL
        End If
    End If
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    Refresh
End Property
Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512
    Set Font = g_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    With g_Font
        .Name = New_Font.Name
        .Size = New_Font.Size
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
        .Strikethrough = New_Font.Strikethrough
    End With
    PropertyChanged "Font"
End Property

Private Sub g_Font_FontChanged(ByVal PropertyName As String)
    Set UserControl.Font = g_Font
    Refresh
End Sub


Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get MouseIcon() As StdPicture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As StdPicture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get ShowFocusRect() As Boolean
    ShowFocusRect = m_ShowFocusRect
End Property

Public Property Let ShowFocusRect(ByVal New_ShowFocusRect As Boolean)
    m_ShowFocusRect = New_ShowFocusRect
    PropertyChanged "ShowFocusRect"
    Refresh
End Property
             
Private Sub RunXTRA3D(RENK As Long, BEVELL As Integer, BEVELDEPTHH As Integer)
    Dim T As Integer
    Dim TEMPRENK As Long
                TEMPRENK = RENK
                BEVELDEPTHH = BEVELDEPTHH * (-1)
                
                For T = BEVELL To 0 Step -1
                    TEMPRENK = COLOR_DarkenLightenColor(TEMPRENK, BEVELDEPTHH)
                    DRAWRECT hdc, T, T, ScaleWidth - T, ScaleHeight - T, TEMPRENK, 0
                Next T
             
                BEVELDEPTHH = BEVELDEPTHH * (-1)
                For T = BEVELL To 0 Step -1
                    RENK = RGB(COLOR_LongToRGB(RENK).Red + BEVELDEPTHH, COLOR_LongToRGB(RENK).Green + BEVELDEPTHH, COLOR_LongToRGB(RENK).blue + BEVELDEPTHH)
                    DrawLine T, T, ScaleWidth - (T + 1), T, RENK
                    DrawLine T, T, T, ScaleHeight - (T + 1), RENK
                    
                Next T
End Sub
Private Sub RunXTRA3D_PRESSED(RENK As Long, BEVELL As Integer, BEVELDEPTHH As Integer)
    Dim ret As Integer
    Dim GRIN As Integer
    Dim BLU As Integer
    Dim T As Integer
                Dim TEMPRENK As Long
                TEMPRENK = RENK
                
                For T = BEVELL To 0 Step -1
                    ret = COLOR_LongToRGB(TEMPRENK).Red + BEVELDEPTHH
                    GRIN = COLOR_LongToRGB(TEMPRENK).Green + BEVELDEPTHH
                    BLU = COLOR_LongToRGB(TEMPRENK).blue + BEVELDEPTHH
                    TEMPRENK = RGB(ret, GRIN, BLU)
                    DRAWRECT hdc, T, T, ScaleWidth - T, ScaleHeight - T, TEMPRENK, 0
                Next T
                
                
                BEVELDEPTHH = BEVELDEPTHH * (-1)
                For T = BEVELL To 0 Step -1
                    RENK = COLOR_DarkenLightenColor(RENK, BEVELDEPTHH)
                    DrawLine T, T, ScaleWidth - (T + 1), T, RENK
                    DrawLine T, T, T, ScaleHeight - (T + 1), RENK
                Next T
End Sub
Private Sub RunShowBorderOnFocus(RENK As Long, BEVELL As Integer, BEVELDEPTHH As Integer)
Dim T As Integer
            If BEVELL < 2 Then
                DRAWRECT hdc, 0, 0, ScaleWidth - 1, ScaleHeight - 1, &H80000010
                DRAWRECT hdc, 0, 0, ScaleWidth + 1, ScaleHeight + 1, &H80000014
                DRAWRECT hdc, -1, -1, ScaleWidth + 1, ScaleHeight + 1, &H80000015
            Else
                RunXTRA3D COLOR_UniColor(UserControl.BackColor), m_BEVEL, BEVELDEPTH + 3
            End If
End Sub
Private Sub XPAdjustColorScheme()
If m_ButtonStyle = gbWinXP Then Exit Sub
    If m_ButtonStyle = gbOfficeXP Then
        If m_TransparentBG = True And g_MouseDown = 0 Then
            Transparentia
        Else
            UserControl.BackColor = m_BackColor
        End If
    Else
        If m_TransparentBG = True Then Transparentia
    End If


    
    
    If m_ButtonStyle = gbOfficeXP Then
        Dim l1 As Double
        Dim l2 As Double
        Dim l3 As Double
        Dim ll As Double
        Dim KOLOR As RGB
        l1 = 171
        l2 = 154
        l3 = 108
        ll = -15
        KOLOR = COLOR_LongToRGB(COLOR_UniColor(&H8000000D))
        If g_MouseDown = 0 And g_MouseIn = 1 Then
                If XPDefaultColors = True Then
                   
                   UserControl.BackColor = RGB(KOLOR.Red + l1, KOLOR.Green + l2, _
                                                                    KOLOR.blue + l3)
                Else
                   UserControl.BackColor = XPColor_Hover
                End If
        End If
        
        If g_MouseDown = 1 Then
                If XPDefaultColors = True Then
                    UserControl.BackColor = RGB(KOLOR.Red + l1 + ll, _
                                    KOLOR.Green + l2 + ll, KOLOR.blue + l3)
                Else
                    UserControl.BackColor = XPColor_Pressed
                End If
        End If
    End If
End Sub
Private Sub Draw3DEffect()
    If Not Ambient.UserMode Then
        If m_ButtonStyle = gbWinXP Then
                DrawWinXPButton 0
        ElseIf m_ButtonStyle = gbOfficeXP Then
                XPAdjustColorScheme
        Else
            If m_BEVEL < 2 Then
                DRAWRECT hdc, 0, 0, ScaleWidth, ScaleHeight, &H80000010
                DRAWRECT hdc, 0, 0, ScaleWidth + 1, ScaleHeight + 1, &H80000014
            Else
                RunXTRA3D COLOR_UniColor(UserControl.BackColor), m_BEVEL, BEVELDEPTH
            End If
        End If
    Exit Sub
    End If
    
    
    
    
    
    
    
    
    If m_ButtonStyle = gbOfficeXP Then
                If Not (XPShowBorderAlways = False And g_MouseIn = 0) Then
                    DRAWRECT hdc, 0, 0, ScaleWidth, ScaleHeight, m_ForeColor
                End If
    ElseIf m_ButtonStyle = gbWinXP Then
            If g_MouseDown = 1 Then DrawWinXPButton 2
            If g_MouseDown = 0 And g_MouseIn = 1 Then DrawWinXPButton 0, 1
            If g_MouseDown = 0 And g_MouseIn = 0 Then DrawWinXPButton 0
    Else
        If g_MouseDown = 1 Then
            If m_BEVEL < 2 Then
                DRAWRECT hdc, 0, 0, ScaleWidth, ScaleHeight, &H80000014
                DRAWRECT hdc, 0, 0, ScaleWidth + 1, ScaleHeight + 1, &H80000010
            Else
                RunXTRA3D_PRESSED COLOR_UniColor(UserControl.BackColor), m_BEVEL, BEVELDEPTH
            End If
        End If
        If g_MouseDown = 0 And g_MouseIn = 1 Then
            If m_BEVEL < 2 Then
                DRAWRECT hdc, 0, 0, ScaleWidth, ScaleHeight, &H80000010
                DRAWRECT hdc, 0, 0, ScaleWidth + 1, ScaleHeight + 1, &H80000014
            Else
                RunXTRA3D COLOR_UniColor(UserControl.BackColor), m_BEVEL, BEVELDEPTH
            End If
        End If
        
        If g_MouseDown = 0 And g_MouseIn = 0 And m_ButtonStyle = gbStandard Then
            If m_BEVEL < 2 Then
                DRAWRECT hdc, 0, 0, ScaleWidth, ScaleHeight, &H80000010
                DRAWRECT hdc, 0, 0, ScaleWidth + 1, ScaleHeight + 1, &H80000014
            Else
                RunXTRA3D COLOR_UniColor(UserControl.BackColor), m_BEVEL, BEVELDEPTH
            End If
        End If
         
          If (g_HasFocus = 1 And m_ButtonStyle = gbStandard And g_MouseDown = 0) Or Extender.Default Then
                    RunShowBorderOnFocus COLOR_UniColor(UserControl.BackColor), m_BEVEL, BEVELDEPTH
         End If
    End If
End Sub
Private Sub OverTimer_Timer()
    Dim P As POINTAPI
    GetCursorPos P
    If hwnd <> WindowFromPoint(P.x, P.y) Then
        OverTimer.Enabled = False
        g_MouseIn = 0
        Set m_Picture = m_PictureOriginal
        RaiseEvent MouseOut(g_Shift)
        Refresh
        If g_MouseDown = 1 Then
            g_MouseDown = 0
            Refresh
            g_MouseDown = 1
        End If
    End If
End Sub

Public Sub GoToURL()
    
    If Left(m_URL, 7) = "mailto:" Then
        Navigate UserControl.Parent, m_URL
        Exit Sub
    End If
        If Not m_URL = "" Then UserControl.Hyperlink.NavigateTo m_URL
End Sub
Private Sub Navigate(frm As Form, ByVal WebPageURL As String)
Dim hBrowse As Long
hBrowse = ShellExecute(frm.hwnd, "open", WebPageURL, "", "", 1)
End Sub
Public Property Get URL() As String
    URL = m_URL
End Property

Public Property Let URL(ByVal New_URL As String)
    m_URL = New_URL
    PropertyChanged "URL"
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property
Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    SetAccessKeys
    Refresh
End Property
Public Property Get ButtonStyle() As XBButtonStyle
    ButtonStyle = m_ButtonStyle
End Property
Public Property Let ButtonStyle(ByVal New_ButtonStyle As XBButtonStyle)
    m_ButtonStyle = New_ButtonStyle
    PropertyChanged "ButtonStyle"
    If m_ButtonStyle = gbWinXP Then TransparentBG = False
    UserControl_Resize
End Property

Public Property Get PicturePosition() As XBPicturePosition
    PicturePosition = m_PicturePosition
End Property
Public Property Let PicturePosition(ByVal New_PicturePosition As XBPicturePosition)
    m_PicturePosition = New_PicturePosition
    PropertyChanged "PicturePosition"
    Refresh
End Property
Public Property Get Picture() As Picture
    Set Picture = m_Picture
End Property
Public Property Set Picture(ByVal New_Picture As Picture)
    Set m_Picture = New_Picture
    Set m_PictureOriginal = New_Picture
    If m_Picture Is Nothing Then
        m_OriginalPicSizeW = 32
        m_OriginalPicSizeH = 32
    Else
        m_OriginalPicSizeW = UserControl.ScaleX(m_Picture.Width, 8, UserControl.ScaleMode)
        m_OriginalPicSizeH = UserControl.ScaleY(m_Picture.Height, 8, UserControl.ScaleMode)
    End If
    PropertyChanged "Picture"
    If m_PictureSize = sizeDefault Then
        m_PictureWidth = UserControl.ScaleX(m_Picture.Width, 8, UserControl.ScaleMode)
        m_PictureHeight = UserControl.ScaleY(m_Picture.Height, 8, UserControl.ScaleMode)
    End If
    Refresh
End Property

Public Property Get PictureWidth() As Long
    PictureWidth = m_PictureWidth
End Property
Public Property Let PictureWidth(ByVal New_PictureWidth As Long)
    m_PictureWidth = New_PictureWidth
    PropertyChanged "PictureWidth"
    Refresh
End Property
Public Property Get PictureHeight() As Long
    PictureHeight = m_PictureHeight
End Property
Public Property Let PictureHeight(ByVal New_PictureHeight As Long)
    m_PictureHeight = New_PictureHeight
    PropertyChanged "PictureHeight"
    Refresh
End Property
Public Property Get PictureSize() As XBPictureSize
    PictureSize = m_PictureSize
End Property
Public Property Let PictureSize(ByVal New_PictureSize As XBPictureSize)
    m_PictureSize = New_PictureSize
    PropertyChanged "PictureSize"
    
    If New_PictureSize = size16x16 Then
        m_PictureWidth = 16
        m_PictureHeight = 16
    ElseIf New_PictureSize = size32x32 Then
        m_PictureWidth = 32
        m_PictureHeight = 32
    ElseIf New_PictureSize = sizeDefault Then
        If Not (m_Picture Is Nothing) Then
            m_PictureWidth = m_OriginalPicSizeW
            m_PictureHeight = m_OriginalPicSizeH
        Else
            m_PictureWidth = 32
            m_PictureHeight = 32
        End If
    End If
   
    Refresh
End Property

Private Sub CalculateCaptionRect()
    Dim mvarWidth, mvarHeight As Long
    Dim mvarFormat As Long
    With mvarDrawTextParams
        .iLeftMargin = 1
        .iRightMargin = 1
        .iTabLength = 1
        .cbSize = Len(mvarDrawTextParams)
    End With
    mvarFormat = &H400 Or &H10 Or &H4 Or &H1
    DrawTextEx hdc, m_Caption, Len(m_Caption), mvarCaptionRect, mvarFormat, mvarDrawTextParams
    mvarWidth = mvarCaptionRect.Right - mvarCaptionRect.Left
    mvarHeight = mvarCaptionRect.Bottom - mvarCaptionRect.Top
    With mvarCaptionRect
        .Left = mvarClientRect.Left + (((mvarClientRect.Right - mvarClientRect.Left) - (mvarCaptionRect.Right - mvarCaptionRect.Left)) \ 2)
        .Top = mvarClientRect.Top + (((mvarClientRect.Bottom - mvarClientRect.Top) - (mvarCaptionRect.Bottom - mvarCaptionRect.Top)) \ 2)
        .Right = mvarCaptionRect.Left + mvarWidth
        .Bottom = mvarCaptionRect.Top + mvarHeight
    End With
End Sub

Private Sub DrawCaption()
    If m_Caption = "" Then Exit Sub
    
    SetTextColor hdc, COLOR_UniColor(m_ForeColor)
    
    Dim mvarForeColor As OLE_COLOR
    mvarOrgRect = mvarCaptionRect
    If g_MouseDown = 1 And m_ButtonStyle <> gbOfficeXP Then
       With mvarCaptionRect
        .Left = mvarCaptionRect.Left + 1
        .Top = mvarCaptionRect.Top + 1
        .Right = mvarCaptionRect.Right + 1
        .Bottom = mvarCaptionRect.Bottom + 1
       End With
    End If
    
    If Not Enabled Then
        Dim g_tmpFontColor As OLE_COLOR
        g_tmpFontColor = UserControl.ForeColor
        
        
        SetTextColor hdc, COLOR_UniColor(&H80000014)
        Dim mvarCaptionRect_Iki As RECT
        With mvarCaptionRect_Iki
            .Bottom = mvarCaptionRect.Bottom
            .Left = mvarCaptionRect.Left + 1
            .Right = mvarCaptionRect.Right + 1
            .Top = mvarCaptionRect.Top + 1
        End With
        DrawTextEx hdc, m_Caption, Len(m_Caption), mvarCaptionRect_Iki, &H10 Or &H4 Or &H1, mvarDrawTextParams
        
        
        SetTextColor hdc, COLOR_UniColor(&H80000010)
        DrawTextEx hdc, m_Caption, Len(m_Caption), mvarCaptionRect, &H10 Or &H4 Or &H1, mvarDrawTextParams
        
        
        SetTextColor hdc, COLOR_UniColor(g_tmpFontColor)
        Exit Sub
    End If
    DrawTextEx hdc, m_Caption, Len(m_Caption), mvarCaptionRect, &H10 Or &H4 Or &H1, mvarDrawTextParams
    mvarCaptionRect = mvarOrgRect
End Sub
Private Sub DrawBitmap(EnabledPic As Byte, CurPictRECT As RECT, _
                            Optional AsShadow As Byte = 0)
Dim DC1 As Long
Dim BM1 As Long
Dim DC2 As Long
Dim BM2 As Long
Dim UZUN1 As Long
Dim UZUN2 As Long
Dim hBrush As Long

DC1 = CreateCompatibleDC(hdc)
DC2 = CreateCompatibleDC(hdc)
BM1 = CreateCompatibleBitmap(hdc, m_OriginalPicSizeW, m_OriginalPicSizeH)
BM2 = CreateCompatibleBitmap(hdc, m_PictureWidth, m_PictureHeight)
UZUN1 = SelectObject(DC1, BM1)
UZUN2 = SelectObject(DC2, BM2)

If EnabledPic = 0 Then
                Dim DC3 As Long
                Dim BM3 As Long
                
                DC3 = CreateCompatibleDC(hdc)
                BM3 = SelectObject(DC3, m_Picture.Handle)
                
                SetBkColor DC1, &HFFFFFF
                 
                DRAWRECT DC1, 0, 0, _
                    m_OriginalPicSizeW, m_OriginalPicSizeH, &HFFFFFF, 1

                TransParentPic DC1, DC1, DC3, 0, 0, _
                    m_OriginalPicSizeW, m_OriginalPicSizeH, 0, 0, m_MaskColor
                
                StretchBlt DC2, 0, 0, _
                    m_PictureWidth, _
                        m_PictureHeight, _
                            DC1, 0, 0, m_OriginalPicSizeW, m_OriginalPicSizeH, &HCC0020
                
                SelectObject DC2, UZUN2
                
                

                
                If AsShadow = 1 Then
                    hBrush = CreateSolidBrush(RGB(146, 146, 146))
                    Call DrawState(hdc, hBrush, 0, BM2, 0, CurPictRECT.Left, _
                                 CurPictRECT.Top, 0, 0, &H80& Or &H4&)
                    DeleteObject hBrush
                Else
                    Call DrawState(hdc, 0, 0, BM2, 0, CurPictRECT.Left, _
                                 CurPictRECT.Top, 0, 0, &H20& Or &H4&)
                End If


                

    DeleteObject BM3
    DeleteDC DC3
                
Else
                
                Call DrawState(DC1, 0, 0, m_Picture, 0, 0, 0, 0, 0, _
                    &H0 Or &H4&)
            
                StretchBlt DC2, 0, 0, _
                    m_PictureWidth, _
                        m_PictureHeight, _
                            DC1, 0, 0, m_OriginalPicSizeW, m_OriginalPicSizeH, &HCC0020
                            
                TransParentPic hdc, hdc, DC2, 0, 0, _
                    m_PictureWidth, m_PictureHeight, _
                     CurPictRECT.Left, CurPictRECT.Top, m_MaskColor
                
End If

    SelectObject DC1, UZUN1
    SelectObject DC2, UZUN2
    DeleteObject BM1
    DeleteObject BM2
    DeleteDC DC1
    DeleteDC DC2
End Sub
Private Sub DrawPIcon(EnabledPic As Byte, CurPictRECT As RECT, Optional AsShadow As Byte = 0)
If EnabledPic = 0 Then
                 Dim DC1 As Long
                Dim BM1 As Long
                Dim DC2 As Long
                Dim BM2 As Long
                Dim UZUN1 As Long
                Dim UZUN2 As Long
                Dim hBrush As Long
                    
                    

                
                DC1 = CreateCompatibleDC(hdc)
                BM1 = CreateCompatibleBitmap(hdc, m_OriginalPicSizeW, m_OriginalPicSizeH)
            
                DC2 = CreateCompatibleDC(hdc)
                BM2 = CreateCompatibleBitmap(hdc, m_PictureWidth, m_PictureHeight)
            
                UZUN1 = SelectObject(DC1, BM1)
                UZUN2 = SelectObject(DC2, BM2)
                
                
                If AsShadow = 1 Then
                    hBrush = CreateSolidBrush(RGB(146, 146, 146))
                    Call DrawState(DC1, hBrush, 0, m_Picture, 0, 0, 0, 0, 0, _
                        &H80& Or &H3&)
                    DeleteObject hBrush
                Else
                    Call DrawState(DC1, 0, 0, m_Picture, 0, 0, 0, 0, 0, _
                       &H20& Or &H3&)
                End If
                
                
                StretchBlt DC2, 0, 0, _
                    CurPictRECT.Right - CurPictRECT.Left, _
                        CurPictRECT.Bottom - CurPictRECT.Top, _
                            DC1, 0, 0, m_OriginalPicSizeW, m_OriginalPicSizeH, &HCC0020
                            
                TransParentPic hdc, hdc, DC2, 0, 0, _
                    m_PictureWidth, m_PictureHeight, _
                      CurPictRECT.Left, CurPictRECT.Top, &H0
                
                SelectObject DC1, UZUN1
                SelectObject DC2, UZUN2
                DeleteObject BM1
                DeleteObject BM2
                DeleteDC DC1
                DeleteDC DC2

Else


            UserControl.PaintPicture m_Picture, CurPictRECT.Left, _
                CurPictRECT.Top, CurPictRECT.Right - CurPictRECT.Left, _
                  CurPictRECT.Bottom - CurPictRECT.Top, 0, 0, _
                    m_OriginalPicSizeW, m_OriginalPicSizeH
End If
End Sub

Private Sub DrawPicture()
    Dim Margin As Integer
    
    If m_Picture Is Nothing Then Exit Sub
    mvarOrgRect = mvarPictureRect
    
    
    If g_MouseDown = 0 And g_MouseIn = 1 And m_ButtonStyle = gbOfficeXP Then
      
        Margin = -3
    ElseIf g_MouseDown = 1 And Not m_ButtonStyle = gbOfficeXP Then
      
        Margin = 1
    End If
    
    With mvarPictureRect
        .Left = .Left + Margin
        .Top = .Top + Margin
        .Right = .Right + Margin
        .Bottom = .Bottom + Margin
    End With


        If m_Picture.Type = 1 Then
            If Not Enabled Then
                DrawBitmap 0, mvarPictureRect
            Else
                If g_MouseDown = 0 And g_MouseIn = 1 And _
                            m_ButtonStyle = gbOfficeXP Then _
                    DrawBitmap 0, mvarOrgRect, 1
                
                DrawBitmap 1, mvarPictureRect
            End If
        ElseIf m_Picture.Type = 3 Then
            If Not Enabled Then
                DrawPIcon 0, mvarPictureRect
            Else
                If g_MouseDown = 0 And g_MouseIn = 1 And _
                        m_ButtonStyle = gbOfficeXP Then _
                    DrawPIcon 0, mvarOrgRect, 1
                    
                DrawPIcon 1, mvarPictureRect
            End If
        End If
mvarPictureRect = mvarOrgRect
End Sub
Private Sub Transparentia()
  On Error Resume Next
Dim RESIM As StdPicture
Dim mem_dc As Long
Dim mem_bm As Long
Dim orig_bm As Long
Dim wid As Long
Dim hgt As Long
Dim IX As Long
Dim YE As Long

IX = ScaleX(Extender.Left, Parent.ScaleMode, ScaleMode)
YE = ScaleY(Extender.Top, Parent.ScaleMode, ScaleMode)

Set RESIM = Parent.Picture
    mem_dc = CreateCompatibleDC(hdc)
    mem_bm = CreateCompatibleBitmap(mem_dc, ScaleWidth, ScaleHeight)
    
    SelectObject mem_dc, RESIM.Handle
    
    BitBlt hdc, 0, 0, ScaleWidth, ScaleHeight, _
        mem_dc, IX, YE, &HCC0020
    

    SelectObject mem_dc, orig_bm
    DeleteObject mem_bm
    DeleteDC mem_dc
    Set RESIM = Nothing
End Sub

Public Property Get PictureHover() As Picture
    Set PictureHover = m_PictureHover
End Property

Public Property Set PictureHover(ByVal New_PictureHover As Picture)
    Set m_PictureHover = New_PictureHover
    PropertyChanged "PictureHover"
End Property
Public Property Get XPColor_Pressed() As OLE_COLOR
    XPColor_Pressed = m_XPColor_Pressed
End Property

Public Property Let XPColor_Pressed(ByVal New_XPColor_Pressed As OLE_COLOR)
    m_XPColor_Pressed = New_XPColor_Pressed
    PropertyChanged "XPColor_Pressed"
End Property
Public Property Get XPColor_Hover() As OLE_COLOR
    XPColor_Hover = m_XPColor_Hover
End Property

Public Property Let XPColor_Hover(ByVal New_XPColor_Hover As OLE_COLOR)
    m_XPColor_Hover = New_XPColor_Hover
    PropertyChanged "XPColor_Hover"
End Property
Public Property Get XPDefaultColors() As Boolean
    XPDefaultColors = m_XPDefaultColors
End Property
Public Property Let XPDefaultColors(ByVal New_XPDefaultColors As Boolean)
    m_XPDefaultColors = New_XPDefaultColors
    PropertyChanged "XPDefaultColors"
End Property
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    UserControl.BackColor = m_BackColor
    Refresh
End Property
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    UserControl.ForeColor = m_ForeColor
    Refresh
End Property
Public Property Get SoundOver() As Variant
    SoundOver = m_SoundOver
End Property
Public Property Let SoundOver(ByVal New_SoundOver As Variant)
    m_SoundOver = New_SoundOver
    PropertyChanged "SoundOver"
End Property
Public Property Get SoundClick() As String
    SoundClick = m_SoundClick
End Property
Public Property Let SoundClick(ByVal New_SoundClick As String)
    m_SoundClick = New_SoundClick
    PropertyChanged "SoundClick"
End Property
Public Property Get version() As String
Attribute version.VB_Description = "FileVersion"
    version = UserControl.Tag
End Property
Public Property Let version(ByVal New_version As String)
End Property
Private Function PlayASound(SoundFile As String) As Byte
 
    PlayASound = PlaySound(SoundFile, 1, &H20000 _
    + &H0 + &H1 + &H2)
End Function
Public Property Get DefCurHand() As Boolean
    DefCurHand = m_DefCurHand
End Property

Public Property Let DefCurHand(ByVal New_DefCurHand As Boolean)
    m_DefCurHand = New_DefCurHand
    PropertyChanged "DefCurHand"
    
    If m_DefCurHand = True Then
        
        
    Else
        
    End If

End Property

Public Property Get XPShowBorderAlways() As Boolean
    XPShowBorderAlways = m_XPShowBorderAlways
End Property

Public Property Let XPShowBorderAlways(ByVal New_XPShowBorderAlways As Boolean)
    m_XPShowBorderAlways = New_XPShowBorderAlways
    PropertyChanged "XPShowBorderAlways"
End Property
Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets the color that specifies transparent areas in the MaskPicture."
    MaskColor = m_MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
    m_MaskColor = New_MaskColor
    PropertyChanged "MaskColor"
    Refresh
End Property
Public Property Get TransparentBG() As Boolean
    TransparentBG = m_TransparentBG
End Property

Public Property Let TransparentBG(ByVal New_TransparentBG As Boolean)
    m_TransparentBG = New_TransparentBG
    PropertyChanged "TransparentBG"
    Refresh
End Property

Public Property Get BEVEL() As Integer
    BEVEL = m_BEVEL
End Property

Public Property Let BEVEL(ByVal New_BEVEL As Integer)
    m_BEVEL = New_BEVEL
    PropertyChanged "BEVEL"
    Refresh
End Property
Public Property Get BEVELDEPTH() As Integer
    BEVELDEPTH = m_BEVELDEPTH
End Property

Public Property Let BEVELDEPTH(ByVal New_BEVELDEPTH As Integer)
    m_BEVELDEPTH = New_BEVELDEPTH
    PropertyChanged "BEVELDEPTH"
    Refresh
End Property

Private Function COLOR_LongToRGB(UniColorValue As Long) As RGB
    Dim BlueS As Double, GreenS As Double, RGBs As String
    COLOR_LongToRGB.blue = Fix((UniColorValue / 256) / 256)
    BlueS = (COLOR_LongToRGB.blue * 256) * 256
    COLOR_LongToRGB.Green = Fix((UniColorValue - BlueS) / 256)
    GreenS = COLOR_LongToRGB.Green * 256
    COLOR_LongToRGB.Red = Fix(UniColorValue - BlueS - GreenS)

End Function
Private Function COLOR_UniColor(ColorVal As Long) As Long

    COLOR_UniColor = ColorVal
    If ColorVal > &HFFFFFF Or ColorVal < 0 Then COLOR_UniColor = GetSysColor(ColorVal And &HFFFFFF)
End Function
Private Function COLOR_DarkenLightenColor(ByVal Color As Long, ByVal Value As Long) As Long
    Dim R As Long, G As Long, B As Long
    B = ((Color \ &H10000) Mod &H100): B = B + ((B * Value) \ &HC0)
    G = ((Color \ &H100) Mod &H100) + Value
    R = (Color And &HFF) + Value
        If R < 0 Then R = 0
        If R > 255 Then R = 255
        If G < 0 Then G = 0
        If G > 255 Then G = 255
        If B < 0 Then B = 0
        If B > 255 Then B = 255
    COLOR_DarkenLightenColor = RGB(R, G, B)
End Function


Private Sub DrawLine(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)
    Dim pt As POINTAPI
    Call DeleteObject(SelectObject(hdc, CreatePen(0, 1, Color)))
    MoveToEx hdc, X1, Y1, pt
    LineTo hdc, X2, Y2
End Sub

Private Sub DRAWRECT(DestHDC As Long, ByVal RectLEFT As Long, _
            ByVal RectTOP As Long, _
            ByVal RectRIGHT As Long, ByVal RectBOTTOM As Long, _
            ByVal MyColor As Long, _
            Optional FillRectWithColor As Byte = 0)
    Dim MyRect As RECT, Firca As Long
    Firca = CreateSolidBrush(COLOR_UniColor(MyColor))
    With MyRect
        .Left = RectLEFT
        .Top = RectTOP
        .Right = RectRIGHT
        .Bottom = RectBOTTOM
    End With
    If FillRectWithColor = 1 Then FillRect DestHDC, MyRect, Firca Else FrameRect DestHDC, MyRect, Firca
    DeleteObject Firca
End Sub

Private Sub DrawWinXPButton(ByVal None_Press_Disabled As Byte, Optional HOVERING As Byte)
Dim x As Long, Intg As Single, curBackColor As Long, OuterBorderColor As Long
Dim KolorHover As Long, KolorPressed As Long
DRAWRECT hdc, 0, 0, Gen, Yuk, m_BackColor, 1
OuterBorderColor = &H80000015
If Enabled Then
    If m_XPDefaultColors = True Then
        KolorPressed = RGB(140, 170, 230)
        KolorHover = RGB(225, 153, 71)
    Else
        KolorPressed = m_XPColor_Pressed
        KolorHover = m_XPColor_Hover
    End If


    If None_Press_Disabled = 0 Then
             Intg = 25 / Yuk: curBackColor = COLOR_DarkenLightenColor(COLOR_UniColor(m_BackColor), 48)
             For x = 1 To Yuk
                 DrawLine 0, x, Gen, x, COLOR_DarkenLightenColor(curBackColor, -Intg * x)
             Next
           
             DRAWRECT hdc, 0, 0, Gen, Yuk, OuterBorderColor
             SetPixel hdc, 1, 1, OuterBorderColor
             SetPixel hdc, 1, Yuk - 2, OuterBorderColor
             SetPixel hdc, Gen - 2, 1, OuterBorderColor
             SetPixel hdc, Gen - 2, Yuk - 2, OuterBorderColor

             If g_HasFocus = 1 Then
                 DRAWRECT hdc, 1, 2, Gen - 1, Yuk - 2, KolorPressed
                 DrawLine 2, Yuk - 2, Gen - 2, Yuk - 2, COLOR_DarkenLightenColor(COLOR_UniColor(KolorPressed), -33)
                 DrawLine 2, 1, Gen - 2, 1, COLOR_DarkenLightenColor(COLOR_UniColor(KolorPressed), 65)
                 DrawLine 1, 2, Gen - 1, 2, COLOR_DarkenLightenColor(COLOR_UniColor(KolorPressed), 50)
                 DrawLine 2, 3, 2, Yuk - 3, COLOR_DarkenLightenColor(COLOR_UniColor(KolorPressed), 31)
                 DrawLine Gen - 3, 3, Gen - 3, Yuk - 3, COLOR_DarkenLightenColor(COLOR_UniColor(KolorPressed), 31)
             Else
                 DrawLine 2, Yuk - 2, Gen - 2, Yuk - 2, COLOR_DarkenLightenColor(curBackColor, -48)
                 DrawLine 1, Yuk - 3, Gen - 2, Yuk - 3, COLOR_DarkenLightenColor(curBackColor, -32)
                 DrawLine Gen - 2, 2, Gen - 2, Yuk - 2, COLOR_DarkenLightenColor(curBackColor, -36)
                 DrawLine Gen - 3, 3, Gen - 3, Yuk - 3, COLOR_DarkenLightenColor(curBackColor, -24)
                 DrawLine 2, 1, Gen - 2, 1, COLOR_DarkenLightenColor(curBackColor, 16)
                 DrawLine 1, 2, Gen - 2, 2, COLOR_DarkenLightenColor(curBackColor, 10)
                 DrawLine 1, 2, 1, Yuk - 2, COLOR_DarkenLightenColor(curBackColor, -5)
                 DrawLine 2, 3, 2, Yuk - 3, COLOR_DarkenLightenColor(curBackColor, -10)
             End If
             If HOVERING = 1 Then
                 DRAWRECT hdc, 1, 2, Gen - 1, Yuk - 2, KolorHover
                 DrawLine 2, Yuk - 2, Gen - 2, Yuk - 2, COLOR_DarkenLightenColor(KolorHover, -40)
                 DrawLine 2, 1, Gen - 2, 1, COLOR_DarkenLightenColor(KolorHover, 90)
                 DrawLine 1, 2, Gen - 1, 2, COLOR_DarkenLightenColor(KolorHover, 35)
                 DrawLine 2, 3, 2, Yuk - 3, COLOR_DarkenLightenColor(KolorHover, 20)
                 DrawLine Gen - 3, 3, Gen - 3, Yuk - 3, COLOR_DarkenLightenColor(KolorHover, 20)
             End If
    ElseIf None_Press_Disabled = 2 Then
            Intg = 15 / Yuk
            curBackColor = COLOR_DarkenLightenColor(COLOR_UniColor(m_BackColor), 48)
            curBackColor = COLOR_DarkenLightenColor(curBackColor, -32)
            For x = 1 To Yuk
                DrawLine 0, Yuk - x, Gen, Yuk - x, COLOR_DarkenLightenColor(curBackColor, -Intg * x)
            Next
            DRAWRECT hdc, 0, 0, Gen, Yuk, OuterBorderColor
            SetPixel hdc, 1, 1, OuterBorderColor
            SetPixel hdc, 1, Yuk - 2, OuterBorderColor
            SetPixel hdc, Gen - 2, 1, OuterBorderColor
            SetPixel hdc, Gen - 2, Yuk - 2, OuterBorderColor
            
            DrawLine 2, Yuk - 2, Gen - 2, Yuk - 2, COLOR_DarkenLightenColor(curBackColor, 16)
            DrawLine 1, Yuk - 3, Gen - 2, Yuk - 3, COLOR_DarkenLightenColor(curBackColor, 10)
            DrawLine Gen - 2, 2, Gen - 2, Yuk - 2, COLOR_DarkenLightenColor(curBackColor, 5)
            DrawLine Gen - 3, 3, Gen - 3, Yuk - 3, curBackColor
            DrawLine 2, 1, Gen - 2, 1, COLOR_DarkenLightenColor(curBackColor, -32)
            DrawLine 1, 2, Gen - 2, 2, COLOR_DarkenLightenColor(curBackColor, -24)
            DrawLine 1, 2, 1, Yuk - 2, COLOR_DarkenLightenColor(curBackColor, -32)
            DrawLine 2, 2, 2, Yuk - 2, COLOR_DarkenLightenColor(curBackColor, -22)
    End If
Else
        curBackColor = COLOR_DarkenLightenColor(COLOR_UniColor(m_BackColor), 48)
        DRAWRECT hdc, 0, 0, Gen, Yuk, COLOR_DarkenLightenColor(curBackColor, -24), 1
        DRAWRECT hdc, 0, 0, Gen, Yuk, COLOR_DarkenLightenColor(curBackColor, -84)
        SetPixel hdc, 1, 1, COLOR_DarkenLightenColor(curBackColor, -72)
        SetPixel hdc, 1, Yuk - 2, COLOR_DarkenLightenColor(curBackColor, -72)
        SetPixel hdc, Gen - 2, 1, COLOR_DarkenLightenColor(curBackColor, -72)
        SetPixel hdc, Gen - 2, Yuk - 2, COLOR_DarkenLightenColor(curBackColor, -72)
End If
End Sub

Private Sub RoundCorners()
Dim Alan1 As Long, Alan2 As Long
    DeleteObject AreaOriginal
    AreaOriginal = CreateRectRgn(0, 0, Gen, Yuk)
    Alan2 = CreateRectRgn(0, 0, 0, 0)
        Alan1 = CreateRectRgn(0, 0, 2, 1)
        CombineRgn Alan2, AreaOriginal, Alan1, 4
        DeleteObject Alan1
        Alan1 = CreateRectRgn(0, Yuk, 2, Yuk - 1)
        CombineRgn AreaOriginal, Alan2, Alan1, 4
        DeleteObject Alan1
        Alan1 = CreateRectRgn(Gen, 0, Gen - 2, 1)
        CombineRgn Alan2, AreaOriginal, Alan1, 4
        DeleteObject Alan1
        Alan1 = CreateRectRgn(Gen, Yuk, Gen - 2, Yuk - 1)
        CombineRgn AreaOriginal, Alan2, Alan1, 4
        DeleteObject Alan1
        Alan1 = CreateRectRgn(0, 1, 1, 2)
        CombineRgn Alan2, AreaOriginal, Alan1, 4
        DeleteObject Alan1
        Alan1 = CreateRectRgn(0, Yuk - 1, 1, Yuk - 2)
        CombineRgn AreaOriginal, Alan2, Alan1, 4
        DeleteObject Alan1
        Alan1 = CreateRectRgn(Gen, 1, Gen - 1, 2)
        CombineRgn Alan2, AreaOriginal, Alan1, 4
        DeleteObject Alan1
        Alan1 = CreateRectRgn(Gen, Yuk - 1, Gen - 1, Yuk - 2)
        CombineRgn AreaOriginal, Alan2, Alan1, 4
        DeleteObject Alan1
DeleteObject Alan2
SetWindowRgn hwnd, AreaOriginal, True
End Sub
Private Sub TransParentPic(DestDC As Long, _
                           DestDCTrans As Long, _
                           SrcDC As Long, _
                           SrcRectLeft As Long, SrcRectTop As Long, _
                           SrcRectRight As Long, SrcRectBottom As Long, _
                           DstX As Long, _
                           DstY As Long, _
                           MaskColor As Long)
   
  Dim nRet As Long, w As Integer, h As Integer
  Dim MonoMaskDC As Long, hMonoMask As Long
  Dim MonoInvDC As Long, hMonoInv As Long
  Dim ResultDstDC As Long, hResultDst As Long
  Dim ResultSrcDC As Long, hResultSrc As Long
  Dim hPrevMask As Long, hPrevInv As Long
  Dim hPrevSrc As Long, hPrevDst As Long
  Dim SrcRect As RECT
  
  With SrcRect
    .Left = SrcRectLeft
    .Top = SrcRectTop
    .Right = SrcRectRight
    .Bottom = SrcRectBottom
  End With

  w = SrcRectRight - SrcRectLeft
  h = SrcRectBottom - SrcRectTop
   
   
 
  MonoMaskDC = CreateCompatibleDC(DestDCTrans)
  MonoInvDC = CreateCompatibleDC(DestDCTrans)
  hMonoMask = CreateBitmap(w, h, 1, 1, ByVal 0&)
  hMonoInv = CreateBitmap(w, h, 1, 1, ByVal 0&)
  hPrevMask = SelectObject(MonoMaskDC, hMonoMask)
  hPrevInv = SelectObject(MonoInvDC, hMonoInv)
   
 
  ResultDstDC = CreateCompatibleDC(DestDCTrans)
  ResultSrcDC = CreateCompatibleDC(DestDCTrans)
  hResultDst = CreateCompatibleBitmap(DestDCTrans, w, h)
  hResultSrc = CreateCompatibleBitmap(DestDCTrans, w, h)
  hPrevDst = SelectObject(ResultDstDC, hResultDst)
  hPrevSrc = SelectObject(ResultSrcDC, hResultSrc)
   

  Dim OldBC As Long
  OldBC = SetBkColor(SrcDC, MaskColor)
  nRet = BitBlt(MonoMaskDC, 0, 0, w, h, SrcDC, _
                SrcRect.Left, SrcRect.Top, &HCC0020)
  MaskColor = SetBkColor(SrcDC, OldBC)
   
 
  nRet = BitBlt(MonoInvDC, 0, 0, w, h, _
                MonoMaskDC, 0, 0, &H330008)
   
 
  nRet = BitBlt(ResultDstDC, 0, 0, w, h, _
                DestDCTrans, DstX, DstY, &HCC0020)
   
 
  nRet = BitBlt(ResultDstDC, 0, 0, w, h, _
                MonoMaskDC, 0, 0, &H8800C6)
   
 
  nRet = BitBlt(ResultSrcDC, 0, 0, w, h, SrcDC, _
                SrcRect.Left, SrcRect.Top, &HCC0020)
   
 
  nRet = BitBlt(ResultSrcDC, 0, 0, w, h, _
                MonoInvDC, 0, 0, &H8800C6)
   

  nRet = BitBlt(ResultDstDC, 0, 0, w, h, _
                ResultSrcDC, 0, 0, &H660046)
   
 
  nRet = BitBlt(DestDC, DstX, DstY, w, h, _
                ResultDstDC, 0, 0, &HCC0020)
                
 
  hMonoMask = SelectObject(MonoMaskDC, hPrevMask)
  DeleteObject hMonoMask

  hMonoInv = SelectObject(MonoInvDC, hPrevInv)
  DeleteObject hMonoInv

  hResultDst = SelectObject(ResultDstDC, hPrevDst)
  DeleteObject hResultDst

  hResultSrc = SelectObject(ResultSrcDC, hPrevSrc)
  DeleteObject hResultSrc

  DeleteDC MonoMaskDC
  DeleteDC MonoInvDC
  DeleteDC ResultDstDC
  DeleteDC ResultSrcDC
End Sub

Private Sub SetAccessKeys()
Dim ampersandPos As Long
If Len(m_Caption) > 1 Then
    ampersandPos = InStr(1, m_Caption, "&", vbTextCompare)
    If (ampersandPos < Len(m_Caption)) And (ampersandPos > 0) Then
        If Mid(m_Caption, ampersandPos + 1, 1) <> "&" Then
            UserControl.AccessKeys = LCase(Mid(m_Caption, ampersandPos + 1, 1))
        Else
            ampersandPos = InStr(ampersandPos + 2, m_Caption, "&", vbTextCompare)
            If Mid(m_Caption, ampersandPos + 1, 1) <> "&" Then
                UserControl.AccessKeys = LCase(Mid(m_Caption, ampersandPos + 1, 1))
            Else
                UserControl.AccessKeys = ""
            End If
        End If
    Else
        UserControl.AccessKeys = ""
    End If
Else
    UserControl.AccessKeys = ""
End If
End Sub
