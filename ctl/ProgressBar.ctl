VERSION 5.00
Begin VB.UserControl ProgressBar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4425
   ScaleHeight     =   20
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   295
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum U_TextAlignments
    [Left Top] = 1
    [Left Middle] = 2
    [Left Bottom] = 3
    [Center Top] = 4
    [Center Middle] = 5
    [Center Bottom] = 6
    [Right Top] = 7
    [Right Middle] = 8
    [Right Bottom] = 9
End Enum

Public Enum U_TextEffects
    [Normal] = 1
    [Embossed] = 2
    [Engraved] = 3
    [Outline] = 4
    [Shadow] = 5
End Enum

Public Enum U_OrientationsS
    [Horizontal] = 1
    [Vertical] = 2

End Enum

Public Enum U_TextStyles
    [PBValue] = 1
    [PBPercentage] = 2
    [CustomText] = 3
    [PBNoneText] = 4
End Enum

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Type cRGB
    Blue As Byte
    Green As Byte
    Red As Byte
End Type

Enum U_Themes
    [IceOrange] = 1
    [IceYellow] = 2
    [IceGreen] = 3
    [IceCyan] = 4
    [IceBangel] = 5
    [IcePurple] = 6
    [IceRed] = 7
    [IceBlue] = 8
    [Vista] = 9
    [Custome] = 10
End Enum
Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Public Enum GRADIENT_DIRECT
    [Left to Right] = &H0
    [Top to Bottom] = &H1
End Enum

Private Type TRIVERTEX
    X As Long
    Y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    ALPHA As Integer
End Type

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Const GRADIENT_FILL_RECT_H As Long = &H0
Const GRADIENT_FILL_RECT_V  As Long = &H1
Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0

Private U_TextStyle As U_TextStyles
Private U_Theme As U_Themes
Private U_Orientation As U_OrientationsS
Private U_Text As String
Private U_TextColor As OLE_COLOR
Private U_TextAlign As U_TextAlignments
Private U_TextFont As Font
Private U_TextEC As OLE_COLOR
Private U_TextEffect As U_TextEffects
Private U_RoundV As Long
Private U_Min As Long
Private U_Value As Long
Private U_Max As Long
Private U_Enabled As Boolean
Private C(16) As Long
Private U_PBSCC1 As OLE_COLOR
Private U_PBSCC2 As OLE_COLOR
Private Sub UserControl_Resize()
Bar_Draw
End Sub

Public Property Let value(ByVal NewValue As Long)
Attribute value.VB_Description = "Progressbar Value."
    If NewValue > U_Max Then NewValue = U_Max
    If NewValue < U_Min Then NewValue = U_Min
    U_Value = NewValue
    
    PropertyChanged "Value"
    Bar_Draw
End Property

Public Property Get value() As Long
    value = U_Value
End Property

Public Property Let Max(ByVal NewValue As Long)
Attribute Max.VB_Description = "Progressbar Max Value."
    If NewValue < 1 Then NewValue = 1
    If NewValue <= U_Min Then NewValue = U_Min + 1
    U_Max = NewValue
    If value > U_Max Then value = U_Max
    PropertyChanged "Max"
    Bar_Draw
End Property
Public Property Get Max() As Long
    Max = U_Max
End Property

Public Property Let Min(ByVal NewValue As Long)
Attribute Min.VB_Description = "Progressbar Min Value."
    If NewValue >= U_Max Then NewValue = Max - 1
    If NewValue < 0 Then NewValue = 0
    U_Min = NewValue
    If value < U_Min Then value = U_Min
    
    PropertyChanged "Min"
    Bar_Draw
End Property
Public Property Get Min() As Long
    Min = U_Min
End Property
Public Property Get RoundedValue() As Long
Attribute RoundedValue.VB_Description = "Progressbar Rounded Corner Value."
RoundedValue = U_RoundV
End Property

Public Property Let RoundedValue(ByVal NewValue As Long)
U_RoundV = NewValue
PropertyChanged "RoundedValue"
Bar_Draw
End Property


Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Progressbar Enabled/Disabled."
Enabled = U_Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
U_Enabled = NewValue
PropertyChanged "Enabled"
Bar_Draw
End Property
Private Sub UserControl_InitProperties()
    Max = 100
    Min = 0
    value = 50
    RoundedValue = 5
    Enabled = True
    Theme = 1
    TextForeColor = vbBlack
    Text = "U11D ProgressBar"
    TextAlignment = [Center Middle]
    TextEffect = Shadow
    TextEffectColor = vbWhite
    TextStyle = CustomText
    Orientations = Horizontal
Set TextFont = Ambient.Font
End Sub
Public Property Let Theme(ByVal NewValue As U_Themes)
Attribute Theme.VB_Description = "Progressbar Styles."

    U_Theme = NewValue
    PropertyChanged "Theme"
Bar_Draw
End Property

Public Property Get Theme() As U_Themes
    Theme = U_Theme
End Property

Public Property Let TextStyle(ByVal NewValue As U_TextStyles)
Attribute TextStyle.VB_Description = "Progressbar Text Style."
    U_TextStyle = NewValue
    PropertyChanged "TextStyle"
Bar_Draw
End Property
Public Property Get TextStyle() As U_TextStyles
    TextStyle = U_TextStyle
End Property


Public Property Get Orientations() As U_OrientationsS
    Orientations = U_Orientation
End Property

Public Property Let Orientations(ByVal NewValue As U_OrientationsS)
    U_Orientation = NewValue
    PropertyChanged "Orientations"
Bar_Draw
End Property

Public Property Get TextAlignment() As U_TextAlignments
Attribute TextAlignment.VB_Description = "Progressbar Text Alignment."
TextAlignment = U_TextAlign
End Property

Public Property Let TextAlignment(ByVal NewValue As U_TextAlignments)
U_TextAlign = NewValue
PropertyChanged "TextAlignment"
Bar_Draw
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Progressbar Text."
Text = U_Text
End Property

Public Property Let Text(ByVal NewValue As String)
U_Text = NewValue
PropertyChanged "Text"
Bar_Draw
End Property
Public Property Get TextEffectColor() As OLE_COLOR
Attribute TextEffectColor.VB_Description = "Progressbar Text Effect Color."
TextEffectColor = U_TextEC
End Property

Public Property Let TextEffectColor(ByVal NewValue As OLE_COLOR)
U_TextEC = NewValue
PropertyChanged "TextEffectColor"
Bar_Draw
End Property

Public Property Get TextEffect() As U_TextEffects
Attribute TextEffect.VB_Description = "Progressbar Text Effect."
TextEffect = U_TextEffect
End Property

Public Property Let TextEffect(ByVal NewValue As U_TextEffects)
U_TextEffect = NewValue
PropertyChanged "TextEffect"
Bar_Draw
End Property

Public Property Get TextForeColor() As OLE_COLOR
Attribute TextForeColor.VB_Description = "Progressbar Text Color."
TextForeColor = U_TextColor
End Property

Public Property Let TextForeColor(ByVal NewValue As OLE_COLOR)
U_TextColor = NewValue
PropertyChanged "TextForeColor"
Bar_Draw
End Property
Public Property Get TextFont() As Font
Attribute TextFont.VB_Description = "Progressbar Text Font."
Set TextFont = U_TextFont
End Property

Public Property Set TextFont(ByVal NewValue As Font)
Set U_TextFont = NewValue
Set UserControl.Font = NewValue
PropertyChanged "TextFont"
Bar_Draw
End Property

Public Property Get PBSCustomeColor1() As OLE_COLOR
Attribute PBSCustomeColor1.VB_Description = "Progressbar Style Custome Color 1."
PBSCustomeColor1 = U_PBSCC1
End Property

Public Property Let PBSCustomeColor1(ByVal NewValue As OLE_COLOR)
U_PBSCC1 = NewValue
PropertyChanged "PBSCustomeColor1"
Bar_Draw
End Property
Public Property Get PBSCustomeColor2() As OLE_COLOR
Attribute PBSCustomeColor2.VB_Description = "Progressbar Style Custome Color 2."
PBSCustomeColor2 = U_PBSCC2
End Property

Public Property Let PBSCustomeColor2(ByVal NewValue As OLE_COLOR)
U_PBSCC2 = NewValue
PropertyChanged "PBSCustomeColor2"
Bar_Draw
End Property
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    With PropBag
    
    Max = .ReadProperty("Max", 100)
    Min = .ReadProperty("Min", 0)
    value = .ReadProperty("Value", 50)
    RoundedValue = .ReadProperty("RoundedValue", 5)
    Enabled = .ReadProperty("Enabled", True)
    Theme = .ReadProperty("Theme", 1)
    TextStyle = .ReadProperty("TextStyle", 1)
    Orientations = .ReadProperty("Orientations", Horizontal)
    Text = .ReadProperty("Text", Ambient.DisplayName)
    TextEffectColor = .ReadProperty("TextEffectColor", RGB(200, 200, 200))
    TextEffect = .ReadProperty("TextEffect", 1)
    TextAlignment = .ReadProperty("TextAlignment", 5)
    Set TextFont = .ReadProperty("TextFont", Ambient.Font)
    TextForeColor = .ReadProperty("TextForeColor", 0)
    PBSCustomeColor2 = .ReadProperty("PBSCustomeColor2", vbBlack)
    PBSCustomeColor1 = .ReadProperty("PBSCustomeColor1", vbBlack)
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
    .WriteProperty "Orientations", U_Orientation, Horizontal
    .WriteProperty "Max", U_Max, 100
    .WriteProperty "Min", U_Min, 0
    .WriteProperty "Value", U_Value, 50
    .WriteProperty "RoundedValue", U_RoundV, 5
    .WriteProperty "Enabled", U_Enabled, True
    .WriteProperty "Theme", U_Theme, 1
    .WriteProperty "TextStyle", U_TextStyle, 1
    .WriteProperty "TextFont", U_TextFont, Ambient.Font
    .WriteProperty "TextForeColor", U_TextColor, vbBlack
    .WriteProperty "TextAlignment", U_TextAlign, 5
    .WriteProperty "Text", U_Text, ""
    .WriteProperty "TextEffectColor", U_TextEC, RGB(200, 200, 200)
    .WriteProperty "TextEffect", U_TextEffect, 1
    .WriteProperty "PBSCustomeColor2", U_PBSCC2, vbBlack
    .WriteProperty "PBSCustomeColor1", U_PBSCC1, vbBlack
    End With
End Sub











Private Sub Bar_Draw()
On Error Resume Next
Dim i, S, z, Y, q As Long
Dim U_LRECT As Long

U_LRECT = CreateRoundRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, U_RoundV, U_RoundV)
SetWindowRgn UserControl.hWnd, U_LRECT, True

    i = U_Max: S = U_Value: z = U_Max
    Y = (S * 100 / z)
    q = (Y * UserControl.ScaleWidth / 100)
    
If Orientations = Vertical Then q = (Y * UserControl.ScaleHeight / 100)

CheckTheme

If Enabled = False Then
Dim II As Byte
For II = 0 To 16
    C(II) = ColourTOGray(C(II))
Next II
End If


UserControl.Cls






If U_Orientation = Horizontal Then



GradientTwoColour UserControl.hdc, [Top to Bottom], C(0), C(2), 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight / 2
GradientTwoColour UserControl.hdc, [Top to Bottom], C(4), C(6), 0, UserControl.ScaleHeight / 2, UserControl.ScaleWidth, UserControl.ScaleHeight

'DrawGradientFourColour UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight / 2, c(0), c(1), c(2), c(3)
'DrawGradientFourColour UserControl.hDC, 0, UserControl.ScaleHeight / 2, UserControl.ScaleWidth, UserControl.ScaleHeight / 2 - 1, c(4), c(5), c(6), c(7)

If value >= 1 Then

GradientTwoColour UserControl.hdc, [Top to Bottom], C(8), C(10), 0, 0, q, UserControl.ScaleHeight / 2
GradientTwoColour UserControl.hdc, [Top to Bottom], C(12), C(14), 0, UserControl.ScaleHeight / 2, q, UserControl.ScaleHeight
'DrawGradientFourColour UserControl.hDC, 0, 0, q, UserControl.ScaleHeight / 2, c(8), c(9), c(10), c(11)
'DrawGradientFourColour UserControl.hDC, 0, UserControl.ScaleHeight / 2, q, UserControl.ScaleHeight / 2 - 1, c(12), c(13), c(14), c(15)
End If



ElseIf U_Orientation = Vertical Then

GradientTwoColour UserControl.hdc, [Left to Right], C(0), C(2), 0, 0, UserControl.ScaleWidth / 2, UserControl.ScaleHeight
GradientTwoColour UserControl.hdc, [Left to Right], C(4), C(6), UserControl.ScaleWidth / 2, 0, UserControl.ScaleWidth, UserControl.ScaleHeight

'DrawGradientFourColour UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight / 2, c(0), c(1), c(2), c(3)
'DrawGradientFourColour UserControl.hDC, 0, UserControl.ScaleHeight / 2, UserControl.ScaleWidth, UserControl.ScaleHeight / 2 - 1, c(4), c(5), c(6), c(7)

If value >= 1 Then

GradientTwoColour UserControl.hdc, [Left to Right], C(8), C(10), 0, 0, UserControl.ScaleWidth / 2, q
GradientTwoColour UserControl.hdc, [Left to Right], C(12), C(14), UserControl.ScaleWidth / 2, 0, UserControl.ScaleWidth, q
'DrawGradientFourColour UserControl.hDC, 0, 0, q, UserControl.ScaleHeight / 2, c(8), c(9), c(10), c(11)
'DrawGradientFourColour UserControl.hDC, 0, UserControl.ScaleHeight / 2, q, UserControl.ScaleHeight / 2 - 1, c(12), c(13), c(14), c(15)
End If
End If




UserControl.ForeColor = C(16)
RoundRect UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, U_RoundV, U_RoundV

If TextStyle = PBValue Then
    DrawCaptionText value, U_TextAlign
ElseIf TextStyle = PBPercentage Then
    DrawCaptionText Y & "%", U_TextAlign
ElseIf TextStyle = CustomText Then
    DrawCaptionText U_Text, U_TextAlign
ElseIf TextStyle = PBNoneText Then
End If
End Sub

Private Sub CheckTheme()
If Theme = 1 Then
'BACK
C(0) = RGB(248, 246, 242)
C(1) = RGB(248, 246, 242)
C(2) = RGB(233, 227, 211)
C(3) = RGB(233, 227, 211)
'\
C(4) = RGB(226, 215, 182)
C(5) = RGB(226, 215, 182)
C(6) = RGB(239, 233, 215)
C(7) = RGB(239, 233, 215)
'FRONT
C(8) = RGB(251, 244, 223)
C(9) = RGB(251, 244, 223)
C(10) = RGB(239, 213, 133)
C(11) = RGB(239, 213, 133)
'\
C(12) = RGB(203, 166, 57)
C(13) = RGB(203, 166, 57)
C(14) = RGB(237, 224, 187)
C(15) = RGB(237, 224, 187)
'FORE COLOUR
C(16) = RGB(204, 168, 62)
ElseIf Theme = 2 Then
C(0) = RGB(228, 179, 11)
C(1) = RGB(228, 179, 11)
C(2) = RGB(228, 179, 11)
C(3) = RGB(228, 179, 11)
'\
C(4) = RGB(228, 179, 11)
C(5) = RGB(228, 179, 11)
C(6) = RGB(228, 179, 11)
C(7) = RGB(228, 179, 11)
'BACK
C(8) = RGB(245, 195, 19)
C(9) = RGB(245, 195, 19)
C(10) = RGB(245, 195, 19)
C(11) = RGB(245, 195, 19)
'\
C(12) = RGB(245, 195, 19)
C(13) = RGB(245, 195, 19)
C(14) = RGB(245, 195, 19)
C(15) = RGB(245, 195, 19)
'FRONT
'FORE COLOUR
C(16) = RGB(245, 195, 19)
ElseIf Theme = 3 Then
'BACK
C(0) = RGB(242, 248, 243)
C(1) = RGB(242, 248, 243)
C(2) = RGB(211, 233, 213)
C(3) = RGB(211, 233, 213)

'\
C(4) = RGB(182, 226, 186)
C(5) = RGB(182, 226, 186)
C(6) = RGB(215, 239, 217)
C(7) = RGB(215, 239, 217)
'FRONT
C(8) = RGB(223, 251, 225)
C(9) = RGB(223, 251, 225)
C(10) = RGB(133, 239, 142)
C(11) = RGB(133, 239, 142)
'\
C(12) = RGB(57, 203, 70)
C(13) = RGB(57, 203, 70)
C(14) = RGB(187, 237, 191)
C(15) = RGB(187, 237, 191)
'FORE COLOUR
C(16) = RGB(62, 204, 74)
ElseIf Theme = 4 Then
'BACK
C(0) = RGB(242, 248, 247)
C(1) = RGB(242, 248, 247)
C(2) = RGB(211, 233, 231)
C(3) = RGB(211, 233, 231)
'\
C(4) = RGB(182, 226, 222)
C(5) = RGB(182, 226, 222)
C(6) = RGB(215, 239, 237)
C(7) = RGB(215, 239, 237)
'FRONT
C(8) = RGB(223, 251, 249)
C(9) = RGB(223, 251, 249)
C(10) = RGB(133, 239, 230)
C(11) = RGB(133, 239, 230)
'\
C(12) = RGB(57, 203, 190)
C(13) = RGB(57, 203, 190)
C(14) = RGB(187, 237, 233)
C(15) = RGB(187, 237, 233)
'FORE COLOUR
C(16) = RGB(62, 204, 192)
ElseIf Theme = 5 Then
'BACK
C(0) = RGB(243, 242, 248)
C(1) = RGB(243, 242, 248)
C(2) = RGB(213, 211, 233)
C(3) = RGB(213, 211, 233)
'\
C(4) = RGB(186, 182, 226)
C(5) = RGB(186, 182, 226)
C(6) = RGB(217, 215, 239)
C(7) = RGB(217, 215, 239)
'FRONT
C(8) = RGB(225, 223, 251)
C(9) = RGB(225, 223, 251)
C(10) = RGB(142, 133, 239)
C(11) = RGB(142, 133, 239)
'\
C(12) = RGB(70, 57, 203)
C(13) = RGB(70, 57, 203)
C(14) = RGB(191, 187, 237)
C(15) = RGB(191, 187, 237)
'FORE COLOUR
C(16) = RGB(74, 62, 204)
ElseIf Theme = 6 Then
'BACK
C(0) = RGB(248, 242, 247)
C(1) = RGB(248, 242, 247)
C(2) = RGB(233, 211, 231)
C(3) = RGB(233, 211, 231)
'\
C(4) = RGB(226, 182, 222)
C(5) = RGB(226, 182, 222)
C(6) = RGB(239, 215, 237)
C(7) = RGB(239, 215, 237)
'FRONT
C(8) = RGB(251, 223, 249)
C(9) = RGB(251, 223, 249)
C(10) = RGB(239, 133, 230)
C(11) = RGB(239, 133, 230)
'\
C(12) = RGB(203, 57, 190)
C(13) = RGB(203, 57, 190)
C(14) = RGB(237, 187, 233)
C(15) = RGB(237, 187, 233)
'FORE COLOUR
C(16) = RGB(204, 62, 192)
ElseIf Theme = 7 Then
'BACK
C(0) = RGB(248, 242, 242)
C(1) = RGB(248, 242, 242)
C(2) = RGB(233, 211, 211)
C(3) = RGB(233, 211, 211)
'\
C(4) = RGB(226, 182, 182)
C(5) = RGB(226, 182, 182)
C(6) = RGB(239, 215, 215)
C(7) = RGB(239, 215, 215)
'FRONT
C(8) = RGB(251, 223, 223)
C(9) = RGB(251, 223, 223)
C(10) = RGB(239, 133, 133)
C(11) = RGB(239, 133, 133)
'\
C(12) = RGB(203, 57, 57)
C(13) = RGB(203, 57, 57)
C(14) = RGB(237, 187, 187)
C(15) = RGB(237, 187, 187)
'FORE COLOUR
C(16) = RGB(204, 62, 62)
ElseIf Theme = 8 Then
'BACK
C(0) = RGB(250, 253, 254)
C(1) = RGB(250, 253, 254)
C(2) = RGB(228, 243, 252)
C(3) = RGB(228, 243, 252)
'\
C(4) = RGB(199, 230, 249)
C(5) = RGB(199, 230, 249)
C(6) = RGB(237, 247, 253)
C(7) = RGB(237, 247, 253)
'FRONT
C(8) = RGB(225, 247, 255)
C(9) = RGB(225, 247, 255)
C(10) = RGB(67, 208, 255)
C(11) = RGB(67, 208, 255)
'\
C(12) = RGB(63, 112, 233)
C(13) = RGB(63, 112, 233)
C(14) = RGB(63, 226, 246)
C(15) = RGB(63, 226, 246)
'FORE COLOUR
C(16) = RGB(23, 139, 211)
ElseIf Theme = 9 Then
'BACK
C(0) = RGB(231, 243, 232)
C(1) = RGB(231, 243, 232)
C(2) = RGB(225, 219, 225)
C(3) = RGB(225, 219, 225)
'\
C(4) = RGB(179, 189, 179)
C(5) = RGB(179, 189, 179)
C(6) = RGB(226, 238, 226)
C(7) = RGB(226, 238, 226)
'FRONT
C(8) = RGB(223, 251, 223)
C(9) = RGB(223, 251, 223)
C(10) = RGB(108, 255, 108)
C(11) = RGB(108, 255, 108)
'\
C(12) = RGB(26, 228, 26)
C(13) = RGB(26, 228, 26)
C(14) = RGB(217, 244, 217)
C(15) = RGB(217, 244, 217)
'FORE COLOUR
C(16) = RGB(188, 184, 188)
ElseIf Theme = 10 Then

'BACK
C(0) = LightenColor(U_PBSCC2, 180)
C(1) = LightenColor(U_PBSCC2, 180)
C(2) = LightenColor(U_PBSCC2, 50)
C(3) = LightenColor(U_PBSCC2, 50)
'\
C(4) = U_PBSCC2
C(5) = U_PBSCC2
C(6) = LightenColor(U_PBSCC2, 80)
C(7) = LightenColor(U_PBSCC2, 80)
'FRONT
C(8) = LightenColor(U_PBSCC1, 180)
C(9) = LightenColor(U_PBSCC1, 180)
C(10) = LightenColor(U_PBSCC1, 50)
C(11) = LightenColor(U_PBSCC1, 50)
'\
C(12) = U_PBSCC1
C(13) = U_PBSCC1
C(14) = LightenColor(U_PBSCC1, 80)
C(15) = LightenColor(U_PBSCC1, 80)
'FORE COLOUR
C(16) = U_PBSCC1
End If
End Sub


























































Private Sub DrawCaptionText(ByVal TextString As String, ByVal Alignment As U_TextAlignments)
Dim lonStartWidth As Long, lonStartHeight As Long
Dim PBTCN, PBTCS As Long

If Enabled = True Then
PBTCN = U_TextColor
PBTCS = U_TextEC
Else
PBTCN = ColourTOGray(U_TextColor)
PBTCS = ColourTOGray(U_TextEC)
End If

UserControl.ForeColor = PBTCN

If Alignment = 1 Then
    lonStartWidth = 1
    lonStartHeight = 0
ElseIf Alignment = 2 Then
    lonStartWidth = 1
    lonStartHeight = (UserControl.ScaleHeight / 2) - (UserControl.TextHeight(TextString) / 2) - 1
ElseIf Alignment = 3 Then
    lonStartWidth = 1
    lonStartHeight = (UserControl.ScaleHeight - UserControl.TextHeight(TextString)) - 1

ElseIf Alignment = 4 Then
    lonStartWidth = (UserControl.ScaleWidth / 2) - (UserControl.TextWidth(TextString) / 2) - 1
    lonStartHeight = 0
ElseIf Alignment = 5 Then
    lonStartWidth = (UserControl.ScaleWidth / 2) - (UserControl.TextWidth(TextString) / 2) - 1
    lonStartHeight = (UserControl.ScaleHeight / 2) - (UserControl.TextHeight(TextString) / 2) - 1
ElseIf Alignment = 6 Then
    lonStartWidth = (UserControl.ScaleWidth / 2) - (UserControl.TextWidth(TextString) / 2) - 1
    lonStartHeight = (UserControl.ScaleHeight - UserControl.TextHeight(TextString)) - 1


ElseIf Alignment = 7 Then
    lonStartWidth = (UserControl.ScaleWidth - UserControl.TextWidth(TextString)) - 3
    lonStartHeight = 0
ElseIf Alignment = 8 Then
    lonStartWidth = (UserControl.ScaleWidth - UserControl.TextWidth(TextString)) - 3
    lonStartHeight = (UserControl.ScaleHeight / 2) - (UserControl.TextHeight(TextString) / 2) - 1
ElseIf Alignment = 9 Then
    lonStartWidth = (UserControl.ScaleWidth - UserControl.TextWidth(TextString)) - 3
    lonStartHeight = (UserControl.ScaleHeight - UserControl.TextHeight(TextString)) - 1
End If



    If U_TextEffect = Normal Then
        UserControl.CurrentX = lonStartWidth
        UserControl.CurrentY = lonStartHeight
        UserControl.Print TextString
    ElseIf U_TextEffect = Engraved Then
        UserControl.ForeColor = PBTCS
        UserControl.CurrentX = lonStartWidth + 1
        UserControl.CurrentY = lonStartHeight + 1
        UserControl.Print TextString
        UserControl.ForeColor = RGB(128, 128, 128)
        UserControl.CurrentX = lonStartWidth - 1
        UserControl.CurrentY = lonStartHeight
        UserControl.Print TextString
        UserControl.ForeColor = PBTCN
        UserControl.CurrentX = lonStartWidth
        UserControl.CurrentY = lonStartHeight
        UserControl.Print TextString
        
    ElseIf U_TextEffect = Embossed Then
        UserControl.ForeColor = PBTCS
        UserControl.CurrentX = lonStartWidth - 1
        UserControl.CurrentY = lonStartHeight - 1
        UserControl.Print TextString
        UserControl.ForeColor = RGB(128, 128, 128)
        UserControl.CurrentX = lonStartWidth + 1
        UserControl.CurrentY = lonStartHeight + 1
        UserControl.Print TextString
        UserControl.ForeColor = PBTCN
        UserControl.CurrentX = lonStartWidth
        UserControl.CurrentY = lonStartHeight
        UserControl.Print TextString
    ElseIf U_TextEffect = Outline Then
        UserControl.ForeColor = PBTCS
        UserControl.CurrentX = lonStartWidth + 1
        UserControl.CurrentY = lonStartHeight
        UserControl.Print TextString
        UserControl.CurrentX = lonStartWidth - 1
        UserControl.CurrentY = lonStartHeight
        UserControl.Print TextString
        UserControl.CurrentY = lonStartHeight - 1
        UserControl.CurrentX = lonStartWidth
        UserControl.Print TextString
        UserControl.CurrentY = lonStartHeight + 1
        UserControl.CurrentX = lonStartWidth
        UserControl.Print TextString
        UserControl.ForeColor = PBTCN
        UserControl.CurrentX = lonStartWidth
        UserControl.CurrentY = lonStartHeight
        UserControl.Print TextString
        
    ElseIf U_TextEffect = Shadow Then
        UserControl.ForeColor = PBTCS
        UserControl.CurrentX = lonStartWidth + 1
        UserControl.CurrentY = lonStartHeight + 1
        UserControl.Print TextString
        UserControl.ForeColor = PBTCN
        UserControl.CurrentX = lonStartWidth
        UserControl.CurrentY = lonStartHeight
        UserControl.Print TextString
    End If


End Sub

Public Function DrawGradientFourColour(ObjectHDC As Long, Left As Long, Top As Long, Width As Long, Height As Long, TopLeftColour As Long, TopRightColour As Long, BottomLeftColour As Long, BottomRightColour As Long)
    Dim bi24BitInfo     As BITMAPINFO
    Dim bBytes()        As Byte
    Dim LeftGrads()     As cRGB
    Dim RightGrads()    As cRGB
    Dim MiddleGrads()   As cRGB
    Dim TopLeft         As cRGB
    Dim TopRight        As cRGB
    Dim BottomLeft      As cRGB
    Dim BottomRight     As cRGB
    Dim iLoop           As Long
    Dim bytesWidth      As Long
    
    With TopLeft
        .Red = Red(TopLeftColour)
        .Green = Green(TopLeftColour)
        .Blue = Blue(TopLeftColour)
    End With
    
    With TopRight
        .Red = Red(TopRightColour)
        .Green = Green(TopRightColour)
        .Blue = Blue(TopRightColour)
    End With
    
    With BottomLeft
        .Red = Red(BottomLeftColour)
        .Green = Green(BottomLeftColour)
        .Blue = Blue(BottomLeftColour)
    End With
    
    With BottomRight
        .Red = Red(BottomRightColour)
        .Green = Green(BottomRightColour)
        .Blue = Blue(BottomRightColour)
    End With
    
    GradateColours LeftGrads, Height, TopLeft, BottomLeft
    GradateColours RightGrads, Height, TopRight, BottomRight
    
    With bi24BitInfo.bmiHeader
        .biBitCount = 24
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(bi24BitInfo.bmiHeader)
        .biWidth = Width
        .biHeight = 1
    End With
    
    ReDim bBytes(1 To bi24BitInfo.bmiHeader.biWidth * bi24BitInfo.bmiHeader.biHeight * 3) As Byte
    
    bytesWidth = (Width) * 3
    
    For iLoop = 0 To Height - 1
        GradateColours MiddleGrads, Width, LeftGrads(iLoop), RightGrads(iLoop)
        CopyMemory bBytes(1), MiddleGrads(0), bytesWidth
        SetDIBitsToDevice ObjectHDC, Left, Top + iLoop, bi24BitInfo.bmiHeader.biWidth, bi24BitInfo.bmiHeader.biHeight, 0, 0, 0, bi24BitInfo.bmiHeader.biHeight, bBytes(1), bi24BitInfo, DIB_RGB_COLORS
    Next iLoop
    
    
End Function

Private Function GradateColours(cResults() As cRGB, Length As Long, Colour1 As cRGB, Colour2 As cRGB)
    Dim fromR   As Integer
    Dim toR     As Integer
    Dim fromG   As Integer
    Dim toG     As Integer
    Dim fromB   As Integer
    Dim toB     As Integer
    Dim stepR   As Single
    Dim stepG   As Single
    Dim stepB   As Single
    Dim iLoop   As Long
    
    ReDim cResults(0 To Length)
    
    fromR = Colour1.Red
    fromG = Colour1.Green
    fromB = Colour1.Blue
    
    toR = Colour2.Red
    toG = Colour2.Green
    toB = Colour2.Blue
    
    stepR = Divide(toR - fromR, Length)
    stepG = Divide(toG - fromG, Length)
    stepB = Divide(toB - fromB, Length)
    
    For iLoop = 0 To Length
        cResults(iLoop).Red = fromR + (stepR * iLoop)
        cResults(iLoop).Green = fromG + (stepG * iLoop)
        cResults(iLoop).Blue = fromB + (stepB * iLoop)
    Next iLoop
End Function

Private Function Blue(Colour As Long) As Long
    Blue = (Colour And &HFF0000) / &H10000
End Function
Private Function Green(Colour As Long) As Long
    Green = (Colour And &HFF00&) / &H100
End Function

Private Function Red(Colour As Long) As Long
    Red = (Colour And &HFF&)
End Function

Private Function Divide(Numerator, Denominator) As Single
    If Numerator = 0 Or Denominator = 0 Then
        Divide = 0
    Else
        Divide = Numerator / Denominator
    End If
End Function
Public Sub GradientTwoColour(ByVal hdc As Long, ByVal Direction As GRADIENT_DIRECT, ByVal StartColor As Long, ByVal EndColor As Long, Left As Long, Top As Long, Width As Long, Height As Long)
Dim udtVert(1) As TRIVERTEX, udtGRect As GRADIENT_RECT
Dim UDTRECT As RECT
'hDCObj.ScaleMode = vbPixels
'hDCObj.AutoRedraw = True
SetRect UDTRECT, Left, Top, Width, Height
With udtVert(0)
    .X = UDTRECT.Left
    .Y = UDTRECT.Top
    .Red = LongToSignedShort(CLng((StartColor And &HFF&) * 256))
    .Green = LongToSignedShort(CLng(((StartColor And &HFF00&) \ &H100&) * 256))
    .Blue = LongToSignedShort(CLng(((StartColor And &HFF0000) \ &H10000) * 256))
    .ALPHA = 0&
End With

With udtVert(1)
    .X = UDTRECT.Right
    .Y = UDTRECT.Bottom
    .Red = LongToSignedShort(CLng((EndColor And &HFF&) * 256))
    .Green = LongToSignedShort(CLng(((EndColor And &HFF00&) \ &H100&) * 256))
    .Blue = LongToSignedShort(CLng(((EndColor And &HFF0000) \ &H10000) * 256))
    .ALPHA = 0&
End With

udtGRect.UpperLeft = 0
udtGRect.LowerRight = 1

GradientFillRect hdc, udtVert(0), 2, udtGRect, 1, Direction
End Sub


Private Function LongToSignedShort(ByVal Unsigned As Long) As Integer
If Unsigned < 32768 Then
    LongToSignedShort = CInt(Unsigned)
Else
    LongToSignedShort = CInt(Unsigned - &H10000)
End If
End Function


Private Function ColourTOGray(ByVal uColor As Long) As Long
Dim Red As Long, Blue As Long, Green As Long
Dim Gray As Long
    Red = uColor Mod 256
    Green = (uColor Mod 65536) / 256
    Blue = uColor / 65536
    Gray = (Red + Green + Blue) / 3
    ColourTOGray = RGB(Gray, Gray, Gray)
End Function
Private Function LightenColor(ByVal uColour As ColorConstants, Optional ByVal offset As Long = 1) As Long
Dim intR As Integer, intG As Integer, intB As Integer
intR = Abs((uColour Mod 256) + offset)
intG = Abs((((uColour And &HFF00) / 256&) Mod 256&) + offset)
intB = Abs(((uColour And &HFF0000) / 65536) + offset)

LightenColor = RGB(intR, intG, intB)
End Function
