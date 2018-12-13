VERSION 5.00
Begin VB.UserControl AkuSayangIbu 
   AutoRedraw      =   -1  'True
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2565
   DefaultCancel   =   -1  'True
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Picture         =   "AkuSayangIbu.ctx":0000
   ScaleHeight     =   390
   ScaleWidth      =   2565
   Begin VB.Timer Timer1 
      Left            =   840
      Top             =   0
   End
End
Attribute VB_Name = "AkuSayangIbu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'Autor: Leandro Ascierto
'Web:   www.leandroascierto.com.ar
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GdiTransparentBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Const IDC_HAND = 32649&

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const DT_SINGLELINE As Long = &H20
Private Const DT_VCENTER As Long = &H4
Private Const DT_CENTER As Long = &H1

Public Enum eBtnStyle
    BtnGrey = 0
    BtnBlue = 1
    BtnGreen = 2
    BtnFlat = 3
End Enum

Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim BtnState As Long
Dim m_BtnStyle As eBtnStyle
Dim m_Caption As String
Dim m_HaveFocus As Boolean
Dim hCursorHands As Long
Dim m_Icon As StdPicture


Private Function RenderStretchFromDC(ByVal destDC As Long, _
                                ByVal destX As Long, _
                                ByVal destY As Long, _
                                ByVal DestW As Long, _
                                ByVal DestH As Long, _
                                ByVal SrcDC As Long, _
                                ByVal X As Long, _
                                ByVal Y As Long, _
                                ByVal Width As Long, _
                                ByVal Height As Long, _
                                ByVal Size As Long, _
                                Optional MaskColor As Long = -1)
 
Dim Sx2 As Long
 
Sx2 = Size * 2
 
If MaskColor <> -1 Then
    Dim mDC         As Long
    Dim mX          As Long
    Dim mY          As Long
    Dim DC          As Long
    Dim hBmp        As Long
    Dim hOldBmp     As Long
 
    mDC = destDC
    DC = GetDC(0)
    destDC = CreateCompatibleDC(0)
    hBmp = CreateCompatibleBitmap(DC, DestW, DestH)
    hOldBmp = SelectObject(destDC, hBmp) ' save the original BMP for later reselection
    mX = destX: mY = destY
    destX = 0: destY = 0
End If
 
SetStretchBltMode destDC, vbPaletteModeNone
 
BitBlt destDC, destX, destY, Size, Size, SrcDC, X, Y, vbSrcCopy  'TOP_LEFT
StretchBlt destDC, destX + Size, destY, DestW - Sx2, Size, SrcDC, X + Size, Y, Width - Sx2, Size, vbSrcCopy 'TOP_CENTER
BitBlt destDC, destX + DestW - Size, destY, Size, Size, SrcDC, X + Width - Size, Y, vbSrcCopy 'TOP_RIGHT
StretchBlt destDC, destX, destY + Size, Size, DestH - Sx2, SrcDC, X, Y + Size, Size, Height - Sx2, vbSrcCopy 'MID_LEFT
StretchBlt destDC, destX + Size, destY + Size, DestW - Sx2, DestH - Sx2, SrcDC, X + Size, Y + Size, Width - Sx2, Height - Sx2, vbSrcCopy 'MID_CENTER
StretchBlt destDC, destX + DestW - Size, destY + Size, Size, DestH - Sx2, SrcDC, X + Width - Size, Y + Size, Size, Height - Sx2, vbSrcCopy 'MID_RIGHT
BitBlt destDC, destX, destY + DestH - Size, Size, Size, SrcDC, X, Y + Height - Size, vbSrcCopy 'BOTTOM_LEFT
StretchBlt destDC, destX + Size, destY + DestH - Size, DestW - Sx2, Size, SrcDC, X + Size, Y + Height - Size, Width - Sx2, Size, vbSrcCopy   'BOTTOM_CENTER
BitBlt destDC, destX + DestW - Size, destY + DestH - Size, Size, Size, SrcDC, X + Width - Size, Y + Height - Size, vbSrcCopy 'BOTTOM_RIGHT

If MaskColor <> -1 Then
    GdiTransparentBlt mDC, mX, mY, DestW, DestH, destDC, 0, 0, DestW, DestH, MaskColor
    SelectObject destDC, hOldBmp
    DeleteObject hBmp
    ReleaseDC 0&, DC
    DeleteDC destDC
End If

End Function
 
 
Private Function RenderStretchFromPicture(ByVal destDC As Long, _
                                ByVal destX As Long, _
                                ByVal destY As Long, _
                                ByVal DestW As Long, _
                                ByVal DestH As Long, _
                                ByVal SrcPicture As StdPicture, _
                                ByVal X As Long, _
                                ByVal Y As Long, _
                                ByVal Width As Long, _
                                ByVal Height As Long, _
                                ByVal Size As Long, _
                                Optional MaskColor As Long = -1)
 
    Dim DC          As Long
    Dim hOldBmp    As Long
 
    DC = CreateCompatibleDC(0)
    hOldBmp = SelectObject(DC, SrcPicture.handle)
 
    RenderStretchFromDC destDC, destX, destY, DestW, DestH, DC, X, Y, Width, Height, Size, MaskColor
 
    hOldBmp = SelectObject(DC, hOldBmp)
    DeleteDC DC
End Function

Public Property Get BtnStyle() As eBtnStyle
    BtnStyle = m_BtnStyle
End Property

Public Property Let BtnStyle(ByVal NewValue As eBtnStyle)
    m_BtnStyle = NewValue
    UserControl.Font.Size = IIf(m_BtnStyle = BtnFlat, 8, 9)
    Draw
    PropertyChanged "BtnStyle"
End Property

Public Property Get Icon() As StdPicture
    Set Icon = m_Icon
End Property

Public Property Set Icon(ByVal NewValue As StdPicture)
    Set m_Icon = NewValue
    Draw
    PropertyChanged "Icon"
End Property



Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal NewValue As String)
    m_Caption = NewValue
    Draw
    PropertyChanged "Caption"
End Property


Private Sub UserControl_DblClick()
    BtnState = 5
    Draw
End Sub


Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
  RaiseEvent Click
End Sub

Private Sub UserControl_ExitFocus()
    m_HaveFocus = False
    Draw
End Sub

Private Sub UserControl_GotFocus()
    m_HaveFocus = True
    Draw
End Sub

Private Sub UserControl_Initialize()
    hCursorHands = LoadCursor(0&, IDC_HAND)
End Sub

Private Sub UserControl_InitProperties()
    m_Caption = Ambient.DisplayName
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If hCursorHands Then SetCursor hCursorHands
    BtnState = 5
    Draw
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If hCursorHands Then SetCursor hCursorHands
    If m_BtnStyle = BtnFlat And Timer1.Interval = 0 Then
        Timer1.Interval = 100
        Draw
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If hCursorHands Then SetCursor hCursorHands
    BtnState = 0
    Draw
    If (X > 0) And (Y > 0) And (X < UserControl.ScaleWidth) And (Y < UserControl.ScaleHeight) Then
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_Resize()
    Draw
End Sub

Private Sub UserControl_Show()
    Draw
End Sub

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
    UserControl.Enabled = Value
    Draw
    PropertyChanged "Enabled"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_BtnStyle = .ReadProperty("BtnStyle", 0)
        Me.Enabled = .ReadProperty("Enabled", True)
        m_Caption = .ReadProperty("Caption", Ambient.DisplayName)
        Set m_Icon = .ReadProperty("Icon", Nothing)
    End With
    UserControl.Font.Size = IIf(m_BtnStyle = BtnFlat, 8, 9)
    Draw
End Sub

Private Function pvIsCursorInUC() As Boolean
    Dim PT As POINTAPI
    GetCursorPos PT
    pvIsCursorInUC = (WindowFromPoint(PT.X, PT.Y) = UserControl.hWnd)
End Function

Private Sub UserControl_Terminate()
    DestroyCursor hCursorHands
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "BtnStyle", m_BtnStyle, 0
        .WriteProperty "Enabled", UserControl.Enabled, True
        .WriteProperty "Caption", m_Caption, Ambient.DisplayName
        .WriteProperty "Icon", m_Icon, Nothing
 
    End With
End Sub

Private Sub Draw()
    Dim IconX As Long
    Dim IconY As Long
    Dim Rec As RECT
    
    If m_BtnStyle = BtnFlat Then
    
        UserControl.ForeColor = vbWhite
        
        If pvIsCursorInUC Then
            UserControl.FontUnderline = True
            If BtnState = 5 Then
                UserControl.FillColor = &HA1674B
            Else
                UserControl.FillColor = &HB7866D
            End If
        Else
            UserControl.FontUnderline = False
            UserControl.FillColor = &HAD7A62
        End If
        
        UserControl.Line (0, 0)-(UserControl.ScaleWidth - 10, UserControl.ScaleHeight - 10), &H88401D, B
    
        Rec.Right = UserControl.ScaleWidth / Screen.TwipsPerPixelX
        Rec.Bottom = UserControl.ScaleHeight / Screen.TwipsPerPixelY
    
    
        If Not m_Icon Is Nothing Then
            IconX = ScaleX(m_Icon.Width, vbHimetric, vbPixels)
            IconY = (Rec.Bottom / 2) - (ScaleY(m_Icon.Height, vbHimetric, vbPixels) / 2)
            Rec.Left = IconX + 16
            Call RenderStdPicture(UserControl.hdc, m_Icon, 8, IconY)
            DrawText UserControl.hdc, m_Caption, -1, Rec, DT_SINGLELINE Or DT_VCENTER
        Else
            DrawText UserControl.hdc, m_Caption, -1, Rec, DT_SINGLELINE Or DT_VCENTER Or DT_CENTER
        
        End If
    

        
        
        UserControl.Refresh
        Exit Sub
    End If

    If UserControl.Enabled Then
        
        If m_BtnStyle = BtnBlue Or m_BtnStyle = BtnGreen Then
            UserControl.ForeColor = vbWhite
        Else
            UserControl.ForeColor = vbBlack
        End If
        
        
        RenderStretchFromPicture UserControl.hdc, 0, 0, UserControl.ScaleWidth / Screen.TwipsPerPixelX, UserControl.ScaleHeight / Screen.TwipsPerPixelY, UserControl.Picture, BtnState + (m_BtnStyle * 10), 0, 5, 26, 2
    Else
        UserControl.ForeColor = &HB8B8B8
        RenderStretchFromPicture UserControl.hdc, 0, 0, UserControl.ScaleWidth / Screen.TwipsPerPixelX, UserControl.ScaleHeight / Screen.TwipsPerPixelY, UserControl.Picture, 30, 0, 5, 26, 2
    End If
    
    Rec.Right = UserControl.ScaleWidth / Screen.TwipsPerPixelX
    Rec.Bottom = UserControl.ScaleHeight / Screen.TwipsPerPixelY
    
    DrawText UserControl.hdc, m_Caption, -1, Rec, DT_SINGLELINE Or DT_VCENTER Or DT_CENTER
    
    If m_HaveFocus Then
        With Rec
            .Left = 2
            .Top = 2
            .Right = (UserControl.ScaleWidth / Screen.TwipsPerPixelX) - 2
            .Bottom = (UserControl.ScaleHeight / Screen.TwipsPerPixelY) - 3
        End With
        DrawFocusRect UserControl.hdc, Rec
    End If
    UserControl.Refresh
End Sub

Private Sub Timer1_Timer()
    If pvIsCursorInUC = False Then
        Timer1.Interval = 0
        Draw
    End If
End Sub

Private Function RenderStdPicture(theTarget As Variant, thePic As StdPicture, _
                             Optional ByVal destX As Long, Optional ByVal destY As Long, _
                             Optional ByVal destWidth As Long, Optional ByVal destHeight As Long, _
                             Optional ByVal SrcX As Long, Optional ByVal SrcY As Long, _
                             Optional ByVal srcWidth As Long, Optional ByVal srcHeight As Long, _
                             Optional ByVal ParamScaleMode As ScaleModeConstants = vbUser, _
                             Optional ByVal Centered As Boolean = False, Optional ByVal ZoomFactor As Single = 1&) As Boolean
                                    
    ' Return Value [out]
    '   If no errors occur, return value is True. If error or invalid parameters passed, value is False
    ' Parameters [in]
    '   theTarget: a VB form, picturebox, usercontrol or a valid hDC (no error checking for valid DC)
    '       ... If Object, then it must expose a ScaleMode and hDC property
    '       ... and if centering and an object, must also expose ScaleWidth & ScaleHeight properties
    '   thePic: a VB ImageControl, stdPicture object, or VB .Picture property
    '   destX: horizontal offset on theTarget where drawing begins, default is zero
    '   destY: vertical offset on theTarget where drawing begins, default is zero
    '   destWidth: rendered image width & will be multiplied against ZoomFactor; default is thePic.Width
    '   destHeight: rendered image height & will be multiplied against ZoomFactor; default is thePic.Height
    '   srcX: horizontal offset of thePic to begin rendering from; default is zero
    '   srcY: vertical offset of thePic to begin rendering from; default is zero
    '   srcWidth: thePic width that will be rendered; default is thePic.Width
    '   srcHeight: thePic height that will be rendered; default is thePic.Height
    '   ParamScaleMode: Scalemode for passed parameters.
    '       If vbUser, then theTarget scalemode is used if theTarget is an Object else vbPixels if theTarget is an hDC
    '   Centered: If True, rendered image is centered in theTarget, offset by destX and/or destY
    '       If theTarget is a DC, then Centered is ignored. You must pass the correct destX,destY values
    '   ZoomFactor: Scaling option. Values>1 zoom out and Values<1||>0 zoom in
    
    ' Tip: To stretch image to a picturebox dimensions, pass destWidth & destHeight
    '   as the picturebox's scalewidth & scaleheight respectively and ZoomFactor of 1
                                    
    If thePic Is Nothing Then Exit Function                 ' sanity checks first
    If thePic.handle = 0& Then Exit Function
    If ZoomFactor <= 0! Then Exit Function
    
    Dim Width As Long, Height As Long, destDC As Long
    
    ' the stdPicture.Render method requires vbPixels for destination and vbHimetrics for source
    Width = ScaleX(thePic.Width, vbHimetric, vbPixels)      ' image size in pixels
    Height = ScaleY(thePic.Height, vbHimetric, vbPixels)
    
    On Error Resume Next
    If IsObject(theTarget) Then         ' passed object? If so, set scalemode if needed
        If theTarget Is Nothing Then Exit Function
        If ParamScaleMode = vbUser Then ParamScaleMode = theTarget.ScaleMode
        destDC = theTarget.hdc
    ElseIf IsNumeric(theTarget) Then    ' passed hDC? If so, set scalemode if needed
        If ParamScaleMode = vbUser Then ParamScaleMode = vbPixels
        destDC = Val(theTarget)
        Centered = False                ' only applicable if theTarget is a VB object
    Else
        Exit Function                   ' unhandled; abort
    End If
    If Err Then                         ' checks above generated an error; probably passing object without scalemode property?
        Err.Clear
        Exit Function
    End If
 
    If destWidth Then                   ' calculate destination width in pixels from ParamScaleMode
        destWidth = ScaleX(destWidth, ParamScaleMode, vbPixels) * ZoomFactor
    Else
        destWidth = Width * ZoomFactor
    End If
    If destHeight Then                  'calculate destination height in pixels from ParamScaleMode
        destHeight = ScaleY(destHeight, ParamScaleMode, vbPixels) * ZoomFactor
    Else
        destHeight = Height * ZoomFactor
    End If
                                        ' get destX,destY in pixels from ParamScaleMode
    If destX Then destX = ScaleX(destX, ParamScaleMode, vbPixels)
    If destY Then destY = ScaleY(destY, ParamScaleMode, vbPixels)
    If Centered Then                    ' Offset destX,destY if centering
        destX = (ScaleX(theTarget.ScaleWidth, theTarget.ScaleMode, vbPixels) - destWidth) / 2 + destX
        destY = (ScaleY(theTarget.ScaleHeight, theTarget.ScaleMode, vbPixels) - destHeight) / 2 + destY
    End If
                                        ' setup source coords/bounds and convert to vbHimetrics
    If SrcX Then SrcX = ScaleX(SrcX, ParamScaleMode, vbHimetric)
    If SrcY Then SrcY = ScaleY(SrcY, ParamScaleMode, vbHimetric)
    If srcWidth Then srcWidth = ScaleX(srcWidth, ParamScaleMode, vbHimetric) Else srcWidth = thePic.Width
    If srcHeight Then srcHeight = ScaleY(srcHeight, ParamScaleMode, vbHimetric) Else srcHeight = thePic.Height
    
    If Err Then                         ' passed bad parameters or
        Err.Clear                       ' passed object that has no ScaleMode property (i.e., VB Frame)
    Else
        With thePic                     ' render, the "Or 0&" below are required else mismatch errors occur
            .Render destDC Or 0&, destX Or 0&, destY Or 0&, destWidth Or 0&, destHeight Or 0&, _
                SrcX Or 0&, .Height - SrcY, srcWidth Or 0&, -srcHeight, ByVal 0&
        End With                        ' return success/failure
        If Err Then Err.Clear Else RenderStdPicture = True
    End If
    On Error GoTo 0
    
End Function




