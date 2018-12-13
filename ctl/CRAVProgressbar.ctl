VERSION 5.00
Begin VB.UserControl CRAVProgressBar 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2535
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   49
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   169
   ToolboxBitmap   =   "CRAVProgressbar.ctx":0000
   Begin VB.PictureBox picProcess 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   840
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   600
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   3
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   48
   End
   Begin VB.PictureBox picBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   360
      Picture         =   "CRAVProgressbar.ctx":0314
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   4
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.PictureBox picBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   120
      Picture         =   "CRAVProgressbar.ctx":04B4
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   90
   End
End
Attribute VB_Name = "CRAVProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CRAVProgressBar Control
'
'Author Candra Ramadhan P.
Option Explicit

' Public Event
Public Event Change()

' Public Enumerations
Public Enum Alignments
   [Left Justify]
   [Right Justify]
   Center
End Enum

Public Enum BarColors
   Default
   Red
   Yellow
   Green
   Cyan
   Blue
   Magenta
End Enum

Public Enum Orientations
   Horizontal
   Vertical
End Enum

' Private Type
Private Type PointAPI
   X                    As Long
   Y                    As Long
End Type

' Private Variables
Private m_Alignment     As Alignments
Private m_BarColor      As BarColors
Private m_BorderDefault As Boolean
Private IsClearing      As Boolean
Private m_Value         As Integer
Private m_ForeColor     As Long
Private m_BorderColor   As OLE_COLOR
Private m_Orientation   As Orientations
Private Step            As Single
Private m_Caption       As String

' Private API
Private Declare Function PlgBlt Lib "GDI32" (ByVal hDCDest As Long, lpPoint As PointAPI, ByVal hDCSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hbmMask As Long, ByVal xMask As Long, ByVal yMask As Long) As Long

Public Property Get Alignment() As Alignments
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."

   Alignment = m_Alignment

End Property

Public Property Let Alignment(ByVal NewAlignment As Alignments)

   m_Alignment = NewAlignment
   PropertyChanged "Alignment"
   
   Call Paint

End Property

Public Property Get BarColor() As BarColors
Attribute BarColor.VB_Description = "Returns/sets the bar color of the progressbar."

   BarColor = m_BarColor

End Property

Public Property Let BarColor(ByVal NewBarColor As BarColors)

   If Not UserControl.Enabled Then Exit Property
   
   m_BarColor = NewBarColor
   PropertyChanged "BarColor"
   
   Call ChangeColor
   Call Paint

End Property

Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the color of an object's border."

   BorderColor = m_BorderColor

End Property

Public Property Let BorderColor(ByVal NewBorderColor As OLE_COLOR)

   If Not UserControl.Enabled Then Exit Property
   
   m_BorderColor = NewBorderColor
   PropertyChanged "BorderColor"
   
   Call DrawBorder
   Call Paint

End Property

Public Property Get BorderDefault() As Boolean
Attribute BorderDefault.VB_Description = "Returns/sets the allow to change the border color. (If True the default border color will be set.)"

   BorderDefault = m_BorderDefault

End Property

Public Property Let BorderDefault(ByVal NewBorderDefault As Boolean)

   If Not UserControl.Enabled Then Exit Property
   If NewBorderDefault Then BorderColor = picBorder.Point(0, 0)
   
   m_BorderDefault = NewBorderDefault
   PropertyChanged "BorderDefault"
   
   Call DrawBorder
   Call Paint

End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."

   Caption = m_Caption

End Property

Public Property Let Caption(ByVal NewCaption As String)

   If Not IsClearing And Not UserControl.Enabled Then Exit Property
   
   m_Caption = NewCaption
   PropertyChanged "Caption"
   
   Call Paint

End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines wheter an object can respond to user-generated events."

   Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(ByVal NewEnabled As Boolean)

   UserControl.Enabled = NewEnabled
   PropertyChanged "Enabled"
   
   Call ChangeColor
   Call DrawBorder
   Call Paint

End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns a Font object."

   Set Font = picProcess.Font

End Property

Public Property Let Font(ByRef NewFont As StdFont)

   Set Font = NewFont

End Property

Public Property Set Font(ByRef NewFont As StdFont)

   If NewFont Is Nothing Then Set NewFont = picProcess.Font
   
   Set picProcess.Font = NewFont
   PropertyChanged "Font"
   
   Call Paint

End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display percentage text in the progressbar."

   ForeColor = m_ForeColor

End Property

Public Property Let ForeColor(ByVal NewForeColor As OLE_COLOR)

   If Not UserControl.Enabled Then Exit Property
   
   m_ForeColor = NewForeColor
   PropertyChanged "ForeColor"
   
   Call Paint

End Property

Public Property Get Orientation() As Orientations
Attribute Orientation.VB_Description = "Returns/sets the orientation of an object."

   Orientation = m_Orientation

End Property

Public Property Let Orientation(ByVal NewOrientation As Orientations)

   m_Orientation = NewOrientation
   PropertyChanged "Orientation"
   
   Call ChangeOrientation

End Property

Public Property Get Value() As Integer
Attribute Value.VB_Description = "Returns/sets the value of an object."

   Value = m_Value

End Property

Public Property Let Value(ByVal NewValue As Integer)

Dim blnChanged As Boolean

   If Not IsClearing Then
      If Not UserControl.Enabled Then Exit Property
      If NewValue < 0 Then NewValue = 0
      If NewValue > 100 Then NewValue = 100
      
      blnChanged = (m_Value <> NewValue)
   End If
   
   m_Value = NewValue
   PropertyChanged "Value"
   
   Call Paint
   
   If blnChanged And UserControl.Enabled Then RaiseEvent Change

End Property

' clears the progressbar
Public Sub Clear()

   IsClearing = True
   Value = 0
   Caption = ""
   IsClearing = False

End Sub

' paints the progressbar
Public Sub Paint()

Dim lngColor As Long
Dim sngX     As Single
Dim sngY     As Single

   With picProcess
      .Cls
      
      If m_Value = 0 Then
         Call DrawBorder
         
         Exit Sub
      End If
      
      .PaintPicture picBuffer.Picture, 1, 1, m_Value * Step, .ScaleHeight - 2, 1, 0, 1, picBuffer.ScaleHeight, vbSrcCopy
      .PaintPicture picBuffer.Picture, 1, 1, 1, .ScaleHeight - 2, 0, 0, 1, picBuffer.ScaleHeight, vbSrcCopy
      .PaintPicture picBuffer.Picture, m_Value * Step, 1, 1, .ScaleHeight - 2, 0, 0, 1, picBuffer.ScaleHeight, vbSrcCopy
      
      If UserControl.Enabled Then
         lngColor = m_ForeColor
         
      Else
         lngColor = &HC0C0C0
      End If
      
      If m_Alignment = [Left Justify] Then
         sngX = 5
         
      ElseIf m_Alignment = [Right Justify] Then
         sngX = .ScaleWidth - .TextWidth(m_Caption) - 5
         
      Else
         sngX = (.ScaleWidth - .TextWidth(m_Caption)) / 2
      End If
      
      sngY = (.ScaleHeight - .TextHeight(m_Caption)) / 2
      .ForeColor = lngColor And &H404040
      .CurrentX = sngX
      .CurrentY = sngY
      picProcess.Print m_Caption
      .ForeColor = lngColor
      .CurrentX = sngX - 1
      .CurrentY = sngY - 1
      picProcess.Print m_Caption
      DoEvents
      
      Call MakeBar(.Image)
   End With

End Sub

Private Sub ChangeColor()

Dim intX     As Integer
Dim intY     As Integer
Dim lngColor As Long

   If Not UserControl.Enabled Or (m_BarColor = Default) Then
      lngColor = vbWhite
      
   Else
      lngColor = Choose(m_BarColor, vbRed, vbYellow, vbGreen, vbCyan, vbBlue, vbMagenta)
   End If
   
   With picBuffer
      .Picture = Nothing
      
      If (m_BarColor = Default) And UserControl.Enabled Then
         For intY = 0 To .Height
            For intX = 2 To 3
               picBuffer.PSet (intX - 2, intY), picBar.Point(intX, intY)
            Next 'intX
         Next 'intY
         
      Else
         For intY = 0 To .Height
            For intX = 0 To 1
               picBuffer.PSet (intX, intY), picBar.Point(intX, intY) And lngColor
            Next 'intX
         Next 'intY
      End If
      
      .Picture = .Image
   End With

End Sub

Private Sub ChangeOrientation()

Dim lngTemp As Long

   If m_Orientation = Horizontal Then
      If Height > Width Then
         lngTemp = Height
         Height = Width
         Width = lngTemp
      End If
      
   ElseIf Width > Height Then
      lngTemp = Height
      Height = Width
      Width = lngTemp
   End If

End Sub

Private Sub DrawBorder()

Dim intPointer As Integer

   With picProcess
      Step = (.ScaleWidth - 2) / 100
      intPointer = 3 And Not UserControl.Enabled
      .PaintPicture picBorder.Picture, 0, 0, 2, .ScaleHeight, intPointer, 0, 2, .ScaleHeight, vbSrcCopy
      .PaintPicture picBorder.Picture, 2, 0, .ScaleWidth - 4, .ScaleHeight, 2 + intPointer, 0, 1, .ScaleHeight, vbSrcCopy
      .PaintPicture picBorder.Picture, .ScaleWidth - 1, 0, -2, .ScaleHeight, 0, 0, 2, .ScaleHeight, vbSrcCopy
      
      If Not m_BorderDefault Then picProcess.Line (0, 0)-(.ScaleWidth - 1, .ScaleHeight - 1), m_BorderColor, B
      
      .Picture = .Image
      
      Call MakeBar(.Picture)
   End With

End Sub

Private Sub MakeBar(ByRef BarImage As Object)

Const PI_PART    As Double = 1.74532925132222E-02

Dim intIndex     As Integer
Dim lngDstWidth  As Long
Dim lngDstHeight As Long
Dim lngSrcWidth  As Long
Dim lngSrcHeight As Long
Dim ptaTemp(3)   As PointAPI
Dim ptaBuffer(3) As PointAPI
Dim sngCosine    As Single
Dim sngHeight    As Single
Dim sngSine      As Single
Dim sngWidth     As Single

   If m_Orientation = Horizontal Then
      Picture = BarImage
      
   Else
      With picProcess
         lngDstWidth = .ScaleWidth / 2
         lngDstHeight = .ScaleHeight / 2
         lngSrcHeight = .ScaleHeight + 1 + (1 And (Screen.TwipsPerPixelX = 15))
         lngSrcWidth = .ScaleWidth + 1 + (1 And (Screen.TwipsPerPixelX = 15))
      End With
      
      sngSine = Sin(270 * PI_PART)
      sngCosine = Cos(270 * PI_PART)
      sngWidth = -lngSrcWidth / 2
      sngHeight = -lngSrcHeight / 2
      
      For intIndex = 0 To 2
         With ptaTemp(intIndex)
            .X = sngWidth + (lngSrcWidth And (intIndex = 1))
            .Y = sngHeight + (lngSrcHeight And (intIndex = 2))
            ptaBuffer(intIndex).X = (.X * sngCosine - .Y * sngSine) + lngDstHeight
            ptaBuffer(intIndex).Y = (.X * sngSine + .Y * sngCosine) + lngDstWidth
         End With
      Next 'intIndex
      
      PlgBlt hDC, ptaBuffer(0), picProcess.hDC, -1, 0, lngSrcWidth, lngSrcHeight, 0, 0, 0
      Picture = Image
   End If
   
   Erase ptaTemp
   Erase ptaBuffer

End Sub

Private Sub UserControl_Initialize()

   Call DrawBorder

End Sub

Private Sub UserControl_InitProperties()

   m_Alignment = Center
   Set Font = Ambient.Font
   
   Call ChangeColor

End Sub

' load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   With PropBag
      m_Alignment = .ReadProperty("Alignment", Center)
      m_BarColor = .ReadProperty("BarColor", Default)
      m_BorderColor = .ReadProperty("BorderColor", picBorder.Point(0, 0))
      m_BorderDefault = .ReadProperty("BorderDefault", True)
      m_Caption = .ReadProperty("Caption", "")
      UserControl.Enabled = .ReadProperty("Enabled", True)
      Set Font = .ReadProperty("Font", Ambient.Font)
      m_ForeColor = .ReadProperty("ForeColor", &HFFFFC0)
      m_Orientation = .ReadProperty("Orientation", Horizontal)
      m_Value = .ReadProperty("Value", 0)
   End With
   
   Call DrawBorder
   Call ChangeColor
   Call Paint

End Sub

Private Sub UserControl_Resize()

Static IsBusy As Boolean

   If IsBusy Then Exit Sub
   
   IsBusy = True
   picBuffer.Height = picBar.Height
   picBuffer.Width = 2
   
   If Width > Height Then
      Orientation = Horizontal
      
   Else
      Orientation = Vertical
   End If
   
   With picProcess
      If m_Orientation = Horizontal Then
         If ScaleHeight <> picBorder.ScaleHeight Then Height = picBorder.ScaleHeight * Screen.TwipsPerPixelY
         
         .Height = ScaleHeight
         .Width = ScaleWidth
         
      Else
         If ScaleWidth <> picBorder.ScaleHeight Then Width = picBorder.ScaleHeight * Screen.TwipsPerPixelY
         
         .Height = ScaleWidth
         .Width = ScaleHeight
      End If
   End With
   
   Call ChangeOrientation
   Call DrawBorder
   
   IsBusy = False

End Sub

' write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   With PropBag
      .WriteProperty "Alignment", m_Alignment, Center
      .WriteProperty "BarColor", m_BarColor, Default
      .WriteProperty "BorderColor", m_BorderColor, picBorder.Point(0, 0)
      .WriteProperty "BorderDefault", m_BorderDefault, True
      .WriteProperty "Caption", m_Caption, ""
      .WriteProperty "Enabled", UserControl.Enabled, True
      .WriteProperty "Font", picProcess.Font, Ambient.Font
      .WriteProperty "ForeColor", m_ForeColor, &HFFFFC0
      .WriteProperty "Orientation", m_Orientation, Horizontal
      .WriteProperty "Value", m_Value, 0
   End With

End Sub
