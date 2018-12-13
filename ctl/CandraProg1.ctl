VERSION 5.00
Begin VB.UserControl CandraProgressbar 
   ClientHeight    =   1350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6405
   ScaleHeight     =   1350
   ScaleWidth      =   6405
   ToolboxBitmap   =   "CandraProg1.ctx":0000
   Begin VB.Label ProgFont 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "Digiface"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   225
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   360
   End
   Begin VB.Shape pic2 
      BorderColor     =   &H00E1E1E1&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   735
   End
   Begin VB.Shape pic31 
      BorderColor     =   &H00E1E1E1&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000CCFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape pic1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "CandraProgressbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'####################################################'
'#                FLEX PROGRESSBAR 1                #'
'#                                                  #'
'#        CREATED BY FLEX SOFTWARE    2005          #'
'#                                                  #'
'#                  FEEL FREE TO EDIT               #'
'#                                                  #'
'#       THERE IS NO LICENSE FOR THIS CONTROL       #'
'####################################################'

Private Inited As Boolean
Private vMax1 As Long, vValue1 As Long

Public Enum ProgStyle
    [Solid] = 0
    [Transparant] = 1
    [Horizontal Line] = 2
    [Vertical Line] = 3
    [Upward Diagonal] = 4
    [Downward Diagonal] = 5
    [Cross] = 6
    [Diagonal Cross] = 7
End Enum

'Private FontTrans As Boolean


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        vMax1 = .ReadProperty("MAX", 100)
        vValue1 = .ReadProperty("VAL", 0)
        'FontTrans = .ReadProperty("FONTTRANS", True)
        Set UserControl.Font = .ReadProperty("FONT", UserControl.Font)
        Set ProgFont.Font = UserControl.Font
        ProgFont.Caption = .ReadProperty("CAPTION", "Text")
        
        
        pic31.FillStyle = .ReadProperty("PROGSTYLE", 0)
        pic31.FillColor = .ReadProperty("PROGCOL", vbBlue)
        pic1.FillColor = .ReadProperty("BACKCOL", vbWhite)
        pic2.BorderColor = .ReadProperty("BORDCOL", vbBlack)
        ProgFont.ForeColor = .ReadProperty("FONTCOL", vbBlack)
        
        'ProgFont.Font = UserControl.Font
        'ProgFont.FontSize = .ReadProperty("FONTSIZE", 8)
        'ProgFont.FontBold = .ReadProperty("FONTBOLD", False)
        'ProgFont.FontItalic = .ReadProperty("FONTITALIC", False)
        'ProgFont.FontUnderline = .ReadProperty("FONTUNDER", False)
        'ProgFont.FontStrikethru = .ReadProperty("FONTSTRIKE", False)
        
    End With
    
    
    
     '   If FontTrans = False Then
    '    ProgFont.BackStyle = 1
    'Else
    '    ProgFont.BackStyle = 0
    'End If
    ProgFont.BackStyle = 0
    Redrawbar
    Inited = True
End Sub

Private Sub UserControl_Resize()
    pic1.Width = UserControl.ScaleWidth ' - 10
    pic1.Height = UserControl.ScaleHeight ' - 10
    pic2.Width = UserControl.ScaleWidth ' - 10
    pic2.Height = UserControl.ScaleHeight ' - 10
    pic31.Height = UserControl.ScaleHeight ' - 10
    If Inited = True Then Redrawbar
End Sub


Public Property Get Font() As Font
    
    Set Font = UserControl.Font
    
End Property

Public Property Set Font(ByRef newFont As Font)
          'PropertyChanged "FONTSIZE"
         'PropertyChanged "FONTBOLD"
         'PropertyChanged "FONTITALIC"
         'PropertyChanged "FONTUNDER"
         'PropertyChanged "FONTSTRIKE"
         PropertyChanged "FONT"
        
        
    Set UserControl.Font = newFont
    
    Set ProgFont.Font = UserControl.Font
    Redrawbar
    'ProgFont.FontSize = UserControl.FontSize
    'ProgFont.FontBold = UserControl.FontBold
    'ProgFont.FontItalic = UserControl.FontItalic
    'ProgFont.FontUnderline = UserControl.FontUnderline
    'ProgFont.FontStrikethru = UserControl.FontStrikethru
End Property




'Public Property Get CaptionTransparant() As Boolean
    
'    CaptionTransparant = FontTrans
    
'End Property

'Public Property Let CaptionTransparant(ByVal newVal As Boolean)
'    FontTrans = newVal
'    If FontTrans = False Then
'        ProgFont.BackStyle = 1
'    Else
'        ProgFont.BackStyle = 0
'    End If
'    PropertyChanged "FONTTRANS"
'End Property

Public Property Get Caption() As String
    
    Caption = ProgFont.Caption
    
End Property

Public Property Let Caption(ByRef newVal As String)
    ProgFont.Caption = newVal
    PropertyChanged "CAPTION"
End Property

Public Property Get ProgressStyle() As ProgStyle
    
    ProgressStyle = pic31.FillStyle
    
End Property

Public Property Let ProgressStyle(ByVal newVal As ProgStyle)
    
    pic31.FillStyle = newVal
    PropertyChanged "PROGSTYLE"
End Property


Public Property Get Max() As Long
    
    Max = vMax1
    
End Property

Public Property Let Max(ByVal newVal As Long)
    
    vMax1 = newVal
    PropertyChanged "MAX"
    Redrawbar
End Property

Public Property Get Value() As Long
    
    Value = vValue1
    
End Property

Public Property Let Value(ByVal newVal As Long)
    
    vValue1 = newVal
    PropertyChanged "VALUE"
    Redrawbar
End Property

Public Property Get FontColor() As OLE_COLOR
    
    FontColor = ProgFont.ForeColor
    
End Property

Public Property Let FontColor(ByVal newCol As OLE_COLOR)
    
    ProgFont.ForeColor = newCol
    PropertyChanged "FONTCOL"
End Property

Public Property Get BackColor() As OLE_COLOR
    
    BackColor = pic1.FillColor
    
End Property

Public Property Let BackColor(ByVal newCol As OLE_COLOR)
    
    pic1.FillColor = newCol
    PropertyChanged "BACKCOL"
End Property


Public Property Get ProgressColor() As OLE_COLOR
    
    ProgressColor = pic31.FillColor
    
End Property

Public Property Let ProgressColor(ByVal newCol As OLE_COLOR)
    
    pic31.FillColor = newCol
    PropertyChanged "PROGCOL"
End Property




Public Property Get BorderColor() As OLE_COLOR
    
    BorderColor = pic2.BorderColor
    
End Property

Public Property Let BorderColor(ByVal newCol As OLE_COLOR)
    
    pic2.BorderColor = newCol
    PropertyChanged "BORDCOL"
End Property

'Public Property Get CaptionBackColor() As OLE_COLOR
    
'    CaptionBackColor = ProgFont.BackColor
    
'End Property

'Public Property Let CaptionBackColor(ByVal newCol As OLE_COLOR)
    
'    ProgFont.BackColor = newCol
'    PropertyChanged "CAPCOL"
'End Property

Private Function Redrawbar()
    ProgFont.Top = pic31.Height / 2 - ProgFont.Height / 2
    ProgFont.Left = pic1.Width / 2 - ProgFont.Width / 2
End Function

Private Sub UserControl_Terminate()
    Inited = False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "MAX", vMax1
        .WriteProperty "VAL", vValue1
        '.WriteProperty "CAPCOL", ProgFont.BackColor
        '.WriteProperty "FONTTRANS", FontTrans
        .WriteProperty "FONT", UserControl.Font
        .WriteProperty "CAPTION", ProgFont.Caption
        .WriteProperty "PROGSTYLE", pic31.FillStyle
        .WriteProperty "PROGCOL", pic31.FillColor
        .WriteProperty "BACKCOL", pic1.FillColor
        .WriteProperty "BORDCOL", pic2.BorderColor
        .WriteProperty "FONTCOL", ProgFont.ForeColor
        
        
        '.WriteProperty "FONTSIZE", ProgFont.FontSize
        '.WriteProperty "FONTBOLD", ProgFont.FontBold
        '.WriteProperty "FONTITALIC", ProgFont.FontItalic
        '.WriteProperty "FONTUNDER", ProgFont.FontUnderline
        '.WriteProperty "FONTSTRIKE", ProgFont.FontStrikethru
    End With
End Sub


Private Property Get FontBold() As Boolean

    FontBold = UserControl.FontBold

End Property

Private Property Let FontBold(ByVal NewValue As Boolean)

    UserControl.FontBold = NewValue


End Property

Private Property Get FontItalic() As Boolean

    FontItalic = UserControl.FontItalic

End Property

Private Property Let FontItalic(ByVal NewValue As Boolean)

    UserControl.FontItalic = NewValue


End Property

Private Property Get FontUnderline() As Boolean

    FontUnderline = UserControl.FontUnderline

End Property

Private Property Let FontUnderline(ByVal NewValue As Boolean)

    UserControl.FontUnderline = NewValue


End Property

Private Property Get FontSize() As Integer

    FontSize = UserControl.FontSize

End Property

Private Property Let FontSize(ByVal NewValue As Integer)

    UserControl.FontSize = NewValue

End Property

Private Property Get FontName() As String

    FontName = UserControl.FontName

End Property

Private Property Let FontName(ByVal NewValue As String)

    UserControl.FontName = NewValue


End Property

