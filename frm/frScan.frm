VERSION 5.00
Begin VB.Form frScan 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   4455
      Left            =   0
      Picture         =   "frScan.frx":0000
      ScaleHeight     =   4395
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   -240
      Width           =   6135
      Begin prjDAA.jcbutton jcbutton1 
         Height          =   495
         Left            =   1440
         TabIndex        =   13
         Top             =   2760
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         ButtonStyle     =   9
         ShowFocusRect   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14800597
         Caption         =   "Hide"
         MousePointer    =   99
         MouseIcon       =   "frScan.frx":2D32
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   120
         ScaleHeight     =   1545
         ScaleWidth      =   4425
         TabIndex        =   11
         Top             =   1080
         Width           =   4455
         Begin VB.PictureBox Picture6 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1095
            Left            =   120
            Picture         =   "frScan.frx":2E94
            ScaleHeight     =   1095
            ScaleWidth      =   1095
            TabIndex        =   12
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lbLogo 
            BackStyle       =   0  'Transparent
            Caption         =   "Real - Time Protection Is On"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   375
            Index           =   7
            Left            =   1320
            TabIndex        =   14
            Top             =   600
            Width           =   3375
         End
      End
      Begin VB.Timer Timer3 
         Interval        =   10000
         Left            =   960
         Top             =   1080
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   600
         Top             =   1080
      End
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   240
         Top             =   1080
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   4680
         MouseIcon       =   "frScan.frx":6B96
         ScaleHeight     =   3375
         ScaleWidth      =   1335
         TabIndex        =   1
         Top             =   120
         Width           =   1335
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   720
            MouseIcon       =   "frScan.frx":6CE8
            MousePointer    =   99  'Custom
            Picture         =   "frScan.frx":6E3A
            ScaleHeight     =   240
            ScaleMode       =   0  'User
            ScaleWidth      =   225.882
            TabIndex        =   9
            Top             =   120
            Width           =   240
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000007&
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   360
            Picture         =   "frScan.frx":706C
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   2
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lbLogo 
            BackStyle       =   0  'Transparent
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Index           =   5
            Left            =   480
            TabIndex        =   8
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label lbLogo 
            BackStyle       =   0  'Transparent
            Caption         =   "O"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Index           =   4
            Left            =   480
            TabIndex        =   7
            Top             =   2400
            Width           =   495
         End
         Begin VB.Label lbLogo 
            BackStyle       =   0  'Transparent
            Caption         =   "F"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Index           =   3
            Left            =   480
            TabIndex        =   6
            Top             =   2760
            Width           =   495
         End
         Begin VB.Label lbLogo 
            BackStyle       =   0  'Transparent
            Caption         =   "U"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Index           =   2
            Left            =   480
            TabIndex        =   5
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label lbLogo 
            BackStyle       =   0  'Transparent
            Caption         =   "Y"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Index           =   1
            Left            =   480
            TabIndex        =   4
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label lbLogo 
            BackStyle       =   0  'Transparent
            Caption         =   "I"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Index           =   0
            Left            =   540
            TabIndex        =   3
            Top             =   960
            Width           =   495
         End
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frScan.frx":ABA6
         Top             =   480
         Width           =   480
      End
      Begin VB.Label lbLogo 
         BackStyle       =   0  'Transparent
         Caption         =   "Real - Time Protection"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Index           =   6
         Left            =   960
         TabIndex        =   10
         Top             =   600
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Oval
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'Akhir oval


Private Sub Form_Load()
Me.Top = Screen.Height
Me.Left = Screen.Width - Me.Width
    LetakanForm frRTP, True
    'membuat form oval
End Sub

Private Sub jcbutton1_Click()
Timer2.Enabled = True
End Sub

Private Sub Picture4_Click()
Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()
If Me.Top <= Screen.Height - (Me.Height + 250) Then Timer1.Enabled = False
Me.Top = Me.Top - 300
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If Me.Top >= 11000 Then
Unload Me
Timer2.Enabled = False
End If
Me.Top = Me.Top + 300
End Sub

Private Sub Timer3_Timer()
Timer2.Enabled = True
End Sub
