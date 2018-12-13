VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H0000CCFF&
   BorderStyle     =   0  'None
   Caption         =   "Bee Antivirus 2014"
   ClientHeight    =   10215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12105
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   Icon            =   "fruit.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "fruit.frx":19F7A
   ScaleHeight     =   681
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   807
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Menu5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   2400
      ScaleHeight     =   4905
      ScaleWidth      =   9225
      TabIndex        =   172
      Top             =   1560
      Width           =   9255
      Begin VB.PictureBox picSec 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   240
         Picture         =   "fruit.frx":125732
         ScaleHeight     =   2415
         ScaleWidth      =   2535
         TabIndex        =   196
         Top             =   960
         Width           =   2535
      End
      Begin VB.PictureBox picSec3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5400
         Picture         =   "fruit.frx":136752
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   195
         Top             =   3000
         Width           =   255
      End
      Begin VB.PictureBox picSec2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5400
         Picture         =   "fruit.frx":136ADC
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   194
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox picSec1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5400
         Picture         =   "fruit.frx":136E66
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   193
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox picSec5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   4080
         Picture         =   "fruit.frx":1371F0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   192
         Top             =   3840
         Width           =   255
      End
      Begin VB.PictureBox picSec4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5400
         Picture         =   "fruit.frx":13757A
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   191
         Top             =   2040
         Width           =   255
      End
      Begin VB.PictureBox picSec6 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5400
         Picture         =   "fruit.frx":137904
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   190
         Top             =   2520
         Width           =   255
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2190
         Left            =   1320
         Picture         =   "fruit.frx":137C8E
         ScaleHeight     =   2190
         ScaleWidth      =   2340
         TabIndex        =   189
         Top             =   2160
         Width           =   2340
      End
      Begin VB.PictureBox picNotSec 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   240
         Picture         =   "fruit.frx":1487BA
         ScaleHeight     =   2415
         ScaleWidth      =   2535
         TabIndex        =   182
         Top             =   960
         Width           =   2535
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5640
         Picture         =   "fruit.frx":1597DA
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   181
         Top             =   2520
         Width           =   255
      End
      Begin VB.PictureBox Picture13 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5640
         Picture         =   "fruit.frx":159B1C
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   180
         Top             =   3000
         Width           =   255
      End
      Begin VB.PictureBox Picture14 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5640
         Picture         =   "fruit.frx":159E5E
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   179
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox Picture15 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5640
         Picture         =   "fruit.frx":15A1A0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   178
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox Picture16 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   4320
         Picture         =   "fruit.frx":15A4E2
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   177
         Top             =   3840
         Width           =   255
      End
      Begin VB.PictureBox Picture17 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5640
         Picture         =   "fruit.frx":15A824
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   176
         Top             =   2040
         Width           =   255
      End
      Begin prjBeeAV.jcbutton jcbutton9 
         Height          =   270
         Left            =   7245
         TabIndex        =   173
         Top             =   1080
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   476
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "Turn On"
         ForeColor       =   192
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin prjBeeAV.jcbutton jcbutton1 
         Height          =   615
         Left            =   6960
         TabIndex        =   174
         Top             =   4080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         ButtonStyle     =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14933984
         Caption         =   "Scan Now >>"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin prjBeeAV.jcbutton jcbutton8 
         Height          =   270
         Left            =   7320
         TabIndex        =   175
         Top             =   2520
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   476
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "Update"
         ForeColor       =   8421376
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "On"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   6120
         TabIndex        =   188
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "On"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   6120
         TabIndex        =   187
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Waiting for Instructions ...."
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   6120
         TabIndex        =   186
         Top             =   3000
         Width           =   3375
      End
      Begin VB.Label lbSec 
         BackStyle       =   0  'Transparent
         Caption         =   "Secured"
         BeginProperty Font 
            Name            =   "Blaster"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   615
         Left            =   2760
         TabIndex        =   185
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "On"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   6120
         TabIndex        =   184
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "2.1.0   Beta"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   6120
         TabIndex        =   183
         Top             =   2520
         Width           =   975
      End
   End
   Begin VB.PictureBox Menu0 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   4935
      Left            =   -7200
      ScaleHeight     =   4905
      ScaleWidth      =   9225
      TabIndex        =   120
      Top             =   4680
      Width           =   9255
      Begin VB.PictureBox picScan 
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
         Height          =   4335
         Index           =   1
         Left            =   0
         ScaleHeight     =   4335
         ScaleWidth      =   9255
         TabIndex        =   167
         Top             =   600
         Width           =   9255
         Begin prjBeeAV.AkuSayangIbu jcbutton2 
            Height          =   495
            Left            =   120
            TabIndex        =   169
            Top             =   3600
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   873
            BtnStyle        =   2
            Caption         =   "Fix "
         End
         Begin prjBeeAV.ucListView lvMal 
            Height          =   3495
            Left            =   120
            TabIndex        =   168
            Top             =   0
            Width           =   11055
            _ExtentX        =   15266
            _ExtentY        =   7435
            Style           =   4
            StyleEx         =   37
         End
         Begin prjBeeAV.AkuSayangIbu jcbutton3 
            Height          =   495
            Left            =   2040
            TabIndex        =   170
            Top             =   3600
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   873
            BtnStyle        =   2
            Caption         =   "Quarantine"
         End
         Begin prjBeeAV.AkuSayangIbu jcbutton4 
            Height          =   495
            Left            =   6720
            TabIndex        =   171
            Top             =   3600
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   873
            BtnStyle        =   2
            Caption         =   "More >>"
         End
         Begin VB.Line Line12 
            BorderColor     =   &H80000000&
            BorderStyle     =   6  'Inside Solid
            X1              =   6600
            X2              =   6600
            Y1              =   3480
            Y2              =   4320
         End
      End
      Begin VB.PictureBox picDetailed 
         Appearance      =   0  'Flat
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
         Height          =   4215
         Left            =   0
         ScaleHeight     =   4215
         ScaleWidth      =   9255
         TabIndex        =   140
         Top             =   600
         Width           =   9255
         Begin prjBeeAV.CRAVProgressBar pScan 
            Height          =   465
            Left            =   120
            TabIndex        =   166
            Top             =   480
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   820
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H00F9F9F9&
            BorderStyle     =   0  'None
            FillColor       =   &H00F9F9F9&
            ForeColor       =   &H80000008&
            Height          =   3135
            Left            =   120
            ScaleHeight     =   3135
            ScaleWidth      =   9015
            TabIndex        =   146
            Top             =   1080
            Width           =   9015
            Begin VB.Line Line6 
               BorderColor     =   &H00E0E0E0&
               BorderStyle     =   3  'Dot
               X1              =   4560
               X2              =   4560
               Y1              =   0
               Y2              =   3120
            End
            Begin VB.Label l 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Files Remaining"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   162
               Top             =   600
               Width           =   1350
            End
            Begin VB.Label l 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Files Scanned"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   161
               Top             =   960
               Width           =   1170
            End
            Begin VB.Label l 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Files Ignored"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   240
               Index           =   8
               Left            =   240
               TabIndex        =   160
               Top             =   1320
               Width           =   1110
            End
            Begin VB.Label l 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Files Infected"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H007080FF&
               Height          =   240
               Index           =   9
               Left            =   240
               TabIndex        =   159
               Top             =   1680
               Width           =   1155
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "Scanning Result"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   255
               Left            =   240
               TabIndex        =   158
               Top             =   2040
               Width           =   2055
            End
            Begin VB.Label lRem 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0 File[s]"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   255
               Left            =   4920
               TabIndex        =   157
               Top             =   600
               Width           =   660
            End
            Begin VB.Label lScanned 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0 File[s]"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   255
               Left            =   4920
               TabIndex        =   156
               Top             =   960
               Width           =   660
            End
            Begin VB.Label lIgnored 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0 File[s]"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   255
               Left            =   4920
               TabIndex        =   155
               Top             =   1320
               Width           =   660
            End
            Begin VB.Label lInFile 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0 File[s]"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H007080FF&
               Height          =   255
               Left            =   4920
               TabIndex        =   154
               Top             =   1680
               Width           =   660
            End
            Begin VB.Label Label27 
               BackStyle       =   0  'Transparent
               Caption         =   "No Threat found."
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   375
               Left            =   4920
               TabIndex        =   153
               Top             =   2040
               Width           =   4095
            End
            Begin VB.Label Label18 
               BackStyle       =   0  'Transparent
               Caption         =   "Status"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   375
               Left            =   4800
               TabIndex        =   152
               Top             =   120
               Width           =   1215
            End
            Begin VB.Image Image14 
               Height          =   3135
               Left            =   0
               Top             =   0
               Width           =   4575
            End
            Begin VB.Line Line5 
               BorderColor     =   &H00E0E0E0&
               X1              =   0
               X2              =   11400
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Category"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   300
               Left            =   120
               TabIndex        =   151
               Top             =   120
               Width           =   930
            End
            Begin VB.Image Image15 
               Height          =   3135
               Left            =   4560
               Top             =   0
               Width           =   4455
            End
            Begin VB.Label lTime 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "00:00:00"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   240
               Left            =   4920
               TabIndex        =   150
               Top             =   2760
               Width           =   750
            End
            Begin VB.Label l 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Time"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   240
               Index           =   7
               Left            =   240
               TabIndex        =   149
               Top             =   2760
               Width           =   420
            End
            Begin VB.Label l 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Speed"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   240
               Index           =   0
               Left            =   240
               TabIndex        =   148
               Top             =   2400
               Width           =   555
            End
            Begin VB.Label lSpeed 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0 File[s]/s"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   240
               Left            =   4920
               TabIndex        =   147
               Top             =   2400
               Width           =   870
            End
         End
         Begin VB.Label lblFile 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000006&
            Height          =   255
            Left            =   3000
            TabIndex        =   145
            Top             =   0
            Width           =   6135
            WordWrap        =   -1  'True
         End
         Begin VB.Label l 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Processed File         :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   120
            Index           =   27
            Left            =   8760
            TabIndex        =   144
            Top             =   4200
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label l 
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   37
            Left            =   8880
            TabIndex        =   143
            Top             =   4200
            Width           =   135
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "NO THREAT FOUND."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   120
            TabIndex        =   142
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Scan Complete,"
            Height          =   255
            Left            =   120
            TabIndex        =   141
            Top             =   0
            Width           =   2415
         End
      End
      Begin VB.PictureBox picScan 
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
         Height          =   4335
         Index           =   0
         Left            =   120
         ScaleHeight     =   4335
         ScaleWidth      =   11295
         TabIndex        =   121
         Top             =   840
         Width           =   11295
         Begin prjBeeAV.AkuSayangIbu Scan 
            Height          =   735
            Index           =   0
            Left            =   6840
            TabIndex        =   133
            Top             =   120
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   1296
            BtnStyle        =   2
            Caption         =   "Scan Now!"
         End
         Begin prjBeeAV.AkuSayangIbu Scan 
            Height          =   735
            Index           =   2
            Left            =   6840
            TabIndex        =   134
            Top             =   3000
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   1296
            BtnStyle        =   2
            Caption         =   "Scan Now!"
         End
         Begin prjBeeAV.AkuSayangIbu Scan 
            Height          =   735
            Index           =   1
            Left            =   6840
            TabIndex        =   135
            Top             =   1080
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   1296
            BtnStyle        =   2
            Caption         =   "Scan Now!"
         End
         Begin prjBeeAV.AkuSayangIbu Scan 
            Height          =   735
            Index           =   3
            Left            =   6840
            TabIndex        =   136
            Top             =   2040
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   1296
            BtnStyle        =   2
            Caption         =   "Scan Now!"
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Performs an in Depth scan of the system"
            Height          =   210
            Index           =   1
            Left            =   1200
            TabIndex        =   132
            Top             =   480
            Width           =   3915
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Performs a quick of your computer's system volume"
            Height          =   210
            Left            =   1200
            TabIndex        =   131
            Top             =   1440
            Width           =   3750
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Performs a full scan of  a  costum folder"
            Height          =   210
            Left            =   1200
            TabIndex        =   130
            Top             =   3360
            Width           =   2910
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Performs a process scan wich underway in of your computer"
            Height          =   210
            Left            =   1200
            TabIndex        =   129
            Top             =   2400
            Width           =   4500
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Full Scan"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   1200
            TabIndex        =   128
            Top             =   240
            Width           =   900
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quick Scan"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   1200
            TabIndex        =   127
            Top             =   1200
            Width           =   1080
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Costum Scan"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   1200
            TabIndex        =   126
            Top             =   3120
            Width           =   1245
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Process Scan"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   1200
            TabIndex        =   125
            Top             =   2160
            Width           =   1275
         End
         Begin VB.Image Image11 
            Height          =   750
            Left            =   240
            Picture         =   "fruit.frx":15AB66
            Top             =   120
            Width           =   750
         End
         Begin VB.Image Image12 
            Height          =   750
            Left            =   240
            Picture         =   "fruit.frx":15C096
            Top             =   3000
            Width           =   750
         End
         Begin VB.Image Image13 
            Height          =   750
            Left            =   240
            Picture         =   "fruit.frx":15D555
            Top             =   1080
            Width           =   750
         End
         Begin VB.Image Image16 
            Height          =   975
            Left            =   0
            Top             =   0
            Width           =   11535
         End
         Begin VB.Image Image17 
            Height          =   975
            Left            =   0
            Top             =   960
            Width           =   11535
         End
         Begin VB.Image Image18 
            Height          =   975
            Left            =   0
            Top             =   1920
            Width           =   11535
         End
         Begin VB.Image Image19 
            Height          =   975
            Left            =   0
            Top             =   2880
            Width           =   11535
         End
         Begin VB.Image Image20 
            Height          =   750
            Left            =   240
            Picture         =   "fruit.frx":15EAC4
            Top             =   2040
            Width           =   750
         End
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00E0E0E0&
         X1              =   120
         X2              =   9000
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Report"
         BeginProperty Font 
            Name            =   "Futura Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   1680
         TabIndex        =   139
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scan Metode"
         BeginProperty Font 
            Name            =   "Futura Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   240
         TabIndex        =   138
         Top             =   120
         Width           =   1170
      End
      Begin VB.Label Label120 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detected"
         BeginProperty Font 
            Name            =   "Futura Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   3360
         TabIndex        =   137
         Top             =   120
         Width           =   795
      End
   End
   Begin VB.PictureBox Menu4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   4935
      Left            =   -7680
      ScaleHeight     =   4905
      ScaleWidth      =   9225
      TabIndex        =   76
      Top             =   6000
      Width           =   9255
      Begin VB.PictureBox picAbout 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4095
         Index           =   2
         Left            =   0
         ScaleHeight     =   4095
         ScaleWidth      =   9255
         TabIndex        =   100
         Top             =   720
         Width           =   9255
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1695
            Left            =   7200
            Picture         =   "fruit.frx":15FF8E
            ScaleHeight     =   113
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   121
            TabIndex        =   119
            Top             =   120
            Width           =   1815
         End
         Begin VB.ListBox lstHeur 
            Appearance      =   0  'Flat
            ForeColor       =   &H003B3B3B&
            Height          =   1395
            ItemData        =   "fruit.frx":16ABDE
            Left            =   4080
            List            =   "fruit.frx":16AC0C
            TabIndex        =   109
            Top             =   240
            Width           =   2895
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Program Info"
            Height          =   975
            Left            =   4080
            TabIndex        =   106
            Top             =   1800
            Width           =   3015
            Begin VB.Label l 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Source Code                  : Bee Av"
               Height          =   195
               Index           =   33
               Left            =   120
               TabIndex        =   108
               Top             =   600
               Width           =   2385
            End
            Begin VB.Label l 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Program Version            : 1.3.0   Beta"
               Height          =   195
               Index           =   32
               Left            =   120
               TabIndex        =   107
               Top             =   360
               Width           =   2745
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Definition Info"
            Height          =   1335
            Left            =   4080
            TabIndex        =   101
            Top             =   2760
            Width           =   3015
            Begin VB.Label lDate 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Unknow"
               Height          =   210
               Left            =   120
               TabIndex        =   105
               Top             =   960
               Width           =   600
            End
            Begin VB.Label l 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Database Definition       :"
               Height          =   210
               Index           =   30
               Left            =   120
               TabIndex        =   104
               Top             =   720
               Width           =   1755
            End
            Begin VB.Label lDB 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               Height          =   210
               Left            =   120
               TabIndex        =   103
               Top             =   480
               Width           =   90
            End
            Begin VB.Label l 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Database Virus             :"
               Height          =   210
               Index           =   29
               Left            =   120
               TabIndex        =   102
               Top             =   240
               Width           =   1755
            End
         End
         Begin prjBeeAV.ucListView lvVirLst 
            Height          =   3735
            Left            =   120
            TabIndex        =   110
            Top             =   240
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   6588
            Style           =   4
            StyleEx         =   33
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Heuristic :"
            Height          =   210
            Left            =   4080
            TabIndex        =   112
            Top             =   0
            Width           =   720
         End
         Begin VB.Label l 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Virus List :"
            Height          =   210
            Index           =   31
            Left            =   120
            TabIndex        =   111
            Top             =   0
            Width           =   780
         End
      End
      Begin VB.PictureBox picAbout 
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
         Height          =   4095
         Index           =   1
         Left            =   -120
         ScaleHeight     =   4095
         ScaleWidth      =   11175
         TabIndex        =   98
         Top             =   600
         Width           =   11175
         Begin VB.ListBox lstThanks 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000011&
            Height          =   3735
            ItemData        =   "fruit.frx":16AC92
            Left            =   240
            List            =   "fruit.frx":16ACF0
            TabIndex        =   99
            Top             =   120
            Width           =   9015
         End
      End
      Begin VB.PictureBox picAbout 
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
         Height          =   4095
         Index           =   0
         Left            =   0
         ScaleHeight     =   4095
         ScaleWidth      =   11175
         TabIndex        =   77
         Top             =   600
         Width           =   11175
         Begin VB.PictureBox Picture7 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1575
            Left            =   9720
            Picture         =   "fruit.frx":16AFB5
            ScaleHeight     =   1575
            ScaleWidth      =   1455
            TabIndex        =   89
            Top             =   2520
            Width           =   1455
         End
         Begin VB.PictureBox Picture8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1455
            Left            =   360
            ScaleHeight     =   1455
            ScaleWidth      =   7575
            TabIndex        =   78
            Top             =   2280
            Width           =   7575
            Begin VB.Shape Shape2 
               BorderColor     =   &H00E1E1E1&
               Height          =   1455
               Index           =   1
               Left            =   0
               Top             =   0
               Width           =   7575
            End
            Begin VB.Label lblEmpty 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Designer"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   225
               Index           =   8
               Left            =   120
               TabIndex        =   88
               Top             =   360
               Width           =   690
            End
            Begin VB.Label lblEmpty 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Virus Hunter"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   225
               Index           =   10
               Left            =   120
               TabIndex        =   87
               Top             =   840
               Width           =   990
            End
            Begin VB.Label lblEmpty 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   ": Candra Ramadhan Prasetya"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   225
               Index           =   11
               Left            =   2640
               TabIndex        =   86
               ToolTipText     =   "Muhammad Dipa Yumansah"
               Top             =   360
               Width           =   2265
            End
            Begin VB.Label lblHelpAbout 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   ": Candra Ramadhan Prasetya"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   225
               Index           =   0
               Left            =   2640
               MousePointer    =   10  'Up Arrow
               TabIndex        =   85
               ToolTipText     =   "Tanio Zamin"
               Top             =   840
               Width           =   2265
            End
            Begin VB.Label lblHelpAbout 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   ": Sigit,Solahudin,Roiyu,Bayu,Arya,Yadi."
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   225
               Index           =   1
               Left            =   2640
               MousePointer    =   10  'Up Arrow
               TabIndex        =   84
               ToolTipText     =   " Aming Anjas Asmara Pamungkas"
               Top             =   600
               Width           =   3045
            End
            Begin VB.Label lblEmpty 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Promotion"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   225
               Index           =   13
               Left            =   120
               TabIndex        =   83
               Top             =   600
               Width           =   855
            End
            Begin VB.Label lblHelpAbout 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   ": Candra Ramadhan Prasetya"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   225
               Index           =   2
               Left            =   2640
               MousePointer    =   10  'Up Arrow
               TabIndex        =   82
               ToolTipText     =   "Aditya Warman"
               Top             =   1080
               Width           =   2265
            End
            Begin VB.Label lblEmpty 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Virus Analysis"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   285
               Index           =   9
               Left            =   120
               TabIndex        =   81
               Top             =   1080
               Width           =   1440
            End
            Begin VB.Label lblEmpty 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Programmer"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   225
               Index           =   7
               Left            =   120
               TabIndex        =   80
               Top             =   120
               Width           =   1005
            End
            Begin VB.Label lblEmpty 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   ": Candra Ramadhan Prasetya"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   225
               Index           =   12
               Left            =   2640
               TabIndex        =   79
               ToolTipText     =   "Muhammad Dipa Yumansah"
               Top             =   120
               Width           =   2265
            End
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00E1E1E1&
            Height          =   1335
            Index           =   0
            Left            =   360
            Top             =   120
            Width           =   6855
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "E-mail"
            BeginProperty Font 
               Name            =   "Futura Md BT"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H007F7F7F&
            Height          =   210
            Index           =   2
            Left            =   600
            TabIndex        =   97
            Top             =   1080
            Width           =   480
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   ": candraramadhan75@gmail.com"
            BeginProperty Font 
               Name            =   "Futura Md BT"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H007F7F7F&
            Height          =   210
            Index           =   6
            Left            =   2760
            TabIndex        =   96
            Top             =   1080
            Width           =   2505
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   ": 13 May 2014"
            BeginProperty Font 
               Name            =   "Futura Md BT"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H007F7F7F&
            Height          =   210
            Index           =   5
            Left            =   2760
            TabIndex        =   95
            Top             =   840
            Width           =   1125
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   ": 1.3.0"
            BeginProperty Font 
               Name            =   "Futura Md BT"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H007F7F7F&
            Height          =   210
            Index           =   4
            Left            =   2760
            TabIndex        =   94
            Top             =   600
            Width           =   510
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Release"
            BeginProperty Font 
               Name            =   "Futura Md BT"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H007F7F7F&
            Height          =   210
            Index           =   1
            Left            =   600
            TabIndex        =   93
            Top             =   840
            Width           =   570
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Program Version"
            BeginProperty Font 
               Name            =   "Futura Md BT"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H007F7F7F&
            Height          =   210
            Index           =   0
            Left            =   600
            TabIndex        =   92
            Top             =   600
            Width           =   1230
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Bee Antivirus :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000CCFF&
            Height          =   375
            Left            =   600
            TabIndex        =   91
            Top             =   240
            Width           =   4695
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "Bee Team :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000CCFF&
            Height          =   375
            Left            =   360
            TabIndex        =   90
            Top             =   1920
            Width           =   4695
         End
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         X1              =   120
         X2              =   9000
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label133 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Thanks To"
         BeginProperty Font 
            Name            =   "Futura Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   1920
         TabIndex        =   118
         Top             =   120
         Width           =   945
      End
      Begin VB.Label Label221 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bee Team Work"
         BeginProperty Font 
            Name            =   "Futura Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   240
         TabIndex        =   117
         Top             =   120
         Width           =   1410
      End
      Begin VB.Label Label1101 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "More Info"
         BeginProperty Font 
            Name            =   "Futura Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   3120
         TabIndex        =   116
         Top             =   120
         Width           =   885
      End
   End
   Begin VB.PictureBox Menu2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   4935
      Left            =   3840
      ScaleHeight     =   4905
      ScaleWidth      =   9225
      TabIndex        =   53
      Top             =   7080
      Width           =   9255
      Begin VB.PictureBox gGr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2655
         Left            =   720
         ScaleHeight     =   2655
         ScaleWidth      =   7455
         TabIndex        =   66
         Top             =   840
         Width           =   7455
         Begin prjBeeAV.mm_checkbox ck 
            Height          =   390
            Index           =   4
            Left            =   0
            TabIndex        =   67
            Top             =   0
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   688
            RoundedValue    =   0
         End
         Begin prjBeeAV.mm_checkbox ck 
            Height          =   390
            Index           =   6
            Left            =   0
            TabIndex        =   69
            Top             =   480
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   688
            Value           =   1
            RoundedValue    =   0
         End
         Begin prjBeeAV.mm_checkbox ck 
            Height          =   390
            Index           =   7
            Left            =   0
            TabIndex        =   71
            Top             =   960
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   688
            Value           =   1
            RoundedValue    =   0
         End
         Begin prjBeeAV.mm_checkbox ck 
            Height          =   390
            Index           =   5
            Left            =   0
            TabIndex        =   73
            Top             =   1440
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   688
            RoundedValue    =   0
         End
         Begin VB.Label CheckInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Enable Context Menu on Explorer ""Scan With Bee"""
            BeginProperty Font 
               Name            =   "BankGothic Md BT"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   180
            Index           =   7
            Left            =   960
            TabIndex        =   74
            Top             =   1560
            Width           =   4725
         End
         Begin VB.Label CheckInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Enable Protection"
            BeginProperty Font 
               Name            =   "BankGothic Md BT"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   180
            Index           =   6
            Left            =   960
            TabIndex        =   72
            Top             =   1080
            Width           =   1725
         End
         Begin VB.Label CheckInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Auto Scan FlashDisk"
            BeginProperty Font 
               Name            =   "BankGothic Md BT"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   180
            Index           =   5
            Left            =   960
            TabIndex        =   70
            Top             =   600
            Width           =   1950
         End
         Begin VB.Label CheckInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Run on Start-Up"
            BeginProperty Font 
               Name            =   "BankGothic Md BT"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   180
            Index           =   4
            Left            =   960
            TabIndex        =   68
            Top             =   120
            Width           =   1560
         End
      End
      Begin prjBeeAV.mm_checkbox ck 
         Height          =   240
         Index           =   0
         Left            =   7560
         TabIndex        =   65
         Top             =   480
         Visible         =   0   'False
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   423
         Small           =   0   'False
         RoundedValue    =   0
      End
      Begin prjBeeAV.mm_checkbox ck 
         Height          =   390
         Index           =   1
         Left            =   720
         TabIndex        =   57
         Top             =   840
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   688
         RoundedValue    =   0
      End
      Begin prjBeeAV.AkuSayangIbu cmdSSet 
         Height          =   375
         Left            =   1320
         TabIndex        =   54
         Top             =   4320
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   661
         BtnStyle        =   2
         Caption         =   "Apply"
      End
      Begin prjBeeAV.mm_checkbox ck 
         Height          =   390
         Index           =   2
         Left            =   720
         TabIndex        =   59
         Top             =   1320
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   688
         RoundedValue    =   0
      End
      Begin prjBeeAV.mm_checkbox ck 
         Height          =   390
         Index           =   3
         Left            =   720
         TabIndex        =   61
         Top             =   1800
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   688
         Value           =   1
         RoundedValue    =   0
      End
      Begin prjBeeAV.mm_checkbox ck 
         Height          =   390
         Index           =   8
         Left            =   720
         TabIndex        =   63
         Top             =   2280
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   688
         RoundedValue    =   0
      End
      Begin VB.Label susu 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Saved !"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Left            =   240
         TabIndex        =   75
         Top             =   4320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label CheckInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Filter by Extentions"
         BeginProperty Font 
            Name            =   "BankGothic Md BT"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   180
         Index           =   3
         Left            =   1680
         TabIndex        =   64
         Top             =   2400
         Width           =   1890
      End
      Begin VB.Label CheckInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Use Heuristic Function for more Accurate Scanning"
         BeginProperty Font 
            Name            =   "BankGothic Md BT"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   180
         Index           =   2
         Left            =   1680
         TabIndex        =   62
         Top             =   1920
         Width           =   4890
      End
      Begin VB.Label CheckInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Filter by Size"
         BeginProperty Font 
            Name            =   "BankGothic Md BT"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   180
         Index           =   1
         Left            =   1680
         TabIndex        =   60
         Top             =   1440
         Width           =   1260
      End
      Begin VB.Label CheckInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Filter by Extentions"
         BeginProperty Font 
            Name            =   "BankGothic Md BT"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   180
         Index           =   0
         Left            =   1680
         TabIndex        =   58
         Top             =   960
         Width           =   1890
      End
      Begin VB.Label label871 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scanning Setting "
         BeginProperty Font 
            Name            =   "Futura Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   240
         TabIndex        =   56
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label298 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Application Setting "
         BeginProperty Font 
            Name            =   "Futura Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   2040
         TabIndex        =   55
         Top             =   120
         Width           =   1755
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         X1              =   120
         X2              =   9120
         Y1              =   480
         Y2              =   480
      End
   End
   Begin VB.PictureBox Menu3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   4935
      Left            =   -720
      ScaleHeight     =   4905
      ScaleWidth      =   9225
      TabIndex        =   51
      Top             =   7200
      Width           =   9255
      Begin prjBeeAV.ucListView lvQuar 
         Height          =   4695
         Left            =   120
         TabIndex        =   52
         Top             =   120
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   8281
         Border          =   0
         Style           =   4
         StyleEx         =   33
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00E1E1E1&
         Height          =   4735
         Left            =   90
         Top             =   90
         Width           =   9085
      End
   End
   Begin VB.PictureBox Menu1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   4935
      Left            =   -6480
      ScaleHeight     =   4905
      ScaleWidth      =   9225
      TabIndex        =   0
      Top             =   6840
      Width           =   9255
      Begin VB.PictureBox picTool 
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
         Height          =   4215
         Index           =   3
         Left            =   0
         ScaleHeight     =   4215
         ScaleWidth      =   9135
         TabIndex        =   7
         Top             =   600
         Width           =   9135
         Begin prjBeeAV.AkuSayangIbu cCancelRegT 
            Height          =   375
            Left            =   5640
            TabIndex        =   49
            Top             =   3600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            BtnStyle        =   2
            Caption         =   "Cancel"
         End
         Begin VB.Frame fCusXP 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Register Owner and Organization"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1575
            Left            =   360
            TabIndex        =   31
            Top             =   2520
            Width           =   3975
            Begin VB.TextBox txtOwner 
               Height          =   315
               Left            =   120
               TabIndex        =   33
               Top             =   480
               Width           =   3735
            End
            Begin VB.TextBox txtOrg 
               Height          =   315
               Left            =   120
               TabIndex        =   32
               Top             =   1080
               Width           =   3735
            End
            Begin VB.Label l 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Owner"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   13
               Left            =   120
               TabIndex        =   35
               Top             =   240
               Width           =   555
            End
            Begin VB.Label l 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Organization"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   10
               Left            =   120
               TabIndex        =   34
               Top             =   840
               Width           =   1035
            End
         End
         Begin VB.Frame fCpanel 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Control Panel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   855
            Left            =   4440
            TabIndex        =   28
            Top             =   0
            Width           =   4335
            Begin VB.CheckBox NoAddRemovePrograms 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Remove Add/Programs"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Tag             =   "Microsoft\Windows\CurrentVersion\Policies\Uninstall\NoAddRemovePrograms"
               Top             =   240
               Width           =   2715
            End
            Begin VB.CheckBox NoControlPanel 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Prohibit access to the Control Panel"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Tag             =   "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
               Top             =   480
               Width           =   3795
            End
         End
         Begin VB.Frame fExplorer 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Windows Explorer"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   2295
            Left            =   4440
            TabIndex        =   19
            Top             =   840
            Width           =   4335
            Begin VB.CheckBox Hidden 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Show Hidden File"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   27
               Tag             =   "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
               Top             =   240
               Width           =   2055
            End
            Begin VB.CheckBox ShowSuperHidden 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Show System File"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   26
               Tag             =   "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
               Top             =   480
               Width           =   2055
            End
            Begin VB.CheckBox NoTrayContextMenu 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Disable Tray Menu"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   25
               Tag             =   "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
               Top             =   960
               Width           =   2055
            End
            Begin VB.CheckBox NoSetTaskbar 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Lock Taksbar Setting"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Tag             =   "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
               Top             =   720
               Width           =   2055
            End
            Begin VB.CheckBox FullPathAddress 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Show Full Path at Address Bar"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Tag             =   "Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState"
               Top             =   1200
               Width           =   3375
            End
            Begin VB.CheckBox NameNumericTail 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Remove the Tildes in Short Filenames ""~"""
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Tag             =   "SYSTEM\CurrentControlSet\Control\FileSystem"
               Top             =   1440
               Width           =   3855
            End
            Begin VB.CheckBox NoViewContextMenu 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Remove Windows Explorer's default context menu"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   21
               Tag             =   "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
               Top             =   1680
               Width           =   4155
            End
            Begin VB.CheckBox NoSaveSettings 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Don't save settings at exit"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   20
               Tag             =   "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
               Top             =   1920
               Width           =   2655
            End
         End
         Begin VB.Frame fSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "System"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1335
            Left            =   360
            TabIndex        =   14
            Top             =   0
            Width           =   3975
            Begin VB.CheckBox DisableRegistryTools 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Disable Registry Editor"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Tag             =   "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
               Top             =   240
               Width           =   2295
            End
            Begin VB.CheckBox DisableTaskMgr 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Disable Task Manager"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   17
               Tag             =   "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
               Top             =   480
               Width           =   3255
            End
            Begin VB.CheckBox DisableCMD 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Disable CMD"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Tag             =   "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
               Top             =   720
               Width           =   2055
            End
            Begin VB.CheckBox NoFolderOptions 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Disable Folder Options"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Tag             =   "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
               Top             =   960
               Width           =   2895
            End
         End
         Begin VB.Frame fStartMenu 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Desktop and Start Menu"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1215
            Left            =   360
            TabIndex        =   8
            Top             =   1320
            Width           =   3975
            Begin VB.CheckBox RestrictRun 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Enable Restrict Run"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   13
               Tag             =   "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
               Top             =   240
               Width           =   2055
            End
            Begin VB.CheckBox NoUserNameInStartMenu 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Remove user name from Start menu"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Left            =   120
               TabIndex        =   12
               Tag             =   "Microsoft\Windows\CurrentVersion\Policies\Explorer"
               Top             =   3000
               Width           =   2775
            End
            Begin VB.CheckBox NoFind 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Remove Search form Start menu"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   11
               Tag             =   "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
               Top             =   720
               Width           =   3135
            End
            Begin VB.CheckBox NoDesktop 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Hide and disable items on desktop"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   10
               Top             =   480
               Width           =   3255
            End
            Begin VB.CheckBox NoRecentDocsHistory 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Delete Recent History"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   9
               Tag             =   "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
               Top             =   960
               Width           =   3075
            End
         End
         Begin prjBeeAV.AkuSayangIbu cSave 
            Height          =   375
            Left            =   7320
            TabIndex        =   50
            Top             =   3600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            BtnStyle        =   2
            Caption         =   "Apply Setting"
         End
      End
      Begin VB.PictureBox picTool 
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
         Height          =   4335
         Index           =   2
         Left            =   0
         ScaleHeight     =   4335
         ScaleWidth      =   9135
         TabIndex        =   5
         Top             =   600
         Width           =   9135
         Begin prjBeeAV.AkuSayangIbu jcbutton5 
            Height          =   855
            Left            =   600
            TabIndex        =   43
            Top             =   840
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   1508
            BtnStyle        =   2
            Caption         =   "Folder Locking"
         End
         Begin prjBeeAV.ucListView lvStartup 
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   19076
            _ExtentY        =   7223
            Style           =   4
            StyleEx         =   33
         End
         Begin prjBeeAV.AkuSayangIbu jcbutton6 
            Height          =   855
            Left            =   600
            TabIndex        =   44
            Top             =   2160
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   1508
            BtnStyle        =   2
            Caption         =   "Fixed Registry"
         End
         Begin prjBeeAV.AkuSayangIbu jcbutton7 
            Height          =   855
            Left            =   5280
            TabIndex        =   45
            Top             =   840
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   1508
            BtnStyle        =   2
            Caption         =   "ICan Write"
         End
         Begin prjBeeAV.AkuSayangIbu jcbutton10 
            Height          =   855
            Left            =   5280
            TabIndex        =   46
            Top             =   2160
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   1508
            BtnStyle        =   2
            Caption         =   "Virtual Keyboard"
         End
      End
      Begin VB.PictureBox picTool 
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
         Height          =   4335
         Index           =   1
         Left            =   0
         ScaleHeight     =   4335
         ScaleWidth      =   9135
         TabIndex        =   3
         Top             =   600
         Width           =   9135
         Begin prjBeeAV.AkuSayangIbu cLock 
            Height          =   495
            Left            =   120
            TabIndex        =   47
            Top             =   3000
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   873
            BtnStyle        =   2
            Caption         =   "Lock Drive"
         End
         Begin prjBeeAV.ucListView lvDlock 
            Height          =   2775
            Left            =   120
            TabIndex        =   4
            Top             =   0
            Width           =   8820
            _ExtentX        =   15558
            _ExtentY        =   4895
            Border          =   0
            StyleEx         =   37
         End
         Begin prjBeeAV.AkuSayangIbu cULock 
            Height          =   495
            Left            =   2520
            TabIndex        =   48
            Top             =   3000
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   873
            BtnStyle        =   2
            Caption         =   "Unlock Drive"
         End
      End
      Begin VB.PictureBox picTool 
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
         Height          =   4335
         Index           =   0
         Left            =   0
         ScaleHeight     =   4335
         ScaleWidth      =   9135
         TabIndex        =   1
         Top             =   600
         Width           =   9135
         Begin prjBeeAV.ucListView lvProcess 
            Height          =   4095
            Left            =   120
            TabIndex        =   2
            Top             =   0
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   7223
            Border          =   0
            StyleEx         =   33
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   0
         X1              =   120
         X2              =   9000
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Tool0 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Process Manager"
         BeginProperty Font 
            Name            =   "Futura Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   120
         TabIndex        =   42
         Top             =   120
         Width           =   1560
      End
      Begin VB.Label Tool3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Registry Tweaker"
         BeginProperty Font 
            Name            =   "Futura Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   5040
         TabIndex        =   41
         Top             =   120
         Width           =   1515
      End
      Begin VB.Label Tool1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Drive Locker"
         BeginProperty Font 
            Name            =   "Futura Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   1920
         TabIndex        =   40
         Top             =   120
         Width           =   1110
      End
      Begin VB.Label Tool2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "StartUp Manager"
         BeginProperty Font 
            Name            =   "Futura Md BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   3240
         TabIndex        =   39
         Top             =   120
         Width           =   1575
      End
   End
   Begin prjBeeAV.jcbutton cNextM1 
      Height          =   495
      Left            =   10980
      TabIndex        =   36
      Top             =   8160
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      Caption         =   ""
      PictureNormal   =   "fruit.frx":16B8A1
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin prjBeeAV.jcbutton cBackM1 
      Height          =   495
      Left            =   240
      TabIndex        =   38
      Top             =   8160
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      Caption         =   ""
      PictureNormal   =   "fruit.frx":16BFB3
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin prjBeeAV.jcbutton cBackAbout 
      Height          =   495
      Left            =   3360
      TabIndex        =   113
      Top             =   7440
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ButtonStyle     =   7
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      Caption         =   ""
      PictureNormal   =   "fruit.frx":16C6C5
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin prjBeeAV.jcbutton cForwardAbout 
      Height          =   495
      Left            =   14100
      TabIndex        =   114
      Top             =   7440
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      Caption         =   ""
      PictureNormal   =   "fruit.frx":16CDD7
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin prjBeeAV.jcbutton cmdBack 
      Height          =   495
      Left            =   4200
      TabIndex        =   122
      Top             =   7200
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ButtonStyle     =   7
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      Caption         =   ""
      PictureNormal   =   "fruit.frx":16D4E9
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin prjBeeAV.jcbutton cmdForward 
      Height          =   495
      Left            =   14940
      TabIndex        =   123
      Top             =   7200
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      Caption         =   ""
      PictureNormal   =   "fruit.frx":16DBFB
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin prjBeeAV.jcbutton cmdSkip 
      Height          =   375
      Left            =   13200
      TabIndex        =   163
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14933984
      Caption         =   "Skip"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin prjBeeAV.jcbutton cmdPause 
      Height          =   375
      Left            =   10440
      TabIndex        =   164
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14933984
      Caption         =   "Pause"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin prjBeeAV.jcbutton cmdStop 
      Height          =   375
      Left            =   7680
      TabIndex        =   165
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14933984
      Caption         =   "Stop"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Virus Definitions"
      BeginProperty Font 
         Name            =   "Control Freak"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   5760
      TabIndex        =   198
      Top             =   7200
      Width           =   1935
   End
   Begin VB.Label lbVirusDef 
      BackStyle       =   0  'Transparent
      Caption         =   "21 Nopember 2012"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   7320
      TabIndex        =   197
      Top             =   7200
      Width           =   2175
   End
   Begin VB.Label lScan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Scan Metode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3480
      TabIndex        =   124
      Top             =   7320
      Width           =   10215
   End
   Begin VB.Label lAbout 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Diyusof Team Work"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4440
      TabIndex        =   115
      Top             =   7920
      Width           =   10095
   End
   Begin VB.Label lTool 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Process Manager"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   37
      Top             =   8280
      Width           =   10215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cBackM1_Click()

End Sub

Private Sub cmdStop_Click()

End Sub

Private Sub jcbutton8_Click()

End Sub

Private Sub l_Click(Index As Integer)

End Sub

Private Sub Label30_Click()

End Sub

Private Sub picAbout_Click(Index As Integer)

End Sub
