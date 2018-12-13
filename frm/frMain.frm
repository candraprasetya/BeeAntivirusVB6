VERSION 5.00
Begin VB.Form frMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dasanggra"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picMainMenu 
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
      Height          =   5175
      Left            =   0
      Picture         =   "frMain.frx":058A
      ScaleHeight     =   5175
      ScaleWidth      =   2415
      TabIndex        =   0
      Top             =   960
      Width           =   2415
      Begin prjDAA.jcbutton Menu 
         Height          =   855
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1508
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16641248
         Caption         =   "Scan"
         PictureNormal   =   "frMain.frx":357CC
         PictureAlign    =   0
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjDAA.jcbutton Menu 
         Height          =   855
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1508
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16641248
         Caption         =   "Tools"
         PictureNormal   =   "frMain.frx":37C1E
         PictureAlign    =   0
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjDAA.jcbutton Menu 
         Height          =   855
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1508
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16641248
         Caption         =   "Setting"
         PictureNormal   =   "frMain.frx":3A070
         PictureAlign    =   0
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjDAA.jcbutton Menu 
         Height          =   855
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   3000
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1508
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16641248
         Caption         =   "Quarantine"
         PictureNormal   =   "frMain.frx":3C4C2
         PictureAlign    =   0
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjDAA.jcbutton Menu 
         Height          =   855
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   3960
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1508
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16641248
         Caption         =   "About"
         PictureNormal   =   "frMain.frx":3E914
         PictureAlign    =   0
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
   End
   Begin VB.Timer tmFD 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1800
      Top             =   5760
   End
   Begin VB.PictureBox pic7 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1440
      Picture         =   "frMain.frx":40D66
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   66
      Top             =   5880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic6 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1200
      Picture         =   "frMain.frx":412F0
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   65
      Top             =   5880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picBuff 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1680
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   53
      Top             =   5880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   960
      Picture         =   "frMain.frx":4187A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   44
      Top             =   5880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.ListBox drvList 
      Height          =   480
      ItemData        =   "frMain.frx":41CBC
      Left            =   480
      List            =   "frMain.frx":41CBE
      TabIndex        =   36
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox pic4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   720
      Picture         =   "frMain.frx":41CC0
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   10
      Top             =   5880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   480
      Picture         =   "frMain.frx":4224A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   9
      Top             =   5880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   240
      Picture         =   "frMain.frx":427D4
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   8
      Top             =   5880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   0
      Picture         =   "frMain.frx":42D5E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer tmBuff 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1800
      Top             =   5400
   End
   Begin VB.Timer tmSpeed 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   5400
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   2400
      Picture         =   "frMain.frx":431A0
      ScaleHeight     =   5175
      ScaleWidth      =   9375
      TabIndex        =   1
      Top             =   960
      Width           =   9375
      Begin VB.PictureBox Menu0 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4935
         Left            =   120
         ScaleHeight     =   4935
         ScaleWidth      =   9135
         TabIndex        =   11
         Top             =   120
         Width           =   9135
         Begin prjDAA.jcbutton cmdBack 
            Height          =   495
            Left            =   0
            TabIndex        =   80
            Top             =   0
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
            PictureNormal   =   "frMain.frx":E0564
            PictureEffectOnOver=   0
            PictureEffectOnDown=   0
            CaptionEffects  =   0
            TooltipBackColor=   0
            ColorScheme     =   3
         End
         Begin prjDAA.jcbutton cmdForward 
            Height          =   495
            Left            =   8640
            TabIndex        =   82
            Top             =   0
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
            PictureNormal   =   "frMain.frx":E15B6
            PictureEffectOnOver=   0
            PictureEffectOnDown=   0
            CaptionEffects  =   0
            TooltipBackColor=   0
            ColorScheme     =   3
         End
         Begin VB.PictureBox picScan 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4335
            Index           =   0
            Left            =   0
            ScaleHeight     =   4335
            ScaleWidth      =   9015
            TabIndex        =   12
            Top             =   480
            Width           =   9015
            Begin VB.PictureBox picDetailed 
               BackColor       =   &H00FFFFFF&
               Height          =   2655
               Left            =   120
               ScaleHeight     =   2595
               ScaleWidth      =   8835
               TabIndex        =   13
               Top             =   1680
               Width           =   8895
               Begin prjDAA.ProgressBar pScan 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   14
                  Top             =   960
                  Width           =   8655
                  _ExtentX        =   15266
                  _ExtentY        =   661
                  Value           =   0
                  Theme           =   8
                  TextStyle       =   3
                  BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "0%"
                  TextEffectColor =   16777215
               End
               Begin prjDAA.jcbutton cmdStop 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   15
                  Top             =   1800
                  Width           =   2535
                  _ExtentX        =   4471
                  _ExtentY        =   661
                  ButtonStyle     =   3
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BackColor       =   16641248
                  Caption         =   "STOP"
                  PictureEffectOnOver=   0
                  PictureEffectOnDown=   0
                  CaptionEffects  =   3
                  TooltipBackColor=   0
               End
               Begin prjDAA.jcbutton cmdSkip 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   16
                  Top             =   2160
                  Width           =   2535
                  _ExtentX        =   4471
                  _ExtentY        =   661
                  ButtonStyle     =   3
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BackColor       =   16641248
                  Caption         =   "SKIP"
                  PictureEffectOnOver=   0
                  PictureEffectOnDown=   0
                  CaptionEffects  =   3
                  TooltipBackColor=   0
               End
               Begin prjDAA.jcbutton cmdPause 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   17
                  Top             =   1440
                  Width           =   2535
                  _ExtentX        =   4471
                  _ExtentY        =   661
                  ButtonStyle     =   3
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BackColor       =   16641248
                  Caption         =   "PAUSE"
                  PictureEffectOnOver=   0
                  PictureEffectOnDown=   0
                  CaptionEffects  =   3
                  TooltipBackColor=   0
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
                  Height          =   255
                  Index           =   24
                  Left            =   6480
                  TabIndex        =   79
                  Top             =   480
                  Width           =   135
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
                  Height          =   255
                  Index           =   23
                  Left            =   6480
                  TabIndex        =   78
                  Top             =   240
                  Width           =   135
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
                  Height          =   255
                  Index           =   22
                  Left            =   2040
                  TabIndex        =   77
                  Top             =   600
                  Width           =   135
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
                  Height          =   255
                  Index           =   21
                  Left            =   2040
                  TabIndex        =   76
                  Top             =   360
                  Width           =   135
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
                  Height          =   255
                  Index           =   20
                  Left            =   2040
                  TabIndex        =   75
                  Top             =   120
                  Width           =   135
               End
               Begin VB.Label lTime 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "00:00:00"
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
                  Left            =   6600
                  TabIndex        =   38
                  Top             =   480
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
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   7
                  Left            =   5160
                  TabIndex        =   37
                  Top             =   480
                  Width           =   480
               End
               Begin VB.Label lblFile 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   975
                  Left            =   2880
                  TabIndex        =   26
                  Top             =   1560
                  Width           =   5895
                  WordWrap        =   -1  'True
               End
               Begin VB.Label l 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "File[s] Remaining"
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
                  Left            =   120
                  TabIndex        =   25
                  Top             =   120
                  Width           =   1665
               End
               Begin VB.Label lRem 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0 File[s]"
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
                  Left            =   2160
                  TabIndex        =   24
                  Top             =   120
                  Width           =   735
               End
               Begin VB.Label l 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "File[s] Scanned"
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
                  Left            =   120
                  TabIndex        =   23
                  Top             =   360
                  Width           =   1470
               End
               Begin VB.Label l 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "File[s] Ignored"
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
                  Index           =   4
                  Left            =   120
                  TabIndex        =   22
                  Top             =   600
                  Width           =   1365
               End
               Begin VB.Label lScanned 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0 File[s]"
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
                  Left            =   2160
                  TabIndex        =   21
                  Top             =   360
                  Width           =   735
               End
               Begin VB.Label lIgnored 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0 File[s]"
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
                  Left            =   2160
                  TabIndex        =   20
                  Top             =   600
                  Width           =   735
               End
               Begin VB.Label l 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Detected"
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
                  Index           =   5
                  Left            =   5160
                  TabIndex        =   19
                  Top             =   240
                  Width           =   840
               End
               Begin VB.Label lInFile 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0 File[s]"
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
                  Left            =   6600
                  TabIndex        =   18
                  Top             =   240
                  Width           =   735
               End
            End
            Begin prjDAA.jcbutton Scan 
               Height          =   975
               Index           =   0
               Left            =   120
               TabIndex        =   27
               Top             =   240
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   1720
               ButtonStyle     =   3
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   16641248
               Caption         =   "Computer"
               PictureNormal   =   "frMain.frx":E2608
               PictureEffectOnOver=   0
               PictureEffectOnDown=   0
               CaptionEffects  =   0
               TooltipBackColor=   0
            End
            Begin prjDAA.jcbutton Scan 
               Height          =   975
               Index           =   1
               Left            =   3120
               TabIndex        =   28
               Top             =   240
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   1720
               ButtonStyle     =   3
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   16641248
               Caption         =   "System"
               PictureNormal   =   "frMain.frx":E665A
               PictureEffectOnOver=   0
               PictureEffectOnDown=   0
               CaptionEffects  =   0
               TooltipBackColor=   0
            End
            Begin prjDAA.jcbutton Scan 
               Height          =   975
               Index           =   2
               Left            =   6120
               TabIndex        =   29
               Top             =   240
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   1720
               ButtonStyle     =   3
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   16641248
               Caption         =   "Custom"
               PictureNormal   =   "frMain.frx":EA6AC
               PictureEffectOnOver=   0
               PictureEffectOnDown=   0
               CaptionEffects  =   0
               TooltipBackColor=   0
            End
         End
         Begin VB.PictureBox picScan 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4215
            Index           =   1
            Left            =   0
            ScaleHeight     =   4215
            ScaleWidth      =   9135
            TabIndex        =   30
            Top             =   600
            Width           =   9135
            Begin prjDAA.ucListView lvMal 
               Height          =   4095
               Left            =   120
               TabIndex        =   31
               Top             =   120
               Width           =   8895
               _ExtentX        =   15266
               _ExtentY        =   7435
               Style           =   4
               StyleEx         =   37
            End
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
            Left            =   960
            TabIndex        =   81
            Top             =   120
            Width           =   7215
         End
      End
      Begin VB.PictureBox Menu2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4935
         Left            =   120
         ScaleHeight     =   4935
         ScaleWidth      =   9135
         TabIndex        =   32
         Top             =   120
         Width           =   9135
         Begin VB.CheckBox ck 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Enable Protection"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   71
            Top             =   2880
            Width           =   8895
         End
         Begin VB.CheckBox ck 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Auto Scan FlashDisk"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   70
            Top             =   2640
            Width           =   8895
         End
         Begin VB.CheckBox ck 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Enable Context Menu on Explorer"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   69
            Top             =   2400
            Width           =   8895
         End
         Begin VB.CheckBox ck 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Run on Start-Up"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   68
            Top             =   2160
            Width           =   8895
         End
         Begin VB.CheckBox ck 
            Caption         =   "Check1"
            Height          =   210
            Index           =   0
            Left            =   2280
            TabIndex        =   41
            Top             =   120
            Visible         =   0   'False
            Width           =   135
         End
         Begin prjDAA.jcbutton cmdSSet 
            Height          =   375
            Left            =   7920
            TabIndex        =   40
            Top             =   4440
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            ButtonStyle     =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16641248
            Caption         =   "Save"
            PictureEffectOnOver=   0
            PictureEffectOnDown=   0
            CaptionEffects  =   0
            TooltipBackColor=   0
         End
         Begin VB.CheckBox ck 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Use Heuristic Function for more Accurate Scanning"
            Height          =   210
            Index           =   3
            Left            =   120
            TabIndex        =   39
            Top             =   1080
            Width           =   8895
         End
         Begin VB.CheckBox ck 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Filter by Size"
            Height          =   210
            Index           =   2
            Left            =   120
            TabIndex        =   35
            Top             =   840
            Width           =   8895
         End
         Begin VB.CheckBox ck 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Filter by Extention"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   34
            Top             =   600
            Width           =   8895
         End
         Begin VB.Label l 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Application Setting : "
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
            Index           =   11
            Left            =   120
            TabIndex        =   67
            Top             =   1680
            Width           =   1995
         End
         Begin VB.Label l 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Scanning Setting : "
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
            Left            =   120
            TabIndex        =   33
            Top             =   120
            Width           =   1815
         End
      End
      Begin VB.PictureBox Menu3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4935
         Left            =   120
         ScaleHeight     =   4935
         ScaleWidth      =   9135
         TabIndex        =   42
         Top             =   120
         Width           =   9135
         Begin prjDAA.ucListView lvQuar 
            Height          =   4695
            Left            =   120
            TabIndex        =   43
            Top             =   120
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   8281
            Style           =   4
            StyleEx         =   33
         End
      End
      Begin VB.PictureBox Menu4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4935
         Left            =   120
         ScaleHeight     =   4935
         ScaleWidth      =   9135
         TabIndex        =   45
         Top             =   120
         Width           =   9135
         Begin VB.ListBox lstThanks 
            Height          =   4050
            ItemData        =   "frMain.frx":EE6FE
            Left            =   3240
            List            =   "frMain.frx":EE804
            TabIndex        =   88
            Top             =   360
            Width           =   3975
         End
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   4935
            Left            =   3000
            Picture         =   "frMain.frx":EECA6
            ScaleHeight     =   4935
            ScaleWidth      =   135
            TabIndex        =   87
            Top             =   0
            Width           =   135
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Definition Info"
            Height          =   1335
            Left            =   120
            TabIndex        =   50
            Top             =   3120
            Width           =   2775
            Begin VB.Label lDate 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Unknow"
               Height          =   210
               Left            =   120
               TabIndex        =   85
               Top             =   960
               Width           =   600
            End
            Begin VB.Label l 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Database Definition       :"
               Height          =   210
               Index           =   0
               Left            =   120
               TabIndex        =   84
               Top             =   720
               Width           =   1755
            End
            Begin VB.Label lDB 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               Height          =   210
               Left            =   120
               TabIndex        =   52
               Top             =   480
               Width           =   90
            End
            Begin VB.Label l 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Database Virus             :"
               Height          =   210
               Index           =   15
               Left            =   120
               TabIndex        =   51
               Top             =   240
               Width           =   1755
            End
         End
         Begin prjDAA.ucListView lvVirLst 
            Height          =   2295
            Left            =   120
            TabIndex        =   49
            Top             =   720
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   4048
            Style           =   4
            StyleEx         =   33
         End
         Begin VB.Image Image2 
            Height          =   2040
            Left            =   7320
            Picture         =   "frMain.frx":18C06A
            Stretch         =   -1  'True
            Top             =   360
            Width           =   1680
         End
         Begin VB.Label l 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   $"frMain.frx":19087A
            Height          =   420
            Index           =   9
            Left            =   3240
            TabIndex        =   89
            Top             =   4440
            Width           =   2250
         End
         Begin VB.Label l 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "More"
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
            Index           =   8
            Left            =   3240
            TabIndex        =   86
            Top             =   120
            Width           =   435
         End
         Begin VB.Label l 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Virus List :"
            Height          =   210
            Index           =   14
            Left            =   120
            TabIndex        =   48
            Top             =   480
            Width           =   780
         End
         Begin VB.Label l 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "© Dasanggra Software, 2010-2011"
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
            Index           =   12
            Left            =   6375
            TabIndex        =   47
            Top             =   4680
            Width           =   2700
         End
         Begin VB.Label l 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Programmer     : HermansyaH"
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
            Index           =   6
            Left            =   120
            TabIndex        =   46
            Top             =   120
            Width           =   2415
         End
      End
      Begin VB.PictureBox Menu1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4935
         Left            =   120
         ScaleHeight     =   4935
         ScaleWidth      =   9135
         TabIndex        =   56
         Top             =   120
         Width           =   9135
         Begin prjDAA.jcbutton cBackM1 
            Height          =   495
            Left            =   0
            TabIndex        =   58
            Top             =   0
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
            PictureNormal   =   "frMain.frx":1908A8
            PictureEffectOnOver=   0
            PictureEffectOnDown=   0
            CaptionEffects  =   0
            TooltipBackColor=   0
            ColorScheme     =   3
         End
         Begin prjDAA.jcbutton cNextM1 
            Height          =   495
            Left            =   8640
            TabIndex        =   57
            Top             =   0
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
            PictureNormal   =   "frMain.frx":1918FA
            PictureEffectOnOver=   0
            PictureEffectOnDown=   0
            CaptionEffects  =   0
            TooltipBackColor=   0
            ColorScheme     =   3
         End
         Begin VB.PictureBox PicTool 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4215
            Index           =   0
            Left            =   120
            ScaleHeight     =   4215
            ScaleWidth      =   8895
            TabIndex        =   59
            Top             =   600
            Width           =   8895
            Begin prjDAA.ucListView lvProcess 
               Height          =   4095
               Left            =   0
               TabIndex        =   60
               Top             =   120
               Width           =   8820
               _ExtentX        =   15558
               _ExtentY        =   7223
               StyleEx         =   33
            End
         End
         Begin VB.PictureBox PicTool 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4095
            Index           =   2
            Left            =   120
            ScaleHeight     =   4095
            ScaleWidth      =   8895
            TabIndex        =   72
            Top             =   720
            Width           =   8895
            Begin prjDAA.jcbutton cmdStartupMan 
               Height          =   855
               Left            =   120
               TabIndex        =   73
               Top             =   0
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   1508
               ButtonStyle     =   3
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   16641248
               Caption         =   "Dasanggra Startup Manager"
               PictureEffectOnOver=   0
               PictureEffectOnDown=   0
               CaptionEffects  =   0
               TooltipBackColor=   0
            End
            Begin prjDAA.jcbutton cmdRegTweak 
               Height          =   855
               Left            =   2280
               TabIndex        =   74
               Top             =   0
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   1508
               ButtonStyle     =   3
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   16641248
               Caption         =   "Dasanggra Registry Tweaker"
               PictureEffectOnOver=   0
               PictureEffectOnDown=   0
               CaptionEffects  =   0
               TooltipBackColor=   0
            End
         End
         Begin VB.PictureBox PicTool 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4215
            Index           =   1
            Left            =   120
            ScaleHeight     =   4215
            ScaleWidth      =   8895
            TabIndex        =   61
            Top             =   600
            Width           =   8895
            Begin prjDAA.jcbutton cULock 
               Height          =   495
               Left            =   2640
               TabIndex        =   64
               Top             =   3000
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   873
               ButtonStyle     =   3
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   16641248
               Caption         =   "Unlock Lock"
               PictureEffectOnOver=   0
               PictureEffectOnDown=   0
               CaptionEffects  =   0
               TooltipBackColor=   0
            End
            Begin prjDAA.jcbutton cLock 
               Height          =   495
               Left            =   120
               TabIndex        =   63
               Top             =   3000
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   873
               ButtonStyle     =   3
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   16641248
               Caption         =   "Lock Disk"
               PictureEffectOnOver=   0
               PictureEffectOnDown=   0
               CaptionEffects  =   0
               TooltipBackColor=   0
            End
            Begin prjDAA.ucListView lvDlock 
               Height          =   2775
               Left            =   0
               TabIndex        =   62
               Top             =   120
               Width           =   8820
               _ExtentX        =   15558
               _ExtentY        =   4895
               StyleEx         =   37
            End
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
            Left            =   960
            TabIndex        =   83
            Top             =   120
            Width           =   7215
         End
      End
   End
   Begin VB.Label lLink2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http:\www.herman-march.blogspot.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   4200
      TabIndex        =   55
      Top             =   720
      Width           =   3330
   End
   Begin VB.Label lLink1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http:\www.pramuka-sman3.blogspot.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   4200
      TabIndex        =   54
      Top             =   480
      Width           =   3450
   End
   Begin prjDAA.UniDialog udScan 
      Left            =   1320
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      FileFlags       =   2621444
      FolderFlags     =   835
      FileCustomFilter=   "frMain.frx":19294C
      FileDefaultExtension=   "frMain.frx":19296C
      FileFilter      =   "frMain.frx":19298C
      FileOpenTitle   =   "frMain.frx":1929D4
      FileSaveTitle   =   "frMain.frx":192A0C
      FolderMessage   =   "frMain.frx":192A44
   End
   Begin VB.Image Image1 
      Height          =   1035
      Left            =   0
      Picture         =   "frMain.frx":192A86
      Top             =   0
      Width           =   11790
   End
   Begin VB.Menu mnLvMal 
      Caption         =   "lvMal"
      Visible         =   0   'False
      Begin VB.Menu mnFAll 
         Caption         =   "Fix All Item[s]"
         Begin VB.Menu mnFADel 
            Caption         =   "Delete"
         End
         Begin VB.Menu mnFAQuar 
            Caption         =   "Quarantine"
         End
      End
      Begin VB.Menu mnFChk 
         Caption         =   "Fix Checked Item[s]"
         Begin VB.Menu mnFCDel 
            Caption         =   "Delete"
         End
         Begin VB.Menu mnFCQuar 
            Caption         =   "Quarantine"
         End
      End
      Begin VB.Menu bm1 
         Caption         =   "-"
      End
      Begin VB.Menu mnIAll 
         Caption         =   "Ignore All Item[s]"
      End
      Begin VB.Menu mnIChk 
         Caption         =   "Ignore Checked Item[s]"
      End
      Begin VB.Menu b283 
         Caption         =   "-"
      End
      Begin VB.Menu mnExMalLV 
         Caption         =   "Explore File"
      End
   End
   Begin VB.Menu mnlvQuar 
      Caption         =   "LVQuar"
      Visible         =   0   'False
      Begin VB.Menu mnRSel 
         Caption         =   "Restore Selected File[s]"
      End
      Begin VB.Menu mnRSelTo 
         Caption         =   "Restore Selected File[s] to . . ."
      End
      Begin VB.Menu bq1 
         Caption         =   "-"
      End
      Begin VB.Menu mnDeleteQ 
         Caption         =   "Delete All File[s]"
      End
      Begin VB.Menu mnDeleteSelQ 
         Caption         =   "Delete Selected File[s]"
      End
   End
   Begin VB.Menu mnlvProcess 
      Caption         =   "lvProcess"
      Visible         =   0   'False
      Begin VB.Menu mnKillProc 
         Caption         =   "Kill Selected Process"
      End
      Begin VB.Menu mnPause 
         Caption         =   "Pause Selected Process"
      End
      Begin VB.Menu mnResProc 
         Caption         =   "Resume Selected Process"
      End
      Begin VB.Menu blvP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnRefreshP 
         Caption         =   "Refresh"
      End
      Begin VB.Menu blvP2 
         Caption         =   "-"
      End
      Begin VB.Menu mnExpProc 
         Caption         =   "Explore Selected Process"
      End
   End
End
Attribute VB_Name = "frMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xToolIndexPic As Long

Private Sub cBackM1_Click()
If xToolIndexPic - 1 < 0 Then Exit Sub
PicTool(xToolIndexPic).Visible = False
PicTool(xToolIndexPic - 1).Visible = True
xToolIndexPic = xToolIndexPic - 1
Select Case xToolIndexPic
Case 0
ENUM_PROSES lvProcess, picBuff
cBackM1.Enabled = False
cNextM1.Enabled = True
lTool.Caption = "Process Manager"
Case 1
lTool.Caption = "Dlock Drive Locker"
lvDlock.ListItems.Clear
GetAllDrive
cLock.Enabled = False
cULock.Enabled = False
cBackM1.Enabled = True
cNextM1.Enabled = True
End Select
End Sub

Private Sub cLock_Click()
Static i As Integer

For i = 1 To lvDlock.ListItems.count
    If lvDlock.ListItems.Item(i).Checked = True Then
    BuatProt lvDlock.ListItems.Item(i).Text
    End If
Next
MsgBox "Finish !", vbInformation, "Dasanggra"
GetAllDrive
cLock.Enabled = False
End Sub

Private Sub cmdBack_Click()
picScan(0).Visible = True
picScan(1).Visible = False
lScan.Caption = "Scan Metode"
cmdBack.Enabled = False
cmdForward.Enabled = True
End Sub

Public Sub cmdForward_Click()
picScan(1).Visible = True
picScan(0).Visible = False
lScan.Caption = "Malware Detected"
cmdForward.Enabled = False
cmdBack.Enabled = True
End Sub

Private Sub cmdPause_Click()
If cmdPause.Caption = "PAUSE" Then
isPause = True
cmdPause.Caption = "RESUME"
Else
isPause = False
cmdPause.Caption = "PAUSE"
End If
End Sub

Private Sub cmdRegTweak_Click()
frRegTweak.Show 0, Me
End Sub

Private Sub cmdSkip_Click()
WithBuffer = False
tmBuff.Enabled = False
cmdStop.Enabled = True
cmdPause.Enabled = True
cmdSkip.Enabled = False
End Sub

Private Sub cmdSSet_Click()
SaveSettingF App.Path & "\Setting.ini"
ApplySetting
MsgBox "Done !", vbInformation, "Information !"
End Sub

Private Sub cmdStartupMan_Click()
frStartup.Show 0, frMain
End Sub

Private Sub cmdStop_Click()
StopScan = True
StopKumpulkan
    Scan(0).Visible = True
    Scan(1).Visible = True
    Scan(2).Visible = True
cmdSkip.Enabled = False
cmdPause.Enabled = True
cmdPause.Caption = "PAUSE"
End Sub

Private Sub cNextM1_Click()
If xToolIndexPic + 1 > 2 Then Exit Sub
PicTool(xToolIndexPic).Visible = False
PicTool(xToolIndexPic + 1).Visible = True
xToolIndexPic = xToolIndexPic + 1
Select Case xToolIndexPic
Case 1
lTool.Caption = "Dlock Drive Locker"
lvDlock.ListItems.Clear
GetAllDrive
cBackM1.Enabled = True
cNextM1.Enabled = True
cLock.Enabled = False
cULock.Enabled = False
Case 2
lTool.Caption = "Other Tools"
cBackM1.Enabled = True
cNextM1.Enabled = False
End Select
End Sub

Private Sub cULock_Click()
Static i As Integer

For i = 1 To lvDlock.ListItems.count
    If lvDlock.ListItems.Item(i).Checked = True Then
    UnProt lvDlock.ListItems.Item(i).Text
    End If
Next
MsgBox "Finish !", vbInformation, "Dasanggra"
GetAllDrive
cULock.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If StopScan = True Then
If isFromContext = True Then
    Unload frRTP
    Unload Me
    End
Else
    Cancel = 1
    Me.WindowState = 1
    Me.Hide
End If
Else
If isFromContext = True Then
    If MsgBox("Scanning is in progress." & vbCrLf & "Are you sure want to exit ?", vbExclamation + vbYesNo, "Warning !") = vbYes Then
    Unload frRTP
    Unload Me
    End
    Else
    Cancel = 1
    End If
Else
    If MsgBox("Scanning is in progress." & vbCrLf & "Do you want to stop Scanning ?", vbExclamation + vbYesNo, "Warning !") = vbYes Then
    Call cmdStop_Click
    Cancel = 1
    Else
    Cancel = 1
    End If
End If
End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lLink1.ForeColor = vbBlue
lLink2.ForeColor = vbBlue
End Sub

Private Sub Label1_Click()

End Sub

Private Sub lLink1_Click()
ShellExecute Me.hWnd, vbNullString, "http://www.pramuka-sman3.blogspot.com", vbNullString, "C:\", 1
End Sub

Private Sub lLink1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lLink1.ForeColor = vbRed
End Sub

Private Sub lLink2_Click()
ShellExecute Me.hWnd, vbNullString, "http://www.herman-march.blogspot.com", vbNullString, "C:\", 1
End Sub

Private Sub lLink2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lLink2.ForeColor = vbRed
End Sub

Private Sub lSpeed_Click()

End Sub

Private Sub lvDlock_ColumnClick(ByVal oColumn As cColumn)
oColumn.Sort
End Sub

Private Sub lvDlock_ItemCheck(ByVal oItem As cListItem, ByVal bCheck As Boolean)
If bCheck = True Then
    If lvDlock.ListItems.Item(oItem).SubItem(2).Text = "UnProtected" Then
        cLock.Enabled = True
        cULock.Enabled = False
    Else
        cLock.Enabled = False
        cULock.Enabled = True
    End If
End If
End Sub

Private Sub lvMal_ColumnClick(ByVal oColumn As cColumn)
oColumn.Sort
End Sub

Private Sub lvMal_ContextMenu(ByVal X As Single, ByVal Y As Single)
PopupMenu mnLvMal
End Sub

Private Sub lvProcess_ColumnClick(ByVal oColumn As cColumn)
oColumn.Sort
End Sub

Private Sub lvProcess_ContextMenu(ByVal X As Single, ByVal Y As Single)
PopupMenu mnlvProcess
End Sub

Private Sub lvQuar_ColumnClick(ByVal oColumn As cColumn)
oColumn.Sort
End Sub

Private Sub LVQuar_ContextMenu(ByVal X As Single, ByVal Y As Single)
PopupMenu mnlvQuar
End Sub

Private Sub lvVirLst_ColumnClick(ByVal oColumn As cColumn)
oColumn.Sort
End Sub

Public Sub Menu_Click(Index As Integer)
Select Case Index
Case 0
    Menu1.Visible = False
    Menu2.Visible = False
    Menu3.Visible = False
    Menu4.Visible = False
    Menu0.Visible = True
    Menu0.Top = 120
    Menu0.Left = 120
    picScan(0).Visible = True
    picScan(1).Visible = False
    cmdBack.Enabled = False
    cmdForward.Enabled = True
    lScan.Caption = "Scan Metode"
Case 1
    Menu0.Visible = False
    Menu2.Visible = False
    Menu3.Visible = False
    Menu4.Visible = False
    Menu1.Visible = True
    Menu1.Top = 120
    Menu1.Left = 120
    xToolIndexPic = 0
    ENUM_PROSES lvProcess, picBuff
    PicTool(0).Visible = True
    GetAllDrive
    PicTool(1).Visible = False
    PicTool(2).Visible = False
    cBackM1.Enabled = False
    cNextM1.Enabled = True
    lTool.Caption = "Process Manager"
Case 2
    Menu0.Visible = False
    Menu1.Visible = False
    Menu3.Visible = False
    Menu4.Visible = False
    Menu2.Visible = True
    Menu2.Top = 120
    Menu2.Left = 120
Case 3
    Menu0.Visible = False
    Menu1.Visible = False
    Menu2.Visible = False
    Menu4.Visible = False
    Menu3.Visible = True
    Menu3.Top = 120
    Menu3.Left = 120
    lvQuar.ListItems.Clear
    GetQuarFile
Case 4
    Menu0.Visible = False
    Menu1.Visible = False
    Menu2.Visible = False
    Menu3.Visible = False
    Menu4.Visible = True
    Menu4.Top = 120
    Menu4.Left = 120
End Select
End Sub

Private Sub mnDeleteQ_Click()
Static i As Integer

For i = 1 To lvQuar.ListItems.count
    HapusFile lvQuar.ListItems.Item(i).SubItem(3).Text
Next
MsgBox "Done !", vbInformation, "A-ViS"
GetQuarFile
End Sub

Private Sub mnDeleteSelQ_Click()
Static i As Integer

For i = 1 To lvQuar.ListItems.count
    If lvQuar.ListItems.Item(i).Selected = True Then
        HapusFile lvQuar.ListItems.Item(i).SubItem(3).Text
    End If
Next
MsgBox "Done !", vbInformation, "A-ViS"
GetQuarFile
End Sub

Private Sub mnExMalLV_Click()
Static i As Long

For i = 1 To lvMal.ListItems.count
    If lvMal.ListItems.Item(i).Selected = True Then
    ExploreTheFile lvMal.ListItems.Item(i).SubItem(2).Text
    End If
Next
End Sub

Private Sub mnExpProc_Click()
Static i As Long

For i = 1 To lvProcess.ListItems.count
    If lvProcess.ListItems.Item(i).Selected = True Then
    ExploreTheFile lvProcess.ListItems.Item(i).SubItem(9).Text
    End If
Next
End Sub

Private Sub mnFADel_Click()
FixVir FixAll, lvMal, del
End Sub

Private Sub mnFAQuar_Click()
FixVir FixAll, lvMal, Quar
End Sub

Private Sub mnFCDel_Click()
FixVir FixChk, lvMal, del
End Sub

Private Sub mnFCQuar_Click()
FixVir FixChk, lvMal, Quar
End Sub

Private Sub mnIAll_Click()
FixVir IgnAll, lvMal, None
End Sub

Private Sub mnIChk_Click()
FixVir IgnChk, lvMal, None
End Sub

Private Sub mnKillProc_Click()
Static i As Long
Dim PID    As Long
Dim sPath  As String

For i = 1 To lvProcess.ListItems.count
    If lvProcess.ListItems.Item(i).Selected = True Then
    If MsgBox("Do you want to kill this Process ?" & vbCrLf & "This Process Maybe needed by System", vbExclamation + vbYesNo, "Warning !") = vbNo Then Exit For
    PID = CLng(lvProcess.ListItems.Item(i).SubItem(3).Text)
    sPath = lvProcess.ListItems.Item(i).SubItem(9).Text
    If KillProses(PID, sPath, False, True) = True Then
            MsgBox "Success", vbInformation, "Informaton"
            ENUM_PROSES lvProcess, picBuff
        Else
            MsgBox "Error !", vbCritical, "Error !"
        End If
    Exit For
    End If
Next
End Sub

Private Sub mnPause_Click()
Static i As Long
Dim PID    As Long

For i = 1 To lvProcess.ListItems.count
    If lvProcess.ListItems.Item(i).Selected = True Then
      PID = CLng(lvProcess.ListItems.Item(i).SubItem(3).Text)
      lvProcess.ListItems.Item(i).SubItem(10).Text = SuspendProses(PID, True)
    End If
Next
End Sub

Private Sub mnRefreshP_Click()
ENUM_PROSES lvProcess, picBuff
End Sub

Private Sub mnResProc_Click()
Static i As Long
Dim PID    As Long

For i = 1 To lvProcess.ListItems.count
    If lvProcess.ListItems.Item(i).Selected = True Then
      PID = CLng(lvProcess.ListItems.Item(i).SubItem(3).Text)
      lvProcess.ListItems.Item(i).SubItem(10).Text = SuspendProses(PID, False)
    End If
Next
End Sub

Private Sub mnRSel_Click()
Static i As Integer
Static MalFrom As String
Static MalSend As String
If MsgBox("This File may be a virus that will Infected your Computer" & vbCrLf & "Are you sure want to restore it ?", vbExclamation + vbYesNo, "Warning !") = vbNo Then Exit Sub
For i = 1 To lvQuar.ListItems.count
    If lvQuar.ListItems.Item(i).Selected = True Then
    MalFrom = lvQuar.ListItems.Item(i).SubItem(3).Text
    MalSend = lvQuar.ListItems.Item(i).SubItem(2).Text
    RestoreFileQuar MalFrom, MalSend
    MsgBox MalSend
    End If
Next
MsgBox "Done !", vbInformation, "Dasanggra Antivirus"
GetQuarFile
End Sub

Private Sub mnRSelTo_Click()
Static i As Integer
Static MalFrom As String
Static MalSend As String
Static BFF As String

If MsgBox("This File may be a virus that will Infected your Computer" & vbCrLf & "Are you sure want to restore it ?", vbExclamation + vbYesNo, "Warning !") = vbNo Then Exit Sub
For i = 1 To lvQuar.ListItems.count
    If lvQuar.ListItems.Item(i).Selected = True Then
        MalFrom = lvQuar.ListItems.Item(i).SubItem(3).Text
        BFF = BrowseForFolder(Me.hWnd, "Choose Path :")
        If Len(BFF) > 0 Then
            MalSend = BFF & "\" & frMain.lvQuar.ListItems.Item(i).Text
            RestoreFileQuar MalFrom, MalSend
        End If
    End If
Next
MsgBox "Done !", vbInformation, "Dasanggra Antivirus"
GetQuarFile
End Sub

Public Sub Scan_Click(Index As Integer)
Static tR As Integer
AllReset
lvMal.ListItems.Clear
pScan.Value = 0
Menu(1).Enabled = False
Menu(2).Enabled = False
Menu(3).Enabled = False
cmdForward.Enabled = False
StopScan = False
If ck(7).Value = 1 Then
ScanRTPmod = False
End If
Select Case Index
Case 0
GetAllDrive
WithBuffer = True
tmBuff.Enabled = True
Scan(0).Enabled = False
Scan(1).Enabled = False
Scan(2).Enabled = False
cmdPause.Enabled = False
    For tR = 0 To drvList.ListCount - 1
        If WithBuffer = False Then Exit For
        BufferPath xDrive(tR), True
    Next
    cmdSkip.Enabled = False
    cmdPause.Enabled = True
    tR = 0
    tmBuff.Enabled = False
    tmSpeed.Enabled = True
    FileRemain = FileToScan
    pScan.Max = FileToScan
    FileToScan = 0
    For tR = 0 To drvList.ListCount - 1
        If StopScan = True Then Exit For
        KumpulkanFile xDrive(tR), True, True
    Next
    Call tmSpeed_Timer
    tmSpeed.Enabled = False
    Scan(0).Enabled = True
    Scan(1).Enabled = True
    Scan(2).Enabled = True
    cmdSkip.Enabled = True
    FinishJob
Case 1
GetQuickPath
    WithBuffer = True
    tmBuff.Enabled = True
    Scan(0).Enabled = False
    Scan(1).Enabled = False
    Scan(2).Enabled = False
    cmdPause.Enabled = False
    For tR = 0 To 2
        If WithBuffer = False Then Exit For
        BufferPath qPath(tR), True
    Next
    cmdPause.Enabled = True
    cmdSkip.Enabled = False
    tR = 0
    tmBuff.Enabled = False
    tmSpeed.Enabled = True
    FileRemain = FileToScan
    pScan.Max = FileToScan
    FileToScan = 0
    For tR = 0 To 2
        If StopScan = True Then Exit For
        KumpulkanFile qPath(tR), True, True
    Next
    Scan(1).Tag = "Scan"
    Call tmSpeed_Timer
    tmSpeed.Enabled = False
    Scan(0).Enabled = True
    Scan(1).Enabled = True
    Scan(2).Enabled = True
    cmdSkip.Enabled = True
    FinishJob
Case 2
If PathCustomScan <> vbNullString Then GoTo l_CustomScan
udScan.ShowFolder
If udScan.FolderPath <> "" Then
cPath = udScan.FolderPath
Else
GoTo l_Akhir
End If
Scan(0).Enabled = False
Scan(1).Enabled = False
Scan(2).Enabled = False
    cmdPause.Enabled = False
    WithBuffer = True
    tmBuff.Enabled = True
    BufferPath cPath, True
    cmdSkip.Enabled = False
    cmdPause.Enabled = True
    tmBuff.Enabled = False
    tmSpeed.Enabled = True
    FileRemain = FileToScan
    pScan.Max = FileToScan
    FileToScan = 0
    KumpulkanFile cPath, True, True
    Call tmSpeed_Timer
    tmSpeed.Enabled = False
    Scan(0).Enabled = True
    Scan(1).Enabled = True
    Scan(2).Enabled = True
    cmdSkip.Enabled = True
    FinishJob
    GoTo l_Akhir
l_CustomScan:
Scan(0).Enabled = False
Scan(1).Enabled = False
Scan(2).Enabled = False
    cmdPause.Enabled = False
    WithBuffer = True
    tmBuff.Enabled = True
    BufferPath PathCustomScan, True
    cmdSkip.Enabled = False
    cmdPause.Enabled = True
    tmBuff.Enabled = False
    tmSpeed.Enabled = True
    FileRemain = FileToScan
    pScan.Max = FileToScan
    FileToScan = 0
    KumpulkanFile PathCustomScan, True, True
    Call tmSpeed_Timer
    tmSpeed.Enabled = False
    Scan(0).Enabled = True
    Scan(1).Enabled = True
    Scan(2).Enabled = True
    cmdSkip.Enabled = True
    FinishJob
End Select
l_Akhir:
Menu(1).Enabled = True
Menu(2).Enabled = True
Menu(3).Enabled = True
If ck(7).Value = 1 Then
ScanRTPmod = True
StopScan = True
cmdForward.Enabled = True
End If
End Sub

Private Sub tmBuff_Timer()
lblFile.Caption = "Getting information from your Computer" & vbCrLf & "File to scan " & FileToScan
End Sub

Private Sub tmFD_Timer()
Dim sDriveName          As String
Dim DriveLabel          As String
Dim nDriveNameLen       As Long
If StopScan = False Or isFromContext = True Then ' jika scan scan jalan
   tmFD.Enabled = False
   Exit Sub
End If
If AdakahFDBaru(LastFlashVolume) = True Then
   nDriveNameLen = 128
   sDriveName = String$(nDriveNameLen, 0)
   If GetVolumeInformationW(StrPtr(Chr(LastFlashVolume) & ":\"), StrPtr(sDriveName), nDriveNameLen, ByVal 0, ByVal 0, ByVal 0, 0, 0) Then
       DriveLabel = Left$(sDriveName, InStr(1, sDriveName, ChrW$(0)) - 1)
   Else
       DriveLabel = vbNullString
   End If
   TampilkanBalon Me, "New Removeable Disk Detected", "Information", NIIF_INFO
   PathCustomScan = Chr(LastFlashVolume) & ":\"
frMain.Show
frMain.Menu_Click (0)
frMain.Scan_Click (2)
End If
End Sub
Public Sub tmSpeed_Timer()
Static xPercent As String
If WithBuffer = True Then
    frMain.pScan.Value = FileScan
    xPercent = pScan.Value * 100 / pScan.Max
    If xPercent >= 10 Then
        pScan.Text = Format$(xPercent, "#0") & "%"
    Else
        pScan.Text = Format$(xPercent, "0") & "%"
    End If
    frMain.lRem.Caption = FileRemain & " File[s]"
Else
frMain.lRem.Caption = "Unknow"
End If
frMain.lScanned.Caption = FileScan & " File[s]"
frMain.lIgnored.Caption = FileIgnore & " File[s]"
frMain.lInFile.Caption = lvMal.ListItems.count & " File[s]"
If lvMal.ListItems.count > 0 Then frMain.lInFile.ForeColor = vbRed
frMain.lblFile.Caption = xScanPath
FileToScan = 0
Detik = Detik + 1
If Detik = 60 Then
   Detik = 0
   Menit = Menit + 1
End If
If Menit = 60 Then
   Menit = 0
   Jam = Jam + 1
End If
lTime.Caption = Format$(Jam, "00") & ":" & Format$(Menit, "00") & ":" & Format$(Detik, "00")
End Sub

Public Sub ScanFromCM()
If ValidFile(PathCustomScan) = True Then
Scan(0).Visible = False
Scan(1).Visible = False
Scan(2).Visible = False
    cmdPause.Enabled = False
    WithBuffer = True
    tmBuff.Enabled = True
    FileToScan = 1
    cmdSkip.Enabled = False
    cmdPause.Enabled = True
    tmBuff.Enabled = False
    tmSpeed.Enabled = True
    FileRemain = FileToScan
    pScan.Max = FileToScan
    FileToScan = 0
                FileScan = FileScan + 1
                FileToScan = FileToScan + 1
                FileRemain = FileRemain - 1
                
                If frMain.ck(1).Value = 1 Then
                    If isProperFile(PathCustomScan, "SYS LNK VBE HTM HTT EXE DLL VBS VMX TML .DB COM SCR BAT INF TML CMD TXT PIF MSI BMP") = True Then
                        Equal (PathCustomScan)
                    Else
                        FileIgnore = FileIgnore + 1
                    End If
                Else
                    Equal (PathCustomScan)
                End If
    Call tmSpeed_Timer
    tmSpeed.Enabled = False
    Scan(0).Visible = True
    Scan(1).Visible = True
    Scan(2).Visible = True
    cmdSkip.Enabled = True
    FinishJob
ElseIf ValidFolder(PathCustomScan) = True Then
Scan(0).Visible = False
Scan(1).Visible = False
Scan(2).Visible = False
    cmdPause.Enabled = False
    WithBuffer = True
    tmBuff.Enabled = True
    BufferPath PathCustomScan, True
    cmdSkip.Enabled = False
    cmdPause.Enabled = True
    tmBuff.Enabled = False
    tmSpeed.Enabled = True
    FileRemain = FileToScan
    pScan.Max = FileToScan
    FileToScan = 0
    KumpulkanFile PathCustomScan, True, True
    Call tmSpeed_Timer
    tmSpeed.Enabled = False
    Scan(0).Visible = True
    Scan(1).Visible = True
    Scan(2).Visible = True
    cmdSkip.Enabled = True
    FinishJob
Else
MsgBox "Can't Scan File or Folder !" & vbCrLf & _
"Invalid or Protected !", vbCritical, "Error !"
End
End If
End Sub

Private Sub uTabSonny1_Click(Index As Integer)

End Sub
