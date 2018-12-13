VERSION 5.00
Begin VB.Form frMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dasanggra Antivirus"
   ClientHeight    =   6225
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
   Icon            =   "frMain2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
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
      Picture         =   "frMain2.frx":058A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   75
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
      Picture         =   "frMain2.frx":0B14
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   74
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
      TabIndex        =   59
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
      Picture         =   "frMain2.frx":109E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   48
      Top             =   5880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.ListBox drvList 
      Height          =   480
      ItemData        =   "frMain2.frx":1628
      Left            =   480
      List            =   "frMain2.frx":162A
      TabIndex        =   40
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
      Picture         =   "frMain2.frx":162C
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
      Picture         =   "frMain2.frx":1BB6
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
      Picture         =   "frMain2.frx":2140
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
      Picture         =   "frMain2.frx":26CA
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
      Picture         =   "frMain2.frx":2C54
      ScaleHeight     =   5175
      ScaleWidth      =   9375
      TabIndex        =   1
      Top             =   1080
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
         Begin VB.PictureBox picScan 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4935
            Index           =   0
            Left            =   120
            ScaleHeight     =   4935
            ScaleWidth      =   8895
            TabIndex        =   12
            Top             =   0
            Width           =   8895
            Begin VB.PictureBox picDetailed 
               BackColor       =   &H00FFFFFF&
               Height          =   3615
               Left            =   0
               ScaleHeight     =   3555
               ScaleWidth      =   8835
               TabIndex        =   13
               Top             =   1200
               Width           =   8895
               Begin prjDAA.jcbutton cmdVResult 
                  Height          =   1455
                  Left            =   4800
                  TabIndex        =   14
                  Top             =   120
                  Width           =   3975
                  _ExtentX        =   7011
                  _ExtentY        =   2566
                  ButtonStyle     =   3
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   24
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BackColor       =   16641248
                  Caption         =   "View Result >>>"
                  PictureEffectOnOver=   0
                  PictureEffectOnDown=   0
                  CaptionEffects  =   0
                  TooltipBackColor=   0
               End
               Begin prjDAA.ProgressBar pScan 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   15
                  Top             =   1920
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
                  TabIndex        =   16
                  Top             =   2760
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
                  TabIndex        =   17
                  Top             =   3120
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
                  TabIndex        =   18
                  Top             =   2400
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
               Begin VB.Label lTime 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   ": 00:00:00"
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
                  TabIndex        =   42
                  Top             =   1320
                  Width           =   870
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
                  Left            =   120
                  TabIndex        =   41
                  Top             =   1320
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
                  TabIndex        =   29
                  Top             =   2520
                  Width           =   5895
                  WordWrap        =   -1  'True
               End
               Begin VB.Label l 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Speed"
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
                  Left            =   120
                  TabIndex        =   28
                  Top             =   120
                  Width           =   615
               End
               Begin VB.Label lSpeed 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   ": 0 File[s]/s"
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
                  TabIndex        =   27
                  Top             =   120
                  Width           =   1005
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
                  TabIndex        =   26
                  Top             =   360
                  Width           =   1665
               End
               Begin VB.Label lRem 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   ": 0 File[s]"
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
                  TabIndex        =   25
                  Top             =   360
                  Width           =   855
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
                  TabIndex        =   24
                  Top             =   600
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
                  TabIndex        =   23
                  Top             =   840
                  Width           =   1365
               End
               Begin VB.Label lScanned 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   ": 0 File[s]"
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
                  TabIndex        =   22
                  Top             =   600
                  Width           =   855
               End
               Begin VB.Label lIgnored 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   ": 0 File[s]"
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
                  Top             =   840
                  Width           =   855
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
                  Left            =   120
                  TabIndex        =   20
                  Top             =   1080
                  Width           =   840
               End
               Begin VB.Label lInFile 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   ": 0 File[s]"
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
                  TabIndex        =   19
                  Top             =   1080
                  Width           =   855
               End
            End
            Begin prjDAA.jcbutton Scan 
               Height          =   975
               Index           =   0
               Left            =   120
               TabIndex        =   30
               Top             =   120
               Width           =   2415
               _ExtentX        =   4260
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
               Caption         =   "Full Scan"
               PictureEffectOnOver=   0
               PictureEffectOnDown=   0
               CaptionEffects  =   0
               TooltipBackColor=   0
            End
            Begin prjDAA.jcbutton Scan 
               Height          =   975
               Index           =   1
               Left            =   3240
               TabIndex        =   31
               Top             =   120
               Width           =   2415
               _ExtentX        =   4260
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
               Caption         =   "Quick Scan"
               PictureEffectOnOver=   0
               PictureEffectOnDown=   0
               CaptionEffects  =   0
               TooltipBackColor=   0
            End
            Begin prjDAA.jcbutton Scan 
               Height          =   975
               Index           =   2
               Left            =   6360
               TabIndex        =   32
               Top             =   120
               Width           =   2415
               _ExtentX        =   4260
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
               Caption         =   "Custom Scan"
               PictureEffectOnOver=   0
               PictureEffectOnDown=   0
               CaptionEffects  =   0
               TooltipBackColor=   0
            End
         End
         Begin VB.PictureBox picScan 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4935
            Index           =   1
            Left            =   120
            ScaleHeight     =   4935
            ScaleWidth      =   8895
            TabIndex        =   33
            Top             =   0
            Width           =   8895
            Begin prjDAA.ucListView lvMal 
               Height          =   4215
               Left            =   120
               TabIndex        =   34
               Top             =   120
               Width           =   8655
               _ExtentX        =   15266
               _ExtentY        =   7435
               Style           =   4
               StyleEx         =   37
            End
            Begin prjDAA.jcbutton cmdBack 
               Height          =   375
               Left            =   240
               TabIndex        =   35
               Top             =   4440
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               ButtonStyle     =   3
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   16641248
               Caption         =   "Back"
               PictureEffectOnOver=   0
               PictureEffectOnDown=   0
               CaptionEffects  =   0
               TooltipBackColor=   0
            End
         End
      End
      Begin VB.PictureBox Menu2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4935
         Left            =   120
         ScaleHeight     =   4935
         ScaleWidth      =   9135
         TabIndex        =   36
         Top             =   120
         Width           =   9135
         Begin VB.CheckBox ck 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Enable Protection"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   80
            Top             =   2880
            Width           =   8895
         End
         Begin VB.CheckBox ck 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Auto Scan FlashDisk"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   79
            Top             =   2640
            Width           =   8895
         End
         Begin VB.CheckBox ck 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Enable Context Menu on Explorer"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   78
            Top             =   2400
            Width           =   8895
         End
         Begin VB.CheckBox ck 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Run on Start-Up"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   77
            Top             =   2160
            Width           =   8895
         End
         Begin VB.CheckBox ck 
            Caption         =   "Check1"
            Height          =   210
            Index           =   0
            Left            =   2280
            TabIndex        =   45
            Top             =   120
            Visible         =   0   'False
            Width           =   135
         End
         Begin prjDAA.jcbutton cmdSSet 
            Height          =   375
            Left            =   7920
            TabIndex        =   44
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
            TabIndex        =   43
            Top             =   1080
            Width           =   8895
         End
         Begin VB.CheckBox ck 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Filter by Size"
            Height          =   210
            Index           =   2
            Left            =   120
            TabIndex        =   39
            Top             =   840
            Width           =   8895
         End
         Begin VB.CheckBox ck 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Filter by Extention"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   38
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
            TabIndex        =   76
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
            TabIndex        =   37
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
         TabIndex        =   46
         Top             =   120
         Width           =   9135
         Begin prjDAA.ucListView lvQuar 
            Height          =   4695
            Left            =   120
            TabIndex        =   47
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
         TabIndex        =   49
         Top             =   120
         Width           =   9135
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   4935
            Left            =   3720
            Picture         =   "frMain2.frx":A0018
            ScaleHeight     =   4935
            ScaleWidth      =   135
            TabIndex        =   81
            Top             =   0
            Width           =   135
         End
         Begin prjDAA.jcbutton cmdADAATeam 
            Height          =   375
            Left            =   120
            TabIndex        =   60
            Top             =   4320
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   661
            ButtonStyle     =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16641248
            Caption         =   "Dasanggra Antivirus Credit"
            PictureEffectOnOver=   0
            PictureEffectOnDown=   0
            CaptionEffects  =   0
            TooltipBackColor=   0
         End
         Begin VB.Frame f 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Definition Info"
            Height          =   615
            Left            =   3960
            TabIndex        =   56
            Top             =   2760
            Width           =   5055
            Begin VB.Label lDB 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ": 0"
               Height          =   210
               Left            =   1560
               TabIndex        =   58
               Top             =   240
               Width           =   180
            End
            Begin VB.Label l 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Database Virus"
               Height          =   210
               Index           =   15
               Left            =   120
               TabIndex        =   57
               Top             =   240
               Width           =   1125
            End
         End
         Begin prjDAA.ucListView lvVirLst 
            Height          =   2295
            Left            =   3960
            TabIndex        =   55
            Top             =   360
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   4260
            Style           =   4
            StyleEx         =   33
         End
         Begin VB.Label l 
            BackStyle       =   0  'Transparent
            Caption         =   $"frMain2.frx":13D3DC
            Height          =   1215
            Index           =   16
            Left            =   3960
            TabIndex        =   82
            Top             =   3360
            Width           =   5055
         End
         Begin VB.Image Image2 
            Height          =   2280
            Left            =   6840
            Picture         =   "frMain2.frx":13D4B0
            Stretch         =   -1  'True
            Top             =   360
            Width           =   2040
         End
         Begin VB.Label l 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Virus List :"
            Height          =   210
            Index           =   14
            Left            =   3960
            TabIndex        =   54
            Top             =   120
            Width           =   780
         End
         Begin VB.Label l 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Some code was adopted from CMC PH#3 SC"
            Height          =   210
            Index           =   13
            Left            =   120
            TabIndex        =   53
            Top             =   4680
            Width           =   3240
         End
         Begin VB.Label l 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "© Copyright by Dasanggra Software, 2010-2011"
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
            Left            =   5280
            TabIndex        =   52
            Top             =   4680
            Width           =   3795
         End
         Begin VB.Label l 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   $"frMain2.frx":141CC0
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2940
            Index           =   8
            Left            =   120
            TabIndex        =   51
            Top             =   360
            Width           =   3390
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
            TabIndex        =   50
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
         TabIndex        =   63
         Top             =   120
         Width           =   9135
         Begin prjDAA.jcbutton cBackM1 
            Height          =   375
            Left            =   120
            TabIndex        =   65
            Top             =   4440
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            ButtonStyle     =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16641248
            Caption         =   "Back"
            PictureEffectOnOver=   0
            PictureEffectOnDown=   0
            CaptionEffects  =   0
            TooltipBackColor=   0
         End
         Begin prjDAA.jcbutton cNextM1 
            Height          =   375
            Left            =   8160
            TabIndex        =   64
            Top             =   4440
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            ButtonStyle     =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16641248
            Caption         =   "Next"
            PictureEffectOnOver=   0
            PictureEffectOnDown=   0
            CaptionEffects  =   0
            TooltipBackColor=   0
         End
         Begin VB.PictureBox PicTool 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4215
            Index           =   0
            Left            =   120
            ScaleHeight     =   4215
            ScaleWidth      =   8895
            TabIndex        =   66
            Top             =   120
            Width           =   8895
            Begin prjDAA.ucListView lvProcess 
               Height          =   3855
               Left            =   0
               TabIndex        =   68
               Top             =   240
               Width           =   8820
               _ExtentX        =   15558
               _ExtentY        =   6800
               StyleEx         =   33
            End
            Begin VB.Label l 
               AutoSize        =   -1  'True
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
               Height          =   210
               Index           =   9
               Left            =   0
               TabIndex        =   67
               Top             =   0
               Width           =   1455
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
            TabIndex        =   69
            Top             =   120
            Width           =   8895
            Begin prjDAA.jcbutton cULock 
               Height          =   495
               Left            =   2640
               TabIndex        =   73
               Top             =   3480
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
               TabIndex        =   72
               Top             =   3480
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
               TabIndex        =   71
               Top             =   240
               Width           =   8820
               _ExtentX        =   15558
               _ExtentY        =   4895
               StyleEx         =   37
            End
            Begin VB.Label l 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dlock"
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
               Index           =   10
               Left            =   0
               TabIndex        =   70
               Top             =   0
               Width           =   450
            End
         End
         Begin VB.PictureBox PicTool 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4215
            Index           =   2
            Left            =   120
            ScaleHeight     =   4215
            ScaleWidth      =   8895
            TabIndex        =   83
            Top             =   120
            Width           =   8895
            Begin prjDAA.jcbutton cmdStartupMan 
               Height          =   855
               Left            =   120
               TabIndex        =   85
               Top             =   360
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
               TabIndex        =   86
               Top             =   360
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
            Begin VB.Label l 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Other Tools"
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
               Index           =   17
               Left            =   0
               TabIndex        =   84
               Top             =   0
               Width           =   975
            End
         End
      End
   End
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
      Height          =   4215
      Left            =   0
      ScaleHeight     =   4215
      ScaleWidth      =   2295
      TabIndex        =   0
      Top             =   1080
      Width           =   2295
      Begin prjDAA.jcbutton Menu 
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1296
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
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjDAA.jcbutton Menu 
         Height          =   735
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1296
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
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjDAA.jcbutton Menu 
         Height          =   735
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1296
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
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjDAA.jcbutton Menu 
         Height          =   735
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   2520
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1296
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
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjDAA.jcbutton Menu 
         Height          =   735
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   3360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1296
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
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
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
      Left            =   7560
      TabIndex        =   62
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
      Left            =   7560
      TabIndex        =   61
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
      FileCustomFilter=   "frMain2.frx":141DF3
      FileDefaultExtension=   "frMain2.frx":141E13
      FileFilter      =   "frMain2.frx":141E33
      FileOpenTitle   =   "frMain2.frx":141E7B
      FileSaveTitle   =   "frMain2.frx":141EB3
      FolderMessage   =   "frMain2.frx":141EEB
   End
   Begin VB.Image Image1 
      Height          =   1035
      Left            =   0
      Picture         =   "frMain2.frx":141F2D
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
Dim xToolIndexPic As Long

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
Case 1
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

Private Sub cmdADAATeam_Click()
frIcon.Show 1, frMain
End Sub

Private Sub cmdBack_Click()
picScan(0).Visible = True
picScan(1).Visible = False
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

Private Sub cmdVResult_Click()
picScan(0).Visible = False
picScan(1).Visible = True
End Sub

Private Sub cNextM1_Click()
If xToolIndexPic + 1 > 2 Then Exit Sub
PicTool(xToolIndexPic).Visible = False
PicTool(xToolIndexPic + 1).Visible = True
xToolIndexPic = xToolIndexPic + 1
Select Case xToolIndexPic
Case 1
lvDlock.ListItems.Clear
GetAllDrive
cBackM1.Enabled = True
cNextM1.Enabled = True
cLock.Enabled = False
cULock.Enabled = False
Case 2
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

Public Sub Form_Load()
BuildLV
If ValidFile(App.Path & "\VirDb\db1.dbC") = True Then
ReadDb App.Path & "\VirDb\db1.dbC"
Else
MsgBox "Database Virus Not Found !", vbCritical + vbOKOnly, "Dasanggra Error !"
Unload frRTP
Unload Me
End
End If
InitPHPattern
LoadDataIcon
GetSettingF App.Path & "\Setting.ini"
ApplySetting
frMain.Menu_Click (0)
xToolIndexPic = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
If isFromContext = True Then
Unload frRTP
Unload Me
End
Else
Cancel = 1
Me.WindowState = 1
Me.Hide
End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lLink1.ForeColor = vbBlue
lLink2.ForeColor = vbBlue
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
    picScan(0).Top = 0
    picScan(0).Left = 120
    picScan(1).Visible = False
    picScan(1).Top = 0
    picScan(1).Left = 120
Case 1
    Menu0.Visible = False
    Menu2.Visible = False
    Menu3.Visible = False
    Menu4.Visible = False
    Menu1.Visible = True
    Menu1.Top = 120
    Menu1.Left = 120
    ENUM_PROSES lvProcess, picBuff
    PicTool(0).Visible = True
    GetAllDrive
    PicTool(1).Visible = False
    PicTool(2).Visible = False
    cBackM1.Enabled = False
    cNextM1.Enabled = True
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
Static tR As Long
AllReset
lvMal.ListItems.Clear
pScan.Value = 0
cmdVResult.Enabled = False
If ck(7).Value = 1 Then
ScanRTPmod = False
End If
Select Case Index
Case 0
GetAllDrive
WithBuffer = True
tmBuff.Enabled = True
Scan(0).Visible = False
Scan(1).Visible = False
Scan(2).Visible = False
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
    Scan(0).Visible = True
    Scan(1).Visible = True
    Scan(2).Visible = True
    cmdSkip.Enabled = True
    FinishJob
Case 1
GetQuickPath
    WithBuffer = True
    tmBuff.Enabled = True
    Scan(0).Visible = False
    Scan(1).Visible = False
    Scan(2).Visible = False
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
    Scan(0).Visible = True
    Scan(1).Visible = True
    Scan(2).Visible = True
    cmdSkip.Enabled = True
    FinishJob
Case 2
udScan.ShowFolder
If udScan.FolderPath <> "" Then
cPath = udScan.FolderPath
Else
Exit Sub
End If
Scan(0).Visible = False
Scan(1).Visible = False
Scan(2).Visible = False
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
    Scan(0).Visible = True
    Scan(1).Visible = True
    Scan(2).Visible = True
    cmdSkip.Enabled = True
    FinishJob
    GoTo l_Akhir
End Select
l_Akhir:
cmdVResult.Enabled = True
If ck(7).Value = 1 Then
ScanRTPmod = True
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
frMain.ScanFromCM False
End If
End Sub
Public Sub tmSpeed_Timer()
Static xPercent As String
If WithBuffer = True Then
    frMain.pScan.Value = FileScan
    xPercent = pScan.Value * 100 / pScan.Max
    pScan.Text = Format(xPercent, "#0") & "%"
    frMain.lRem.Caption = ": " & FileRemain & " File[s]"
Else
frMain.lRem.Caption = ": Unknow"
End If
lSpeed.Caption = ": " & FileToScan & " File[s]/s"
frMain.lScanned.Caption = ": " & FileScan & " File[s]"
frMain.lIgnored.Caption = ": " & FileIgnore & " File[s]"
frMain.lInFile.Caption = ": " & lvMal.ListItems.count & " File[s]"
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
lTime.Caption = ": " & Format(Jam, "00") & ":" & Format(Menit, "00") & ":" & Format(Detik, "00")
End Sub

Public Sub ScanFromCM(isContexMenu As Boolean)
If ValidFile3(PathCustomScan) = False Then MsgBox "File or Folder is not Valid !", vbExclamation, "Error !": Exit Sub
AllReset
lvMal.ListItems.Clear
pScan.Value = 0
cmdVResult.Enabled = False
If isContexMenu = False Then
If ck(7).Value = 1 Then
ScanRTPmod = False
End If
End If
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
End Select
cmdVResult.Enabled = True
If isContexMenu = False Then
If ck(7).Value = 1 Then
ScanRTPmod = True
End If
End If
End Sub
