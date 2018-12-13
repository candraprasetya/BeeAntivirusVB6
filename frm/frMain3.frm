VERSION 5.00
Begin VB.Form frMain1 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Diyusof Antivirus : {QWERTYqwes,cmkj12388kausJJ}"
   ClientHeight    =   7440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11550
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frMain3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frMain3.frx":000C
   ScaleHeight     =   7440
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture20 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   9720
      Picture         =   "frMain3.frx":4510
      ScaleHeight     =   765
      ScaleWidth      =   1665
      TabIndex        =   63
      ToolTipText     =   "About"
      Top             =   1215
      Width           =   1695
   End
   Begin VB.PictureBox Picture19 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   7800
      Picture         =   "frMain3.frx":4C34
      ScaleHeight     =   765
      ScaleWidth      =   1905
      TabIndex        =   62
      ToolTipText     =   "Quarantine"
      Top             =   1215
      Width           =   1935
   End
   Begin VB.PictureBox Picture18 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   5880
      Picture         =   "frMain3.frx":56B9
      ScaleHeight     =   765
      ScaleWidth      =   1905
      TabIndex        =   61
      ToolTipText     =   "Tools"
      Top             =   1215
      Width           =   1935
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   120
      Picture         =   "frMain3.frx":5D54
      ScaleHeight     =   765
      ScaleWidth      =   1905
      TabIndex        =   60
      ToolTipText     =   "Overview"
      Top             =   1215
      Width           =   1935
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   2040
      Picture         =   "frMain3.frx":667D
      ScaleHeight     =   765
      ScaleWidth      =   1905
      TabIndex        =   59
      ToolTipText     =   "Scan Area"
      Top             =   1215
      Width           =   1935
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   3960
      Picture         =   "frMain3.frx":6F96
      ScaleHeight     =   765
      ScaleWidth      =   1905
      TabIndex        =   58
      ToolTipText     =   "Settings"
      Top             =   1215
      Width           =   1935
   End
   Begin VB.PictureBox picMenu7 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   9720
      Picture         =   "frMain3.frx":78A4
      ScaleHeight     =   765
      ScaleWidth      =   1665
      TabIndex        =   57
      ToolTipText     =   "About"
      Top             =   1215
      Width           =   1695
   End
   Begin VB.PictureBox picMenu6 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   7800
      Picture         =   "frMain3.frx":7E79
      ScaleHeight     =   765
      ScaleWidth      =   1905
      TabIndex        =   56
      ToolTipText     =   "Quarantine"
      Top             =   1215
      Width           =   1935
   End
   Begin VB.PictureBox picMenu4 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   5880
      Picture         =   "frMain3.frx":8697
      ScaleHeight     =   765
      ScaleWidth      =   1905
      TabIndex        =   55
      ToolTipText     =   "Tools"
      Top             =   1215
      Width           =   1935
   End
   Begin VB.PictureBox picMenu1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   120
      Picture         =   "frMain3.frx":8C1B
      ScaleHeight     =   765
      ScaleWidth      =   1905
      TabIndex        =   54
      ToolTipText     =   "Overview"
      Top             =   1215
      Width           =   1935
   End
   Begin VB.PictureBox picMenu2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   2040
      Picture         =   "frMain3.frx":9327
      ScaleHeight     =   765
      ScaleWidth      =   1905
      TabIndex        =   53
      ToolTipText     =   "Scan Area"
      Top             =   1215
      Width           =   1935
   End
   Begin VB.PictureBox picMenu3 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   3960
      Picture         =   "frMain3.frx":9A4E
      ScaleHeight     =   765
      ScaleWidth      =   1905
      TabIndex        =   52
      ToolTipText     =   "Settings"
      Top             =   1215
      Width           =   1935
   End
   Begin VB.PictureBox pqDsib 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   7800
      Picture         =   "frMain3.frx":A175
      ScaleHeight     =   765
      ScaleWidth      =   1905
      TabIndex        =   51
      Top             =   1215
      Width           =   1935
   End
   Begin VB.PictureBox pTDsib 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   5880
      Picture         =   "frMain3.frx":A9D5
      ScaleHeight     =   765
      ScaleWidth      =   1905
      TabIndex        =   50
      Top             =   1215
      Width           =   1935
   End
   Begin VB.PictureBox psDsib 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   3960
      Picture         =   "frMain3.frx":AF69
      ScaleHeight     =   765
      ScaleWidth      =   1905
      TabIndex        =   49
      Top             =   1215
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   10800
      Picture         =   "frMain3.frx":B6B8
      ScaleHeight     =   270
      ScaleMode       =   0  'User
      ScaleWidth      =   660
      TabIndex        =   41
      Top             =   0
      Width           =   666
   End
   Begin VB.PictureBox cCloseSys 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   10800
      Picture         =   "frMain3.frx":BB2F
      ScaleHeight     =   270
      ScaleWidth      =   660
      TabIndex        =   40
      Top             =   0
      Width           =   666
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   275
      Left            =   10120
      Picture         =   "frMain3.frx":BF67
      ScaleHeight     =   270
      ScaleWidth      =   660
      TabIndex        =   39
      Top             =   0
      Width           =   666
   End
   Begin VB.PictureBox cMinSys 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   275
      Left            =   10120
      Picture         =   "frMain3.frx":C338
      ScaleHeight     =   270
      ScaleWidth      =   660
      TabIndex        =   38
      ToolTipText     =   "Minimize"
      Top             =   0
      Width           =   666
   End
   Begin VB.PictureBox picIconDAV 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   500
      Left            =   240
      Picture         =   "frMain3.frx":C6F3
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   27
      Top             =   360
      Width           =   500
   End
   Begin VB.PictureBox Menu5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   600
      ScaleHeight     =   4905
      ScaleWidth      =   11280
      TabIndex        =   9
      Top             =   2880
      Width           =   11315
      Begin prjBeeAV.jcbutton jcbutton9 
         Height          =   270
         Left            =   7245
         TabIndex        =   65
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
         Left            =   9000
         TabIndex        =   64
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
         TabIndex        =   48
         Top             =   2880
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
      Begin VB.PictureBox picSec6 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5280
         Picture         =   "frMain3.frx":CBDA
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   45
         Top             =   2880
         Width           =   255
      End
      Begin VB.PictureBox picSec 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   360
         Picture         =   "frMain3.frx":CF64
         ScaleHeight     =   2415
         ScaleWidth      =   2535
         TabIndex        =   35
         Top             =   240
         Width           =   2535
      End
      Begin VB.PictureBox picSec4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5280
         Picture         =   "frMain3.frx":1DF82
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   11
         Top             =   2160
         Width           =   255
      End
      Begin VB.PictureBox picSec5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5280
         Picture         =   "frMain3.frx":1E30C
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   12
         Top             =   2520
         Width           =   255
      End
      Begin VB.PictureBox picSec1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5280
         Picture         =   "frMain3.frx":1E696
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   13
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox picSec2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5280
         Picture         =   "frMain3.frx":1EA20
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   14
         Top             =   1440
         Width           =   255
      End
      Begin VB.PictureBox picSec3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5280
         Picture         =   "frMain3.frx":1EDAA
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   15
         Top             =   1800
         Width           =   255
      End
      Begin VB.PictureBox Picture17 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5280
         Picture         =   "frMain3.frx":1F134
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   34
         Top             =   2160
         Width           =   255
      End
      Begin VB.PictureBox Picture16 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5280
         Picture         =   "frMain3.frx":1F476
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   33
         Top             =   2520
         Width           =   255
      End
      Begin VB.PictureBox Picture15 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5280
         Picture         =   "frMain3.frx":1F7B8
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   32
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox Picture14 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5280
         Picture         =   "frMain3.frx":1FAFA
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   31
         Top             =   1440
         Width           =   255
      End
      Begin VB.PictureBox Picture13 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5280
         Picture         =   "frMain3.frx":1FE3C
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   30
         Top             =   1800
         Width           =   255
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5280
         Picture         =   "frMain3.frx":2017E
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   46
         Top             =   2880
         Width           =   255
      End
      Begin VB.PictureBox picNotSec 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   360
         Picture         =   "frMain3.frx":204C0
         ScaleHeight     =   2415
         ScaleWidth      =   2535
         TabIndex        =   10
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label11 
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
         Left            =   5640
         TabIndex        =   47
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "+ Program Version"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   44
         Top             =   2880
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
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   5640
         TabIndex        =   29
         Top             =   2520
         Width           =   2175
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
         Left            =   5640
         TabIndex        =   28
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lbSec 
         BackStyle       =   0  'Transparent
         Caption         =   "Secured"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   615
         Left            =   3120
         TabIndex        =   26
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "+ Virus Definitions"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   23
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "+ Auto Scan FlashDisk"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   22
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "+ Scanning Path"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   21
         Top             =   1800
         Width           =   1695
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
         Left            =   5640
         TabIndex        =   20
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Label Label19 
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
         Left            =   5640
         TabIndex        =   19
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "+ Use Heuristic"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   18
         Top             =   2160
         Width           =   1935
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
         Left            =   5640
         TabIndex        =   17
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "+ Real - Time Protection"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   16
         Top             =   1080
         Width           =   2175
      End
   End
   Begin VB.Timer tmFD 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   2160
      Top             =   2520
   End
   Begin VB.PictureBox pic7 
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
      Height          =   240
      Left            =   1800
      Picture         =   "frMain3.frx":314DE
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic6 
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
      Height          =   240
      Left            =   1440
      Picture         =   "frMain3.frx":31A68
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picBuff 
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
      Height          =   255
      Left            =   2040
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic5 
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
      Height          =   240
      Left            =   1200
      Picture         =   "frMain3.frx":31FF2
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   2640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.ListBox drvList 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      ItemData        =   "frMain3.frx":32334
      Left            =   720
      List            =   "frMain3.frx":32336
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox pic4 
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
      Height          =   240
      Left            =   960
      Picture         =   "frMain3.frx":32338
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic3 
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
      Height          =   240
      Left            =   720
      Picture         =   "frMain3.frx":330B2
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   2640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic2 
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
      Height          =   240
      Left            =   480
      Picture         =   "frMain3.frx":333F4
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic1 
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
      Height          =   240
      Left            =   240
      Picture         =   "frMain3.frx":33736
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   8
      Top             =   3000
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer tmBuff 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2040
      Top             =   2040
   End
   Begin VB.Timer tmSpeed 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   2160
   End
   Begin VB.PictureBox pic8 
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
      Height          =   240
      Left            =   240
      Picture         =   "frMain3.frx":33A78
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   24
      Top             =   3240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000A&
      X1              =   10560
      X2              =   10560
      Y1              =   360
      Y2              =   600
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   10680
      TabIndex        =   43
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "About Team"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   9480
      TabIndex        =   42
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version : BETA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   10080
      TabIndex        =   37
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "2013"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3960
      TabIndex        =   36
      Top             =   360
      Width           =   1095
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   120
      X2              =   11400
      Y1              =   2025
      Y2              =   2025
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   120
      X2              =   11400
      Y1              =   1180
      Y2              =   1180
   End
   Begin prjBeeAV.UniDialog udScan 
      Left            =   1440
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      FileFlags       =   2621444
      FolderFlags     =   835
      FileCustomFilter=   "frMain3.frx":33DBA
      FileDefaultExtension=   "frMain3.frx":33DDA
      FileFilter      =   "frMain3.frx":33DFA
      FileOpenTitle   =   "frMain3.frx":33E42
      FileSaveTitle   =   "frMain3.frx":33E7A
      FolderMessage   =   "frMain3.frx":33EB2
   End
   Begin VB.Label Label48 
      BackStyle       =   0  'Transparent
      Caption         =   "Diyusof Antivirus"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   840
      TabIndex        =   25
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "frMain1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Oval
Private Declare Function CreateRoundRectRgn Lib "GDI32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "GDI32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'Akhir oval

'menggerakan form tanpa border
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
'akhirnya border
Public xToolIndexPic As Long
Public xAboutIndexPic As Long

Dim UserDown As Boolean
Dim UserStartX, UserStartY

Private Sub cBackAbout_Click()
If xAboutIndexPic - 1 < 0 Then Exit Sub
picAbout(xAboutIndexPic).Visible = False
picAbout(xAboutIndexPic - 1).Visible = True
xAboutIndexPic = xAboutIndexPic - 1
Select Case xAboutIndexPic
Case 0
cBackAbout.Enabled = False
cForwardAbout.Enabled = True
lAbout.Caption = "Diyusof Team"
Case 1
cBackAbout.Enabled = True
cForwardAbout.Enabled = True
lAbout.Caption = "Thank's to"
End Select
End Sub

Private Sub cBackM1_Click()
If xToolIndexPic - 1 < 0 Then Exit Sub
picTool(xToolIndexPic).Visible = False
picTool(xToolIndexPic - 1).Visible = True
xToolIndexPic = xToolIndexPic - 1
Select Case xToolIndexPic
Case 0
ENUM_PROSES lvProcess, picBuff
cBackM1.Enabled = False
cNextM1.Enabled = True
lTool.Caption = "Process Manager"
Case 1
lTool.Caption = "Diyusof - Lock"
lvDlock.ListItems.Clear
GetAllDrive
cBackM1.Enabled = True
cNextM1.Enabled = True
Case 2
lTool.Caption = "Tools External"
GetRegStartup lvStartup
cBackM1.Enabled = True
cNextM1.Enabled = True
End Select
End Sub

Private Sub cCancelRegT_Click()
GetRegTweakSetting
End Sub

Private Sub cCloseSys_Click()
Unload Me
End Sub

Private Sub cCloseSys_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.Visible = True
Picture20.Visible = True
Picture5.Visible = True
Picture6.Visible = True
If StopScan = True Then
Picture18.Visible = True
Picture19.Visible = True
Picture3.Visible = True
Else
Picture18.Visible = False
Picture19.Visible = False
Picture3.Visible = False
End If
End Sub

Private Sub ck_Click(Index As Integer)
Select Case Index
Case 3
If ck(3).Value = 1 Then
    picSec.Visible = True
    lbSec.Caption = "Secured"
    lbSec.ForeColor = &H8000&
    picSec4.Visible = True
    Label6.ForeColor = &H8000&
    Label6.Caption = "On"
Else
    ck(3).Value = 0
    picSec.Visible = False
     lbSec.Caption = "Not secured"
     lbSec.ForeColor = &HC0&
     picSec4.Visible = False
     Label6.ForeColor = &HC0&
     Label6.Caption = "Off"
End If


Case 6
If ck(6).Value = 1 Then
picSec.Visible = True
    lbSec.Caption = "Secured"
    lbSec.ForeColor = &H8000&
    picSec2.Visible = True
    Label19.ForeColor = &H8000&
    Label19.Caption = "On"
Else
    ck(6).Value = 0
    picSec.Visible = False
     lbSec.Caption = "Not secured"
     lbSec.ForeColor = &HC0&
     picSec2.Visible = False
     Label19.ForeColor = &HC0&
     Label19.Caption = "Off"
End If

Case 7
If ck(7).Value = 1 Then
picSec.Visible = True
    jcbutton9.Visible = False
    lbSec.Caption = "Secured"
    lbSec.ForeColor = &H8000&
    picSec1.Visible = True
    Label15.ForeColor = &H8000&
    Label15.Caption = "On"
    frRTP.mnEP.Checked = True
    UpdateIcon Me.Icon, "Diyusof Real Time - Protection is ON", frRTP
    TampilkanBalon frRTP, "Your PC is Protect", "Protection Active", NIIF_INFO

Else
   ck(7).Value = 0
   picSec.Visible = False
   jcbutton9.Visible = True
    lbSec.Caption = "Not secured"
    lbSec.ForeColor = &HC0&
    picSec1.Visible = False
    Label15.ForeColor = &HC0&
    Label15.Caption = "Off"
    frRTP.mnEP.Checked = False
    UpdateIcon Me.Icon, "Diyusof Real Time - Protection is OFF", frRTP
    TampilkanBalon frRTP, "Your PC is Not Protect", "Protection Not Active", NIIF_ERROR
    jcbutton9.Visible = True
    End If
End Select
End Sub

Private Sub cForwardAbout_Click()
If xAboutIndexPic + 1 > 2 Then Exit Sub
picAbout(xAboutIndexPic).Visible = False
picAbout(xAboutIndexPic + 1).Visible = True
xAboutIndexPic = xAboutIndexPic + 1
Select Case xAboutIndexPic
Case 1
cBackAbout.Enabled = True
cForwardAbout.Enabled = True
lAbout.Caption = "Thank's to"
Case 2
cBackAbout.Enabled = True
cForwardAbout.Enabled = False
lAbout.Caption = "Diyusof Antivirus Information"
End Select
End Sub


Private Sub cLock_Click()
Static i As Integer

For i = 1 To lvDlock.ListItems.count
    If lvDlock.ListItems.Item(i).Checked = True Then
    BuatProt lvDlock.ListItems.Item(i).Text
    End If
Next
MsgBox "Finish !", vbInformation
GetAllDrive
End Sub

Private Sub cmdBack_Click()
picScan(0).Visible = True
picScan(1).Visible = False
lScan.Caption = "Scan Area"
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
If cmdPause.Caption = "Pause" Then
cmdStop.Enabled = False
isPause = True
cmdPause.Caption = "Resume"
Else
isPause = False
cmdStop.Enabled = True
cmdPause.Caption = "Pause"
End If
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

Private Sub cmdStop_Click()
StopScan = True
StopKumpulkan
    Scan(0).Visible = True
    Scan(1).Visible = True
    Scan(2).Visible = True
picMenu3.Visible = True
picMenu4.Visible = True
picMenu6.Visible = True
Picture18.Visible = True
Picture19.Visible = True
Picture3.Visible = True
cmdSkip.Enabled = False
cmdPause.Enabled = True
cmdPause.Caption = "Pause"
End Sub

Private Sub cMinSys_Click()
Me.WindowState = 1
Shell_NotifyIcon NIM_DELETE, nID
UpdateIcon frRTP.Icon, "Diyusof RT-Protection, Diyusof 2012 Ver. 2.1", frRTP
End Sub

Private Sub cMinSys_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture2.Visible = True
Picture20.Visible = True
Picture5.Visible = True
Picture6.Visible = True
If StopScan = True Then
Picture18.Visible = True
Picture19.Visible = True
Picture3.Visible = True
Else
Picture18.Visible = False
Picture19.Visible = False
Picture3.Visible = False
End If
End Sub

Private Sub cNextM1_Click()
If xToolIndexPic + 1 > 3 Then Exit Sub
picTool(xToolIndexPic).Visible = False
picTool(xToolIndexPic + 1).Visible = True
xToolIndexPic = xToolIndexPic + 1
Select Case xToolIndexPic
Case 1
lTool.Caption = "Diyusof - Lock"
lvDlock.ListItems.Clear
GetAllDrive
cBackM1.Enabled = True
cNextM1.Enabled = True
Case 2
lTool.Caption = "Tools External"
GetRegStartup lvStartup
cBackM1.Enabled = True
cNextM1.Enabled = True
Case 3
lTool.Caption = "Diyusof Registry Tweaker"
GetRegTweakSetting
cBackM1.Enabled = True
cNextM1.Enabled = False
End Select
End Sub

Private Sub cSave_Click()
    Dim X
    On Error Resume Next
    ' Update according to each control
    For Each X In Controls
        If X.Tag <> "" Then _
            SetDwordValue SingkatanKey("HKCU"), X.Tag, Replace(X.name, " ", " "), X.Value
    Next
    
    Call SetStringValue(SingkatanKey("HKLM"), "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner", txtOwner.Text)
    Call SetStringValue(SingkatanKey("HKLM"), "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOrganization", txtOrg.Text)
    
    MsgBox "Done !", vbInformation
    If MsgBox("Do you want to Restart Explorer to take the effect ?", vbQuestion + vbYesNo) = vbYes Then
        KillByProccess "explorer.exe"
        ShellExecute Me.hWnd, vbNullString, GetSpecFolder(WINDOWS_DIR) & "\explorer.exe", vbNullString, "C:\", 1
    End If
End Sub

Private Sub cULock_Click()
Static i As Integer

For i = 1 To lvDlock.ListItems.count
    If lvDlock.ListItems.Item(i).Checked = True Then
    UnProt lvDlock.ListItems.Item(i).Text
    End If
Next
MsgBox "Finish !", vbInformation
GetAllDrive
End Sub

Private Sub Form_Load()
Me.Icon = frIcon.Icon
BuildLV
GetSettingF App.Path & "\Setting.ini"
ApplySetting
InitPHPattern
LoadDataIcon
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then 'left click
ReleaseCapture
SendMessage Me.hWnd, &HA1, 2, 0
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label34.Font.Underline = False
Label9.Font.Underline = False
Picture1.Visible = True
Picture2.Visible = True
Picture20.Visible = True
Picture5.Visible = True
Picture6.Visible = True
If StopScan = True Then
Picture18.Visible = True
Picture19.Visible = True
Picture3.Visible = True
Else
Picture18.Visible = False
Picture19.Visible = False
Picture3.Visible = False
End If
End Sub

Public Sub Form_Unload(Cancel As Integer)
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

Private Sub jcbutton1_Click()
    Menu1.Visible = False
    Menu2.Visible = False
    Menu3.Visible = False
    Menu4.Visible = False
    Menu0.Visible = True
    Menu5.Visible = False
    picScan(0).Visible = True
    picScan(1).Visible = False
    cmdBack.Enabled = False
    cmdForward.Enabled = True
    lScan.Caption = "Scan Area"
End Sub


Private Sub jcbutton10_Click()
If Dir(App.Path & "\Tools\Virtual Keyboard.exe") = "" Then
MsgBox "Sorry, File Virtual Keyboard.exe Not Found.", 0 + vbExclamation, "Error"
Else
ShellExecute Me.hWnd, vbNullString, App.Path & "\Tools\Virtual Keyboard.exe", vbNullString, "C:\", 1
End If
End Sub

Private Sub jcbutton2_Click()
FixVir FixChk, frMain.lvMal, del
End Sub

Private Sub jcbutton3_Click()
FixVir FixChk, frMain.lvMal, Quar
End Sub

Private Sub jcbutton4_Click()
PopupMenu frIcon.mnLvMal
End Sub


Private Sub jcbutton5_Click()
If Dir(App.Path & "\Tools\Site Blocking.exe") = "" Then
MsgBox "Sorry, File Site Blocking.exe Not Found.", 0 + vbExclamation, "Error"
Else
ShellExecute Me.hWnd, vbNullString, App.Path & "\Tools\Site Blocking.exe", vbNullString, "C:\", 1
End If
End Sub

Private Sub jcbutton6_Click()
If Dir(App.Path & "\Tools\Fixed Registry.exe") = "" Then
MsgBox "Sorry, File Fixed Registry.exe Not Found.", 0 + vbExclamation, "Error"
Else
ShellExecute Me.hWnd, vbNullString, App.Path & "\Tools\Fixed Registry.exe", vbNullString, "C:\", 1
End If
End Sub

Private Sub jcbutton7_Click()
If Dir(App.Path & "\Tools\Diyusof - Set Attributes.exe") = "" Then
MsgBox "Sorry, File Diyusof - Set Attributes Not Found.", 0 + vbExclamation, "Error"
Else
ShellExecute Me.hWnd, vbNullString, App.Path & "\Tools\Diyusof - Set Attributes.exe", vbNullString, "C:\", 1
End If

End Sub

Private Sub jcbutton9_Click()
    picSec.Visible = True
    ck(7).Value = 1
    lbSec.Caption = "Secured"
    lbSec.ForeColor = &H8000&
    picSec1.Visible = True
    Label15.ForeColor = &H8000&
    Label15.Caption = "On"
    frRTP.mnEP.Checked = True
    UpdateIcon Me.Icon, "Diyusof Real Time - Protection is ON", frRTP
    TampilkanBalon frRTP, "Your PC is Protect", "Protection Active", NIIF_INFO
    jcbutton9.Visible = False
End Sub

Private Sub Label34_Click()
    Menu0.Visible = False
    Menu1.Visible = False
    Menu2.Visible = False
    Menu3.Visible = False
    Menu5.Visible = False
    Menu4.Visible = True
    xAboutIndexPic = 1
    Call cBackAbout_Click
End Sub

Private Sub Label34_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label34.Font.Underline = True
End Sub



Private Sub Label48_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then 'left click
ReleaseCapture
SendMessage Me.hWnd, &HA1, 2, 0
End If
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then 'left click
ReleaseCapture
SendMessage Me.hWnd, &HA1, 2, 0
End If
End Sub

Private Sub Label9_Click()
    Menu0.Visible = False
    Menu1.Visible = False
    Menu3.Visible = False
    Menu4.Visible = False
    Menu5.Visible = False
    Menu2.Visible = True
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.Font.Underline = True
End Sub

Private Sub lvDlock_ColumnClick(ByVal oColumn As cColumn)
oColumn.Sort
End Sub

Private Sub lvMal_ColumnClick(ByVal oColumn As cColumn)
oColumn.Sort
End Sub

Private Sub lvMal_ContextMenu(ByVal X As Single, ByVal Y As Single)
PopupMenu frIcon.mnLvMal
End Sub

Private Sub lvProcess_ColumnClick(ByVal oColumn As cColumn)
oColumn.Sort
End Sub

Private Sub lvProcess_ContextMenu(ByVal X As Single, ByVal Y As Single)
PopupMenu frIcon.mnlvProcess
End Sub

Private Sub lvQuar_ColumnClick(ByVal oColumn As cColumn)
oColumn.Sort
End Sub

Private Sub LVQuar_ContextMenu(ByVal X As Single, ByVal Y As Single)
PopupMenu frIcon.mnlvQuar
End Sub

Private Sub lvStartup_ColumnClick(ByVal oColumn As cColumn)
oColumn.Sort
End Sub

Private Sub lvStartup_ContextMenu(ByVal X As Single, ByVal Y As Single)
Static Inter As Integer
Dim pCada   As String
Dim tStatus  As String

Inter = 1
For Inter = 1 To lvStartup.ListItems.count
    If lvStartup.ListItems.Item(Inter).Selected = True Then
        tStatus = lvStartup.ListItems.Item(Inter).SubItem(6).Text
        If tStatus = "Enable" Then
        frIcon.mnES.Enabled = False
        frIcon.mnDS.Enabled = True
        PopupMenu frIcon.mnST
        Else
        frIcon.mnES.Enabled = True
        frIcon.mnDS.Enabled = False
        PopupMenu frIcon.mnST
        End If
    End If
Next
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
    Menu5.Visible = False
    picScan(0).Visible = True
    picScan(1).Visible = False
    cmdBack.Enabled = False
    cmdForward.Enabled = True
    lScan.Caption = "Scan Area"
Case 1
    Menu0.Visible = False
    Menu2.Visible = False
    Menu3.Visible = False
    Menu4.Visible = False
    Menu5.Visible = False
    Menu1.Visible = True
    xToolIndexPic = 0
    ENUM_PROSES lvProcess, picBuff
    picTool(0).Visible = True
    GetAllDrive
    picTool(1).Visible = False
    picTool(2).Visible = False
    cBackM1.Enabled = False
    cNextM1.Enabled = True
    lTool.Caption = "Process Manager"
Case 2
    Menu0.Visible = False
    Menu1.Visible = False
    Menu3.Visible = False
    Menu4.Visible = False
    Menu5.Visible = False
    Menu2.Visible = True
Case 3
    Menu0.Visible = False
    Menu1.Visible = False
    Menu2.Visible = False
    Menu4.Visible = False
    Menu5.Visible = False
    Menu3.Visible = True
    lvQuar.ListItems.Clear
    GetQuarFile
Case 4
    Menu0.Visible = False
    Menu1.Visible = False
    Menu2.Visible = False
    Menu3.Visible = False
    Menu5.Visible = False
    Menu4.Visible = True
    xAboutIndexPic = 1
    Call cBackAbout_Click
Case 5
    Menu0.Visible = False
    Menu1.Visible = False
    Menu2.Visible = False
    Menu3.Visible = False
    Menu5.Visible = True
    Menu4.Visible = False
End Select
End Sub




Private Sub picBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then 'left click
ReleaseCapture
SendMessage Me.hWnd, &HA1, 2, 0
End If
End Sub


Private Sub Menu0_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture20.Visible = True
Picture5.Visible = True
Picture6.Visible = True
If StopScan = True Then
Picture18.Visible = True
Picture19.Visible = True
Picture3.Visible = True
Else
Picture18.Visible = False
Picture19.Visible = False
Picture3.Visible = False
End If
End Sub

Private Sub Menu1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture20.Visible = True
Picture5.Visible = True
Picture6.Visible = True
Picture18.Visible = True
Picture19.Visible = True
Picture3.Visible = True

End Sub

Private Sub Menu2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture20.Visible = True
Picture5.Visible = True
Picture6.Visible = True
If StopScan = True Then
Picture18.Visible = True
Picture19.Visible = True
Picture3.Visible = True
Else
Picture18.Visible = False
Picture19.Visible = False
Picture3.Visible = False
End If
End Sub

Private Sub Menu3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture20.Visible = True
Picture5.Visible = True
Picture6.Visible = True
If StopScan = True Then
Picture18.Visible = True
Picture19.Visible = True
Picture3.Visible = True
Else
Picture18.Visible = False
Picture19.Visible = False
Picture3.Visible = False
End If
End Sub

Private Sub Menu4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture20.Visible = True
Picture5.Visible = True
Picture6.Visible = True
If StopScan = True Then
Picture18.Visible = True
Picture19.Visible = True
Picture3.Visible = True
Else
Picture18.Visible = False
Picture19.Visible = False
Picture3.Visible = False
End If
End Sub

Private Sub Menu5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture20.Visible = True
Picture5.Visible = True
Picture6.Visible = True
If StopScan = True Then
Picture18.Visible = True
Picture19.Visible = True
Picture3.Visible = True
Else
Picture18.Visible = False
Picture19.Visible = False
Picture3.Visible = False
End If
End Sub




Private Sub picAbout_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture20.Visible = True
Picture5.Visible = True
Picture6.Visible = True
If StopScan = True Then
Picture18.Visible = True
Picture19.Visible = True
Picture3.Visible = True
Else
Picture18.Visible = False
Picture19.Visible = False
Picture3.Visible = False
End If
End Sub


Private Sub picIconDAV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then 'left click
ReleaseCapture
SendMessage Me.hWnd, &HA1, 2, 0
End If
End Sub

Private Sub picMenu1_Click()
    Menu0.Visible = False
    Menu1.Visible = False
    Menu2.Visible = False
    Menu3.Visible = False
    Menu5.Visible = True
    Menu4.Visible = False
End Sub

Private Sub picMenu2_Click()
    Menu1.Visible = False
    Menu2.Visible = False
    Menu3.Visible = False
    Menu4.Visible = False
    Menu0.Visible = True
    Menu5.Visible = False
    picScan(0).Visible = True
    picScan(1).Visible = False
    cmdBack.Enabled = False
    cmdForward.Enabled = True
    lScan.Caption = "Scan Area"
End Sub

Private Sub picMenu3_Click()
    Menu0.Visible = False
    Menu1.Visible = False
    Menu3.Visible = False
    Menu4.Visible = False
    Menu5.Visible = False
    Menu2.Visible = True
End Sub

Private Sub picMenu4_Click()
    Menu0.Visible = False
    Menu2.Visible = False
    Menu3.Visible = False
    Menu4.Visible = False
    Menu5.Visible = False
    Menu1.Visible = True
    xToolIndexPic = 0
    ENUM_PROSES lvProcess, picBuff
    picTool(0).Visible = True
    GetAllDrive
    picTool(1).Visible = False
    picTool(2).Visible = False
    cBackM1.Enabled = False
    cNextM1.Enabled = True
    lTool.Caption = "Process Manager"
End Sub

Private Sub picMenu6_Click()
    Menu0.Visible = False
    Menu1.Visible = False
    Menu2.Visible = False
    Menu4.Visible = False
    Menu5.Visible = False
    Menu3.Visible = True
    lvQuar.ListItems.Clear
    GetQuarFile
End Sub



Private Sub picMenu7_Click()
    Menu0.Visible = False
    Menu1.Visible = False
    Menu2.Visible = False
    Menu3.Visible = False
    Menu5.Visible = False
    Menu4.Visible = True
    xAboutIndexPic = 1
    Call cBackAbout_Click
End Sub



Private Sub picScan_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture20.Visible = True
Picture5.Visible = True
Picture6.Visible = True
If StopScan = True Then
Picture18.Visible = True
Picture19.Visible = True
Picture3.Visible = True
Else
Picture18.Visible = False
Picture19.Visible = False
Picture3.Visible = False
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.Visible = False
End Sub



Private Sub Picture18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture18.Visible = False
Picture5.Visible = True
Picture6.Visible = True
Picture3.Visible = True
Picture19.Visible = True
Picture20.Visible = True
End Sub

Private Sub Picture19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture5.Visible = True
Picture6.Visible = True
Picture18.Visible = True
Picture3.Visible = True
Picture20.Visible = True
Picture19.Visible = False

End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture2.Visible = False
End Sub


Private Sub Picture20_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture20.Visible = False
Picture5.Visible = True
Picture6.Visible = True
If StopScan = True Then
Picture18.Visible = True
Picture19.Visible = True
Picture3.Visible = True
Else
Picture18.Visible = False
Picture19.Visible = False
Picture3.Visible = False
End If

End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3.Visible = False
Picture5.Visible = True
Picture6.Visible = True
Picture18.Visible = True
Picture19.Visible = True
Picture20.Visible = True
End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture5.Visible = False
Picture6.Visible = True
If StopScan = True Then
Picture18.Visible = True
Picture19.Visible = True
Picture3.Visible = True
Else
Picture18.Visible = False
Picture19.Visible = False
Picture3.Visible = False
End If
Picture20.Visible = True
End Sub

Private Sub Picture6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture6.Visible = False
Picture5.Visible = True
Picture20.Visible = True
If StopScan = True Then
Picture18.Visible = True
Picture19.Visible = True
Picture3.Visible = True
Else
Picture18.Visible = False
Picture19.Visible = False
Picture3.Visible = False
End If

End Sub

Public Sub Scan_Click(Index As Integer)
Static tR As Integer
AllReset
lvMal.ListItems.Clear
pScan.Value = 0
Picture3.Visible = False
Picture18.Visible = False
Picture19.Visible = False
picMenu3.Visible = False
picMenu4.Visible = False
picMenu6.Visible = False
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
Label20.Caption = "Full system Scan..."
Label4.Caption = "Scan Running...."
Label3.Caption = "Please Wait...."
Label3.ForeColor = &HC0&
Scan(0).Enabled = False
Scan(1).Enabled = False
Scan(2).Enabled = False
Scan(3).Enabled = False
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
    Scan(3).Enabled = True
    cmdSkip.Enabled = True
    FinishJob
Case 1
GetQuickPath
    WithBuffer = True
    tmBuff.Enabled = True
    Label20.Caption = "Quick Scan..."
    Label4.Caption = "Scan Running...."
    Label3.Caption = "Please Wait...."
    Label3.ForeColor = &HC0&
    Scan(0).Enabled = False
    Scan(1).Enabled = False
    Scan(2).Enabled = False
    Scan(3).Enabled = False
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
    Scan(3).Enabled = True
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
Scan(3).Enabled = False
Label20.Caption = "Scanning Path..."
Label4.Caption = "Scan Running...."
Label3.Caption = "Please Wait...."
    cmdPause.Enabled = False
    Label3.ForeColor = &HC0&
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
    Scan(3).Enabled = True
    cmdSkip.Enabled = True
    FinishJob
    GoTo l_Akhir
l_CustomScan:
Scan(0).Enabled = False
Scan(1).Enabled = False
Scan(2).Enabled = False
Scan(3).Enabled = False
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
    Scan(3).Enabled = True
    cmdSkip.Enabled = True
    FinishJob
Case 3
    WithBuffer = True
    Label20.Caption = "Scanning Proces..."
    Label4.Caption = "Scan Running...."
    Label3.Caption = "Please Wait...."
    Label3.ForeColor = &HC0&
    Scan(0).Enabled = False
    Scan(1).Enabled = False
    Scan(2).Enabled = False
    Scan(3).Enabled = False
    cmdPause.Enabled = False
    cmdPause.Enabled = True
    cmdSkip.Enabled = False
    tmSpeed.Enabled = True
    FileRemain = FileToScan
    FileToScan = 0
    ScanProses True, lblFile
    Call tmSpeed_Timer
    tmSpeed.Enabled = False
    Scan(0).Enabled = True
    Scan(1).Enabled = True
    Scan(2).Enabled = True
    Scan(3).Enabled = True
    cmdSkip.Enabled = True
    FinishJob
End Select
l_Akhir:
Label20.Caption = "Waiting for Instructions ...."
picMenu3.Visible = True
picMenu4.Visible = True
picMenu6.Visible = True
Picture3.Visible = True
Picture18.Visible = True
Picture19.Visible = True
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
   frMain.Scan_Click (2)
   frMain.Menu_Click (0)
   frMain.Show
   frMain.WindowState = 0
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
lSpeed.Caption = ": " & FileToScan & " File[s]/s"
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
Static pokLet As Long
Static pokGet As Long
pokLet = 0
pokGet = UBound(PathContextScan()) - 1
For pokLet = 0 To pokGet
If ValidFile(PathContextScan(pokLet)) = True Then
Scan(0).Enabled = False
Scan(1).Enabled = False
Scan(2).Enabled = False
Scan(3).Enabled = False
    cmdPause.Enabled = False
    WithBuffer = True
    tmBuff.Enabled = True
    FileToScan = pokGet + 1
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
                    If isProperFile(PathContextScan(pokLet), "SYS LNK VBE HTM HTT EXE DLL VBS VMX TML .DB COM SCR BAT INF TML CMD TXT PIF MSI BMP") = True Then
                        Equal (PathContextScan(pokLet))
                    Else
                        FileIgnore = FileIgnore + 1
                    End If
                Else
                    Equal (PathContextScan(pokLet))
                End If
    FileRemain = 0
    Call tmSpeed_Timer
    pScan.Value = 0
    frMain.lblFile.Caption = GetLongFileName(PathContextScan(pokLet))
    tmSpeed.Enabled = False
    Scan(0).Enabled = True
    Scan(1).Enabled = True
    Scan(2).Enabled = True
    Scan(3).Enabled = True
    cmdSkip.Enabled = True
ElseIf ValidFolder(PathContextScan(pokLet)) = True Then
Scan(0).Enabled = False
Scan(1).Enabled = False
Scan(2).Enabled = False
Scan(3).Enabled = False
    cmdPause.Enabled = False
    WithBuffer = True
    tmBuff.Enabled = True
    BufferPath PathContextScan(pokLet), True
    cmdSkip.Enabled = False
    cmdPause.Enabled = True
    tmBuff.Enabled = False
    tmSpeed.Enabled = True
    FileRemain = FileToScan
    pScan.Max = FileToScan
    FileToScan = 0
    KumpulkanFile PathContextScan(pokLet), True, True
    FileRemain = 0
    Call tmSpeed_Timer
    pScan.Value = 0
    tmSpeed.Enabled = False
    Scan(0).Enabled = True
    Scan(1).Enabled = True
    Scan(2).Enabled = True
    Scan(3).Enabled = True
    cmdSkip.Enabled = True
Else
MsgBox "Can't Scan File or Folder !" & vbCrLf & _
"Invalid or Protected !", vbCritical, "Error !"
End
End If
Next
FinishJob
End Sub

Public Sub GetRegTweakSetting()
    Dim X, X1
    On Error Resume Next
    For Each X In Controls
        If X.Tag <> "" Then
            X1 = GetDwordValue(SingkatanKey("HKCU"), X.Tag, Replace(X.name, " ", " "))
        If X1 = vbNullString Then X.Value = 0 Else X.Value = Int(X1)
        End If
    Next
    
    txtOwner.Text = GetStringValue(SingkatanKey("HKLM"), "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner")
    txtOrg.Text = GetStringValue(SingkatanKey("HKLM"), "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOrganization")

End Sub

