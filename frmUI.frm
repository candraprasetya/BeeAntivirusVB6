VERSION 5.00
Begin VB.Form frMain 
   Appearance      =   0  'Flat
   BackColor       =   &H0000CCFF&
   BorderStyle     =   0  'None
   Caption         =   "Bee Antivirus 2014"
   ClientHeight    =   6975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11775
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
   Icon            =   "frmUI.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmUI.frx":19F7A
   ScaleHeight     =   465
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   785
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   275
      Left            =   10320
      Picture         =   "frmUI.frx":125732
      ScaleHeight     =   270
      ScaleWidth      =   660
      TabIndex        =   230
      Top             =   60
      Width           =   666
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   11040
      Picture         =   "frmUI.frx":1260BE
      ScaleHeight     =   270
      ScaleMode       =   0  'User
      ScaleWidth      =   660
      TabIndex        =   229
      Top             =   60
      Width           =   666
   End
   Begin VB.PictureBox cMinSys 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   275
      Left            =   10320
      MouseIcon       =   "frmUI.frx":126A4A
      MousePointer    =   99  'Custom
      Picture         =   "frmUI.frx":126B9C
      ScaleHeight     =   270
      ScaleWidth      =   660
      TabIndex        =   228
      ToolTipText     =   "Minimize"
      Top             =   60
      Width           =   666
   End
   Begin VB.PictureBox cCloseSys 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   11040
      MouseIcon       =   "frmUI.frx":127528
      MousePointer    =   99  'Custom
      Picture         =   "frmUI.frx":12767A
      ScaleHeight     =   270
      ScaleWidth      =   660
      TabIndex        =   227
      Top             =   60
      Width           =   666
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      MousePointer    =   99  'Custom
      Picture         =   "frmUI.frx":128006
      ScaleHeight     =   735
      ScaleWidth      =   2175
      TabIndex        =   226
      ToolTipText     =   "Settings"
      Top             =   3240
      Width           =   2175
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      MousePointer    =   99  'Custom
      Picture         =   "frmUI.frx":12D3BE
      ScaleHeight     =   735
      ScaleWidth      =   2175
      TabIndex        =   225
      ToolTipText     =   "Scan Area"
      Top             =   2400
      Width           =   2175
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      MousePointer    =   99  'Custom
      Picture         =   "frmUI.frx":132776
      ScaleHeight     =   735
      ScaleWidth      =   2175
      TabIndex        =   224
      ToolTipText     =   "Overview"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.PictureBox Picture18 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      MousePointer    =   99  'Custom
      Picture         =   "frmUI.frx":137B2E
      ScaleHeight     =   735
      ScaleWidth      =   2175
      TabIndex        =   223
      ToolTipText     =   "Tools"
      Top             =   4080
      Width           =   2175
   End
   Begin VB.PictureBox Picture19 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      MousePointer    =   99  'Custom
      Picture         =   "frmUI.frx":13CEE6
      ScaleHeight     =   735
      ScaleWidth      =   2175
      TabIndex        =   222
      ToolTipText     =   "Quarantine"
      Top             =   4920
      Width           =   2175
   End
   Begin VB.PictureBox Picture20 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      MousePointer    =   99  'Custom
      Picture         =   "frmUI.frx":14229E
      ScaleHeight     =   735
      ScaleWidth      =   2175
      TabIndex        =   221
      ToolTipText     =   "About"
      Top             =   5760
      Width           =   2175
   End
   Begin VB.PictureBox picMenu4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      MouseIcon       =   "frmUI.frx":147656
      MousePointer    =   99  'Custom
      Picture         =   "frmUI.frx":1477A8
      ScaleHeight     =   735
      ScaleWidth      =   2295
      TabIndex        =   220
      ToolTipText     =   "Tools"
      Top             =   4080
      Width           =   2295
   End
   Begin VB.PictureBox picMenu6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      MouseIcon       =   "frmUI.frx":14CFF8
      MousePointer    =   99  'Custom
      Picture         =   "frmUI.frx":14D14A
      ScaleHeight     =   735
      ScaleWidth      =   2295
      TabIndex        =   219
      ToolTipText     =   "Quarantine"
      Top             =   4920
      Width           =   2295
   End
   Begin VB.PictureBox picMenu7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      MouseIcon       =   "frmUI.frx":15299A
      MousePointer    =   99  'Custom
      Picture         =   "frmUI.frx":152AEC
      ScaleHeight     =   735
      ScaleWidth      =   2295
      TabIndex        =   218
      ToolTipText     =   "About"
      Top             =   5760
      Width           =   2295
   End
   Begin VB.PictureBox picMenu3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      MouseIcon       =   "frmUI.frx":15833C
      MousePointer    =   99  'Custom
      Picture         =   "frmUI.frx":15848E
      ScaleHeight     =   735
      ScaleWidth      =   2295
      TabIndex        =   217
      ToolTipText     =   "Settings"
      Top             =   3240
      Width           =   2295
   End
   Begin VB.PictureBox picMenu2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      MouseIcon       =   "frmUI.frx":15DCDE
      MousePointer    =   99  'Custom
      Picture         =   "frmUI.frx":15DE30
      ScaleHeight     =   735
      ScaleWidth      =   2295
      TabIndex        =   216
      ToolTipText     =   "Scan Area"
      Top             =   2400
      Width           =   2295
   End
   Begin VB.PictureBox picMenu1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      MouseIcon       =   "frmUI.frx":163680
      MousePointer    =   99  'Custom
      Picture         =   "frmUI.frx":1637D2
      ScaleHeight     =   735
      ScaleWidth      =   2295
      TabIndex        =   215
      ToolTipText     =   "Overview"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.PictureBox psDsib 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      Picture         =   "frmUI.frx":169022
      ScaleHeight     =   735
      ScaleWidth      =   2295
      TabIndex        =   214
      Top             =   3240
      Width           =   2295
   End
   Begin VB.PictureBox pTDsib 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      Picture         =   "frmUI.frx":16E872
      ScaleHeight     =   735
      ScaleWidth      =   2295
      TabIndex        =   213
      Top             =   4080
      Width           =   2295
   End
   Begin VB.PictureBox pqDsib 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      Picture         =   "frmUI.frx":1740C2
      ScaleHeight     =   735
      ScaleWidth      =   2295
      TabIndex        =   212
      Top             =   4920
      Width           =   2295
   End
   Begin VB.PictureBox pic7 
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   3000
      Picture         =   "frmUI.frx":179912
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   210
      Top             =   7560
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic6 
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   2760
      Picture         =   "frmUI.frx":179E9C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   209
      Top             =   7560
      Visible         =   0   'False
      Width           =   240
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
      Left            =   1560
      Picture         =   "frmUI.frx":17A426
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   208
      Top             =   8160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer tmSpeed 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   7080
   End
   Begin VB.Timer tmBuff 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3360
      Top             =   6960
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
      Left            =   1560
      Picture         =   "frmUI.frx":17A768
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   207
      Top             =   7920
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
      Left            =   1800
      Picture         =   "frmUI.frx":17AAAA
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   206
      Top             =   7560
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
      Left            =   2040
      Picture         =   "frmUI.frx":17ADEC
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   205
      Top             =   7560
      Visible         =   0   'False
      Width           =   240
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
      Left            =   2280
      Picture         =   "frmUI.frx":17B12E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   204
      Top             =   7560
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
      ItemData        =   "frmUI.frx":17B470
      Left            =   2040
      List            =   "frmUI.frx":17B472
      TabIndex        =   203
      Top             =   7080
      Visible         =   0   'False
      Width           =   735
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
      Left            =   2520
      Picture         =   "frmUI.frx":17B474
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   202
      Top             =   7560
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
      Left            =   3360
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   201
      Top             =   7560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer tmFD 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   3480
      Top             =   7440
   End
   Begin VB.PictureBox picSec5 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4200
      Picture         =   "frmUI.frx":17B7B6
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   196
      Top             =   7440
      Width           =   255
   End
   Begin VB.PictureBox Picture16 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4440
      Picture         =   "frmUI.frx":17BAF8
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   195
      Top             =   7440
      Width           =   255
   End
   Begin VB.PictureBox Menu5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   2400
      ScaleHeight     =   4935
      ScaleWidth      =   9255
      TabIndex        =   171
      Top             =   1560
      Width           =   9255
      Begin prjBeeAV.AkuSayangIbu jcbutton1 
         Height          =   615
         Left            =   6720
         TabIndex        =   197
         Top             =   4080
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         BtnStyle        =   2
         Caption         =   "Scan Now >>"
      End
      Begin VB.PictureBox picSec 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   240
         Picture         =   "frmUI.frx":17BE3A
         ScaleHeight     =   2415
         ScaleWidth      =   2535
         TabIndex        =   192
         Top             =   960
         Width           =   2535
      End
      Begin VB.PictureBox picSec3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5640
         Picture         =   "frmUI.frx":18CE5A
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   191
         Top             =   3000
         Width           =   255
      End
      Begin VB.PictureBox picSec2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5640
         Picture         =   "frmUI.frx":18D19C
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   190
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox picSec1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5640
         Picture         =   "frmUI.frx":18D4DE
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   189
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox picSec4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5640
         Picture         =   "frmUI.frx":18D820
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   188
         Top             =   2040
         Width           =   255
      End
      Begin VB.PictureBox picSec6 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5640
         Picture         =   "frmUI.frx":18DB62
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   187
         Top             =   2520
         Width           =   255
      End
      Begin VB.PictureBox Picture3s 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2190
         Left            =   2880
         Picture         =   "frmUI.frx":18DEA4
         ScaleHeight     =   2190
         ScaleWidth      =   2340
         TabIndex        =   186
         Top             =   1080
         Width           =   2340
      End
      Begin VB.PictureBox picNotSec 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   240
         Picture         =   "frmUI.frx":19E9D0
         ScaleHeight     =   2415
         ScaleWidth      =   2535
         TabIndex        =   179
         Top             =   960
         Width           =   2535
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5640
         Picture         =   "frmUI.frx":1AF9F0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   178
         Top             =   2520
         Width           =   255
      End
      Begin VB.PictureBox Picture13 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5640
         Picture         =   "frmUI.frx":1AFD32
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   177
         Top             =   3000
         Width           =   255
      End
      Begin VB.PictureBox Picture14 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5640
         Picture         =   "frmUI.frx":1B0074
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   176
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox Picture15 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5640
         Picture         =   "frmUI.frx":1B03B6
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   175
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox Picture17 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5640
         Picture         =   "frmUI.frx":1B06F8
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   174
         Top             =   2040
         Width           =   255
      End
      Begin prjBeeAV.jcbutton jcbutton9 
         Height          =   270
         Left            =   7245
         TabIndex        =   172
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
         ForeColor       =   10404431
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin prjBeeAV.jcbutton jcbutton8 
         Height          =   270
         Left            =   7320
         TabIndex        =   173
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
         ForeColor       =   5090297
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
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   6120
         TabIndex        =   185
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
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   6120
         TabIndex        =   184
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
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   6120
         TabIndex        =   183
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
         TabIndex        =   182
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
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   6120
         TabIndex        =   181
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "1.3.0   Beta"
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
         Left            =   6120
         TabIndex        =   180
         Top             =   2520
         Width           =   975
      End
   End
   Begin VB.PictureBox Menu0 
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
      Height          =   4935
      Left            =   2400
      ScaleHeight     =   4935
      ScaleWidth      =   9255
      TabIndex        =   120
      Top             =   1560
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
         TabIndex        =   166
         Top             =   600
         Width           =   9255
         Begin prjBeeAV.AkuSayangIbu jcbutton2 
            Height          =   495
            Left            =   120
            TabIndex        =   168
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
            TabIndex        =   167
            Top             =   0
            Width           =   9015
            _extentx        =   15901
            _extenty        =   6165
            border          =   0
            style           =   4
            styleex         =   37
         End
         Begin prjBeeAV.AkuSayangIbu jcbutton3 
            Height          =   495
            Left            =   2040
            TabIndex        =   169
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
            TabIndex        =   170
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
         Begin prjBeeAV.ProgressBar pScan 
            Height          =   375
            Left            =   120
            TabIndex        =   233
            Top             =   600
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   661
            Value           =   0
            RoundedValue    =   2
            Theme           =   2
            TextStyle       =   2
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Digiface"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextForeColor   =   8421504
            Text            =   "ProgressBar1"
            TextEffectColor =   16777215
            TextEffect      =   4
         End
         Begin VB.PictureBox Picture21 
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
         Begin VB.Line Line8 
            BorderColor     =   &H00E1E1E1&
            X1              =   8640
            X2              =   8640
            Y1              =   360
            Y2              =   720
         End
         Begin VB.Line Line7 
            BorderColor     =   &H00E1E1E1&
            X1              =   7920
            X2              =   7920
            Y1              =   360
            Y2              =   720
         End
         Begin VB.Label cMdan3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Skip"
            ForeColor       =   &H80000011&
            Height          =   195
            Left            =   8760
            MouseIcon       =   "frmUI.frx":1B0A3A
            MousePointer    =   99  'Custom
            TabIndex        =   200
            Top             =   360
            Width           =   330
         End
         Begin VB.Label cMdan2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pause"
            ForeColor       =   &H80000011&
            Height          =   195
            Left            =   8040
            MouseIcon       =   "frmUI.frx":1B0B8C
            MousePointer    =   99  'Custom
            TabIndex        =   199
            Top             =   360
            Width           =   450
         End
         Begin VB.Label cMdan1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Stop"
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   7440
            MouseIcon       =   "frmUI.frx":1B0CDE
            MousePointer    =   99  'Custom
            TabIndex        =   198
            Top             =   360
            Width           =   360
         End
         Begin VB.Label lblFile 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000006&
            Height          =   255
            Left            =   2280
            TabIndex        =   145
            Top             =   0
            Width           =   6735
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
            ForeColor       =   &H009EC24F&
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
            Left            =   240
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
         Begin VB.Label Label91 
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
            Picture         =   "frmUI.frx":1B0E30
            Top             =   120
            Width           =   750
         End
         Begin VB.Image Image12 
            Height          =   750
            Left            =   240
            Picture         =   "frmUI.frx":1B2360
            Top             =   3000
            Width           =   750
         End
         Begin VB.Image Image13 
            Height          =   750
            Left            =   240
            Picture         =   "frmUI.frx":1B381F
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
            Picture         =   "frmUI.frx":1B4D8E
            Top             =   2040
            Width           =   750
         End
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00E0E0E0&
         X1              =   120
         X2              =   9120
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
         MouseIcon       =   "frmUI.frx":1B6258
         MousePointer    =   99  'Custom
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
         MouseIcon       =   "frmUI.frx":1B63AA
         MousePointer    =   99  'Custom
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
         MouseIcon       =   "frmUI.frx":1B64FC
         MousePointer    =   99  'Custom
         TabIndex        =   137
         Top             =   120
         Width           =   795
      End
   End
   Begin VB.PictureBox Menu4 
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
      Height          =   4935
      Left            =   2400
      ScaleHeight     =   4935
      ScaleWidth      =   9255
      TabIndex        =   76
      Top             =   1560
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
         Begin VB.PictureBox Picture11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1695
            Left            =   7200
            Picture         =   "frmUI.frx":1B664E
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
            ItemData        =   "frmUI.frx":1C129E
            Left            =   4080
            List            =   "frmUI.frx":1C12CC
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
            _extentx        =   6800
            _extenty        =   6588
            style           =   4
            styleex         =   33
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
            ItemData        =   "frmUI.frx":1C1352
            Left            =   240
            List            =   "frmUI.frx":1C13B0
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
            Picture         =   "frmUI.frx":1C1675
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
         MouseIcon       =   "frmUI.frx":1C1F61
         MousePointer    =   99  'Custom
         TabIndex        =   118
         Top             =   120
         Width           =   945
      End
      Begin VB.Label Label1221 
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
         MouseIcon       =   "frmUI.frx":1C20B3
         MousePointer    =   99  'Custom
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
         MouseIcon       =   "frmUI.frx":1C2205
         MousePointer    =   99  'Custom
         TabIndex        =   116
         Top             =   120
         Width           =   885
      End
   End
   Begin VB.PictureBox Menu2 
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
      Height          =   4935
      Left            =   2400
      ScaleHeight     =   4935
      ScaleWidth      =   9255
      TabIndex        =   53
      Top             =   1560
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
         MouseIcon       =   "frmUI.frx":1C2357
         MousePointer    =   99  'Custom
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
         MouseIcon       =   "frmUI.frx":1C24A9
         MousePointer    =   99  'Custom
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
      Height          =   4935
      Left            =   2400
      ScaleHeight     =   4935
      ScaleWidth      =   9255
      TabIndex        =   51
      Top             =   1560
      Width           =   9255
      Begin prjBeeAV.ucListView lvQuar 
         Height          =   4695
         Left            =   120
         TabIndex        =   52
         Top             =   120
         Width           =   9015
         _extentx        =   15901
         _extenty        =   8281
         border          =   0
         style           =   4
         styleex         =   33
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
      Height          =   4935
      Left            =   2400
      ScaleHeight     =   4935
      ScaleWidth      =   9255
      TabIndex        =   0
      Top             =   1560
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
            _extentx        =   19076
            _extenty        =   7223
            style           =   4
            styleex         =   33
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
            _extentx        =   15558
            _extenty        =   4895
            border          =   0
            styleex         =   37
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
            _extentx        =   15637
            _extenty        =   7223
            border          =   0
            styleex         =   33
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
         MouseIcon       =   "frmUI.frx":1C25FB
         MousePointer    =   99  'Custom
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
         MouseIcon       =   "frmUI.frx":1C274D
         MousePointer    =   99  'Custom
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
         MouseIcon       =   "frmUI.frx":1C289F
         MousePointer    =   99  'Custom
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
         MouseIcon       =   "frmUI.frx":1C29F1
         MousePointer    =   99  'Custom
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
      PictureNormal   =   "frmUI.frx":1C2B43
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
      PictureNormal   =   "frmUI.frx":1C3255
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
      PictureNormal   =   "frmUI.frx":1C3967
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
      PictureNormal   =   "frmUI.frx":1C4079
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
      PictureNormal   =   "frmUI.frx":1C478B
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
      PictureNormal   =   "frmUI.frx":1C4E9D
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin prjBeeAV.jcbutton cmdSkip 
      Height          =   375
      Left            =   13680
      TabIndex        =   163
      Top             =   3480
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
      Left            =   14640
      TabIndex        =   164
      Top             =   3000
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
      Left            =   11880
      TabIndex        =   165
      Top             =   3000
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
      Left            =   9600
      TabIndex        =   232
      Top             =   1200
      Width           =   1095
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
      Left            =   10800
      TabIndex        =   231
      Top             =   1200
      Width           =   735
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00E1E1E1&
      X1              =   712
      X2              =   712
      Y1              =   80
      Y2              =   96
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "CI-Soft Software 2014"
      Height          =   615
      Left            =   9960
      TabIndex        =   211
      Top             =   6720
      Width           =   3375
   End
   Begin prjBeeAV.UniDialog udScan 
      Left            =   2760
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      FileFlags       =   2621444
      FolderFlags     =   835
      FileCustomFilter=   "frmUI.frx":1C55AF
      FileDefaultExtension=   "frmUI.frx":1C55CF
      FileFilter      =   "frmUI.frx":1C55EF
      FileOpenTitle   =   "frmUI.frx":1C5637
      FileSaveTitle   =   "frmUI.frx":1C566F
      FolderMessage   =   "frmUI.frx":1C56A7
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
      TabIndex        =   194
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
      TabIndex        =   193
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
Attribute VB_Name = "frMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Oval
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
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
lAbout.Caption = "Bee Team Work"
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
lTool.Caption = "Bee - Lock"
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
If ck(3).value = 1 Then
    picSec.Visible = True
    lbSec.Caption = "Secured"
    lbSec.ForeColor = &H8000&
    picSec4.Visible = True
    Label6.ForeColor = &H8000&
    Label6.Caption = "On"
Else
    ck(3).value = 0
    picSec.Visible = False
     lbSec.Caption = "Not secured"
     lbSec.ForeColor = &HC0&
     picSec4.Visible = False
     Label6.ForeColor = &HC0&
     Label6.Caption = "Off"
End If


Case 6
If ck(6).value = 1 Then
picSec.Visible = True
    lbSec.Caption = "Secured"
    lbSec.ForeColor = &H8000&
    picSec2.Visible = True
    Label19.ForeColor = &H8000&
    Label19.Caption = "On"
Else
    ck(6).value = 0
    picSec.Visible = False
     lbSec.Caption = "Not secured"
     lbSec.ForeColor = &HC0&
     picSec2.Visible = False
     Label19.ForeColor = &HC0&
     Label19.Caption = "Off"
End If

Case 7
If ck(7).value = 1 Then
picSec.Visible = True
    jcbutton9.Visible = False
    lbSec.Caption = "Secured"
    lbSec.ForeColor = &H8000&
    picSec1.Visible = True
    Label15.ForeColor = &H8000&
    Label15.Caption = "On"
    frRTP.mnEP.Checked = True
    UpdateIcon Me.Icon, "Bee Real Time - Protection is ON", frRTP
    TampilkanBalon frRTP, "Your PC is Protect", "Protection Active", NIIF_INFO

Else
   ck(7).value = 0
   picSec.Visible = False
   jcbutton9.Visible = True
    lbSec.Caption = "Not secured"
    lbSec.ForeColor = &HC0&
    picSec1.Visible = False
    Label15.ForeColor = &HC0&
    Label15.Caption = "Off"
    frRTP.mnEP.Checked = False
    UpdateIcon Me.Icon, "Bee Real Time - Protection is OFF", frRTP
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
lAbout.Caption = "Bee Antivirus Information"
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
susu.Visible = True
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
UpdateIcon frRTP.Icon, "Bee RT-Protection, Bee 2014 Ver. 1.3", frRTP
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
lTool.Caption = "Bee - Lock"
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
lTool.Caption = "Bee Registry Tweaker"
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
            SetDwordValue SingkatanKey("HKCU"), X.Tag, Replace(X.name, " ", " "), X.value
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
If Dir(App.Path & "\Tools\Bee - Set Attributes.exe") = "" Then
MsgBox "Sorry, File Bee - Set Attributes Not Found.", 0 + vbExclamation, "Error"
Else
ShellExecute Me.hWnd, vbNullString, App.Path & "\Tools\Bee - Set Attributes.exe", vbNullString, "C:\", 1
End If

End Sub

Private Sub jcbutton9_Click()
    picSec.Visible = True
    ck(7).value = 1
    lbSec.Caption = "Secured"
    lbSec.ForeColor = &H8000&
    picSec1.Visible = True
    Label15.ForeColor = &H8000&
    Label15.Caption = "On"
    frRTP.mnEP.Checked = True
    UpdateIcon Me.Icon, "Bee Real Time - Protection is ON", frRTP
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


Private Sub picIconBee_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
Call Label11_Click
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
Call label871_Click
susu.Visible = False
Menu0.Visible = False
    Menu1.Visible = False
    Menu3.Visible = False
    Menu4.Visible = False
    Menu5.Visible = False
    Menu2.Visible = True
End Sub

Private Sub picMenu4_Click()
    Call Tool0_Click
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
Call Label1221_Click
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
pScan.value = 0
Picture3.Visible = False
Picture18.Visible = False
Picture19.Visible = False
picMenu3.Visible = False
picMenu4.Visible = False
picMenu6.Visible = False
cmdForward.Enabled = False
StopScan = False
If ck(7).value = 1 Then
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
If ck(7).value = 1 Then
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
    frMain.pScan.value = FileScan
    xPercent = pScan.value * 100 / pScan.Max
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
                
                If frMain.ck(1).value = 1 Then
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
    pScan.value = 0
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
    pScan.value = 0
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
        If X1 = vbNullString Then X.value = 0 Else X.value = Int(X1)
        End If
    Next
    
    txtOwner.Text = GetStringValue(SingkatanKey("HKLM"), "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner")
    txtOrg.Text = GetStringValue(SingkatanKey("HKLM"), "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOrganization")

End Sub


Private Sub cMdan1_Click()
Call cmdStop_Click
End Sub

Private Sub cMdan2_Click()
Call cmdPause_Click
End Sub

Private Sub cMdan3_Click()
Call cmdSkip_Click
End Sub

Private Sub cMdan1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cMdan1.ForeColor = &HCCFF&
cMdan2.ForeColor = &H80000011
cMdan3.ForeColor = &H80000011
End Sub

Private Sub cMdan2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cMdan2.ForeColor = &HCCFF&
cMdan1.ForeColor = &H80000011
cMdan3.ForeColor = &H80000011
End Sub

Private Sub cMdan3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cMdan3.ForeColor = &HCCFF&
cMdan2.ForeColor = &H80000011
cMdan1.ForeColor = &H80000011
End Sub

Private Sub picDetailed_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cMdan3.ForeColor = &H80000011
cMdan2.ForeColor = &H80000011
cMdan1.ForeColor = &H80000011
End Sub
Private Sub Label11_Click()
picScan(0).Visible = True
picScan(1).Visible = False
picDetailed.Visible = False
Label11.FontBold = True
Label120.FontBold = False
Label2.FontBold = False
End Sub

Private Sub Label1101_Click()
picAbout(2).Visible = True
picAbout(1).Visible = False
picAbout(0).Visible = False
Label1101.FontBold = True
Label133.FontBold = False
Label1221.FontBold = False
End Sub

Private Sub Label120_Click()
picScan(0).Visible = False
picScan(1).Visible = True
picDetailed.Visible = False
Label11.FontBold = False
Label120.FontBold = True
Label2.FontBold = False
End Sub

Private Sub Label133_Click()
Label1101.FontBold = False
Label133.FontBold = True
Label1221.FontBold = False
picAbout(1).Visible = True
picAbout(0).Visible = False
picAbout(2).Visible = False
End Sub

Private Sub Label2_Click()
picScan(0).Visible = False
picScan(1).Visible = False
picDetailed.Visible = True
Label11.FontBold = False
Label120.FontBold = False
Label2.FontBold = True
End Sub

Private Sub Label1221_Click()
Label1101.FontBold = False
Label133.FontBold = False
Label1221.FontBold = True
picAbout(0).Visible = True
picAbout(1).Visible = False
picAbout(2).Visible = False
End Sub

Private Sub Label298_Click()
label871.FontBold = False
Label298.FontBold = True
gGr.Visible = True
End Sub

Private Sub label871_Click()
label871.FontBold = True
Label298.FontBold = False
gGr.Visible = False
End Sub

Private Sub Tool0_Click()
Tool0.FontBold = True
Tool1.FontBold = False
Tool2.FontBold = False
Tool3.FontBold = False
picTool(0).Visible = True
picTool(1).Visible = False
picTool(2).Visible = False
picTool(3).Visible = False
End Sub
Private Sub Tool1_Click()
Tool1.FontBold = True
Tool0.FontBold = False
Tool2.FontBold = False
Tool3.FontBold = False
picTool(1).Visible = True
picTool(0).Visible = False
picTool(2).Visible = False
picTool(3).Visible = False
End Sub
Private Sub Tool2_Click()
Tool2.FontBold = True
Tool1.FontBold = False
Tool0.FontBold = False
Tool3.FontBold = False
picTool(2).Visible = True
picTool(1).Visible = False
picTool(0).Visible = False
picTool(3).Visible = False
End Sub
Private Sub Tool3_Click()
Tool3.FontBold = True
Tool1.FontBold = False
Tool2.FontBold = False
Tool0.FontBold = False
picTool(3).Visible = True
picTool(1).Visible = False
picTool(2).Visible = False
picTool(0).Visible = False
End Sub


