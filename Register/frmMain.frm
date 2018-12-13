VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Software title goes here."
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9855
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar pbx 
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   7440
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   120
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   7185
      ScaleWidth      =   9585
      TabIndex        =   4
      Top             =   120
      Width           =   9615
      Begin VB.Label lblProductVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "ProductVersion"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   9135
      End
      Begin VB.Line LineMiddle 
         BorderColor     =   &H00FFFFFF&
         X1              =   240
         X2              =   9360
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label lblProductSlogan 
         BackStyle       =   0  'Transparent
         Caption         =   "...Your slogan goes here."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   840
         Width           =   8775
      End
      Begin VB.Label lblProductName 
         BackStyle       =   0  'Transparent
         Caption         =   "ProductName"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   9135
      End
   End
   Begin SimpleTrial.XpBs CmdExit 
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Top             =   7800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Exit"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin SimpleTrial.XpBs CmdKGen 
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   7800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Key Generator"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin SimpleTrial.XpBs CmdEnter 
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   7800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Enter Trial"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin SimpleTrial.XpBs CmdEntSerial 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   7800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Enter Serial"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin VB.Timer TmrMain 
      Interval        =   100
      Left            =   10080
      Top             =   3960
   End
   Begin VB.Label lblTrialDelay 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Trial Delay:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   7440
      Width           =   9615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*********************************************'
'                                             '
' SimpleTrial                                 '
' Feel free to re-distrubute this code, since '
' this code is freeware :).                   '
'                                             '
' Please vote for me.                         '
'                                             '
'*********************************************'

Private Sub CmdAbout_Click()

    'Show details about your software.
        MsgBox "Company Name: " & App.CompanyName & vbCrLf & "Product Name: " & App.ProductName & vbCrLf & "Version: " & App.Major & "." & App.Revision & "." & App.Minor & vbCrLf & vbCrLf & "Little message about your product here.."

End Sub

Private Sub CmdEnter_Click()

    'Load the software.
        frmSoftware.Show
        
    'Add the unregistered status to the software.
        frmSoftware.Caption = "" & App.ProductName & " (Unregistered Version)": Unload Me

End Sub

Private Sub CmdExit_Click()

    'Terminate the program if the user decides to.
        Unload Me

End Sub

Private Sub CmdKGen_Click()

    'Load the Key Generator form.
        MsgBox "Please ensure that you remove this from your software, otherwise people will be registering your product for free.", vbInformation, "Quick Notice!"
        frmKeyGen.Show

End Sub

Private Sub Form_Load()

    On Error Resume Next

    Dim USR As String
    Dim SERIAL As String
    
    lblProductName.Caption = "Your product name."
    lblProductSlogan.Caption = "...Your slogan goes here."
    lblProductVersion.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
    
    USR = EnigmaDecrypt(GetSetting("SimpleTrial", "Main", "User"))
    SERIAL = EnigmaDecrypt(GetSetting("SimpleTrial", "Main", "Serial"))
    
    If KeyGen(USR, "™¤9DP0l@4Q", 2) = SERIAL Then frmSoftware.Show: Unload Me
    
End Sub

Private Sub CmdEntSerial_Click()

    'Load the details entry form.
        EnterDetails.Show
End Sub

Private Sub TmrMain_Timer()
        
    'Start the trial delay..
    pbx.Value = pbx.Value + 1
    If pbx.Value = "100" Then CmdEnter.Enabled = True: TmrMain.Enabled = False
    
End Sub
