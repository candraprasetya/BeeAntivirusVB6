VERSION 5.00
Begin VB.Form frMases 
   BorderStyle     =   0  'None
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5175
   Icon            =   "frMases.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frMases.frx":000C
   ScaleHeight     =   2535
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   4440
      Picture         =   "frMases.frx":2AC3C
      ScaleHeight     =   270
      ScaleMode       =   0  'User
      ScaleWidth      =   660
      TabIndex        =   4
      Top             =   0
      Width           =   666
   End
   Begin VB.Timer Timer3 
      Interval        =   5000
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   2760
      ScaleHeight     =   945
      ScaleWidth      =   4425
      TabIndex        =   0
      Top             =   2640
      Width           =   4455
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         Picture         =   "frMases.frx":2B0B3
         ScaleHeight     =   735
         ScaleWidth      =   735
         TabIndex        =   2
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Real - Time Protection Is ON"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   3180
      End
   End
   Begin VB.PictureBox cCloseSys 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   4440
      Picture         =   "frMases.frx":2B7D5
      ScaleHeight     =   270
      ScaleWidth      =   660
      TabIndex        =   5
      Top             =   0
      Width           =   666
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Real - Time Protection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   330
      Left            =   2040
      TabIndex        =   1
      Top             =   1200
      Width           =   3045
   End
End
Attribute VB_Name = "frMases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cCloseSys_Click()
Timer2.Enabled = True
End Sub

Private Sub Form_Load()
LetakanForm frMases, True
Me.Top = Screen.Height
Me.Left = Screen.Width - Me.Width
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3.Visible = True
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3.Visible = False
End Sub

Private Sub Timer1_Timer()
If Me.Top <= Screen.Height - (Me.Height + 250) Then Timer1.Enabled = False
Me.Top = Me.Top - 200
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If Me.Top >= 11000 Then
Timer2.Enabled = False
Unload Me
End If
Me.Top = Me.Top + 300
Unload Me
End Sub
Private Sub Timer3_Timer()
Timer2.Enabled = True
Timer3.Enabled = False
End Sub



