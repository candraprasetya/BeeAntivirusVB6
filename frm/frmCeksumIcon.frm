VERSION 5.00
Begin VB.Form frmCeksumIcon 
   Caption         =   "Form1"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTmpIcon 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   3
      ToolTipText     =   "Pengaturan PicTmpIcon HArus seperti in ( Standarnya )"
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Ceksum"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   2280
      Width           =   9255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Text            =   "F:\virus\fileman.exe"
      Top             =   960
      Width           =   4095
   End
End
Attribute VB_Name = "frmCeksumIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DRAW_ICO Text1.Text, picTmpIcon
End Sub
