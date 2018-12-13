VERSION 5.00
Begin VB.Form frSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "dqf wgwege"
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5175
   Icon            =   "frSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tm 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2880
      Top             =   1920
   End
   Begin prjBeeAV.ProgressBar pBsplash 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   4695
      _ExtentX        =   8281
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ee Antivirus"
      BeginProperty Font 
         Name            =   "Backslash"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000CCFF&
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   720
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00E1E1E1&
      Height          =   2655
      Left            =   0
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version : 1 . 3 . 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1380
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Your PC is Protected by Bee - Anti Virus"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2011- 2014 CI-Soft Software"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   2040
      TabIndex        =   1
      Top             =   2400
      Width           =   3000
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading......"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "D3 Honeycombism"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000CCFF&
      Height          =   555
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   3975
   End
End
Attribute VB_Name = "frSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Dint As Long

Public Sub Loading()
BuildLV
If ValidFile(App.Path & "\Definition\BeeDef.Bee") = True Then
ReadDb App.Path & "\Definition\BeeDef.Bee"
Else
MsgBox "Database Virus Not Found !", vbCritical + vbOKOnly, "Bee Error !"
Unload frRTP
Unload frMain
Unload Me
End
End If
InitPHPattern
LoadDataIcon
If GetSettingF(App.Path & "\Setting.ini") = False Then
MsgBox Chr(34) & "Setting.ini" & Chr(34) & "wasn't found" & vbCrLf & "Bee will use Default Setting", vbExclamation, "Bee"
End If
ApplySetting
frMain.Show
frMain.xToolIndexPic = 0
Unload Me
End Sub

Private Sub Form_Load()
If Dir(App.Path & "\Definition\BeeDef.Bee") = "" Then
    tm.Enabled = False
        LetakanForm frSplash, False
            MsgBox "Sorry, Database Virus Not Found", vbCritical + vbOKOnly, "Bee Error !"
                End
                    End If
If Dir(App.Path & "\Setting.ini") = "" Then
    tm.Enabled = False
        LetakanForm frSplash, False
            MsgBox "Sorry, File Component Not Found", vbCritical + vbOKOnly, "Bee Error !"
                End
                    End If
Dint = 0
Label8.Caption = vbNullString
If Dir(App.Path & "\Quarantine", vbDirectory) = "" Then MkDir (App.Path & "\Quarantine")
LetakanForm frSplash, True
End Sub

Private Sub tm_Timer()
If Dint = 4 Then tm.Enabled = False: Exit Sub
Dint = Dint + 1
Select Case Dint
Case 1
pBsplash.value = 5
pBsplash.Text = "5%"
Label8.Caption = "Loading Definition [5 %]"
ReadDb App.Path & "\Definition\BeeDef.Bee"
Exit Sub
Case 2
pBsplash.Text = "100%"
pBsplash.value = 100
Label8.Caption = "Starting Application [100 %]"
Exit Sub
Case 3
frSplash.Hide '<-- Formnya Disembunyikan
frMain.xToolIndexPic = 0
Exit Sub
Case 4
LetakanForm frSplash, False
Unload Me
frMases.Show
End Select
End Sub

Public Sub Loading2()
BuildLV
If ValidFile(App.Path & "\Definition\BeeDef.Bee") = True Then
ReadDb App.Path & "\Definition\BeeDef.Bee"
Else
MsgBox "Database Virus Not Found !", vbCritical + vbOKOnly, "Bee Error !"
Unload frRTP
Unload frMain
Unload Me
End
End If
InitPHPattern
LoadDataIcon
If GetSettingF(App.Path & "\Setting.ini") = False Then
MsgBox Chr(34) & "Setting.ini" & Chr(34) & "wasn't found" & vbCrLf & "Bee will use Default Setting", vbExclamation, "Bee"
End If
ApplySetting
frMain.Show
frMain.xToolIndexPic = 0
Unload Me
End Sub
