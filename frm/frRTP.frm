VERSION 5.00
Begin VB.Form frRTP 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bee Real - Time Protection"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9240
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frRTP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000CCFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   9255
      TabIndex        =   5
      Top             =   0
      Width           =   9255
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Virus Detected !!!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   120
         Width           =   3375
      End
   End
   Begin prjBeeAV.jcbutton jcbutton3 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   1695
      _extentx        =   2990
      _extenty        =   873
      buttonstyle     =   8
      font            =   "frRTP.frx":19F7A
      backcolor       =   52479
      caption         =   "More   "
      picturenormal   =   "frRTP.frx":19F9E
      pictureeffectonover=   0
      pictureeffectondown=   0
      captioneffects  =   0
      picturealign    =   3
      forecolor       =   8421504
      tooltipbackcolor=   0
      colorscheme     =   3
   End
   Begin prjBeeAV.jcbutton jcbutton2 
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   3600
      Width           =   2175
      _extentx        =   3836
      _extenty        =   873
      buttonstyle     =   8
      font            =   "frRTP.frx":1A338
      backcolor       =   52479
      caption         =   "Quarantine All"
      picturenormal   =   "frRTP.frx":1A35C
      pictureeffectonover=   0
      pictureeffectondown=   0
      captioneffects  =   0
      forecolor       =   8421504
      tooltipbackcolor=   0
      colorscheme     =   3
   End
   Begin prjBeeAV.jcbutton jcbutton1 
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   3600
      Width           =   1815
      _extentx        =   3201
      _extenty        =   873
      buttonstyle     =   8
      font            =   "frRTP.frx":1A6F6
      backcolor       =   52479
      caption         =   "Delete All"
      picturenormal   =   "frRTP.frx":1A71A
      pictureeffectonover=   0
      pictureeffectondown=   0
      captioneffects  =   0
      forecolor       =   8421504
      tooltipbackcolor=   0
      colorscheme     =   3
   End
   Begin prjBeeAV.jcbutton cClose 
      Height          =   495
      Left            =   7560
      TabIndex        =   1
      Top             =   3600
      Width           =   1575
      _extentx        =   2778
      _extenty        =   873
      buttonstyle     =   8
      font            =   "frRTP.frx":1AAB4
      backcolor       =   52479
      caption         =   "Close"
      pictureeffectonover=   0
      pictureeffectondown=   0
      captioneffects  =   0
      forecolor       =   8421504
      tooltipbackcolor=   0
      colorscheme     =   3
   End
   Begin prjBeeAV.ucListView lvRTP 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   9015
      _extentx        =   12726
      _extenty        =   4895
      border          =   0
      style           =   4
      styleex         =   37
   End
   Begin prjBeeAV.rtp_mode rtp_mode1 
      Index           =   0
      Left            =   8280
      Top             =   0
      _extentx        =   1720
      _extenty        =   450
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   6  'Inside Solid
      X1              =   7200
      X2              =   7200
      Y1              =   3240
      Y2              =   4200
   End
   Begin VB.Menu mnFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnOP 
         Caption         =   "Open User Interfaces"
      End
      Begin VB.Menu bR1 
         Caption         =   "-"
      End
      Begin VB.Menu mnROS 
         Caption         =   "Run on Start-Up"
      End
      Begin VB.Menu mnEP 
         Caption         =   "Enable Protection"
      End
      Begin VB.Menu mnTO 
         Caption         =   "Tools External"
         Begin VB.Menu mnSBL 
            Caption         =   "Site Blocking"
         End
         Begin VB.Menu mnFG 
            Caption         =   "Fixed Registry"
         End
         Begin VB.Menu mnSAT 
            Caption         =   "Set Attributes"
         End
         Begin VB.Menu mnVK 
            Caption         =   "Virtual Keyboard"
         End
      End
      Begin VB.Menu bR2 
         Caption         =   "-"
      End
      Begin VB.Menu mnAD 
         Caption         =   "About Bee Antivirus"
      End
      Begin VB.Menu mnEA 
         Caption         =   "Exit Application"
      End
   End
   Begin VB.Menu mnlvRTP 
      Caption         =   "lvRTP"
      Visible         =   0   'False
      Begin VB.Menu mnFA 
         Caption         =   "Fix All Item[s]"
         Begin VB.Menu mnDel 
            Caption         =   "Delete"
         End
         Begin VB.Menu mnQuar 
            Caption         =   "Quarantine"
         End
      End
      Begin VB.Menu mnFC 
         Caption         =   "Fix Checked Item[s]"
         Begin VB.Menu mnDelC 
            Caption         =   "Delete"
         End
         Begin VB.Menu mnQuarC 
            Caption         =   "Quarantine"
         End
      End
      Begin VB.Menu b3 
         Caption         =   "-"
      End
      Begin VB.Menu mnIA 
         Caption         =   "Ignore All Item[s]"
      End
      Begin VB.Menu mnIC 
         Caption         =   "Ignore Checked Item[s]"
      End
   End
End
Attribute VB_Name = "frRTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ShellIE As SHDocVw.ShellWindows
Attribute ShellIE.VB_VarHelpID = -1
Dim isCompatch As Boolean
Private Sub cClose_Click()
lvRTP.ListItems.Clear
Me.Hide
LetakanForm frRTP, False
End Sub

Public Sub Form_Load()
On Error Resume Next
frRTP.Icon = frIcon.Icon
CompactObject
DisableClose frRTP, True
LetakanForm frRTP, True
If isFromContext = False Then
TampilkanBalon frRTP, "Your PC is Protect", "Protection Active", NIIF_INFO
UpdateIcon Me.Icon, "Bee RT-Protection, Bee Antivirus 2014 Ver. 1.3", frRTP
End If
End Sub

Private Sub jcbutton1_Click()
FixVir FixAll, lvRTP, del
Call cClose_Click
End Sub

Private Sub jcbutton2_Click()
FixVir FixAll, lvRTP, Quar
Call cClose_Click
End Sub

Private Sub jcbutton3_Click()
PopupMenu mnlvRTP
End Sub

Private Sub mnAD_Click()
frMain.WindowState = 0
frMain.Show
With frMain
    .Menu0.Visible = False
    .Menu1.Visible = False
    .Menu2.Visible = False
    .Menu3.Visible = False
    .Menu5.Visible = False
    .Menu4.Visible = True
    .xAboutIndexPic = 1
    If .xAboutIndexPic - 1 < 0 Then Exit Sub
    .picAbout(.xAboutIndexPic).Visible = False
    .picAbout(.xAboutIndexPic - 1).Visible = True
    .xAboutIndexPic = .xAboutIndexPic - 1
        Select Case .xAboutIndexPic
    Case 0
    .cBackAbout.Enabled = False
    .cForwardAbout.Enabled = True
    .lAbout.Caption = "Bee Team Work"
    Case 1
    .cBackAbout.Enabled = True
    .cForwardAbout.Enabled = True
    .lAbout.Caption = "Thank's to"
    End Select

End With
End Sub

Private Sub mnEA_Click()
If StopScan = False Then
If MsgBox("Scanning is in progress !" & vbCrLf & "Do you want to exit now ?", vbExclamation + vbYesNo, "Warning !") = vbYes Then
Shell_NotifyIcon NIM_DELETE, nID
Unload frMain
Unload Me
End
End If
Else
If MsgBox("Are you sure want to exit ?", vbExclamation + vbYesNo, "Warning !") = vbYes Then
Shell_NotifyIcon NIM_DELETE, nID
Unload frMain
Unload Me
End
End If
End If
End Sub

Private Sub mnEP_Click()
If mnEP.Checked = True Then
    frMain.jcbutton9.Visible = True
    frMain.picSec.Visible = False
    frMain.lbSec.Caption = "Not secured"
    frMain.lbSec.ForeColor = &HC0&
    frMain.picSec1.Visible = False
    frMain.Label15.ForeColor = &HC0&
    frMain.Label15.Caption = "Off"
    ScanRTPmod = False
    frRTP.mnEP.Checked = False
    frMain.ck(7).Value = 0
    UpdateIcon Me.Icon, "Bee Real Time - Protection is OFF", frRTP
    TampilkanBalon frRTP, "Your PC is Not Protect", "Protection Not Active", NIIF_ERROR
Else
    frMain.jcbutton9.Visible = False
    frMain.picSec.Visible = True
    frMain.lbSec.Caption = "Secured"
    frMain.lbSec.ForeColor = &H8000&
    frMain.picSec1.Visible = True
    frMain.Label15.ForeColor = &H8000&
    frMain.Label15.Caption = "On"
    ScanRTPmod = True
    frRTP.mnEP.Checked = True
    frMain.ck(7).Value = 1
    UpdateIcon Me.Icon, "Bee Real Time - Protection is ON", frRTP
    TampilkanBalon frRTP, "Your PC is Protect", "Protection Active", NIIF_INFO
End If
End Sub

Private Sub mnFG_Click()
If Dir(App.Path & "\Tools\Fixed Registry.exe") = "" Then
MsgBox "Sorry, File Fixed Registry.exe Not Found.", 0 + vbExclamation, "Error"
Else
ShellExecute Me.hWnd, vbNullString, App.Path & "\Tools\Fixed Registry.exe", vbNullString, "C:\", 1
End If
End Sub

Private Sub mnOP_Click()
frMain.WindowState = 0
frMain.Show
End Sub

Private Sub mnROS_Click()
If mnROS.Checked = True Then
    RunSU 0
Else
    RunSU 1
End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lHasil  As Long
Dim HorX    As Long
    
    If Me.ScaleMode = vbPixels Then
        HorX = X
    Else
        HorX = X / Screen.TwipsPerPixelX
    End If
    
    Select Case HorX
        Case WM_LBUTTONDBLCLK
            frMain.WindowState = vbNormal
            frMain.Show
            lHasil = SetForegroundWindow(Me.hWnd)
        Case WM_RBUTTONUP 'Tampilkan menu Popup saat klik kanan.
            lHasil = SetForegroundWindow(Me.hWnd)
            Me.PopupMenu Me.mnFile, , , , Me.mnOP
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
End Sub

Private Sub lvRTP_ContextMenu(ByVal X As Single, ByVal Y As Single)
PopupMenu mnlvRTP
End Sub

Private Sub mnDel_Click()
FixVir FixAll, lvRTP, del
Call cClose_Click
End Sub

Private Sub mnDelC_Click()
FixVir FixChk, lvRTP, del
End Sub

Private Sub mnIA_Click()
FixVir IgnAll, lvRTP, None
Call cClose_Click
End Sub

Private Sub mnIC_Click()
FixVir IgnChk, lvRTP, None
End Sub

Private Sub mnOS_Click()
frMain.Show
End Sub

Private Sub mnQuar_Click()
FixVir FixAll, lvRTP, Quar
Call cClose_Click
End Sub

Private Sub mnQuarC_Click()
FixVir FixChk, lvRTP, Quar
End Sub

Sub CompactObject() ' untuk aktifkan rtp
On Error Resume Next
isCompatch = True
   Dim i As Integer, Cnt As Integer
   For i = 0 To rtp_mode1.count - 1
       rtp_mode1(i).SetIENothing
   Next i
       
   Set ShellIE = Nothing
   For i = 1 To rtp_mode1.count - 1
        Unload rtp_mode1(i)
   Next i
   
   Set ShellIE = New SHDocVw.ShellWindows
   Cnt = ShellIE.count - 1
   For i = 0 To Cnt
       If i > 0 Then
          AddIEObj i
       End If
          rtp_mode1(i).AddSubClass ShellIE(i)
   Next i
isCompatch = False
End Sub
Sub AddIEObj(Index As Integer)
On Error GoTo salah
    Load rtp_mode1(Index)
salah:
End Sub

Function FindID(ID As Long) As Boolean
On Error GoTo salah
    Dim i As Integer
    For i = 0 To rtp_mode1.count - 1
        If rtp_mode1(i).IEKey = ID Then
           FindID = True
        End If
    Next i
salah:
End Function

Function file_isFolder(Path As String) As Long
On Error GoTo salah

Dim ret As VbFileAttribute
    ret = GetAttr(Path) And vbDirectory
    If ret = vbDirectory Then
        file_isFolder = 1
    Else
        file_isFolder = 0
    End If
    
Exit Function

salah:
file_isFolder = -1
End Function

Private Sub mnSAT_Click()
If Dir(App.Path & "\Tools\Bee - Set Attributes.exe") = "" Then
MsgBox "Sorry, File Bee - Set Attributes Not Found.", 0 + vbExclamation, "Error"
Else
ShellExecute Me.hWnd, vbNullString, App.Path & "\Tools\Bee - Set Attributes.exe", vbNullString, "C:\", 1
End If
End Sub

Private Sub mnSBL_Click()
If Dir(App.Path & "\Tools\Site Blocking.exe") = "" Then
MsgBox "Sorry, File Site Blocking.exe Not Found.", 0 + vbExclamation, "Error"
Else
ShellExecute Me.hWnd, vbNullString, App.Path & "\Tools\Site Blocking.exe", vbNullString, "C:\", 1
End If
End Sub

Private Sub mnVK_Click()
If Dir(App.Path & "\Tools\Virtual Keyboard.exe") = "" Then
MsgBox "Sorry, File Virtual Keyboard.exe Not Found.", 0 + vbExclamation, "Error"
Else
ShellExecute Me.hWnd, vbNullString, App.Path & "\Tools\Virtual Keyboard.exe", vbNullString, "C:\", 1
End If
End Sub

Private Sub rtp_mode1_PathChange(Index As Integer, strPath As String)
If ScanRTPmod = True Then
    ScanRTP strPath
    If frRTP.lvRTP.ListItems.count > 0 Then
    frRTP.WindowState = 0
    frRTP.Show
    LetakanForm frRTP, True
    End If
End If
End Sub

Private Sub ShellIE_WindowRegistered(ByVal lCookie As Long)
    Intip
End Sub
Sub Intip()
If isCompatch = False Then
   Dim i As Integer, Cnt As Integer
   Cnt = ShellIE.count - 1
   For i = 0 To Cnt
       If (rtp_mode1.count - 1) < Cnt Then
          AddIEObj i
       End If
          If FindID(ShellIE(i).hWnd) = False Then
             rtp_mode1(i).EnabledMonitoring True
             rtp_mode1(i).AddSubClass ShellIE(i)
          End If
   Next i
End If
End Sub
