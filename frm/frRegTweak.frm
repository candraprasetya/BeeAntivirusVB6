VERSION 5.00
Begin VB.Form frRegTweak 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registry Tweaker"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frRegTweak.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   8895
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fCusXP 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Register Owner and Organization"
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   120
      TabIndex        =   24
      Top             =   3480
      Width           =   3975
      Begin VB.TextBox txtOrg 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox txtOwner 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Organization"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Owner"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   555
      End
   End
   Begin prjDAA.jcbutton cSave 
      Height          =   375
      Left            =   7200
      TabIndex        =   23
      Top             =   4680
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   "Save Setting"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.Frame fStartMenu 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Desktop and Start Menu"
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   3975
      Begin VB.CheckBox NoRecentDocsHistory 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Do not keep history of recent opened documents "
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   120
         TabIndex        =   20
         Tag             =   "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
         Top             =   1200
         Width           =   3075
      End
      Begin VB.CheckBox NoDesktop 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hide and disable items on desktop"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   3255
      End
      Begin VB.CheckBox NoFind 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Remove Search form Start menu"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Tag             =   "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
         Top             =   960
         Width           =   3135
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
         TabIndex        =   17
         Tag             =   "Microsoft\Windows\CurrentVersion\Policies\Explorer"
         Top             =   3000
         Width           =   2775
      End
      Begin VB.CheckBox RestrictRun 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enable Restrict Run"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Tag             =   "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame fExplorer 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Windows Explorer"
      ForeColor       =   &H00FF0000&
      Height          =   3135
      Left            =   4200
      TabIndex        =   7
      Top             =   1200
      Width           =   4575
      Begin VB.CheckBox NoSaveSettings 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Don't save settings at exit"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Tag             =   "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
         Top             =   2760
         Width           =   2655
      End
      Begin VB.CheckBox NoViewContextMenu 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Remove Windows Explorer's default context menu"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   120
         TabIndex        =   21
         Tag             =   "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
         Top             =   2280
         Width           =   4395
      End
      Begin VB.CheckBox NameNumericTail 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Remove the Tildes in Short Filenames ""~"""
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Tag             =   "SYSTEM\CurrentControlSet\Control\FileSystem"
         Top             =   2040
         Width           =   3735
      End
      Begin VB.CheckBox FullPathAddress 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Full Path at Address Bar"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Tag             =   "Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState"
         Top             =   1680
         Width           =   3375
      End
      Begin VB.CheckBox NoSetTaskbar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lock Taksbar Setting"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Tag             =   "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
         Top             =   960
         Width           =   2055
      End
      Begin VB.CheckBox NoTrayContextMenu 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Disable Tray Menu"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Tag             =   "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CheckBox ShowSuperHidden 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show System File"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Tag             =   "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
         Top             =   600
         Width           =   2055
      End
      Begin VB.CheckBox Hidden 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Hidden File"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Tag             =   "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame fCpanel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Control Panel"
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   4200
      TabIndex        =   4
      Top             =   120
      Width           =   4575
      Begin VB.CheckBox NoControlPanel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Prohibit access to the Control Panel"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Tag             =   "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
         Top             =   600
         Width           =   3795
      End
      Begin VB.CheckBox NoAddRemovePrograms 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Remove Add/Programs"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Tag             =   "Microsoft\Windows\CurrentVersion\Policies\Uninstall\NoAddRemovePrograms"
         Top             =   240
         Width           =   2715
      End
   End
   Begin VB.Frame fSystem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "System"
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.CheckBox NoFolderOptions 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Disable Folder Options"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Tag             =   "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
         Top             =   1320
         Width           =   2895
      End
      Begin VB.CheckBox DisableCMD 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Disable CMD"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Tag             =   "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
         Top             =   960
         Width           =   2055
      End
      Begin VB.CheckBox DisableTaskMgr 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Disable Task Manager"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Tag             =   "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
         Top             =   600
         Width           =   3255
      End
      Begin VB.CheckBox DisableRegistryTools 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Disable Registry Editor"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Tag             =   "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
         Top             =   240
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frRegTweak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cSave_Click()
    Dim X
    
    ' Update according to each control
    For Each X In Controls
        If X.Tag <> "" Then _
            SetDwordValue SingkatanKey("HKCU"), X.Tag, Replace(X.Name, " ", " "), X.Value
    Next
    
    Call SetStringValue(SingkatanKey("HKLM"), "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner", txtOwner.Text)
    Call SetStringValue(SingkatanKey("HKLM"), "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOrganization", txtOrg.Text)
    
    MsgBox "Done !", vbInformation, "Dasanggra"
    If MsgBox("Do you want to Restart Explorer to take the effect ?", vbQuestion + vbYesNo, "Dasanggra") = vbYes Then
        KillByProccess "explorer.exe"
        ShellExecute Me.hwnd, vbNullString, GetSpecFolder(WINDOWS_DIR) & "\explorer.exe", vbNullString, "C:\", 1
    End If
End Sub

Private Sub Form_Load()
    Dim X, X1
    
    For Each X In Controls
        If X.Tag <> "" Then
            X1 = GetDwordValue(SingkatanKey("HKCU"), X.Tag, Replace(X.Name, " ", " "))
        If X1 = vbNullString Then X.Value = 0 Else X.Value = Int(X1)
        End If
    Next
    
    txtOwner.Text = GetStringValue(SingkatanKey("HKLM"), "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner")
    txtOrg.Text = GetStringValue(SingkatanKey("HKLM"), "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOrganization")
End Sub
