VERSION 5.00
Begin VB.Form frIcon 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Auto Scan FlashDisk"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5400
   Icon            =   "frIcon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text4 
      Height          =   975
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "frIcon.frx":19F7A
      Top             =   3720
      Width           =   4935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Descripty"
      Height          =   615
      Left            =   1920
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   975
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frIcon.frx":19F97
      Top             =   2640
      Width           =   4935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Enpryp"
      Height          =   615
      Left            =   3600
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frIcon.frx":19FBC
      Top             =   1560
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Cheksum"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frIcon.frx":19FDB
      Top             =   120
      Width           =   4935
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
   Begin VB.Menu mnST 
      Caption         =   "Startup"
      Visible         =   0   'False
      Begin VB.Menu mnRef 
         Caption         =   "Refresh"
      End
      Begin VB.Menu b2 
         Caption         =   "-"
      End
      Begin VB.Menu mnES 
         Caption         =   "Enable Start-Up"
      End
      Begin VB.Menu mnDS 
         Caption         =   "Disable Start-Up"
      End
      Begin VB.Menu mnDelS 
         Caption         =   "Delete Start-Up"
      End
      Begin VB.Menu b1 
         Caption         =   "-"
      End
      Begin VB.Menu mnEx 
         Caption         =   "Explore File"
      End
      Begin VB.Menu mnP 
         Caption         =   "Properties"
      End
   End
End
Attribute VB_Name = "frIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim TmpHGlobal   As Long
Dim sPath As String
sPath = Text1.Text
TmpHGlobal = GetHandleFile(sPath)
Text2.Text = MYCeksum(sPath, TmpHGlobal)
End Sub



Private Sub Command2_Click()
CryptVirus Text1.Text, Text3.Text
End Sub

Private Sub Command3_Click()
DeCryptVirus Text1.Text, Text4.Text
End Sub



Private Sub mnDeleteQ_Click()
Static i As Integer

For i = 1 To frMain.lvQuar.ListItems.count
    HapusFile frMain.lvQuar.ListItems.Item(i).SubItem(3).Text
Next
MsgBox "Done !", vbInformation
GetQuarFile
End Sub

Private Sub mnDeleteSelQ_Click()
Static i As Integer

For i = 1 To frMain.lvQuar.ListItems.count
    If frMain.lvQuar.ListItems.Item(i).Selected = True Then
        HapusFile frMain.lvQuar.ListItems.Item(i).SubItem(3).Text
    End If
Next
MsgBox "Done !", vbInformation
GetQuarFile
End Sub

Private Sub mnExMalLV_Click()
Static i As Long

For i = 1 To frMain.lvMal.ListItems.count
    If frMain.lvMal.ListItems.Item(i).Selected = True Then
    ExploreTheFile frMain.lvMal.ListItems.Item(i).SubItem(2).Text
    End If
Next
End Sub

Private Sub mnExpProc_Click()
Static i As Long

For i = 1 To frMain.lvProcess.ListItems.count
    If frMain.lvProcess.ListItems.Item(i).Selected = True Then
    ExploreTheFile frMain.lvProcess.ListItems.Item(i).SubItem(9).Text
    End If
Next
End Sub

Private Sub mnFADel_Click()
FixVir FixAll, frMain.lvMal, del
End Sub

Private Sub mnFAQuar_Click()
FixVir FixAll, frMain.lvMal, Quar
End Sub

Private Sub mnFCDel_Click()
FixVir FixChk, frMain.lvMal, del
End Sub

Private Sub mnFCQuar_Click()
FixVir FixChk, frMain.lvMal, Quar
End Sub

Private Sub mnIAll_Click()
FixVir IgnAll, frMain.lvMal, None
End Sub

Private Sub mnIChk_Click()
FixVir IgnChk, frMain.lvMal, None
End Sub

Private Sub mnKillProc_Click()
Static i As Long
Dim PID    As Long
Dim sPath  As String

For i = 1 To frMain.lvProcess.ListItems.count
    If frMain.lvProcess.ListItems.Item(i).Selected = True Then
    If MsgBox("Do you want to kill this Process ?" & vbCrLf & "This Process Maybe needed by System", vbExclamation + vbYesNo, "Warning !") = vbNo Then Exit For
    PID = CLng(frMain.lvProcess.ListItems.Item(i).SubItem(3).Text)
    sPath = frMain.lvProcess.ListItems.Item(i).SubItem(9).Text
    If sPath = App.Path & "\" & App.EXEName & ".exe" Then MsgBox "Cannot Kill it self !", vbCritical + vbOKOnly, "Error !": Exit Sub
    If KillProses(PID, sPath, False, True) = True Then
            MsgBox "Success", vbInformation, "Informaton"
            ENUM_PROSES frMain.lvProcess, frMain.picBuff
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

For i = 1 To frMain.lvProcess.ListItems.count
    If frMain.lvProcess.ListItems.Item(i).Selected = True Then
      PID = CLng(frMain.lvProcess.ListItems.Item(i).SubItem(3).Text)
      frMain.lvProcess.ListItems.Item(i).SubItem(10).Text = SuspendProses(PID, True)
    End If
Next
End Sub

Private Sub mnRefreshP_Click()
ENUM_PROSES frMain.lvProcess, frMain.picBuff
End Sub

Private Sub mnResProc_Click()
Static i As Long
Dim PID    As Long

For i = 1 To frMain.lvProcess.ListItems.count
    If frMain.lvProcess.ListItems.Item(i).Selected = True Then
      PID = CLng(frMain.lvProcess.ListItems.Item(i).SubItem(3).Text)
      frMain.lvProcess.ListItems.Item(i).SubItem(10).Text = SuspendProses(PID, False)
    End If
Next
End Sub

Private Sub mnRSel_Click()
Static i As Integer
Static MalFrom As String
Static MalSend As String
If MsgBox("This File may be a virus that will Infected your Computer" & vbCrLf & "Are you sure want to restore it ?", vbExclamation + vbYesNo, "Warning !") = vbNo Then Exit Sub
For i = 1 To frMain.lvQuar.ListItems.count
    If frMain.lvQuar.ListItems.Item(i).Selected = True Then
    MalFrom = frMain.lvQuar.ListItems.Item(i).SubItem(3).Text
    MalSend = frMain.lvQuar.ListItems.Item(i).SubItem(2).Text
    RestoreFileQuar MalFrom, MalSend
    End If
Next
MsgBox "Done !", vbInformation
GetQuarFile
End Sub

Private Sub mnRSelTo_Click()
Static i As Integer
Static MalFrom As String
Static MalSend As String
Static BFF As String

If MsgBox("This File may be a virus that will Infected your Computer" & vbCrLf & "Are you sure want to restore it ?", vbExclamation + vbYesNo, "Warning !") = vbNo Then Exit Sub
For i = 1 To frMain.lvQuar.ListItems.count
    If frMain.lvQuar.ListItems.Item(i).Selected = True Then
        MalFrom = frMain.lvQuar.ListItems.Item(i).SubItem(3).Text
        BFF = BrowseForFolder(Me.hWnd, "Choose Path :")
        If Len(BFF) > 0 Then
            MalSend = BFF & "\" & frMain.lvQuar.ListItems.Item(i).Text
            RestoreFileQuar MalFrom, MalSend
        End If
    End If
Next
MsgBox "Done !", vbInformation
GetQuarFile
End Sub

Private Sub mnDelS_Click()
Static Inter As Integer
Dim pCada    As String
Dim mKeyReg  As Long
Dim pReg     As String
Dim sNama    As String
Dim sData    As String

Inter = 1
For Inter = 1 To frMain.lvStartup.ListItems.count
    If frMain.lvStartup.ListItems.Item(Inter).Selected = True Then
    pCada = frMain.lvStartup.ListItems.Item(Inter).SubItem(2).Text
        mKeyReg = SingkatanKey(Left$(pCada, 4))
        pReg = Right$(pCada, Len(pCada) - 5)
        sNama = frMain.lvStartup.ListItems.Item(Inter).Text
        sData = frMain.lvStartup.ListItems.Item(Inter).SubItem(3).Text
        If sData <> "None" Then
        Call DeleteValue(mKeyReg, pReg, sNama)
        Else
        KillByProccess GetPath(frMain.lvStartup.ListItems.Item(Inter).SubItem(4).Text, FileName)
        If HapusFile(frMain.lvStartup.ListItems.Item(Inter).SubItem(4).Text) = False Then MsgBox "Can't delete Startup !" + Chr(13) + "File is in use !"
        End If
        MsgBox "Success !", vbInformation
        GetRegStartup frMain.lvStartup
        Exit For
    End If
Next

End Sub

Private Sub mnDS_Click()
Static Inter As Integer
Dim pCada   As String
Dim mKeyReg As Long
Dim pReg    As String
Dim sData   As String
Dim sNama   As String
Dim sTarget As String

Inter = 1
For Inter = 1 To frMain.lvStartup.ListItems.count
    If frMain.lvStartup.ListItems.Item(Inter).Selected = True Then
    pCada = frMain.lvStartup.ListItems.Item(Inter).SubItem(2).Text
        mKeyReg = SingkatanKey(Left$(pCada, 4))
        sData = frMain.lvStartup.ListItems.Item(Inter).SubItem(3).Text
        sTarget = frMain.lvStartup.ListItems.Item(Inter).SubItem(4).Text
        pReg = Right$(pCada, Len(pCada) - 5)
        sNama = frMain.lvStartup.ListItems.Item(Inter).Text
        If sData <> "None" Then
        Call CreateKeyReg(mKeyReg, "Software\Microsoft\Windows\CurrentVersion\Run-")
        Call DeleteValue(mKeyReg, pReg, sNama)
        Call SetStringValue(mKeyReg, "Software\Microsoft\Windows\CurrentVersion\Run-", sNama, sData)
        Else
        BuatFolder App.Path & "\Plus Fitur\Startup Manager"
        BuatFolder App.Path & "\Plus Fitur\Startup Manager\Disable-CU"
        BuatFolder App.Path & "\Plus Fitur\Startup Manager\Disable-AU"
            If lvStartup.ListItems.Item(Inter).SubItem(2).Text = GetSpecFolder(ALL_USER_STARTUP) Then
                KillByProccess GetPath(sTarget, FileName)
                Call CopiFile(sTarget, App.Path + "\Plus Fitur\Startup Manager\Disable-AU\" + GetPath(sTarget, FileName), False)
                If HapusFile(sTarget) = False Then MsgBox "Can't Disable !": HapusFile App.Path + "\Plus Fitur\Startup Manager\Disable-AU\" + GetPath(sTarget, FileName): Exit Sub
            Else
                KillByProccess GetPath(sTarget, FileName)
                Call CopiFile(sTarget, App.Path + "\Plus Fitur\Startup Manager\Disable-CU\" + GetPath(sTarget, FileName), False)
                If HapusFile(sTarget) = False Then MsgBox "Can't Disable !": HapusFile App.Path + "\Plus Fitur\Startup Manager\Disable-CU\" + GetPath(sTarget, FileName): Exit Sub
            End If
        End If
        MsgBox "Success !", vbInformation
        GetRegStartup frMain.lvStartup
        Exit For
    End If
Next
End Sub

Private Sub mnES_Click()
Static Inter As Integer
Dim pCada   As String
Dim mKeyReg As Long
Dim pReg    As String
Dim sData   As String
Dim sNama   As String
Dim sTarget As String
Dim sDet    As String

Inter = 1
For Inter = 1 To frMain.lvStartup.ListItems.count
    If frMain.lvStartup.ListItems.Item(Inter).Selected = True Then
    pCada = lfrMain.vStartup.ListItems.Item(Inter).SubItem(2).Text
        mKeyReg = SingkatanKey(Left$(pCada, 4))
        pReg = Right$(pCada, Len(pCada) - 5)
        sData = frMain.lvStartup.ListItems.Item(Inter).SubItem(3).Text
        sNama = frMain.lvStartup.ListItems.Item(Inter).Text
        sTarget = frMain.lvStartup.ListItems.Item(Inter).SubItem(2).Text
        sDet = frMain.lvStartup.ListItems.Item(Inter).SubItem(4).Text
        If sData <> "None" Then
        Call DeleteValue(mKeyReg, pReg, sNama)
        Call SetStringValue(mKeyReg, "Software\Microsoft\Windows\CurrentVersion\Run", sNama, sData)
        Else
            If sTarget = GetSpecFolder(ALL_USER_STARTUP) Then
                KillByProccess GetPath(sDet, FileName)
                Call CopiFile(sDet, sTarget + "\" + GetPath(sDet, FileName), False)
                If HapusFile(sDet) = False Then MsgBox "Can't Enable !" + Chr(13) + "File is in Use !": HapusFile sTarget + "\" + GetPath(sDet, FileName): Exit Sub
            Else
                KillByProccess GetPath(sDet, FileName)
                Call CopiFile(sDet, sTarget + "\" + GetPath(sDet, FileName), False)
                If HapusFile(sDet) = False Then MsgBox "Can't Enable !" + Chr(13) + "File is in Use !": HapusFile sTarget + "\" + GetPath(sDet, FileName): Exit Sub
            End If
        End If
        MsgBox "Success !", vbInformation
        GetRegStartup frMain.lvStartup
        Exit For
    End If
Next
End Sub

Private Sub mnEx_Click()
Static Inter As Integer
Static pPath As String

For Inter = 1 To frMain.lvStartup.ListItems.count
    If frMain.lvStartup.ListItems.Item(Inter).Selected = True Then
    pPath = frMain.lvStartup.ListItems.Item(Inter).SubItem(4).Text
        If ValidFile(pPath) = True Then
            ExploreTheFile pPath
        Else
            MsgBox "Can't find file !", vbExclamation + vbOKOnly, "Warning !"
        End If
    End If
Next
End Sub

Private Sub mnP_Click()
Static Inter As Integer
Static pPath As String

For Inter = 1 To frMain.lvStartup.ListItems.count
    If frMain.lvStartup.ListItems.Item(Inter).Selected = True Then
    pPath = frMain.lvStartup.ListItems.Item(Inter).SubItem(4).Text
        If ValidFile(pPath) = True Then
            ShowProperties pPath, frMain.hWnd
        Else
            MsgBox "Can't find file !", vbExclamation + vbOKOnly, "Warning !"
        End If
    End If
Next
End Sub

Private Sub mnRef_Click()
GetRegStartup frMain.lvStartup
End Sub

