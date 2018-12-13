VERSION 5.00
Begin VB.Form frStartup 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dasanggra Start-Up Manager"
   ClientHeight    =   3975
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   12255
   Icon            =   "frStartup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   12255
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
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
Attribute VB_Name = "frStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
With frStartup
    With .lvStartup
        .View = lvwDetails
        .Font.FaceName = "Arial"
        .Columns.Add , "Startup Name", , , lvwAlignCenter, 1800
        .Columns.Add , "Startup Path", , , lvwAlignLeft, 5000
        .Columns.Add , "Startup Reg Data", , , lvwAlignLeft, 5000
        .Columns.Add , "File Path", , , lvwAlignLeft, 5000
        .Columns.Add , "Evaluation", , , lvwAlignLeft, 1600
        .Columns.Add , "Status", , , lvwAlignLeft, 1300
    End With
End With
GetRegStartup lvStartup
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
        Me.mnES.Enabled = False
        Me.mnDS.Enabled = True
        PopupMenu mnST
        Else
        Me.mnES.Enabled = True
        Me.mnDS.Enabled = False
        PopupMenu mnST
        End If
    End If
Next
End Sub

Private Sub mnDelS_Click()
Static Inter As Integer
Dim pCada    As String
Dim mKeyReg  As Long
Dim pReg     As String
Dim sNama    As String
Dim sData    As String

Inter = 1
For Inter = 1 To lvStartup.ListItems.count
    If lvStartup.ListItems.Item(Inter).Selected = True Then
    pCada = lvStartup.ListItems.Item(Inter).SubItem(2).Text
        mKeyReg = SingkatanKey(Left$(pCada, 4))
        pReg = Right$(pCada, Len(pCada) - 5)
        sNama = lvStartup.ListItems.Item(Inter).Text
        sData = lvStartup.ListItems.Item(Inter).SubItem(3).Text
        If sData <> "None" Then
        Call DeleteValue(mKeyReg, pReg, sNama)
        Else
        KillByProccess GetPath(lvStartup.ListItems.Item(Inter).SubItem(4).Text, fileName)
        If HapusFile(lvStartup.ListItems.Item(Inter).SubItem(4).Text) = False Then MsgBox "Can't delete Startup !" + Chr(13) + "File is in use !"
        End If
        MsgBox "Success !", vbInformation, "Dasanggra"
        GetRegStartup lvStartup
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
For Inter = 1 To lvStartup.ListItems.count
    If lvStartup.ListItems.Item(Inter).Selected = True Then
    pCada = lvStartup.ListItems.Item(Inter).SubItem(2).Text
        mKeyReg = SingkatanKey(Left$(pCada, 4))
        sData = lvStartup.ListItems.Item(Inter).SubItem(3).Text
        sTarget = lvStartup.ListItems.Item(Inter).SubItem(4).Text
        pReg = Right$(pCada, Len(pCada) - 5)
        sNama = lvStartup.ListItems.Item(Inter).Text
        If sData <> "None" Then
        Call CreateKeyReg(mKeyReg, "Software\Microsoft\Windows\CurrentVersion\Run-")
        Call DeleteValue(mKeyReg, pReg, sNama)
        Call SetStringValue(mKeyReg, "Software\Microsoft\Windows\CurrentVersion\Run-", sNama, sData)
        Else
        BuatFolder App.Path & "\Plus Fitur\Startup Manager"
        BuatFolder App.Path & "\Plus Fitur\Startup Manager\Disable-CU"
        BuatFolder App.Path & "\Plus Fitur\Startup Manager\Disable-AU"
            If lvStartup.ListItems.Item(Inter).SubItem(2).Text = GetSpecFolder(ALL_USER_STARTUP) Then
                KillByProccess GetPath(sTarget, fileName)
                Call CopiFile(sTarget, App.Path + "\Plus Fitur\Startup Manager\Disable-AU\" + GetPath(sTarget, fileName), False)
                If HapusFile(sTarget) = False Then MsgBox "Can't Disable !": HapusFile App.Path + "\Plus Fitur\Startup Manager\Disable-AU\" + GetPath(sTarget, fileName): Exit Sub
            Else
                KillByProccess GetPath(sTarget, fileName)
                Call CopiFile(sTarget, App.Path + "\Plus Fitur\Startup Manager\Disable-CU\" + GetPath(sTarget, fileName), False)
                If HapusFile(sTarget) = False Then MsgBox "Can't Disable !": HapusFile App.Path + "\Plus Fitur\Startup Manager\Disable-CU\" + GetPath(sTarget, fileName): Exit Sub
            End If
        End If
        MsgBox "Success !", vbInformation, "Dasanggra"
        GetRegStartup lvStartup
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
For Inter = 1 To lvStartup.ListItems.count
    If lvStartup.ListItems.Item(Inter).Selected = True Then
    pCada = lvStartup.ListItems.Item(Inter).SubItem(2).Text
        mKeyReg = SingkatanKey(Left$(pCada, 4))
        pReg = Right$(pCada, Len(pCada) - 5)
        sData = lvStartup.ListItems.Item(Inter).SubItem(3).Text
        sNama = lvStartup.ListItems.Item(Inter).Text
        sTarget = lvStartup.ListItems.Item(Inter).SubItem(2).Text
        sDet = lvStartup.ListItems.Item(Inter).SubItem(4).Text
        If sData <> "None" Then
        Call DeleteValue(mKeyReg, pReg, sNama)
        Call SetStringValue(mKeyReg, "Software\Microsoft\Windows\CurrentVersion\Run", sNama, sData)
        Else
            If sTarget = GetSpecFolder(ALL_USER_STARTUP) Then
                KillByProccess GetPath(sDet, fileName)
                Call CopiFile(sDet, sTarget + "\" + GetPath(sDet, fileName), False)
                If HapusFile(sDet) = False Then MsgBox "Can't Enable !" + Chr(13) + "File is in Use !": HapusFile sTarget + "\" + GetPath(sDet, fileName): Exit Sub
            Else
                KillByProccess GetPath(sDet, fileName)
                Call CopiFile(sDet, sTarget + "\" + GetPath(sDet, fileName), False)
                If HapusFile(sDet) = False Then MsgBox "Can't Enable !" + Chr(13) + "File is in Use !": HapusFile sTarget + "\" + GetPath(sDet, fileName): Exit Sub
            End If
        End If
        MsgBox "Success !", vbInformation, "Dasanggra"
        GetRegStartup lvStartup
        Exit For
    End If
Next
End Sub

Private Sub mnEx_Click()
Static Inter As Integer
Static pPath As String

For Inter = 1 To lvStartup.ListItems.count
    If lvStartup.ListItems.Item(Inter).Selected = True Then
    pPath = lvStartup.ListItems.Item(Inter).SubItem(4).Text
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

For Inter = 1 To lvStartup.ListItems.count
    If lvStartup.ListItems.Item(Inter).Selected = True Then
    pPath = lvStartup.ListItems.Item(Inter).SubItem(4).Text
        If ValidFile(pPath) = True Then
            ShowProperties pPath, Me.hWnd
        Else
            MsgBox "Can't find file !", vbExclamation + vbOKOnly, "Warning !"
        End If
    End If
Next
End Sub

Private Sub mnRef_Click()
GetRegStartup lvStartup
End Sub
