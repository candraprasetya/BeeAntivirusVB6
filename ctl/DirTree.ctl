VERSION 5.00
Begin VB.UserControl DirTree 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Picture         =   "DirTree.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox picBuffer 
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
      Index           =   10
      Left            =   3960
      Picture         =   "DirTree.ctx":058A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   11
      Top             =   2640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picBuffer 
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
      Index           =   7
      Left            =   3240
      Picture         =   "DirTree.ctx":0B14
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   10
      Top             =   2160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picBuffer 
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
      Index           =   0
      Left            =   3000
      Picture         =   "DirTree.ctx":109E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picBuffer 
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
      Index           =   1
      Left            =   3240
      Picture         =   "DirTree.ctx":1628
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picBuffer 
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
      Index           =   2
      Left            =   3720
      Picture         =   "DirTree.ctx":1BB2
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picBuffer 
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
      Index           =   3
      Left            =   3600
      Picture         =   "DirTree.ctx":213C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picBuffer 
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
      Index           =   4
      Left            =   4080
      Picture         =   "DirTree.ctx":26C6
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picBuffer 
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
      Index           =   5
      Left            =   4200
      Picture         =   "DirTree.ctx":2C50
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picBuffer 
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
      Index           =   6
      Left            =   3000
      Picture         =   "DirTree.ctx":31DA
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picBuffer 
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
      Index           =   8
      Left            =   3480
      Picture         =   "DirTree.ctx":3764
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picBuffer 
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
      Index           =   9
      Left            =   3720
      Picture         =   "DirTree.ctx":3CEE
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   240
   End
   Begin prjDAA.ucTreeView tree1 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   5741
   End
End
Attribute VB_Name = "DirTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'###:Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetVolumeInformationW Lib "kernel32" (ByVal pv_lpRootPathName As Long, ByVal pv_lpVolumeNameBuffer As Long, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal pv_lpFileSystemNameBuffer As Long, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryW" (ByVal pszPath As String) As Long
    Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
    Private Declare Function FindFirstFileW Lib "kernel32" _
        (ByVal lpFileName As Long, ByVal lpFindFileData As Long) As Long
    Private Declare Function FindNextFileW Lib "kernel32" _
        (ByVal hFindFile As Long, ByVal lpFindFileData As Long) As Long
    Private Const MAX_PATH As Long = 260, ALTERNATE As Long = 14, MAXDWORD = &HFFFF, INVALID_HANDLE_VALUE = -1, FILE_ATTRIBUTE_ARCHIVE = &H20, FILE_ATTRIBUTE_DIRECTORY = &H10, FILE_ATTRIBUTE_HIDDEN = &H2, FILE_ATTRIBUTE_NORMAL = &H80, FILE_ATTRIBUTE_READONLY = &H1, FILE_ATTRIBUTE_SYSTEM = &H4, FILE_ATTRIBUTE_TEMPORARY = &H100
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
    
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime   As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime  As FILETIME
    nFileSizeHigh    As Long
    nFileSizeLow     As Long
    dwReserved0      As Long
    dwReserved1      As Long
    cFileName        As String * MAX_PATH
    cAlternate       As String * ALTERNATE
End Type

Public SelesaiLoad  As Boolean

Dim dirIcon          As Long

Public Sub OutPutPath(lst As Collection)
    If tree1.NodeChecked(tree1.NodeRoot) = False Then Exit Sub
    Call PeriksaCek(tree1.NodeRoot, lst)
End Sub
Private Sub PeriksaCek(hNode As Long, lst As Collection)
    Dim anak As Long, i As Integer
    With tree1
    anak = .NodeChild(hNode)
    For i = 1 To .NodeChildren(hNode)
        If i > 1 Then anak = .NodeNextSibling(anak)
        If .IsChildCheckedAll(anak) = True Then lst.Add .GetNodeKey(anak): GoTo Lanjut
        If .NodeChecked(anak) = True Then PeriksaCek anak, lst
Lanjut:
    Next
    End With
End Sub


Private Sub tree1_BeforeExpand(ByVal hNode As Long, ByVal ExpandedOnce As Boolean)
    On Error Resume Next
    Dim i      As Integer, nodeku As Long
    If Left$(tree1.GetNodeKey(hNode), 2) = "::" Then Exit Sub
    With tree1
    nodeku = .NodeChild(hNode)
    For i = 1 To .NodeChildren(hNode)
        If i > 1 Then nodeku = .NodeNextSibling(nodeku)
        If tree1.NodeChildren(nodeku) <> 0 Then Exit Sub
        LoadFolder nodeku
    Next
    End With
End Sub
Private Sub tree1_NodeCheck(ByVal hNode As Long)
Dim sKeyNode As String
    If tree1.NodeChecked(hNode) = False Then
       sKeyNode = tree1.GetNodeKey(hNode)
       If sKeyNode = "::ssta" Then
          StartUpNode = False
       ElseIf sKeyNode = "::sreg" Then
          RegNode = False
       ElseIf sKeyNode = "::spro" Then
          ProsesNode = False
       ElseIf sKeyNode = "::pwin" Then
          WinNode = False
       ElseIf sKeyNode = "::pdoc" Then
          DocNode = False
       ElseIf sKeyNode = "::ppro" Then
          ProgNode = False
       ElseIf sKeyNode = "::plain" Then
          WinNode = False
          DocNode = False
          ProgNode = False
       ElseIf sKeyNode = "::sistem" Then
          StartUpNode = False
          RegNode = False
          ProsesNode = False
       End If
       tree1.NodeChecked(hNode) = False
    Else
       sKeyNode = tree1.GetNodeKey(hNode)
       If sKeyNode = "::ssta" Then
          StartUpNode = True
       ElseIf sKeyNode = "::sreg" Then
          RegNode = True
       ElseIf sKeyNode = "::spro" Then
          ProsesNode = True
       ElseIf sKeyNode = "::pwin" Then
          WinNode = True
       ElseIf sKeyNode = "::pdoc" Then
          DocNode = True
       ElseIf sKeyNode = "::ppro" Then
          ProgNode = True
       ElseIf sKeyNode = "::plain" Then
          WinNode = True
          DocNode = True
          ProgNode = True
       ElseIf sKeyNode = "::sistem" Then
          StartUpNode = True
          RegNode = True
          ProsesNode = True
       End If
       tree1.NodeChecked(hNode) = True
    End If
    
    tree1.CheckChildren hNode, tree1.NodeChecked(hNode)
    tree1.CheckParent hNode, tree1.NodeChecked(hNode)
End Sub
Public Sub Matikan()
    tree1.Matikan_Tree
End Sub
Private Sub UserControl_Initialize()
    With tree1
        Call .Initialize
        Call .InitializeImageList(16, 16)
        Dim i As Integer
        For i = 0 To 10
            Call .AddIcon(picBuffer(i).Picture)
        Next
        .ItemHeight = 19
        .HasButtons = True
        .HasLines = True
        .HasRootLines = True
        .CheckBoxes = True
        .TrackSelect = True
        .ItemIndent = 22
    End With
    dirIcon = 0
    SelesaiLoad = False
End Sub
Private Sub UserControl_Resize()
    tree1.Width = UserControl.Width
    tree1.Height = UserControl.Height
End Sub
Public Sub LoadTreeDir(Optional cekSis As Boolean = True, Optional cekKomp As Boolean = False)
    Dim DriveNum      As Integer, node1 As Long, DriveType As Long, ikonku As Integer, nodeKomp As Long, nodeSis As Long
    Dim NodePathLain  As Long
    DriveNum = 64: tree1.Clear: SelesaiLoad = False
    nodeKomp = AddFolder(tree1, "My Computer", False, , "::komputer", 5, 5)
    'list untuk sistem
    nodeSis = AddFolder(tree1, j_bahasa(10), False, , "::sistem", 6, 6)
    AddFolder tree1, j_bahasa(11), True, nodeSis, "::spro", 7, 7
    AddFolder tree1, b_bahasa(2), True, nodeSis, "::sreg", 8, 8
    AddFolder tree1, e_bahasa(9), True, nodeSis, "::ssta", 9, 9
    
    ' Tambahkan Path Lain
    NodePathLain = AddFolder(tree1, j_bahasa(35), False, , "::plain", 0, 0)
    AddFolder tree1, "Windows", True, NodePathLain, "::pwin", 0, 0
    AddFolder tree1, "Program Files", True, NodePathLain, "::ppro", 0, 0
    AddFolder tree1, "My Documents", True, NodePathLain, "::pdoc", 10, 10
    
    
    Do
        DriveNum = DriveNum + 1
        DriveType = GetDriveType(Chr$(DriveNum) & ":\")
        If DriveNum > 90 Then Exit Do
        If DriveNum = 65 Then GoTo lanjutkan
        Select Case DriveType
            Case 3: ikonku = 2
            Case 2: ikonku = 4
            Case 5: ikonku = 3
            Case Else: GoTo lanjutkan
        End Select
        On Error Resume Next
        node1 = AddFolder(tree1, DriveLabel(Chr$(DriveNum)) & " (" & Chr$(DriveNum) & ":)", True, nodeKomp, Chr$(DriveNum) & ":\", ikonku, ikonku)
        LoadFolder node1
lanjutkan:
    Loop
    tree1.Expand nodeKomp
    tree1.Expand nodeSis
    If cekSis = True Then tree1.NodeChecked(nodeSis) = True: tree1.CheckChildren nodeSis, True
    If cekKomp = True Then tree1.NodeChecked(nodeKomp) = True: tree1.CheckChildren nodeKomp, True
End Sub
Private Function AddFolder(TreeView As ucTreeView, Text As String, Optional IsChild As Boolean = True, Optional atas As Long, Optional sKey As String, Optional IkonApa As Integer = -1, Optional IkonSel As Integer = -1) As Long
    On Error Resume Next
    If IsChild Then
        AddFolder = TreeView.AddNode(atas, , sKey, Text, IkonApa, IkonSel)
        TreeView.CheckChildren atas, TreeView.NodeChecked(atas)
    Else
        AddFolder = TreeView.AddNode(, , sKey, Text, IkonApa, IkonSel)
    End If
    dirIcon = dirIcon + 1
End Function
Private Function StripNulls(ByVal OriginalStr As String) As String
    If (InStr(OriginalStr, Chr$(0)) > 0) Then
        OriginalStr = Left$(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function
Private Function LoadFolder(nodeku As Long)
    Dim Filename      As String, hSearch As Long
    Dim WFD           As WIN32_FIND_DATA, Cont As Integer
    Path = tree1.GetNodeKey(nodeku)
    'If Left$(path, 2) = "::" Then Exit Function
    If tree1.NodeChildren(nodeku) > 0 Then Exit Function
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    Cont = True
    hSearch = FindFirstFileW(StrPtr(Path & "*"), VarPtr(WFD))
    If hSearch <> INVALID_HANDLE_VALUE Then
        While Cont
            Filename = StripNulls((WFD.cFileName)): DoEvents
            If SelesaiLoad Then Exit Function
            If (Filename <> ".") And (Filename <> "..") Then
                'Hitung = Hitung + (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
                If WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
                    'Hitung jumlah Folder
                    DoEvents
                    Dim hid As Long: hid = AddFolder(tree1, Filename, , nodeku, Path & Filename, 0, 1)
                    If WFD.dwFileAttributes And FILE_ATTRIBUTE_HIDDEN Then tree1.NodeGhosted(hid) = True
                End If
            End If
            Cont = FindNextFileW(hSearch, VarPtr(WFD)) ' Get next file
        Wend
        Cont = FindClose(hSearch)
    End If
    tree1.SetRedrawMode True
End Function

Private Function DriveLabel(ByVal sDrive As String) As String
    '###:Const clMaxLen As Long = 100
    '###:Dim lSerial      As Long, sDriveName As String * clMaxLen, sFileSystemName As String * clMaxLen
    Dim sDriveName          As String
    Dim nDriveNameLen       As Long
        nDriveNameLen = 128
        sDriveName = String$(nDriveNameLen, 0)
    sDrive = Left$(sDrive, 1) & ":\"
    'dapatkan info drive
    '###:If GetVolumeInformation(sDrive, sDriveName, clMaxLen, lSerial, 0, 0, sFileSystemName, clMaxLen) Then
    If GetVolumeInformationW(StrPtr(sDrive), StrPtr(sDriveName), nDriveNameLen, ByVal 0, ByVal 0, ByVal 0, 0, 0) Then
        '###:DriveLabel = Left$(sDriveName, InStr(1, sDriveName, vbNullChar) - 1)
        DriveLabel = Left$(sDriveName, InStr(1, sDriveName, ChrW$(0)) - 1)
    Else
        DriveLabel = vbNullString
    End If
    If Len(DriveLabel) > 0 Then Exit Function
    Select Case GetDriveType(sDrive)
    Case 3: DriveLabel = "Local Disk"
    Case 5: DriveLabel = "CD/DVD-Drive"
    End Select
End Function
Private Sub UserControl_Terminate()
    tree1.Matikan_Tree
    SelesaiLoad = True
End Sub


