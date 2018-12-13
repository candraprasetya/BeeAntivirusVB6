VERSION 5.00
Begin VB.UserControl ucDirTree 
   ClientHeight    =   4035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5460
   ScaleHeight     =   4035
   ScaleWidth      =   5460
   Begin VB.PictureBox tree1 
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3195
      ScaleWidth      =   3315
      TabIndex        =   1
      Top             =   0
      Width           =   3375
   End
   Begin VB.PictureBox picBuffer 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   3480
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "ucDirTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpFileName As String) As Long
Private Const MAXDWORD = &HFFFF, FILE_ATTRIBUTE_ARCHIVE = &H20, FILE_ATTRIBUTE_DIRECTORY = &H10, FILE_ATTRIBUTE_HIDDEN = &H2, FILE_ATTRIBUTE_NORMAL = &H80, FILE_ATTRIBUTE_READONLY = &H1, FILE_ATTRIBUTE_SYSTEM = &H4, FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const ALTERNATE As Long = 14
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const MAX_PATH  As Long = 260
'***
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFileW Lib "kernel32" _
(ByVal lpFileName As Long, ByVal lpFindFileData As Long) As Long
Private Declare Function FindNextFileW Lib "kernel32" _
(ByVal hFindFile As Long, ByVal lpFindFileData As Long) As Long
'***
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
Dim cCtlFld          As New gComCtl
Dim dirIcon(4)       As Long
Public SelesaiLoad   As Boolean
Public Event ItemChecked(FullPath As String, Checked As Boolean)
Private Sub PeriksaNode(nodeku As cNode, lst As ListBox)
Dim i                As Integer, node1 As cNode
    If nodeku.ChildCount < 1 Then Exit Sub
    Set node1 = nodeku.GetNode(tvwGetNodeFirstChild)
    For i = 1 To nodeku.ChildCount
        If i > 1 Then Set node1 = node1.GetNode(tvwGetNodeNextSibling)
        If node1.Bold = True Then lst.AddItem node1.FullPath
        If node1.ChildCount > 0 Then PeriksaNode node1, lst
    Next
End Sub
Public Sub CekLokasi(ListApa As ListBox)
    PeriksaCek ListApa
End Sub
Private Sub PeriksaCek(lst As ListBox)
    Dim i      As Integer, node1 As cNode, nodeku As cNode
    Set nodeku = tree1.Root: lst.Clear
    If nodeku.ChildCount < 1 Then Exit Sub
    Set node1 = nodeku.GetNode(tvwGetNodeFirstChild)
    For i = 1 To nodeku.ChildCount
        If i > 1 Then Set node1 = node1.GetNode(tvwGetNodeNextSibling)
        If node1.ChildCount > 0 Then
            PeriksaNode node1, lst
        End If
        If node1.Bold = True Then lst.AddItem node1.FullPath
    Next
End Sub
Private Function AddFolder(TreeView As ucTreeView, Text As String, IsChild As Boolean, path As String, picTemp As PictureBox, imglist As Integer, Optional nodeku As cNode, Optional sKey As String) As cNode
    With picTemp
        .Cls: .AutoRedraw = True
        RetrieveIcon path, picTemp, ricnSmall
    End With
    TreeView.ImageList.AddFromDc picTemp.hdc, 16, 16
    If IsChild Then
        Set AddFolder = nodeku.AddChildNode(, sKey, Text, dirIcon(imglist))
    Else
        Set AddFolder = TreeView.Nodes.Add(, , sKey, Text, dirIcon(imglist))
    End If
    dirIcon(imglist) = CLng(dirIcon(imglist)) + 1
End Function

Private Sub tree1_Expand(ByVal oNode As cNode)
    On Error Resume Next
    Dim i      As Integer, node1 As cNode
    If oNode.ChildCount < 1 Or oNode.Bold = True Then Exit Sub
    Set node1 = oNode.GetNode(tvwGetNodeFirstChild)
    For i = 1 To oNode.ChildCount
        If i > 1 Then Set node1 = node1.GetNode(tvwGetNodeNextSibling)
        LoadFolder node1
    Next
End Sub
Private Sub tree1_NodeClick(ByVal oNode As cNode, ByVal iHitTestCode As eTreeViewHitTest)
    If oNode.Bold = False Then
        LoadFolder oNode
    End If
End Sub
Private Sub UserControl_Initialize()
    Set tree1.ImageList = cCtlFld.NewImageList(16, 16, imlColor32)
    tree1.ItemHeight = 19
    SelesaiLoad = False
End Sub
Private Sub UserControl_Resize()
    tree1.Width = UserControl.Width
    tree1.Height = UserControl.Height
End Sub
Public Sub LoadTreeDir()
    Dim DriveNum      As Integer, node1 As Object, DriveType As Long
    DriveNum = 64: tree1.Nodes.Clear:  SelesaiLoad = False ': lstPath.Clear: lstPath.ListIndex = 0
    Do
        DriveNum = DriveNum + 1
        DriveType = GetDriveType(Chr$(DriveNum) & ":\")
        If DriveNum > 90 Then Exit Do
        If DriveNum = 65 Then GoTo lanjutkan
        If InStr("023456", CStr(DriveType)) <> 0 Then
            On Error Resume Next
            Set node1 = AddFolder(tree1, DriveLabel(Chr$(DriveNum)) & " (" & Chr$(DriveNum) & ":)", _
            False, Chr$(DriveNum) & ":\", picBuffer, 0, , Chr$(DriveNum) & ":\")
            'lstPath.AddItem Chr$(DriveNum) & ":\"
            LoadFolder node1
        End If
lanjutkan:
    Loop
End Sub
Public Sub getAktifDrive(drv() As String)
    Dim DriveNum      As Integer, DriveType As Long, sKumpul As String
    DriveNum = 64
    Do
        DriveNum = DriveNum + 1
        DriveType = GetDriveType(Chr$(DriveNum) & ":\")
        If DriveNum > 90 Then Exit Do
        If DriveType = 0 Or DriveType = 2 Or DriveType = 3 Or DriveType = 4 Or DriveType = 5 Or DriveType = 6 Then
            sKumpul = sKumpul & "|" & Chr$(DriveNum)
        End If
    Loop
    'array mulai dari 1 bukan 0
    drv = Split(sKumpul, "|")
End Sub
Private Sub tree1_NodeCheck(ByVal oNode As cNode)
    On Error Resume Next
    If oNode.Bold = False Then
        oNode.Bold = True
        oNode.Expanded = False
        oNode.DeleteChildren
        oNode.ShowPlusMinus = False
    Else
        oNode.Bold = False
        LoadFolder oNode
        oNode.Expanded = False
    End If
    RaiseEvent ItemChecked(oNode.FullPath, oNode.Bold)
End Sub
Private Sub tree1_NodeDblClick(ByVal oNode As cNode, ByVal iHitTestCode As eTreeViewHitTest)
    On Error Resume Next
    If iHitTestCode = 64 Then
        oNode.Bold = Not oNode.Bold
        RaiseEvent ItemChecked(oNode.FullPath, oNode.Bold)
    End If
End Sub
Private Function RemoveNulls(OriginalString As String) As String
    Dim pos As Long
    pos = InStr(OriginalString, Chr$(0))
    If pos > 1 Then
        RemoveNulls = Mid$(OriginalString, 1, pos - 1)
    Else
        RemoveNulls = OriginalString
    End If
End Function
Private Function LoadFolder(ByVal nodeku As cNode)
    Dim FileName      As String, hSearch As Long
    Dim WFD           As WIN32_FIND_DATA, Cont As Integer, path As String
    path = nodeku.Key 'lstPath.List(CLng(nodeku.Key))
    If nodeku.ChildCount > 0 Then Exit Function
    If Right(path, 1) <> "\" Then path = path & "\"
    Cont = True
    hSearch = FindFirstFileW(StrPtr(path & "*"), VarPtr(WFD))
    If hSearch <> INVALID_HANDLE_VALUE Then
        While Cont
            FileName = RemoveNulls((WFD.cFileName)): DoEvents
            If SelesaiLoad Then Exit Function
            If (FileName <> ".") And (FileName <> "..") Then
                If WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
                    DoEvents
                    AddFolder tree1, FileName, True, path & FileName, picBuffer, 0, nodeku, path & FileName
                    'lstPath.AddItem path & FileName
                End If
            End If
            Cont = FindNextFileW(hSearch, VarPtr(WFD)) ' Get next file
        Wend
        Cont = FindClose(hSearch)
        If nodeku.ChildCount > 0 Then nodeku.ShowPlusMinus = True
    End If
End Function
Private Sub UserControl_Terminate()
    Set cCtlFld = Nothing
    SelesaiLoad = True
End Sub

Public Function GetFromUnicode(Index As Long) As String
    GetFromUnicode = UniList1.List(Index)
End Function

Private Function DriveLabel(ByVal sDrive As String) As String
    Const clMaxLen As Long = 100
    Dim lSerial      As Long, sDriveName As String * clMaxLen, sFileSystemName As String * clMaxLen
    sDrive = Left$(sDrive, 1) & ":\"
    'dapatkan info drive
    If GetVolumeInformation(sDrive, sDriveName, clMaxLen, lSerial, 0, 0, sFileSystemName, clMaxLen) Then
        DriveLabel = Left$(sDriveName, InStr(1, sDriveName, vbNullChar) - 1)
    Else
        DriveLabel = ""
    End If
    If DriveLabel <> "" Then Exit Function
    Select Case GetDriveType(sDrive)
    Case 3: DriveLabel = "Local Disk"
    Case 5: DriveLabel = "CD/DVD-Drive"
    End Select
End Function


