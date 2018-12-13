VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Archive Explorer"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   13245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnUnzip 
      Caption         =   "Extract selected ones"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   4200
      Width           =   9015
   End
   Begin VB.FileListBox FileList 
      Height          =   3990
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.DirListBox DirList 
      Height          =   3465
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.DriveListBox DriveList 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   1935
   End
   Begin MSComctlLib.ListView lstInZip 
      Height          =   3615
      Left            =   4080
      TabIndex        =   3
      Top             =   480
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label lblHeadLine 
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   120
      Width           =   9015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ZF As New Cls_GetFileType
Private Filetype(10) As String

Private Sub btnUnzip_Click()
    Dim FileUnzip() As Boolean
    Dim ToDir As String
    Dim Sel As Boolean
    Dim X As Long
    Dim RetVal As Boolean
    With lstInZip
        ReDim FileUnzip(.ListItems.Count)
        For X = 1 To .ListItems.Count
            If .ListItems(X).Selected Then
                Sel = True
                Exit For
            End If
        Next
        For X = 1 To .ListItems.Count
            If .ListItems(X).Selected = Sel Then
                FileUnzip(X) = True
            End If
        Next
    End With
    ToDir = tsGetPathFromUser
    If ToDir = "" Then
        MsgBox "No path to store files"
        Exit Sub
    End If
    MousePointer = vbHourglass
    RetVal = ZF.UnPack(FileUnzip, ToDir)
'    RetVal = ZF.Unzip(FileUnzip, ToDir)
    MousePointer = vbNormal
End Sub
    
Private Sub DirList_Change()
    FileList.Path = DirList.Path
End Sub

Private Sub DriveList_Change()
    DirList.Path = DriveList.Drive
End Sub

Private Sub FileList_Click()
    If FileList.FileName <> "" Then
        lstInZip.ListItems.Clear
        Call Show_ZipContents
        If Len(ZF.CommentsPack) > 0 Then
            MsgBox ZF.CommentsPack
        End If
    End If
End Sub

Private Sub Show_ZipContents()
    Dim X As Long
    Dim Enc As String
    Dim DirCnt As Long
    Dim FileCnt As Long
    Dim Temp As Long
    ZF.Get_Contents (DirList.Path & "\" & FileList.FileName)
    For X = 1 To lstInZip.ListItems.Count
        lstInZip.ListItems(X).Selected = False
    Next
    For X = 1 To ZF.FileCount
        With lstInZip
            Enc = " "
            If ZF.Encrypted(X) Then Enc = "+"
            If Not ZF.IsDir(X) Then
                FileCnt = FileCnt + 1
                .ListItems.Add X, , Enc & ZF.FileName(X)
                .ListItems(X).SubItems(1) = ZF.Method(X)
                Temp = ZF.CRC32(X)
                If Temp = 0 Then
                    .ListItems(X).SubItems(2) = "?"
                Else
                    .ListItems(X).SubItems(2) = Hex(Temp)
                End If
                Temp = ZF.Compressed_Size(X)
                If Temp = 0 Then
                    .ListItems(X).SubItems(3) = "?"
                Else
                    .ListItems(X).SubItems(3) = Temp
                End If
                .ListItems(X).SubItems(4) = ZF.UnCompressed_Size(X)
                .ListItems(X).SubItems(5) = ZF.FileDateTime(X)
            Else
                DirCnt = DirCnt + 1
                .ListItems.Add X, , Enc & ZF.FileName(X)
                .ListItems(X).SubItems(1) = ZF.Method(X)
                .ListItems(X).SubItems(2) = "Directory Entry"
                .ListItems(X).SubItems(3) = "Directory Entry"
                .ListItems(X).SubItems(4) = "Directory Entry"
                .ListItems(X).SubItems(5) = ZF.FileDateTime(X)
            End If
        End With
    Next
    If ZF.FileCount > 0 Then
        lblHeadLine.Caption = "Contents of " & Filetype(PackFileType) & " file " & _
                              FileList.FileName & " -> " & _
                              DirCnt & " directories and " & _
                              FileCnt & " files"
    Else
        lblHeadLine.Caption = "Not supported format"
    End If
    If ZF.CanUnpack Then
        btnUnzip.Enabled = True
    Else
        btnUnzip.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Call Insert_Header
'    FileList.Pattern = "*.zip;*.gz;*.tgz;*.tar;*.arj"
    Filetype(ZipFileType) = "ZIP"
    Filetype(GZFileType) = "GZIP"
    Filetype(TARFileType) = "TAR"
    Filetype(RARFileType) = "RAR"
    Filetype(ARJFileType) = "ARJ"
    Filetype(LZHFileType) = "LZH/LHA"
    Filetype(CABFileType) = "Cabinet"
'    DirList.Path = "d:\download\new\archives"
    btnUnzip.Enabled = False
End Sub

Private Sub Insert_Header()
    With lstInZip
        .ColumnHeaders.Add , , "File Name"
        .ColumnHeaders.Add , , "Compression Method"
        .ColumnHeaders.Add , , "CRC-32"
        .ColumnHeaders.Add , , "Compressed Size"
        .ColumnHeaders.Add , , "Decompressed Size"
        .ColumnHeaders.Add , , "File date"
    End With
End Sub

