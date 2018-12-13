VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_GetFileType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'This class is used to examen the file if it is supported by the
'rest of the class files

Private TAR As New Cls_Tar
Private GZ As New Cls_GZip
Private ZIP As New Cls_Zip
Private Archive As Object

Public Function Get_Contents(ZipName As String) As Integer
    Dim FileNum As Long
    Dim FileLenght As Long
    Dim Header As Integer
    PackFileType = 0
    If ZipName = "" Then
        Get_Contents = -1       'file don't exist
        Exit Function
    End If
    If Right(UCase(ZipName), 4) = ".TAR" Then
        Set Archive = TAR
        Archive.Get_Contents ZipName
    End If
    If PackFileType <> 0 Then Exit Function
    FileNum = FreeFile
    Open ZipName For Binary Access Read As #FileNum
    FileLenght = LOF(FileNum)
    Get #FileNum, , Header
    Close #FileNum
    Select Case Header
    Case ZipHeader
        Set Archive = ZIP
        Archive.Get_Contents ZipName
    Case GZipHeader
        Set Archive = GZ
        Archive.Get_Contents ZipName
    Case Else
        PackFileType = 0
        PackTotFiles = 0
        Get_Contents = -2       'no compressed files
    End Select
    If PackFileType = 0 Then Set Archive = Nothing
End Function

'Unzip as file and return 0 for good decompression or others for error
Public Function UnPack(ZippedFile() As Boolean, ToPath As String) As Integer
    If PackFileType = 0 Then Exit Function
    UnPack = Archive.UnPack(ZippedFile, ToPath)
End Function

Public Property Get CanUnpack() As Boolean
    If PackFileType = 0 Then Exit Property
    CanUnpack = Archive.CanUnpack
End Property

Public Property Get FileCount() As Long
    If PackFileType = 0 Then Exit Property
    FileCount = PackTotFiles
End Property

Public Property Get FileName(FileNum As Long) As String
    If PackFileType = 0 Then Exit Property
    FileName = Archive.FileName(FileNum)
End Property

Public Property Get CommentsFile(FileNum) As String
    If PackFileType = 0 Then Exit Property
    CommentsFile = Archive.CommentsFile(FileNum)
End Property

Public Property Get CommentsPack() As String
    If PackFileType = 0 Then Exit Property
    CommentsPack = Archive.CommentsPack
End Property

Public Property Get IsDir(FileNum As Long) As Boolean
    If PackFileType = 0 Then Exit Property
    IsDir = Archive.IsDir(FileNum)
End Property

Public Property Get Method(FileNum As Long) As String
    If PackFileType = 0 Then Exit Property
    Method = Archive.Method(FileNum)
End Property

Public Property Get CRC32(FileNum As Long) As Long
    If PackFileType = 0 Then Exit Property
    CRC32 = Archive.CRC32(FileNum)
End Property

Public Property Get Compressed_Size(FileNum As Long) As Long
    If PackFileType = 0 Then Exit Property
    Compressed_Size = Archive.Compressed_Size(FileNum)
End Property

Public Property Get UnCompressed_Size(FileNum As Long) As Long
    If PackFileType = 0 Then Exit Property
    UnCompressed_Size = Archive.UnCompressed_Size(FileNum)
End Property

Public Property Get Encrypted(FileNum As Long) As Boolean
    If PackFileType = 0 Then Exit Property
    Encrypted = Archive.Encrypted(FileNum)
End Property

Public Property Get FileDateTime(FileNum As Long) As Date
    If PackFileType = 0 Then Exit Property
    FileDateTime = Archive.FileDateTime(FileNum)
End Property

Public Property Get SystemMadeBy(FileNum As Long) As String
    If PackFileType = 0 Then Exit Property
    SystemMadeBy = Archive.SystemMadeBy(FileNum)
End Property

Public Property Get VersionMadeBy(FileNum As Long) As String
    If PackFileType = 0 Then Exit Property
    VersionMadeBy = Archive.VersionMadeBy(FileNum)
End Property

Public Property Get SystemNeeded(FileNum As Long) As String
    If PackFileType = 0 Then Exit Property
    SystemNeeded = Archive.SystemNeeded(FileNum)
End Property

Public Property Get VersionNeeded(FileNum As Long) As String
    If PackFileType = 0 Then Exit Property
    VersionNeeded = Archive.VersionNeeded(FileNum)
End Property

