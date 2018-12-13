VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_Arc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'This class file can be used to show the contents of an ARC-archive
        
Private Type ARCHeaderType
    Id As Byte              'Arc signature
    Method As Byte          'Compression method
    FileName As String      'FileName
    CSize As Long           'Compressed filesize
    FDate As Integer        'Date
    FTime As Integer        'Time
    CRC16 As Integer        'CRC 16
    USize As Long           'Original filesize
    DataOffSet As Long      'Offset of the compressed data
End Type

Private ARCFiles() As ARCHeaderType
Private Const m_Unpack_Supported As Boolean = False

Public Function Get_Contents(ZipName As String) As Integer
    Dim FileNum As Long
    Dim ByteVal As Byte
    Dim FName As String * 13
    PackFileName = ZipName
    PackComments = ""
    PackFileType = 0
    FileNum = FreeFile
    Open PackFileName For Binary Access Read As #FileNum
    PackTotFiles = 0
    Erase ARCFiles
    Do
        Get #FileNum, , ByteVal
        If ByteVal <> ARCHeader Then Exit Do            'No arc file
        Get #FileNum, , ByteVal
        If ByteVal < 1 Or ByteVal > 9 Then Exit Do  'probably No arc file or EOF
        PackTotFiles = PackTotFiles + 1
        ReDim Preserve ARCFiles(PackTotFiles)
        With ARCFiles(PackTotFiles)
            .Id = ARCHeader
            .Method = ByteVal
            Get #FileNum, , FName
            .FileName = Trim(Replace(FName, vbNullChar, ""))
            Get #FileNum, , .CSize
            Get #FileNum, , .FDate
            Get #FileNum, , .FTime
            Get #FileNum, , .CRC16
            Get #FileNum, , .USize
            .DataOffSet = Seek(FileNum)
            Seek #FileNum, .DataOffSet + .CSize
        End With
    Loop
    ReDim Preserve ARCFiles(PackTotFiles)
    If PackTotFiles > 0 Then PackFileType = ARCFileType
    Close FileNum
End Function

'Unzip as file and return 0 for good decompression or others for error
Public Function UnPack(ZippedFile() As Boolean, ToPath As String) As Integer

    Erase PackData
End Function

Public Function Pack(ZipName As String, Files() As String, CompType As Integer, CompLevel As Integer, Optional IncludeDir As String = "") As Integer
    
End Function

Public Property Get CanUnpack() As Boolean
    CanUnpack = m_Unpack_Supported
End Property

Public Property Get FileName(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    FileName = ARCFiles(FileNum).FileName
End Property

Public Property Get CommentsFile(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    CommentsFile = "Not Supported"
End Property

Public Property Get CommentsPack() As String
    CommentsPack = ""
End Property

Public Property Get IsDir(FileNum As Long) As Boolean
    If NotGood(FileNum) Then Exit Property
    If ARCFiles(FileNum).USize = 0 Then
        If Right(ARCFiles(FileNum).FileName, 1) = "/" Then IsDir = True
    End If
End Property

Public Property Get Method(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    Method = Methods(ARCFiles(FileNum).Method)
End Property

'Not totaly correct but what the hack
Public Property Get CRC32(FileNum As Long) As Long
    If NotGood(FileNum) Then Exit Property
    CRC32 = ARCFiles(FileNum).CRC16
End Property

Public Property Get Compressed_Size(FileNum As Long) As Long
    If NotGood(FileNum) Then Exit Property
    Compressed_Size = ARCFiles(FileNum).CSize
End Property

Public Property Get UnCompressed_Size(FileNum As Long) As Long
    If NotGood(FileNum) Then Exit Property
    UnCompressed_Size = ARCFiles(FileNum).USize
End Property

Public Property Get Encrypted(FileNum As Long) As Boolean
    If NotGood(FileNum) Then Exit Property
    Encrypted = False
End Property

Public Property Get FileDateTime(FileNum As Long) As Date
    If NotGood(FileNum) Then Exit Property
    FileDateTime = GetZipDate(ARCFiles(FileNum).FDate, ARCFiles(FileNum).FTime)
End Property

Public Property Get SystemMadeBy(FileNum As Long) As String
    SystemMadeBy = "UnKnown"
End Property

Public Property Get VersionMadeBy(FileNum As Long) As String
    VersionMadeBy = "UnKnown"
End Property

Public Property Get SystemNeeded(FileNum As Long) As String
    SystemNeeded = "UnKnown"
End Property

Public Property Get VersionNeeded(FileNum As Long) As String
    VersionNeeded = "UnKnown"
End Property

Private Function Methods(MethodType As Byte) As String
    Select Case MethodType
        Case 1: Methods = "unpacked (obsolete)"
        Case 2: Methods = "unpacked"
        Case 3: Methods = "packed"
        Case 4: Methods = "squeezed (after packing)"
        Case 5: Methods = "crunched (obsolete)"
        Case 6: Methods = "crunched (after packing) (obsolete)"
        Case 7: Methods = "crunched (after packing, using faster hash algorithm)"
        Case 8: Methods = "crunched (after packing, using dynamic LZW variations)"
        Case 9: Methods = "Squashed c/o Phil Katz (no packing) (var. on crunching)"
        Case Else: Methods = "Unknown"
    End Select
End Function

Private Function NotGood(FileNum As Long) As Boolean
    If FileNum = 0 Then NotGood = True: Exit Function
    If FileNum > PackTotFiles Then NotGood = True: Exit Function
    If PackFileType = 0 Then NotGood = True: Exit Function
End Function


