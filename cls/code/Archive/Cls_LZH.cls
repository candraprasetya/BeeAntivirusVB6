VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_LZH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'This class file can be used to show the contents of an LZH/LHA-archive

Private Type LzhLhaType
    Headersize As Byte      'Size of archived file header
    HeaderCRC As Byte       'Checksum of remaining bytes
    Minc1 As Byte           '-
    Method As String * 3    'Compression methods used lzs/lh7
    Minc2 As Byte           '-
    CSize As Long           'Compressed size
    USize As Long           'Uncompressed size
    FDateTime As Long       'File Date and time
    FTime As Integer        'File Time
    FDate As Integer        'File Date
    Attrib As Integer       'File attribute
    FLen As Byte            'Filename Lenght
    FileName As String      'FileName
    CRC16 As Integer        'CRC16 of the data
    DataOffSet As Long      'Start Compressed data
End Type

Private LZHFiles() As LzhLhaType
Private Const m_Unpack_Supported As Boolean = False

Public Function Get_Contents(ZipName As String) As Integer
    Dim FileNum As Long
    Dim FileLenght As Long
    Dim ByteVal(2) As Byte
    Dim Meth As String * 3
    Dim Temp As LzhLhaType
    Dim NPos As Long
    PackFileName = ZipName
    PackComments = ""
    PackFileType = 0
    FileNum = FreeFile
    Open PackFileName For Binary Access Read As #FileNum
    FileLenght = LOF(FileNum)
    PackTotFiles = 0
    ReDim LZHFiles(100)
    Do
        Get #FileNum, , ByteVal(0)
        If ByteVal(0) < 24 Then Exit Do         'HeaderSize to small
        NPos = Seek(FileNum)
        Get #FileNum, , ByteVal(1)
        Get #FileNum, , ByteVal(2)
        If ByteVal(2) <> 45 Then Exit Do        'No LZH file
        Get #FileNum, , Meth
        If Left(Meth, 1) <> "l" Then Exit Do    'No LZH file
        
        PackTotFiles = PackTotFiles + 1
        With LZHFiles(PackTotFiles)
            .Headersize = ByteVal(0)
            .HeaderCRC = ByteVal(1)
            .Minc1 = ByteVal(2)
            .Method = Meth
            Get #FileNum, , .Minc2
            Get #FileNum, , .CSize
            Get #FileNum, , .USize
            Get #FileNum, , .FDateTime
            Get #FileNum, , .Attrib
            Get #FileNum, , .FLen
            .FileName = String(.FLen, 0)
            Get #FileNum, , .FileName
            Get #FileNum, , .CRC16
            .DataOffSet = Seek(FileNum)
            Seek #FileNum, NPos + .CSize + 1 + .Headersize
            .FDate = Lng2Int((.FDateTime And &HFFFF0000) \ &HFFFF&)
            .FTime = Lng2Int(.FDateTime And &HFFFF&)
        End With
    Loop
    ReDim Preserve LZHFiles(PackTotFiles)
    If PackTotFiles > 0 Then PackFileType = LZHFileType
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
    FileName = LZHFiles(FileNum).FileName
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
    If LZHFiles(FileNum).Method = "-lhd-" Then IsDir = True: Exit Property
    If LZHFiles(FileNum).USize = 0 Then
        If Right(LZHFiles(FileNum).FileName, 1) = "/" Then IsDir = True
    End If
End Property

Public Property Get Method(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    Method = Methods(LZHFiles(FileNum).Method)
End Property

'Not totaly correct but what the hack
Public Property Get CRC32(FileNum As Long) As Long
    If NotGood(FileNum) Then Exit Property
    CRC32 = LZHFiles(FileNum).CRC16
End Property

Public Property Get Compressed_Size(FileNum As Long) As Long
    If NotGood(FileNum) Then Exit Property
    Compressed_Size = LZHFiles(FileNum).CSize
End Property

Public Property Get UnCompressed_Size(FileNum As Long) As Long
    If NotGood(FileNum) Then Exit Property
    UnCompressed_Size = LZHFiles(FileNum).USize
End Property

Public Property Get Encrypted(FileNum As Long) As Boolean
    If NotGood(FileNum) Then Exit Property
    Encrypted = False
End Property

Public Property Get FileDateTime(FileNum As Long) As Date
    If NotGood(FileNum) Then Exit Property
    FileDateTime = GetZipDate(LZHFiles(FileNum).FDate, LZHFiles(FileNum).FTime)
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

Private Function Methods(MethodType As String) As String
    Select Case MethodType
        Case "lzs": Methods = "LZSS with 2kb window"
        Case "lz4": Methods = "no compression"
        Case "lz5": Methods = "LZSS with 4kb window"
        Case "lhd": Methods = "no compression (this is a directory, not a file!)"
        Case "lh0": Methods = "no compression (could be a directory or a file)"
        Case "lh1": Methods = "LZH with 4kb window, dynamic Huffman"
        Case "lh2": Methods = "LZH with 8kb window, dynamic Huffman"
        Case "lh3": Methods = "LZH with 8kb window, static Huffman"
        Case "lh4": Methods = "LZH with 4kb window, static canonical Huffman"
        Case "lh5": Methods = "LZH with 8kb window, static canonical Huffman"
        Case "lh6": Methods = "LZH with 32kb window, static canonical Huffman"
        Case "lh7": Methods = "LZH with 64kb window, static canonical Huffman"
        Case Else: Methods = "Unknown"
    End Select
End Function

Private Function NotGood(FileNum As Long) As Boolean
    If FileNum = 0 Then NotGood = True: Exit Function
    If FileNum > PackTotFiles Then NotGood = True: Exit Function
    If PackFileType = 0 Then NotGood = True: Exit Function
End Function

Private Function Lng2Int(LngValue As Long) As Integer
    If LngValue > 32767 Then Lng2Int = LngValue - &HFFFF& - 1 Else Lng2Int = LngValue
End Function
