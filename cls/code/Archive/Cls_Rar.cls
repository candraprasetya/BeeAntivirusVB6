VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_Rar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'This class file can be used to show the contents of an RAR-archive

Private Type BlockMarkerType
    HEAD_CRC As Integer         'CRC of the header
    HEAD_TYPE As Byte           'Header type
    HEAD_FLAGS As Integer       'Flags
    HEAD_SIZE As Integer        'Size of the header
End Type
Private Type ArchiveHeaderType
    HEAD_CRC As Integer         'CRC of fields HEAD_TYPE to RESERVED2
    HEAD_TYPE As Byte           '&h73
    HEAD_FLAGS As Integer       '&h01    - Volume attribute (archive volume)
'                                &h02    - Archive comment present
'                                &h04    - Archive lock attribute
'                                &h08    - Solid attribute (solid archive)
'                                &h10 -Unused
'                                &h20    - Authenticity information present
'                                other bits in HEAD_FLAGS are reserved for internal use
    HEAD_SIZE As Integer        'Archive header total size including archive comments
    reserved1 As Integer        'Reserved
    reserved2 As Long           'Reserved
    COMMENT As String           'present if (HEAD_FLAGS and &h02) <> 0

End Type
Private Type FileHeaderType
    HEAD_CRC As Integer         'CRC of fields from HEAD_TYPE to FILEATTR and file name
    HEAD_TYPE As Byte           '&h74
    HEAD_FLAGS As Integer       '&h01 - file continued from previous volume
'                                &h02 - file continued in next volume
'                                &h04 - file encrypted with password
'                                &h08 - file comment present
'                                &h10 - information from previous files is used (solid flag)
'                                       (for RAR 2.0 and later)
'
'                                bits 7 6 5 (for RAR 2.0 and later)
'
'                                     0 0 0    - dictionary size   64 Kb
'                                     0 0 1    - dictionary size  128 Kb
'                                     0 1 0    - dictionary size  256 Kb
'                                     0 1 1    - dictionary size  512 Kb
'                                     1 0 0    - dictionary size 1024 Kb
'                                     1 0 1    - reserved
'                                     1 1 0    - reserved
'                                     1 1 1    - file is directory
'
'                                (HEAD_FLAGS And &h8000) == 1, because full block Size Is Head_Size + PACK_SIZE
    HEAD_SIZE As Long        'File header full size including file name and comments
    CSize As Long               'Compressed file size
    USize As Long               'Unompressed file size
    Host_OS As Byte             'Operating system used for archiving
'                                0 - MS DOS
'                                1 - OS/2
'                                2 - Win32
'                                3 - Unix
    File_CRC As Long            'CRC32 of the file
    FDateTime As Long           'Date and time in standard MS DOS format
    FDate As Integer            'Calculated data
    FTime As Integer            'Calculated time
    VersionNeeded As Byte       'Minimum version need to extract
    Method As Byte              'Compression method
    FNameLen As Integer         'Size of FileName
    Attrib As Long              'File Attributes
    FileName As String          'File Name
    COMMENT As String           'present if (HEAD_FLAGS and &h08) <> 0
    DataOffSet As Long          'Start position of packed data in archive
End Type
Private Type CommentHeaderType
    HEAD_CRC As Integer         'CRC of fields from HEAD_TYPE to COMM_CRC
    HEAD_TYPE As Byte           '&h75
    HEAD_FLAGS As Integer       'Bit flags
    HEAD_SIZE As Integer        'Comment header size + comment size
    USize As Integer            'Uncompressed comment size
    VerNeeded As Byte           'RAR version needed to extract comment
    Method As Byte              'Packing method
    COMM_CRC As Integer         'Comment CRC
    COMMENT As String           'Comment text
End Type
Private Type InfoHeaderType
    HEAD_CRC As Integer         'Block CRC
    HEAD_TYPE As Byte           'Header type: 0x76
    HEAD_FLAGS As Integer       'Bit flags
    HEAD_SIZE As Integer        'Total block size
    INFO As String              'Other data
End Type
Private Type SubblockType
    HEAD_CRC As Integer         'Block CRC
    HEAD_TYPE As Byte           'Header type: 0x77
    HEAD_FLAGS As Integer       'Bit flags
'                                (HEAD_FLAGS & 0x8000) == 1, because full
'                                block Size Is HEAD_SIZE + DATA_SIZE
    HEAD_SIZE As Integer        'Total block size
    DATA_SIZE As Long           'Total data size
    SUB_TYPE As Integer         'Subblock type
    RESERVED As Byte            'Must be 0
    OtherFields As String       'Other fields depending on the subblock type
End Type

Private RARArchHead As ArchiveHeaderType
Private RARFiles() As FileHeaderType
Private Const m_Unpack_Supported As Boolean = False

Public Function Get_Contents(ZipName As String) As Integer
    Dim X As Long
    Dim FileNum As Long
    Dim ByteVal As Byte
    Dim TextBytes() As Byte
    Dim Bpointer As Integer
    Dim Temp As BlockMarkerType
    Dim TempHead As ArchiveHeaderType
    Dim TempComment As CommentHeaderType
    Dim TempInfo As InfoHeaderType
    Dim TempSub As SubblockType
    Dim AddHeadSize As Long
    PackFileName = ZipName
    PackComments = ""
    PackTotFiles = 0
    PackFileType = 0
    FileNum = FreeFile
    Open PackFileName For Binary Access Read As #FileNum
    If LOF(FileNum) < 7 Then
        Close #FileNum
        Exit Function
    End If
    'get the end of central date
    Get #FileNum, , Temp
    If Temp.HEAD_CRC <> &H6152 Then Close FileNum: Exit Function
    If Temp.HEAD_TYPE <> &H72 Then Close FileNum: Exit Function
    If Temp.HEAD_FLAGS <> &H1A21 Then Close FileNum: Exit Function
    If Temp.HEAD_SIZE <> &H7 Then Close FileNum: Exit Function
'Header OK  find out what type of header
    PackFileType = RARFileType
    Do
        If Seek(FileNum) >= LOF(FileNum) Then Exit Do    'EOF
        Get #FileNum, , Temp
        AddHeadSize = 0
'        If Temp.HEAD_FLAGS And &H8000& <> 0 Then
'            Get #FileNum, , AddHeadSize
'        End If
        Select Case Temp.HEAD_TYPE
        Case &H72          'marker block    =   first block of the file
            'BlockMarker already read
        Case &H73          'archive header
            With TempHead
                .HEAD_CRC = Temp.HEAD_CRC
                .HEAD_TYPE = Temp.HEAD_TYPE
                .HEAD_FLAGS = Temp.HEAD_FLAGS
                .HEAD_SIZE = Temp.HEAD_SIZE + AddHeadSize
                Get #FileNum, , .reserved1
                Get #FileNum, , .reserved2
                If (.HEAD_FLAGS And 2) > 0 Then
                    ReDim TextBytes(0 To 100)
                    Bpointer = 0
                    Do
                        Get #FileNum, , ByteVal
                        If ByteVal = 0 Then Exit Do
                        TextBytes(Bpointer) = ByteVal
                        Bpointer = Bpointer + 1
                        If Bpointer > UBound(TextBytes) Then ReDim Preserve TextBytes(Bpointer + 50)
                    Loop
                    ReDim Preserve TextBytes(Bpointer - 1)
                    .COMMENT = StrConv(TextBytes, vbUnicode)
                End If
            End With
            RARArchHead = TempHead
        Case &H74          'file header
            PackTotFiles = PackTotFiles + 1
            ReDim Preserve RARFiles(PackTotFiles)
            With RARFiles(PackTotFiles)
                .HEAD_CRC = Temp.HEAD_CRC
                .HEAD_TYPE = Temp.HEAD_TYPE
                .HEAD_FLAGS = Temp.HEAD_FLAGS
                .HEAD_SIZE = Temp.HEAD_SIZE + AddHeadSize
                Get #FileNum, , .CSize
                Get #FileNum, , .USize
                Get #FileNum, , .Host_OS
                Get #FileNum, , .File_CRC
                Get #FileNum, , .FDateTime
                Get #FileNum, , .VersionNeeded
                Get #FileNum, , .Method
                Get #FileNum, , .FNameLen
                Get #FileNum, , .Attrib
                ReDim TextBytes(0 To Int2Lng(.FNameLen) - 1)
                Get #FileNum, , TextBytes
                .FileName = StrConv(TextBytes, vbUnicode)
                If (.HEAD_FLAGS And 2) > 0 Then
                    ReDim TextBytes(0 To 100)
                    Bpointer = 0
                    Do
                        Get #FileNum, , ByteVal
                        If ByteVal = 0 Then Exit Do
                        TextBytes(Bpointer) = ByteVal
                        Bpointer = Bpointer + 1
                        If Bpointer > UBound(TextBytes) Then ReDim Preserve TextBytes(Bpointer + 50)
                    Loop
                    ReDim Preserve TextBytes(Bpointer - 1)
                    .COMMENT = StrConv(TextBytes, vbUnicode)
                End If
                .DataOffSet = Seek(FileNum)
                .FDate = Lng2Int((.FDateTime And &HFFFF0000) \ &HFFFF&)
                .FTime = Lng2Int(.FDateTime And &HFFFF&)
                Seek #FileNum, Seek(FileNum) + .CSize
            End With
        Case &H75          'comment header
            With TempComment
                .HEAD_CRC = Temp.HEAD_CRC
                .HEAD_TYPE = Temp.HEAD_TYPE
                .HEAD_FLAGS = Temp.HEAD_FLAGS
                .HEAD_SIZE = Temp.HEAD_SIZE + AddHeadSize
                Get #FileNum, , .USize
                Get #FileNum, , .VerNeeded
                Get #FileNum, , .Method
                Get #FileNum, , .COMM_CRC
                ReDim TextBytes(0 To Int2Lng(.HEAD_SIZE - 13) - 1)
                Get #FileNum, , TextBytes
                .COMMENT = StrConv(TextBytes, vbUnicode)
            End With
        Case &H76          'extra information
            With TempInfo
                .HEAD_CRC = Temp.HEAD_CRC
                .HEAD_TYPE = Temp.HEAD_TYPE
                .HEAD_FLAGS = Temp.HEAD_FLAGS
                .HEAD_SIZE = Temp.HEAD_SIZE + AddHeadSize
                ReDim TextBytes(0 To Int2Lng(.HEAD_SIZE - 7) - 1)
                Get #FileNum, , TextBytes
                .INFO = StrConv(TextBytes, vbUnicode)
            End With
        Case &H77          'subblock
            With TempSub
                .HEAD_CRC = Temp.HEAD_CRC
                .HEAD_TYPE = Temp.HEAD_TYPE
                .HEAD_FLAGS = Temp.HEAD_FLAGS
                .HEAD_SIZE = Temp.HEAD_SIZE + AddHeadSize
                Get #FileNum, , .DATA_SIZE
                Get #FileNum, , .SUB_TYPE
                Get #FileNum, , .RESERVED
                ReDim TextBytes(0 To Int2Lng(.HEAD_SIZE) - 14 - 1)
                Get #FileNum, , TextBytes
                .OtherFields = StrConv(TextBytes, vbUnicode)
            End With
        Case &H78          'recovery record
        Case Else
            'Unknown Headertype
        End Select
    Loop
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

Public Property Get FileCount() As Long
    FileCount = PackTotFiles
End Property

Public Property Get FileName(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    FileName = RARFiles(FileNum).FileName
End Property

Public Property Get CommentsFile(FileNum As Long) As String
    CommentsFile = RARFiles(FileNum).COMMENT
End Property

Public Property Get CommentsPack() As String
    CommentsPack = RARArchHead.COMMENT
End Property

Public Property Get IsDir(FileNum As Long) As Boolean
    If NotGood(FileNum) Then Exit Property
    If (RARFiles(FileNum).HEAD_FLAGS And 224) = 224 Then IsDir = True
End Property

Public Property Get Method(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    Method = Methods(CInt(RARFiles(FileNum).Method))
End Property

Public Property Get CRC32(FileNum As Long) As Long
    If NotGood(FileNum) Then Exit Property
    CRC32 = RARFiles(FileNum).File_CRC
End Property

Public Property Get Compressed_Size(FileNum As Long) As Long
    If NotGood(FileNum) Then Exit Property
    Compressed_Size = RARFiles(FileNum).CSize
End Property

Public Property Get UnCompressed_Size(FileNum As Long) As Long
    If NotGood(FileNum) Then Exit Property
    UnCompressed_Size = RARFiles(FileNum).USize
End Property

Public Property Get Encrypted(FileNum As Long) As Boolean
    If NotGood(FileNum) Then Exit Property
    Encrypted = ((RARFiles(FileNum).HEAD_FLAGS And &H4) > 0)
End Property

Public Property Get FileDateTime(FileNum As Long) As Date
    If NotGood(FileNum) Then Exit Property
    FileDateTime = GetZipDate(RARFiles(FileNum).FDate, RARFiles(FileNum).FTime)
End Property

Public Property Get SystemMadeBy(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    SystemMadeBy = SystemName(RARFiles(FileNum).Host_OS)
End Property

Public Property Get VersionMadeBy(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    VersionMadeBy = "Unknown"
End Property

Public Property Get SystemNeeded(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    SystemNeeded = "Unknown"
End Property

Public Property Get VersionNeeded(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    VersionNeeded = VersionTo(RARFiles(FileNum).VersionNeeded)
End Property

Private Function NotGood(FileNum As Long) As Boolean
    If FileNum = 0 Then NotGood = True: Exit Function
    If FileNum > PackTotFiles Then NotGood = True: Exit Function
    If PackFileType = 0 Then NotGood = True: Exit Function
End Function

Private Function DataSize() As Long
    On Error Resume Next
    DataSize = UBound(PackData) + 1
    If Err.Number <> 0 Then
        Err.Clear
        DataSize = 0
    End If
End Function

Private Function SystemName(System As Byte) As String
    Select Case System
    Case 0:     SystemName = "MS-DOS"
    Case 1:     SystemName = "Win32"
    Case 2:     SystemName = "OS/2"
    Case 3:     SystemName = "UNIX"
    Case Else:  SystemName = "unKnown"
    End Select
End Function

Private Function VersionTo(Version As Byte) As String
    VersionTo = Fix(Version / 10) & "." & Version Mod 10
End Function

Private Function Methods(MethodType As Integer) As String
    Select Case MethodType
        Case &H30: Methods = "No Compression"
        Case &H31: Methods = "Fastest"
        Case &H32: Methods = "Fast"
        Case &H33: Methods = "Normal"
        Case &H34: Methods = "Good"
        Case &H35: Methods = "Maximum"
        Case Else: Methods = "Unknown"
    End Select
End Function

Private Function Lng2Int(LngValue As Long) As Integer
    If LngValue > 32767 Then Lng2Int = LngValue - &HFFFF& - 1 Else Lng2Int = LngValue
End Function

Private Function Int2Lng(Value As Integer) As Long
     If Value < 0 Then Int2Lng = &HFFFF& + Value + 1 Else Int2Lng = Value
End Function

