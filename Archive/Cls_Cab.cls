VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_Cab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'This class file can be used to show the contents of an CAB-archive

Private Type CabFileHeaderType
    signature     As Long        ' MSCF (cabinet file signature )
    HeadCRC       As Long        'CRC 32 of the cab header
    cbCabinet     As Long        'size of this cabinet file in bytes
    FolderCRC     As Long        'CRC 32 of the folder header
    coffFiles     As Long        'offset of the first CFFILE entry
    FilesCRC      As Long        'CRC 32 of the Files header
    versionMinor  As Byte        'cabinet file format version, minor
    versionMajor  As Byte        'cabinet file format version, major
    cFolders      As Integer     'number of CFFOLDER entries in this cabinet
    cFiles        As Integer     'number of CFFILE entries in this cabinet
    Flags         As Integer     'cabinet file option indicators
                                 'bit 0 = Has Previous Cab file
                                 'bit 1 = Has Next cab file
                                 'Bit 2 = reserve
    setID         As Integer     'must be the same for all cabinets in a set
    iCabinet      As Integer     'number of this cabinet file in a set
  '  cbCFHeader    As Integer     '(optional) size of per-cabinet reserved area
  '  cbCFFolder    As Byte        '(optional) size of per-folder reserved area
  '  cbCFData      As Byte        '(optional) size of per-datablock reserved area
  '  abReserve     As Byte        '(optional) per-cabinet reserved area
  '  szCabinetPrev As Byte        '(optional) name of previous cabinet file
  '  szDiskPrev    As Byte        '(optional) name of previous disk
  '  szCabinetNext As Byte        '(optional) name of next cabinet file
  '  szDiskNext    As Byte        '(optional) name of next disk
End Type


Private Type CFFolderType
    coffCabStart  As Long     'offset of the first CFDATA block in this folder
    cCFData       As Integer  '??number of CFDATA blocks in this folder
    typeCompress  As Integer  'compression type indicator
End Type

Private Type CFFileType
    USize            As Long     'uncompressed size of this file in bytes
    UoffFolderStart  As Long     'uncompressed offset of this file in the folder
    IFolder          As Integer  'index into the CFFOLDER area
                                 '&h0000 = FIRST
                                 '&h0001 = NEXT
                                 '&hFFFE = SPLIT
                                 '&hFFFF = CONTINUED
    FDate            As Integer  'date stamp for this file
    FTime            As Integer  'time stamp for this file
    Attribs          As Integer  'attribute flags for this file
                                 'and &h0001 = READONLY
                                 'and &h0002 = HIDDEN
                                 'and &h0004 = SYSTEM
                                 'and &h0008 = VOLUME
                                 'and &h0010 = DIRECTORY
                                 'and &h0020 = ARCHIVE
    'szName is variable length string with Chr$(0) terminator
    'See GetInfo to see how seek is adjusted for block alignment
    FileName         As String   'name of this file
End Type

'Would have been nice if the Crc and
'Compressed size were in CFFILE above

Private Type CFDataType
    CRC32      As Long    'checksum of this CFDATA entry
    CSize      As Integer 'number of compressed bytes in this block
    cbUncomp   As Integer 'number of uncompressed bytes in this block
  '  abReserve  As Byte    '(optional) per-datablock reserved area
  '  ab[cbData] As Byte    'compressed data bytes
End Type

Private Type CabFileDataType
    USize            As Long     'uncompressed size of this file in bytes
    UoffFolderStart  As Long     'uncompressed offset of this file in the folder
    IFolder          As Integer  'index into the CFFOLDER area
                                 '&h0000 = FIRST
                                 '&h0001 = NEXT
                                 '&hFFFE = SPLIT
                                 '&hFFFF = CONTINUED
    FDate            As Integer  'date stamp for this file
    FTime            As Integer  'time stamp for this file
    Attribs          As Integer  'attribute flags for this file
                                 'and &h0001 = READONLY
                                 'and &h0002 = HIDDEN
                                 'and &h0004 = SYSTEM
                                 'and &h0008 = VOLUME
                                 'and &h0010 = DIRECTORY
                                 'and &h0020 = ARCHIVE
    'szName is variable length string with Chr$(0) terminator
    'See GetInfo to see how seek is adjusted for block alignment
    FileName         As String   'name of this file
    CRC32      As Long           'checksum of this CFDATA entry
    CSize      As Integer        'number of compressed bytes in this block
    cbUncomp   As Integer        'number of uncompressed bytes in this block
    DataOffSet As Long           'start position if the compressed data
    Method     As Integer
  '  abReserve  As Byte    '(optional) per-datablock reserved area
  '  ab[cbData] As Byte    'compressed data bytes
End Type

Private CabFiles() As CabFileDataType
Private CabHead As CabFileHeaderType
Private Const m_Unpack_Supported As Boolean = False

Public Function Get_Contents(ZipName As String) As Integer
    Dim FileNum As Long
    Dim FileLenght As Long
    Dim ByteVal As Byte
    Dim LN As Long
    Dim X As Long
    Dim CabFolder() As CFFolderType
    Dim CabReserve As Integer
    Dim FolderReserve As Byte
    Dim dataReserve As Byte
    Dim CAbPrevName As String
    Dim CabPrevDisk As String
    Dim CabNextName As String
    Dim CabNextDist As String
    PackFileName = ZipName
    PackComments = ""
    PackFileType = 0
    FileNum = FreeFile
    Open PackFileName For Binary Access Read As #FileNum
    If LOF(FileNum) < Len(CabHead) Then
        Close #FileNum
        Exit Function
    End If
    'get the end of central date
    Get #FileNum, , CabHead
    If CabHead.signature = &H4643534D Then
        PackFileType = CABFileType
        If (CabHead.Flags And 4) Then       'reserve
            Get #FileNum, , CabReserve      'Reserved header space
            Get #FileNum, , FolderReserve   'Reserved folder space
            Get #FileNum, , dataReserve     'Reserved Datablock space
            If CabReserve > 0 Then
                Seek #FileNum, Seek(FileNum) + CabReserve + 1   'Skip reserved block
            End If
        End If
        If (CabHead.Flags And 1) Then       'Has Previous
            Do
                Get #FileNum, , ByteVal
                If ByteVal = 0 Then Exit Do
                CAbPrevName = CAbPrevName & Chr(ByteVal)
            Loop
            Do
                Get #FileNum, , ByteVal
                If ByteVal = 0 Then Exit Do
                CabPrevDisk = CabPrevDisk & Chr(ByteVal)
            Loop
        End If
        If (CabHead.Flags And 2) Then       'Has Next
            Do
                Get #FileNum, , ByteVal
                If ByteVal = 0 Then Exit Do
                CabNextName = CabNextName & Chr(ByteVal)
            Loop
            Do
                Get #FileNum, , ByteVal
                If ByteVal = 0 Then Exit Do
                CabNextDist = CabNextDist & Chr(ByteVal)
            Loop
        End If
        ReDim CabFolder(CabHead.cFolders)
        For X = 1 To CabHead.cFolders
            Get #FileNum, , CabFolder(X)
            If FolderReserve > 0 Then
                Seek #FileNum, Seek(FileNum) + FolderReserve + 1   'Skip reserved block
            End If
        Next
        ReDim CabFiles(CabHead.cFiles)
        If Seek(FileNum) <> CabHead.coffFiles + 1 Then Seek #FileNum, CabHead.coffFiles + 1
        PackTotFiles = CabHead.cFiles
        For X = 1 To PackTotFiles
            With CabFiles(X)
                Get #FileNum, , .USize
                Get #FileNum, , .UoffFolderStart
                Get #FileNum, , .IFolder
                Get #FileNum, , .FDate
                Get #FileNum, , .FTime
                Get #FileNum, , .Attribs
                Do
                    Get #FileNum, , ByteVal
                    If ByteVal = 0 Then Exit Do
                    .FileName = .FileName & Chr(ByteVal)
                Loop
                .Method = CabFolder(1).typeCompress
            End With
        Next
'At this point the CFDatablock begin
'These are compressed blocks from uncompressed blocks up to 32K
'The files are stored into a buff of 32K until its full, After that the compression
'starts. That's why there are no CRC-value of the independed files

'        If Seek(FileNum) <> CabFolder(1).coffCabStart + 1 Then Seek #FileNum, CabFolder(1).coffCabStart + 1
'        For X = 1 To PackTotFiles
'            With CabFiles(X)
'                Get #FileNum, , .CRC32
'                Get #FileNum, , .CSize
'                Get #FileNum, , .cbUncomp
'                .DataOffSet = Seek(FileNum)
'                .Method = CabFolder(1).typeCompress
'                Seek #FileNum, Seek(FileNum) + .CSize
'            End With
'        Next
    End If
'    Close FileNum
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
    FileName = CabFiles(FileNum).FileName
End Property

Public Property Get CommentsFile(FileNum As Long) As String
    CommentsFile = ""
End Property

Public Property Get CommentsPack() As String
    CommentsPack = ""
End Property

Public Property Get IsDir(FileNum As Long) As Boolean
    If NotGood(FileNum) Then Exit Property
    If (CabFiles(FileNum).Attribs And &H10) > 0 Then IsDir = True
End Property

Public Property Get Method(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    Method = Methods(CInt(CabFiles(FileNum).Method And &HF))
End Property

Public Property Get CRC32(FileNum As Long) As Long
    If NotGood(FileNum) Then Exit Property
    CRC32 = CabFiles(FileNum).CRC32
End Property

Public Property Get Compressed_Size(FileNum As Long) As Long
    If NotGood(FileNum) Then Exit Property
    Compressed_Size = CabFiles(FileNum).CSize
End Property

Public Property Get UnCompressed_Size(FileNum As Long) As Long
    If NotGood(FileNum) Then Exit Property
    UnCompressed_Size = CabFiles(FileNum).USize
End Property

Public Property Get Encrypted(FileNum As Long) As Boolean
    If NotGood(FileNum) Then Exit Property
    Encrypted = False
End Property

Public Property Get FileDateTime(FileNum As Long) As Date
    If NotGood(FileNum) Then Exit Property
    FileDateTime = GetZipDate(CabFiles(FileNum).FDate, CabFiles(FileNum).FTime)
End Property

Public Property Get SystemMadeBy(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    SystemMadeBy = "UnKnown"
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
    VersionNeeded = Trim(CabHead.versionMajor & "." & CabHead.versionMinor)
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

Private Function VersionTo(Version As Byte) As String
    VersionTo = Fix(Version / 10) & "." & Version Mod 10
End Function

Private Function Methods(MethodType As Integer) As String
    Select Case MethodType
        Case 0: Methods = "No Compression"
        Case 1: Methods = "MsZip"
        Case 2: Methods = "Quantum"
        Case 3: Methods = "Lzx"
        Case Else: Methods = "Unknown"
    End Select
End Function


