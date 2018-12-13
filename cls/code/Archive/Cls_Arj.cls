VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_Arj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'This class file can be used to show the contents of an ARJ-archive

Private Type ARJmainheader
   Id             As Integer
   Headersize     As Integer
   Firsthdrsize   As Byte
   Version        As Byte
   Minversion     As Byte
   Archiveos      As Byte
   Flags          As Byte
   Secversion     As Byte
   Filetype       As Byte
   X_reserved     As Byte
   Createtime     As Long
   Modifytime     As Long
   FileSize       As Long
   Secenvpos      As Long
   Filespecpos    As Integer
   Secenvlength   As Integer
   X_notused      As Integer
End Type
Private Type ARJlocalheader
   Id             As Integer
   Headersize     As Integer
   Firsthdrsize   As Byte
   Version        As Byte
   Minversion     As Byte
   Archiveos      As Byte
   Flags          As Byte
   Method         As Byte
   Filetype       As Byte
   X_reserved     As Byte
   Datemodify     As Long
   Sizecompr      As Long
   Sizeorig       As Long
   Origcrc        As Long
   Filespecpos    As Integer
   Accessmode     As Integer
   Hostdata       As Integer
End Type
Private Type ARJFileType
   Id             As Integer
   Headersize     As Integer
   Firsthdrsize   As Byte
   Version        As Byte
   Minversion     As Byte
   Archiveos      As Byte
   Flags          As Byte
   Method         As Byte
   Filetype       As Byte
   X_reserved     As Byte
   FTime          As Integer
   FDate          As Integer
   Sizecompr      As Long
   Sizeorig       As Long
   Origcrc        As Long
   Filespecpos    As Integer
   Accessmode     As Integer
   Hostdata       As Integer
   StartSplit     As Long
   FileName       As String
   FileComment    As String
   HeaderCRC      As Long
   ExtHeadSize    As Integer
   extHeader      As String
   ExtHeadCRC     As Long
   DataOffSet     As Long
End Type

Private ArjFileData() As ARJFileType
Private Const m_Unpack_Supported As Boolean = False

Public Function Get_Contents(ZipName As String) As Integer
    Dim FileNum As Long
    Dim FileLenght As Long
    Dim LngVal As Long
    Dim IntVal As Integer
    Dim IntVal2 As Integer
    Dim ByteVal As Byte
    Dim LN As Long
    Dim X As Long
    PackFileName = ZipName
    PackComments = ""
    PackTotFiles = 0
    PackFileType = 0
    FileNum = FreeFile
    Open PackFileName For Binary Access Read As #FileNum
    If LOF(FileNum) < 2 Then
        Close #FileNum
        Exit Function
    End If
    'get the end of central date
    Get #FileNum, , IntVal
    If IntVal = ARJHeader Then                              'arj header
        Get #FileNum, , IntVal                              'total header bytes
        Get #FileNum, Seek(FileNum) + IntVal, LngVal        'Header CRC
        Get #FileNum, , IntVal                              'Lenght extra header data
        If IntVal > 0 Then
            Get #FileNum, Seek(FileNum) + IntVal, LngVal    'Extra Header CRC
        End If
        PackFileType = ARJFileType
'Whe reached the local header area so lets collecting the data
        Get #FileNum, , IntVal
        Do While IntVal = ARJHeader                         'arj header
            Get #FileNum, , IntVal2
            If IntVal2 = 0 Then Exit Do                     'HeaderSize
            PackTotFiles = PackTotFiles + 1
            ReDim Preserve ArjFileData(PackTotFiles)
            With ArjFileData(PackTotFiles)
                .Id = IntVal
                .Headersize = IntVal2
                Get #FileNum, , .Firsthdrsize
                Get #FileNum, , .Version
                Get #FileNum, , .Minversion
                Get #FileNum, , .Archiveos
                Get #FileNum, , .Flags
                Get #FileNum, , .Method
                Get #FileNum, , .Filetype
                Get #FileNum, , .X_reserved
                Get #FileNum, , .FTime
                Get #FileNum, , .FDate
                Get #FileNum, , .Sizecompr
                Get #FileNum, , .Sizeorig
                Get #FileNum, , .Origcrc
                Get #FileNum, , .Filespecpos
                Get #FileNum, , .Accessmode
                Get #FileNum, , .Hostdata
                
                If (.Flags And 8) Then Get #FileNum, , .StartSplit
                Do
                    Get #FileNum, , ByteVal
                    If ByteVal = 0 Then Exit Do 'filename complete
                    .FileName = .FileName & Chr(ByteVal)
                Loop
                Do
                    Get #FileNum, , ByteVal
                    If ByteVal = 0 Then Exit Do 'filecomment complete
                    .FileComment = .FileComment & Chr(ByteVal)
                Loop
                Get #FileNum, , .HeaderCRC
                
                Get #FileNum, , .ExtHeadSize
                If .ExtHeadSize > 0 Then
                    .extHeader = String(CLng(.ExtHeadSize), 0)
                    Get #FileNum, , .extHeader
                    Get #FileNum, , .ExtHeadCRC
                End If
                .DataOffSet = Seek(FileNum)
                Get #FileNum, Seek(FileNum) + .Sizecompr, IntVal                'get new header
            End With
        Loop
    End If
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
    FileName = ArjFileData(FileNum).FileName
End Property

Public Property Get CommentsFile(FileNum As Long) As String
    CommentsFile = ArjFileData(FileNum).FileComment
End Property

Public Property Get CommentsPack() As String
    CommentsPack = PackComments
End Property

Public Property Get IsDir(FileNum As Long) As Boolean
    If NotGood(FileNum) Then Exit Property
    If (ArjFileData(FileNum).Flags And 2) > 0 Then IsDir = True
End Property

Public Property Get Method(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    Method = Methods(CInt(ArjFileData(FileNum).Method))
End Property

Public Property Get CRC32(FileNum As Long) As Long
    If NotGood(FileNum) Then Exit Property
    CRC32 = ArjFileData(FileNum).Origcrc
End Property

Public Property Get Compressed_Size(FileNum As Long) As Long
    If NotGood(FileNum) Then Exit Property
    Compressed_Size = ArjFileData(FileNum).Sizecompr
End Property

Public Property Get UnCompressed_Size(FileNum As Long) As Long
    If NotGood(FileNum) Then Exit Property
    UnCompressed_Size = ArjFileData(FileNum).Sizeorig
End Property

Public Property Get Encrypted(FileNum As Long) As Boolean
    If NotGood(FileNum) Then Exit Property
    Encrypted = (ArjFileData(FileNum).Flags And 1) = 1
End Property

Public Property Get FileDateTime(FileNum As Long) As Date
    If NotGood(FileNum) Then Exit Property
    FileDateTime = GetZipDate(ArjFileData(FileNum).FDate, ArjFileData(FileNum).FTime)
End Property

Public Property Get SystemMadeBy(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    SystemMadeBy = SystemName(ArjFileData(FileNum).Archiveos)
End Property

Public Property Get VersionMadeBy(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    VersionMadeBy = VersionTo(ArjFileData(FileNum).Version)
End Property

Public Property Get SystemNeeded(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    SystemNeeded = SystemName(99)
End Property

Public Property Get VersionNeeded(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    VersionNeeded = VersionTo(ArjFileData(FileNum).Version)
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
    Case 0:     SystemName = "MS-DOS and OS/2 (FAT / VFAT / FAT32 file systems)"
    Case 1:     SystemName = "Primos"
    Case 2:     SystemName = "UNIX"
    Case 3:     SystemName = "Amiga"
    Case 4:     SystemName = "MAC-OS"
    Case 5:     SystemName = "OS/2"
    Case 6:     SystemName = "Apple GS"
    Case 7:     SystemName = "Atari ST"
    Case 8:     SystemName = "Next"
    Case 9:     SystemName = "VAX VMS"
    Case Else:  SystemName = "unKnown"
    End Select
End Function

Private Function VersionTo(Version As Byte) As String
    VersionTo = Fix(Version / 10) & "." & Version Mod 10
End Function

Private Function Methods(MethodType As Integer) As String
    Select Case MethodType
        Case 0: Methods = "No Compression"
        Case 1: Methods = "Maximum"
        Case 2: Methods = "Normal"
        Case 3: Methods = "Small"
        Case 4: Methods = "Fastest"
        Case Else: Methods = "Unknown"
    End Select
End Function

