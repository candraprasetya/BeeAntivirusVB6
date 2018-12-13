VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_GZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'This class file can be used to show/extract the contents of an gZip-archive

Private Type GZipType
    CompType As Byte                    'Compression method
    Flags As Byte                       'Flags
    FDateUnix As Long                       'last modified FileDate in UNIX format
    ExtFlags As Byte                    'Extended Flags
    OStype As Byte                      'Used OS type
    ExtField As String                  'Extended fields
    FileName As String                  'FileName
    COMMENT As String                   'Comment field
    CRC16 As Integer                    'least significant 2 bytes from CRC32 from header
    Buff(11) As Byte                    'hold encryption header
    CRC32 As Long                       'CRC 32
    CSize As Long                       'Compressed size
    USize As Long                       'Uncompressed size
    DataStart As Long                   'OffSet to start of data
    FDate As Integer                    'Zip date format
    FTime As Integer                    'zip time format
End Type

'Flag values for GZip-files
Private Const GzFlgAscII As Byte = 1        'bit 0 set = file is ascii file
Private Const GzFlgCRC16 As Byte = 2        'bit 1 set Included CRC16 ofthe header
Private Const GzFlgExtFld As Byte = 4       'bit 2 set Extra field is set
Private Const GzFlgOrgName As Byte = 8      'bit 3 set original name included
Private Const GzFlgComment As Byte = 16     'bit 4 set comment are included
Private Const GzFlgEncrypt As Byte = 32     'bit 5 set file is encrypted
Private Const GzFlgReserved As Byte = 192   'bit 6,7 reserved

Private GZipData As GZipType
Private CRC As New Cls_CRC32
Private Encrypt As New Cls_Encrypt
Private Const m_Unpack_Supported As Boolean = True

Public Function Get_Contents(ZipName As String) As Integer
    Dim NextByte As Byte        '1 byte
    Dim Byte2 As Integer        '2 bytes
    Dim FileNum As Long
    Dim FileLenght As Long
    Dim Header As Integer
    Dim LngHeader As Long
    Dim LN As Long
    Dim X As Long
    Dim Temp As String
    If ZipName = "" Then
        Get_Contents = -1       'file don't exist
        Exit Function
    End If
    PackFileName = ZipName
    PackComments = ""
    PackFileType = 0
    FileNum = FreeFile
    Open PackFileName For Binary Access Read As #FileNum
    FileLenght = LOF(FileNum)
    Get #FileNum, , Header
    PackFileType = GZFileType
    PackTotFiles = 1                 'GZip files contain 1 file
    Get #FileNum, , GZipData.CompType
    If GZipData.CompType <> 8 Then GZipData.CompType = 99 'only deflate allowed
    Get #FileNum, , GZipData.Flags              'get the flags
    Get #FileNum, , GZipData.FDateUnix          'get unix date
    Get #FileNum, , GZipData.ExtFlags           'get extra flags
    Get #FileNum, , GZipData.OStype             'get os type
    GZipData.ExtField = ""
    If (GZipData.Flags And GzFlgExtFld) > 0 Then        'read extra field
        Get #FileNum, , Byte2
        LN = Int2Lng(Byte2)
        GZipData.ExtField = String(LN, 0)
        Get #FileNum, , GZipData.ExtField
    End If
    GZipData.FileName = ""
    If (GZipData.Flags And GzFlgOrgName) > 0 Then       'Read original filename (NULL terminated)
        Do
            Get #FileNum, , NextByte
            If NextByte = 0 Then Exit Do 'filename complete
            GZipData.FileName = GZipData.FileName & Chr(NextByte)
        Loop
    End If
    If GZipData.FileName = "" Then GZipData.FileName = GetNameFromFileName
    GZipData.COMMENT = ""
    If (GZipData.Flags And GzFlgComment) > 0 Then      'read comments (NULL terminated)
        Do
            Get #FileNum, , NextByte
            If NextByte = 0 Then Exit Do 'filename complete
            GZipData.COMMENT = GZipData.COMMENT & Chr(NextByte)
        Loop
    End If
    If (GZipData.Flags And GzFlgCRC16) > 0 Then       'get CRC16 of header
        Get #FileNum, , Byte2
        GZipData.CRC16 = Byte2
    End If
    If (GZipData.Flags And GzFlgEncrypt) > 0 Then
        Get #FileNum, , GZipData.Buff
    End If
'here is where the compressed data is
    GZipData.CSize = FileLenght - Seek(FileNum) - 8
    GZipData.DataStart = Seek(FileNum)
    Get #FileNum, FileLenght - 7, GZipData.CRC32
    Get #FileNum, , GZipData.USize
    Close FileNum
'Translate unix time to zip time
    GZipData.FDate = GetIntegerDate(GZipData.FDateUnix)
    GZipData.FTime = GetIntegerTime(GZipData.FDateUnix)
End Function

'Unzip as file and return 0 for good decompression or others for error
Public Function UnPack(ZippedFile() As Boolean, ToPath As String) As Integer
    Dim X As Long
    Dim FileNum As Long
    Dim Y As Long
    Dim PassWord As String
    Dim TotDir As String                    'Used for new pathnames
    If PackTotFiles = 0 Then UnPack = -10: Exit Function 'nothing to UnPack
    Erase PackData
    FileNum = FreeFile
    Open PackFileName For Binary Access Read As #FileNum
    For X = 1 To PackTotFiles
        If ZippedFile(X) = True Then
            If Encrypted(X) Then
                If PassWord = "" Then
                    PassWord = InputBox("Give Password", "Password requered")
                    If PassWord = "" Then
                        UnPack = -1
                        Close FileNum
                        MsgBox "Password is incorrect"
                        Exit Function
                    End If
                End If
                Encrypt.ZipPrepareKey PackData, PassWord
                If PackData(11) <> (((GZipData.CRC32 And &HFF000000) \ &H1000000) And 255&) Then
                    UnPack = -1
                    Close FileNum
                    MsgBox "Password is incorrect"
                    Exit Function
                End If
'adjust the size of instream to delete the decryption data
                For Y = 0 To UBound(PackData) - 12
                    PackData(Y) = PackData(Y + 12)
                Next
                ReDim Preserve PackData(UBound(PackData) - 12)
                Encrypt.ZipDecryptArray PackData
            End If
            ReDim PackData(GZipData.CSize)
            Get #FileNum, GZipData.DataStart, PackData
            Call Inflate(PackData, GZipData.USize - 1)
            Call Write_Uncompressed_Data(1, ToPath)
        End If
    Next
    Close FileNum
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
    FileName = GZipData.FileName
End Property

Public Property Get CommentsFile(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    CommentsFile = GZipData.COMMENT
End Property

Public Property Get CommentsPack() As String
    CommentsPack = GZipData.COMMENT
End Property

Public Property Get IsDir(FileNum As Long) As Boolean
    IsDir = False   'No dirs allowed in gzip so must be a file
End Property

Public Property Get Method(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    Method = Methods(CInt(GZipData.CompType))
End Property

Public Property Get CRC32(FileNum As Long) As Long
    If NotGood(FileNum) Then Exit Property
    CRC32 = GZipData.CRC32
End Property

Public Property Get Compressed_Size(FileNum As Long) As Long
    If NotGood(FileNum) Then Exit Property
    Compressed_Size = GZipData.CSize
End Property

Public Property Get UnCompressed_Size(FileNum As Long) As Long
    If NotGood(FileNum) Then Exit Property
    UnCompressed_Size = GZipData.USize
End Property

Public Property Get Encrypted(FileNum As Long) As Boolean
    If NotGood(FileNum) Then Exit Property
    Encrypted = (GZipData.Flags And GzFlgEncrypt) = 1
End Property

Public Property Get FileDateTime(FileNum As Long) As Date
    If NotGood(FileNum) Then Exit Property
    FileDateTime = GetZipDate(GZipData.FDate, GZipData.FTime)
End Property


Public Property Get SystemMadeBy(FileNum As Long) As String
    SystemMadeBy = SystemName(GZipData.OStype)
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



Private Function NotGood(FileNum As Long) As Boolean
    If FileNum = 0 Then NotGood = True: Exit Function
    If FileNum > PackTotFiles Then NotGood = True: Exit Function
    If PackFileType = 0 Then NotGood = True: Exit Function
End Function

Private Function Int2Lng(Value As Integer) As Long
     If Value < 0 Then Int2Lng = &HFFFF& + Value + 1 Else Int2Lng = Value
End Function

Private Function SystemName(System As Byte) As String
    Select Case System
    Case 0:     SystemName = "MS-DOS and OS/2 (FAT / VFAT / FAT32 file systems)"
    Case 1:     SystemName = "Amiga"
    Case 2:     SystemName = "2 - OpenVMS"
    Case 3:     SystemName = "UNIX"
    Case 4:     SystemName = "VM/CMS"
    Case 5:     SystemName = "Atari ST"
    Case 6:     SystemName = "OS/2 H.P.F.S."
    Case 7:     SystemName = "Macintosh"
    Case 8:     SystemName = "Z-System"
    Case 9:     SystemName = "CP/M"
    Case 10:    SystemName = "Tops-20"
    Case 11:    SystemName = "NTFS filesystem"
    Case 12:    SystemName = "QDos"
    Case 13:    SystemName = "Acorn Risc"
    Case Else:  SystemName = "UnKnown"
    End Select
End Function

Private Sub Write_Uncompressed_Data(FileNum As Long, ToPath As String)
    Dim DSize As Long
    Static sDs As String
    DSize = DataSize
    If DSize > 0 Then
        If GZipData.CRC32 <> CRC.CalcCRC32File(PackData) Then
            MsgBox "CRC calculation failed"
        End If
    Else
        If GZipData.CRC32 <> 0 Then
            MsgBox "CRC error"
        End If
    End If
    If GZipData.USize <> DSize Then
        MsgBox "Error in decompressed size"
    End If
    IsiDataArc = PackData
    sDs = StrConv(IsiDataArc, vbUnicode)
    sDs = Len(sDs)
    EqualArc IsiDataArc, CStr(sDs), GZipData.FileName, PathDalamArc
    'If Write_File(GZipData.FileName, ToPath, PackData, GZipData.FDate, GZipData.FTime) Then
    '    MsgBox "error writing file"
    'End If
End Sub

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
        Case 0: Methods = "Stored"
        Case 8: Methods = "Deflate"
        Case Else: Methods = "Unknown"
    End Select
End Function

Private Function GetNameFromFileName()
    Dim name As String
    name = mbStripFileName(PackFileName, False)
    If InStr(UCase(name), ".GZ") Then
        GetNameFromFileName = Left(name, InStr(UCase(name), ".GZ") - 1)
        Exit Function
    End If
    If InStr(UCase(name), ".Z") Then
        GetNameFromFileName = Left(name, InStr(UCase(name), ".Z") - 1)
        Exit Function
    End If
    If InStr(UCase(name), ".TGZ") Then
        GetNameFromFileName = Left(name, InStr(UCase(name), ".TGZ") - 1) & ".tar"
        Exit Function
    End If
    If InStr(UCase(name), ".TAZ") Then
        GetNameFromFileName = Left(name, InStr(UCase(name), ".TAZ") - 1) & ".tar"
        Exit Function
    End If
End Function

'----------------------------------------------------------
'This function is used to extract the filename
'input:
'Stripfile = Filename with or without directory
'StripBaseOnly = Treu if only filename is needed
'                False if also the extension is needed
'return:
'StripFileName = filename
'----------------------------------------------------------
Private Function mbStripFileName(Stripfile As String, StripBaseOnly As Boolean) As String

    Dim Counter As Integer, Stripped As String
    On Local Error Resume Next

    If InStr(Stripfile, "\") Then
        For Counter = Len(Stripfile) To 1 Step -1
            If Mid$(Stripfile, Counter, 1) = "\" Then
                Stripped = Mid$(Stripfile, Counter + 1)
            Exit For
            End If
        Next Counter
    ElseIf InStr(Stripfile, ":") = 2 Then
        Stripped = Mid$(Stripfile, 3)
    Else
        Stripped = Stripfile
    End If
  
    If StripBaseOnly = True Then
        If InStr(Stripped, ".") > 0 Then
            Stripped = Left$(Stripped, InStr(Stripped, ".") - 1)
        End If
    End If

    mbStripFileName = Stripped
End Function

Private Function StampToData(Stamp) As Date
    StampToData = CDate(CDbl(DateSerial(1970, 1, 1)) + CDbl((CLng(Stamp) / 86400)))
End Function

Private Function GetIntegerDate(Stamp As Long) As Integer
    Dim Dat As String
    Dim FD As Long
    Dat = StampToData(Stamp)
    FD = (Year(Dat) - 1980) * 2 ^ 9
    FD = FD + (Month(Dat) * 2 ^ 5)
    FD = FD + Day(Dat)
    If FD > 32767 Then GetIntegerDate = FD - &HFFFF& - 1 Else GetIntegerDate = FD
End Function

Private Function GetIntegerTime(Stamp As Long) As Integer
    Dim Dat As String
    Dim FT As Long
    Dat = StampToData(Stamp)
    FT = Hour(Dat) * 2 ^ 11
    FT = FT + (Minute(Dat) * 2 ^ 5)
    FT = FT + Second(Dat)
    If FT > 32767 Then GetIntegerTime = FT - &HFFFF& - 1 Else GetIntegerTime = FT
End Function

