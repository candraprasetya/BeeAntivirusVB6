VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_Tar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'This class file can be used to show the contents of a TAR-file

'data for TAR files
Private Type TarHeaderType          'byte offset
    FName As String * 100           '  0
    Mode As String * 8              '100
    Uid As String * 8               '108
    Gid As String * 8               '116
    Size As String * 12             '124
    Mtime As String * 12            '136
    Chksum As String * 8            '148
    Typeflag  As String * 1         '156
    Linkname As String * 100        '157
    Magic As String * 6             '257
    Version As String * 2           '263
    Uname As String * 32            '265
    Gname As String * 32            '297
    Devmajor As String * 8          '329
    Devminor As String * 8          '337
    Prefix As String * 155          '345
    Total As String * 12            '500    only to total the block size to 512
End Type                            '512
Private Type TarFilesType
    FileName As String
    FileDateUnix As Long
    FDate As Integer
    FTime As Integer
    DataLenght As Long
    DataOffSet As Long
    SumHeader As Long
End Type

Private TarHead As TarHeaderType
Private TarFiles() As TarFilesType
Private Const m_Unpack_Supported As Boolean = True

Public Function Get_Contents(ZipName As String) As Integer
    Dim FileNum As Long
    Dim FileLenght As Long
    Dim LN As Long
    PackFileName = ZipName
    PackComments = ""
    FileNum = FreeFile
    Open PackFileName For Binary Access Read As #FileNum
    FileLenght = LOF(FileNum)
    PackFileType = 0
    PackTotFiles = 0
    ReDim TarFiles(100)
    Do
        If PackTotFiles = UBound(TarFiles) Then ReDim Preserve TarFiles(PackTotFiles + 50)
        Get #FileNum, , TarHead
        TarFiles(PackTotFiles + 1).FileName = Replace(TarHead.FName, vbNullChar, "")
        If TarFiles(PackTotFiles + 1).FileName = "" Then Exit Do
        PackTotFiles = PackTotFiles + 1
        TarFiles(PackTotFiles).FileDateUnix = OctToLng(TarHead.Mtime)
        TarFiles(PackTotFiles).SumHeader = OctToLng(TarHead.Chksum)
        TarFiles(PackTotFiles).FDate = GetIntegerDate(TarFiles(PackTotFiles).FileDateUnix)
        TarFiles(PackTotFiles).FTime = GetIntegerTime(TarFiles(PackTotFiles).FileDateUnix)
        LN = OctToLng(TarHead.Size)
        TarFiles(PackTotFiles).DataLenght = LN
        TarFiles(PackTotFiles).DataOffSet = Seek(FileNum)
        Do While LN > 0
            Seek #FileNum, Seek(FileNum) + 512
            LN = LN - 512
        Loop
    Loop
    ReDim Preserve TarFiles(PackTotFiles)
    If PackTotFiles > 0 Then PackFileType = TARFileType
    Close FileNum
End Function

'Unzip as file and return 0 for good decompression or others for error
Public Function UnPack(ZippedFile() As Boolean, ToPath As String) As Integer
'    Dim ZipHead As Local_Header_Type        'Local Zip Header
    Dim Header As Long
    Dim X As Long
    Dim FileNum As Long
    Dim Y As Long
    Dim TotDir As String                    'Used for new pathnames
    If PackTotFiles = 0 Then UnPack = -10: Exit Function 'nothing to UnPack
    If PackTotFiles <> UBound(ZippedFile) Then
        UnPack = -10                         'need same amount as files in zipfile
        Exit Function
    End If
    Erase PackData
    FileNum = FreeFile
    Open PackFileName For Binary Access Read As #FileNum
    For X = 1 To PackTotFiles
        If ZippedFile(X) = True Then
            If Not IsDir(X) Then            'extract data if no dir
                ReDim PackData(TarFiles(X).DataLenght - 1)
                Seek #FileNum, TarFiles(X).DataOffSet
                If TarFiles(X).DataLenght = 0 Then
                    Erase PackData
                Else
                    ReDim PackData(TarFiles(X).DataLenght - 1)
                    Get #FileNum, , PackData()          'Read the compressed file
                End If
                Call Write_Uncompressed_Data(X, ToPath)
            Else
                TotDir = ToPath
                If Right(TotDir, 1) <> "\" And Right(TotDir, 1) <> "/" Then TotDir = TotDir & "\"
                TotDir = TotDir & TarFiles(X).FileName
                If CreatePath(TotDir) = False Then
                    MsgBox ("error creating directory " & TotDir)
                End If
            End If
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

Public Property Get FileName(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    FileName = TarFiles(FileNum).FileName
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
    If TarFiles(FileNum).DataLenght = 0 Then
        If Right(TarFiles(FileNum).FileName, 1) = "/" Then IsDir = True
    End If
End Property

Public Property Get Method(FileNum As Long) As String
    If NotGood(FileNum) Then Exit Property
    Method = "Stored"
End Property

Public Property Get CRC32(FileNum As Long) As Long
    CRC32 = 0
End Property

Public Property Get Compressed_Size(FileNum As Long) As Long
    If NotGood(FileNum) Then Exit Property
    Compressed_Size = TarFiles(FileNum).DataLenght
End Property

Public Property Get UnCompressed_Size(FileNum As Long) As Long
    If NotGood(FileNum) Then Exit Property
    UnCompressed_Size = TarFiles(FileNum).DataLenght
End Property

Public Property Get Encrypted(FileNum As Long) As Boolean
    If NotGood(FileNum) Then Exit Property
    Encrypted = False
End Property

Public Property Get FileDateTime(FileNum As Long) As Date
    If NotGood(FileNum) Then Exit Property
    FileDateTime = GetZipDate(TarFiles(FileNum).FDate, TarFiles(FileNum).FTime)
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

Private Function OctToLng(OctVal As String) As Long
    Dim OctNum As String
    Dim X As Integer
    OctNum = Trim(Replace(OctVal, vbNullChar, ""))
    Do While Left(OctNum, 1) = "0"
        OctNum = Mid(OctNum, 2)
    Loop
    For X = 1 To Len(OctNum)
        OctToLng = OctToLng + CLng(Val(Mid(OctNum, X, 1))) * 8 ^ (Len(OctNum) - X)
    Next
End Function

Private Function NotGood(FileNum As Long) As Boolean
    If FileNum = 0 Then NotGood = True: Exit Function
    If FileNum > PackTotFiles Then NotGood = True: Exit Function
    If PackFileType = 0 Then NotGood = True: Exit Function
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

Private Sub Write_Uncompressed_Data(FileNum As Long, ToPath As String)
    Dim DSize As Long
    DSize = DataSize
    If TarFiles(FileNum).DataLenght <> DSize Then
        MsgBox "Error in decompressed size"
    End If
    If Write_File(TarFiles(FileNum).FileName, ToPath, PackData(), TarFiles(FileNum).FDate, TarFiles(FileNum).FTime) <> 0 Then
        MsgBox "error writing file"
    End If
End Sub

Private Function DataSize() As Long
    On Error Resume Next
    DataSize = UBound(PackData) + 1
    If Err.Number <> 0 Then
        Err.Clear
        DataSize = 0
    End If
End Function

