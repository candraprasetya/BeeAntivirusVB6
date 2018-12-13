VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_Encrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0
'This module is created by John Korejwa

Private Key0        As Long 'Zip Encryption Keys
Private Key1(3)     As Long 'Key1(0-3) = LSB to MSB of Key1
Private Key2        As Long
Private Key1Mul(3)  As Long 'Multiplier (constant) used in update_keys
Private Key1Tmp(3)  As Long 'Temp for intermediate Key1
Private CRC As New Cls_CRC32

'ZIP Decryption/Encryption
Private Function decrypt_byte() As Long
    Dim i As Long
    Dim j As Long

    i = (Key2 And 255&) Or 2&
    j = (Key2 And &HFF00&)
    decrypt_byte = (((i Xor 1&) * (i + j) + i * j) And &HFF00&) \ 256&
End Function

Private Sub update_keys(TheChar As Long)
    Dim i    As Long
    Dim j    As Long
    Dim reg  As Long

    Key0 = (((Key0 And &HFFFFFF00) \ &H100&) And &HFFFFFF) Xor (CRC.GetcrcTable((Key0 And &HFF&) Xor TheChar))

    reg = (Key0 And 255&)             'Key1Tmp = Key1 + (Key0 And 255)
    For i = 0 To 3
        reg = reg + Key1(i)
        Key1Tmp(i) = reg And 255&
        reg = reg \ 256&
        Key1(i) = 0&
    Next i

    Key1(0) = 1                       'Key1 = 1 + Key1Tmp * Key1Mul
    For i = 0 To 3
        reg = 0
        For j = 0 To 3 - i
            reg = reg + Key1(i + j) + Key1Tmp(i) * Key1Mul(j)
            Key1(i + j) = reg And 255&
            reg = reg \ 256&
        Next j
    Next i

    Key2 = (((Key2 And &HFFFFFF00) \ &H100&) And &HFFFFFF) Xor (CRC.GetcrcTable((Key2 And &HFF&) Xor Key1(3)))

End Sub

Public Sub ZipPrepareKey(Data() As Byte, PassWord As String)
    Dim C     As Long
    Dim i     As Long

   'Initialize Keys
    Key0 = &H12345678
    Key1(3) = &H23
    Key1(2) = &H45
    Key1(1) = &H67
    Key1(0) = &H89
    Key2 = &H34567890

   'Prepare Keys with password
    For i = 0 To Len(PassWord) - 1
        C = Asc(Mid$(PassWord, i + 1, 1))
        update_keys C
    Next i

   'Crypt 12 byte header
    For i = 0 To 11
        C = Data(i) Xor decrypt_byte()
        update_keys C
        Data(i) = C
    Next i

End Sub

Public Sub ZipDecryptArray(Data() As Byte)
    Dim i As Long
    Dim C As Long
    Dim Min As Long
    Dim Max As Long
    Min = LBound(Data)
    Max = UBound(Data)
    For i = Min To Max
        C = Data(i) Xor decrypt_byte()
        update_keys C
        Data(i) = C
    Next i
End Sub

Public Sub ZipEncryptArray(Data() As Byte, PassWord As String, CRC32 As Long)
    Dim C     As Long 'This procedure Not Tested
    Dim i     As Long
    Dim j     As Long
    Dim Max   As Long

   'Shift data 12 bytes, to make room for encryption header
    ReDim Preserve Data(UBound(Data) + 12)
    Max = UBound(Data)
    For i = Max To 12 Step -1 '    CopyMemory data(12), data(0), UBound(data) - 11
        Data(i) = Data(i - 12)
    Next i

   'Fill first 11 bytes with random data, 12th byte with the MSB of the crc32 value
    Randomize Timer
    For i = 0 To 10
        j = CLng((Max - 11) * Rnd + 12)
        Data(i) = Data(j) Xor Int(256 * Rnd)
    Next i
    j = 0
    For i = 0 To 10
        j = (j + Asc(Mid$(PassWord, (i Mod Len(PassWord)) + 1, 1))) Mod 11
        C = Data(i)
        Data(i) = Data(j)
        Data(j) = C
    Next i
    Data(11) = ((CRC32 And &HFF000000) \ &H1000000) And 255&

    ZipPrepareKey Data, PassWord

   'Encrypt the data
    For i = 12 To UBound(Data)
        C = Data(i) Xor decrypt_byte
        update_keys C
        Data(i) = C
    Next i

End Sub

Private Sub Class_Initialize()
    Key1Mul(3) = &H8
    Key1Mul(2) = &H8
    Key1Mul(1) = &H84
    Key1Mul(0) = &H5
End Sub
