VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_CRC32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'This class is used to calculate a CRC32 checksum

Private crcTable(255) As Long
Private m_CRC32 As Long

'This sub Resets to the initial crc32 value
'This must be done explicitly when using CalcCRC32Byte
'Usage: ???.ResetCRC32
Public Sub ResetCRC32()
    m_CRC32 = -1
End Sub

'This property updates the current CRC32 value by adding new values to it
'Usage: ???.CalcCRC32Byte = value to add
Public Property Let CalcCRC32Byte(NewVal As Byte)
    m_CRC32 = (((m_CRC32 And &HFFFFFF00) \ &H100) And &HFFFFFF) Xor (crcTable((m_CRC32 And &HFF) Xor NewVal))
End Property

'This property show u the current CRC32 value
'usage: CRC = ???.CRC32
Public Property Get CRC32() As Long
    CRC32 = m_CRC32 Xor &HFFFFFFFF
End Property

'This property sets the CRC32 value to an initial position
'Usage: ???.CRC32 = CRC
Public Property Let CRC32(CrcValue As Long)
    m_CRC32 = CrcValue
End Property

Public Property Get GetcrcTable(TabNum As Integer) As Long
    GetcrcTable = crcTable(TabNum)
End Property

'This property calculates the CRC32 for a complete file
'usage: FileCRC32 = ???.CalcCRC32File(Array of the file)
Public Property Get CalcCRC32File(ByteArray() As Byte) As Long
    Dim i As Long
    Call ResetCRC32
    For i = 0 To UBound(ByteArray)
        m_CRC32 = (((m_CRC32 And &HFFFFFF00) \ &H100) And &HFFFFFF) Xor (crcTable((m_CRC32 And &HFF) Xor ByteArray(i)))
    Next i
    m_CRC32 = m_CRC32 Xor &HFFFFFFFF
    CalcCRC32File = m_CRC32
End Property

'This sub is private and initialates the CRCTable used by the properties
Private Sub Class_Initialize()
    Dim i As Long
    Dim j As Long
    Dim CRC As Long
    For i = 0 To 255
        CRC = i
        For j = 0 To 7
            If CRC And 1 Then
              CRC = (((CRC And &HFFFFFFFE) \ 2) And &H7FFFFFFF) Xor &HEDB88320
            Else
              CRC = ((CRC And &HFFFFFFFE) \ 2) And &H7FFFFFFF
            End If
        Next j
        crcTable(i) = CRC
    Next i
End Sub
