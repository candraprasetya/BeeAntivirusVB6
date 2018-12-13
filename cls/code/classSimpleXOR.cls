VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classSimpleXOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Key()             As Byte
Private m_KeyLen            As Long
Private m_KeyValue          As String
Public Event progress(Percent As Long)

Public Sub DecryptByte(ByteArray() As Byte, _
                       Optional Key As String)

    Call EncryptByte(ByteArray(), Key)

End Sub

Public Sub DecryptFile(SourceFile As String, _
                       DestFile As String, _
                       Optional Key As String)

  Dim Filenr      As Long
  Dim ByteArray() As Byte

   ' If (Not FileExist(SourceFile)) Then
    '    Exit Sub
    'End If
    Filenr = FreeFile
    Open SourceFile For Binary As #Filenr
    ReDim ByteArray(0 To LOF(Filenr) - 1)
    Get #Filenr, , ByteArray()
    Close #Filenr
    Call DecryptByte(ByteArray(), Key)
  '  If (FileExist(DestFile)) Then
   '     Kill DestFile
   ' End If
    Filenr = FreeFile
    Open DestFile For Binary As #Filenr
    Put #Filenr, , ByteArray()
    Close #Filenr

End Sub

Public Sub EncryptByte(ByteArray() As Byte, _
                       Optional ByVal Key As String)

  
  Dim offset      As Long
  Dim ByteLen     As Long
  Dim ResultLen   As Long
  Dim CurrPercent As Long
  Dim NextPercent As Long

    If (Len(Key) > 0) Then
        Me.Key = Key
    End If
    ByteLen = UBound(ByteArray) + 1
    ResultLen = ByteLen
    For offset = 0 To (ByteLen - 1)
        ByteArray(offset) = ByteArray(offset) Xor m_Key(offset Mod m_KeyLen)
        If (offset >= NextPercent) Then
            CurrPercent = Int((offset / ResultLen) * 100)
            NextPercent = (ResultLen * ((CurrPercent + 1) / 100)) + 1
            RaiseEvent progress(CurrPercent)
        End If
    Next '  OFFSET
    If (CurrPercent <> 100) Then
        RaiseEvent progress(100)
    End If

End Sub

Public Sub EncryptFile(SourceFile As String, _
                       DestFile As String, _
                       Optional Key As String)

  Dim Filenr      As Long
  Dim ByteArray() As Byte

    'If (Not FileExist(SourceFile)) Then
        'MsgBox "Source file doesn't exist", "SimpleXOR Error Procedure"
        'Exit Sub
    'End If
    Filenr = FreeFile
    Open SourceFile For Binary As #Filenr
    ReDim ByteArray(0 To LOF(Filenr) - 1)
    Get #Filenr, , ByteArray()
    Close #Filenr
    Call EncryptByte(ByteArray(), Key)
    'If (FileExist(DestFile)) Then
       ' Kill DestFile
   ' End If
    Filenr = FreeFile
    Open DestFile For Binary As #Filenr
    Put #Filenr, , ByteArray()
    Close #Filenr

End Sub

Public Property Let Key(New_Value As String)

    If (m_KeyValue = New_Value) Then
        Exit Property
    End If
    m_KeyValue = New_Value
    m_KeyLen = Len(New_Value)
    m_Key() = StrConv(m_KeyValue, vbFromUnicode)

End Property


