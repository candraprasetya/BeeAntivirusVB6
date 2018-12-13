Attribute VB_Name = "basChecksum"
' ########################################################
' Module untuk penanganan ceksum
'
'

' Ceksumer Non PE
Public Function MYCeksum(ByRef sFile As String, ByVal hFile As Long) As String ' var hFile gak ngaruh disini
On Error Resume Next
Dim DataOut()   As Byte
Dim TheSize     As Long
Dim iCount      As Long

TheSize = GetSizeFile(hFile)

If TheSize <= 0 Then
    MYCeksum = vbNullString
    Exit Function
End If

' Walapun 300 tapi variatif
If TheSize > 4300 Then
    Call ReadUnicodeFile2(hFile, 4000, 300, DataOut)
Else
    '[APTX] hilangkan salin antar variable yang nggak perlu
    If TheSize > 300 Then
        Call ReadUnicodeFile2(hFile, 1, 300, DataOut)
    Else
        Call ReadUnicodeFile2(hFile, 1, TheSize, DataOut)
    End If
End If

For iCount = 0 To 299 Step 10
    MYCeksum = MYCeksum & Hex$(DataOut(iCount))
Next iCount

MYCeksum = StrReverse(MYCeksum) ' dibalik  supaya variatif

'If MYCeksum = String(Len(MYCeksum), "0") Then MYCeksum = "Tolong pakai ceksum Cadangan !"
Erase DataOut
End Function

' Fungsi kedua-nya sama cuma yang satu ini digunakan lebih kusus (1 file) aj dalam looping -->pada saat ArrS
Function MYCeksum2(ByRef sFile As String) As String
On Error Resume Next
Dim sTmpFile As String
Dim IsiFile  As String
Dim iCount   As Integer


sTmpFile = ReadUnicodeFile(sFile)

If Len(sTmpFile) = 0 Then
    MYCeksum2 = vbNullString
    Exit Function
End If

If Len(sTmpFile) > 4300 Then
    IsiFile = Mid(sTmpFile, 4000)
Else
    IsiFile = sTmpFile
End If

For iCount = 1 To 300 Step 10
    MYCeksum2 = MYCeksum2 & Hex$(Asc(Mid(IsiFile, iCount, 1)))
Next iCount

MYCeksum2 = StrReverse(MYCeksum2) ' dibalik  supaya variatif

End Function


' Apapun bentuk file virusnya, pasti ada nilainya. Tpi untuk antisipasi
' dipakai jika ceksum di atas mengahasilkan nilai 00000 atau vbNullString
Public Function MYCeksumCadangan(ByRef sFile As String, hFile As Long) As String
On Error Resume Next
Dim DataOut()   As Byte
Dim TmpCount    As Long
Dim iCount      As Long
Dim TheSize     As Long

TheSize = GetSizeFile(hFile)
If TheSize <= 0 Then
    MYCeksumCadangan = vbNullString
    Exit Function
End If

If TheSize > 400 Then ' ambil lebh banyak krena hanya cdangan
    Call ReadUnicodeFile2(hFile, 1, 400, DataOut)
Else
    Call ReadUnicodeFile2(hFile, 1, TheSize, DataOut)
End If

For iCount = 0 To 199
    TmpCount = TmpCount + DataOut(iCount) ^ 2.2
Next iCount

MYCeksumCadangan = Hex$(TmpCount)
TmpCount = 0

For iCount = 200 To 399
    TmpCount = TmpCount + DataOut(iCount) ^ 2.2
Next iCount

MYCeksumCadangan = MYCeksumCadangan & Hex$(TmpCount)

Erase DataOut

End Function

