Attribute VB_Name = "basPEFuntion"
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByVal pv6432_lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long

Private Declare Sub RtlMoveMemory Lib "ntdll.dll" (ByVal pDestBuffer As Long, ByVal pSourceBuffer As Long, ByVal nBufferLengthToMove As Long) '<---sebenarnya namanya kurang sesuai, karena yang dilakukan adalah menyalin (copy) isi dari src ke dst.
Private Declare Sub RtlFillMemory Lib "ntdll.dll" (ByVal pDestBuffer As Long, ByVal nDestLengthToFill As Long, ByVal nByteNumber As Long) '<---harusnya byte,tapi memori 32 bit, jadi nggak apa-apa, asal tetap bernilai antara 0 sampai 255.
Private Declare Sub RtlZeroMemory Lib "ntdll.dll" (ByVal pDestBuffer As Long, ByVal nDestLengthToFillWithZeroBytes As Long) '<---reset isi dst yaitu mengisinya dengan bytenumber = 0.
Dim RPT As New classFile

Public Function ValidPE(sPathX As String) As Boolean
Dim IDOSH       As IMAGE_DOS_HEADER
Dim INTSIGN     As IMAGE_NT_HEADERS
Dim DataOut()   As Byte
Dim NewHeader   As Long

If ReadDataFile(sPathX, 1, 64, DataOut) = True Then
   Call RtlMoveMemory(VarPtr(IDOSH), VarPtr(DataOut(0)), Len(IDOSH))
   If IDOSH.e_magic = &H5A4D Then
      NewHeader = IDOSH.e_lfanew
      Call ReadDataFile(sPathX, NewHeader + 1, Len(INTSIGN), DataOut)
      Call RtlMoveMemory(VarPtr(INTSIGN), VarPtr(DataOut(0)), Len(INTSIGN))
      If INTSIGN.SignatureLow = &H4550 And INTSIGN.SignatureHigh = 0 Then
         ValidPE = True
      Else
         ValidPE = False
      End If
   Else
      ValidPE = False
   End If
End If
End Function

Public Function ReadDataFile(ByRef szFileToRead As String, ByVal nStartBase1 As Long, ByVal nLenghtRead As Long, ByRef DataOut() As Byte) As Boolean
Dim nOperation  As Long
Dim hFile       As Long
  
  hFile = RPT.VbOpenFile(szFileToRead, FOR_BINARY_ACCESS_READ, LOCK_NONE)
  
  If hFile > 0 Then
     nOperation = RPT.VbReadFileB(hFile, nStartBase1, nLenghtRead, DataOut)
     If nOperation > 0 Then
        ReadDataFile = True
     Else
        ReadDataFile = False
     End If
  Else
     ReadDataFile = False
  End If
  Call RPT.VbCloseFile(hFile)
End Function

Public Function Valid_PE(hFile As Long) As Boolean
   
   Dim Buffer(12)      As Byte
   Dim lngBytesread    As Long
   Dim tDosHeader      As IMAGE_DOS_HEADER
   
   If (hFile > 0) Then
      ReadFile hFile, tDosHeader, ByVal Len(tDosHeader), lngBytesread, ByVal 0&
      CopyMemory Buffer(0), tDosHeader.e_magic, 2
      If (Chr(Buffer(0)) & Chr(Buffer(1)) = "MZ") Then
         SetFilePointer hFile, tDosHeader.e_lfanew, 0, 0
         ReadFile hFile, Buffer(0), 4, lngBytesread, ByVal 0&
         If (Chr(Buffer(0)) = "P") And (Chr(Buffer(1)) = "E") And (Buffer(2) = 0) And (Buffer(3) = 0) Then
            Valid_PE = True
            Exit Function
         End If
      End If
   End If
   
   Valid_PE = False
   
End Function
