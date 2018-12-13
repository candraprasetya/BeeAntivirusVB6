Attribute VB_Name = "basInfectVir"
Private Declare Sub RtlMoveMemory Lib "ntdll.dll" (ByVal pDestBuffer As Long, ByVal pSourceBuffer As Long, ByVal nBufferLengthToMove As Long) '<---sebenarnya namanya kurang sesuai, karena yang dilakukan adalah menyalin (copy) isi dari src ke dst.
Private Declare Sub RtlFillMemory Lib "ntdll.dll" (ByVal pDestBuffer As Long, ByVal nDestLengthToFill As Long, ByVal nByteNumber As Long) '<---harusnya byte,tapi memori 32 bit, jadi nggak apa-apa, asal tetap bernilai antara 0 sampai 255.
Private Declare Sub RtlZeroMemory Lib "ntdll.dll" (ByVal pDestBuffer As Long, ByVal nDestLengthToFillWithZeroBytes As Long) '<---reset isi dst yaitu mengisinya dengan bytenumber = 0.
Private Type NT_HEADER
  INTSIGN   As IMAGE_NT_HEADERS
  IFILEH    As IMAGE_FILE_HEADER
  IOPTH32   As IMAGE_OPTIONAL_HEADER_32
End Type

Public Function InfectVirX(sFilePE As String) As Boolean
Dim IDOSH      As IMAGE_DOS_HEADER
Dim NTHPE      As NT_HEADER
Dim ISECH()    As IMAGE_SECTION_HEADER
Dim DataOut()  As Byte
Dim NewHeader  As Long
Dim nSection   As Long
Dim iCount     As Long
Dim EPPhysic   As Long
If isProperFile(sFilePE, "EXE DLL SYS") = False Then Exit Function
If ReadDataFile(sFilePE, 1, 64, DataOut) = True Then  ' IDOSH
   Call RtlMoveMemory(VarPtr(IDOSH), VarPtr(DataOut(0)), Len(IDOSH))
   
   If IDOSH.e_magic <> &H5A4D Then ' tidak valid (gagal)
      GoTo LBL_GAGAL
   End If
   
   NewHeader = IDOSH.e_lfanew
   
   Call ReadDataFile(sFilePE, NewHeader + 1, Len(NTHPE), DataOut)
   Call RtlMoveMemory(VarPtr(NTHPE), VarPtr(DataOut(0)), Len(NTHPE))
   
   If NTHPE.INTSIGN.SignatureLow <> &H4550 And NTHPE.INTSIGN.SignatureLow <> 0 Then ' tidak valid (gagal)
      GoTo LBL_GAGAL
   End If
   
   nSection = NTHPE.IFILEH.NumberOfSections ' jumah sectionya
   
   ReDim ISECH(nSection - 1) As IMAGE_SECTION_HEADER
   
   Call ReadDataFile(sFilePE, NewHeader + Len(NTHPE) + 1, Len(ISECH(0)) * nSection, DataOut)
   Call RtlMoveMemory(VarPtr(ISECH(0)), VarPtr(DataOut(0)), Len(ISECH(0)) * nSection)


   For iCount = 0 To nSection - 1
      If (NTHPE.IOPTH32.AddressOfEntryPoint >= ISECH(iCount).VirtualAddress) And (NTHPE.IOPTH32.AddressOfEntryPoint < (ISECH(iCount).VirtualAddress + ISECH(iCount).VirtualSize)) Then
         EPPhysic = ISECH(iCount).PointerToRawData + (NTHPE.IOPTH32.AddressOfEntryPoint - ISECH(iCount).VirtualAddress)
         Exit For
      End If
   Next
   'Tenga
   Call ReadDataFile(sFilePE, EPPhysic + 1, 24, DataOut)
   If DataOut(0) = &H52 And DataOut(1) = &H60 And DataOut(2) = &HB9 Then ' level 1
      If DataOut(7) = &HE8 And DataOut(12) = &H5F And DataOut(13) = &H4F Then ' level 2
         If DataOut(14) = &H66 And DataOut(17) = &H66 And DataOut(22) = &H75 Then ' level 3
            InfectVirX = True
            XinjectVir = "Tenga"
            Exit Function
         End If
      End If
   End If
   'Runonce
   Call ReadDataFile(sFilePE, EPPhysic + 1, 22, DataOut)
   If DataOut(0) = &H60 And DataOut(1) = &HE8 And DataOut(6) = &H8B Then ' level 1
      If DataOut(10) = &HE8 And DataOut(15) = &H61 And DataOut(16) = &H68 Then ' level 2
         If DataOut(21) = &HC3 Then ' level 3
            InfectVirX = True
            XinjectVir = "Runonce"
            Exit Function
         End If
      End If
   End If
   'Sality
   Call ReadDataFile(sFilePE, EPPhysic + 1, 19, DataOut)
   If DataOut(0) = &H60 And DataOut(1) = &HE8 And DataOut(6) = &H33 Then ' level 1
      If DataOut(8) = &H8B And DataOut(11) = &H90 And DataOut(12) = &H81 Then ' level 2
         If DataOut(18) = &H81 Then ' level 3
            InfectVirX = True
            XinjectVir = "Sality"
         End If
      End If
   End If
End If

LBL_GAGAL:
InfectVirX = False
End Function

Public Function CheckAlman(Where As String, hFile As Long, nSize As Long) As Boolean
Dim Awal         As Long
Dim Panjang      As Long
Dim OutData()    As Byte
Dim Alman(1)     As String
Dim IsiFile      As String

On Error GoTo KELUAR
Alman(0) = "¯EI5œ‚ÞùWç‘Ï" ' :: Alman A
Alman(1) = "µí§¶ýÚÿ×Ðþÿÿ·hþÿÿÿÿÿï¡ùÿÿÿÿÿÿÿÿÿÿÿÿ" ':: Alman B


If nSize > 40970 Then 'And isValidPE32(hFile) > 0 Then '+ Yakinkan PE
    Awal = nSize - 40000 ' yah sekitar 40KB an aj ambil datanya
    Panjang = 40000
    Call ReadUnicodeFile2(hFile, Awal, Panjang, OutData)
    IsiFile = StrConv(OutData, vbUnicode)

    If InStr(IsiFile, Alman(0)) > 1 Then 'Or InStr(isiFile, Alman(1)) > 1 Then
       CheckAlman = True
       TutupFile hFile ' nutupnya klo TRUE saja
    Else
       CheckAlman = False
    End If
Else
    CheckAlman = False
End If

KELUAR:
End Function

