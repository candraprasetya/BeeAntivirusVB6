Attribute VB_Name = "basPEChecksum"
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByVal pv6432_lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal pv_lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, ByVal pv_lpNumberOfBytesRead As Long, ByVal pv_lpOverlapped As Long) As Long

Public Function GetHIBCeksum(AlamatFile As String, nBased As Long) As String
Dim IDOSH               As IMAGE_DOS_HEADER
Dim INTH32              As IMAGE_NT_HEADERS_32
Dim ISECH()             As IMAGE_SECTION_HEADER
Dim AddNewHeaderBase0   As Long
Dim nNumberBytesOpsRet  As Long
Dim nSection            As Long
Dim pPhysicEP           As Long
Dim hFilePE             As Long
Dim RetFunct            As Long
Dim iCount              As Integer
Dim OutData()           As Byte

Dim CFILE       As classFile
    Set CFILE = New classFile

    hFilePE = CFILE.VbOpenFile(AlamatFile, FOR_BINARY_ACCESS_READ, LOCK_NONE)
    If hFilePE <= 0 Then
        GoTo LBL_GAGAL
    End If
    nFileLen = CFILE.VbFileLen(hFilePE)

    If nFileLen <= Len(IDOSH) Then '---nggak memeriksa file yg berukuran kurang dari atau sama dengan dosheader atau lebih besar dari 2GB.
         GoTo LBL_GAGAL
    End If

    Call SetFilePointer(hFilePE, 0, 0, 0)  '---Base0, set ke pointer pertama
    RetFunct = ReadFile(hFilePE, VarPtr(IDOSH), Len(IDOSH), VarPtr(nNumberBytesOpsRet), 0) ' base1
   
    If RetFunct = 0 Then
        GoTo LBL_GAGAL
    End If
    
    ' cek header DOS
    If IDOSH.e_magic <> &H5A4D Then '---"MZ" mungkin udah DOS valid tapi.... cek lagi gak ya :D
        GoTo LBL_GAGAL
    End If
    
    AddNewHeaderBase0 = IDOSH.e_lfanew

    Call SetFilePointer(hFilePE, AddNewHeaderBase0, 0, 0) '---Base0.
    
    RetFunct = ReadFile(hFilePE, VarPtr(INTH32), Len(INTH32), VarPtr(nNumberBytesOpsRet), 0)
       
    If RetFunct = 0 Then
        GoTo LBL_GAGAL
    End If
    
    
    ' cek apa benar-benar PE 32 byte
    If (INTH32.FileHeader.Characteristics And IMAGE_FILE_32BIT_MACHINE) <> IMAGE_FILE_32BIT_MACHINE Then '---hanya cek di file pe32bit.
         GoTo LBL_GAGAL
    End If
  
    nSection = INTH32.FileHeader.NumberOfSections
  
    If nSection <= 0 Then ' masak 0 jumlah sectionya
       GoTo LBL_GAGAL
    End If
    
    
    '---cek sectionheader: (disini bisa ketemu EP fisik nya)
    ReDim ISECH(nSection - 1) As IMAGE_SECTION_HEADER
    Call SetFilePointer(hFilePE, AddNewHeaderBase0 + Len(INTH32), 0, 0) '---Base0. INTH32=248 Bytes, set pointernya - lebih irit
    RetFunct = ReadFile(hFilePE, VarPtr(ISECH(0)), Len(ISECH(0)) * nSection, VarPtr(nNumberBytesOpsRet), 0) ' yang akan dibaca ukuran type section (40bytes) x jumlah section
    For iCount = 0 To nSection - 1 ' untuk mencari Entri Point Fisik ajh lalu ditata
      If (INTH32.OptionalHeader.AddressOfEntryPoint >= ISECH(iCount).VirtualAddress) And (INTH32.OptionalHeader.AddressOfEntryPoint < (ISECH(iCount).VirtualAddress + ISECH(iCount).VirtualSize)) Then
           pPhysicEP = ISECH(iCount).PointerToRawData + (INTH32.OptionalHeader.AddressOfEntryPoint - ISECH(iCount).VirtualAddress)
           '---EP-di-file-fisik-ya ketemu,deh!
           Call CFILE.VbReadFileB(hFilePE, pPhysicEP + 1, nBased, OutData()) '--Base0.' nBased banyaknya data yang dibaca
           GetHIBCeksum = TataByte(OutData)
       End If
    Next
   CFILE.VbCloseFile hFilePE
Exit Function

LBL_GAGAL:
    GetHIBCeksum = ""
    CFILE.VbCloseFile hFilePE
End Function

' Menata byte ke bentuk string
Function TataByte(sByte() As Byte) As String
Dim i As Integer
For i = 1 To UBound(sByte) + 1
    TataByte = TataByte & ":" & Hex(sByte(i - 1))
Next
End Function

