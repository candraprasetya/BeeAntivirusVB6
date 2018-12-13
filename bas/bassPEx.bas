Attribute VB_Name = "basPEx"
' ALL ABOUT PE
' I Love PE

Public Const IMAGE_FILE_32BIT_MACHINE = &H100 ' Karektristic 32bit

Private Type LARGE_INTEGER
    LoPart                      As Long
    HiPart                      As Long
End Type

Public Type IMAGE_DOS_HEADER
    e_magic                 As Integer ' Magic number "MZ"
    e_cblp                  As Integer ' Bytes on last page of file
    e_cp                    As Integer ' Pages in file
    e_crlc                  As Integer ' Relocations
    e_cparhdr               As Integer ' Size of header in paragraphs
    e_minalloc              As Integer ' Minimum extra paragraphs needed
    e_maxalloc              As Integer ' Maximum extra paragraphs needed
    e_ss                    As Integer ' Initial (relative) SS value
    e_sp                    As Integer ' Initial SP value
    e_csum                  As Integer ' Checksum
    e_ip                    As Integer ' Initial IP value
    e_cs                    As Integer ' Initial (relative) CS value
    e_lfarlc                As Integer ' File address of relocation table
    e_ovno                  As Integer ' Overlay number
    e_res(0 To 3)           As Integer ' Reserved words
    e_oemid                 As Integer ' OEM identifier (for e_oeminfo)
    e_oeminfo               As Integer ' OEM information; e_oemid specific
    e_res2(0 To 9)          As Integer ' Reserved words
    e_lfanew                As Long ' File address of new exe header
End Type

Public Type IMAGE_DATA_DIRECTORY_32
    VirtualAddress          As Long
    nSize                   As Long
End Type

Public Type IMAGE_DATA_DIRECTORY_64
    VirtualAddress          As Long
    nSize                   As Long
End Type

Public Type IMAGE_FILE_HEADER '---20 bytes.
    Machine                 As Integer
    NumberOfSections        As Integer
    TimeDateStamp           As Long
    PointerToSymbolTable    As Long
    NumberOfSymbols         As Long
    SizeOfOptionalHeader    As Integer
    Characteristics         As Integer
End Type

Public Type IMAGE_OPTIONAL_HEADER_32
    '---Standard fields:
    Magic                       As Integer
    MajorLinkerVersion          As Byte
    MinorLinkerVersion          As Byte
    SizeOfCode                  As Long
    SizeOfInitializedData       As Long
    SizeOfUninitializedData     As Long
    AddressOfEntryPoint         As Long 'PEHDR+40
    BaseOfCode                  As Long
    BaseOfData                  As Long
    '---NT additional fields:
    ImageBase                   As Long
    SectionAlignment            As Long
    FileAlignment               As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion           As Integer
    MinorImageVersion           As Integer
    MajorSubsystemVersion       As Integer
    MinorSubsystemVersion       As Integer
    Win32VersionValue           As Long
    SizeOfImage                 As Long
    SizeOfHeaders               As Long
    CheckSum                    As Long
    Subsystem                   As Integer
    DllCharacteristics          As Integer
    SizeOfStackReserve          As Long
    SizeOfStackCommit           As Long
    SizeOfHeapReserve           As Long
    SizeOfHeapCommit            As Long
    LoaderFlags                 As Long
    NumberOfRvaAndSizes         As Long
    DataDirectory(0 To 15)      As IMAGE_DATA_DIRECTORY_32 '%IMAGE_DIRECTORY_ENTRY_EXPORT       =  0   ' Export Directory
End Type

Public Type IMAGE_OPTIONAL_HEADER_64
    '---Standard fields:
    Magic                       As Integer
    MajorLinkerVersion          As Byte
    MinorLinkerVersion          As Byte
    SizeOfCode                  As Long
    SizeOfInitializedData       As Long
    SizeOfUninitializedData     As Long
    AddressOfEntryPoint         As Long
    BaseOfCode                  As Long
    '---NT additional fields:
    ImageBase                   As LARGE_INTEGER
    SectionAlignment            As Long
    FileAlignment               As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion           As Integer
    MinorImageVersion           As Integer
    MajorSubsystemVersion       As Integer
    MinorSubsystemVersion       As Integer
    Win32VersionValue           As Long
    SizeOfImage                 As Long
    SizeOfHeaders               As Long
    CheckSum                    As Long
    Subsystem                   As Integer
    DllCharacteristics          As Integer
    SizeOfStackReserve          As LARGE_INTEGER
    SizeOfStackCommit           As LARGE_INTEGER
    SizeOfHeapReserve           As LARGE_INTEGER
    SizeOfHeapCommit            As LARGE_INTEGER
    LoaderFlags                 As Long
    NumberOfRvaAndSizes         As Long
    DataDirectory(0 To 15)      As IMAGE_DATA_DIRECTORY_64
End Type

Public Type IMAGE_NT_HEADERS_32 '---248bytes.
    SignatureLow            As Integer '2 "PE"
    SignatureHigh           As Integer '2
    FileHeader              As IMAGE_FILE_HEADER '20
    OptionalHeader          As IMAGE_OPTIONAL_HEADER_32
End Type

Public Type IMAGE_NT_HEADERS_64 '---264bytes.
    SignatureLow            As Integer '2 "PE"
    SignatureHigh           As Integer '2
    FileHeader              As IMAGE_FILE_HEADER '20
    OptionalHeader          As IMAGE_OPTIONAL_HEADER_64
End Type

Public Type IMAGE_SECTION_HEADER '---40bytes. lokasi struktur header ini berada persis setelah struktur header "IMAGE_NT_HEADERS_32" untuk w32bit atau "IMAGE_NT_HEADERS_64" untuk w64bit.
    SectionName(7)          As Byte '---8 bytes.
    VirtualSize             As Long '---ukuran virtual di memori (bila dieksekusi)
    VirtualAddress          As Long '---alamat virtual default di memori (bila dieksekusi)
    SizeOfRawData           As Long '---ukuran section di file fisik.
    PointerToRawData        As Long '---posisi awal section di file fisik. base 0 alias perhitungan dimulai dari 0, bukan 1.
    PointerToRelocations    As Long
    PointerToLinenumbers    As Long
    NumberOfRelocations     As Integer
    NumberOfLinenumbers     As Integer
    Characteristics         As Long
End Type

Private Type SignPE32
    Signature1            As Integer '2 "PE"
    Signature2            As Integer '2
End Type

' Gambar AH struktur PE biar hapal....
' --------------------------------------
' DOS SIGNATURE -> MZ       = 2 byes
' DOS HEADER                = 62 bytes
' -- Penyekat/DOS STUB ------- = ... bytes (optional)
' NT SIGNATURE -> PE00      = 4 bytes
' FILE STUB                 = 20 bytes
' OPTIONAL HEADER           = 224 bytes
' --------------------------------------
' CUKUP karena udah masuk section :D
Private Const IMAGE_FILE_RELOCS_STRIPPED = &H1                 ' Relocation info stripped from file.
Private Const IMAGE_FILE_EXECUTABLE_IMAGE = &H2                ' File is executable  (i.e. no unresolved externel references).
Private Const IMAGE_FILE_LINE_NUMS_STRIPPED = &H4              ' Line nunbers stripped from file.
Private Const IMAGE_FILE_LOCAL_SYMS_STRIPPED = &H8             ' Local symbols stripped from file.
Private Const IMAGE_FILE_AGGRESIVE_WS_TRIM = &H10              ' Agressively trim working set
Private Const IMAGE_FILE_LARGE_ADDRESS_AWARE = &H20            ' App can handle >2gb addresses
Private Const IMAGE_FILE_BYTES_REVERSED_LO = &H80              ' Bytes of machine word are reversed.
'Private Const IMAGE_FILE_32BIT_MACHINE = &H100                 ' 32 bit word machine.
Private Const IMAGE_FILE_DEBUG_STRIPPED = &H200                ' Debugging info stripped from file in .DBG file
Private Const IMAGE_FILE_REMOVABLE_RUN_FROM_SWAP = &H400       ' If Image is on removable media, copy and run from the swap file.
Private Const IMAGE_FILE_NET_RUN_FROM_SWAP = &H800             ' If Image is on Net, copy and run from the swap file.
Private Const IMAGE_FILE_SYSTEM = &H1000                       ' System File.
Private Const IMAGE_FILE_DLL = &H2000                          ' File is a DLL.
Private Const IMAGE_FILE_UP_SYSTEM_ONLY = &H4000               ' File should only be run on a UP machine
Private Const IMAGE_FILE_BYTES_REVERSED_HI = &H8000            ' Bytes of machine word are reversed.

Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByVal pv6432_lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal pv_lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, ByVal pv_lpNumberOfBytesRead As Long, ByVal pv_lpOverlapped As Long) As Long
Private Declare Sub RtlZeroMemory Lib "ntdll.dll" (ByVal pDestBuffer As Long, ByVal nDestLengthToFillWithZeroBytes As Long) '<---reset isi dst yaitu mengisinya dengan bytenumber = 0.

Dim CFILE   As New classFile

' Buat cek apakah valid PE (64 byte tidak valid)? - Berdasarkan Handle File (Nilai Benar kembali pada e_lfalnew karena eman-eman posisi itu bisa dipakai lagi, biar efisien he9)
Public Function IsValidPE32(hFile As Long) As Long ' jumlah byte yang diakses 88bytes
Dim IDOSH               As IMAGE_DOS_HEADER ' 64
Dim IFH                 As IMAGE_FILE_HEADER '20
Dim SPE                 As SignPE32 '4
Dim RetFunct            As Long
Dim nNumberBytesOpsRet  As Long
    
    Call SetFilePointer(hFile, 0, 0, 0) '---Base0. 'start dari 0
    RetFunct = ReadFile(hFile, VarPtr(IDOSH), Len(IDOSH), VarPtr(nNumberBytesOpsRet), 0) ' base1
    
    If RetFunct = 0 Then
       GoTo LBL_FALSE
    End If
    
    If IDOSH.e_magic <> &H5A4D Then '---"MZ" mungkin udah DOS valid tapi.... cek lagi gak ya :D
       GoTo LBL_FALSE
    End If
    
    Call SetFilePointer(hFile, IDOSH.e_lfanew, 0, 0) '---Base0.
    
    RetFunct = ReadFile(hFile, VarPtr(SPE), Len(SPE), VarPtr(nNumberBytesOpsRet), 0)

    If SPE.Signature1 <> &H4550 Then ' PE
       GoTo LBL_FALSE
    End If
    
    Call SetFilePointer(hFile, IDOSH.e_lfanew + 4, 0, 0) '---Base0. lgsung menuju target (characteristic)
    RetFunct = ReadFile(hFile, VarPtr(IFH), Len(IFH), VarPtr(nNumberBytesOpsRet), 0)

    If (IFH.Characteristics And IMAGE_FILE_32BIT_MACHINE) <> IMAGE_FILE_32BIT_MACHINE Then '---hanya cek di file pe32bit.
        GoTo LBL_FALSE
    End If
        
    If (IFH.Characteristics And IMAGE_FILE_DLL) = IMAGE_FILE_DLL Then ' cek dari DLL dulu
       IsPE32EXE = False
    ElseIf (IFH.Characteristics And IMAGE_FILE_EXECUTABLE_IMAGE) = IMAGE_FILE_EXECUTABLE_IMAGE Then
       IsPE32EXE = True
    Else
       IsPE32EXE = False
    End If

    IsValidPE32 = IDOSH.e_lfanew ' nilainya pasti lebih dari 68

LBL_CLEAN:
    Call RtlZeroMemory(VarPtr(IDOSH), Len(IDOSH))
    Call RtlZeroMemory(VarPtr(SPE), Len(SPE))

Exit Function
LBL_FALSE:
    IsValidPE32 = 0 ' klo tidak valid balik ke 0 atau <= 68
    IsPE32EXE = False
End Function

  
'  return: 0 --> bukan PE yg valid 32bit maupun 64bit.
'  return: 1 --> PE 32 bit untuk *.EXE (default).
'  return: 2 --> PE 32 bit untuk *.DLL (default).
'  return: 3 --> PE 32 bit untuk *.SYS (default).
'  return: 4 --> PE 32 bit untuk lainnya (default).
'  return: 5 --> PE 64 bit untuk *.EXE (default).
'  return: 6 --> PE 64 bit untuk *.DLL (default).
'  return: 7 --> PE 64 bit untuk *.SYS (default).
'  return: 8 --> PE 64 bit untuk lainnya (default).
'  return: 88 --> PE 32 valid tapi struktur tidak diketahui (tidak sesuai standar 32/64 bit).
'  return: 99 --> PE 64 valid tapi struktur tidak diketahui (tidak sesuai standar 32/64 bit).
Public Function GetPE3264Type(ByVal hFile As Long) As Long
Dim IDOSH               As IMAGE_DOS_HEADER ' 64
Dim IMNTH32             As IMAGE_NT_HEADERS_32
Dim IMNTH64             As IMAGE_NT_HEADERS_64
Dim IOPTH64             As IMAGE_OPTIONAL_HEADER_64
Dim nFileLen            As Long
Dim RetFunct            As Long
Dim nNumberBytesOpsRet  As Long
    nFileLen = GetFileSize(hFile, ByVal 0)
    If nFileLen <= Len(IDOSH) Then
        GetPE3264Type = 0
        GoTo LBL_TERAKHIR
    End If
    Call SetFilePointer(hFile, 0, 0, 0)
    RetFunct = ReadFile(hFile, VarPtr(IDOSH), Len(IDOSH), VarPtr(nNumberBytesOpsRet), 0)
    If RetFunct = 0 Then
        GetPE3264Type = 0
        GoTo LBL_TERAKHIR
    End If
    If IDOSH.e_magic <> &H5A4D Then '"MZ"
        GetPE3264Type = 0
        GoTo LBL_FREE_MEMORY
    End If
    If IDOSH.e_lfanew <= 0 Or IDOSH.e_lfanew >= nFileLen Then
        GetPE3264Type = 0
        GoTo LBL_FREE_MEMORY
    End If
    If (IDOSH.e_lfanew + Len(IMNTH32)) >= nFileLen Then '---besar pasak daripada tiang.
        GetPE3264Type = 0
        GoTo LBL_FREE_MEMORY
    End If
    Call SetFilePointer(hFile, IDOSH.e_lfanew, 0, 0)
    RetFunct = ReadFile(hFile, VarPtr(IMNTH32), Len(IMNTH32), VarPtr(nNumberBytesOpsRet), 0)
    If RetFunct = 0 Then
        GetPE3264Type = 0
        GoTo LBL_FREE_MEMORY
    End If
    If IMNTH32.SignatureLow <> &H4550 Then '"PE"
        GetPE3264Type = 0
        GoTo LBL_FREE_MEMORY
    End If
    '---yg berhasil melewati baris ini adalah valid PE. sekarang cari tahu karakteristik (tipe) PE:
    'MsgBox "VALID!"
    If IMNTH32.FileHeader.SizeOfOptionalHeader >= Len(IOPTH64) Then '---kemungkinan besar 64 bit.
        Call SetFilePointer(hFile, IDOSH.e_lfanew, 0, 0)
        RetFunct = ReadFile(hFile, VarPtr(IMNTH64), Len(IMNTH64), VarPtr(nNumberBytesOpsRet), 0)
        If RetFunct = 0 Then
            GetPE3264Type = 99 '??PE 64 valid tapi struktur tidak diketahui (tidak sesuai standar 32/64 bit)
            GoTo LBL_FREE_MEMORY
        End If
LBL_JUMP_TO_64_BIT_CHARACTERISTICS:
        If (IMNTH64.FileHeader.Characteristics And IMAGE_FILE_32BIT_MACHINE) = IMAGE_FILE_32BIT_MACHINE Then
            GoTo LBL_JUMP_TO_32_BIT_CHARACTERISTICS
        ElseIf (IMNTH64.FileHeader.Characteristics And IMAGE_FILE_DLL) = IMAGE_FILE_DLL Then
            GetPE3264Type = 6 'PE 64 bit untuk *.DLL (default).
            GoTo LBL_FREE_MEMORY
        ElseIf (IMNTH64.FileHeader.Characteristics And IMAGE_FILE_SYSTEM) = IMAGE_FILE_SYSTEM Then
            GetPE3264Type = 7 'PE 64 bit untuk *.SYS (default).
            GoTo LBL_FREE_MEMORY
        ElseIf (IMNTH64.FileHeader.Characteristics And IMAGE_FILE_EXECUTABLE_IMAGE) = IMAGE_FILE_EXECUTABLE_IMAGE Then
            GetPE3264Type = 5 'PE 64 bit untuk *.EXE (default).
            GoTo LBL_FREE_MEMORY
        Else
            GetPE3264Type = 8 'PE 64 bit untuk lainnya (default).
            GoTo LBL_FREE_MEMORY
        End If
    Else
LBL_JUMP_TO_32_BIT_CHARACTERISTICS:
        If (IMNTH32.FileHeader.Characteristics And IMAGE_FILE_32BIT_MACHINE) <> IMAGE_FILE_32BIT_MACHINE Then
            GoTo LBL_JUMP_TO_64_BIT_CHARACTERISTICS
        ElseIf (IMNTH32.FileHeader.Characteristics And IMAGE_FILE_DLL) = IMAGE_FILE_DLL Then
            GetPE3264Type = 2 'PE 32 bit untuk *.DLL (default).
            GoTo LBL_FREE_MEMORY
        ElseIf (IMNTH32.FileHeader.Characteristics And IMAGE_FILE_SYSTEM) = IMAGE_FILE_SYSTEM Then
            GetPE3264Type = 3 'PE 32 bit untuk *.SYS (default).
            GoTo LBL_FREE_MEMORY
        ElseIf (IMNTH32.FileHeader.Characteristics And IMAGE_FILE_EXECUTABLE_IMAGE) = IMAGE_FILE_EXECUTABLE_IMAGE Then
            GetPE3264Type = 1 'PE 32 bit untuk *.EXE (default).
            GoTo LBL_FREE_MEMORY
        Else
            GetPE3264Type = 4 'PE 32 bit untuk lainnya (default).
            GoTo LBL_FREE_MEMORY
        End If
    End If
LBL_FREE_MEMORY:
    Call RtlZeroMemory(VarPtr(IDOSH), Len(IDOSH))
    Call RtlZeroMemory(VarPtr(IMNTH32), Len(IMNTH32))
    Call RtlZeroMemory(VarPtr(IMNTH64), Len(IMNTH64))
    
LBL_TERAKHIR:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Function

' Klo yang ini dipakai untuk mendapatkan ukuran PE32 asli, tapi tidak pada kondisi berulang
Public Function GetRealSizePE32(sFile As String, ByRef SizeInFisik As Long) As Long
Dim fReturn                 As Long
Dim nFileLen                As Long
Dim hFile                   As Long
Dim nNumberBytesOpsRet      As Long
Dim iCount                  As Long
Dim IDOSH                   As IMAGE_DOS_HEADER
Dim INTH32                  As IMAGE_NT_HEADERS_32
Dim ISECH()                 As IMAGE_SECTION_HEADER
Dim nSection                As Long
Dim BiggestSectionOff       As Long
Dim SectionToSize           As Long
Dim AddNewHeaderBase0       As Long

LBL_LOCAL_INITIALIZATION:
Dim CFILE       As classFile '---class file ini ada di proyek cms kemarin.
    Set CFILE = New classFile

LBL_OBJECT_TRY_OPEN:
    hFile = CFILE.VbOpenFile(sFile, FOR_BINARY_ACCESS_READ, LOCK_NONE)
    If hFile <= 0 Then
        GoTo LBL_TERAKHIR
    End If
    nFileLen = CFILE.VbFileLen(hFile)

LBL_VALID:
    If nFileLen <= Len(IDOSH) Then '---nggak memeriksa file yg berukuran kurang dari atau sama dengan dosheader atau lebih besar dari 2GB.
         GoTo LBL_TERAKHIR
    End If

LBL_READ:
    Call SetFilePointer(hFile, 0, 0, 0) 'jaga-jaga
    fReturn = ReadFile(hFile, VarPtr(IDOSH), Len(IDOSH), VarPtr(nNumberBytesOpsRet), 0)
    If fReturn = 0 Then
       GoTo LBL_TERAKHIR
    End If
    If IDOSH.e_magic <> &H5A4D Then '---"MZ"="Mark Zbikowski",singkatan nama orang.
       GoTo LBL_TERAKHIR
    End If
    If (IDOSH.e_lfanew >= (nFileLen - Len(INTH32))) Or (IDOSH.e_lfanew <= 0) Then '--lokasi peheader harus berada di dalam file.
       GoTo LBL_TERAKHIR
    End If
    AddNewHeaderBase0 = IDOSH.e_lfanew
'---cek ntheader:
    Call SetFilePointer(hFile, IDOSH.e_lfanew, 0, 0) '---Base0.
    fReturn = ReadFile(hFile, VarPtr(INTH32), Len(INTH32), VarPtr(nNumberBytesOpsRet), 0)
    If fReturn = 0 Then
       GoTo LBL_TERAKHIR
    End If
    
    If INTH32.SignatureLow <> &H4550 Then '---"PE"="Portable Executable",singkatan istilah.
        GoTo LBL_TERAKHIR
    End If
    
    If (INTH32.FileHeader.Characteristics And IMAGE_FILE_32BIT_MACHINE) <> IMAGE_FILE_32BIT_MACHINE Then '---hanya cari di file pe32bit.
        GoTo LBL_TERAKHIR
    End If
    
    If INTH32.FileHeader.NumberOfSections <= 0 Then '---lho,kok nggak ada section-nya,sih?
        GoTo LBL_TERAKHIR
    End If
 '---cek sectionheader:
  nSection = INTH32.FileHeader.NumberOfSections
  ReDim ISECH(nSection - 1) As IMAGE_SECTION_HEADER
  Call SetFilePointer(hFile, AddNewHeaderBase0 + Len(INTH32), 0, 0) '---Base0. INTH32=248 Bytes, set pointernya
  fReturn = ReadFile(hFile, VarPtr(ISECH(0)), Len(ISECH(0)) * nSection, VarPtr(nNumberBytesOpsRet), 0) ' yang akan dibaca ukuran type section (40bytes) x jumlah section
  For iCount = 0 To nSection - 1
       If iCount > 0 Then
          If ISECH(iCount).PointerToRawData > BiggestSectionOff Then
             BiggestSectionOff = ISECH(iCount).PointerToRawData ' biasanya section terakhir
             SectionToSize = ISECH(iCount).SizeOfRawData
          End If
       Else
            BiggestSectionOff = ISECH(iCount).PointerToRawData ' awalnya baygkan terbesar ada yang pertama
            SectionToSize = ISECH(iCount).SizeOfRawData
       End If
   Next
   

   'Ketemu ukuran Real dari EXE :D
   GetRealSizePE32 = BiggestSectionOff + SectionToSize
   SizeInFisik = nFileLen ' ukuran di HD

            
LBL_CLEAN_MEMORY:
    Erase ISECH()
    Call RtlZeroMemory(VarPtr(IDOSH), Len(IDOSH))
    Call RtlZeroMemory(VarPtr(INTH32), Len(INTH32))
    CFILE.VbCloseFile hFile

Exit Function
LBL_TERAKHIR:
    GetRealSizePE32 = 0
    SizeInFisik = 0
    CFILE.VbCloseFile hFile
End Function


' Mencari jumlah muatan byte tambahn di exe/dll, dipakai idak pada perulangan
Public Function DeteksiMuatanPE32(sFile As String) As Long
Dim RealSz   As String
Dim FisikSz  As Long

RealSz = GetRealSizePE32(sFile, FisikSz)

If RealSz > 0 Then
   DeteksiMuatanPE32 = FisikSz - RealSz ' bisa minus (artinya korupt PE)
Else
   DeteksiMuatanPE32 = 0 ' gagal anggap gak ada muatan
End If
End Function

