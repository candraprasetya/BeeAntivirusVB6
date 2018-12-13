Attribute VB_Name = "basPE"
' ALL ABOUT PE
' I Love PE


Public Type LARGE_INTEGER
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

Public Type IMAGE_NT_HEADERS '---panjang struktur untuk signature saja : 4 bytes.
    SignatureLow                As Integer  'WORD[2]        : Signature untuk PE 32 dan 64 bit adalah &H4550 = "PE"
    SignatureHigh               As Integer  'WORD[2]        : Signature tambahan untuk PE 32 dan 64 bit (biasanya) &H0000 = Chr$(0) & Chr$(0)
    '---dilanjutkan dengan FileHeader                   (=IMAGE_FILE_HEADER).
    '---dilanjutkan dengan OptionalHeader               (=IMAGE_OPTIONAL_HEADER_32 / IMAGE_OPTIONAL_HEADER_64).
    '---dilanjutkan dengan array dari SectionHeaders    (=IMAGE_SECTION_HEADER).
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


' --------------------------------------
' DOS SIGNATURE -> MZ       = 2 byes
' DOS HEADER                = 62 bytes
' -- Penyekat/DOS STUB ------- = ... bytes (optional)
' NT SIGNATURE -> PE00      = 4 bytes
' FILE STUB                 = 20 bytes
' OPTIONAL HEADER           = 224 bytes
' --------------------------------------


Public Const IMAGE_FILE_RELOCS_STRIPPED = &H1                 ' Relocation info stripped from file.
Public Const IMAGE_FILE_EXECUTABLE_IMAGE = &H2                ' File is executable  (i.e. no unresolved externel references).
Public Const IMAGE_FILE_LINE_NUMS_STRIPPED = &H4              ' Line nunbers stripped from file.
Public Const IMAGE_FILE_LOCAL_SYMS_STRIPPED = &H8             ' Local symbols stripped from file.
Public Const IMAGE_FILE_AGGRESIVE_WS_TRIM = &H10              ' Agressively trim working set
Public Const IMAGE_FILE_LARGE_ADDRESS_AWARE = &H20            ' App can handle >2gb addresses
Public Const IMAGE_FILE_BYTES_REVERSED_LO = &H80              ' Bytes of machine word are reversed.
Public Const IMAGE_FILE_32BIT_MACHINE = &H100                 ' 32 bit word machine.
Public Const IMAGE_FILE_DEBUG_STRIPPED = &H200                ' Debugging info stripped from file in .DBG file
Public Const IMAGE_FILE_REMOVABLE_RUN_FROM_SWAP = &H400       ' If Image is on removable media, copy and run from the swap file.
Public Const IMAGE_FILE_NET_RUN_FROM_SWAP = &H800             ' If Image is on Net, copy and run from the swap file.
Public Const IMAGE_FILE_SYSTEM = &H1000                       ' System File.
Public Const IMAGE_FILE_DLL = &H2000                          ' File is a DLL.
Public Const IMAGE_FILE_UP_SYSTEM_ONLY = &H4000               ' File should only be run on a UP machine
Public Const IMAGE_FILE_BYTES_REVERSED_HI = &H8000            ' Bytes of machine word are reversed.

Dim CFILE   As New classFile
