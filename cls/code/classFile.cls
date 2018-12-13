VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
';+------------------------------------------------------------------+;
';Modul untuk membaca file standar fungsi API, via Kernel32.dll
';karena VB6 tidak mendukung pengoperasian file secara unicode -
';dalam hal akses alamat filenya.
';Bisa dianggap sebagai "simulasi" instruksi vb6: Open ... As #n
';Bukan yang terbaik, tapi relatif lebih stabil diantara cara-cara -
';yang lainnya.
';Pengecualian:
';  ~tidak bisa membaca + menulis file yang dikunci oleh proses lain.
';  ~...
';+------------------------------------------------------------------+;
';PMZ-Indonesia, 24-25@Agustus@2008.
';+------------------------------------------------------------------+;

Private Const MAX_PATH      As Long = 260               '00-FF
Private Const MAX_BUFFER    As Long = (MAX_PATH * 2)    '00 00 - FF FF

Private Const SYNCHRONIZE = &H100000    'penting! sinkronisasi data dan akses dengan proses lain.
Private Const READ_CONTROL = &H20000    'penting! ijin untuk mengoperasikan file.
Private Const FILE_READ_DATA = (&H1)    'penting! operasi: membaca file.
Private Const FILE_WRITE_DATA = (&H2)   'penting! operasi: menulis file.

Private Const FILE_SHARE_READ = &H1     'dapat diakses baca oleh proses lain.
Private Const FILE_SHARE_WRITE = &H2    'dapat diakses tulis oleh proses lain.
Private Const FILE_SHARE_DELETE = &H4   'dapat diakses hapus oleh proses lain.

Private Const FILE_ATTRIBUTE_NORMAL = &H80 'untuk file standar.

'operasi alternatif untuk file yang akan dibuat ataupun dibuka:
Private Const FILE_DISPOSE_CREATE_NEW = 1           'hanya akan membuat file baru. bila file sudah ada sebelumnya, fungsi gagal.
Private Const FILE_DISPOSE_CREATE_ALWAYS = 2        'hapus file yang lama (bila ada), dan akan membuat file yang baru.
Private Const FILE_DISPOSE_OPEN_EXISTING = 3        'hanya akan membuka file yang sudah ada, bila file tidak ada, fungsi gagal.
Private Const FILE_DISPOSE_OPEN_ALWAYS = 4          'membuka file yang ada (bila ada), dan akan membuat file yang baru bila file belum ada.
Private Const FILE_DISPOSE_TRUNCATE_EXISTING = 5    'membuka file yang sudah ada, dan menghapus semua isinya terlebih dahulu. fungsi gagal bila file tidak ada.

'membuka file:
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFileA Lib "kernel32" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

'mencari ukuran file:
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

'menggeser posisi pointer ke file:
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long

'operasi dasar untuk file yang telah dibuka:
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long

'menutup file yang telah dibuka:
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'\\Ingat!:
Public Enum CREATE_ACCESS_OPTIONS
    FOR_BINARY_ACCESS_READ = 1                  'hanya membaca isi dari file, tanpa memodifikasi isi file. bila file tidak ada, fungsi gagal.
    FOR_BINARY_ACCESS_WRITE = 2                 'hanya menulis isi ke file, bila file belum ada, akan dibuatkan file baru.
    FOR_BINARY_ACCESS_READ_WRITE = 3            'untuk membaca dan menulis file, bila file belum ada, akan dibuatkan file baru.
    FOR_BINARY_ACCESS_READ_WRITE_EXISTING = 4   'untuk membaca dan menulis file, bila file belum ada, fungsi akan gagal. lebih aman.
End Enum

'\\Ingat!:
Public Enum SHARE_ACCESS_OPTIONS
    LOCK_READ_WRITE = 1 'hanya "sharing delete access", untuk proses yang lain.
    LOCK_READ = 2       'hanya "sharing write + delete access", untuk proses yang lain.
    LOCK_WRITE = 3      'hanya "sharing read + delete access", untuk proses yang lain.
    LOCK_NONE = 4       '"sharing" semuanya, tanpa terkecuali.
End Enum

Public Function VbOpenFile(ByVal szFileName As String, ByVal opCreateOption As CREATE_ACCESS_OPTIONS, ByVal opShareAccess As SHARE_ACCESS_OPTIONS) As Long
On Error Resume Next 'memberi nomor handle ke file bila berhasil, 0 jika gagal.
Dim KeResult        As Long 'result dari kernel32.
Dim KeCreateAccess  As Long
Dim KeCreateOption  As Long
Dim KeShareAccess   As Long
    
    Select Case opCreateOption
        Case FOR_BINARY_ACCESS_READ                 '1
            KeCreateAccess = SYNCHRONIZE Or READ_CONTROL Or FILE_READ_DATA
            KeCreateOption = FILE_DISPOSE_OPEN_EXISTING
        Case FOR_BINARY_ACCESS_WRITE                '2
            KeCreateAccess = SYNCHRONIZE Or READ_CONTROL Or FILE_WRITE_DATA
            KeCreateOption = FILE_DISPOSE_OPEN_ALWAYS
        Case FOR_BINARY_ACCESS_READ_WRITE           '3
            KeCreateAccess = SYNCHRONIZE Or READ_CONTROL Or FILE_READ_DATA Or FILE_WRITE_DATA
            KeCreateOption = FILE_DISPOSE_OPEN_ALWAYS
        Case FOR_BINARY_ACCESS_READ_WRITE_EXISTING  '4
            KeCreateAccess = SYNCHRONIZE Or READ_CONTROL Or FILE_READ_DATA Or FILE_WRITE_DATA
            KeCreateOption = FILE_DISPOSE_OPEN_EXISTING
    End Select
    
    Select Case opShareAccess
        Case LOCK_READ_WRITE    '1
            KeShareAccess = FILE_SHARE_DELETE
        Case LOCK_READ          '2
            KeShareAccess = FILE_SHARE_WRITE Or FILE_SHARE_DELETE
        Case LOCK_WRITE         '3
            KeShareAccess = FILE_SHARE_READ Or FILE_SHARE_DELETE
        Case LOCK_NONE          '4
            KeShareAccess = FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE
    End Select
    
    'pakai cara utama (unicode):
    KeResult = CreateFileW(StrPtr(szFileName), KeCreateAccess, KeShareAccess, ByVal 0, KeCreateOption, FILE_ATTRIBUTE_NORMAL, 0)
    If KeResult > 0 Then        'sukses pakai cara unicode.
        VbOpenFile = KeResult 'masukkan ke fungsi (return): nomor handle menuju ke file.
        GoTo TERAKHIR
    End If
    
    'pakai cara cadangan (ansi):
    KeResult = 0                'reset, sekarang coba pakai ansi:
    KeResult = CreateFileA(szFileName, KeCreateAccess, KeShareAccess, ByVal 0, KeCreateOption, FILE_ATTRIBUTE_NORMAL, 0)
    If KeResult > 0 Then        'sukses pakai cara ansi.
        VbOpenFile = KeResult  'masukkan ke fungsi (return): nomor handle menuju ke file.
        GoTo TERAKHIR
    End If
    
    VbOpenFile = 0             'gagal membuka file :(
    
TERAKHIR:
    If Err.Number > 0 Then
        Err.Clear
    End If
End Function

Public Function VbFileLen(ByVal nFileHandle As Long) As Long
On Error Resume Next 'memberi nilai angka sebesar ukuran file dalam bytes.
    VbFileLen = GetFileSize(nFileHandle, 0)
TERAKHIR:
    If Err.Number > 0 Then
        Err.Clear
    End If
End Function

Public Function VbCloseFile(ByVal nFileHandle As Long) As Long
On Error Resume Next 'memberi nilai 1 jika berhasil, 0 jika gagal.
    VbCloseFile = CloseHandle(nFileHandle)
TERAKHIR:
    If Err.Number > 0 Then
        Err.Clear
    End If
End Function

Public Function VbReadFileB(ByVal nFileHandle As Long, ByVal nStartPos As Long, ByVal nReadLength As Long, ByRef OutFileData() As Byte) As Long
On Error Resume Next 'memberi isi ukuran file (buffer) dalam bytes sebagai pengembalian (return) + isi buffer.

Erase OutFileData                   'reset memori data.

'lanjut yang baru:
Dim nTrueLen        As Long
    nTrueLen = GetFileSize(nFileHandle, 0) 'cari ukuran filenya.
Dim nRequestStart   As Long
Dim nRequestLen     As Long
Dim nApproxLen      As Long
Dim KeResult        As Long

    'optimisasi opsional, dapat diganti sesuai keinginan:
    If nTrueLen <= -1 Then
        VbReadFileB = -1 'error: file tidak ada.
        GoTo TERAKHIR
    ElseIf nTrueLen = 0 Then
        VbReadFileB = -2 'error: file isi kosong.
        GoTo TERAKHIR
    End If
    If nStartPos > nTrueLen Then
        VbReadFileB = -3 'error: start melebihi akhir.
        GoTo TERAKHIR
    End If
    If nStartPos <= 0 Then
        VbReadFileB = -4 'error: start pointer tidak sesuai.
        GoTo TERAKHIR
    End If
    If nReadLength <= 0 Then
        VbReadFileB = -5 'error: panjang yang diminta tidak sesuai.
        GoTo TERAKHIR
    End If
    If nReadLength > nTrueLen Then
        VbReadFileB = -6 'error: panjang yang diminta melebihi akhir.
        GoTo TERAKHIR
    End If
    
    nRequestStart = nStartPos 'start pointer ke data (base 1).
    nRequestLen = nReadLength 'panjang data.
    
    ReDim OutFileData(nRequestLen - 1) As Byte 'persiapkan buffer data (base 0).
    
    SetFilePointer nFileHandle, (nRequestStart - 1), 0, 0 'set start pointer ke handle (base 0).
    KeResult = ReadFile(nFileHandle, OutFileData(0), nRequestLen, nApproxLen, ByVal 0)
    
    If nApproxLen <> nRequestLen Then 'test ukuran buffer dengan isi datanya.
        ReDim Preserve OutFileData(nApproxLen - 1) As Byte 'sesuaikan lagi ukuran buffer (base 0).
    End If
    
    VbReadFileB = nApproxLen  '<-- beritahu jumlah bytes yang berhasil dibaca.
       
TERAKHIR:
    If Err.Number > 0 Then
        Err.Clear
    End If
End Function

Public Function VbWriteFileB(ByVal nFileHandle As Long, ByVal nStartPos As Long, ByVal nWriteLength As Long, ByRef InToFileData() As Byte) As Long
On Error Resume Next 'memberi isi ukuran file (buffer) dalam bytes sebagai pengembalian (return) + isi buffer.
'nStartPos=Base1--->startpos untuk posisi di dalam file, bukan posisi di dalam byte array!
'lanjut yang baru:
Dim nTrueLen        As Long
    nTrueLen = GetFileSize(nFileHandle, 0) 'cari ukuran filenya.
Dim nApproxLen      As Long
Dim KeResult        As Long
    '---filter (umum):
    If nStartPos <= 0 Then '---nStartPos=Base1.
        VbWriteFileB = -1 'error: start pointer tidak sesuai.
        GoTo TERAKHIR
    End If
    If nWriteLength <= 0 Then
        VbWriteFileB = -2 'error: panjang data tidak sesuai.
        GoTo TERAKHIR
    End If
    If nWriteLength > (UBound(InToFileData()) + 1) Then
        VbWriteFileB = -3 'error: panjang data tidak sesuai.
        GoTo TERAKHIR
    End If
    '---tambahkan filter di sini:
    '...
    '---set posisi penulisan:
    SetFilePointer nFileHandle, (nStartPos - 1), 0, 0 'set start pointer ke handle (base 0).
    '---menulis:
    KeResult = WriteFile(nFileHandle, InToFileData(0), nWriteLength, nApproxLen, ByVal 0)
    
    VbWriteFileB = nApproxLen  '<-- beritahu jumlah bytes yang berhasil ditulis.
       
TERAKHIR:
    If Err.Number > 0 Then
        Err.Clear
    End If
End Function

Private Sub Class_Initialize()
On Error Resume Next
    DoEvents
TERAKHIR:
    If Err.Number > 0 Then
        Err.Clear
    End If
End Sub

Private Sub Class_Terminate()
On Error Resume Next
    DoEvents
TERAKHIR:
    If Err.Number > 0 Then
        Err.Clear
    End If
End Sub




