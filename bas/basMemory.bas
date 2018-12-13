Attribute VB_Name = "basMemory"
Public Type LARGE_INTEGER
LowPart As Long
HighPart As Long
End Type

Public Type MEMORYSTATUSEX
dwLength As Long
dwMemoryLoad As Long
ullTotalPhys As LARGE_INTEGER
ullAvailPhys As LARGE_INTEGER
ullTotalPageFile As LARGE_INTEGER
ullAvailPageFile As LARGE_INTEGER
ullTotalVirtual As LARGE_INTEGER
ullAvailVirtual As LARGE_INTEGER
ullAvailExtendedVirtual As LARGE_INTEGER
End Type

Declare Function GlobalMemoryStatusEx Lib "kernel32.dll" (ByRef lpBuffer As MEMORYSTATUSEX) As Long

Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Function MemoryAvailable() As Long
Dim MemStat As MEMORYSTATUSEX
Dim TotalPhys As Currency

On Error Resume Next

'initialize structure
MemStat.dwLength = Len(MemStat)

'retireve memory information
GlobalMemoryStatusEx MemStat

'convert large integer to currency
TotalPhys = LargeIntToCurrency(MemStat.ullTotalPhys)

MemoryAvailable = Int(TotalPhys / (1024 ^ 2))
End Function

Private Function LargeIntToCurrency(liInput As LARGE_INTEGER) As Currency
On Error Resume Next

'copy 8 bytes from the large integer to an empty currency
CopyMemory LargeIntToCurrency, liInput, LenB(liInput)
'adjust it
LargeIntToCurrency = LargeIntToCurrency * 10000
End Function


