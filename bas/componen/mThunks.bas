Attribute VB_Name = "mThunks"
Option Explicit

Public Enum eThunk
    tnkCommonDialogProc = 101
    tnkSubclassProc = 102
    tnkHookProc = 103
    tnkTimerProc = 104
    tnkHookClientProc = 105
    tnkTimerClientProc = 106
    tnkRichEditProc = 107
    tnkLVCompareProc = 108
End Enum

Public Function Thunk_Alloc(ByVal iThunk As eThunk, Optional ByVal iIDEPatchOffset As Long = -1&) As Long
    Dim lyBytes()      As Byte:  lyBytes = LoadResData(iThunk, "ASM")
    Dim liLBound       As Long:   liLBound = LBound(lyBytes)
    Dim liSize         As Long:     liSize = UBound(lyBytes) - liLBound + 1&
    
    Thunk_Alloc = MemAlloc(liSize)
    If Thunk_Alloc Then
        CopyMemory ByVal Thunk_Alloc, lyBytes(liLBound), liSize
        If iIDEPatchOffset > -1& Then
            If InIDE() Then MemWord(ByVal UnsignedAdd(Thunk_Alloc, iIDEPatchOffset)) = &H9090
        End If
    End If
End Function

Public Sub Thunk_Patch(ByVal pThunk As Long, ByVal iOffset As Long, ByVal iValue As Long)
    MemLong(ByVal UnsignedAdd(pThunk, iOffset)) = iValue
End Sub

Public Sub Thunk_PatchFuncAddr(ByVal pThunk As Long, ByVal iOffset As Long, ByRef sLib As String, ByRef sName As String)
    Dim lsAnsiLib       As String:    lsAnsiLib = StrConv(sLib & vbNullChar, vbFromUnicode)
    Dim lsAnsiName      As String:   lsAnsiName = StrConv(sName & vbNullChar, vbFromUnicode)
    Dim lpPatch         As Long:        lpPatch = UnsignedAdd(pThunk, iOffset)
    Dim liAddr          As Long:         liAddr = GetProcAddress(GetModuleHandle(ByVal StrPtr(lsAnsiLib)), ByVal StrPtr(lsAnsiName))
    'debug.assert liAddr
    MemLong(ByVal lpPatch) = liAddr - lpPatch - 4&
End Sub

Private Function InIDE() As Boolean
    'debug.assert pTrue(InIDE)
End Function

Private Function pTrue(ByRef B As Boolean) As Boolean
    B = True
    pTrue = B
End Function
