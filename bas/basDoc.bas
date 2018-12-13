Attribute VB_Name = "basDoc"
Public Function isInfectedDoc(sFileExe As String) As Boolean
Static IsiFile As String
Dim dR() As Byte
Call ReadUnicodeFile2(GetHandleFile(sFileExe), 1, FileLen(sFileExe), dR)
IsiFile = StrConv(dR, vbUnicode)
If Left(IsiFile, 2) = "MZ" And InStr(IsiFile, DocHeader) > 0 Then
    isInfectedDoc = True
Else
    isInfectedDoc = False
End If
End Function

Public Function HealDoc(sFileExe As String, sTarget As String)
Static iPointer As Long
Static IsiFile As String
Dim dR() As Byte
Call ReadUnicodeFile2(GetHandleFile(sFileExe), 1, FileLen(sFileExe), dR)
IsiFile = StrConv(dR, vbUnicode)
iPointer = InStr(IsiFile, DocHeader)
IsiFile = Mid(IsiFile, iPointer)
MsgBox IsiFile
WriteFileUniSim sTarget, IsiFile
End Function
