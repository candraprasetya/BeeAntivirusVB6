Attribute VB_Name = "basMPG"
Public Function CekMPG(sPath As String) As String
Static OutDat() As Byte
Static sData As String
CekMPG = vbNullString
If LCase$(Right$(sPath, 3)) <> "mpg" Then Exit Function

Call ReadUnicodeFile2(hFile, 1, 70, OutDat)
sData = StrConv(OutDat, vbUnicode)
Erase OutDat

End Function
