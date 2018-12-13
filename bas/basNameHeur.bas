Attribute VB_Name = "basNameHeur"

Public Function ChkNamVir(xPath As String, xFsize As String) As String
Static IsiFile As String
xPath = UCase$(xPath)
If isProperFile(xPath, "VMX") = True Then
IsiFile = ReadUnicodeFile(xPath)
If Left(IsiFile, 2) <> "MZ" Then ChkNamVir = "": Exit Function
    If InStr(xPath, "JWGKVSQ") > 0 Then
        If xFsize > 150000 Then
        ChkNamVir = "Conficker.Variants[Wrm]"
        Exit Function
        End If
    End If
    If InStr(xPath, "JWG") > 0 Then
        If xFsize > 150000 Then
        ChkNamVir = "Conficker.Variants[Trj]"
        Exit Function
        End If
    End If
    If InStr(xPath, "KVSQ") > 0 Then
        If xFsize > 150000 Then
        ChkNamVir = "Conficker.Variants[Vrt]"
        Exit Function
        End If
    End If
End If
ChkNamVir = ""
End Function
