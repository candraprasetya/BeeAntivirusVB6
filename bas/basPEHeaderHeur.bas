Attribute VB_Name = "basPEHeaderHeur"

Public Function CekHeaderPE(xLastSect As String, xNLastSect As String) As String

xNLastSect = LCase$(xNLastSect)
If xNLastSect = "" Then Exit Function
If InStr(xNLastSect, "dat") > 0 Or InStr(xNLastSect, "text") > 0 Then
    If xLastSect = "E0000020" Then
        CekHeaderPE = "Win32/Sality"
        Exit Function
    End If
End If
If Len(xNLastSect) > 5 Then
If InStr(xNLastSect, "data") > 0 Then
If Left(xNamaSectionAkhir2, 1) <> "." Then
    If xLastSect = "40000040" Then
    If xLastSect = xSectionAkhir2 Then
        CekHeaderPE = "Win32/Mabezat"
        Exit Function
    End If
    End If
End If
End If
End If
If InStr(xNLastSect, "text") > 0 Then
If xSectionJum <> 1 Then
    If xLastSect = "E0000020" Then
        CekHeaderPE = "Win32/Ramnit.G"
        Exit Function
    End If
End If
End If
If Left(xNLastSect, 1) <> "." Then
    If InStr(xNamaSectionAkhir2, "data") > 0 Or InStr(xNamaSectionAkhir2, "rsrc") > 0 Then
        If Right(xSectionAkhir2, 2) = 60 Then
            If xLastSect = "E0000060" Then
                CekHeaderPE = "Win32/Virut"
                Exit Function
            End If
        End If
    End If
End If
If InStr(xNLastSect, "reloc") > 0 Then
    If xLastSect = "E2000060" Then
        CekHeaderPE = "Win32/Virut"
        Exit Function
    End If
End If
If InStr(xNLastSect, "rsrc") > 0 Then
    If xLastSect = "E0000060" Then
        CekHeaderPE = "Win32/Virut"
        Exit Function
    End If
End If
If InStr(xNLastSect, "idata") > 0 Then
    If xLastSect = "E00000E0" Then
        CekHeaderPE = "Win32/Troxa.B"
        Exit Function
    End If
End If
CekHeaderPE = vbNullString
End Function

