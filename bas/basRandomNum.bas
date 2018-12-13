Attribute VB_Name = "basRandomNum"
Option Explicit

Dim unique As Boolean
Dim uniqueNumber As Boolean
Dim sepbyComma As Boolean
Dim yesUcase As Boolean

Public Function GetRandomPassword(ByVal length As Integer) As String
    Dim x, n: Dim stmp As String: Const s_chars As String = _
    "abcdefghijklmnopqrstuvwxyzADCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    On Local Error GoTo ErrorRandom
    unique = True
    yesUcase = True
    sepbyComma = False
    uniqueNumber = False
    If uniqueNumber Then
        x = RandomNumbers(100, 1, length)
        For n = LBound(x) To UBound(x)
            If sepbyComma Then stmp = stmp & x(n) & "-" Else stmp = stmp & x(n)
        Next n
    Else
        x = RandomNumbers(Len(s_chars), 1, length)
        For n = LBound(x) To UBound(x)
            If sepbyComma Then stmp = stmp + Mid$(s_chars, x(n), Len(length)) & "-" _
            Else stmp = stmp + Mid$(s_chars, x(n), Len(length))
        Next n
    End If
    
    If uniqueNumber Then
        If sepbyComma Then GetRandomPassword = Mid$(stmp, 1, Len(stmp) - 1) Else GetRandomPassword = stmp
    Else
        GetRandomPassword = Mid$(stmp, 1, Len(stmp) - length)
        If Mid$(GetRandomPassword, Len(GetRandomPassword), Len(GetRandomPassword)) = "-" Then
            GetRandomPassword = Mid$(stmp, 1, Len(stmp) - length - 1)
        End If
        If yesUcase Then GetRandomPassword = UCase(GetRandomPassword)
    End If
    
Exit Function
ErrorRandom:
        GetRandomPassword = "Error #" & Err.Number
    Err.Clear
End Function

Private Function RandomNumbers(Upper As Integer, Optional Lower As Integer = 1, Optional HowMany As Integer = 1) As Variant
    On Error GoTo LocalError
        If HowMany > ((Upper + 1) - (Lower - 1)) Then Exit Function
    Dim x As Integer
    Dim n As Integer
    Dim arrNums() As Variant
    Dim colNumbers As New Collection

    ReDim arrNums(HowMany - 1)
    With colNumbers
        ' .... First populate the collection
            For x = Lower To Upper
                .Add x
            Next x
        For x = 0 To HowMany - 1
            n = RandomNumber(0, colNumbers.Count + 1)
            arrNums(x) = colNumbers(n)
                If unique Then
                    colNumbers.Remove n
                End If
        Next x
    End With
    Set colNumbers = Nothing
    RandomNumbers = arrNums
Exit Function
LocalError:
    RandomNumbers = ""
    Err.Clear
End Function

Private Function RandomNumber(Upper As Integer, Lower As Integer) As Integer
    Randomize
    RandomNumber = Int((Upper - Lower + 1) * Rnd + Lower)
End Function


