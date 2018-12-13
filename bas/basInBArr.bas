Attribute VB_Name = "basInBArr"
Option Explicit
Public Function InBArr(ByRef ByteArray() As Byte, ByRef KeyWord As String, Optional StartPos As Long = 0, Optional Compare As Byte = vbBinaryCompare) As Long
    Static KeyBuffer() As Byte, KeyBufferU() As Byte, KeyPtr As Long, OldCompare As Byte
    Dim A As Long, B As Long, C As Long, KeyLen As Long, KeyUpper As Long
    Dim FirstKeyByte As Byte, LastKeyByte As Byte, TempByte As Byte
    Dim FirstKeyByteU As Byte, LastKeyByteU As Byte
    If KeyWord = vbNullString Then InBArr = -1: Exit Function
    KeyLen = StrPtr(KeyWord)
    If Not (KeyPtr = KeyLen And Compare = OldCompare) Then
        KeyPtr = KeyLen
        OldCompare = Compare
        If Compare = vbBinaryCompare Then
            KeyBuffer = KeyWord
        Else
            KeyBufferU = UCase$(KeyWord)
            KeyBuffer = LCase$(KeyWord)
        End If
    End If
    KeyLen = UBound(KeyBuffer) - 1
    KeyUpper = KeyLen \ 2
    If KeyUpper > UBound(ByteArray) Then InBArr = -1: Exit Function
    If StartPos < LBound(ByteArray) Then StartPos = LBound(ByteArray)
    If StartPos > UBound(ByteArray) - KeyUpper Then StartPos = UBound(ByteArray) - KeyUpper
    FirstKeyByte = KeyBuffer(0)
    LastKeyByte = KeyBuffer(UBound(KeyBuffer) - 1)
    If Compare = vbBinaryCompare Then
        'loop through the array
        For A = StartPos To UBound(ByteArray) - KeyUpper
            If ByteArray(A) = FirstKeyByte Then
                If ByteArray(A + KeyUpper) = LastKeyByte Then
                    If KeyLen > 4 Then
                        'check if keyword is found from the array
                        C = A + 1
                        For B = 2 To KeyLen Step 2
                            If Not (ByteArray(C) = KeyBuffer(B)) Then Exit For
                            C = C + 1
                        Next B
                        'keyword is found!
                        If B > KeyLen Then
                            InBArr = A
                            Exit Function
                        End If
                    Else
                        InBArr = A
                        Exit Function
                    End If
                End If
            End If
        Next A
    Else 'vbTextCompare
        FirstKeyByteU = KeyBufferU(0)
        LastKeyByteU = KeyBufferU(UBound(KeyBuffer) - 1)
        'loop through the array
        For A = StartPos To UBound(ByteArray) - KeyUpper
            TempByte = ByteArray(A)
            If TempByte = FirstKeyByte Or TempByte = FirstKeyByteU Then
                TempByte = ByteArray(A + KeyUpper)
                If TempByte = LastKeyByte Or TempByte = LastKeyByteU Then
                    If KeyLen > 4 Then
                        'check if keyword is found from the array
                        C = A + 1
                        For B = 2 To KeyLen Step 2
                            TempByte = ByteArray(C)
                            If Not (TempByte = KeyBuffer(B) Or TempByte = KeyBufferU(B)) Then Exit For
                            C = C + 1
                        Next B
                        'keyword is found!
                        If B > KeyLen Then
                            InBArr = A
                            Exit Function
                        End If
                    Else
                        InBArr = A
                        Exit Function
                    End If
                End If
            End If
        Next A
    End If
    InBArr = -1
End Function
Public Function InBArrRev(ByRef ByteArray() As Byte, ByRef KeyWord As String, Optional StartPos As Long = -1, Optional Compare As Byte = vbBinaryCompare) As Long
    Static KeyBuffer() As Byte, KeyBufferU() As Byte, KeyPtr As Long, OldCompare As Byte
    Dim A As Long, B As Long, C As Long, KeyLen As Long, KeyUpper As Long
    Dim FirstKeyByte As Byte, LastKeyByte As Byte, TempByte As Byte
    Dim FirstKeyByteU As Byte, LastKeyByteU As Byte
    If KeyWord = vbNullString Then InBArrRev = -1: Exit Function
    KeyLen = StrPtr(KeyWord)
    If Not (KeyPtr = KeyLen And Compare = OldCompare) Then
        KeyPtr = KeyLen
        OldCompare = Compare
        If Compare = vbBinaryCompare Then
            KeyBuffer = KeyWord
        Else
            KeyBufferU = UCase$(KeyWord)
            KeyBuffer = LCase$(KeyWord)
        End If
    End If
    KeyLen = UBound(KeyBuffer) - 1
    KeyUpper = KeyLen \ 2
    If KeyUpper > UBound(ByteArray) Then InBArrRev = -1: Exit Function
    If StartPos < 0 Or StartPos > UBound(ByteArray) - KeyUpper Then StartPos = UBound(ByteArray) - KeyUpper
    FirstKeyByte = KeyBuffer(0)
    LastKeyByte = KeyBuffer(UBound(KeyBuffer) - 1)
    If Compare = vbBinaryCompare Then
        'loop through the array
        For A = StartPos To 0 Step -1
            If ByteArray(A) = FirstKeyByte Then
                If ByteArray(A + KeyUpper) = LastKeyByte Then
                    If KeyLen > 4 Then
                        'check if keyword is found from the array
                        C = A + 1
                        For B = 2 To KeyLen Step 2
                            If Not (ByteArray(C) = KeyBuffer(B)) Then Exit For
                            C = C + 1
                        Next B
                        'keyword is found!
                        If B > KeyLen Then
                            InBArrRev = A
                            Exit Function
                        End If
                    Else
                        InBArrRev = A
                        Exit Function
                    End If
                End If
            End If
        Next A
    Else 'vbTextCompare
        FirstKeyByteU = KeyBufferU(0)
        LastKeyByteU = KeyBufferU(UBound(KeyBuffer) - 1)
        'loop through the array
        For A = StartPos To 0 Step -1
            TempByte = ByteArray(A)
            If TempByte = FirstKeyByte Or TempByte = FirstKeyByteU Then
                TempByte = ByteArray(A + KeyUpper)
                If TempByte = LastKeyByte Or TempByte = LastKeyByteU Then
                    If KeyLen > 4 Then
                        'check if keyword is found from the array
                        C = A + 1
                        For B = 2 To KeyLen Step 2
                            TempByte = ByteArray(C)
                            If Not (TempByte = KeyBuffer(B) Or TempByte = KeyBufferU(B)) Then Exit For
                            C = C + 1
                        Next B
                        'keyword is found!
                        If B > KeyLen Then
                            InBArrRev = A
                            Exit Function
                        End If
                    Else
                        InBArrRev = A
                        Exit Function
                    End If
                End If
            End If
        Next A
    End If
    InBArrRev = -1
End Function

