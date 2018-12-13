Attribute VB_Name = "Mod_Reduce"
Option Explicit
'This mode is not tested cause i couldn't find a file wich was reduced

Private Type Data_Type
    Data() As Byte
    Pos As Long
    BitPos As Integer
End Type

Private BitMask(15) As Long

Private Cdata As Data_Type
Private Udata As Data_Type

Public Function UnReduce(ByteArray() As Byte, Level As Integer, UncompressedSize As Long) As Integer
    Dim S(256, 32) As Integer
    Dim N(256) As Integer
    Dim B(64) As Integer
    Dim j As Integer, i As Integer, LastC As Integer, State As Integer, C As Byte
    Dim LN As Integer, Dist As Integer, Cnt As Integer
    Dim Temp()
    Dim X As Long
    Temp = Array(0, 1, 1, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 4, 4, 4, 4, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 6)
    For X = 0 To 64
        B(X) = Temp(X)
    Next
    ReDim Udata.Data(UncompressedSize)
    Udata.Pos = 0
    Cdata.Data = ByteArray
    Cdata.Pos = -1
    Cdata.BitPos = 0
    For X = 0 To 15
        BitMask(X) = 2 ^ X - 1
    Next
    Cnt = 0
    For j = 255 To 0 Step -1
        N(j) = GetBits(6)
        If N(j) > 32 Then
            UnReduce = -1               'Follower set to large
            N(j) = 32
        End If
        For i = 0 To N(j) - 1
            S(j, i) = GetBits(8)
            Cnt = Cnt + 1
        Next
    Next
    LastC = 0
    State = 0
    Do While Udata.Pos <= UncompressedSize
        If Not N(LastC) Then
            C = GetBits(8)
        Else
            If GetBits(1) Then
                C = GetBits(8)
            Else
                C = 0
                If N(LastC) <> 0 Then C = GetBits(B(N(LastC)))
                C = S(LastC, C)
            End If
        End If
        LastC = C
        Select Case State
        Case 0:
            If C <> 144 Then
                Call PutByte(C)
            Else
                State = 1
            End If
        Case 1:
            If C Then
                X = 9 - Level
                Dist = Fix(C / 2 ^ X) * 256
                LN = (2 ^ X - 1) And C
                State = 3
                If LN = (2 ^ X - 1) Then State = 2
            Else
                Call PutByte(144)
                State = 0
            End If
        Case 2:
            LN = LN + C
            State = 3
        Case 3:
            Dist = Dist + (C + 1)
            LN = LN + 3
            Do While LN
                Call PutByte(Udata.Data(Udata.Pos - Dist))
                LN = LN - 1
            Loop
            State = 0
        End Select
    Loop
    UnReduce = 0
End Function
 
Private Function GetBits(Numbits As Integer) As Long
    Dim NB As Integer
    Dim Value As Long
    If Numbits = 0 Then Exit Function
    If Cdata.BitPos = 0 Then Cdata.Pos = Cdata.Pos + 1
    NB = 8 - Cdata.BitPos
    Value = Fix(Cdata.Data(Cdata.Pos) / (2 ^ Cdata.BitPos))
    Do While NB < Numbits
        Cdata.Pos = Cdata.Pos + 1
        Value = Value + (Cdata.Data(Cdata.Pos) * (2 ^ NB))
        NB = NB + 8
    Loop
    Cdata.BitPos = (Cdata.BitPos + Numbits) Mod 8
    GetBits = Value And BitMask(Numbits)
End Function

Private Sub PutByte(Char As Byte)
    Udata.Data(Udata.Pos) = Char
    Udata.Pos = Udata.Pos + 1
End Sub

