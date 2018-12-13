Attribute VB_Name = "Mod_Inflate64"
Option Explicit
'This mod is the famoes Inflate routine used by several different
'Compression programs like ZIP,gZip,PNG,etc..
'This module is created by Marco v/d Berg but is heavely optimized by John Korejwa
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type CodesType
    Lenght() As Long
    code() As Long
End Type

Private OutStream() As Byte
Private OutPos As Long
Private InStream() As Byte
Private Inpos As Long
Private ByteBuff As Long
Private BitNum As Long
Private BitMask(16) As Long
Private Pow2(16) As Long

Private LC As CodesType
Private dc As CodesType
Private LitLen As CodesType          'Literal/length tree
Private Dist As CodesType            'Distance tree
Private TempLit As CodesType
Private TempDist As CodesType

Private LenOrder(18) As Long
Private MinLLenght As Long           'Minimum length used in literal/lenght codes
Private MaxLLenght As Long           'Maximum length used in literal/lenght codes
Private MinDLenght As Long           'Minimum length used in distance codes
Private MaxDLenght As Long           'Maximum length used in distance codes
Private IsStaticBuild As Boolean

Public Function Inflate(ByteArray() As Byte, UncompressedSize As Long, Optional ZIP64 As Boolean = False) As Long
'On Error GoTo errhandle

    Dim IsLastBlock  As Boolean
    Dim CompType     As Long
    Dim Char         As Long
    Dim Nubits       As Long
    Dim L1           As Long
    Dim L2           As Long
    Dim X            As Long


    InStream = ByteArray 'Copy local array to global array
    Call Init_Inflate(UncompressedSize) 'Init global variables

    Do
        IsLastBlock = GetBits(1) 'Read Last Block Flag
        CompType = GetBits(2)    'Read Block Type

        If CompType = 0 Then              'Block is Stored
            If Inpos + 4 > UBound(InStream) Then
                Inflate = -1 'InStream depleated
                Exit Do
            End If
'this is done couse if bitnum >= then next byte is already in ByteBuff
            Do While BitNum >= 8
                Inpos = Inpos - 1
                BitNum = BitNum - 8
            Loop
            CopyMemory L1, InStream(Inpos), 2&      'Read Count
            CopyMemory L2, InStream(Inpos + 2), 2&  'Read ones compliment of Count
            Inpos = Inpos + 4
            If L1 - (Not (L2) And &HFFFF&) Then Inflate = -2
            If Inpos + L1 - 1 > UBound(InStream) Then
                Inflate = -1 'InStream depleated
                Exit Do
            End If
            If OutPos + L1 - 1 > UBound(OutStream) Then
                Inflate = -1 'OutStream overflow
                Exit Do
            End If
            CopyMemory OutStream(OutPos), InStream(Inpos), L1 'Copy stored Block
            OutPos = OutPos + L1
            Inpos = Inpos + L1
            ByteBuff = 0
            BitNum = 0
        ElseIf CompType = 3 Then          'Error in compressed data
            Inflate = -1
            Exit Do
        Else
            If CompType = 1 Then          'Static Compression
                If Create_Static_Tree <> 0 Then
                    MsgBox "Error in tree creation (Static)"
                    Exit Function
                End If
            Else 'CompType = 2            'Dynamic Compression
                If Create_Dynamic_Tree <> 0 Then
                    MsgBox "Error in tree creation (Static)"
                    Exit Function
                End If
            End If
            Do
                NeedBits MaxLLenght
                Nubits = MinLLenght
                Do While LitLen.Lenght(ByteBuff And BitMask(Nubits)) <> Nubits
                    Nubits = Nubits + 1
                Loop
                Char = LitLen.code(ByteBuff And BitMask(Nubits))
                DropBits Nubits
                If Char < 256 Then  'Character is Literal
                    OutStream(OutPos) = Char        'output the character
                    OutPos = OutPos + 1
                ElseIf Char > 256 Then 'Character is Length Symbol
                    'Decode Length L1
                    Char = Char - 257
                    L1 = LC.code(Char) + GetBits(LC.Lenght(Char))
                    If (L1 = 258) And ZIP64 Then L1 = GetBits(16) + 3
                    'Decode Distance L2 Symbol
                    NeedBits MaxDLenght
                    Nubits = MinDLenght
                    Do While Dist.Lenght(ByteBuff And BitMask(Nubits)) <> Nubits
                        Nubits = Nubits + 1
                    Loop
                    Char = Dist.code(ByteBuff And BitMask(Nubits))
                    DropBits Nubits
                    L2 = dc.code(Char) + GetBits(dc.Lenght(Char)) 'Decode Distance L2
                    'copy L2 positions back L1 characters
                    For X = 1 To L1
                        OutStream(OutPos) = OutStream(OutPos - L2)
                        OutPos = OutPos + 1
                    Next X
                End If
            Loop While Char <> 256 'EOB
        End If
    Loop While Not IsLastBlock
Stop_Decompression:
    If OutPos > 0 Then
        ReDim Preserve OutStream(OutPos - 1)
    Else
        Erase OutStream
    End If
'Clear memory
    Erase InStream
    Erase BitMask
    Erase Pow2
    Erase LC.code
    Erase LC.Lenght
    Erase dc.code
    Erase dc.Lenght
    Erase LitLen.code
    Erase LitLen.Lenght
    Erase Dist.code
    Erase Dist.Lenght
    Erase LenOrder
    ByteArray = OutStream
    
    Exit Function
errhandle:
If OutPos > UBound(OutStream) Then
    MsgBox "Incorrect Uncompressed Size"
    GoTo Stop_Decompression
ElseIf Inpos > UBound(InStream) Then
    MsgBox "Unexpected End of File"
    GoTo Stop_Decompression
Else
    Err.Raise Err.Number, , Err.Description
End If

End Function

'This sub is used to create a static huffmann tree for inflate
Private Function Create_Static_Tree()
    Dim X As Long
    Dim Lenght(287) As Long
    If IsStaticBuild = False Then
        For X = 0 To 143: Lenght(X) = 8: Next
        For X = 144 To 255: Lenght(X) = 9: Next
        For X = 256 To 279: Lenght(X) = 7: Next
        For X = 280 To 287: Lenght(X) = 8: Next
        If Create_Codes(TempLit, Lenght, 287, MaxLLenght, MinLLenght) <> 0 Then
            Create_Static_Tree = -1
            Exit Function
        End If
    
        For X = 0 To 31: Lenght(X) = 5: Next
        Create_Static_Tree = Create_Codes(TempDist, Lenght, 31, MaxDLenght, MinDLenght)
        IsStaticBuild = True
    Else
        MinLLenght = 7
        MaxLLenght = 9
        MinDLenght = 5
        MaxDLenght = 5
    End If
    LitLen = TempLit
    Dist = TempDist
End Function

'This sub is used to create a dynamic tree for inflate
Private Function Create_Dynamic_Tree() As Long
    Dim Lenght() As Long
    Dim Bl_Tree As CodesType
    Dim MinBL As Long
    Dim MaxBL As Long
    Dim NumLen As Long
    Dim Numdis As Long
    Dim NumCod As Long
    Dim Char As Long
    Dim Nubits As Long
    Dim LN As Long
    Dim Pos As Long
    Dim X As Long

    NumLen = GetBits(5) + 257   'Get lenght of the literal/lenght tree
    Numdis = GetBits(5) + 1     'Get lenght of the distance tree
    NumCod = GetBits(4) + 4     'Get number of codes for the tree to form the other trees
    ReDim Lenght(18)
    'read the lengths per code
    For X = 0 To NumCod - 1
        Lenght(LenOrder(X)) = GetBits(3)
    Next
    'codes not used get lenght 0
    For X = NumCod To 18
        Lenght(LenOrder(X)) = 0
    Next
    'create the construction tree
    If Create_Codes(Bl_Tree, Lenght, 18, MaxBL, MinBL) <> 0 Then
        Create_Dynamic_Tree = -1
        Exit Function
    End If

'Get the codes for the literal/lenght and distance trees
    ReDim Lenght(NumLen + Numdis)
    Pos = 0
    Do While Pos < NumLen + Numdis
        NeedBits MaxBL
        Nubits = MinBL
        Do While Bl_Tree.Lenght(ByteBuff And BitMask(Nubits)) <> Nubits
            Nubits = Nubits + 1
        Loop
        Char = Bl_Tree.code(ByteBuff And BitMask(Nubits))
        DropBits Nubits
        
        If Char < 16 Then
            Lenght(Pos) = Char
            Pos = Pos + 1
        Else
            If Char = 16 Then
                If Pos = 0 Then  'no last lenght
                    Create_Dynamic_Tree = -5
                    Exit Function
                End If
                LN = Lenght(Pos - 1)
                Char = 3 + GetBits(2)
            ElseIf Char = 17 Then
                Char = 3 + GetBits(3)
                LN = 0
            Else
                Char = 11 + GetBits(7)
                LN = 0
            End If
            If Pos + Char > NumLen + Numdis Then            'to many lenghts
                Create_Dynamic_Tree = -6
                Exit Function
            End If
            Do While Char > 0
                Char = Char - 1
                Lenght(Pos) = LN
                Pos = Pos + 1
            Loop
        End If
    Loop
    'create the literal/lenght tree
    If Create_Codes(LitLen, Lenght, NumLen - 1, MaxLLenght, MinLLenght) <> 0 Then
        Create_Dynamic_Tree = -1
        Exit Function
    End If
    For X = 0 To Numdis
        Lenght(X) = Lenght(X + NumLen)
    Next
    'create the distance tree
    Create_Dynamic_Tree = Create_Codes(Dist, Lenght, Numdis - 1, MaxDLenght, MinDLenght)
End Function

'This function is used to retrieve the codes belonging to the huffmann-trees
Private Function Create_Codes(tree As CodesType, Lenghts() As Long, NumCodes As Long, MaxBits As Long, Minbits As Long) As Long
    Dim bits(16) As Long
    Dim next_code(16) As Long
    Dim code As Long
    Dim LN As Long
    Dim X As Long

'retrieve the bitlenght count and minimum and maximum bitlenghts
    Minbits = 16
    For X = 0 To NumCodes
        bits(Lenghts(X)) = bits(Lenghts(X)) + 1
        If Lenghts(X) > MaxBits Then MaxBits = Lenghts(X)
        If Lenghts(X) < Minbits And Lenghts(X) > 0 Then Minbits = Lenghts(X)
    Next
    LN = 1
    For X = 1 To MaxBits
        LN = LN + LN
        LN = LN - bits(X)
        If LN < 0 Then Create_Codes = LN: Exit Function 'Over subscribe, Return negative
    Next
    Create_Codes = LN

    ReDim tree.code(2 ^ MaxBits - 1)  'set the right dimensions
    ReDim tree.Lenght(2 ^ MaxBits - 1)
    code = 0
    bits(0) = 0
    For X = 1 To MaxBits
        code = (code + bits(X - 1)) * 2
        next_code(X) = code
    Next
    For X = 0 To NumCodes
        LN = Lenghts(X)
        If LN <> 0 Then
            code = Bit_Reverse(next_code(LN), LN)
            tree.Lenght(code) = LN
            tree.code(code) = X
            next_code(LN) = next_code(LN) + 1
        End If
    Next
End Function

'Inflated codes are stored in reversed order so this funtion will
'reverse the stored order to get the original value back
Private Function Bit_Reverse(ByVal Value As Long, ByVal Numbits As Long)
    Do While Numbits > 0
        Bit_Reverse = Bit_Reverse * 2 + (Value And 1)
        Numbits = Numbits - 1
        Value = Value \ 2
    Loop
End Function

Private Sub Init_Inflate(UncompressedSize As Long)
    Dim Temp()
    Dim X As Long
    ReDim OutStream(UncompressedSize)
    Erase LitLen.code
    Erase LitLen.Lenght
    Erase Dist.code
    Erase Dist.Lenght
    ReDim LC.code(31)
    ReDim LC.Lenght(31)
    ReDim dc.code(31)
    ReDim dc.Lenght(31)

    'Create the read order array
    Temp() = Array(16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15)
    For X = 0 To UBound(Temp): LenOrder(X) = Temp(X): Next
    'Create the Start lenghts array
    Temp() = Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258)
    For X = 0 To UBound(Temp): LC.code(X) = Temp(X): Next
    'Create the Extra lenght bits array
    Temp() = Array(0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0)
    For X = 0 To UBound(Temp): LC.Lenght(X) = Temp(X): Next
    'Create the distance code array
    Temp() = Array(1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577, 32769, 49153)
    For X = 0 To UBound(Temp): dc.code(X) = Temp(X): Next
    'Create the extra bits distance codes
    Temp() = Array(0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13, 14, 14)
    For X = 0 To UBound(Temp): dc.Lenght(X) = Temp(X): Next


    For X = 0 To 16
        BitMask(X) = 2 ^ X - 1
        Pow2(X) = 2 ^ X
    Next
    OutPos = 0
    Inpos = 0
    ByteBuff = 0
    BitNum = 0
End Sub

Private Sub PutByte(Char As Byte)
    If OutPos > UBound(OutStream) Then ReDim Preserve OutStream(OutPos + 1000)
    OutStream(OutPos) = Char
    OutPos = OutPos + 1
End Sub

'This sub Makes sure that there are at least the number of requested bits
'in ByteBuff
Private Sub NeedBits(Numbits As Long)
    While BitNum < Numbits
        If Inpos > UBound(InStream) Then Exit Sub   'do not past end
        ByteBuff = ByteBuff + (InStream(Inpos) * Pow2(BitNum))
        BitNum = BitNum + 8
        Inpos = Inpos + 1
    Wend
End Sub

'This sub will drop the amount of bits requested
Private Sub DropBits(Numbits As Long)
    ByteBuff = ByteBuff \ Pow2(Numbits)
    BitNum = BitNum - Numbits
End Sub

Private Function GetBits(Numbits As Long) As Long
    While BitNum < Numbits
        ByteBuff = ByteBuff + (InStream(Inpos) * Pow2(BitNum))
        BitNum = BitNum + 8
        Inpos = Inpos + 1
    Wend
    GetBits = ByteBuff And BitMask(Numbits)
    ByteBuff = ByteBuff \ Pow2(Numbits)
    BitNum = BitNum - Numbits
End Function

