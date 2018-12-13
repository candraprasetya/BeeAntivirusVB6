Attribute VB_Name = "mMemory"
'==================================================================================================
'mMemory.bas                                        1/17/05
'
'           PURPOSE:
'               General memory access and pointer manipulation.
'
'==================================================================================================

Option Explicit

Private myRand(0 To 255) As Byte
Private myArray()  As Byte
Private moArrPtr   As pcArrayPtr

Public Function UnsignedAdd(ByVal iStart As Long, ByVal iInc As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/05
    ' Purpose   : Emulate unsigned addition.
    '---------------------------------------------------------------------------------------
    UnsignedAdd = (iStart Xor &H80000000) + iInc Xor &H80000000
End Function

Public Sub Incr(ByRef i As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/05
    ' Purpose   : Increment value by one, wrapping if necessary.
    '---------------------------------------------------------------------------------------
    If i <> &H7FFFFFFF _
        Then i = i + 1& _
    Else i = &H80000000
End Sub

Public Property Get MemOffset32(ByVal iStart As Long, ByVal iOffset As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/05
    ' Purpose   : Return the 32 bit value at the given offset of the memory address.
    '---------------------------------------------------------------------------------------
    MemOffset32 = MemLong(ByVal UnsignedAdd(iStart, iOffset))
End Property

Public Property Let MemOffset32(ByVal iStart As Long, ByVal iOffset As Long, ByVal iNew As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/05
    ' Purpose   : Set the 32 bit value at the given offset of the memory address.
    '---------------------------------------------------------------------------------------
    MemLong(ByVal UnsignedAdd(iStart, iOffset)) = iNew
End Property

Public Property Get MemOffset16(ByVal iStart As Long, ByVal iOffset As Long) As Integer
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/05
    ' Purpose   : Return the 16 bit value at the given offset of the memory address.
    '---------------------------------------------------------------------------------------
    MemOffset16 = MemWord(ByVal UnsignedAdd(iStart, iOffset))
End Property

Public Property Let MemOffset16(ByVal iStart As Long, ByVal iOffset As Long, ByVal iNew As Integer)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/05
    ' Purpose   : Set the 16 bit value at the given offset of the memory address.
    '---------------------------------------------------------------------------------------
    MemWord(ByVal UnsignedAdd(iStart, iOffset)) = iNew
End Property

Public Property Get MemOffset8(ByVal iStart As Long, ByVal iOffset As Long) As Byte
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/05
    ' Purpose   : Return the 8 bit value at the given offset of the memory address.
    '---------------------------------------------------------------------------------------
    MemOffset8 = MemByte(ByVal UnsignedAdd(iStart, iOffset))
End Property

Public Property Let MemOffset8(ByVal iStart As Long, ByVal iOffset As Long, ByVal yNew As Byte)
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/05
    ' Purpose   : Set the 8 bit value at the given offset of the memory address.
    '---------------------------------------------------------------------------------------
    MemByte(ByVal UnsignedAdd(iStart, iOffset)) = yNew
End Property

Public Function MakeLong(ByVal iLower As Integer, ByVal iUpper As Integer) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/05
    ' Purpose   : Combine two words into a dword.
    '---------------------------------------------------------------------------------------
    MakeLong = iLower Or (iUpper * &H10000)
End Function

Public Function MemAlloc(ByVal iBytes As Long, Optional ByVal bZeroInit As Boolean) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/05
    ' Purpose   : Allocate a given amount of memory.
    '---------------------------------------------------------------------------------------
    MemAlloc = HeapAlloc(GetProcessHeap(), -HEAP_ZERO_MEMORY * bZeroInit, iBytes)
End Function

Public Function MemAllocFromString(ByVal pBStr As Long, ByVal iLength As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/05
    ' Purpose   : Allocate a given amount of memory and copy into it the data in the string.
    '---------------------------------------------------------------------------------------
    MemAllocFromString = MemAlloc(iLength)
    If MemAllocFromString Then
        If pBStr Then
            Dim liBStrLen      As Long
            liBStrLen = MemOffset32(pBStr, -4&)
            If liBStrLen > ZeroL Then
                If liBStrLen > iLength Then liBStrLen = iLength
                CopyMemory ByVal MemAllocFromString, ByVal pBStr, liBStrLen
                If liBStrLen < iLength _
                    Then MemOffset8(MemAllocFromString, liBStrLen) = ZeroY _
                Else MemOffset8(MemAllocFromString, liBStrLen - OneL) = ZeroY
                ElseIf iLength > ZeroL Then
                    MemOffset8(MemAllocFromString, ZeroL) = ZeroY
                End If
            End If
        End If
End Function

Public Function MemFree(ByVal hMem As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/05
    ' Purpose   : Free the block of memory.
    '---------------------------------------------------------------------------------------
    MemFree = HeapFree(GetProcessHeap(), ZeroL, hMem)
End Function








Private Sub pInitRand()
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/05
    ' Purpose   : Initialize the hash array.
    '---------------------------------------------------------------------------------------
    myRand(0) = 229: myRand(1) = 67: myRand(2) = 128: myRand(3) = 52: myRand(4) = 141: myRand(5) = 20: myRand(6) = 246: myRand(7) = 89: myRand(8) = 202: myRand(9) = 150: myRand(10) = 7: myRand(11) = 164: myRand(12) = 171: myRand(13) = 236: myRand(14) = 87: myRand(15) = 251: myRand(16) = 13: myRand(17) = 208: myRand(18) = 131: myRand(19) = 15: myRand(20) = 53: myRand(21) = 151: myRand(22) = 103: myRand(23) = 199: myRand(24) = 43: myRand(25) = 159: myRand(26) = 213: myRand(27) = 51: myRand(28) = 4: myRand(29) = 252: myRand(30) = 72: myRand(31) = 115
    myRand(32) = 181: myRand(33) = 220: myRand(34) = 126: myRand(35) = 144: myRand(36) = 0: myRand(37) = 112: myRand(38) = 206: myRand(39) = 224: myRand(40) = 85: myRand(41) = 97: myRand(42) = 162: myRand(43) = 35: myRand(44) = 130: myRand(45) = 138: myRand(46) = 82: myRand(47) = 122: myRand(48) = 121: myRand(49) = 68: myRand(50) = 247: myRand(51) = 214: myRand(52) = 192: myRand(53) = 14: myRand(54) = 172: myRand(55) = 250: myRand(56) = 233: myRand(57) = 234: myRand(58) = 18: myRand(59) = 155: myRand(60) = 238: myRand(61) = 92: myRand(62) = 57: myRand(63) = 117
    myRand(64) = 152: myRand(65) = 23: myRand(66) = 205: myRand(67) = 129: myRand(68) = 149: myRand(69) = 80: myRand(70) = 79: myRand(71) = 231: myRand(72) = 179: myRand(73) = 187: myRand(74) = 114: myRand(75) = 216: myRand(76) = 177: myRand(77) = 239: myRand(78) = 12: myRand(79) = 167: myRand(80) = 74: myRand(81) = 120: myRand(82) = 244: myRand(83) = 142: myRand(84) = 109: myRand(85) = 174: myRand(86) = 217: myRand(87) = 104: myRand(88) = 157: myRand(89) = 99: myRand(90) = 195: myRand(91) = 132: myRand(92) = 76: myRand(93) = 198: myRand(94) = 137: myRand(95) = 212
    myRand(96) = 91: myRand(97) = 60: myRand(98) = 222: myRand(99) = 50: myRand(100) = 70: myRand(101) = 189: myRand(102) = 66: myRand(103) = 255: myRand(104) = 22: myRand(105) = 81: myRand(106) = 158: myRand(107) = 226: myRand(108) = 49: myRand(109) = 140: myRand(110) = 248: myRand(111) = 133: myRand(112) = 96: myRand(113) = 156: myRand(114) = 65: myRand(115) = 105: myRand(116) = 116: myRand(117) = 249: myRand(118) = 225: myRand(119) = 40: myRand(120) = 59: myRand(121) = 253: myRand(122) = 178: myRand(123) = 254: myRand(124) = 83: myRand(125) = 39: myRand(126) = 32: myRand(127) = 61
    myRand(128) = 243: myRand(129) = 240: myRand(130) = 25: myRand(131) = 176: myRand(132) = 84: myRand(133) = 118: myRand(134) = 136: myRand(135) = 110: myRand(136) = 221: myRand(137) = 27: myRand(138) = 58: myRand(139) = 193: myRand(140) = 47: myRand(141) = 101: myRand(142) = 184: myRand(143) = 33: myRand(144) = 191: myRand(145) = 190: myRand(146) = 42: myRand(147) = 94: myRand(148) = 98: myRand(149) = 154: myRand(150) = 143: myRand(151) = 203: myRand(152) = 77: myRand(153) = 124: myRand(154) = 2: myRand(155) = 73: myRand(156) = 30: myRand(157) = 183: myRand(158) = 185: myRand(159) = 21
    myRand(160) = 88: myRand(161) = 160: myRand(162) = 145: myRand(163) = 147: myRand(164) = 19: myRand(165) = 165: myRand(166) = 168: myRand(167) = 215: myRand(168) = 64: myRand(169) = 55: myRand(170) = 10: myRand(171) = 245: myRand(172) = 26: myRand(173) = 37: myRand(174) = 6: myRand(175) = 31: myRand(176) = 230: myRand(177) = 242: myRand(178) = 93: myRand(179) = 38: myRand(180) = 71: myRand(181) = 62: myRand(182) = 232: myRand(183) = 237: myRand(184) = 180: myRand(185) = 1: myRand(186) = 36: myRand(187) = 75: myRand(188) = 127: myRand(189) = 11: myRand(190) = 3: myRand(191) = 163
    myRand(192) = 134: myRand(193) = 170: myRand(194) = 5: myRand(195) = 24: myRand(196) = 207: myRand(197) = 54: myRand(198) = 8: myRand(199) = 106: myRand(200) = 63: myRand(201) = 56: myRand(202) = 48: myRand(203) = 235: myRand(204) = 186: myRand(205) = 209: myRand(206) = 201: myRand(207) = 44: myRand(208) = 241: myRand(209) = 100: myRand(210) = 139: myRand(211) = 46: myRand(212) = 95: myRand(213) = 113: myRand(214) = 17: myRand(215) = 197: myRand(216) = 194: myRand(217) = 41: myRand(218) = 111: myRand(219) = 28: myRand(220) = 78: myRand(221) = 169: myRand(222) = 188: myRand(223) = 29
    myRand(224) = 69: myRand(225) = 108: myRand(226) = 153: myRand(227) = 34: myRand(228) = 146: myRand(229) = 135: myRand(230) = 228: myRand(231) = 210: myRand(232) = 166: myRand(233) = 119: myRand(234) = 200: myRand(235) = 227: myRand(236) = 182: myRand(237) = 107: myRand(238) = 45: myRand(239) = 161: myRand(240) = 86: myRand(241) = 148: myRand(242) = 175: myRand(243) = 125: myRand(244) = 9: myRand(245) = 218: myRand(246) = 123: myRand(247) = 173: myRand(248) = 16: myRand(249) = 102: myRand(250) = 90: myRand(251) = 223: myRand(252) = 219: myRand(253) = 204: myRand(254) = 211: myRand(255) = 196
End Sub

Private Function pIdeFix() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/05
    ' Purpose   : Still don't know why this is necessary.
    '---------------------------------------------------------------------------------------
    pIdeFix = True
    moArrPtr.SetArrayByte myArray
End Function

Public Function HashLong(ByVal iLong As Long) As Byte
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/05
    ' Purpose   : Return a hash of the 32 bit value.
    '---------------------------------------------------------------------------------------
    If myRand(0) = ZeroY Then pInitRand
    If moArrPtr Is Nothing Then
        Set moArrPtr = New pcArrayPtr
        moArrPtr.SetArrayByte myArray
    End If
    
    'debug.assert pIdeFix
    
    moArrPtr.PointToLong iLong
    
    Dim i      As Long
    For i = ZeroL To 3&
        HashLong = myRand(HashLong Xor myArray(i))
    Next
End Function

Public Function Hash(ByVal iPtr As Long, ByVal iLen As Long) As Byte
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/05
    ' Purpose   : Return a hash of the given length of memory.
    '---------------------------------------------------------------------------------------
    If myRand(0) = ZeroY Then pInitRand
    If moArrPtr Is Nothing Then
        Set moArrPtr = New pcArrayPtr
        moArrPtr.SetArrayByte myArray
    End If
    
    'debug.assert iPtr
    'debug.assert pIdeFix
    
    moArrPtr.POINT iPtr, iLen
    Dim i      As Long
    For i = ZeroL To iLen - OneL
        Hash = myRand(Hash Xor myArray(i))
    Next
End Function

Public Function MemCmp(ByVal iPtr1 As Long, ByVal iPtr2 As Long, ByVal iLen As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 1/17/05
    ' Purpose   : Compare the two memory areas and return a value indicating which is greater.
    '---------------------------------------------------------------------------------------
    'debug.assert (iLen Mod 4&) = ZeroL
    
    Static oArrPtr1 As pcArrayPtr
    Static oArrPtr2 As pcArrayPtr
    Static iArr1() As Long
    Static iArr2() As Long

    If oArrPtr1 Is Nothing Then
        Set oArrPtr1 = New pcArrayPtr
        oArrPtr1.SetArrayLong iArr1
    End If
    If oArrPtr2 Is Nothing Then
        Set oArrPtr2 = New pcArrayPtr
        oArrPtr2.SetArrayLong iArr2
    End If
    
    oArrPtr1.POINT iPtr1, iLen
    oArrPtr2.POINT iPtr2, iLen
    
    Dim i1      As Long
    Dim i2      As Long
    
    MemCmp = ZeroL
    
    For iLen = ZeroL To (iLen \ 4&) - OneL
        i1 = iArr1(iLen)
        i2 = iArr2(iLen)
        If i1 > i2 Then
            MemCmp = NegOneL
            Exit Function
        ElseIf i1 <> i2 Then
            MemCmp = OneL
            Exit Function
        End If
    Next
    
End Function

Public Property Get hiword(ByVal i As Long) As Integer
    hiword = MemWord(ByVal UnsignedAdd(VarPtr(i), 2))
End Property
Public Property Get loword(ByVal i As Long) As Integer
    loword = MemWord(ByVal VarPtr(i))
End Property


'Public Function IsGoodRand() As Boolean
'    pInitRand
'    Dim i As Long
'    Dim loColl As New Collection
'    For i = 0 To 255
'        loColl.Add "", CStr(myRand(i))
'    Next
'    IsGoodRand = True
'End Function
'
'Public Sub GenRand()
'Dim loColl As New Collection
'
'Randomize Timer
'
'Dim lyB As Byte
'
'On Error Resume Next
'Do Until loColl.Count = 256
'    lyB = Rnd * 256
'    loColl.Add lyB, CStr(lyB)
'Loop
'
'Dim v As Variant
'Dim ls As String
'Dim i As Long
'
'ls = "Private Sub InitRand()" & vbCrLf & vbTab
'For Each v In loColl
'    ls = ls & "myRand(" & i & ") = " & v
'    If i < 255 Then
'        If (i + 1) Mod 32 = 0 Then
'            ls = ls & vbCrLf & vbTab
'        Else
'            ls = ls & ": "
'        End If
'    End If
'    i = i + OneL
'Next
'ls = ls & vbCrLf & "End Sub"
'
'Clipboard.Clear
'Clipboard.SetText ls
'
'End Sub
