Attribute VB_Name = "kGen"
Option Explicit

Function KeyGen(kNamev As Variant, kPass As String, kType As Integer) As String
'****************************************************************************
'* KeyGen v2.01 Build 01                                                    *
'* Copyright © 2000 W.G.Griffiths                                           *
'* E-Mail: w.g.griffiths@telinco.co.uk                                      *
'****************************************************************************

On Error Resume Next         'still here just as a precaution

Dim cTable(512) As Integer   'character map
Dim nKeys(16) As Integer     'xor keys used for pArray(x) xor nkeys(x)
Dim s0(512) As Integer       'swap-box data used to map character table
Dim nArray(16) As Integer    'name array data
Dim pArray(16) As Integer    'password array data
Dim n As Integer             'for next loop counter
Dim nPtr As Integer          'name pointer (used for counting)
Dim cPtr As Integer          'character pointer (used for counting)
Dim cFlip As Boolean         'character flip (used to flip between numeric and alpha)
Dim sIni As Integer          'holds s-box values
Dim temp As Integer          'holds s-box values
Dim rtn As Integer           'holds generated key values used agains chr map
Dim gKey As String           'generated key as string
Dim nLen As Integer          'number of chr's in name
Dim pLen As Integer          'number of chr's in password
Dim kPtr As Integer          'key pointer
Dim sPtr As Integer          'space pointer (used in hex key)
Dim nOffset As Integer       'name offset
Dim pOffset As Integer       'password offset
Dim tOffset As Integer       'total offset
Dim KeySize As Integer       'the size of the key to make

Const nXor As Integer = 18   'name xor value
Const pXor As Integer = 25   'password xor value
Const cLw As Integer = 65    'character lower limit 65 = A ** do not change **
Const nLw As Integer = 48    'number lower limit 48 = 0 ** do not change **
Const sOffset As Integer = 0 'character map offset

'****************************************************************************
'Thanks to Chris Fournier for his suggestions for adding support for        *
'Strings, Objects and String() as arrays                                    *
'Your comments please                                                       *
'****************************************************************************
Dim VarType As String
Dim kName As String
Dim AryCtl As Integer
Dim AryCtrl As Control

VarType = TypeName(kNamev)

Select Case VarType
    Case "String"
        kName = kNamev
    Case "TextBox"
        kName = kNamev.Text
    Case "Object"
        For Each AryCtrl In kNamev
            If AryCtrl.Text <> "" Then
                kName = kName & AryCtrl.Text & "|"
            End If
        Next
        kName = Left(kName, Len(kName) - 1)
    Case "String()"
        For AryCtl = LBound(kNamev) To UBound(kNamev)
            If kNamev(AryCtl) <> "" Then
                kName = kName & kNamev(AryCtl) & "|"
            End If
        Next
        kName = Left(kName, Len(kName) - 1)
        Case Else
            MsgBox VarType & " is an unsupported type to be passed to KeyGen"
End Select
'****************************************************************************

nLen = Len(kName)
pLen = Len(kPass)

'password xor keys ** change to make keygen unique **
nKeys(1) = 46
nKeys(2) = 89
nKeys(3) = 142
nKeys(4) = 63
nKeys(5) = 231
nKeys(6) = 32
nKeys(7) = 129
nKeys(8) = 51
nKeys(9) = 28
nKeys(10) = 97
nKeys(11) = 248
nKeys(12) = 41
nKeys(13) = 136
nKeys(14) = 53
nKeys(15) = 78
nKeys(16) = 164

sIni = 0

'set s boxes
For n = 0 To 512
    s0(n) = n
Next n

For n = 0 To 512
    sIni = (sOffset + sIni + n) Mod 256
    temp = s0(n)
    s0(n) = s0(sIni)
    s0(sIni) = temp
Next n

If kType = 1 Then       '(numeric)
    
    nPtr = 0
    KeySize = 16
    gKey = String(16, " ")
    
    For n = 0 To 512
        cTable(s0(n)) = (nLw + (nPtr))
        nPtr = nPtr + 1
        If nPtr = 10 Then nPtr = 0
    Next n
    
    

ElseIf kType = 2 Then   '(alphanumeric)
    
    nPtr = 0
    cPtr = 0
    KeySize = 16
    gKey = String(16, " ")
    
    cFlip = False
    For n = 0 To 512
        If cFlip Then
            cTable(s0(n)) = (nLw + nPtr)
            nPtr = nPtr + 1
            If nPtr = 10 Then nPtr = 0
            cFlip = False
        Else
            cTable(s0(n)) = (cLw + cPtr)
            cPtr = cPtr + 1
            If cPtr = 26 Then cPtr = 0
            cFlip = True
        End If
    Next n
    
    
    
Else  '(hex)

    KeySize = 8
    gKey = String(19, " ")
    
End If

kPtr = 1

For n = 1 To nLen 'name
  nArray(kPtr) = nArray(kPtr) + Asc(Mid(kName, n, 1)) Xor nXor
  nOffset = nOffset + nArray(kPtr)
  kPtr = kPtr + 1
    If kPtr = 9 Then kPtr = 1
Next n

For n = 1 To pLen 'password
  pArray(kPtr) = pArray(kPtr) + Asc(Mid(kPass, n, 1)) Xor pXor
  pOffset = pOffset + pArray(kPtr)
  kPtr = kPtr + 1
    If kPtr = 9 Then kPtr = 1
Next n

tOffset = (nOffset + pOffset) Mod 512

kPtr = 1
sPtr = 1
For n = 1 To KeySize
  pArray(n) = pArray(n) Xor nKeys(n)
  rtn = Abs(((nArray(n) Xor pArray(n)) Mod 512) - tOffset)
  
  If kType = 3 Then 'hex key
        If rtn < 16 Then
            Mid(gKey, kPtr, 2) = "0" & Hex(rtn)
        Else
            Mid(gKey, kPtr, 2) = Hex(rtn)
        End If
            If sPtr = 2 And kPtr < 18 Then
                kPtr = kPtr + 1
                Mid(gKey, kPtr + 1, 1) = "-"
            End If
        kPtr = kPtr + 2
        sPtr = sPtr + 1
        If sPtr = 3 Then sPtr = 1
  Else  'numeric - alphanumeric key
    Mid(gKey, n, 1) = Chr(cTable(rtn))
  End If
Next

KeyGen = gKey

End Function
