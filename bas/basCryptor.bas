Attribute VB_Name = "basCryptor"
' Module untuk cryptografi teknik "by join"
' HrXxX

Dim sRangkai As String
Dim cRoungDiv As Single
Dim sTail As String
Dim iTail As Integer
Dim sJoin() As String

Public Function EncryptFile(fPath As String, KeyNum As Long) As String
Static sNum As Single
Static incre As Single
Static isiFile As String

isiFile = BacaFile(fPath)

incre = 1
sRangkai = ""
iTail = Len(isiFile) Mod KeyNum
sTail = Right(isiFile, iTail)
cRoungDiv = (Len(isiFile) - iTail) / KeyNum

If cRoungDiv Mod 2 = 1 Then
   cRoungDiv = cRoungDiv - 1
   sTail = Right(isiFile, iTail + KeyNum)
End If

ReDim sJoin(cRoungDiv - 1) As String

For sNum = 0 To cRoungDiv - 1
DoEvents
    sJoin(sNum) = Mid(isiFile, incre, KeyNum)
    incre = incre + KeyNum
Next

sNum = 0
incre = 0

For sNum = 1 To ((UBound(sJoin()) + 1) / 2)
DoEvents
    sRangkai = sRangkai & JoinAB(sJoin(incre), sJoin(incre + 1))
    incre = incre + 2
Next

Kill fPath
BuatFile sRangkai & sTail, fPath

End Function

Public Function DecryptFile(fPath As String, KeyNum As Long) As String
Static sNum As Single
Static incre As Single
Static isiFile As String

isiFile = BacaFile(fPath)

incre = 1
sRangkai = ""
iTail = Len(isiFile) Mod KeyNum
sTail = Right(isiFile, iTail)
cRoungDiv = (Len(isiFile) - iTail) / KeyNum

If cRoungDiv Mod 2 = 1 Then
   cRoungDiv = cRoungDiv - 1
   sTail = Right(isiFile, iTail + KeyNum)
End If

ReDim sJoin(cRoungDiv - 1) As String

For sNum = 0 To cRoungDiv - 1
DoEvents
    sJoin(sNum) = Mid(isiFile, incre, KeyNum)
    incre = incre + KeyNum
Next

sNum = 0
incre = 0

For sNum = 1 To ((UBound(sJoin()) + 1) / 2)
DoEvents
    sRangkai = sRangkai & JoinAB(sJoin(incre), sJoin(incre + 1))
    incre = incre + 2
Next

Kill fPath
BuatFile sRangkai & sTail, fPath

End Function


Private Function BacaFile(source As String) As String
Static TMP As String
Open source For Binary As #1
    TMP = Space(LOF(1))
    Get #1, , TMP
Close #1
BacaFile = TMP
End Function

Private Function BuatFile(isi As String, cpath As String)
Open cpath For Binary As #1
    Put #1, , isi
Close #1
End Function
Private Function JoinAB(a As String, B As String) As String
    JoinAB = B & a
End Function


