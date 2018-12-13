Attribute VB_Name = "basReadDB"

Public Function ReadDb(sPath As String)
Static sHuff As New classHuffman
Static sSXOr As New classSimpleXOR
Static ByteArr() As Byte
Static DataSementara As String

Static sTemp As String
Static sTmp() As String
Static sTmp2() As String
Static pisah As String
Static iCount As Integer
Static iTemp As Integer
Static xHnd As Long
Static xHnd2 As Long
pisah = Chr(13)

'xHnd2 = GetHandleFile(App.Path & "\VirDb\Decrypt\db1.dbC2")
'Call ReadUnicodeFile2(xHnd2, 1, GetSizeFile(xHnd2), ByteArr())
'sHuff.EncodeByte ByteArr(), UBound(ByteArr) + 1
'DataSementara = StrConv(ByteArr, vbUnicode)
'WriteFileUniSim App.Path & "\db1.dbC", DataSementara
'CryptVirus App.Path & "\db1.dbC", App.Path & "\db1.dbC"
'Exit Function

DeCryptVirus App.Path & "\VirDb\db1.dbC", GetSpecFolder(SYSTEM_DIR) & "\database.db"
xHnd = GetHandleFile(GetSpecFolder(SYSTEM_DIR) & "\database.db")
Call ReadUnicodeFile2(xHnd, 1, GetSizeFile(xHnd), ByteArr())
DataSementara = ReadUnicodeFile(GetSpecFolder(SYSTEM_DIR) & "\database.db")
sHuff.DecodeByte ByteArr(), UBound(ByteArr) + 1
DataSementara = StrConv(ByteArr, vbUnicode)
sTemp = DataSementara
sTmp() = Split(sTemp, pisah)
    
    iTemp = UBound(sTmp())
    ReDim xNumChecksum(iTemp) As String
    ReDim xNamChecksum(iTemp) As String
        
    For iCount = 1 To iTemp
        sTmp2() = Split(sTmp(iCount), ":")
        xNumChecksum(iCount) = Mid(sTmp2(0), 2)
        xNamChecksum(iCount) = sTmp2(1)
    frMain.lvVirLst.ListItems.Add , xNamChecksum(iCount), , 0
    Next
    
xJumChecksum = iTemp
frMain.lDB = iTemp
DbDefinition = Right$(sTmp(0), Len(sTmp(0)) - 8)
DbDefinition = Left$(DbDefinition, Len(DbDefinition) - 1)
frMain.lDate.Caption = DbDefinition
TutupFile xHnd
End Function
