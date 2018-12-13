Attribute VB_Name = "basReadDB"

Public Function ReadDb(spath As String)
Static sTemp As String
Static stmp() As String
Static sTmp2() As String
Static pisah As String
Static iCount As Integer
Static iTemp As Integer
pisah = Chr(13)
sTemp = ReadUnicodeFile(spath)
stmp() = Split(sTemp, pisah)
    
    iTemp = UBound(stmp())
    ReDim xNumChecksum(iTemp) As String
    ReDim xNamChecksum(iTemp) As String
        
    For iCount = 1 To iTemp
        sTmp2() = Split(stmp(iCount), ":")
        xNumChecksum(iCount) = Mid(sTmp2(0), 2)
        xNamChecksum(iCount) = sTmp2(1)
    frMain.lvVirLst.ListItems.Add , xNamChecksum(iCount), , 0
    Next
    
xJumChecksum = iTemp
frMain.lDB = ": " & iTemp
End Function
