Attribute VB_Name = "basScanArchive"
Dim ZF As New Cls_GetFileType
Public PathDalamArc As String

Public Function GetFileInArc(sPath As String)
Static tFileCount As Long
Static sCount As Long
Dim FileArc() As Boolean
If DapatkanUkuranFile(sPath) > 5242880 Then Exit Function
ZF.Get_Contents sPath
tFileCount = ZF.FileCount
PathDalamArc = sPath
ReDim FileArc(tFileCount) As Boolean
sCount = 1
For sCount = 1 To tFileCount
FileArc(sCount) = True
    ZF.UnPack FileArc, App.Path & "\tmp\a.tmp"
Next
End Function
