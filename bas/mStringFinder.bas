Attribute VB_Name = "mStringFinder"
Private cStrSearch As classFindString

Public Function StringSearch(sPath As String, sString As String, sStart As String, sLeght As String) As Boolean
StringSearch = False
cStrSearch.SearchAlgorithm = Asm_BMHA
If cStrSearch.FileMapSearch(sPath, sString, sStart) > sLeght Then StringSearch = True
End Function
