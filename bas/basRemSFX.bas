Attribute VB_Name = "basRemSFX"
Public Sub For7Z(sPath As String)
Static IsiSFX As String
IsiSFX = ReadUnicodeFile(sPath)
IsiSFX = Right$(IsiSFX, Len(IsiSFX) - InStrRev(IsiSFX, "7z¼") + 1)
WriteFileUniSim App.Path & "\a.7z", IsiSFX
End Sub

Public Sub Extract7z(sPathArc As String)
MsgBox Shell("7za.exe 7z e " & sPathArc & " -o" & App.Path & "\tmp *.* -r", vbHide)
End Sub

Public Sub Scan7z(sPath As String)
Equal (sPath)
End Sub
