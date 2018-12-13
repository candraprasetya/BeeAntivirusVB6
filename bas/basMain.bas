Attribute VB_Name = "basMain"
Sub Main()
StopScan = True
Select Case UCase$(Left$(Command, 2))
Case "/S"
StopScan = False
PathCustomScan = Mid$(Command, 4, Len(Command) - 3)
ReadPathFromCM PathCustomScan
isFromContext = True
frSplash.Loading2
frMain.Show
AllReset
frMain.ScanFromCM

Case "/U"
isFromContext = False
AllReset
frSplash.Show
LetakanForm frSplash, True
frSplash.Loading
LetakanForm frSplash, False

Case Else
isFromContext = False
frSplash.Show
LetakanForm frSplash, True
frSplash.tm.Enabled = True
LetakanForm frSplash, False
End Select
End Sub
