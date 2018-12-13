Attribute VB_Name = "basRTP"
Option Explicit
Dim L1(100) As String
Dim L2(100) As String
Dim tL1 As Long, tL2 As Long
Dim iewindow As InternetExplorer
Private currentwindows As ShellWindows
Sub iniT()
Set currentwindows = New ShellWindows
Timer.Enabled = True
Me.WindowState = 1
End Sub
Private Sub Timer_Timer()
Dim i As Long
On Error GoTo TheEnd
If currentwindows.count > 0 Then
Erase L2
tL2 = 0
    For Each iewindow In currentwindows
        DoEvents
        If iewindow.Busy Then GoTo busysignal
    Dim currentlocation As String
        currentlocation = iewindow.LocationURL
        If Mid$(currentlocation, 1, 7) = "file://" Then
                 currentlocation = Replace(currentlocation, "file:///", "")
                 currentlocation = Replace(currentlocation, "%20", " ")
                 currentlocation = Replace(currentlocation, "/", "\")
                 currentlocation = Replace(currentlocation, "%5B", "[")
                 currentlocation = Replace(currentlocation, "%5D", "]")
         L2(tL2) = currentlocation
         tL2 = inc(tL2)
         Dim k As Long
         For k = 0 To dec(tL1)
            If currentlocation = L1(k) Then GoTo busysignal
         Next k
        ' scanfolder currentlocation
        MsgBox currentlocation, vbSystemModal, "ojanblank"
End If
busysignal:
    Next
    Erase L1
    tL1 = 0
    For k = 0 To dec(tL2)
        L1(k) = L2(k)
        tL1 = inc(tL1)
    Next k
    End If
TheEnd:
End Sub
Private Function inc(ByVal a As Long) As Long
a = a + 1
End Function
Private Function dec(ByVal a As Long) As Long
a = a - 1
End Function


