Attribute VB_Name = "basFix"
Public Enum typFix
    FixAll = 0
    FixChk = 1
    IgnAll = 2
    IgnChk = 3
End Enum

Public Enum FixAct
    del = 0
    Quar = 1
    None = 2
End Enum

Public Enum viStatf
    Error = 0
    Success = 1
    UnSuccess = 2
    Moved = 3
    UnMoved = 4
End Enum
Dim xSXor As classSimpleXOR
Dim xHuff As classHuffman


Public Function FixVir(ByVal xTipe As typFix, Lv As ucListView, ByVal xstatus As FixAct)
Static xViPathF As String
Static xCount As Integer



Select Case xTipe
Case 0
    xCount = 1
    For xCount = 1 To Lv.ListItems.count
        xViPathF = Lv.ListItems.Item(xCount).SubItem(2).Text
        If ValidFile(xViPathF) = False Then
            If InStr(xViPathF, "|>") > 0 Then
                xViPathF = Left(xViPathF, InStr(xViPathF, "|>") - 1)
            End If
        End If
        If ValidFile(xViPathF) = False Then
            MKviStat Error, Lv, xCount
            GoTo SkippedXF
        End If
        Lv.ListItems.Item(xCount).Checked = True
        Lv.ListItems.Item(xCount).Selected = True
        Select Case xstatus
        Case 0
        KillByProccess GetPath(xViPathF, FileName)
            If HapusFile(xViPathF) = True Then
                MKviStat Success, Lv, xCount
            Else
                MKviStat UnSuccess, Lv, xCount
            End If
        DoEvents
        Case 1
        KillByProccess GetPath(xViPathF, FileName)
            If InsPath2File(xViPathF) = True Then
                MKviStat Moved, Lv, xCount
            Else
                MKviStat UnMoved, Lv, xCount
            End If
        DoEvents
        End Select
SkippedXF:
    Next
    Exit Function
Case 1
    xCount = 1
    For xCount = 1 To Lv.ListItems.count
      If Lv.ListItems.Item(xCount).Checked = True Then
        xViPathF = Lv.ListItems.Item(xCount).SubItem(2).Text
        If ValidFile(xViPathF) = False Then
            If InStr(xViPathF, "|>") > 0 Then
                xViPathF = Left(xViPathF, InStr(xViPathF, "|>") - 1)
            End If
        End If
        If ValidFile(xViPathF) = False Then
            MKviStat Error, Lv, xCount
            GoTo SkippedXF2
        End If
        Select Case xstatus
        Case 0
        KillByProccess GetPath(xViPathF, FileName)
            If HapusFile(xViPathF) = True Then
                MKviStat Success, Lv, xCount
            Else
                MKviStat UnSuccess, Lv, xCount
            End If
        DoEvents
        Case 1
        KillByProccess GetPath(xViPathF, FileName)
            If InsPath2File(xViPathF) = True Then
                MKviStat Moved, Lv, xCount
            Else
                MKviStat UnMoved, Lv, xCount
            End If
        DoEvents
        End Select
SkippedXF2:
    End If
    Next
    Exit Function
Case 2
    Lv.ListItems.Clear
Case 3
    xCount = 1
    For xCount = 1 To Lv.ListItems.count
    If xCount > Lv.ListItems.count Then Exit For
      If Lv.ListItems.Item(xCount).Checked = True Then
        Lv.ListItems.Remove (xCount)
      End If
    Next
End Select
End Function

Public Function MKviStat(virtStat As viStatf, Lv As ucListView, itmDX As Integer)
Static xJum As String
xJum = Left(frMain.lInFile.Caption, Len(frMain.lInFile.Caption) - 8)
Select Case virtStat
    Case 0
    Lv.ListItems.Item(itmDX).IconIndex = 3
    Lv.ListItems.Item(itmDX).SubItem(4).Text = "Invalid File !"
    Case 1
    Lv.ListItems.Item(itmDX).IconIndex = 1
    Lv.ListItems.Item(itmDX).SubItem(4).Text = "Success !"
                If xJum - 1 < 0 Then Exit Function
                frMain.lInFile.Caption = xJum - 1 & " File[s]"
                If xJum - 1 = 0 Then
                    frMain.lInFile.ForeColor = vbBlack
                End If
    Case 2
    Lv.ListItems.Item(itmDX).IconIndex = 2
    Lv.ListItems.Item(itmDX).SubItem(4).Text = "Can't Fix it !"
    Case 3
    Lv.ListItems.Item(itmDX).IconIndex = 1
    Lv.ListItems.Item(itmDX).SubItem(4).Text = "Moved To Qurantine"
                If xJum - 1 < 0 Then Exit Function
                frMain.lInFile.Caption = xJum - 1 & " File[s]"
                If xJum - 1 = 0 Then
                    frMain.lInFile.ForeColor = vbBlack
                End If
    Case 3
    Lv.ListItems.Item(itmDX).IconIndex = 2
    Lv.ListItems.Item(itmDX).SubItem(4).Text = "Moved To Qurantine, but Can't Delete Original File"
End Select
End Function

Public Function MKquarDir()
BuatFolder App.Path & "\Quarantine"
End Function

Private Function InsPath2File(xTarget As String) As Boolean
Static IsiFileX As String
Static IsiFileX2 As String
Static NamaFileQuar As String
Static BanyakOldPath As String
Static xDesnt As String
Static Inter As Integer

Inter = 8
BanyakOldPath = Len(xTarget)
Cek_Lge:
NamaFileQuar = Format$(BanyakOldPath, "0000") & Right$(xTarget, Len(xTarget) - InStrRev(xTarget, "\")) & GetRandomPassword(Inter)
xDesnt = App.Path & "\Quarantine\" & NamaFileQuar & ".BeeQuar"
If ValidFile(xDesnt) = True Then GoSub Cek_Lge
CryptVirus xTarget, xDesnt
IsiFileX = ReadUnicodeFile(xDesnt)
IsiFileX2 = xTarget & IsiFileX
IsiFileX2 = StrConv(IsiFileX2, vbUnicode)
WriteFileUniSim xDesnt, IsiFileX2
If HapusFile(xTarget) = True Then
InsPath2File = True
Else
InsPath2File = False
End If
End Function

Public Function CryptVirus(sPath As String, sDestination As String)
Static hFileP As Long, OutConfig() As Byte, IsiFile As String
hFileP = GetHandleFile(sPath)
If ValidFile(sPath) = False Then Exit Function
    Call ReadUnicodeFile2(hFileP, 1, GetSizeFile(hFileP), OutConfig())
    Set xSXor = New classSimpleXOR
    Call xSXor.EncryptByte(OutConfig(), "DAXAAngKOt")
    Set xSXor = Nothing
IsiFile = StrConv(OutConfig, vbUnicode)
WriteFileUniSim sDestination, IsiFile
TutupFile hFileP
End Function

Public Function DeCryptVirus(sPath As String, sDestination As String)
Static hFileP As Long, OutConfig() As Byte, IsiFile As String
hFileP = GetHandleFile(sPath)
If ValidFile(sPath) = False Then Exit Function
    Call ReadUnicodeFile2(hFileP, 1, GetSizeFile(hFileP), OutConfig())
    Set xSXor = New classSimpleXOR
    Call xSXor.DecryptByte(OutConfig(), "DAXAAngKOt")
    Set xSXor = Nothing
IsiFile = StrConv(OutConfig, vbUnicode)
WriteFileUniSim sDestination, IsiFile
TutupFile hFileP
End Function

Public Function KillByProccess(Exe_Name As String)
On Error Resume Next
    Dim ProccessList As Object
    Dim WMI As Object
    Dim Proccess As Object
        Set WMI = GetObject("winmgmts:")
    If IsNull(WMI) = False Then
            Set ProccessList = WMI.InstancesOf("win32_process")
            For Each Proccess In ProccessList
                If UCase(Proccess.name) = UCase(Exe_Name) Then
                Proccess.Terminate (0)
            End If
        Next
    Else
        End If
        Set ProccessList = Nothing
    Set WMI = Nothing
End Function

