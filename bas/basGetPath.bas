Attribute VB_Name = "basGetPath"
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Public Enum PathTypes
    FileName = 1
    JustName = 2
    FileExtension = 3
    FilePath = 4
    Drive = 5
    LastFolder = 6
    FirstFolder = 7
    LastFolderAndFileName = 8
    DriveAndFirstFolder = 9
    Fullpath = 10
End Enum

Public Function GetPath(ByVal Path As String, Optional ByVal PathType As PathTypes = 1) As String
Dim strPath As String
Dim ThisType As PathTypes
Dim i As Integer
Dim j As Integer

strPath = Path

If InStr(strPath, "\") = 0 And InStr(strPath, ".") > 0 And InStr(strPath, ":") = 0 Then
    ThisType = FileName
ElseIf InStrRev(strPath, "\") = Len(strPath) And Len(strPath) > 3 Then
    ThisType = FilePath
ElseIf Len(strPath) = 3 And Mid(strPath, 2, 2) = ":\" Then
    ThisType = Drive
ElseIf Len(strPath) = 2 And Mid(strPath, 2, 1) = ":" Then
    ThisType = Drive
ElseIf InStrRev(strPath, "\") > InStrRev(strPath, ".") Then
    ThisType = JustName
ElseIf InStr(strPath, "\") > 0 And InStr(strPath, ".") > 0 Then
    ThisType = Fullpath
Else
'    MsgBox "Cannot determine the type of the path"
    Exit Function
End If

Select Case PathType
    Case 1
        If ThisType = Fullpath Or ThisType = JustName Then
            GetPath = Right(strPath, Len(strPath) - InStrRev(strPath, "\"))
        ElseIf ThisType = FileName Then
            GetPath = strPath
        End If
    Case 2
        If ThisType = Fullpath Then
            strPath = StrReverse(strPath)
            i = InStr(strPath, ".") + 1
            j = InStr(strPath, "\")
            strPath = Mid(strPath, i, j - i)
            GetPath = StrReverse(strPath)
        ElseIf ThisType = FileName Then
            GetPath = Left(strPath, InStrRev(strPath, ".") - 1)
        ElseIf ThisType = JustName Then
            GetPath = Right(strPath, Len(strPath) - InStrRev(strPath, "\"))
        End If
    Case 3
        If ThisType = Fullpath Or ThisType = FileName Then
            GetPath = Right(strPath, Len(strPath) - InStrRev(strPath, "."))
        End If
    Case 4
        If ThisType = Fullpath Or ThisType = JustName Then
            strPath = Left(strPath, InStrRev(strPath, "\") - 1)
        ElseIf ThisType = FilePath Then
            strPath = Left(strPath, Len(strPath) - 1)
        End If
        If Left(strPath, 1) = "\" Then
            strPath = Right(strPath, Len(strPath) - 1)
        End If
        GetPath = strPath
    Case 5
        If ThisType = FilePath Or ThisType = Fullpath Or ThisType = Drive Or ThisType = JustName Then
            If Mid(strPath, 2, 1) = ":" Then
                GetPath = Left(strPath, 2)
            End If
        End If
    Case 6
        If ThisType = Fullpath Or ThisType = JustName Or ThisType = FilePath Then
            strPath = Left(strPath, InStrRev(strPath, "\") - 1)
            GetPath = Right(strPath, Len(strPath) - InStrRev(strPath, "\"))
        End If
    Case 7
        If Mid(strPath, 2, 1) <> ":" And Left(strPath, 1) <> "\" Then
            strPath = "\" & strPath
        End If
        If ThisType = Fullpath Or ThisType = JustName Or ThisType = FilePath Then
            strPath = Right(strPath, Len(strPath) - InStr(strPath, "\"))
            If InStr(strPath, "\") = 0 Then
                Exit Function
            End If
            GetPath = Left(strPath, InStr(strPath, "\") - 1)
        End If
    Case 8
        If ThisType = Fullpath Or ThisType = JustName Then
            strPath = Left(strPath, InStrRev(strPath, "\") - 1)
            GetPath = Right(strPath, Len(strPath) - InStrRev(strPath, "\"))
            GetPath = GetPath & Right(Path, Len(Path) - InStrRev(Path, "\") + 1)
        End If
    Case 9
        If ThisType = Fullpath Or ThisType = JustName Or ThisType = FilePath Then
            If Mid(strPath, 2, 1) = ":" Then
                strPath = Right(strPath, Len(strPath) - InStr(strPath, "\"))
                GetPath = Left(Path, 3) & Left(strPath, InStr(strPath, "\") - 1)
            End If
        End If
    Case 10
        GetPath = strPath
End Select
End Function

Public Function GetAllDrive()
Static ing As Integer
ing = 0
frMain.drvList.Clear
frMain.lvDlock.ListItems.Clear
GetHardDrive
GetFlashDrive
GetCDDrive
ReDim xDrive(frMain.drvList.ListCount - 1) As String
For ing = 0 To frMain.drvList.ListCount - 1
    xDrive(ing) = frMain.drvList.List(ing)
Next
End Function

Public Function GetQuickPath()
qPath(0) = GetSpecFolder(PROGRAM_FILE)
qPath(1) = GetSpecFolder(WINDOWS_DIR)
qPath(2) = GetSpecFolder(USER_DOC)
End Function

Public Function GetFlashDrive() As String
Dim PKey As Byte
Dim xClItem As cListItem

For PKey = 0 To 23 ' Mulai dari Drive C:\ --> Z:\
    If GetDriveType(Chr(67 + PKey) & ":\") = 2 Then
        GetFlashDrive = Chr(67 + PKey) & ":\"
        frMain.drvList.AddItem GetFlashDrive
        Set xClItem = frMain.lvDlock.ListItems.Add(, GetFlashDrive, , 0)
        xClItem.SubItem(2).Text = CekProtecFrom(GetFlashDrive)
    Else
        GetFlashDrive = "Nothing"
    End If
Next

End Function

Public Function GetFlashDrive2() As String
Dim PKey As Byte
Dim xClItem As cListItem

For PKey = 0 To 23 ' Mulai dari Drive C:\ --> Z:\
    If GetDriveType(Chr(67 + PKey) & ":\") = 2 Then
        GetFlashDrive2 = Chr(67 + PKey) & ":\"
        frMain.drvList.AddItem GetFlashDrive
        Set xClItem = frMain.lvDlock.ListItems.Add(, GetFlashDrive, , 0)
        xClItem.SubItem(2).Text = CekProtecFrom(GetFlashDrive)
    Else
        GetFlashDrive2 = "Nothing"
    End If
Next

End Function

Public Function GetHardDrive() As String
Dim PKey As Byte
Dim xClItem As cListItem

For PKey = 0 To 23 ' Mulai dari Drive C:\ --> Z:\
    If GetDriveType(Chr(67 + PKey) & ":\") = 3 Then
        GetHardDrive = Chr(67 + PKey) & ":\"
        frMain.drvList.AddItem GetHardDrive
        Set xClItem = frMain.lvDlock.ListItems.Add(, GetHardDrive, , 0)
        xClItem.SubItem(2).Text = CekProtecFrom(GetHardDrive)
    Else
        GetHardDrive = "Nothing"
    End If
Next
End Function

Public Function GetCDDrive() As String
Dim PKey As Byte
    
For PKey = 0 To 23 ' Mulai dari Drive C:\ --> Z:\
    If GetDriveType(Chr(67 + PKey) & ":\") = 5 Then
        GetCDDrive = Chr(67 + PKey) & ":\"
        frMain.drvList.AddItem GetCDDrive
    Else
        GetCDDrive = "Nothing"
    End If
Next
End Function

Private Function CekProtecFrom(sNamaDrive As String) As String
Dim xFSO As New FileSystemObject

If xFSO.FolderExists(sNamaDrive & "autorun.inf\con\aux\nul.  This autorun.inf is LOCKED by SMAD" & ChrW(916) & "V to protect your Flash-Disk from virus infection.") = True Then
CekProtecFrom = "Protected"

ElseIf xFSO.FolderExists(sNamaDrive & "autorun.inf\aux.morphost\data\nul. Folder ini sudah diproteksi Morphost Antivirus") = True Then
CekProtecFrom = "Protected"

ElseIf ValidFile3(sNamaDrive & "autorun.inf\aux\con\nul.[Beelock] Folder sudah terproteksi oleh Bee Antivirus [BeeLock].") = True Then
CekProtecFrom = "Protected By Bee Antivirus"

Else
CekProtecFrom = "UnProtected"

End If
End Function
