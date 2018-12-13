Attribute VB_Name = "Mod_DiskIO"
Option Explicit

Private Type FILETIME
    dwLowDateTime     As Long
    dwHighDateTime    As Long
End Type
Private Type SYSTEMTIME
    wYear             As Integer
    wMonth            As Integer
    wDayOfWeek        As Integer
    wDay              As Integer
    wHour             As Integer
    wMinute           As Integer
    wSecond           As Integer
    wMilliseconds     As Integer
End Type

Private Const FILE_SHARE_READ   As Long = &H1
Private Const FILE_SHARE_WRITE  As Long = &H2
Private Const OPEN_EXISTING     As Long = &H3
Private Const GENERIC_WRITE     As Long = &H40000000
Private Const GENERIC_READ      As Long = &H80000000

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal NoSecurity As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long

Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function DosDateTimeToFileTime Lib "kernel32" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFileTime As FILETIME) As Long
Private Declare Function FileTimeToDosDateTime Lib "kernel32" (lpFileTime As FILETIME, ByVal lpFatDate As Long, ByVal lpFatTime As Long) As Long

Private Const m_LocalTimes As Boolean = True
Private PathDilimiter As String

Public Function GetFileDate(FileName As String, FDate As Integer, FTime As Integer) As Boolean
    Dim hFile         As Long     'Get file created/modified/access times
    Dim fCreated      As FILETIME 'Returns True on success
    Dim fModified     As FILETIME
    Dim fAccessed     As FILETIME 'Note:  Accessing a file with this function
    Dim FilTime       As FILETIME '          will modify its File Access Time
    Dim SysTime       As SYSTEMTIME
    Dim FD As Long
    Dim FT As Long

    hFile = CreateFile(FileName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    If hFile <> 0 Then
        GetFileDate = (GetFileTime(hFile, fCreated, fAccessed, fModified) <> 0)
        CloseHandle hFile

        If m_LocalTimes Then                    'Convert FILETIMEs to Local
            FileTimeToLocalFileTime fModified, FilTime
            fModified = FilTime
        End If
        FileTimeToSystemTime fModified, SysTime 'Convert FILETIMEs to Dates
        With SysTime
            FD = (.wYear - 1980) * 2 ^ 9
            FD = FD + (.wMonth * 2 ^ 5)
            FD = FD + .wDay
            FT = .wHour * 2 ^ 11
            FT = FT + (.wMinute * 2 ^ 5)
            FT = FT + .wSecond
        End With
    End If
    If FD > 32767 Then FDate = FD - &HFFFF& - 1 Else FDate = FD
    If FT > 32767 Then FTime = FT - &HFFFF& - 1 Else FTime = FT
End Function

Public Function SetFileDate(FileName As String, dCreated As Date, dModified As Date, dAccessed As Date) As Boolean
    Dim hFile         As Long
    Dim fCreated      As FILETIME
    Dim fModified     As FILETIME
    Dim fAccessed     As FILETIME
    Dim FilTime       As FILETIME
    Dim SysTime       As SYSTEMTIME

    With SysTime                           'Convert Dates to FILETIMEs
        .wYear = Year(dCreated)
        .wMonth = Month(dCreated)
        .wDay = Day(dCreated)
        .wHour = Hour(dCreated)
        .wMinute = Minute(dCreated)
        .wSecond = Second(dCreated)
    End With
    SystemTimeToFileTime SysTime, fCreated

    With SysTime
        .wYear = Year(dModified)
        .wMonth = Month(dModified)
        .wDay = Day(dModified)
        .wHour = Hour(dModified)
        .wMinute = Minute(dModified)
        .wSecond = Second(dModified)
    End With
    SystemTimeToFileTime SysTime, fModified

    With SysTime
        .wYear = Year(dAccessed)
        .wMonth = Month(dAccessed)
        .wDay = Day(dAccessed)
        .wHour = Hour(dAccessed)
        .wMinute = Minute(dAccessed)
        .wSecond = Second(dAccessed)
    End With
    SystemTimeToFileTime SysTime, fAccessed

    If m_LocalTimes Then                    'Convert FILETIMEs from Local
        LocalFileTimeToFileTime fCreated, FilTime
        fCreated = FilTime
        LocalFileTimeToFileTime fModified, FilTime
        fModified = FilTime
        LocalFileTimeToFileTime fAccessed, FilTime
        fAccessed = FilTime
    End If

    hFile = CreateFile(FileName, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    If hFile <> 0 Then
        SetFileDate = (SetFileTime(hFile, fCreated, fAccessed, fModified) <> 0)
        CloseHandle hFile
    End If

End Function

Public Function SetUnZippedFileDate(FileName As String, FDate As Integer, FTime As Integer) As Boolean
    Dim dCreated      As Date
    Dim dModified     As Date
    Dim dAccessed     As Date
    dModified = GetZipDate(FDate, FTime)
    dCreated = dModified
    dAccessed = Now
    SetUnZippedFileDate = SetFileDate(FileName, dCreated, dModified, dAccessed)
End Function

Public Function GetZipDate(FDate As Integer, FTime As Integer) As Date
    Dim fModified     As FILETIME
    Dim SysTime       As SYSTEMTIME

    DosDateTimeToFileTime FDate, FTime, fModified
    FileTimeToSystemTime fModified, SysTime
    With SysTime
    GetZipDate = CDate(Format$(.wMonth) & "/" & _
                       Format$(.wDay) & "/" & _
                       Format$(.wYear) & " " & _
                       Format$(.wHour) & ":" & _
                       Format$(.wMinute, "00") & ":" & _
                       Format$(.wSecond, "00"))
    End With

End Function

'This function is used to write a file
'It will overwrite existing file without prompting
'It sets the filedate and time and checks if the directories exist
Public Function Write_File(FileName As String, _
                            PathName As String, _
                            Data() As Byte, _
                            FDate As Integer, _
                            FTime As Integer) As Integer
    Dim FLnum As Long
    Dim TotName As String
    If PathDilimiter = "" Then PathDilimiter = GetPathdilimiter
    If Right(PathName, 1) <> "\" And Right(PathName, 1) <> "/" Then PathName = PathName & PathDilimiter
    TotName = PathName & FileName
    If CreatePath(mbStripDirName(TotName)) = False Then
'room for error message
    End If
    If Dir(TotName, vbNormal) <> "" Then
        On Error Resume Next
        Kill TotName
    End If
    FLnum = FreeFile
    Open TotName For Binary Access Write As #FLnum
    Put #FLnum, , Data()
    Close FLnum
    If FDate <> 0 Or FTime <> 0 Then
        If SetUnZippedFileDate(TotName, FDate, FTime) = False Then
    'room for error message
        End If
    End If
End Function

Public Function CreatePath(ByVal DestPath$) As Boolean
    Dim BackPos As Integer, ForePos As Integer
    Dim Temp$
    Dim TMP$
    Dim ThisDir As String

    If PathDilimiter = "" Then PathDilimiter = GetPathdilimiter
    DestPath$ = Replace(DestPath$, "\", PathDilimiter)  'set dilimiters in the right direction
    DestPath$ = Replace(DestPath$, "/", PathDilimiter)  'set dilimiters in the right direction
    '-------------------------------------------------------
    '- Add slash to end of path if not there already
    '-------------------------------------------------------
    If Right$(DestPath$, 1) <> PathDilimiter Then DestPath$ = DestPath$ + PathDilimiter

    '-------------------------------------------------------
    '- Quick check on existance if destination path
    '-------------------------------------------------------
    Temp = Dir(Left(DestPath$, Len(DestPath$) - 1), vbDirectory)
    If Temp <> "" Then CreatePath = True: Exit Function

    ThisDir = CurDir$
    '-------------------------------------------------------
    '- Change to the root dir of the drive
    '-------------------------------------------------------
    On Error Resume Next
    ChDrive DestPath$
    If Err <> 0 Then GoTo errorOut

    ChDir PathDilimiter

    '-------------------------------------------------------
    '- Attempt to make each directory, then change to it
    '-------------------------------------------------------
    BackPos = 3
    ForePos = InStr(4, DestPath$, PathDilimiter)
    Do While ForePos <> 0
        Temp$ = Mid$(DestPath$, BackPos + 1, ForePos - BackPos - 1)
        TMP = Dir(Temp$, vbDirectory)
        If TMP = "" Then
            Err = 0
            MkDir Temp$
            If Err <> 0 And Err <> 75 Then GoTo errorOut
        End If
        Err = 0
        ChDir Temp$
        If Err <> 0 Then GoTo errorOut
        BackPos = ForePos
        ForePos = InStr(BackPos + 1, DestPath$, PathDilimiter)
    Loop
    ChDir ThisDir
    CreatePath = True
    Exit Function

errorOut:
    MsgBox "Error While Attempting to Create Directories on Destination Drive.", 48, "SETUP"
    ChDir ThisDir
    CreatePath = False
End Function

'----------------------------------------------------------
'This function is used to retrieve a pathname from a filename
'input:
'Stripfile = Filename with or without pathname
'return:
'StripDirName = Pathname
'----------------------------------------------------------
Private Function mbStripDirName(Stripfile As String) As String

    Dim Counter As Integer, Stripped As String
    On Error Resume Next
    If PathDilimiter = "" Then PathDilimiter = GetPathdilimiter
    Stripfile = Replace(Stripfile, "\", PathDilimiter)  'set dilimiters in the right direction
    Stripfile = Replace(Stripfile, "/", PathDilimiter)  'set dilimiters in the right direction

    If InStr(Stripfile, PathDilimiter) > 0 Then
        For Counter = Len(Stripfile) To 1 Step -1
            If Mid$(Stripfile, Counter, 1) = PathDilimiter Then
                Stripped = Left$(Stripfile, Counter)
                Exit For
            End If
        Next Counter
    ElseIf InStr(Stripfile, ":") = 2 Then
        Stripped = CurDir$(Stripfile)
        If Len(Stripped) = 0 Then
            Stripped = CurDir$
        End If
    Else
        Stripped = CurDir$
    End If

    If Right$(Stripped, 1) <> PathDilimiter Then
        Stripped = Stripped + PathDilimiter
    End If
    mbStripDirName = UCase(Stripped)
End Function

Private Function GetPathdilimiter() As String
    Dim Temp As String
    Temp = CurDir$
    If InStr(Temp, "\") > 0 Then
        GetPathdilimiter = "\"
    Else
        GetPathdilimiter = "/"
    End If
End Function

