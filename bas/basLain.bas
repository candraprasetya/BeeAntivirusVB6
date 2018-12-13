Attribute VB_Name = "basLain"
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal CSIDL As Long, ByVal fCreate As Boolean) As Boolean
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function GetDriveType& Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Public Declare Function ShellExecuteEx Lib "shell32.dll" (SEI As SHELLEXECUTEINFO) As Long
Declare Function MessageBoxW Lib "user32" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal wType As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameW" (ByVal lpszLongPath As Long, ByVal lpszShortPath As Long, ByVal cchBuffer As Long) As Long
Private Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameW" (ByVal lpszShortPath As Long, ByVal lpszLongPath As Long, ByVal cchBuffer As Long) As Long
Private Const INVALID_ARG As Long = 5
Private Const EMPTY_STR As String = "Empty string passed."

Public Enum IDFolder
    ALL_USER_STARTUP = &H18
    WINDOWS_DIR = &H24
    SYSTEM_DIR = &H25
    PROGRAM_FILE = &H26
    USER_DOC = &H5
    USER_STARTUP = &H7
    RECENT_DOC = &H8
    DEKSTOP_PATH = &H19
End Enum

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

' Konstanta peletakan form
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400

Public Sub PlaySound(File As String)
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    Svar = sndPlaySound(App.Path & "\sounds\" & File & ".wav", wFlags%) 'Send the sound to the big world
End Sub

Public Function GetSpecFolder(ByVal lpCSIDL As IDFolder) As String

Dim sPath As String
Dim lRet As Long
    
    sPath = String$(255, 0)
    
    lRet = SHGetSpecialFolderPath(0&, sPath, lpCSIDL, False)
    
    If lRet <> 0 Then
        GetSpecFolder = FixBuffer(sPath)
    End If
    
End Function

Private Function FixBuffer(ByVal sBuffer As String) As String

Dim NullPos As Long
    
    NullPos = InStr(sBuffer, Chr$(0))
    
    If NullPos > 0 Then
        FixBuffer = Left$(sBuffer, NullPos - 1)
    End If
    
End Function

Function ConvertSeconds(Seconds) As String     'As Date
Dim tm   As String
    tm = Format(Int(Seconds / 60), "00") & ":" & Format(Seconds Mod 60, "00")
    ConvertSeconds = tm
End Function

Public Function GetHHMMSS(ByVal ms As Long) As String
    sg = Int(ms / 1000)
    mn = Int(sg / 60)
    hh = Fix(mn / 60)
    If mn > 59 Then
        mn = mn Mod 60
    End If
    zz = sg Mod 60
    GetHHMMSS = Format(Str$(hh), "0#") + ":" + Format(Str$(mn), "0#") + ":" + Format(Str$(zz), "0#")
End Function

Public Function GetTargetLink(ByRef TheFullPath As String, ByVal WithArgumen As Boolean) As String

    Dim LinkShell As New wshShell
    Dim LinkShortCut As WshShortcut
    Set LinkShortCut = LinkShell.CreateShortcut(TheFullPath)

   GetTargetLink = LinkShortCut.TargetPath
   ' klo ada argumen VBS ambil argumenya aja
   If WithArgumen = True Then
     If UCase(Left(LinkShortCut.Arguments, 12)) = "//E:VBSCRIPT" Then
        GetTargetLink = ArgumenToPath(LinkShortCut.Arguments, TheFullPath)
     End If
   End If

Set LinkShell = Nothing
Set LinkShortCut = Nothing
End Function

Private Function ArgumenToPath(ByRef sArgumenFull As String, ByRef ScPath As String) As String
If InStr(sArgumenFull, Chr(34)) > 0 Then
   ArgumenToPath = Mid(sArgumenFull, InStr(sArgumenFull, " ") + 1)
   ArgumenToPath = Left(ArgumenToPath, InStr(ArgumenToPath, " ") - 1)
Else
   ArgumenToPath = Mid(sArgumenFull, InStr(sArgumenFull, " ") + 1)
End If
   If Mid(ArgumenToPath, 2, 1) <> ":" Then ArgumenToPath = GetFilePath(ScPath) & "\" & ArgumenToPath

End Function

Public Function LetakanForm(Frm As Form, bTopMost As Boolean)
   If bTopMost Then
       Call SetWindowPos(Frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
   Else
       Call SetWindowPos(Frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
   End If
End Function

Public Sub ShowProperties(FileName As String, OwnerhWnd As Long)
On Error Resume Next
    Dim SEI As SHELLEXECUTEINFO
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hwnd = OwnerhWnd
        .lpVerb = "properties"
        .lpFile = FileName
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = App.hInstance
        .lpIDList = 0
    End With
    ShellExecuteEx SEI
End Sub

Public Function BuangSpaceAwal(ByVal sKar As String) As String
If Left$(sKar, 1) = Chr(32) Then
    BuangSpaceAwal = Mid(sKar, 2)
Else
    BuangSpaceAwal = sKar
End If

End Function

Public Function MsgBoxU(sPesan As String, sCaption As String, lType As Long, FrmOwn As Form)
    MessageBoxW FrmOwn.hwnd, StrPtr(sPesan), StrPtr(sCaption), lType
End Function

Public Function GetShortFileName(ByVal FileName As String) As String
  Dim rc As Long
  Dim ShortPath As String
  ShortPath = String$(Len(FileName) + 1, 0)
  rc = GetShortPathName(StrPtr(FileName), StrPtr(ShortPath), Len(FileName) + 1)
  GetShortFileName = (Left$(ShortPath, rc))
End Function

Public Function GetLongFileName(sShortPath As String) As String
    If LenB(sShortPath) = ZERO Then Err.Raise INVALID_ARG, , EMPTY_STR

    GetLongFileName = sShortPath

    On Error GoTo GetFailed
    Dim sPath As String
    Dim lResult As Long

    sPath = String$(MAX_PATH, vbNullChar)
    lResult = GetLongPathName(StrPtr(sShortPath), StrPtr(sPath), MAX_PATH)
    If (lResult) Then GetLongFileName = Trim$(sPath)
GetFailed:
End Function

Public Function ReadPathFromCM(GString As String)
Static Ptmp() As String
Static PiTemp As Long
Static lCount As Long
Ptmp = Split(GString, "|")
PiTemp = UBound(Ptmp())
ReDim PathContextScan(PiTemp) As String
For lCount = 0 To PiTemp
    PathContextScan(lCount) = Ptmp(lCount)
Next
End Function
