Attribute VB_Name = "basBFF"
Private Declare Function lstrcat Lib _
    "kernel32" Alias "lstrcatA" _
    (ByVal lpString1 As String, _
    ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib _
    "shell32" (lpBI As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib _
    "shell32" (ByVal pidList As Long, _
    ByVal lpBuffer As String) As Long
Private Declare Sub CoTaskMemFree Lib _
    "ole32.dll" (ByVal hMem As Long)
    
Private Type BrowseInfo
    lnghWnd As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Const BIF_NEWDIALOGSTYLE As Long = &H40
Private Const BIF_EDITBOX As Long = &H10
Private Const MAX_PATH As Integer = 260
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400


Public Function BrowseForFolder(ByVal hWndOwner As Long, _
    ByVal strPrompt As String) As String
    
    On Error GoTo ErrHandle
    
    Dim intNull As Integer
    Dim lngIDList As Long, lngResult As Long
    Dim strPath As String
    Dim udtBI As BrowseInfo
    
    With udtBI
        .lnghWnd = hWndOwner
        .lpszTitle = lstrcat(strPrompt, "")
        .ulFlags = BIF_NEWDIALOGSTYLE + BIF_EDITBOX
    End With
    
    lngIDList = SHBrowseForFolder(udtBI)
    
    If lngIDList <> 0 Then
        strPath = String(MAX_PATH, 0)
        lngResult = SHGetPathFromIDList(lngIDList, _
            strPath)
        Call CoTaskMemFree(lngIDList)
        intNull = InStr(strPath, vbNullChar)
            If intNull > 0 Then
                strPath = Left(strPath, intNull - 1)
            End If
    End If
     
    BrowseForFolder = strPath
    
    Exit Function
    
ErrHandle:
    BrowseForFolder = Empty
    
End Function




