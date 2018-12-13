VERSION 5.00
Begin VB.UserControl UniDialog 
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   Picture         =   "UniDialog.ctx":0000
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   32
   ToolboxBitmap   =   "UniDialog.ctx":0282
   Windowless      =   -1  'True
End
Attribute VB_Name = "UniDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************************************
'* UniDialog 0.6 - Unicode common dialog control
'* ---------------------------------------------
'* By Vesa Piittinen aka Merri, http://vesa.piittinen.name/ <vesa@piittinen.name>
'* Unicode on 2000/XP/Vista
'*
'* LICENSE
'* -------
'* http://creativecommons.org/licenses/by-sa/1.0/fi/deed.en
'*
'* Terms: 1) If you make your own version, share using this same license.
'*        2) When used in a program, mention my name in the program's credits.
'*        3) May not be used as a part of commercial (unicode) controls suite.
'*        4) Free for any other commercial and non-commercial usage.
'*        5) Use at your own risk. No support guaranteed.
'*
'* SUPPORT FOR UNICONTROLS
'* -----------------------
'* http://www.vbforums.com/showthread.php?t=500026
'* http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=69738&lngWId=1
'*
'* REQUIREMENTS
'* ------------
'* No special requirements.
'*
'* HOW TO ADD TO YOUR PROGRAM
'* --------------------------
'* 1) Copy UniDialog.ctl and UniDialog.ctx to your project folder.
'* 2) In your project, add UniDialog.ctl.
'*
'* VERSION HISTORY
'* ---------------
'* Version 0.6 BETA (2008-06-19)
'* - initial release: no Color, Font, Help, Printer or Save dialog.
'*************************************************************************************************
Option Explicit

Public Event FolderCancel(ByVal CancelType As UniDialogFolderCancel)
Public Event FolderSelect(ByVal Path As String)
Public Event OpenCancel(ByVal CancelType As UniDialogFileCancel)
Public Event OpenFile(ByVal FileName As String)
Public Event SaveCancel(ByVal CancelType As UniDialogFileCancel)
Public Event SaveFile(ByVal FileName As String)

Private Const WM_USER = &H400

' BrowseForFolder constants
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELECTIONCHANGED = 2
Private Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)

Public Enum UniDialogFileCancel
    [No file error] = &H0
    [Invalid lStructSize] = &H1
    [Initialization failure] = &H2
    [No template] = &H3
    [No instance handle] = &H4
    [Required string load failed] = &H5
    [Required resource not found] = &H6
    [Required resource load failed] = &H7
    [Required resource lock failed] = &H8
    [Memory allocation failure] = &H9
    [Memory lock failure] = &HA
    [No hook] = &HB
    [RegisterWindowMessage failure] = &HC
    [Buffer too small] = &H3003
    [Invalid filename] = &H3002
    [Subclass failure] = &H3001
    [DialogBox failure] = &HFFFF
End Enum

Public Enum UniDialogFileFlags
    [Read Only] = &H1
    [Overwrite Prompt] = &H2
    [Hide Read Only] = &H4
    [No Change Directory] = &H8
    [Help Button] = &H10
    [No Validate] = &H100
    [Allow Multi Select] = &H200
    [Extension Different] = &H400
    [Path Must Exist] = &H800
    [File Must Exist] = &H1000
    [Create Prompt] = &H2000
    [Share Aware] = &H4000
    [No Read Only Return] = &H8000
    [Explorer Style] = &H80000
    [No Dereference Links] = &H100000
    [Long Filenames] = &H200000
End Enum

Public Enum UniDialogFolderCancel
    [No folder error] = &H0
    [Not a path]
End Enum

Public Enum UniDialogFolderFlags
    [Return Only File System Dirs] = &H1
    [Don't Go Below Domain] = &H2
    [Status Text] = &H4&
    [Return File System Ancestors] = &H8
    [Show Edit Box] = &H10
    [Validate Edit Box] = &H20
    [New Dialog Style] = &H40
    [Browse Include URLs] = &H80
    [Usage Hint] = &H100
    [No New Folder Button] = &H200
    [No Translate Targets] = &H400
    [Browse For Computer] = &H1000
    [Browse For Printer] = &H2000
    [Browse Include Files] = &H4000
    [Show Shareable Resources] = &H8000
End Enum

Private Type BrowseInfo
    hWndOwner As Long
    pidlRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfn As Long
    lParam As Long
    lImage As Long
End Type

Private Type OpenFilename
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As Long
    lpstrCustomFilter As Long
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As Long
    nMaxFile As Long
    lpstrFileTitle As Long
    nMaxFileTitle As Long
    lpstrInitialDir As Long
    lpstrTitle As Long
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type

' properties
Private m_FileCustomFilter As String
Private m_FileDefaultExtension As String
Private m_FileFilter As String
Private m_FileFlags As Long
Private m_Filename As String
Private m_FileInitialDirectory As String
Private m_FileOpenTitle As String
Private m_FileSaveTitle As String
Private m_FileTemplateName As String
Private m_FileTitle As String
Private m_FolderFlags As Long
Private m_FolderMessage As String
Private m_FolderPath As String

' private
Private m_Buffer As String
Private m_FileDialog As OpenFilename
Private m_FolderDialog As BrowseInfo

' BrowseForFolder
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SHBrowseForFolderW Lib "shell32" (lpBrowseInfo As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDListW Lib "shell32" (ByVal pidl As Long, ByVal pszPath As Long) As Long

' CommonDialog
Private Declare Function CommDlgExtendedError Lib "comdlg32" () As Long
Private Declare Function GetOpenFileNameW Lib "comdlg32" (pOpenFilename As OpenFilename) As Long
Private Declare Function GetSaveFileNameW Lib "comdlg32" (pOpenFilename As OpenFilename) As Long
Private Declare Function SendMessageW Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Property Get AllowInvalidChars() As Boolean
    AllowInvalidChars = (m_FileFlags And &H100&) = &H100&
End Property
Public Property Let AllowInvalidChars(ByVal NewValue As Boolean)
    If NewValue Then
        m_FileFlags = m_FileFlags Or &H100&
    Else
        m_FileFlags = m_FileFlags And Not &H100&
    End If
End Property
Public Property Get CreatePrompt() As Boolean
    CreatePrompt = (m_FileFlags And &H2000&) = &H2000&
End Property
Public Property Let CreatePrompt(ByVal NewValue As Boolean)
    If NewValue Then
        m_FileFlags = m_FileFlags Or &H2000&
    Else
        m_FileFlags = m_FileFlags And Not &H2000&
    End If
End Property
Public Property Get ExplorerStyle() As Boolean
    ExplorerStyle = (m_FileFlags And &H80000) = &H80000
End Property
Public Property Let ExplorerStyle(ByVal NewValue As Boolean)
    If NewValue Then
        m_FileFlags = m_FileFlags Or &H80000
    Else
        m_FileFlags = m_FileFlags And Not &H80000
    End If
End Property
Public Property Get ExtensionDifferent() As Boolean
    ExtensionDifferent = (m_FileFlags And &H400&) = &H400&
End Property
Public Property Let ExtensionDifferent(ByVal NewValue As Boolean)
    If NewValue Then
        m_FileFlags = m_FileFlags Or &H400&
    Else
        m_FileFlags = m_FileFlags And Not &H400&
    End If
End Property
Public Property Get FileCustomFilter() As String
    FileCustomFilter = Replace(m_FileCustomFilter, vbNullChar, "|")
End Property
Public Property Let FileCustomFilter(ByVal NewValue As String)
    m_FileCustomFilter = Replace(NewValue, "|", vbNullChar)
    m_FileDialog.lpstrCustomFilter = StrPtr(m_FileCustomFilter)
End Property
Public Property Get FileDefaultExtension() As String
    FileDefaultExtension = m_FileDefaultExtension
End Property
Public Property Let FileDefaultExtension(ByVal NewValue As String)
    m_FileDefaultExtension = NewValue
    m_FileDialog.lpstrDefExt = StrPtr(m_FileDefaultExtension)
End Property
Public Property Get FileFilter() As String
    FileFilter = Replace(m_FileFilter, vbNullChar, "|")
End Property
Public Property Let FileFilter(ByVal NewValue As String)
    m_FileFilter = Replace(NewValue, "|", vbNullChar)
    If AscW(Right$(m_FileFilter, 1)) <> 0& Then m_FileFilter = m_FileFilter & vbNullChar
    m_FileDialog.lpstrFilter = StrPtr(m_FileFilter)
End Property
Public Property Get FileFlags() As UniDialogFileFlags
Attribute FileFlags.VB_MemberFlags = "400"
    FileFlags = m_FileFlags
End Property
Public Property Let FileFlags(ByVal NewValue As UniDialogFileFlags)
    m_FileFlags = NewValue
End Property
Public Property Get FileInitialDirectory() As String
    FileInitialDirectory = m_FileInitialDirectory
End Property
Public Property Let FileInitialDirectory(ByVal NewValue As String)
    m_FileInitialDirectory = NewValue
    m_FileDialog.lpstrInitialDir = StrPtr(m_FileInitialDirectory)
End Property
Public Property Get FileMustExist() As Boolean
    FileMustExist = (m_FileFlags And &H1000&) = &H1000&
End Property
Public Property Let FileMustExist(ByVal NewValue As Boolean)
    If NewValue Then
        m_FileFlags = m_FileFlags Or &H1000&
    Else
        m_FileFlags = m_FileFlags And Not &H1000&
    End If
End Property
Public Property Get FileName() As String
Attribute FileName.VB_MemberFlags = "600"
    FileName = m_Filename
End Property
Public Property Let FileName(ByVal NewValue As String)
    m_Filename = NewValue
    m_FileDialog.lpstrFile = StrPtr(m_Filename)
End Property
Public Property Get FileOpenTitle() As String
    FileOpenTitle = m_FileOpenTitle
End Property
Public Property Let FileOpenTitle(ByRef NewValue As String)
    m_FileOpenTitle = NewValue
End Property
Public Property Get FileSaveTitle() As String
    FileSaveTitle = m_FileSaveTitle
End Property
Public Property Let FileSaveTitle(ByRef NewValue As String)
    m_FileSaveTitle = NewValue
End Property
Public Property Get FileTemplateName() As String
    FileTemplateName = m_FileTemplateName
End Property
Public Property Let FileTemplateName(ByVal NewValue As String)
    m_FileTemplateName = NewValue
    m_FileDialog.lpTemplateName = StrPtr(m_FileTemplateName)
End Property
Public Property Get FileTitle() As String
    FileTitle = m_FileTitle
End Property
Public Property Let FileTitle(ByVal NewValue As String)
    m_FileTitle = NewValue
    m_FileDialog.lpstrFileTitle = StrPtr(m_FileTitle)
End Property
Public Property Get FolderFileSystemOnly() As Boolean
    FolderFileSystemOnly = (m_FolderFlags And &H1&) = &H1&
End Property
Public Property Let FolderFileSystemOnly(ByVal NewValue As Boolean)
    If NewValue Then
        m_FolderFlags = m_FolderFlags Or &H1&
    Else
        m_FolderFlags = m_FolderFlags And Not &H1&
    End If
End Property
Public Property Get FolderFlags() As UniDialogFolderFlags
    FolderFlags = m_FolderFlags
End Property
Public Property Let FolderFlags(ByVal NewValue As UniDialogFolderFlags)
    m_FolderFlags = NewValue
End Property
Public Property Get FolderMessage() As String
    FolderMessage = m_FolderMessage
End Property
Public Property Let FolderMessage(ByRef NewValue As String)
    m_FolderMessage = NewValue
    m_FolderDialog.lpszTitle = StrPtr(NewValue)
End Property
Public Property Get FolderPath() As String
    FolderPath = m_FolderPath
End Property
Public Property Let FolderPath(ByRef NewValue As String)
    m_FolderPath = NewValue
    m_FolderDialog.pszDisplayName = StrPtr(NewValue)
End Property
Public Property Get FolderShowEditBox() As Boolean
    FolderShowEditBox = (m_FolderFlags And &H20&) = &H20&
End Property
Public Property Let FolderShowEditBox(ByVal NewValue As Boolean)
    If NewValue Then
        m_FolderFlags = m_FolderFlags Or &H20&
    Else
        m_FolderFlags = m_FolderFlags And Not &H20&
    End If
End Property
Public Property Get FolderShowFiles() As Boolean
    FolderShowFiles = (m_FolderFlags And &H4000&) = &H4000&
End Property
Public Property Let FolderShowFiles(ByVal NewValue As Boolean)
    If NewValue Then
        m_FolderFlags = m_FolderFlags Or &H4000&
    Else
        m_FolderFlags = m_FolderFlags And Not &H4000&
    End If
End Property
Public Property Get FolderShowNewButton() As Boolean
    FolderShowNewButton = (m_FolderFlags And &H200&) = 0&
End Property
Public Property Let FolderShowNewButton(ByVal NewValue As Boolean)
    If NewValue Then
        m_FolderFlags = m_FolderFlags And Not &H200&
    Else
        m_FolderFlags = m_FolderFlags Or &H200&
    End If
End Property
Public Property Get FolderShowShareable() As Boolean
    FolderShowShareable = (m_FolderFlags And &H8000&) = &H8000&
End Property
Public Property Let FolderShowShareable(ByVal NewValue As Boolean)
    If NewValue Then
        m_FolderFlags = m_FolderFlags Or &H8000&
    Else
        m_FolderFlags = m_FolderFlags And Not &H8000&
    End If
End Property
Public Property Get HelpButton() As Boolean
    HelpButton = (m_FileFlags And &H10&) = &H10&
End Property
Public Property Let HelpButton(ByVal NewValue As Boolean)
    If NewValue Then
        m_FileFlags = m_FileFlags Or &H10&
    Else
        m_FileFlags = m_FileFlags And Not &H10&
    End If
End Property
Public Property Get HideReadOnly() As Boolean
    HideReadOnly = (m_FileFlags And &H4&) = &H4&
End Property
Public Property Let HideReadOnly(ByVal NewValue As Boolean)
    If NewValue Then
        m_FileFlags = m_FileFlags Or &H4&
    Else
        m_FileFlags = m_FileFlags And Not &H4&
    End If
End Property
Public Property Get IgnoreShareErrors() As Boolean
    IgnoreShareErrors = (m_FileFlags And &H4000&) = &H4000&
End Property
Public Property Let IgnoreShareErrors(ByVal NewValue As Boolean)
    If NewValue Then
        m_FileFlags = m_FileFlags Or &H4000&
    Else
        m_FileFlags = m_FileFlags And Not &H4000&
    End If
End Property
Public Property Get LongFilenames() As Boolean
    LongFilenames = (m_FileFlags And &H200000) = &H200000
End Property
Public Property Let LongFilenames(ByVal NewValue As Boolean)
    If NewValue Then
        m_FileFlags = m_FileFlags Or &H200000
    Else
        m_FileFlags = m_FileFlags And Not &H200000
    End If
End Property
Public Property Get MultiSelect() As Boolean
    MultiSelect = (m_FileFlags And &H200&) = &H200&
End Property
Public Property Let MultiSelect(ByVal NewValue As Boolean)
    If NewValue Then
        m_FileFlags = m_FileFlags Or &H200&
    Else
        m_FileFlags = m_FileFlags And Not &H200&
    End If
End Property
Public Property Get NoChangeDir() As Boolean
    NoChangeDir = (m_FileFlags And &H8&) = &H8&
End Property
Public Property Let NoChangeDir(ByVal NewValue As Boolean)
    If NewValue Then
        m_FileFlags = m_FileFlags Or &H8&
    Else
        m_FileFlags = m_FileFlags And Not &H8&
    End If
End Property
Public Property Get NoDereferenceLinks() As Boolean
    NoDereferenceLinks = (m_FileFlags And &H100000) = &H100000
End Property
Public Property Let NoDereferenceLinks(ByVal NewValue As Boolean)
    If NewValue Then
        m_FileFlags = m_FileFlags Or &H100000
    Else
        m_FileFlags = m_FileFlags And Not &H100000
    End If
End Property
Public Property Get NoReadOnlyFiles() As Boolean
    NoReadOnlyFiles = (m_FileFlags And &H8000&) = &H8000&
End Property
Public Property Let NoReadOnlyFiles(ByVal NewValue As Boolean)
    If NewValue Then
        m_FileFlags = m_FileFlags Or &H8000&
    Else
        m_FileFlags = m_FileFlags And Not &H8000&
    End If
End Property
Public Property Get PathMustExist() As Boolean
    PathMustExist = (m_FileFlags And &H800&) = &H800&
End Property
Public Property Let PathMustExist(ByVal NewValue As Boolean)
    If NewValue Then
        m_FileFlags = m_FileFlags Or &H800&
    Else
        m_FileFlags = m_FileFlags And Not &H800&
    End If
End Property
Private Sub Private_Init()
    With m_FileDialog
        .lStructSize = Len(m_FileDialog)
        .hInstance = App.hInstance
        .hWndOwner = Parent.hwnd
    End With
    m_FolderDialog.hWndOwner = m_FileDialog.hWndOwner
End Sub
Public Property Get ReadOnly() As Boolean
    ReadOnly = (m_FileFlags And &H1&) = &H1&
End Property
Public Property Let ReadOnly(ByVal NewValue As Boolean)
    If NewValue Then
        m_FileFlags = m_FileFlags Or &H1&
    Else
        m_FileFlags = m_FileFlags And Not &H1&
    End If
End Property
Public Property Get SavePromptOverwrite() As Boolean
    SavePromptOverwrite = (m_FileFlags And &H2&) = &H2&
End Property
Public Property Let SavePromptOverwrite(ByVal NewValue As Boolean)
    If NewValue Then
        m_FileFlags = m_FileFlags Or &H2&
    Else
        m_FileFlags = m_FileFlags And Not &H2&
    End If
End Property
Public Function ShowColor(Optional ByRef Color As Long) As Boolean

End Function
Public Function ShowFolder(Optional ByRef Path As String) As Boolean
    Dim lngPathMem As Long, lngPidl As Long, lngPos As Long, strBuffer As String * 260, strPath As String
    ' make another path buffer
    strPath = strBuffer
    ' fill the other buffer?
    If LenB(Path) Then strBuffer = Path Else strBuffer = m_FolderPath
    ' allocate string
    lngPathMem = LocalAlloc(&H40&, Len(strBuffer) + 1)
    CopyMemory ByVal lngPathMem, ByVal strBuffer, Len(strBuffer) + 1
    ' set generic settings
    With m_FolderDialog
        .lParam = lngPathMem
        .lpszTitle = StrPtr(m_FolderMessage)
        .ulFlags = m_FolderFlags
    End With
    ' attempt to get a folder
    lngPidl = SHBrowseForFolderW(m_FolderDialog)
    ' null if no selection
    If lngPidl Then
        ' get the path from pidl
        If SHGetPathFromIDListW(lngPidl, StrPtr(strPath)) Then
            ' determine string length
            lngPos = InStr(strPath, vbNullChar)
            If lngPos > 0 Then
                m_FolderPath = Left$(strPath, lngPos - 1)
            Else
                m_FolderPath = strPath
            End If
            RaiseEvent FolderSelect(m_FolderPath)
            ' success!
            ShowFolder = True
        Else
            ' the pidl was not a path
            RaiseEvent FolderCancel([Not a path])
        End If
        CoTaskMemFree lngPidl
    Else
        RaiseEvent FolderCancel([No folder error])
    End If
    LocalFree lngPathMem
End Function
Public Function ShowFont(Optional ByRef Font As Font) As Boolean

End Function
Public Function ShowHelp() As Boolean

End Function
Public Function ShowOpen(Optional ByRef FileName As String) As Boolean
    Dim strBuffer As String, strFiles() As String
    Dim lngA As Long
    With m_FileDialog
        .Flags = m_FileFlags
        ' set window title
        .lpstrTitle = StrPtr(m_FileOpenTitle)
        ' prepare string buffer
        strBuffer = FileName & m_Buffer
        .lpstrFile = StrPtr(strBuffer)
        .nMaxFile = Len(strBuffer)
        ' show the dialog
        ShowOpen = GetOpenFileNameW(m_FileDialog) <> 0&
        If ShowOpen Then
            ' remove extra data from the buffer?
            lngA = InStr(strBuffer, vbNullChar & vbNullChar)
            If lngA = 1 Then
                ' fail!
                Exit Function
            ElseIf lngA > 1 Then
                ' remove extra
                strFiles = Split(Left$(strBuffer, lngA - 1), vbNullChar)
            Else
                ' buffer was fully filled...
                strFiles = Split(strBuffer, vbNullChar)
            End If
            ' now we have the number of files...
            lngA = UBound(strFiles)
            If lngA = 0 Then
                ' one file
                FileName = strBuffer
                RaiseEvent OpenFile(strBuffer)
            ElseIf Right$(strFiles(0), Len(strFiles(1))) = strFiles(1) Then
                ' one file
                FileName = strFiles(0)
                RaiseEvent OpenFile(strFiles(0))
            Else
                ' many files
                FileName = vbNullString
                If AscW(Right$(strFiles(0), 1)) <> &H5C Then strFiles(0) = strFiles(0) & "\"
                For lngA = 1 To lngA - 1
                    FileName = FileName & strFiles(0) & strFiles(lngA) & "|"
                    RaiseEvent OpenFile(strFiles(0) & strFiles(lngA))
                Next lngA
                FileName = FileName & strFiles(0) & strFiles(lngA)
                RaiseEvent OpenFile(strFiles(0) & strFiles(lngA))
            End If
            ' remember this
            m_Filename = FileName
            ' success!
            ShowOpen = True
        Else
            RaiseEvent OpenCancel(CommDlgExtendedError)
        End If
    End With
End Function
Public Function ShowPrinter() As Boolean

End Function
Public Function ShowSave(Optional ByRef FileName As String) As Boolean
    
End Function
Private Sub UserControl_Initialize()
    m_Buffer = String$(65535, vbNullChar)
End Sub
Private Sub UserControl_InitProperties()
    m_FileFlags = [Hide Read Only] Or [Explorer Style] Or [Long Filenames]
    m_FileFilter = "All files (*.*)" & vbNullChar & "*.*" & vbNullChar
    m_FileOpenTitle = "Open file..."
    m_FileSaveTitle = "Save file..."
    m_FolderFlags = [Return Only File System Dirs] Or [Don't Go Below Domain] Or [New Dialog Style] Or [Usage Hint]
    m_FolderMessage = "Select a file:"
    Private_Init
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim bytBuffer() As Byte, bytEmpty() As Byte
    m_FileFlags = PropBag.ReadProperty("FileFlags", [Hide Read Only] Or [Explorer Style] Or [Long Filenames])
    m_FolderFlags = PropBag.ReadProperty("FolderFlags", [Return Only File System Dirs] Or [Don't Go Below Domain] Or [New Dialog Style] Or [Usage Hint])
    
    bytBuffer = PropBag.ReadProperty("FileCustomFilter", bytEmpty)
    m_FileCustomFilter = bytBuffer
    
    bytBuffer = PropBag.ReadProperty("FileDefaultExtension", bytEmpty)
    m_FileDefaultExtension = bytBuffer
    
    bytBuffer = PropBag.ReadProperty("FileFilter", bytEmpty)
    m_FileFilter = bytBuffer
    
    bytBuffer = PropBag.ReadProperty("FileOpenTitle", bytEmpty)
    m_FileOpenTitle = bytBuffer
    
    bytBuffer = PropBag.ReadProperty("FileSaveTitle", bytEmpty)
    m_FileSaveTitle = bytBuffer
    
    bytBuffer = PropBag.ReadProperty("FolderMessage", bytEmpty)
    m_FolderMessage = bytBuffer
    
    Private_Init
    
    With m_FileDialog
        .lpstrCustomFilter = StrPtr(m_FileCustomFilter)
        .lpstrDefExt = StrPtr(m_FileDefaultExtension)
        .lpstrFilter = StrPtr(m_FileFilter)
    End With
End Sub
Private Sub UserControl_Resize()
    UserControl.SIZE 480, 480
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim bytBuffer() As Byte
    PropBag.WriteProperty "FileFlags", m_FileFlags
    PropBag.WriteProperty "FolderFlags", m_FolderFlags
    
    bytBuffer = m_FileCustomFilter
    PropBag.WriteProperty "FileCustomFilter", bytBuffer
    
    bytBuffer = m_FileDefaultExtension
    PropBag.WriteProperty "FileDefaultExtension", bytBuffer
    
    bytBuffer = m_FileFilter
    PropBag.WriteProperty "FileFilter", bytBuffer
    
    bytBuffer = m_FileOpenTitle
    PropBag.WriteProperty "FileOpenTitle", bytBuffer
    
    bytBuffer = m_FileSaveTitle
    PropBag.WriteProperty "FileSaveTitle", bytBuffer
    
    bytBuffer = m_FolderMessage
    PropBag.WriteProperty "FolderMessage", bytBuffer
End Sub
