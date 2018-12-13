Attribute VB_Name = "basIconRet"
' Hanya module untuk menggambar ke picture box saja

Option Explicit
'Retrieve Icon
Private Const MAX_PATH As Long = 260
Private Const SHGFI_DISPLAYNAME = &H200, SHGFI_EXETYPE = &H2000, SHGFI_SYSICONINDEX = &H4000, SHGFI_LARGEICON = &H0, SHGFI_SMALLICON = &H1, SHGFI_SHELLICONSIZE = &H4, SHGFI_TYPENAME = &H400, ILD_TRANSPARENT = &H1, BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
Private Type SHFILEINFO
    hIcon As Long: iIcon As Long: dwAttributes As Long: szDisplayName As String * MAX_PATH: szTypeName As String * 80
End Type
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoW" (ByVal pszPath As Long, ByVal dwFileAttributes As Long, ByVal psfi As Long, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDest As Long, ByVal X As Long, ByVal Y As Long, ByVal Flags As Long) As Long
Private shinfo   As SHFILEINFO, sshinfo As SHFILEINFO
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

Public Enum IconRetrieve
    ricnLarge = 32
    ricnSmall = 16
End Enum

'ICON FUNCTION
Public Sub RetrieveIcon(fName As String, DC As PictureBox, icnSize As IconRetrieve)
    Dim hImgSmall, hImgLarge As Long                                                                                                                               'the handle to the system image list
    
    Select Case icnSize
    Case ricnSmall
        hImgSmall = SHGetFileInfo(StrPtr(fName$), 0&, VarPtr(shinfo), Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
        Call ImageList_Draw(hImgSmall, shinfo.iIcon, DC.hdc, 0, 0, ILD_TRANSPARENT)
    Case ricnLarge
        hImgLarge& = SHGetFileInfo(StrPtr(fName$), 0&, VarPtr(shinfo), Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
        Call ImageList_Draw(hImgLarge, shinfo.iIcon, DC.hdc, 0, 0, ILD_TRANSPARENT)
    End Select
End Sub
