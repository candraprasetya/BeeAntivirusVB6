Attribute VB_Name = "basIcon"
' Module Untuk Mendapatkan Ceksum icon dan Cek Icon

Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExW" (ByVal lpszFile As Long, ByVal nIconIndex As Long, ByRef phiconLarge As Long, ByRef phiconSmall As Long, ByVal nIcons As Long) As Long
Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean

Dim InterIco(50) As String
Dim NumZ         As Byte
Dim CFLFL        As New classFile

Public Sub LoadDataIcon() ' di init saat load
InterIco(0) = "15A550412FF69A":    InterIco(1) = "12CD58F10C578B":    InterIco(2) = "12C64AA11F31D5"
InterIco(3) = "15A55041309047":    InterIco(4) = "179B281181FE83":    InterIco(5) = "179B281181FE83"
InterIco(6) = "18CB20E10D6585":    InterIco(7) = "1166113170EEA5":    InterIco(8) = "1888F08178DF96"
End Sub

Private Function CEK_ICON(Ceksum_Icon As String, Path As String) As Boolean
For NumZ = 0 To 8
    If InterIco(NumZ) = Ceksum_Icon Then
       CEK_ICON = True
       Exit Function
    End If
Next
CEK_ICON = False
End Function
' ----------------------------------------------      CEK ICON      -------------------------------------

Public Function DRAW_ICO(PathToDraw As String, PicBox As PictureBox) As Boolean ' Yang dipanggil untuk Cek Icon
Dim hIcon       As Long
Dim IconExist   As Long

Dim HashIco As String
Dim SaveTmp As String
Dim Ukuran  As String

On Error GoTo KELUAR
DoEvents
DRAW_ICO = False ' init nilainya ke False
PicBox.Cls

SaveTmp = "C:\ico.tmp"
IconExist = ExtractIconEx(StrPtr(PathToDraw), 0, ByVal 0&, hIcon, 1)

If IconExist <= 0 Then
    IconExist = ExtractIconEx(StrPtr(PathToDraw), 0, hIcon, ByVal 0&, 1)
    If IconExist <= 0 Then Exit Function
End If

DrawIconEx PicBox.hdc, 0, 0, hIcon, 0, 0, 0, 0, &H3

SavePicture PicBox.Image, SaveTmp  ' Simpan Dulu Gambarnya
HashIco = CALC_BYTE_ICON(SaveTmp) ' Calculasikan Byte Simpanan

If CEK_ICON(HashIco, PathToDraw) = True Then
    DRAW_ICO = True
Else
    DRAW_ICO = False
End If

HapusFile SaveTmp ' Hapus Simpanan Icon
KELUAR:
End Function

' sementara
Private Function CALC_BYTE_ICON(Path As String) As String ' Kalkulasikan Byte Icon
On Error Resume Next
    Dim hFileIcon  As Long
    Dim iTurn      As Long
    
    hFileIcon = CFLFL.VbOpenFile(Path, FOR_BINARY_ACCESS_READ, LOCK_NONE)
    
    If hFileIcon > 0 Then
       CALC_BYTE_ICON = MYCeksumCadangan(Path, hFileIcon)
    Else
       CALC_BYTE_ICON = "00"
    End If
    
    frmCeksumIcon.Text2.Text = CALC_BYTE_ICON
    CFLFL.VbCloseFile hFileIcon
End Function

Public Function CheckIcon(Where As String, hFile As Long) As Boolean
On Error GoTo KELUAR
If IsPE32EXE = False Then GoTo KELUAR

If DRAW_ICO(Where, frMain.picBuff) = True Then
    CheckIcon = True
    TutupFile hFile
    Exit Function
Else
    CheckIcon = False
End If
KELUAR:
End Function
