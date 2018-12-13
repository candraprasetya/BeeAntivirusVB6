Attribute VB_Name = "mLoadShellLib"
'==================================================================================================
'mLoadShellLib.bas      9/2/05
'
'           PURPOSE:
'               Load one handle of shell32.dll.  Doing this when the first usercontrol
'               is initialized and releasing it when the last usercontrol terminates
'               prevents a crash at shutdown when linked to cc version 6.
'
'==================================================================================================

Option Explicit

Private mhMod       As Long
Private miModCount  As Long

Public Sub LoadShellMod()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/13/05
    ' Purpose   : Load the handle to shell32.dll
    '---------------------------------------------------------------------------------------
    If CheckCCVersion(6) Then
        If (mhMod Or miModCount) = ZeroL Then
            Dim lsAnsi      As String
            lsAnsi = StrConv("Shell32.dll" & vbNullChar, vbFromUnicode)
            mhMod = LoadLibrary(ByVal StrPtr(lsAnsi))
        End If
        miModCount = miModCount + OneL
    End If
End Sub

Public Sub ReleaseShellMod()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/13/05
    ' Purpose   : Release our handle to shell32.dll
    '---------------------------------------------------------------------------------------
    If CheckCCVersion(6) Then
        miModCount = miModCount - OneL
        If miModCount = ZeroL And mhMod <> ZeroL Then
            FreeLibrary mhMod
            mhMod = ZeroL
        End If
    End If
End Sub
