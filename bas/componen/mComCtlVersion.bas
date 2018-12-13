Attribute VB_Name = "mComCtlVersion"
'==================================================================================================
'mComCtlVersion.bas      2/21/05
'
'           PURPOSE:
'               Initialize and check the version of comctl32.dll.
'
'           LINEAGE:
'               www.vbaccelerator.com
'
'==================================================================================================
Option Explicit

Private miMajor  As Long
Private miMinor  As Long
Private miBuild  As Long

Public Sub InitCC(Optional ByVal iInit As Long = NegOneL)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Call InitCommonControls or InitCommonControlsEx.
    '---------------------------------------------------------------------------------------
    On Error GoTo VerCheck
    
    If iInit <> NegOneL Then
        Dim tIccex      As INITCOMMONCONTROLSEX
    
        With tIccex
            .dwSize = LenB(tIccex)
            .dwICC = iInit
        End With
        Call vbComCtlTlb.INITCOMMONCONTROLSEX(tIccex)
    Else
        InitCommonControls
    End If
VerCheck:
    pInitVersion
End Sub

Private Sub pInitVersion()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Call the api once to get the version, then store those results.
    '---------------------------------------------------------------------------------------
    Static bInitVersion As Boolean
    If bInitVersion Then Exit Sub
    bInitVersion = True
    
    On Error GoTo unsupported
    
    Dim ltVer      As DLLVERSIONINFO
    ltVer.cbSize = LenB(ltVer)

    If DllGetVersion(ltVer) = ZeroL Then
        With ltVer
            miMajor = .dwMajorVersion
            miMinor = .dwMinorVersion
            miBuild = .dwBuildNumber
        End With
    Else
unsupported:
        miMajor = 4
        miMinor = 0
        miBuild = 0
    End If
    On Error GoTo 0
End Sub

Public Function CheckCCVersion( _
ByVal iMajor As Long, _
Optional ByVal iMinor As Long, _
Optional ByVal iBuild As Long) _
As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Return true if at least the specified version is supported.
    '---------------------------------------------------------------------------------------
    pInitVersion
    If miMajor > iMajor Then
        CheckCCVersion = True
    ElseIf miMajor = iMajor Then
        If miMinor > iMinor Then
            CheckCCVersion = True
        ElseIf miMinor = iMinor Then
            CheckCCVersion = CBool(miBuild >= iBuild)
        End If
    End If
End Function

Public Sub GetCCVersion( _
ByRef iMajor As Long, _
Optional ByRef iMinor As Long, _
Optional ByRef iBuild As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Get the version values.
    '---------------------------------------------------------------------------------------
    pInitVersion
    iMajor = miMajor
    iMinor = miMinor
    iBuild = miBuild
End Sub


