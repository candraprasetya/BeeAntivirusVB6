VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcMemDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'pcMemDC.cls                  8/25/04
'
'           PURPOSE:
'               Memory DC for flicker free gdi drawing.
'
'           LINEAGE:
'               http://www.vbaccelerator.com/home/VB/Code/Controls/Menus/Popup_Menu_ActiveX_DLL/VB6_PopupMenu_DLL_Full_Source.asp
'               pcMemDc.cls
'
'==================================================================================================

Option Explicit

Private mhDc      As Long
Private mhBmpOld  As Long
Private mhBmp     As Long
Private miWidth   As Long
Private miHeight  As Long

Public Sub CreateFromPicture(ByVal sPic As IPicture)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Create a dc and bitmap and copy a picture into it.
    '---------------------------------------------------------------------------------------
    Dim tB            As BITMAP
    Dim lhDcTemp      As Long
    Dim lhDc          As Long
    Dim lhBmpOld      As Long

    vbComCtlTlb.GetObject sPic.handle, Len(tB), tB
    Create tB.bmWidth, tB.bmHeight
    lhDcTemp = CreateDisplayDC()
    If lhDcTemp Then
        lhDc = CreateCompatibleDC(lhDcTemp)
        If lhDc Then
            lhBmpOld = SelectObject(lhDc, sPic.handle)
            If lhBmpOld Then
                BitBlt mhDc, ZeroL, ZeroL, tB.bmWidth, tB.bmHeight, lhDc, 0, 0, vbSrcCopy
                SelectObject lhDc, lhBmpOld
            End If
            
            DeleteDC lhDc
        End If
        DeleteDC lhDcTemp
    End If
End Sub

Public Property Get hdc() As Long
    hdc = mhDc
End Property

Public Property Get Width() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Return the width of the bitmap
    '---------------------------------------------------------------------------------------
    Width = miWidth
End Property

Public Property Get Height() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Return the height of the bitmap
    '---------------------------------------------------------------------------------------
    Height = miHeight
End Property

Public Sub Create(ByVal lW As Long, ByVal lH As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Create a new dc and bitmap with the given dimensions.
    '---------------------------------------------------------------------------------------
    Dim lhDc      As Long
    pDestroy
    
    lhDc = CreateDisplayDC()
    
    If lhDc Then
        mhDc = CreateCompatibleDC(lhDc)
        If mhDc Then
            mhBmp = CreateCompatibleBitmap(lhDc, lW, lH)
            If mhBmp Then
                mhBmpOld = SelectObject(mhDc, mhBmp)
                If mhBmpOld = ZeroL Then
                    pDestroy
                Else
                    miHeight = lH
                    miWidth = lW
                End If
            Else
                pDestroy
            End If
        End If
        DeleteDC lhDc
    End If
End Sub

Private Sub pDestroy()
    '---------------------------------------------------------------------------------------
    ' Date      : 2/21/05
    ' Purpose   : Destroy the dc.
    '---------------------------------------------------------------------------------------
    If mhBmpOld Then
        SelectObject mhDc, mhBmpOld
        mhBmpOld = ZeroL
    End If
    If mhBmp Then
        DeleteObject mhBmp
        mhBmp = ZeroL
    End If
    miWidth = ZeroL
    miHeight = ZeroL
    If mhDc Then
        DeleteDC mhDc
        mhDc = ZeroL
    End If
End Sub

Private Sub Class_Terminate()
    pDestroy
End Sub
