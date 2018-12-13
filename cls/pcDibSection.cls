VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcDibSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'pcDibSection.bas      4/21/05
'
'           PURPOSE:
'               Create and manage a dib section.
'
'           LINEAGE:
'               www.vbaccelerator.com
'
'==================================================================================================

Option Explicit

Private mtBMI  As BITMAPINFO
Private mtSA   As SAFEARRAY2D

Private mhDib As Long
Private mpBits As Long

Public Function Create(ByVal iWidth As Long, ByVal iHeight As Long, Optional ByVal hdc As Long, Optional ByVal iBitCount As Long = 24) As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Create the dib with the given dimensions.
    '---------------------------------------------------------------------------------------
    
    Destroy
    
    Dim lhDc      As Long
    If hdc Then lhDc = hdc Else lhDc = GetDC(ZeroL)
    
    Dim ltBMI As BITMAPINFO
    With ltBMI.bmiHeader
        .biSize = Len(ltBMI.bmiHeader)
        .biWidth = iWidth
        .biHeight = iHeight
        .biPlanes = OneL
        .biBitCount = iBitCount
        .biCompression = ZeroL
        .biSizeImage = ((.biWidth * .biBitCount \ 8&) And &HFFFFFFFC) * .biHeight
    End With
    
    mhDib = CreateDIBSection(lhDc, ltBMI, ZeroL, mpBits, ZeroL, ZeroL)
    
    If mhDib Then
        LSet mtBMI = ltBMI
        Create = True
    End If
    
    'debug.assert Create
    
    If hdc <> lhDc Then ReleaseDC ZeroL, lhDc
    
End Function

Public Function Destroy() As Boolean
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Destroy the dib.
    '---------------------------------------------------------------------------------------
    If mhDib Then
        If DeleteObject(mhDib) Then
            Destroy = True
            mhDib = ZeroL
            mpBits = ZeroL
        End If
        Dim ltBMI      As BITMAPINFO: LSet mtBMI = ltBMI
    End If
End Function

Public Property Get pBits() As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Return the pointer to the dib bits.
    '---------------------------------------------------------------------------------------
    pBits = mpBits
End Property

Public Property Get hBitmap() As Long
    hBitmap = mhDib
End Property

Public Property Get Height() As Long
    Height = mtBMI.bmiHeader.biHeight
End Property

Public Property Get Width() As Long
    Width = mtBMI.bmiHeader.biWidth
End Property

Public Property Get BitCount() As Long
    BitCount = mtBMI.bmiHeader.biBitCount
End Property

Public Property Get ArrPtr(ByVal cbElements As Long) As Long
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Return a safearray pointer that can be used to access the bitmap.
    '---------------------------------------------------------------------------------------
    Const FADF_FIXEDSIZE As Long = &H10
    Const FADF_STATIC As Long = &H2
    Const FADF_AUTO As Long = &H1
    
    Const FADF_Flags = FADF_FIXEDSIZE Or FADF_STATIC Or FADF_AUTO
    
    With mtSA
        .fFeatures = FADF_Flags
        .cbElements = cbElements
        .cDims = TwoL
        .Bounds(0).lLbound = ZeroL
        .Bounds(0).cElements = mtBMI.bmiHeader.biHeight
        .Bounds(1).lLbound = ZeroL
        .Bounds(1).cElements = ((mtBMI.bmiHeader.biWidth * mtBMI.bmiHeader.biBitCount \ 8&) \ cbElements) And &HFFFFFFFC
        .pvData = mpBits
    End With
    ArrPtr = VarPtr(mtSA)
End Property

Private Sub Class_Terminate()
    Destroy
End Sub
