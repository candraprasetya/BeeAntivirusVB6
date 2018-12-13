VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcMnemonics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'cMnemonics.cls                      1/26/05
'
'           PURPOSE:
'               Manage a win32 ACCEL structure.
'
'           LINEAGE:
'               http://www.vbaccelerator.com/home/VB/Code/Controls/Explorer_Bar_Control/VB6_Explorer_Bar_Control_Full_Source.zip
'               pcMnemonics.cls
'
'==================================================================================================

Option Explicit

Event Changed()

Private mhAccel       As Long
Private mtAccel()     As ACCEL
Private miRefCount()  As Long
Private miCount       As Long

Public Property Get Count() As Long
    Count = miCount
End Property

Public Property Get hAccel() As Long
    hAccel = mhAccel
End Property

Public Sub Add(ByRef sCaption As String)
    Dim liKey      As Long
    liKey = AccelChar(sCaption)
    If liKey Then liKey = GetVirtKey(liKey)
    
    If liKey Then
        Dim i      As Long
        For i = ZeroL To miCount - OneL
            If mtAccel(i).Key = liKey Then
                miRefCount(i) = miRefCount(i) + OneL
                Exit Sub
            End If
        Next
        
        If miCount Then
            ReDim Preserve mtAccel(ZeroL To miCount)
            ReDim Preserve miRefCount(ZeroL To miCount)
        Else
            ReDim mtAccel(ZeroL To miCount)
            ReDim miRefCount(ZeroL To miCount)
        End If
        
        With mtAccel(miCount)
            .fVirt = FALT Or FVIRTKEY
            mtAccel(miCount).cmd = NextCommandId
            mtAccel(miCount).Key = liKey
        End With
        
        miRefCount(miCount) = ZeroL
        
        miCount = miCount + OneL
        
        If mhAccel Then DestroyAcceleratorTable mhAccel
        mhAccel = CreateAcceleratorTable(mtAccel(ZeroL), miCount)
        
        RaiseEvent Changed
        
    End If
End Sub

Public Sub Remove(ByRef sCaption As String)
    Dim liKey      As Long
    liKey = AccelChar(sCaption)
    
    If liKey Then
        Dim i      As Long
        
        For i = ZeroL To miCount - OneL
            If mtAccel(i).Key = liKey Then
                If miRefCount(i) Then
                    miRefCount(i) = miRefCount(i) - OneL
                    Exit Sub
                Else
                    Exit For
                End If
            End If
        Next
        
        ''debug.assert i < miCount
        
        If i < miCount Then
            For i = i To miCount - TwoL
                LSet mtAccel(i) = mtAccel(i + OneL)
                miRefCount(i) = miRefCount(i + OneL)
            Next
            
            miCount = miCount - OneL
            If miCount Then
                ReDim Preserve mtAccel(ZeroL To miCount - OneL)
                ReDim Preserve miRefCount(ZeroL To miCount - OneL)
            End If
            If mhAccel Then DestroyAcceleratorTable mhAccel
            mhAccel = CreateAcceleratorTable(mtAccel(ZeroL), miCount)
            
            RaiseEvent Changed
            
        End If
        
    End If
End Sub

Private Sub Class_Terminate()
    If mhAccel Then DestroyAcceleratorTable mhAccel
End Sub

