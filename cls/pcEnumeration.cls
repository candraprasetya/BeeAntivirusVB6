VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcEnumeration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'pcEnumeration.cls                       7/4/04
'
'           PURPOSE:
'               Holds a placemarker and other client-specific data for an enumeration.
'
'==================================================================================================
Option Explicit

Implements Interfaces.IEnumVARIANT

Private mtEnum   As tEnum
Private moOwner  As iEnumerable

Private Sub IEnumVARIANT_Clone(ppEnum As Interfaces.IEnumVARIANT)
    'dummy - see mEnumVariant.bas
End Sub

Private Sub IEnumVARIANT_Next(ByVal Celt As Long, rgVar As Variant, ByVal pCeltFetched As Long)
    'dummy - see mEnumVariant.bas
End Sub

Private Sub IEnumVARIANT_Reset()
    'dummy - see mEnumVariant.bas
End Sub

Private Sub IEnumVARIANT_Skip(ByVal Celt As Long)
    'dummy - see mEnumVariant.bas
End Sub


Public Function GetEnum( _
ByVal oWho As iEnumerable, _
ByVal iControl As Long, _
Optional ByVal iData As Long, _
Optional ByVal iIndex As Long = -1&) _
As IUnknown
                
    With mtEnum
        .iIndex = iIndex
        .iControl = iControl
        .iData = iData
    End With

    Set GetEnum = Me
    Set moOwner = oWho
End Function

Private Sub Class_Initialize()
    ReplaceIEnumVARIANT Me
End Sub

Private Sub Class_Terminate()
    If Not moOwner Is Nothing Then moOwner.Cleanup mtEnum
    Set moOwner = Nothing
    RestoreIEnumVARIANT Me
End Sub

Friend Sub GetNextItem( _
ByRef vNextItem As Variant, _
ByRef bNoMoreItems As Boolean)
    'debug.assert Not moOwner Is Nothing
    If Not moOwner Is Nothing Then moOwner.GetNextItem mtEnum, vNextItem, bNoMoreItems Else bNoMoreItems = True
End Sub

Friend Sub Reset()
    'debug.assert Not moOwner Is Nothing
    If Not moOwner Is Nothing Then moOwner.Reset mtEnum
End Sub

Friend Sub skip( _
ByVal iSkipCount As Long, _
ByRef bSkippedAll As Boolean)
    'debug.assert Not moOwner Is Nothing
    If Not moOwner Is Nothing Then moOwner.skip mtEnum, iSkipCount, bSkippedAll
End Sub
