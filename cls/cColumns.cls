VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'cColumns.cls                    12/15/04
'
'           PURPOSE:
'               Represent a collection of listview columns.
'
'==================================================================================================

Option Explicit

Implements iEnumerable

Private moOwner  As ucListView

Friend Sub fInit(ByVal oOwner As ucListView)
    Set moOwner = oOwner
End Sub

Public Function Add( _
Optional ByRef sKey As String, _
Optional ByRef sText As String, _
Optional ByVal iIcon As Long = NegOneL, _
Optional ByVal iSortType As eListViewSortType, _
Optional ByVal iAlignment As eListViewColumnAlign, _
Optional ByVal fWidth As Single = NegOneF, _
Optional ByRef sFormat As String, _
Optional ByVal vColumnBefore As Variant) _
As cColumn
    Set Add = moOwner.fColumns_Add(sKey, sText, iIcon, iSortType, iAlignment, fWidth, sFormat, vColumnBefore)
End Function

Public Sub Clear()
    moOwner.fColumns_Clear
End Sub

Public Property Get Count() As Long
    Count = moOwner.fColumns_Count
End Property

Public Property Get Item(ByVal vColumn As Variant) As cColumn
    Set Item = moOwner.fColumns_Item(vColumn)
End Property

Public Sub Remove(ByVal vColumn As Variant)
    moOwner.fColumns_Remove vColumn
End Sub

Public Property Get Exists(ByVal vColumn As Variant) As Boolean
    Exists = moOwner.fColumns_Exists(vColumn)
End Property


Private Sub iEnumerable_Cleanup(tEnum As tEnum)
    '
End Sub

Private Sub iEnumerable_GetNextItem(tEnum As tEnum, vNextItem As Variant, bNoMoreItems As Boolean)
    moOwner.fColumns_NextItem tEnum, vNextItem, bNoMoreItems
End Sub

Private Sub iEnumerable_Reset(tEnum As tEnum)
    tEnum.iIndex = NegOneL
    tEnum.iControl = moOwner.fColumns_Control()
End Sub

Private Sub iEnumerable_Skip(tEnum As tEnum, ByVal iSkipCount As Long, bSkippedAll As Boolean)
    tEnum.iIndex = tEnum.iIndex + iSkipCount
    bSkippedAll = tEnum.iIndex < moOwner.fColumns_Count()
End Sub


Public Property Get NewEnum() As stdole.IUnknown
    Dim loEnum      As pcEnumeration
    Set loEnum = New pcEnumeration
    
    Set NewEnum = loEnum.GetEnum(Me, moOwner.fColumns_Control())
End Property
