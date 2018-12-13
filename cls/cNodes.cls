VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'cNodes.cls                    12/15/04
'
'           PURPOSE:
'               Represent a collection of all the nodes on a treeview control.
'
'==================================================================================================

Option Explicit

Implements iEnumerable

Private moOwner  As ucTreeView

Friend Sub fInit(ByVal oOwner As ucTreeView)
    Set moOwner = oOwner
End Sub

Private Sub iEnumerable_Cleanup(tEnum As tEnum)
    '
End Sub

Private Sub iEnumerable_GetNextItem(tEnum As tEnum, vNextItem As Variant, bNoMoreItems As Boolean)
    moOwner.fItems_Enum_GetNextItem tEnum, vNextItem, bNoMoreItems
End Sub

Private Sub iEnumerable_Reset(tEnum As tEnum)
    tEnum.iControl = ZeroL
    tEnum.iData = ZeroL
End Sub

Private Sub iEnumerable_Skip(tEnum As tEnum, ByVal iSkipCount As Long, bSkippedAll As Boolean)
    moOwner.fItems_Enum_Skip tEnum, iSkipCount, bSkippedAll
End Sub




Public Function Add( _
Optional ByVal vNodeRelative As Variant, _
Optional ByVal iRelation As eTreeViewNodeRelation, _
Optional ByRef sKey As String, _
Optional ByRef sText As String, _
Optional ByVal iIconIndex As Long, _
Optional ByVal iIconIndexSelected As Long = NegOneL, _
Optional ByVal iIconIndexState As Long, _
Optional ByVal iItemData As Long, _
Optional ByVal iItemNumber As Long = NegOneL, _
Optional ByVal bShowPlusMinus As Boolean) _
As cNode
    If iIconIndexSelected = NegOneL Then iIconIndexSelected = iIconIndex
    Set Add = moOwner.fItems_Add(vNodeRelative, iRelation, sKey, sText, iIconIndex, iIconIndexSelected, iIconIndexState, iItemData, iItemNumber, bShowPlusMinus)
End Function

Public Sub Remove(ByVal vNode As Variant)
    moOwner.fItems_Remove vNode
End Sub

Public Property Get Item(ByVal vNode As Variant) As cNode
    Set Item = moOwner.fItems_Item(vNode)
End Property

Public Sub Clear()
    moOwner.fItems_Clear
End Sub

Public Property Get Exists(ByVal vNode As Variant) As Boolean
    Exists = moOwner.fItems_Exists(vNode)
End Property

Public Property Get Count() As Long
    Count = moOwner.fItems_Count
End Property

Public Property Get NewEnum() As IUnknown
    Set NewEnum = moOwner.fItems_NewEnum(Me)
End Property
