VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cListItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'cListItems.cls                    12/15/04
'
'           PURPOSE:
'               Represent a collection of list items in a listview control.
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
Optional ByRef sToolTipText As String, _
Optional ByVal iIconIndex As Long = NegOneL, _
Optional ByVal iItemData As Long, _
Optional ByVal iIndent As Long, _
Optional ByVal vItemBefore As Variant, _
Optional ByVal vItemGroup As Variant, _
Optional ByVal vSubItems As Variant) _
As cListItem
    
    Set Add = moOwner.fItems_Add(sKey, sText, sToolTipText, iIconIndex, iItemData, iIndent, vItemBefore, vItemGroup, vSubItems)
    
End Function

Public Sub Remove(ByVal vItem As Variant)
    moOwner.fItems_Remove vItem
End Sub

Public Property Get Count() As Long
    Count = moOwner.fItems_Count()
End Property

Public Property Get Exists(ByVal vItem As Variant) As Boolean
    Exists = moOwner.fItems_Exists(vItem)
End Property

Public Property Get Item(ByVal vItem As Variant) As cListItem
    Set Item = moOwner.fItems_Item(vItem)
End Property


Public Sub Clear()
    moOwner.fItems_Clear
End Sub

Public Sub InitStorage(ByVal iAdditionalItems As Long)
    moOwner.fItems_InitStorage iAdditionalItems
End Sub


Private Sub iEnumerable_Cleanup(tEnum As tEnum)
    '
End Sub

Private Sub iEnumerable_GetNextItem(tEnum As tEnum, vNextItem As Variant, bNoMoreItems As Boolean)
    moOwner.fItems_NextItem tEnum, vNextItem, bNoMoreItems
End Sub

Private Sub iEnumerable_Reset(tEnum As tEnum)
    tEnum.iIndex = NegOneL
    tEnum.iControl = moOwner.fItems_Control()
End Sub

Private Sub iEnumerable_Skip(tEnum As tEnum, ByVal iSkipCount As Long, bSkippedAll As Boolean)
    tEnum.iIndex = tEnum.iIndex + iSkipCount
End Sub

Public Property Get NewEnum() As stdole.IUnknown
    Dim loEnum      As pcEnumeration
    Set loEnum = New pcEnumeration
    
    Set NewEnum = loEnum.GetEnum(Me, moOwner.fItems_Control())
End Property
