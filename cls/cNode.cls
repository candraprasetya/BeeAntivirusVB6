VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'cNode.cls                    12/15/04
'
'           PURPOSE:
'               Represent a node on a treeview control.
'
'==================================================================================================


Option Explicit

Implements iEnumerable

Private moOwner  As ucTreeView
Private mhItem   As Long

Private Sub iEnumerable_Cleanup(tEnum As tEnum)
    '
End Sub

Private Sub iEnumerable_GetNextItem(tEnum As tEnum, vNextItem As Variant, bNoMoreItems As Boolean)
    moOwner.fItem_Enum_GetNextItem mhItem, tEnum, vNextItem, bNoMoreItems
End Sub

Private Sub iEnumerable_Reset(tEnum As tEnum)
    tEnum.iData = ZeroL
End Sub

Private Sub iEnumerable_Skip(tEnum As tEnum, ByVal iSkipCount As Long, bSkippedAll As Boolean)
    moOwner.fItem_Enum_Skip mhItem, tEnum, iSkipCount, bSkippedAll
End Sub


Friend Sub fInit(ByVal oOwner As ucTreeView, ByVal hItem As Long)
    Set moOwner = oOwner
    mhItem = hItem
End Sub

Friend Property Get fIsMine(ByVal oWho As ucTreeView)
    fIsMine = (oWho Is moOwner)
End Property

Friend Property Get fhItem() As Long
    fhItem = mhItem
End Property


Public Function AddChildNodeSorted( _
Optional ByVal sKey As String, _
Optional ByVal sText As String, _
Optional ByVal iIconIndex As Long = NegOneL, _
Optional ByVal iIconIndexSelected As Long = NegOneL, _
Optional ByVal iIconIndexState As Long = NegOneL, _
Optional ByVal iItemData As Long, _
Optional ByVal iItemNumber As Long = NegOneL, _
Optional ByVal bShowPlusMinus As Boolean) _
As cNode
    Set AddChildNodeSorted = moOwner.fItem_AddChildNodeSorted(mhItem, sKey, sText, iIconIndex, iIconIndexSelected, iIconIndexState, iItemData, iItemNumber, bShowPlusMinus)
End Function

Public Function AddChildNodeFirst( _
Optional ByVal sKey As String, _
Optional ByVal sText As String, _
Optional ByVal iIconIndex As Long = NegOneL, _
Optional ByVal iIconIndexSelected As Long = NegOneL, _
Optional ByVal iIconIndexState As Long = NegOneL, _
Optional ByVal iItemData As Long, _
Optional ByVal iItemNumber As Long = NegOneL, _
Optional ByVal bShowPlusMinus As Boolean) _
As cNode
    Set AddChildNodeFirst = moOwner.fItem_AddChildNodeFirst(mhItem, sKey, sText, iIconIndex, iIconIndexSelected, iIconIndexState, iItemData, iItemNumber, bShowPlusMinus)
End Function

Public Function AddChildNode( _
Optional ByVal vNodeAfter As Variant, _
Optional ByVal sKey As String, _
Optional ByVal sText As String, _
Optional ByVal iIconIndex As Long = NegOneL, _
Optional ByVal iIconIndexSelected As Long = NegOneL, _
Optional ByVal iIconIndexState As Long = NegOneL, _
Optional ByVal iItemData As Long, _
Optional ByVal iItemNumber As Long = NegOneL, _
Optional ByVal bShowPlusMinus As Boolean) _
As cNode
    Set AddChildNode = moOwner.fItem_AddChildNode(mhItem, vNodeAfter, sKey, sText, iIconIndex, iIconIndexSelected, iIconIndexState, iItemData, iItemNumber, bShowPlusMinus)
End Function

Public Property Get GetNode(ByVal iNode As eTreeViewGetNode) As cNode
    Set GetNode = moOwner.fItem_GetNode(mhItem, iNode)
End Property

Public Sub Delete()
    moOwner.fItem_Delete mhItem
End Sub

Public Sub DeleteChildren()
    moOwner.fItem_DeleteChildren mhItem
End Sub

Public Property Get ChildCount() As Long
    ChildCount = moOwner.fItem_ChildCount(mhItem)
End Property

Public Property Get hItem() As Long
    hItem = moOwner.fItem_hItem(mhItem)
End Property

Public Property Get Bold() As Boolean
    Bold = moOwner.fItem_Bold(mhItem)
End Property
Public Property Let Bold(ByVal bNew As Boolean)
    moOwner.fItem_Bold(mhItem) = bNew
End Property

Public Property Get IconIndexState() As Long
    IconIndexState = moOwner.fItem_IconIndexState(mhItem)
End Property
Public Property Let IconIndexState(ByVal iNew As Long)
    moOwner.fItem_IconIndexState(mhItem) = iNew
End Property

Public Property Get IconIndex() As Long
    IconIndex = moOwner.fItem_IconIndex(mhItem)
End Property
Public Property Let IconIndex(ByVal iNew As Long)
    moOwner.fItem_IconIndex(mhItem) = iNew
End Property

Public Property Get IconIndexSelected() As Long
    IconIndexSelected = moOwner.fItem_IconIndexSelected(mhItem)
End Property
Public Property Let IconIndexSelected(ByVal iNew As Long)
    moOwner.fItem_IconIndexSelected(mhItem) = iNew
End Property

Public Property Get DropHighlighted() As Boolean
    DropHighlighted = moOwner.fItem_DropHighlighted(mhItem)
End Property
Public Property Let DropHighlighted(ByVal bNew As Boolean)
    moOwner.fItem_DropHighlighted(mhItem) = bNew
End Property

Public Sub EnsureVisible()
    moOwner.fItem_EnsureVisible mhItem
End Sub

Public Property Get Expanded() As Boolean
    Expanded = moOwner.fItem_Expanded(mhItem)
End Property
Public Property Let Expanded(ByVal bNew As Boolean)
    moOwner.fItem_Expanded(mhItem) = bNew
End Property

Public Property Get FullPath() As String
    FullPath = moOwner.fItem_Expanded(mhItem)
End Property

Public Property Get ItemData() As Long
    ItemData = moOwner.fItem_ItemData(mhItem)
End Property
Public Property Let ItemData(ByVal iNew As Long)
    moOwner.fItem_ItemData(mhItem) = iNew
End Property

Public Property Get ItemNumber() As Long
    ItemNumber = moOwner.fItem_ItemNumber(mhItem)
End Property
Public Property Let ItemNumber(ByVal iNew As Long)
    moOwner.fItem_ItemNumber(mhItem) = iNew
End Property

Public Property Get Key() As String
    Key = moOwner.fItem_Key(mhItem)
End Property
Public Property Let Key(ByRef sNew As String)
    moOwner.fItem_Key(mhItem) = sNew
End Property

Public Property Get Selected() As Boolean
    Selected = moOwner.fItem_Selected(mhItem)
End Property
Public Property Let Selected(ByVal bNew As Boolean)
    moOwner.fItem_Selected(mhItem) = bNew
End Property

Public Property Get ShowPlusMinus() As Boolean
    ShowPlusMinus = moOwner.fItem_ShowPlusMinus(mhItem)
End Property
Public Property Let ShowPlusMinus(ByRef bNew As Boolean)
    moOwner.fItem_ShowPlusMinus(mhItem) = bNew
End Property

Public Property Get Text() As String
    Text = moOwner.fItem_Text(mhItem)
End Property
Public Property Let Text(ByRef sNew As String)
    moOwner.fItem_Text(mhItem) = sNew
End Property

Public Property Get Cut() As Boolean
    Cut = moOwner.fItem_Cut(mhItem)
End Property
Public Property Let Cut(ByVal bNew As Boolean)
    moOwner.fItem_Cut(mhItem) = bNew
End Property

Public Sub Sort()
    moOwner.fItem_Sort mhItem
End Sub

Public Property Get Left() As Single
    Left = moOwner.fItem_Left(mhItem)
End Property
Public Property Get Top() As Single
    Top = moOwner.fItem_Top(mhItem)
End Property
Public Property Get Width() As Single
    Width = moOwner.fItem_Width(mhItem)
End Property
Public Property Get Height() As Single
    Height = moOwner.fItem_Height(mhItem)
End Property

Public Sub Move(ByVal vNodeRelative As Variant, ByVal iRelation As eTreeViewNodeRelation)
    moOwner.fItem_Move mhItem, vNodeRelative, iRelation
End Sub

Public Property Get HasChildren() As Boolean
    HasChildren = moOwner.fItem_HasChildren(mhItem)
End Property

Public Property Get NewEnum() As IUnknown
    'verify that we're still an existing node
    Dim i      As Long
    i = moOwner.fItem_hItem(mhItem)
    
    Dim loEnum      As pcEnumeration
    Set loEnum = New pcEnumeration
    Set NewEnum = loEnum.GetEnum(Me, ZeroL)
End Property
