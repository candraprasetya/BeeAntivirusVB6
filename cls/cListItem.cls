VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'cListItem.cls                    12/15/04
'
'           PURPOSE:
'               Represent a list item in a listview control.
'
'==================================================================================================

Option Explicit

Private moOwner   As ucListView
Private milpItem  As Long
Private miIndex   As Long

Friend Sub fInit(ByVal oOwner As ucListView, ByVal lpItem As Long, ByVal iIndex As Long)
    Set moOwner = oOwner
    milpItem = lpItem
    miIndex = iIndex
End Sub

Friend Function fIsOwner(ByVal oOwner As ucListView) As Boolean
    fIsOwner = (oOwner Is moOwner)
End Function

Public Property Get ToolTipText() As String
    ToolTipText = moOwner.fItem_ToolTipText(milpItem, miIndex)
End Property
Public Property Let ToolTipText(ByRef sNew As String)
    moOwner.fItem_ToolTipText(milpItem, miIndex) = sNew
End Property

Public Property Get ItemData() As Long
    ItemData = moOwner.fItem_ItemData(milpItem, miIndex)
End Property
Public Property Let ItemData(ByRef iNew As Long)
    moOwner.fItem_ItemData(milpItem, miIndex) = iNew
End Property

Public Property Get Text() As String
    Text = moOwner.fItem_Text(milpItem, miIndex)
End Property
Public Property Let Text(ByVal sNew As String)
    moOwner.fItem_Text(milpItem, miIndex) = sNew
End Property

Public Property Get Key() As String
    Key = moOwner.fItem_Key(milpItem, miIndex)
End Property

Public Property Let Key(ByRef sNew As String)
    moOwner.fItem_Key(milpItem, miIndex) = sNew
End Property

Public Property Get Index() As Long
    Index = moOwner.fItem_Index(milpItem, miIndex)
End Property

Public Property Get IconIndex() As Long
    IconIndex = moOwner.fItem_IconIndex(milpItem, miIndex)
End Property
Public Property Let IconIndex(ByVal iNew As Long)
    moOwner.fItem_IconIndex(milpItem, miIndex) = iNew
End Property

Public Property Get Checked() As Boolean
    Checked = moOwner.fItem_Checked(milpItem, miIndex)
End Property
Public Property Let Checked(ByVal bNew As Boolean)
    moOwner.fItem_Checked(milpItem, miIndex) = bNew
End Property

Public Property Get Cut() As Boolean
    Cut = moOwner.fItem_Cut(milpItem, miIndex)
End Property
Public Property Let Cut(ByVal bNew As Boolean)
    moOwner.fItem_Cut(milpItem, miIndex) = bNew
End Property

Public Property Get Indent() As Long
    Indent = moOwner.fItem_Indent(milpItem, miIndex)
End Property
Public Property Let Indent(ByVal iNew As Long)
    moOwner.fItem_Indent(milpItem, miIndex) = iNew
End Property

Public Property Get Selected() As Boolean
    Selected = moOwner.fItem_Selected(milpItem, miIndex)
End Property
Public Property Let Selected(ByVal bNew As Boolean)
    moOwner.fItem_Selected(milpItem, miIndex) = bNew
End Property

Public Property Get Top() As Single
    Top = moOwner.fItem_Top(milpItem, miIndex)
End Property
Public Property Let Top(ByVal fNew As Single)
    moOwner.fItem_Top(milpItem, miIndex) = fNew
End Property

Public Property Get Left() As Single
    Left = moOwner.fItem_Left(milpItem, miIndex)
End Property
Public Property Let Left(ByVal fNew As Single)
    moOwner.fItem_Left(milpItem, miIndex) = fNew
End Property

Public Property Get SubItem(ByVal vColumn As Variant) As cListSubItem
    Set SubItem = moOwner.fItem_SubItem(milpItem, miIndex, vColumn)
End Property

Public Property Get Group() As cItemGroup
    Set Group = moOwner.fItem_Group(milpItem, miIndex)
End Property

Public Property Let Group(ByVal cG As cItemGroup)
    Set moOwner.fItem_Group(milpItem, miIndex) = cG
End Property
Public Property Set Group(ByVal cG As cItemGroup)
    Set moOwner.fItem_Group(milpItem, miIndex) = cG
End Property

Public Sub SetGroup(ByVal vKeyOrIndex As Variant)
    moOwner.fItem_SetGroup milpItem, miIndex, vKeyOrIndex
End Sub

Public Sub SetTileViewItems(ByRef iItems() As Long)
    moOwner.fItem_SetTileViewItems milpItem, miIndex, iItems
End Sub

