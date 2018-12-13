VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cListSubItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'cListSubItem.cls                    12/15/04
'
'           PURPOSE:
'               Represent a subitem in a listview control.
'
'==================================================================================================

Option Explicit

Private moOwner        As ucListView
Private milpItem       As Long
Private miIndex        As Long
Private miColumnIndex  As Long
Private miColumnId     As Long

Friend Sub fInit(ByVal oOwner As ucListView, ByVal lpItem As Long, ByVal iIndex As Long, ByVal iColumnIndex As Long, ByVal iColumnId As Long)
    Set moOwner = oOwner
    milpItem = lpItem
    miIndex = iIndex
    miColumnIndex = iColumnIndex
    miColumnId = iColumnId
End Sub

Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0
    Text = moOwner.fSubItem_Text(milpItem, miIndex, miColumnIndex, miColumnId)
End Property

Public Property Let Text(ByRef sNew As String)
    moOwner.fSubItem_Text(milpItem, miIndex, miColumnIndex, miColumnId) = sNew
End Property

Public Property Get IconIndex() As Long
    IconIndex = moOwner.fSubItem_IconIndex(milpItem, miIndex, miColumnIndex, miColumnId)
End Property
Public Property Let IconIndex(ByVal iNew As Long)
    moOwner.fSubItem_IconIndex(milpItem, miIndex, miColumnIndex, miColumnId) = iNew
End Property

Public Property Get ShowInTileView() As Boolean
    ShowInTileView = moOwner.fSubItem_ShowInTileView(milpItem, miIndex, miColumnIndex, miColumnId)
End Property
Public Property Let ShowInTileView(ByVal bNew As Boolean)
    moOwner.fSubItem_ShowInTileView(milpItem, miIndex, miColumnIndex, miColumnId) = bNew
End Property

Public Sub SetFormattedText(ByVal vData As Variant)
    moOwner.fSubItem_SetFormattedText milpItem, miIndex, miColumnIndex, miColumnId, vData
End Sub
