VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'cColumn.cls                    12/15/04
'
'           PURPOSE:
'               Represent a listview column.
'
'==================================================================================================

Option Explicit

Private moOwner  As ucListView
Private miIndex  As Long
Private miId     As Long

Friend Sub fInit(ByVal oOwner As ucListView, ByVal iIndex As Long, ByVal iId As Long)
    Set moOwner = oOwner
    miIndex = iIndex
    miId = iId
End Sub

Friend Function fIsOwner(ByVal oOwner As ucListView) As Boolean
    fIsOwner = (oOwner Is moOwner)
End Function

Public Property Get Width() As Single
    Width = moOwner.fColumn_Width(miIndex, miId)
End Property
Public Property Let Width(ByVal fNew As Single)
    moOwner.fColumn_Width(miIndex, miId) = fNew
End Property

Public Property Get Text() As String
Attribute Text.VB_MemberFlags = "200"
    Text = moOwner.fColumn_Text(miIndex, miId)
End Property
Public Property Let Text(ByVal sNew As String)
    moOwner.fColumn_Text(miIndex, miId) = sNew
End Property

Public Property Get Key() As String
    Key = moOwner.fColumn_Key(miIndex, miId)
End Property
Public Property Let Key(ByVal sNew As String)
    moOwner.fColumn_Key(miIndex, miId) = sNew
End Property

Public Property Get IconIndex() As Long
    IconIndex = moOwner.fColumn_IconIndex(miIndex, miId)
End Property
Public Property Let IconIndex(ByVal iNew As Long)
    moOwner.fColumn_IconIndex(miIndex, miId) = iNew
End Property

Public Property Get ImageOnRight() As Boolean
    ImageOnRight = moOwner.fColumn_ImageOnRight(miIndex, miId)
End Property
Public Property Let ImageOnRight(ByVal bNew As Boolean)
    moOwner.fColumn_ImageOnRight(miIndex, miId) = bNew
End Property

Public Property Get Alignment() As eListViewColumnAlign
    Alignment = moOwner.fColumn_Alignment(miIndex, miId)
End Property
Public Property Let Alignment(ByVal iNew As eListViewColumnAlign)
    moOwner.fColumn_Alignment(miIndex, miId) = iNew
End Property

Public Property Get SortOrder() As eListViewSortOrder
    SortOrder = moOwner.fColumn_SortOrder(miIndex, miId)
End Property
Public Property Let SortOrder(ByVal iNew As eListViewSortOrder)
    moOwner.fColumn_SortOrder(miIndex, miId) = iNew
End Property

Public Property Get SortType() As eListViewSortType
    SortType = moOwner.fColumn_SortType(miIndex, miId)
End Property
Public Property Let SortType(ByVal iNew As eListViewSortType)
    moOwner.fColumn_SortType(miIndex, miId) = iNew
End Property

Public Property Get Position() As Long
    Position = moOwner.fColumn_Position(miIndex, miId)
End Property
Public Property Let Position(ByVal iNew As Long)
    moOwner.fColumn_Position(miIndex, miId) = iNew
End Property

Public Property Get Format() As String
    Format = moOwner.fColumn_Format(miIndex, miId)
End Property
Public Property Let Format(ByRef sNew As String)
    moOwner.fColumn_Format(miIndex, miId) = sNew
End Property

Public Property Get Index() As Long
    Index = moOwner.fColumn_Index(miIndex, miId)
End Property

Public Sub AutoSize(Optional ByVal iSize As eListViewColumnAutoSize = lvwColumnSizeToItemText)
    moOwner.fColumn_AutoSize miIndex, miId, iSize
End Sub

Public Sub Sort(Optional ByVal iType As eListViewSortType = NegOneL, Optional ByVal iOrder As eListViewSortOrder = NegOneL)
    moOwner.fColumn_Sort miIndex, miId, iType, iOrder
End Sub
