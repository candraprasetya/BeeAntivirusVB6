VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'cTab.cls                    12/15/04
'
'           PURPOSE:
'               Represent a tab in a tab control.
'
'==================================================================================================

Option Explicit

Private moOwner  As ucTabStrip
Private mhTab    As Long
Private miIndex  As Long

Friend Sub fInit(ByVal oOwner As ucTabStrip, ByVal hTab As Long, ByVal iIndex As Long)
    Set moOwner = oOwner
    mhTab = hTab
    miIndex = iIndex
End Sub

Friend Property Get fIsOwner(ByVal oOwner As ucTabStrip)
    fIsOwner = CBool(oOwner Is moOwner)
End Property

Public Property Get Text() As String
    Text = moOwner.fTab_Text(mhTab, miIndex)
End Property
Public Property Let Text(ByRef sNew As String)
    moOwner.fTab_Text(mhTab, miIndex) = sNew
End Property

Public Property Get IconIndex() As Long
    IconIndex = moOwner.fTab_IconIndex(mhTab, miIndex)
End Property
Public Property Let IconIndex(ByVal iNew As Long)
    moOwner.fTab_IconIndex(mhTab, miIndex) = iNew
End Property

Public Property Get Key() As String
    Key = moOwner.fTab_Key(mhTab, miIndex)
End Property
Public Property Let Key(ByRef sNew As String)
    moOwner.fTab_Key(mhTab, miIndex) = sNew
End Property

Public Property Get Index() As Long
    Index = moOwner.fTab_Index(mhTab, miIndex)
End Property
