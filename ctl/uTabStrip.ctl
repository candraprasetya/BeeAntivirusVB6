VERSION 5.00
Begin VB.UserControl uTabStrip 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox tab1 
      Height          =   3015
      Left            =   0
      ScaleHeight     =   2955
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   0
      Width           =   3015
   End
   Begin VB.TextBox txt 
      Height          =   285
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "uTabStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'===================================
'ucTabSonny coded by Sonny Lazuardi
'using vbcomctl and ucTabStrip
'(c) Copyleft 2009 Bandung Indonesia
'===================================
'uctabsonny mampu menampung objek secara aktif (active drawing)
'mengganti judul tab dan merubah banyak tab secara aktif
'enjoy :D
Dim i_count As Integer
Dim i_aktif As Integer
Dim i_laktif As Integer
Dim m_lMoveOffset As Long
Dim s_Judul() As String
Dim s_Tampung() As String
Const def_Count As Integer = 1
Const def_AktifTab As Integer = 1
Public Event Click(Index As Integer)
Public Sub GantiJudul(TabAktif As Integer, Judul As String)
If TabAktif > i_count Or TabAktif < 1 Then Exit Sub
On Error Resume Next
tab1.Tabs.Item(TabAktif).Text = Judul
s_Judul(TabAktif) = Judul
Reload i_count
End Sub

Private Sub tab1_Click(ByVal oTab As cTab)
txt = oTab.Index
End Sub

Private Sub txt_Change()
    On Error Resume Next
    RaiseEvent Click(CInt(txt.Text))
    Call UbahAktif(CInt(txt.Text))
End Sub

Private Sub UserControl_InitProperties()
i_count = def_Count
i_aktif = def_AktifTab
UbahJudul def_Count
End Sub
Private Sub UserControl_Resize()
tab1.Width = UserControl.Width
tab1.Height = UserControl.Height
End Sub
Public Property Get JudulTab() As String
JudulTab = tab1.Tabs.Item(i_aktif).Text
End Property
Public Property Let JudulTab(vKop As String)
tab1.Tabs.Item(i_aktif).Text = vKop
s_Judul(i_aktif) = vKop
Reload i_aktif
End Property
Public Property Get AktifTab() As Integer
AktifTab = i_aktif
tab1.SetSelectedTab i_aktif
End Property
Public Property Let AktifTab(vAktif As Integer)
On Error Resume Next
If vAktif > i_count Or vAktif = 0 Then Exit Property
UbahAktif vAktif
End Property
Private Sub UbahAktif(vAct As Integer)
If vAct = i_aktif Then Exit Sub
If UserControl.Parent.Width > 10000 Then
    m_lMoveOffset = UserControl.Parent.Width + 1000
Else
    m_lMoveOffset = 10000
End If
i_laktif = i_aktif
i_aktif = vAct
tab1.SetSelectedTab vAct
UserControl.PropertyChanged AktifTab
UserControl.PropertyChanged JudulTab
HandleContainedControls vAct
End Sub
Public Property Get TabCount() As Integer
TabCount = i_count
End Property
Public Property Let TabCount(ByVal vTabCount As Integer)
If vTabCount < 1 Then Exit Property
Dim i As Integer
i_count = vTabCount
UbahJudul vTabCount
End Property
Private Sub UbahJudul(vTabs As Integer)
On Error Resume Next
tab1.Tabs.Clear
ReDim s_Judul(vTabs) As String
For i = 1 To vTabs
If s_Tampung(i) <> "" Then s_Judul(i) = s_Tampung(i)
tab1.Tabs.Add s_Judul(i)
Next
Reload vTabs
End Sub
Private Sub Reload(vTabs As Integer)
ReDim s_Tampung(vTabs) As String
For i = 1 To vTabs
s_Tampung(i) = s_Judul(i)
Next
End Sub
Private Sub HandleContainedControls(ByVal New_ActiveTab As Long)
    On Error Resume Next
    Dim Ctl As Control
    Dim MoveVal As Long
 
    On Error Resume Next
    '   The difference between what was the active
    '   Tab and the newly set activetab
    MoveVal = (New_ActiveTab - i_laktif)
    '   Move the controls by a Factor which is
    '   tied to the Tab Diff....the default value
    '   is set to 10K, but for Objects greater than
    '   this, size the Width + 1000 will be used.
    MoveVal = (MoveVal * m_lMoveOffset)
    '   This is what creates the illusion of
    '   Changing the Tab of a tab control
    For Each Ctl In UserControl.ContainedControls
         Ctl.Left = (Ctl.Left + MoveVal)
    Next Ctl

End Sub

Private Sub UserControl_Show()
tab1.SetSelectedTab i_aktif
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "tabcount", i_count, def_Count
.WriteProperty "aktif", i_aktif, def_AktifTab
For i = 1 To i_count
.WriteProperty "judul(" & i & ")", s_Judul(i), ""
Next
End With
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
i_count = .ReadProperty("tabcount", def_Count)
i_aktif = .ReadProperty("aktif", def_AktifTab)
ReDim s_Tampung(i_count) As String
For i = 1 To i_count
s_Tampung(i) = .ReadProperty("judul(" & i & ")", "")
Next
End With
UbahJudul i_count
End Sub

