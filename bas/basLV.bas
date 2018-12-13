Attribute VB_Name = "basLV"
Dim cImgMal  As New gComCtl

Public Function BuildLV()
'LV MAl
With frMain
       
    With .lvMal
        .View = lvwDetails
        .Font.FaceName = "Arial"
        .Columns.Add , "Virus Name", , , lvwAlignCenter, 2200
        .Columns.Add , "Virus Path", , , lvwAlignLeft, 5000
        .Columns.Add , "Size [B]", , , lvwAlignRight, 1300
        .Columns.Add , "Virus Status", , , lvwAlignLeft, 3200
        Set .ImageList = cImgMal.NewImageList(16, 16, imlColor32)
    End With
    
    With .lvQuar
        .View = lvwDetails
        .Font.FaceName = "Arial"
        .Columns.Add , "File Name", , , lvwAlignCenter, 2200
        .Columns.Add , "File old Path", , , lvwAlignLeft, 5000
        .Columns.Add , "Quar Path", , , lvwAlignLeft, 5000
        Set .ImageList = cImgMal.NewImageList(16, 16, imlColor32)
    End With
    
    With .lvVirLst
        .View = lvwDetails
        .Font.FaceName = "Arial"
        .Columns.Add , "Virus Name", , , lvwAlignCenter, 4400
        Set .ImageList = cImgMal.NewImageList(16, 16, imlColor32)
    End With
    
    With .lvProcess
          .Font.FaceName = "Arial"
          .Columns.Add , "Process Name", , , lvwAlignLeft, 2000
          .Columns.Add , "Startup", , , lvwAlignCenter, 1200
          .Columns.Add , "PID", , , lvwAlignCenter, 1200
          .Columns.Add , "Parent PID", , , lvwAlignCenter, 1300
          .Columns.Add , "Hidden", , , lvwAlignLeft, 1200
          .Columns.Add , "In Debug", , , lvwAlignLeft, 1300
          .Columns.Add , "Locked", , , lvwAlignLeft, 1300
          .Columns.Add , "Size [B]", , , lvwAlignRight, 1300
          .Columns.Add , "Update Path", , , lvwAlignLeft, 5000
          .Columns.Add , "Status", , , lvwAlignRight, 1200
    Set .ImageList = cImgMal.NewImageList(16, 16, imlColor32)
    End With
    
    With .lvDlock
    .Font.FaceName = "Arial"
    .Columns.Add , "Drive Name", , , lvwAlignCenter, 3000
    .Columns.Add , "Status", , , lvwAlignLeft, 3500
    Set .ImageList = cImgMal.NewImageList(16, 16, imlColor32)
    End With
    
    With .lvStartup
        .View = lvwDetails
        .Font.FaceName = "Arial"
        .Columns.Add , "Startup Name", , , lvwAlignCenter, 1800
        .Columns.Add , "Startup Path", , , lvwAlignLeft, 5000
        .Columns.Add , "Startup Reg Data", , , lvwAlignLeft, 5000
        .Columns.Add , "File Path", , , lvwAlignLeft, 5000
        .Columns.Add , "Evaluation", , , lvwAlignLeft, 1600
        .Columns.Add , "Status", , , lvwAlignLeft, 1300
    End With
End With

With frRTP
    With .lvRTP
        .View = lvwDetails
        .Font.FaceName = "Arial"
        .Columns.Add , "Virus Name", , , lvwAlignCenter, 2200
        .Columns.Add , "Virus Path", , , lvwAlignLeft, 5000
        .Columns.Add , "Size [B]", , , lvwAlignRight, 1300
        .Columns.Add , "Virus Status", , , lvwAlignLeft, 3200
        Set .ImageList = cImgMal.NewImageList(16, 16, imlColor32)
    End With
End With
InitImageList
End Function

Public Function InitImageList()
With frMain
    .lvMal.ImageList.AddFromDc .pic1.hdc, 16, 16
    .lvMal.ImageList.AddFromDc .pic2.hdc, 16, 16
    .lvMal.ImageList.AddFromDc .pic3.hdc, 16, 16
    .lvMal.ImageList.AddFromDc .pic4.hdc, 16, 16
        
    .lvQuar.ImageList.AddFromDc .pic8.hdc, 16, 16
    
    .lvVirLst.ImageList.AddFromDc .pic2.hdc, 16, 16
    
    .lvDlock.ImageList.AddFromDc frMain.pic6.hdc, 16, 16
    .lvDlock.ImageList.AddFromDc frMain.pic7.hdc, 16, 16
    
    frRTP.lvRTP.ImageList.AddFromDc .pic1.hdc, 16, 16
    frRTP.lvRTP.ImageList.AddFromDc .pic2.hdc, 16, 16
    frRTP.lvRTP.ImageList.AddFromDc .pic3.hdc, 16, 16
    frRTP.lvRTP.ImageList.AddFromDc .pic4.hdc, 16, 16
End With

End Function

Public Sub AddToLVMal(Lv As ucListView, sItem As String, sSub1 As String, sSub2 As String, sSub3 As String, iIcon As Long, nScroll As Long)
Dim lstLV As cListItem

Set lstLV = Lv.ListItems.Add(, sItem, , iIcon, , , , , "")
    lstLV.SubItem(2).Text = sSub1
    lstLV.SubItem(3).Text = sSub2
    lstLV.SubItem(4).Text = sSub3
If Lv.ListItems.count > nScroll Then Lv.Scroll 0, 25

Set lstLV = Nothing
End Sub

Public Sub AddToLVQuar(Lv As ucListView, sItem As String, sSub1 As String, sSub2 As String, iIcon As Long, nScroll As Long)
Dim lstLV As cListItem

Set lstLV = Lv.ListItems.Add(, sItem, , iIcon, , , , , "")
    lstLV.SubItem(2).Text = sSub1
    lstLV.SubItem(3).Text = sSub2
If Lv.ListItems.count > nScroll Then Lv.Scroll 0, 25

Set lstLV = Nothing
End Sub

Public Sub AddToLVStUP(Lv As ucListView, sItem As String, sSub1 As String, sSub2 As String, sSub3 As String, sTag1 As String, sTag2 As String, iIcon As Long, nScroll As Long)
Dim lstLV As cListItem

Set lstLV = Lv.ListItems.Add(, sItem, , iIcon, , , , , "")
    lstLV.SubItem(2).Text = sSub1
    
    lstLV.SubItem(3).Text = sSub2
    lstLV.SubItem(4).Text = sSub3
If Lv.ListItems.count > nScroll Then Lv.Scroll 0, 25

Set lstLV = Nothing
End Sub
