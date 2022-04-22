Sub NowToolbar()
' 添加工具栏菜单的过程
      Dim arr As Variant
      Dim id As Variant
      Dim i As Integer
      Dim Toolbar As CommandBar
      On Error Resume Next
      Application.CommandBars("MyToolbar").Delete
      arr = Array("Tool1", "Tool2", "Tool3", "Tool4", "Tool5", "Tool6")
      id = Array(9893, 284, 9590, 9614, 707, 986)
      Set Toolbar = Application.CommandBars.Add("MyToolbar", msoBarTop)
          With Toolbar
              .Protection = msoBarNoResize
              .Visible = True
              For i = 0 To 5
                  With .Controls.Add(Type:=msoControlButton)
                      .Caption = arr(i)
                      .FaceId = id(i)
                      .BeginGroup = True
                      .Style = msoButtonIconAndCaptionBelow
                  End With
              Next
          End With
      Set Toolbar = Nothing
End Sub

Sub DelToolbar()
'删除自定义工具栏的过程   
    Application.CommandBars("MyToolbar").Delete    
End Sub


Sub del_header()
'恢复工资表
    Rows("6:9").Select
    For i = 1 To ((ActiveSheet.UsedRange.Rows.Count \ 5) - 1)
        Application.CutCopyMode = False
        Selection.Delete Shift:=xlUp
        ActiveCell.Offset(1, 0).Rows("1:3").EntireRow.Select
    Next    
End Sub


Sub row_count()
    Debug.Print ActiveSheet.UsedRange.Rows.Count
End Sub

Sub copy_header()
'生成工资表
    Rows("1:4").Select    
    For i = 1 To ((ActiveSheet.UsedRange.Rows.Count - 4) - 1)   
        Selection.Copy
        ActiveCell.Offset(5, 0).Rows("1:4").EntireRow.Select
        Selection.Insert Shift:=xlDown    
    Next    
End Sub





