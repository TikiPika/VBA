Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' 定义需要弹出日期选择控件的单元格区域
    Dim rng As Range
    Set rng = Range("A1:A10, B1:B10, C1:C10") ' 修改为你的目标单元格区域

    ' 判断双击的单元格是否在指定区域内
    If Not Intersect(Target, rng) Is Nothing Then
        Dim dtPicker As Object
        Set dtPicker = CreateObject("MSComCtl2.DTPicker.2")
        
        ' 如果目标单元格是合并单元格
        If Target.MergeCells Then
            With Target.MergeArea
                dtPicker.Top = .Top
                dtPicker.Left = .Left
                dtPicker.Width = .Width
                dtPicker.Height = .Height
            End With
        Else
            ' 如果目标单元格不是合并单元格
            With Target
                dtPicker.Top = .Top
                dtPicker.Left = .Left
                dtPicker.Width = .Width
                dtPicker.Height = .Height
            End With
        End If
        
        With dtPicker
            .Visible = True
            .Value = Date
            .DropDown
        End With
        
        Cancel = True
        Target.Value = dtPicker.Value
        dtPicker.Visible = False
        Set dtPicker = Nothing
    End If
End Sub
