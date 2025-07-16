Attribute VB_Name = "D_round_off_tablecolumn"
'对表格中的某一列四舍五入，完成后显示红色字体
Sub d_Round_off_TableColumn_Withoutkimi()
    Dim tbl As Table
    Dim cell As cell
    Dim colNum As Integer
    Dim cellValue As String
    Dim numericValue As Double
    Dim i As Integer
    
    ' 检查是否有选中的表格
    If Selection.Information(wdWithInTable) = False Then
        MsgBox "请先选中表格中的某一列！", vbExclamation
        Exit Sub
    End If
    
    ' 获取选中的列号
    colNum = Selection.Information(wdStartOfRangeColumnNumber)
    
    ' 获取选中的表格
    Set tbl = Selection.Tables(1)
    
    ' 遍历表格的每一行
    For i = 1 To tbl.Rows.Count
        ' 获取当前行的指定列的单元格
        Set cell = tbl.cell(i, colNum)
        
        ' 获取单元格内容并去除尾部的段落标记
        cellValue = Trim(cell.Range.text)
        If Right(cellValue, 2) = Chr(13) & Chr(7) Then
            cellValue = Left(cellValue, Len(cellValue) - 2)
        End If
        
        ' 检查是否为数字
        If IsNumeric(cellValue) Then
            ' 将字符串转换为数字
            numericValue = CDbl(cellValue)
            
            ' 四舍五入保留两位小数
            numericValue = Round(numericValue, 2)
            
            ' 将结果写回单元格，并设置字体颜色为红色
            With cell.Range
                .text = numericValue
                .Font.Color = wdColorRed
            End With
        End If
    Next i
    
    MsgBox "操作完成，修改内容已设置为红色字体！", vbInformation
End Sub
