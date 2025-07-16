Attribute VB_Name = "D_ChiSquareTest_2x3"
Sub ChiSquareTest_2x3()
    On Error GoTo ErrorHandler

    ' 声明变量（保持不变）
    Dim oSelection As Object
    Dim cellCount As Integer, rowCount As Integer, colCount As Integer
    Dim table(0 To 1, 0 To 2) As Double
    Dim chiSquare As Double, pValue As Double
    Dim rowTotal(0 To 1) As Double, colTotal(0 To 2) As Double
    Dim total As Double, i As Integer, j As Integer
    Dim result As String
    Dim expected As Double, diff As Double
    Dim cmt As Object
    Dim cellText As String
    Dim bErrorOccurred As Boolean

    Dim observed() As Double
    Dim success As Double, totalNum As Double, fail As Double

    ' 初始化
    bErrorOccurred = False
    result = "卡方检验结果(2×3):" & vbCrLf & vbCrLf

    Set oSelection = Application.selection

    If oSelection Is Nothing Or oSelection.Cells.Count = 0 Then
        result = result & "错误: 请先选中表格中的单元格。" & vbCrLf
        GoTo FinalOutput
    End If

    cellCount = oSelection.Cells.Count
    rowCount = oSelection.Rows.Count
    colCount = oSelection.Columns.Count

    ' 模式一：选择6个单元格（2×3格式）（保持不变）
    If cellCount = 6 Then
        ReDim observed(0 To 5)
        For i = 1 To 6
            On Error Resume Next
            cellText = oSelection.Cells(i).Range.text
            cellText = CleanCellText(cellText)
            observed(i - 1) = ExtractNumber(cellText)
            If Err.Number <> 0 Then
                result = result & "错误: 单元格" & i & "数据提取失败。" & vbCrLf
                bErrorOccurred = True
                Exit For
            End If
            On Error GoTo ErrorHandler
        Next i

        If Not bErrorOccurred Then
            If rowCount = 2 And colCount = 3 Then
                For i = 0 To 1
                    For j = 0 To 2
                        table(i, j) = observed(i * 3 + j)
                    Next j
                Next i
            ElseIf rowCount = 3 And colCount = 2 Then
                For i = 0 To 2
                    For j = 0 To 1
                        table(j, i) = observed(i * 2 + j)
                    Next j
                Next i
            ElseIf rowCount = 6 And colCount = 1 Then
                For i = 0 To 2
                    table(0, i) = observed(i)
                    table(1, i) = observed(i + 3)
                Next i
            ElseIf rowCount = 1 And colCount = 6 Then
                For i = 0 To 2
                    table(0, i) = observed(i)
                    table(1, i) = observed(i + 3)
                Next i
            Else
                result = result & "错误: 请选择2×3、3×2、6×1或1×6的单元格区域！" & vbCrLf
                bErrorOccurred = True
            End If
        End If

    ' 模式二：选择3个单元格（n(N%)格式）（核心修正）
    ElseIf cellCount = 3 Then
        ReDim observed(0 To 5)
        For i = 1 To 3
            On Error Resume Next
            cellText = oSelection.Cells(i).Range.text
            cellText = CleanCellText(cellText)   ' 清理文本
            
            ' 提取括号前的成功数（强化校验）
            success = ExtractNumberBeforeParenthesis(cellText)
            ' 提取总人数（基于括号内的百分比）
            totalNum = ExtractTotalFromPercentage(cellText, success)
            ' 计算失败数（组2数据）
            fail = totalNum - success

            If Err.Number <> 0 Then
                result = result & "错误: 单元格" & i & "数据提取失败：" & Err.Description & vbCrLf
                bErrorOccurred = True
                Exit For
            End If
            On Error GoTo ErrorHandler

            observed((i - 1) * 2) = success     ' 组1数据
            observed((i - 1) * 2 + 1) = fail     ' 组2数据
        Next i

        If Not bErrorOccurred Then
            For i = 0 To 2
                table(0, i) = observed(i * 2)
                table(1, i) = observed(i * 2 + 1)
            Next i
        End If
    Else
        result = result & "错误: 请选择3个（n(N%)格式）或6个（2×3格式）单元格！" & vbCrLf
        bErrorOccurred = True
    End If

    ' 后续校验及计算（保持不变）
    If Not bErrorOccurred Then
        For i = 0 To 1
            For j = 0 To 2
                If table(i, j) < 0 Or Not IsNumeric(table(i, j)) Then
                    result = result & "错误: 数据无效，单元格(" & i & "," & j & ")包含无效数据。" & vbCrLf
                    bErrorOccurred = True
                    Exit For
                End If
            Next j
            If bErrorOccurred Then Exit For
        Next i

        If Not bErrorOccurred Then
            rowTotal(0) = table(0, 0) + table(0, 1) + table(0, 2)
            rowTotal(1) = table(1, 0) + table(1, 1) + table(1, 2)
            colTotal(0) = table(0, 0) + table(1, 0)
            colTotal(1) = table(0, 1) + table(1, 1)
            colTotal(2) = table(0, 2) + table(1, 2)
            total = rowTotal(0) + rowTotal(1)

            For i = 0 To 1
                For j = 0 To 2
                    expected = (rowTotal(i) * colTotal(j)) / total
                    If expected < 1 Then
                        result = result & "警告: 期望值(" & i & "," & j & ")=" & Format(expected, "0.00") & " 过小，卡方检验可能不适用。" & vbCrLf
                    End If
                Next j
            Next i

            chiSquare = 0
            For i = 0 To 1
                For j = 0 To 2
                    expected = (rowTotal(i) * colTotal(j)) / total
                    If expected > 0 Then
                        diff = table(i, j) - expected
                        chiSquare = chiSquare + (diff * diff) / expected
                    End If
                Next j
            Next i

            pValue = 1 - Chi2CDF(chiSquare, 2)

            result = result & "观测数据:" & vbCrLf
            result = result & "组1: " & table(0, 0) & " | " & table(0, 1) & " | " & table(0, 2) & vbCrLf
            result = result & "组2: " & table(1, 0) & " | " & table(1, 1) & " | " & table(1, 2) & vbCrLf & vbCrLf
            result = result & "卡方值 = " & Format(chiSquare, "0.0000") & vbCrLf
            result = result & "P值 = " & Format(pValue, "0.0000") & vbCrLf
            result = result & "自由度 = 2"
        End If
    End If

FinalOutput:
    On Error Resume Next
    If Not oSelection.Comments Is Nothing Then
        If oSelection.Comments.Count > 0 Then oSelection.Comments(1).Delete
    End If

    If Not ActiveDocument Is Nothing Then
        Set cmt = ActiveDocument.Comments.add(oSelection.Range, result)
    End If

    Debug.Print result
    Exit Sub

ErrorHandler:
    result = result & "运行时错误: " & Err.Description & " (错误号: " & Err.Number & ")" & vbCrLf
    Resume FinalOutput
End Sub

' ========== 辅助函数（修正后） ==========

' 清理单元格文本（增强版）
Function CleanCellText(text As String) As String
    text = Replace(text, Chr(13), "")         ' 移除换行符
    text = Replace(text, Chr(7), "")          ' 移除隐藏字符
    text = Replace(text, ChrW(12288), " ")    ' 全角空格转半角
    text = Replace(text, "（", "(")           ' 全角括号转半角
    text = Replace(text, "）", ")")
    text = Replace(text, " ", "")             ' 移除所有空格（关键修正：避免空格干扰提取）
    CleanCellText = text
End Function

' 提取括号前的数字（强化校验）
Function ExtractNumberBeforeParenthesis(text As String) As Double
    Dim pos As Integer
    ' 查找左括号（半角或全角）
    pos = InStr(text, "(")
    If pos = 0 Then pos = InStr(text, "（")
    
    If pos = 0 Then
        Err.Raise 1001, , "未找到左括号，请检查格式（应为n(百分比)）"
    End If
    
    ' 提取括号前的内容并转换为数字
    Dim numText As String
    numText = Left(text, pos - 1)
    If Not IsNumeric(numText) Then
        Err.Raise 1002, , "括号前的内容不是有效数字：" & numText
    End If
    
    ExtractNumberBeforeParenthesis = CDbl(numText)
End Function

' 从括号中提取总人数（核心修正：使用正则提取数字）
Function ExtractTotalFromPercentage(text As String, success As Double) As Double
    Dim pos1 As Integer, pos2 As Integer
    Dim percentText As String
    Dim percent As Double
    Dim regEx As Object, matches As Object
    
    ' 初始化正则对象（用于提取数字）
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "\d+(\.\d+)?"  ' 匹配整数或小数（如19.10、38.3）
    regEx.Global = False
    
    ' 查找括号位置
    pos1 = InStr(text, "(")
    If pos1 = 0 Then pos1 = InStr(text, "（")
    pos2 = InStr(text, ")")
    If pos2 = 0 Then pos2 = InStr(text, "）")
    
    ' 校验括号有效性
    If pos1 = 0 Or pos2 = 0 Or pos1 >= pos2 Then
        Err.Raise 1003, , "未找到有效括号对，请检查格式"
    End If
    
    ' 提取括号内的内容
    percentText = Mid(text, pos1 + 1, pos2 - pos1 - 1)
    
    ' 用正则提取括号内的数字（忽略其他字符）
    Set matches = regEx.Execute(percentText)
    If matches.Count = 0 Then
        Err.Raise 1004, , "括号内未找到有效数字（应为百分比，如19.10）"
    End If
    percentText = matches(0).value  ' 提取第一个匹配的数字
    
    ' 转换为百分比数值
    If Not IsNumeric(percentText) Then
        Err.Raise 1005, , "提取的百分比不是有效数字：" & percentText
    End If
    percent = CDbl(percentText)
    
    ' 处理百分比格式（如19.10 → 0.1910）
    If percent > 1 Then
        percent = percent / 100
    End If
    
    ' 校验百分比有效性
    If percent <= 0 Then
        Err.Raise 1006, , "百分比必须大于0（当前为：" & percentText & "）"
    End If
    
    ' 计算总人数（四舍五入为整数）
    ExtractTotalFromPercentage = Round(success / percent, 0)
    
    ' 清理对象
    Set matches = Nothing
    Set regEx = Nothing
End Function

' 其他辅助函数（保持不变）
Function ExtractNumber(text As String) As Double
    On Error Resume Next
    If Len(Trim(text)) = 0 Then
        ExtractNumber = 0
        Exit Function
    End If
    ExtractNumber = CDbl(Trim(text))
    If Err.Number <> 0 Then Err.Raise 13
End Function

Function Chi2CDF(x As Double, df As Integer) As Double
    If df <> 2 Then
        Chi2CDF = CVErr(1)
        Exit Function
    End If
    Chi2CDF = 1 - Exp(-x / 2)
End Function
