Attribute VB_Name = "D_COX_0"
Sub D_Cox_0()
    Dim selectedText As String
    Dim a As Double
    Dim b As Double
    Dim commentText As String
    Dim comment As comment
    Dim Beta As Double
    Dim SE As Double
    Dim HR As Double
    Dim ChiSquare As Double
    Dim PValue As Double
    Dim isSingleCell As Boolean
    Dim leftCell As Range
    Dim rightCell As Range
    Dim numMatches As Object
    Dim regex As Object
    Dim bracketContent As String
    
    ' 获取选中的文本
    selectedText = Selection.text
    
    ' 判断是单个单元格还是多个单元格
    If Selection.Cells.Count = 1 Then
        ' 单个单元格，提取括号内的内容（如果有）
        bracketContent = ExtractBracketContent(selectedText)
        If bracketContent = "" Then
            ' 如果没有括号，则使用整个文本
            bracketContent = selectedText
        End If
        
        ' 提取数字
        Set regex = CreateObject("VBScript.RegExp")
        regex.Global = True
        regex.IgnoreCase = True
        regex.Pattern = "-?\d+\.?\d*" ' 匹配正负整数或小数
        
        ' 查找所有匹配的数字
        Set numMatches = regex.Execute(bracketContent)
        
        ' 检查是否找到至少两个数字
        If numMatches.Count < 2 Then
            MsgBox "未找到足够的数字，无法提取 A 和 B。", vbExclamation
            Exit Sub
        End If
        
        ' 提取 A 和 B
        a = Val(numMatches(0).value)
        b = Val(numMatches(1).value)
        
        ' 处理 A-B 形式的文本（确保 B 为正数）
        If InStr(bracketContent, "-") > 0 And numMatches.Count = 2 Then
            b = Abs(b) ' 取 B 的绝对值
        End If
    ElseIf Selection.Cells.Count = 2 Then
        ' 两个单元格，分别提取 A 和 B
        Set leftCell = Selection.Cells(1).Range
        Set rightCell = Selection.Cells(2).Range
        a = Val(leftCell.text)
        b = Val(rightCell.text)
    Else
        MsgBox "请选中一个单元格（包含 A 和 B）或两个相邻单元格（A 在左，B 在右）。", vbExclamation
        Exit Sub
    End If
    
    ' 验证 A 和 B
    If b <= a Then
        MsgBox "B 必须大于 A。", vbExclamation
        Exit Sub
    End If
    
    ' 计算 Cox 回归系数 β
    Beta = (Log(a) + Log(b)) / 2
    
    ' 计算标准误 SE
    SE = (Log(b) - Log(a)) / (2 * 1.96)
    
    ' 计算风险比 HR
    HR = Exp(Beta)
    
    ' 计算卡方值 (Wald检验)
    ChiSquare = (Beta / SE) ^ 2
    
    ' 计算P值 (使用优化的近似方法)
    PValue = ImprovedChiSquarePValue(ChiSquare, 1)
    
    ' 生成批注内容
    commentText = "HR: " & Format(HR, "0.0000") & vbCrLf & _
                  "95%CI: " & Format(a, "0.0000") & " - " & Format(b, "0.0000") & vbCrLf & _
                  "β: " & Format(Beta, "0.0000") & vbCrLf & _
                  "SE: " & Format(SE, "0.0000") & vbCrLf & _
                  "χ2: " & Format(ChiSquare, "0.0000") & vbCrLf & _
                  "P值: " & Format(PValue, "0.0000")
    
    ' 添加批注
    With Selection
        ' 删除原有的批注（如果存在）
        If .Comments.Count > 0 Then
            .Comments(1).Delete
        End If
        ' 添加新批注
        Set comment = .Range.Comments.Add(Range:=.Range, text:=commentText)
    End With
End Sub

Function ExtractBracketContent(text As String) As String
    ' 提取括号内的内容（支持全角和半角括号）
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "[$（].*?[$）]" ' 匹配括号及其内容
    
    ' 查找匹配的括号内容
    Dim matches As Object
    Set matches = regex.Execute(text)
    
    If matches.Count > 0 Then
        ' 提取括号内的内容（去掉括号）
        ExtractBracketContent = Mid(matches(0).value, 2, Len(matches(0).value) - 2)
    Else
        ' 未找到括号
        ExtractBracketContent = ""
    End If
End Function

Function ImprovedChiSquarePValue(ChiSquare As Double, df As Integer) As Double
    ' 优化的卡方P值计算方法（适用于df=1）
    ' 使用分段近似法提高精度
    
    ' 预先计算的关键点（卡方值, P值）
    Static keyPoints(1 To 7, 1 To 2) As Double
    keyPoints(1, 1) = 0#:     keyPoints(1, 2) = 1#
    keyPoints(2, 1) = 0.455:  keyPoints(2, 2) = 0.5
    keyPoints(3, 1) = 1.642:  keyPoints(3, 2) = 0.2
    keyPoints(4, 1) = 2.706:  keyPoints(4, 2) = 0.1
    keyPoints(5, 1) = 3.841:  keyPoints(5, 2) = 0.05
    keyPoints(6, 1) = 5.024:  keyPoints(6, 2) = 0.025
    keyPoints(7, 1) = 6.635:  keyPoints(7, 2) = 0.01
    
    ' 边界检查
    If ChiSquare <= 0 Then
        ImprovedChiSquarePValue = 1
        Exit Function
    ElseIf ChiSquare >= keyPoints(7, 1) Then
        ImprovedChiSquarePValue = keyPoints(7, 2)
        Exit Function
    End If
    
    ' 线性插值
    Dim i As Integer
    For i = 1 To 6
        If ChiSquare >= keyPoints(i, 1) And ChiSquare < keyPoints(i + 1, 1) Then
            ImprovedChiSquarePValue = keyPoints(i, 2) + (keyPoints(i + 1, 2) - keyPoints(i, 2)) * _
                                    (ChiSquare - keyPoints(i, 1)) / (keyPoints(i + 1, 1) - keyPoints(i, 1))
            Exit Function
        End If
    Next i
    
    ' 默认返回最小P值
    ImprovedChiSquarePValue = keyPoints(7, 2)
End Function

