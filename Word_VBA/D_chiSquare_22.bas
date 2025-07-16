Attribute VB_Name = "a_Check_chiSquare_WithoutKimi"
Sub D_ChiSquare_0()
    Dim sel As Selection
    Dim cellText1 As String, cellText2 As String, cellText3 As String, cellText4 As String
    Dim a As Double, b As Double, C As Double, D As Double
    Dim O As Double, p As Double, x As Double, y As Double
    Dim ChiSquare As Double, pValue As Double
    Dim result As String
    Dim aExp As Double, bExp As Double, cExp As Double, dExp As Double
    Dim hasSmallValue As Boolean

    ' 获取当前选中的内容
    Set sel = Selection

    ' 检查选中的单元格数量
    Select Case sel.Cells.Count
        Case 4 ' 模式一：直接提取 A, B, C, D
            ' 获取四个单元格的内容
            cellText1 = sel.Cells(1).Range.text
            cellText2 = sel.Cells(2).Range.text
            cellText3 = sel.Cells(3).Range.text
            cellText4 = sel.Cells(4).Range.text
            ' 去除表格符号（如段落标记）并转换为半角
            cellText1 = StrConv(Replace(cellText1, vbCr, ""), vbNarrow)
            cellText2 = StrConv(Replace(cellText2, vbCr, ""), vbNarrow)
            cellText3 = StrConv(Replace(cellText3, vbCr, ""), vbNarrow)
            cellText4 = StrConv(Replace(cellText4, vbCr, ""), vbNarrow)
            ' 提取数字（A, B, C, D）
            a = ExtractNumber(cellText1)
            b = ExtractNumber(cellText2)
            C = ExtractNumber(cellText3)
            D = ExtractNumber(cellText4)
        Case 2 ' 模式二：提取 A, B, O, P，计算 X, Y, C, D
            ' 获取两个单元格的内容
            cellText1 = sel.Cells(1).Range.text
            cellText2 = sel.Cells(2).Range.text
            ' 去除表格符号（如段落标记）并转换为半角
            cellText1 = StrConv(Replace(cellText1, vbCr, ""), vbNarrow)
            cellText2 = StrConv(Replace(cellText2, vbCr, ""), vbNarrow)
            ' 提取括号前的数字（A 和 B）
            a = ExtractNumberBeforeParenthesis(cellText1)
            b = ExtractNumberBeforeParenthesis(cellText2)
            ' 提取括号后的数字（O 和 P）
            O = ExtractNumberAfterParenthesis(cellText1)
            p = ExtractNumberAfterParenthesis(cellText2)
            ' 计算 X, Y, C, D
            x = Round((a / O) * 100, 0)
            y = Round((b / p) * 100, 0)
            C = x - a
            D = y - b
        Case Else
            MsgBox "请选中两个或四个单元格！", vbExclamation
            Exit Sub
    End Select

    ' 计算理论频数
    Dim total As Double
    total = a + b + C + D
    aExp = (a + b) * (a + C) / total
    bExp = (a + b) * (b + D) / total
    cExp = (C + D) * (a + C) / total
    dExp = (C + D) * (b + D) / total
    
    ' 检查是否有理论频数小于5的单元格
    hasSmallValue = (aExp < 5 Or bExp < 5 Or cExp < 5 Or dExp < 5)

    ' 计算卡方值（使用Pearson卡方检验公式）
    ChiSquare = ((a - aExp) ^ 2 / aExp) + _
                ((b - bExp) ^ 2 / bExp) + _
                ((C - cExp) ^ 2 / cExp) + _
                ((D - dExp) ^ 2 / dExp)

    ' 计算p值 (自由度=1)
    pValue = ExactChiSquarePValue(ChiSquare, 1)

    ' 构建结果字符串（不显示理论频数）
    result = "A: " & a & vbCrLf & _
             "B: " & b & vbCrLf & _
             "C: " & C & vbCrLf & _
             "D: " & D & vbCrLf & _
             "卡方值: " & Round(ChiSquare, 4) & vbCrLf & _
             "P 值: " & Format(pValue, "0.0000")
    
    ' 只有当存在理论频数小于5时才添加警告
    If hasSmallValue Then
        result = result & vbCrLf & vbCrLf & _
                 "警告: 存在理论频数小于5的单元格(" & _
                 IIf(aExp < 5, "A ", "") & _
                 IIf(bExp < 5, "B ", "") & _
                 IIf(cExp < 5, "C ", "") & _
                 IIf(dExp < 5, "D", "") & _
                 ")，建议使用Fisher精确检验！"
    End If

    ' 删除现有批注
    Dim cmt As comment
    For Each cmt In sel.Range.Comments
        cmt.Delete
    Next cmt
    
    ' 添加新批注
    sel.Range.Comments.Add sel.Range, result
End Sub

' ========== 辅助函数 ==========

' 提取数字
Function ExtractNumber(text As String) As Double
    Dim numStr As String
    Dim i As Integer
    Dim char As String
    ' 提取数字部分
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        If IsNumeric(char) Or char = "." Then
            numStr = numStr & char
        End If
    Next i
    ' 转换为数字
    If IsNumeric(numStr) Then
        ExtractNumber = CDbl(numStr)
    Else
        ExtractNumber = 0
    End If
End Function

' 提取括号前的数字
Function ExtractNumberBeforeParenthesis(text As String) As Double
    Dim pos As Integer
    Dim numStr As String
    ' 查找括号的位置
    pos = InStr(text, "(")
    ' 提取括号前的数字
    If pos > 1 Then
        numStr = Trim(Left(text, pos - 1))
    Else
        numStr = ""
    End If
    ' 转换为数字
    If IsNumeric(numStr) Then
        ExtractNumberBeforeParenthesis = CDbl(numStr)
    Else
        ExtractNumberBeforeParenthesis = 0
    End If
End Function

' 提取括号内的数字
Function ExtractNumberAfterParenthesis(text As String) As Double
    Dim posStart As Integer, posEnd As Integer
    Dim numStr As String
    ' 查找括号的位置
    posStart = InStr(text, "(")
    posEnd = InStr(text, ")")
    ' 提取括号内的数字
    If posStart > 0 And posEnd > posStart Then
        numStr = Trim(Mid(text, posStart + 1, posEnd - posStart - 1))
    Else
        numStr = ""
    End If
    ' 去除百分号并转换为数字
    numStr = Replace(numStr, "%", "")
    If IsNumeric(numStr) Then
        ExtractNumberAfterParenthesis = CDbl(numStr)
    Else
        ExtractNumberAfterParenthesis = 0
    End If
End Function

' ========== 统计计算函数 ==========

' 修正后的精确卡方分布p值计算函数
Function ExactChiSquarePValue(x As Double, df As Integer) As Double
    ' 确保输入有效性
    If x <= 0 Or df <= 0 Then
        ExactChiSquarePValue = 1
        Exit Function
    End If
    
    ' 对极小的卡方值直接返回1
    If x < 0.000001 Then
        ExactChiSquarePValue = 1
        Exit Function
    End If
    
    ' 对自由度=1的特殊处理
    If df = 1 Then
        ' 确保p值不超过1
        Dim sqrtX As Double
        sqrtX = Sqr(x)
        If sqrtX > 8 Then  ' 对于非常大的x值，直接返回0
            ExactChiSquarePValue = 0
        Else
            ExactChiSquarePValue = 2 * (1 - NormalCDF(sqrtX))
            ' 确保p值在0-1范围内
            If ExactChiSquarePValue > 1 Then ExactChiSquarePValue = 1
            If ExactChiSquarePValue < 0 Then ExactChiSquarePValue = 0
        End If
    Else
        ' 对于其他自由度，使用不完全Gamma函数
        Dim gammaResult As Double
        gammaResult = IncompleteGamma(df / 2, x / 2)
        ExactChiSquarePValue = 1 - gammaResult
        ' 确保p值在0-1范围内
        If ExactChiSquarePValue > 1 Then ExactChiSquarePValue = 1
        If ExactChiSquarePValue < 0 Then ExactChiSquarePValue = 0
    End If
End Function

' 高精度标准正态CDF计算（修正版）
Function NormalCDF(x As Double) As Double
    ' 使用Hart近似算法，增加边界检查
    Dim z As Double, t As Double, y As Double
    
    ' 处理极大值
    If x > 8 Then
        NormalCDF = 1
        Exit Function
    End If
    
    ' 处理极小值
    If x < -8 Then
        NormalCDF = 0
        Exit Function
    End If
    
    Const a1 As Double = 0.254829592
    Const a2 As Double = -0.284496736
    Const a3 As Double = 1.421413741
    Const a4 As Double = -1.453152027
    Const a5 As Double = 1.061405429
    Const pp As Double = 0.3275911
    
    z = Abs(x)
    t = 1# / (1# + pp * z)
    y = 1# - (((((a5 * t + a4) * t + a3) * t + a2) * t + a1) * t) * Exp(-z * z)
    
    ' 确保结果在0-1范围内
    If y > 1 Then y = 1
    If y < 0 Then y = 0
    
    If x > 0 Then
        NormalCDF = y
    Else
        NormalCDF = 1 - y
    End If
End Function

' 不完全Gamma函数实现（增加边界检查）
Function IncompleteGamma(a As Double, x As Double) As Double
    ' 检查输入有效性
    If x < 0 Or a <= 0 Then
        IncompleteGamma = 0
        Exit Function
    End If
    
    ' 对极小的x值直接返回0
    If x < 0.0000000001 Then
        IncompleteGamma = 0
        Exit Function
    End If
    
    ' 对极大的x值直接返回1
    If x > 100000 Then
        IncompleteGamma = 1
        Exit Function
    End If
    
    Dim gamser As Double, gln As Double
    Dim gamcf As Double
    
    If x < a + 1 Then
        ' 使用级数展开
        Call GammaSeries(gamser, a, x, gln)
        ' 确保结果在0-1范围内
        If gamser > 1 Then gamser = 1
        If gamser < 0 Then gamser = 0
        IncompleteGamma = gamser
    Else
        ' 使用连分式展开
        Call GammaCF(gamcf, a, x, gln)
        ' 确保结果在0-1范围内
        If gamcf > 1 Then gamcf = 1
        If gamcf < 0 Then gamcf = 0
        IncompleteGamma = 1 - gamcf
    End If
End Function

' Gamma级数展开
Sub GammaSeries(gamser As Double, a As Double, x As Double, gln As Double)
    Dim n As Integer
    Dim sum As Double, del As Double, ap As Double
    
    Const ITMAX As Integer = 100
    Const EPS As Double = 0.000000000000001
    
    gln = LogGamma(a)
    If x <= 0 Then
        gamser = 0
        Exit Sub
    End If
    
    ap = a
    sum = 1 / a
    del = sum
    
    For n = 1 To ITMAX
        ap = ap + 1
        del = del * x / ap
        sum = sum + del
        If Abs(del) < Abs(sum) * EPS Then Exit For
    Next n
    
    gamser = sum * Exp(-x + a * Log(x) - gln)
End Sub

' Gamma连分式展开
Sub GammaCF(gamcf As Double, a As Double, x As Double, gln As Double)
    Dim n As Integer
    Dim gold As Double, g As Double, fac As Double
    Dim b1 As Double, b0 As Double
    Dim anf As Double, ana As Double
    Dim an As Double, a1 As Double
    
    Const ITMAX As Integer = 100
    Const EPS As Double = 0.000000000000001
    
    gln = LogGamma(a)
    gold = 0
    a1 = 1
    b0 = 1
    b1 = x
    fac = 1
    
    For n = 1 To ITMAX
        an = CDbl(n)
        ana = an - a
        a1 = (a1 + ana) * fac
        b1 = x * b1 + ana * a1
        fac = 1 / a1
        g = b1 * fac
        If Abs((g - gold) / g) < EPS Then Exit For
        gold = g
    Next n
    
    gamcf = Exp(-x + a * Log(x) - gln) * g
End Sub

' 计算Gamma函数的对数
Function LogGamma(x As Double) As Double
    Dim y As Double, tmp As Double, ser As Double
    Dim cof(6) As Double
    Dim j As Integer
    
    cof(0) = 76.1800917294715
    cof(1) = -86.5053203294168
    cof(2) = 24.0140982408309
    cof(3) = -1.23173957245015
    cof(4) = 1.20865097386618E-03
    cof(5) = -5.395239384953E-06
    
    y = x
    tmp = x + 5.5
    tmp = tmp - (x + 0.5) * Log(tmp)
    ser = 1.00000000019001
    
    For j = 0 To 5
        y = y + 1
        ser = ser + cof(j) / y
    Next j
    
    LogGamma = -tmp + Log(2.506628274631 * ser / x)
End Function

