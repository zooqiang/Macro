Attribute VB_Name = "D_t_text0"
' 两组均数的t检验（修正版）
Sub D_ExtractAnd_t_test_0()
    On Error GoTo ErrorHandler
    Dim cell1 As Range
    Dim cell2 As Range
    Dim text1 As String
    Dim text2 As String
    Dim x1 As Double, s1 As Double
    Dim x2 As Double, s2 As Double
    Dim n1 As Double, n2 As Double
    Dim t As Double, df As Double
    Dim pValue As Double
    
    ' 检查是否选中了两个单元格
    If Selection.Cells.Count <> 2 Then
        MsgBox "请选中两个单元格。"
        Exit Sub
    End If
    
    ' 获取两个单元格的内容
    Set cell1 = Selection.Cells(1).Range
    Set cell2 = Selection.Cells(2).Range
    
    ' 去除表格符号并转换为纯文本
    text1 = Replace(cell1.text, Chr(13), "") ' 去除换行符
    text1 = Replace(text1, Chr(7), "")      ' 去除表格符号
    text1 = StrConv(text1, vbNarrow)        ' 转换为半角字符
    text2 = Replace(cell2.text, Chr(13), "") ' 去除换行符
    text2 = Replace(text2, Chr(7), "")      ' 去除表格符号
    text2 = StrConv(text2, vbNarrow)        ' 转换为半角字符
    
    ' 提取x±s格式的内容
    If Not ExtractXS(text1, x1, s1) Then
        MsgBox "未在第一个单元格中找到有效的x±s格式内容。"
        Exit Sub
    End If
    If Not ExtractXS(text2, x2, s2) Then
        MsgBox "未在第二个单元格中找到有效的x±s格式内容。"
        Exit Sub
    End If
    
    ' 弹窗提示用户输入样本量，并验证输入是否为数字
    n1 = InputBox("请输入第一组的样本量：", "输入样本量")
    If Not IsNumeric(n1) Or n1 <= 1 Then
        MsgBox "请输入有效的样本量（大于1）。"
        Exit Sub
    End If
    n2 = InputBox("请输入第二组的样本量：", "输入样本量")
    If Not IsNumeric(n2) Or n2 <= 1 Then
        MsgBox "请输入有效的样本量（大于1）。"
        Exit Sub
    End If
    
    ' 计算t值和自由度df（使用独立样本t检验公式）
    t = CalculateT(x1, x2, s1, s2, n1, n2)
    df = n1 + n2 - 2
    
    ' 计算P值（使用更精确的双尾检验）
    pValue = CalculatePValue(t, df)
    
    ' 在批注中显示结果
    With ActiveDocument.Comments.Add(Range:=Selection.Range)
        .Range.text = _
            "【t检验结果】" & vbCrLf & _
            "第一组数据（x1±s1）：" & Format(x1, "0.00") & " ± " & Format(s1, "0.00") & vbCrLf & _
            "样本量 n1：" & n1 & vbCrLf & _
            "第二组数据（x2±s2）：" & Format(x2, "0.00") & " ± " & Format(s2, "0.00") & vbCrLf & _
            "样本量 n2：" & n2 & vbCrLf & _
            "t值 = " & Format(t, "0.0000") & vbCrLf & _
            "自由度 = " & df & vbCrLf & _
            "P值 = " & Format(pValue, "0.0000") & vbCrLf & _
            "【提示】这是假设方差齐性的独立样本t检验结果"
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "发生错误: " & Err.Description
End Sub

Function ExtractXS(text As String, ByRef x As Double, ByRef s As Double) As Boolean
    Dim regex As Object
    Dim matches As Object
    Dim parts() As String
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "(\d+\.?\d*)±(\d+\.?\d*)"
    regex.Global = True
    
    Set matches = regex.Execute(text)
    
    If matches.Count > 0 Then
        parts = Split(matches(0).value, "±")
        x = CDbl(parts(0))
        s = CDbl(parts(1))
        ExtractXS = True
    Else
        ExtractXS = False
    End If
End Function

' 修正后的t值计算（独立样本t检验，假设方差齐性）
Function CalculateT(x1 As Double, x2 As Double, s1 As Double, s2 As Double, n1 As Double, n2 As Double) As Double
    Dim pooledVar As Double
    
    ' 计算合并方差
    pooledVar = ((n1 - 1) * s1 ^ 2 + (n2 - 1) * s2 ^ 2) / (n1 + n2 - 2)
    
    ' 计算t值
    CalculateT = (x1 - x2) / Sqr(pooledVar * (1 / n1 + 1 / n2))
End Function

' 修正后的P值计算（纯VBA实现）
Function CalculatePValue(t As Double, df As Double) As Double
    ' 使用改进的纯VBA实现的双尾P值计算
    Dim p As Double
    
    ' 计算单侧概率
    p = T_Dist_Right(Abs(t), df)
    
    ' 双尾检验，乘以2
    CalculatePValue = 2 * p
    
    ' 确保P值在0和1之间
    If CalculatePValue < 0 Then CalculatePValue = 0
    If CalculatePValue > 1 Then CalculatePValue = 1
End Function

' 计算t分布的右侧概率
Function T_Dist_Right(t As Double, df As Double) As Double
    Dim x As Double, a As Double, b As Double
    
    x = df / (df + t * t)
    a = df / 2
    b = 0.5
    
    T_Dist_Right = 0.5 * BetaI(x, a, b)
End Function

' 正则化不完全Beta函数
Function BetaI(x As Double, a As Double, b As Double) As Double
    Dim bt As Double
    Dim eps As Double
    Dim a1 As Double, b1 As Double
    Dim m As Integer
    
    eps = 0.0000001
    
    ' 检查x的范围
    If x <= 0 Then
        BetaI = 0
        Exit Function
    ElseIf x >= 1 Then
        BetaI = 1
        Exit Function
    End If
    
    ' 计算前置因子
    bt = Exp(GammaLn(a + b) - GammaLn(a) - GammaLn(b) + a * Log(x) + b * Log(1 - x))
    
    ' 根据x值选择计算方法
    If x < (a + 1) / (a + b + 2) Then
        BetaI = bt * BetaCF(x, a, b) / a
    Else
        BetaI = 1 - bt * BetaCF(1 - x, b, a) / b
    End If
End Function

' 连分式展开计算BetaCF（修正版）
Function BetaCF(x As Double, a As Double, b As Double) As Double
    Dim qab As Double, qap As Double, qam As Double
    Dim c As Double, d As Double, h As Double
    Dim m As Integer
    Dim aa As Double, del As Double
    Dim maxIter As Integer
    
    qab = a + b
    qap = a + 1
    qam = a - 1
    c = 1
    d = 1 - qab * x / qap
    If Abs(d) < 1E-30 Then d = 1E-30
    d = 1 / d
    h = d
    maxIter = 100
    
    For m = 1 To maxIter
        aa = m * (b - m) * x / ((qam + 2 * m) * (a + 2 * m))
        d = 1 + aa * d
        If Abs(d) < 1E-30 Then d = 1E-30
        c = 1 + aa / c
        If Abs(c) < 1E-30 Then c = 1E-30
        d = 1 / d
        h = h * d * c
        
        aa = -(a + m) * (qab + m) * x / ((a + 2 * m) * (qap + 2 * m))
        d = 1 + aa * d
        If Abs(d) < 1E-30 Then d = 1E-30
        c = 1 + aa / c
        If Abs(c) < 1E-30 Then c = 1E-30
        d = 1 / d
        del = d * c
        h = h * del
        
        If Abs(del - 1) < 0.0000001 Then Exit For
    Next m
    
    BetaCF = h
End Function

' 对数Gamma函数（修正版）
Function GammaLn(x As Double) As Double
    Dim cof(6) As Double
    Dim stp As Double
    Dim tmp As Double, ser As Double
    Dim j As Integer
    Dim y As Double
    
    cof(1) = 76.1800917294715
    cof(2) = -86.5053203294168
    cof(3) = 24.0140982408309
    cof(4) = -1.23173957245015
    cof(5) = 1.20865097386618E-03
    cof(6) = -5.395239384953E-06
    stp = 2.506628274631
    
    y = x
    tmp = x + 5.5
    tmp = (x + 0.5) * Log(tmp) - tmp
    ser = 1.00000000019001
    
    For j = 1 To 6
        y = y + 1
        ser = ser + cof(j) / y
    Next j
    
    GammaLn = tmp + Log(stp * ser / x)
End Function
