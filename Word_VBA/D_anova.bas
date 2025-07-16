Attribute VB_Name = "D_anova"
' 方差分析，上下或者左右连续的三个单元格
Sub D_anova()
    On Error GoTo ErrorHandler
    Dim cell1 As Range, cell2 As Range, cell3 As Range
    Dim text1 As String, text2 As String, text3 As String
    Dim x1 As Double, s1 As Double
    Dim x2 As Double, s2 As Double
    Dim x3 As Double, s3 As Double
    Dim a As Double, b As Double, c As Double
    Dim f As Double, df1 As Double, df2 As Double
    Dim pValue As Double
    Dim resultComment As String
    
    ' 检查是否选中了三个单元格
    If selection.cells.Count <> 3 Then
        MsgBox "请选中三个连续的单元格。", vbExclamation, "输入错误"
        Exit Sub
    End If
    
    ' 获取三个单元格的内容
    Set cell1 = selection.cells(1).Range
    Set cell2 = selection.cells(2).Range
    Set cell3 = selection.cells(3).Range
    
    ' 去除表格符号并转换为纯文本
    text1 = CleanText(cell1.text)
    text2 = CleanText(cell2.text)
    text3 = CleanText(cell3.text)
    
    ' 提取x±s格式的内容
    If Not ExtractXS(text1, x1, s1) Then
        MsgBox "第一个单元格中未找到有效的x±s格式内容。", vbExclamation, "数据格式错误"
        Exit Sub
    End If
    If Not ExtractXS(text2, x2, s2) Then
        MsgBox "第二个单元格中未找到有效的x±s格式内容。", vbExclamation, "数据格式错误"
        Exit Sub
    End If
    If Not ExtractXS(text3, x3, s3) Then
        MsgBox "第三个单元格中未找到有效的x±s格式内容。", vbExclamation, "数据格式错误"
        Exit Sub
    End If
    
    ' 验证标准差是否为正数
    If s1 <= 0 Or s2 <= 0 Or s3 <= 0 Then
        MsgBox "标准差必须为正数。", vbExclamation, "数据错误"
        Exit Sub
    End If
    
    ' 弹窗提示用户输入样本量，并验证输入
    a = GetSampleSize("单元格A")
    If a <= 0 Then Exit Sub
    b = GetSampleSize("单元格B")
    If b <= 0 Then Exit Sub
    c = GetSampleSize("单元格C")
    If c <= 0 Then Exit Sub
    
    ' 计算F值和自由度
    If Not CalculateANOVA(x1, x2, x3, s1, s2, s3, a, b, c, f, df1, df2, pValue) Then
        MsgBox "方差分析计算失败，请检查输入数据。", vbExclamation, "计算错误"
        Exit Sub
    End If
    
    ' 准备结果注释
    resultComment = BuildResultComment(x1, s1, x2, s2, x3, s3, a, b, c, f, df1, df2, pValue)
    
    ' 添加结果注释
    AddSimpleComment resultComment
    
    Exit Sub
    
ErrorHandler:
    MsgBox "发生错误: " & Err.Description & vbCrLf & "错误代码: " & Err.Number, vbCritical, "系统错误"
End Sub

' ======================
' 辅助函数
' ======================

' 清理文本内容
Function CleanText(inputText As String) As String
    CleanText = StrConv(Replace(Replace(inputText, Chr(13), ""), Chr(7), ""), vbNarrow)
End Function

' 获取样本量输入
Function GetSampleSize(cellName As String) As Double
    Dim inputValue As String
    Dim result As Double
    
    inputValue = InputBox("请输入" & cellName & "的样本量：", "输入样本量")
    
    If inputValue = "" Then
        GetSampleSize = -1
        Exit Function
    End If
    
    If IsNumeric(inputValue) Then
        result = CDbl(inputValue)
        If result > 0 Then
            GetSampleSize = result
        Else
            MsgBox "样本量必须为正数。", vbExclamation, "输入错误"
            GetSampleSize = -1
        End If
    Else
        MsgBox "请输入有效的数字。", vbExclamation, "输入错误"
        GetSampleSize = -1
    End If
End Function

' 提取x±s格式数据
Function ExtractXS(text As String, ByRef x As Double, ByRef s As Double) As Boolean
    Dim regEx As Object
    Dim matches As Object
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "(-?\d+\.?\d*)\s*±\s*(-?\d+\.?\d*)"
    regEx.Global = True
    
    Set matches = regEx.Execute(text)
    
    If matches.Count > 0 Then
        x = CDbl(matches(0).SubMatches(0))
        s = CDbl(matches(0).SubMatches(1))
        ExtractXS = True
    Else
        ExtractXS = False
    End If
End Function

' 计算ANOVA结果
Function CalculateANOVA(x1 As Double, x2 As Double, x3 As Double, _
                        s1 As Double, s2 As Double, s3 As Double, _
                        a As Double, b As Double, c As Double, _
                        ByRef f As Double, ByRef df1 As Double, ByRef df2 As Double, _
                        ByRef pValue As Double) As Boolean
    On Error GoTo ErrorHandler
    
    Dim grandMean As Double, totalN As Double
    Dim SSB As Double, SSW As Double
    Dim MSB As Double, MSW As Double
    
    ' 计算总样本量和总均值
    totalN = a + b + c
    If totalN <= 3 Then
        MsgBox "总样本量太小，无法进行计算。", vbExclamation, "数据错误"
        CalculateANOVA = False
        Exit Function
    End If
    
    grandMean = (x1 * a + x2 * b + x3 * c) / totalN
    
    ' 计算组间平方和 (SSB)
    SSB = a * (x1 - grandMean) ^ 2 + b * (x2 - grandMean) ^ 2 + c * (x3 - grandMean) ^ 2
    
    ' 计算组内平方和 (SSW)
    SSW = (a - 1) * s1 ^ 2 + (b - 1) * s2 ^ 2 + (c - 1) * s3 ^ 2
    
    ' 计算自由度
    df1 = 2 ' 组间自由度 (k-1)
    df2 = totalN - 3 ' 组内自由度 (N-k)
    
    ' 计算均方
    MSB = SSB / df1
    MSW = SSW / df2
    
    ' 计算F值
    If MSW <= 0 Then
        MsgBox "组内变异为零，无法计算F值。", vbExclamation, "计算错误"
        CalculateANOVA = False
        Exit Function
    End If
    
    f = MSB / MSW
    
    ' 计算P值
    pValue = GetPValue(f, df1, df2)
    If pValue < 0 Then
        CalculateANOVA = False
        Exit Function
    End If
    
    CalculateANOVA = True
    Exit Function
    
ErrorHandler:
    CalculateANOVA = False
    Debug.Print "CalculateANOVA Error: " & Err.Description
End Function

' 获取P值（优先使用Excel函数）
Function GetPValue(f As Double, df1 As Double, df2 As Double) As Double
    Dim excelP As Double
    
    ' 首先尝试使用Excel函数
    excelP = GetExcelPValue(f, df1, df2)
    If excelP >= 0 Then
        GetPValue = excelP
        Exit Function
    End If
    
    ' 如果Excel不可用，使用VBA算法
    GetPValue = CalculatePValueWithVBA(f, df1, df2)
    
    ' 验证结果
    If GetPValue < 0 Or GetPValue > 1 Then
        MsgBox "P值计算结果超出范围，请检查输入数据。", vbExclamation, "计算错误"
        GetPValue = -1
    End If
End Function

' 使用Excel计算P值
Function GetExcelPValue(f As Double, df1 As Double, df2 As Double) As Double
    On Error Resume Next
    Dim excelApp As Object
    
    Set excelApp = CreateObject("Excel.Application")
    If Err.Number <> 0 Then
        GetExcelPValue = -1
        Exit Function
    End If
    
    ' 尝试使用新版函数
    GetExcelPValue = excelApp.WorksheetFunction.FDist_RT(f, df1, df2)
    If Err.Number = 0 Then
        excelApp.Quit
        Set excelApp = Nothing
        Exit Function
    End If
    
    ' 尝试使用旧版函数
    Err.Clear
    GetExcelPValue = excelApp.WorksheetFunction.FDist(f, df1, df2)
    If Err.Number = 0 Then
        excelApp.Quit
        Set excelApp = Nothing
        Exit Function
    End If
    
    ' 两种方法都失败
    excelApp.Quit
    Set excelApp = Nothing
    GetExcelPValue = -1
End Function

' 使用VBA算法计算P值
Function CalculatePValueWithVBA(f As Double, df1 As Double, df2 As Double) As Double
    On Error GoTo ErrorHandler
    
    ' 检查输入参数
    If f <= 0 Or df1 <= 0 Or df2 <= 0 Then
        CalculatePValueWithVBA = -1
        Exit Function
    End If
    
    Dim x As Double
    x = df2 / (df2 + df1 * f)
    
    ' 边界检查
    If x <= 0 Then
        CalculatePValueWithVBA = 1
        Exit Function
    ElseIf x >= 1 Then
        CalculatePValueWithVBA = 0
        Exit Function
    End If
    
    Dim a As Double, b As Double
    a = df2 / 2
    b = df1 / 2
    
    ' 计算不完全Beta函数
    Dim betaResult As Double
    betaResult = IncompleteBeta(x, a, b)
    
    ' 计算P值（右尾概率）
    CalculatePValueWithVBA = 1 - betaResult
    
    ' 验证结果
    If CalculatePValueWithVBA < 0 Then CalculatePValueWithVBA = 0
    If CalculatePValueWithVBA > 1 Then CalculatePValueWithVBA = 1
    
    Exit Function
    
ErrorHandler:
    Debug.Print "CalculatePValueWithVBA Error: " & Err.Description
    CalculatePValueWithVBA = -1
End Function

' 不完全Beta函数计算
Function IncompleteBeta(x As Double, a As Double, b As Double) As Double
    On Error GoTo ErrorHandler
    
    ' 参数检查
    If x <= 0 Then
        IncompleteBeta = 0
        Exit Function
    ElseIf x >= 1 Then
        IncompleteBeta = 1
        Exit Function
    End If
    
    If a <= 0 Or b <= 0 Then
        IncompleteBeta = 0
        Exit Function
    End If
    
    ' 使用连分数展开法
    Dim eps As Double: eps = 1E-16
    Dim maxIter As Integer: maxIter = 200
    Dim m, m2, aa, c, d, h As Double
    Dim i As Integer
    
    ' 初始化
    m = 1
    c = 1
    d = 1 - (a + b) * x / (a + 1)
    If Abs(d) < eps Then d = eps
    d = 1 / d
    h = d
    
    ' 连分数展开
    For i = 1 To maxIter
        m2 = 2 * i
        
        ' 第一部分
        aa = i * (b - i) * x / ((a + m2 - 1) * (a + m2))
        d = 1 + aa * d
        If Abs(d) < eps Then d = eps
        c = 1 + aa / c
        If Abs(c) < eps Then c = eps
        d = 1 / d
        h = h * d * c
        
        ' 第二部分
        aa = -(a + i) * (a + b + i) * x / ((a + m2) * (a + m2 + 1))
        d = 1 + aa * d
        If Abs(d) < eps Then d = eps
        c = 1 + aa / c
        If Abs(c) < eps Then c = eps
        d = 1 / d
        h = h * d * c
        
        ' 检查收敛
        If Abs(d * c - 1) < eps Then Exit For
    Next i
    
    ' 计算Beta函数值
    Dim betaVal As Double
    betaVal = Beta(a, b)
    
    ' 计算最终结果
    If betaVal > 0 Then
        Dim term As Double
        term = x ^ a * (1 - x) ^ b / (a * betaVal)
        IncompleteBeta = h * term
    Else
        IncompleteBeta = 0
    End If
    
    ' 确保结果在0-1范围内
    If IncompleteBeta < 0 Then IncompleteBeta = 0
    If IncompleteBeta > 1 Then IncompleteBeta = 1
    
    Exit Function
    
ErrorHandler:
    Debug.Print "IncompleteBeta Error: " & Err.Description
    IncompleteBeta = 0
End Function

' Beta函数计算
Function Beta(a As Double, b As Double) As Double
    On Error GoTo ErrorHandler
    
    ' 参数检查
    If a <= 0 Or b <= 0 Then
        Beta = 0
        Exit Function
    End If
    
    ' 使用对数Gamma函数计算
    Dim lgA As Double, lgB As Double, lgAB As Double
    
    lgA = LogGamma(a)
    lgB = LogGamma(b)
    lgAB = LogGamma(a + b)
    
    ' 计算结果
    Beta = Exp(lgA + lgB - lgAB)
    
    ' 确保结果有效
    If Beta <= 0 Then Beta = 0
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Beta Error: " & Err.Description
    Beta = 0
End Function

' 对数Gamma函数计算
Function LogGamma(x As Double) As Double
    On Error GoTo ErrorHandler
    
    ' 参数检查
    If x <= 0 Then
        LogGamma = 1E+308
        Exit Function
    End If
    
    ' Lanczos近似系数
    Dim coef(6) As Double
    coef(0) = 76.1800917294715
    coef(1) = -86.5053203294168
    coef(2) = 24.0140982408309
    coef(3) = -1.23173957245015
    coef(4) = 1.20865097386618E-03
    coef(5) = -5.395239384953E-06
    
    Dim tmp As Double, ser As Double
    Dim i As Integer
    
    tmp = x + 5.5
    tmp = (x + 0.5) * Log(tmp) - tmp
    ser = 1.00000000019001
    
    For i = 0 To 5
        x = x + 1
        ser = ser + coef(i) / x
    Next i
    
    LogGamma = tmp + Log(2.506628274631 * ser / x)
    
    Exit Function
    
ErrorHandler:
    Debug.Print "LogGamma Error: " & Err.Description
    LogGamma = 0
End Function

' 构建结果注释文本
Function BuildResultComment(x1 As Double, s1 As Double, _
                           x2 As Double, s2 As Double, _
                           x3 As Double, s3 As Double, _
                           a As Double, b As Double, c As Double, _
                           f As Double, df1 As Double, df2 As Double, _
                           pValue As Double) As String
    Dim sig As String
    
    ' 判断显著性
    If pValue < 0.001 Then
        sig = "***"
    ElseIf pValue < 0.01 Then
        sig = "**"
    ElseIf pValue < 0.05 Then
        sig = "*"
    Else
        sig = "不显著"
    End If
    
    BuildResultComment = "方差分析结果:" & vbCrLf & _
        "------------------------" & vbCrLf & _
        "组1: " & Format(x1, "0.00") & " ± " & Format(s1, "0.00") & " (n=" & a & ")" & vbCrLf & _
        "组2: " & Format(x2, "0.00") & " ± " & Format(s2, "0.00") & " (n=" & b & ")" & vbCrLf & _
        "组3: " & Format(x3, "0.00") & " ± " & Format(s3, "0.00") & " (n=" & c & ")" & vbCrLf & _
        "------------------------" & vbCrLf & _
        "F(" & df1 & "," & df2 & ") = " & Format(f, "0.000") & vbCrLf & _
        "P值 = " & Format(pValue, "0.0000") & " "
End Function

' 添加分析结果注释
Sub AddSimpleComment(commentText As String)
    ' 删除现有注释
    Dim cmt As comment
    For Each cmt In selection.Comments
        cmt.Delete
    Next
    
    ' 添加新注释
    selection.Comments.add Range:=selection.Range, text:=commentText
End Sub
