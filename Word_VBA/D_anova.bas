Attribute VB_Name = "D_anova"
' ������������»�������������������Ԫ��
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
    
    ' ����Ƿ�ѡ����������Ԫ��
    If selection.cells.Count <> 3 Then
        MsgBox "��ѡ�����������ĵ�Ԫ��", vbExclamation, "�������"
        Exit Sub
    End If
    
    ' ��ȡ������Ԫ�������
    Set cell1 = selection.cells(1).Range
    Set cell2 = selection.cells(2).Range
    Set cell3 = selection.cells(3).Range
    
    ' ȥ�������Ų�ת��Ϊ���ı�
    text1 = CleanText(cell1.text)
    text2 = CleanText(cell2.text)
    text3 = CleanText(cell3.text)
    
    ' ��ȡx��s��ʽ������
    If Not ExtractXS(text1, x1, s1) Then
        MsgBox "��һ����Ԫ����δ�ҵ���Ч��x��s��ʽ���ݡ�", vbExclamation, "���ݸ�ʽ����"
        Exit Sub
    End If
    If Not ExtractXS(text2, x2, s2) Then
        MsgBox "�ڶ�����Ԫ����δ�ҵ���Ч��x��s��ʽ���ݡ�", vbExclamation, "���ݸ�ʽ����"
        Exit Sub
    End If
    If Not ExtractXS(text3, x3, s3) Then
        MsgBox "��������Ԫ����δ�ҵ���Ч��x��s��ʽ���ݡ�", vbExclamation, "���ݸ�ʽ����"
        Exit Sub
    End If
    
    ' ��֤��׼���Ƿ�Ϊ����
    If s1 <= 0 Or s2 <= 0 Or s3 <= 0 Then
        MsgBox "��׼�����Ϊ������", vbExclamation, "���ݴ���"
        Exit Sub
    End If
    
    ' ������ʾ�û�����������������֤����
    a = GetSampleSize("��Ԫ��A")
    If a <= 0 Then Exit Sub
    b = GetSampleSize("��Ԫ��B")
    If b <= 0 Then Exit Sub
    c = GetSampleSize("��Ԫ��C")
    If c <= 0 Then Exit Sub
    
    ' ����Fֵ�����ɶ�
    If Not CalculateANOVA(x1, x2, x3, s1, s2, s3, a, b, c, f, df1, df2, pValue) Then
        MsgBox "�����������ʧ�ܣ������������ݡ�", vbExclamation, "�������"
        Exit Sub
    End If
    
    ' ׼�����ע��
    resultComment = BuildResultComment(x1, s1, x2, s2, x3, s3, a, b, c, f, df1, df2, pValue)
    
    ' ��ӽ��ע��
    AddSimpleComment resultComment
    
    Exit Sub
    
ErrorHandler:
    MsgBox "��������: " & Err.Description & vbCrLf & "�������: " & Err.Number, vbCritical, "ϵͳ����"
End Sub

' ======================
' ��������
' ======================

' �����ı�����
Function CleanText(inputText As String) As String
    CleanText = StrConv(Replace(Replace(inputText, Chr(13), ""), Chr(7), ""), vbNarrow)
End Function

' ��ȡ����������
Function GetSampleSize(cellName As String) As Double
    Dim inputValue As String
    Dim result As Double
    
    inputValue = InputBox("������" & cellName & "����������", "����������")
    
    If inputValue = "" Then
        GetSampleSize = -1
        Exit Function
    End If
    
    If IsNumeric(inputValue) Then
        result = CDbl(inputValue)
        If result > 0 Then
            GetSampleSize = result
        Else
            MsgBox "����������Ϊ������", vbExclamation, "�������"
            GetSampleSize = -1
        End If
    Else
        MsgBox "��������Ч�����֡�", vbExclamation, "�������"
        GetSampleSize = -1
    End If
End Function

' ��ȡx��s��ʽ����
Function ExtractXS(text As String, ByRef x As Double, ByRef s As Double) As Boolean
    Dim regEx As Object
    Dim matches As Object
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "(-?\d+\.?\d*)\s*��\s*(-?\d+\.?\d*)"
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

' ����ANOVA���
Function CalculateANOVA(x1 As Double, x2 As Double, x3 As Double, _
                        s1 As Double, s2 As Double, s3 As Double, _
                        a As Double, b As Double, c As Double, _
                        ByRef f As Double, ByRef df1 As Double, ByRef df2 As Double, _
                        ByRef pValue As Double) As Boolean
    On Error GoTo ErrorHandler
    
    Dim grandMean As Double, totalN As Double
    Dim SSB As Double, SSW As Double
    Dim MSB As Double, MSW As Double
    
    ' ���������������ܾ�ֵ
    totalN = a + b + c
    If totalN <= 3 Then
        MsgBox "��������̫С���޷����м��㡣", vbExclamation, "���ݴ���"
        CalculateANOVA = False
        Exit Function
    End If
    
    grandMean = (x1 * a + x2 * b + x3 * c) / totalN
    
    ' �������ƽ���� (SSB)
    SSB = a * (x1 - grandMean) ^ 2 + b * (x2 - grandMean) ^ 2 + c * (x3 - grandMean) ^ 2
    
    ' ��������ƽ���� (SSW)
    SSW = (a - 1) * s1 ^ 2 + (b - 1) * s2 ^ 2 + (c - 1) * s3 ^ 2
    
    ' �������ɶ�
    df1 = 2 ' ������ɶ� (k-1)
    df2 = totalN - 3 ' �������ɶ� (N-k)
    
    ' �������
    MSB = SSB / df1
    MSW = SSW / df2
    
    ' ����Fֵ
    If MSW <= 0 Then
        MsgBox "���ڱ���Ϊ�㣬�޷�����Fֵ��", vbExclamation, "�������"
        CalculateANOVA = False
        Exit Function
    End If
    
    f = MSB / MSW
    
    ' ����Pֵ
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

' ��ȡPֵ������ʹ��Excel������
Function GetPValue(f As Double, df1 As Double, df2 As Double) As Double
    Dim excelP As Double
    
    ' ���ȳ���ʹ��Excel����
    excelP = GetExcelPValue(f, df1, df2)
    If excelP >= 0 Then
        GetPValue = excelP
        Exit Function
    End If
    
    ' ���Excel�����ã�ʹ��VBA�㷨
    GetPValue = CalculatePValueWithVBA(f, df1, df2)
    
    ' ��֤���
    If GetPValue < 0 Or GetPValue > 1 Then
        MsgBox "Pֵ������������Χ�������������ݡ�", vbExclamation, "�������"
        GetPValue = -1
    End If
End Function

' ʹ��Excel����Pֵ
Function GetExcelPValue(f As Double, df1 As Double, df2 As Double) As Double
    On Error Resume Next
    Dim excelApp As Object
    
    Set excelApp = CreateObject("Excel.Application")
    If Err.Number <> 0 Then
        GetExcelPValue = -1
        Exit Function
    End If
    
    ' ����ʹ���°溯��
    GetExcelPValue = excelApp.WorksheetFunction.FDist_RT(f, df1, df2)
    If Err.Number = 0 Then
        excelApp.Quit
        Set excelApp = Nothing
        Exit Function
    End If
    
    ' ����ʹ�þɰ溯��
    Err.Clear
    GetExcelPValue = excelApp.WorksheetFunction.FDist(f, df1, df2)
    If Err.Number = 0 Then
        excelApp.Quit
        Set excelApp = Nothing
        Exit Function
    End If
    
    ' ���ַ�����ʧ��
    excelApp.Quit
    Set excelApp = Nothing
    GetExcelPValue = -1
End Function

' ʹ��VBA�㷨����Pֵ
Function CalculatePValueWithVBA(f As Double, df1 As Double, df2 As Double) As Double
    On Error GoTo ErrorHandler
    
    ' ����������
    If f <= 0 Or df1 <= 0 Or df2 <= 0 Then
        CalculatePValueWithVBA = -1
        Exit Function
    End If
    
    Dim x As Double
    x = df2 / (df2 + df1 * f)
    
    ' �߽���
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
    
    ' ���㲻��ȫBeta����
    Dim betaResult As Double
    betaResult = IncompleteBeta(x, a, b)
    
    ' ����Pֵ����β���ʣ�
    CalculatePValueWithVBA = 1 - betaResult
    
    ' ��֤���
    If CalculatePValueWithVBA < 0 Then CalculatePValueWithVBA = 0
    If CalculatePValueWithVBA > 1 Then CalculatePValueWithVBA = 1
    
    Exit Function
    
ErrorHandler:
    Debug.Print "CalculatePValueWithVBA Error: " & Err.Description
    CalculatePValueWithVBA = -1
End Function

' ����ȫBeta��������
Function IncompleteBeta(x As Double, a As Double, b As Double) As Double
    On Error GoTo ErrorHandler
    
    ' �������
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
    
    ' ʹ��������չ����
    Dim eps As Double: eps = 1E-16
    Dim maxIter As Integer: maxIter = 200
    Dim m, m2, aa, c, d, h As Double
    Dim i As Integer
    
    ' ��ʼ��
    m = 1
    c = 1
    d = 1 - (a + b) * x / (a + 1)
    If Abs(d) < eps Then d = eps
    d = 1 / d
    h = d
    
    ' ������չ��
    For i = 1 To maxIter
        m2 = 2 * i
        
        ' ��һ����
        aa = i * (b - i) * x / ((a + m2 - 1) * (a + m2))
        d = 1 + aa * d
        If Abs(d) < eps Then d = eps
        c = 1 + aa / c
        If Abs(c) < eps Then c = eps
        d = 1 / d
        h = h * d * c
        
        ' �ڶ�����
        aa = -(a + i) * (a + b + i) * x / ((a + m2) * (a + m2 + 1))
        d = 1 + aa * d
        If Abs(d) < eps Then d = eps
        c = 1 + aa / c
        If Abs(c) < eps Then c = eps
        d = 1 / d
        h = h * d * c
        
        ' �������
        If Abs(d * c - 1) < eps Then Exit For
    Next i
    
    ' ����Beta����ֵ
    Dim betaVal As Double
    betaVal = Beta(a, b)
    
    ' �������ս��
    If betaVal > 0 Then
        Dim term As Double
        term = x ^ a * (1 - x) ^ b / (a * betaVal)
        IncompleteBeta = h * term
    Else
        IncompleteBeta = 0
    End If
    
    ' ȷ�������0-1��Χ��
    If IncompleteBeta < 0 Then IncompleteBeta = 0
    If IncompleteBeta > 1 Then IncompleteBeta = 1
    
    Exit Function
    
ErrorHandler:
    Debug.Print "IncompleteBeta Error: " & Err.Description
    IncompleteBeta = 0
End Function

' Beta��������
Function Beta(a As Double, b As Double) As Double
    On Error GoTo ErrorHandler
    
    ' �������
    If a <= 0 Or b <= 0 Then
        Beta = 0
        Exit Function
    End If
    
    ' ʹ�ö���Gamma��������
    Dim lgA As Double, lgB As Double, lgAB As Double
    
    lgA = LogGamma(a)
    lgB = LogGamma(b)
    lgAB = LogGamma(a + b)
    
    ' ������
    Beta = Exp(lgA + lgB - lgAB)
    
    ' ȷ�������Ч
    If Beta <= 0 Then Beta = 0
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Beta Error: " & Err.Description
    Beta = 0
End Function

' ����Gamma��������
Function LogGamma(x As Double) As Double
    On Error GoTo ErrorHandler
    
    ' �������
    If x <= 0 Then
        LogGamma = 1E+308
        Exit Function
    End If
    
    ' Lanczos����ϵ��
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

' �������ע���ı�
Function BuildResultComment(x1 As Double, s1 As Double, _
                           x2 As Double, s2 As Double, _
                           x3 As Double, s3 As Double, _
                           a As Double, b As Double, c As Double, _
                           f As Double, df1 As Double, df2 As Double, _
                           pValue As Double) As String
    Dim sig As String
    
    ' �ж�������
    If pValue < 0.001 Then
        sig = "***"
    ElseIf pValue < 0.01 Then
        sig = "**"
    ElseIf pValue < 0.05 Then
        sig = "*"
    Else
        sig = "������"
    End If
    
    BuildResultComment = "����������:" & vbCrLf & _
        "------------------------" & vbCrLf & _
        "��1: " & Format(x1, "0.00") & " �� " & Format(s1, "0.00") & " (n=" & a & ")" & vbCrLf & _
        "��2: " & Format(x2, "0.00") & " �� " & Format(s2, "0.00") & " (n=" & b & ")" & vbCrLf & _
        "��3: " & Format(x3, "0.00") & " �� " & Format(s3, "0.00") & " (n=" & c & ")" & vbCrLf & _
        "------------------------" & vbCrLf & _
        "F(" & df1 & "," & df2 & ") = " & Format(f, "0.000") & vbCrLf & _
        "Pֵ = " & Format(pValue, "0.0000") & " "
End Function

' ��ӷ������ע��
Sub AddSimpleComment(commentText As String)
    ' ɾ������ע��
    Dim cmt As comment
    For Each cmt In selection.Comments
        cmt.Delete
    Next
    
    ' �����ע��
    selection.Comments.add Range:=selection.Range, text:=commentText
End Sub
