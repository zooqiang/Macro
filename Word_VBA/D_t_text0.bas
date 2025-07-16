Attribute VB_Name = "D_t_text0"
' ���������t���飨�����棩
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
    
    ' ����Ƿ�ѡ����������Ԫ��
    If Selection.Cells.Count <> 2 Then
        MsgBox "��ѡ��������Ԫ��"
        Exit Sub
    End If
    
    ' ��ȡ������Ԫ�������
    Set cell1 = Selection.Cells(1).Range
    Set cell2 = Selection.Cells(2).Range
    
    ' ȥ�������Ų�ת��Ϊ���ı�
    text1 = Replace(cell1.text, Chr(13), "") ' ȥ�����з�
    text1 = Replace(text1, Chr(7), "")      ' ȥ��������
    text1 = StrConv(text1, vbNarrow)        ' ת��Ϊ����ַ�
    text2 = Replace(cell2.text, Chr(13), "") ' ȥ�����з�
    text2 = Replace(text2, Chr(7), "")      ' ȥ��������
    text2 = StrConv(text2, vbNarrow)        ' ת��Ϊ����ַ�
    
    ' ��ȡx��s��ʽ������
    If Not ExtractXS(text1, x1, s1) Then
        MsgBox "δ�ڵ�һ����Ԫ�����ҵ���Ч��x��s��ʽ���ݡ�"
        Exit Sub
    End If
    If Not ExtractXS(text2, x2, s2) Then
        MsgBox "δ�ڵڶ�����Ԫ�����ҵ���Ч��x��s��ʽ���ݡ�"
        Exit Sub
    End If
    
    ' ������ʾ�û�����������������֤�����Ƿ�Ϊ����
    n1 = InputBox("�������һ�����������", "����������")
    If Not IsNumeric(n1) Or n1 <= 1 Then
        MsgBox "��������Ч��������������1����"
        Exit Sub
    End If
    n2 = InputBox("������ڶ������������", "����������")
    If Not IsNumeric(n2) Or n2 <= 1 Then
        MsgBox "��������Ч��������������1����"
        Exit Sub
    End If
    
    ' ����tֵ�����ɶ�df��ʹ�ö�������t���鹫ʽ��
    t = CalculateT(x1, x2, s1, s2, n1, n2)
    df = n1 + n2 - 2
    
    ' ����Pֵ��ʹ�ø���ȷ��˫β���飩
    pValue = CalculatePValue(t, df)
    
    ' ����ע����ʾ���
    With ActiveDocument.Comments.Add(Range:=Selection.Range)
        .Range.text = _
            "��t��������" & vbCrLf & _
            "��һ�����ݣ�x1��s1����" & Format(x1, "0.00") & " �� " & Format(s1, "0.00") & vbCrLf & _
            "������ n1��" & n1 & vbCrLf & _
            "�ڶ������ݣ�x2��s2����" & Format(x2, "0.00") & " �� " & Format(s2, "0.00") & vbCrLf & _
            "������ n2��" & n2 & vbCrLf & _
            "tֵ = " & Format(t, "0.0000") & vbCrLf & _
            "���ɶ� = " & df & vbCrLf & _
            "Pֵ = " & Format(pValue, "0.0000") & vbCrLf & _
            "����ʾ�����Ǽ��跽�����ԵĶ�������t������"
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "��������: " & Err.Description
End Sub

Function ExtractXS(text As String, ByRef x As Double, ByRef s As Double) As Boolean
    Dim regex As Object
    Dim matches As Object
    Dim parts() As String
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "(\d+\.?\d*)��(\d+\.?\d*)"
    regex.Global = True
    
    Set matches = regex.Execute(text)
    
    If matches.Count > 0 Then
        parts = Split(matches(0).value, "��")
        x = CDbl(parts(0))
        s = CDbl(parts(1))
        ExtractXS = True
    Else
        ExtractXS = False
    End If
End Function

' �������tֵ���㣨��������t���飬���跽�����ԣ�
Function CalculateT(x1 As Double, x2 As Double, s1 As Double, s2 As Double, n1 As Double, n2 As Double) As Double
    Dim pooledVar As Double
    
    ' ����ϲ�����
    pooledVar = ((n1 - 1) * s1 ^ 2 + (n2 - 1) * s2 ^ 2) / (n1 + n2 - 2)
    
    ' ����tֵ
    CalculateT = (x1 - x2) / Sqr(pooledVar * (1 / n1 + 1 / n2))
End Function

' �������Pֵ���㣨��VBAʵ�֣�
Function CalculatePValue(t As Double, df As Double) As Double
    ' ʹ�øĽ��Ĵ�VBAʵ�ֵ�˫βPֵ����
    Dim p As Double
    
    ' ���㵥�����
    p = T_Dist_Right(Abs(t), df)
    
    ' ˫β���飬����2
    CalculatePValue = 2 * p
    
    ' ȷ��Pֵ��0��1֮��
    If CalculatePValue < 0 Then CalculatePValue = 0
    If CalculatePValue > 1 Then CalculatePValue = 1
End Function

' ����t�ֲ����Ҳ����
Function T_Dist_Right(t As Double, df As Double) As Double
    Dim x As Double, a As Double, b As Double
    
    x = df / (df + t * t)
    a = df / 2
    b = 0.5
    
    T_Dist_Right = 0.5 * BetaI(x, a, b)
End Function

' ���򻯲���ȫBeta����
Function BetaI(x As Double, a As Double, b As Double) As Double
    Dim bt As Double
    Dim eps As Double
    Dim a1 As Double, b1 As Double
    Dim m As Integer
    
    eps = 0.0000001
    
    ' ���x�ķ�Χ
    If x <= 0 Then
        BetaI = 0
        Exit Function
    ElseIf x >= 1 Then
        BetaI = 1
        Exit Function
    End If
    
    ' ����ǰ������
    bt = Exp(GammaLn(a + b) - GammaLn(a) - GammaLn(b) + a * Log(x) + b * Log(1 - x))
    
    ' ����xֵѡ����㷽��
    If x < (a + 1) / (a + b + 2) Then
        BetaI = bt * BetaCF(x, a, b) / a
    Else
        BetaI = 1 - bt * BetaCF(1 - x, b, a) / b
    End If
End Function

' ����ʽչ������BetaCF�������棩
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

' ����Gamma�����������棩
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
