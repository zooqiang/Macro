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

    ' ��ȡ��ǰѡ�е�����
    Set sel = Selection

    ' ���ѡ�еĵ�Ԫ������
    Select Case sel.Cells.Count
        Case 4 ' ģʽһ��ֱ����ȡ A, B, C, D
            ' ��ȡ�ĸ���Ԫ�������
            cellText1 = sel.Cells(1).Range.text
            cellText2 = sel.Cells(2).Range.text
            cellText3 = sel.Cells(3).Range.text
            cellText4 = sel.Cells(4).Range.text
            ' ȥ�������ţ�������ǣ���ת��Ϊ���
            cellText1 = StrConv(Replace(cellText1, vbCr, ""), vbNarrow)
            cellText2 = StrConv(Replace(cellText2, vbCr, ""), vbNarrow)
            cellText3 = StrConv(Replace(cellText3, vbCr, ""), vbNarrow)
            cellText4 = StrConv(Replace(cellText4, vbCr, ""), vbNarrow)
            ' ��ȡ���֣�A, B, C, D��
            a = ExtractNumber(cellText1)
            b = ExtractNumber(cellText2)
            C = ExtractNumber(cellText3)
            D = ExtractNumber(cellText4)
        Case 2 ' ģʽ������ȡ A, B, O, P������ X, Y, C, D
            ' ��ȡ������Ԫ�������
            cellText1 = sel.Cells(1).Range.text
            cellText2 = sel.Cells(2).Range.text
            ' ȥ�������ţ�������ǣ���ת��Ϊ���
            cellText1 = StrConv(Replace(cellText1, vbCr, ""), vbNarrow)
            cellText2 = StrConv(Replace(cellText2, vbCr, ""), vbNarrow)
            ' ��ȡ����ǰ�����֣�A �� B��
            a = ExtractNumberBeforeParenthesis(cellText1)
            b = ExtractNumberBeforeParenthesis(cellText2)
            ' ��ȡ���ź�����֣�O �� P��
            O = ExtractNumberAfterParenthesis(cellText1)
            p = ExtractNumberAfterParenthesis(cellText2)
            ' ���� X, Y, C, D
            x = Round((a / O) * 100, 0)
            y = Round((b / p) * 100, 0)
            C = x - a
            D = y - b
        Case Else
            MsgBox "��ѡ���������ĸ���Ԫ��", vbExclamation
            Exit Sub
    End Select

    ' ��������Ƶ��
    Dim total As Double
    total = a + b + C + D
    aExp = (a + b) * (a + C) / total
    bExp = (a + b) * (b + D) / total
    cExp = (C + D) * (a + C) / total
    dExp = (C + D) * (b + D) / total
    
    ' ����Ƿ�������Ƶ��С��5�ĵ�Ԫ��
    hasSmallValue = (aExp < 5 Or bExp < 5 Or cExp < 5 Or dExp < 5)

    ' ���㿨��ֵ��ʹ��Pearson�������鹫ʽ��
    ChiSquare = ((a - aExp) ^ 2 / aExp) + _
                ((b - bExp) ^ 2 / bExp) + _
                ((C - cExp) ^ 2 / cExp) + _
                ((D - dExp) ^ 2 / dExp)

    ' ����pֵ (���ɶ�=1)
    pValue = ExactChiSquarePValue(ChiSquare, 1)

    ' ��������ַ���������ʾ����Ƶ����
    result = "A: " & a & vbCrLf & _
             "B: " & b & vbCrLf & _
             "C: " & C & vbCrLf & _
             "D: " & D & vbCrLf & _
             "����ֵ: " & Round(ChiSquare, 4) & vbCrLf & _
             "P ֵ: " & Format(pValue, "0.0000")
    
    ' ֻ�е���������Ƶ��С��5ʱ����Ӿ���
    If hasSmallValue Then
        result = result & vbCrLf & vbCrLf & _
                 "����: ��������Ƶ��С��5�ĵ�Ԫ��(" & _
                 IIf(aExp < 5, "A ", "") & _
                 IIf(bExp < 5, "B ", "") & _
                 IIf(cExp < 5, "C ", "") & _
                 IIf(dExp < 5, "D", "") & _
                 ")������ʹ��Fisher��ȷ���飡"
    End If

    ' ɾ��������ע
    Dim cmt As comment
    For Each cmt In sel.Range.Comments
        cmt.Delete
    Next cmt
    
    ' �������ע
    sel.Range.Comments.Add sel.Range, result
End Sub

' ========== �������� ==========

' ��ȡ����
Function ExtractNumber(text As String) As Double
    Dim numStr As String
    Dim i As Integer
    Dim char As String
    ' ��ȡ���ֲ���
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        If IsNumeric(char) Or char = "." Then
            numStr = numStr & char
        End If
    Next i
    ' ת��Ϊ����
    If IsNumeric(numStr) Then
        ExtractNumber = CDbl(numStr)
    Else
        ExtractNumber = 0
    End If
End Function

' ��ȡ����ǰ������
Function ExtractNumberBeforeParenthesis(text As String) As Double
    Dim pos As Integer
    Dim numStr As String
    ' �������ŵ�λ��
    pos = InStr(text, "(")
    ' ��ȡ����ǰ������
    If pos > 1 Then
        numStr = Trim(Left(text, pos - 1))
    Else
        numStr = ""
    End If
    ' ת��Ϊ����
    If IsNumeric(numStr) Then
        ExtractNumberBeforeParenthesis = CDbl(numStr)
    Else
        ExtractNumberBeforeParenthesis = 0
    End If
End Function

' ��ȡ�����ڵ�����
Function ExtractNumberAfterParenthesis(text As String) As Double
    Dim posStart As Integer, posEnd As Integer
    Dim numStr As String
    ' �������ŵ�λ��
    posStart = InStr(text, "(")
    posEnd = InStr(text, ")")
    ' ��ȡ�����ڵ�����
    If posStart > 0 And posEnd > posStart Then
        numStr = Trim(Mid(text, posStart + 1, posEnd - posStart - 1))
    Else
        numStr = ""
    End If
    ' ȥ���ٷֺŲ�ת��Ϊ����
    numStr = Replace(numStr, "%", "")
    If IsNumeric(numStr) Then
        ExtractNumberAfterParenthesis = CDbl(numStr)
    Else
        ExtractNumberAfterParenthesis = 0
    End If
End Function

' ========== ͳ�Ƽ��㺯�� ==========

' ������ľ�ȷ�����ֲ�pֵ���㺯��
Function ExactChiSquarePValue(x As Double, df As Integer) As Double
    ' ȷ��������Ч��
    If x <= 0 Or df <= 0 Then
        ExactChiSquarePValue = 1
        Exit Function
    End If
    
    ' �Լ�С�Ŀ���ֱֵ�ӷ���1
    If x < 0.000001 Then
        ExactChiSquarePValue = 1
        Exit Function
    End If
    
    ' �����ɶ�=1�����⴦��
    If df = 1 Then
        ' ȷ��pֵ������1
        Dim sqrtX As Double
        sqrtX = Sqr(x)
        If sqrtX > 8 Then  ' ���ڷǳ����xֵ��ֱ�ӷ���0
            ExactChiSquarePValue = 0
        Else
            ExactChiSquarePValue = 2 * (1 - NormalCDF(sqrtX))
            ' ȷ��pֵ��0-1��Χ��
            If ExactChiSquarePValue > 1 Then ExactChiSquarePValue = 1
            If ExactChiSquarePValue < 0 Then ExactChiSquarePValue = 0
        End If
    Else
        ' �����������ɶȣ�ʹ�ò���ȫGamma����
        Dim gammaResult As Double
        gammaResult = IncompleteGamma(df / 2, x / 2)
        ExactChiSquarePValue = 1 - gammaResult
        ' ȷ��pֵ��0-1��Χ��
        If ExactChiSquarePValue > 1 Then ExactChiSquarePValue = 1
        If ExactChiSquarePValue < 0 Then ExactChiSquarePValue = 0
    End If
End Function

' �߾��ȱ�׼��̬CDF���㣨�����棩
Function NormalCDF(x As Double) As Double
    ' ʹ��Hart�����㷨�����ӱ߽���
    Dim z As Double, t As Double, y As Double
    
    ' ������ֵ
    If x > 8 Then
        NormalCDF = 1
        Exit Function
    End If
    
    ' ����Сֵ
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
    
    ' ȷ�������0-1��Χ��
    If y > 1 Then y = 1
    If y < 0 Then y = 0
    
    If x > 0 Then
        NormalCDF = y
    Else
        NormalCDF = 1 - y
    End If
End Function

' ����ȫGamma����ʵ�֣����ӱ߽��飩
Function IncompleteGamma(a As Double, x As Double) As Double
    ' ���������Ч��
    If x < 0 Or a <= 0 Then
        IncompleteGamma = 0
        Exit Function
    End If
    
    ' �Լ�С��xֱֵ�ӷ���0
    If x < 0.0000000001 Then
        IncompleteGamma = 0
        Exit Function
    End If
    
    ' �Լ����xֱֵ�ӷ���1
    If x > 100000 Then
        IncompleteGamma = 1
        Exit Function
    End If
    
    Dim gamser As Double, gln As Double
    Dim gamcf As Double
    
    If x < a + 1 Then
        ' ʹ�ü���չ��
        Call GammaSeries(gamser, a, x, gln)
        ' ȷ�������0-1��Χ��
        If gamser > 1 Then gamser = 1
        If gamser < 0 Then gamser = 0
        IncompleteGamma = gamser
    Else
        ' ʹ������ʽչ��
        Call GammaCF(gamcf, a, x, gln)
        ' ȷ�������0-1��Χ��
        If gamcf > 1 Then gamcf = 1
        If gamcf < 0 Then gamcf = 0
        IncompleteGamma = 1 - gamcf
    End If
End Function

' Gamma����չ��
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

' Gamma����ʽչ��
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

' ����Gamma�����Ķ���
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

