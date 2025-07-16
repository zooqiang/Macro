Attribute VB_Name = "D_ChiSquareTest_2x3"
Sub ChiSquareTest_2x3()
    On Error GoTo ErrorHandler

    ' �������������ֲ��䣩
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

    ' ��ʼ��
    bErrorOccurred = False
    result = "����������(2��3):" & vbCrLf & vbCrLf

    Set oSelection = Application.selection

    If oSelection Is Nothing Or oSelection.Cells.Count = 0 Then
        result = result & "����: ����ѡ�б���еĵ�Ԫ��" & vbCrLf
        GoTo FinalOutput
    End If

    cellCount = oSelection.Cells.Count
    rowCount = oSelection.Rows.Count
    colCount = oSelection.Columns.Count

    ' ģʽһ��ѡ��6����Ԫ��2��3��ʽ�������ֲ��䣩
    If cellCount = 6 Then
        ReDim observed(0 To 5)
        For i = 1 To 6
            On Error Resume Next
            cellText = oSelection.Cells(i).Range.text
            cellText = CleanCellText(cellText)
            observed(i - 1) = ExtractNumber(cellText)
            If Err.Number <> 0 Then
                result = result & "����: ��Ԫ��" & i & "������ȡʧ�ܡ�" & vbCrLf
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
                result = result & "����: ��ѡ��2��3��3��2��6��1��1��6�ĵ�Ԫ������" & vbCrLf
                bErrorOccurred = True
            End If
        End If

    ' ģʽ����ѡ��3����Ԫ��n(N%)��ʽ��������������
    ElseIf cellCount = 3 Then
        ReDim observed(0 To 5)
        For i = 1 To 3
            On Error Resume Next
            cellText = oSelection.Cells(i).Range.text
            cellText = CleanCellText(cellText)   ' �����ı�
            
            ' ��ȡ����ǰ�ĳɹ�����ǿ��У�飩
            success = ExtractNumberBeforeParenthesis(cellText)
            ' ��ȡ�����������������ڵİٷֱȣ�
            totalNum = ExtractTotalFromPercentage(cellText, success)
            ' ����ʧ��������2���ݣ�
            fail = totalNum - success

            If Err.Number <> 0 Then
                result = result & "����: ��Ԫ��" & i & "������ȡʧ�ܣ�" & Err.Description & vbCrLf
                bErrorOccurred = True
                Exit For
            End If
            On Error GoTo ErrorHandler

            observed((i - 1) * 2) = success     ' ��1����
            observed((i - 1) * 2 + 1) = fail     ' ��2����
        Next i

        If Not bErrorOccurred Then
            For i = 0 To 2
                table(0, i) = observed(i * 2)
                table(1, i) = observed(i * 2 + 1)
            Next i
        End If
    Else
        result = result & "����: ��ѡ��3����n(N%)��ʽ����6����2��3��ʽ����Ԫ��" & vbCrLf
        bErrorOccurred = True
    End If

    ' ����У�鼰���㣨���ֲ��䣩
    If Not bErrorOccurred Then
        For i = 0 To 1
            For j = 0 To 2
                If table(i, j) < 0 Or Not IsNumeric(table(i, j)) Then
                    result = result & "����: ������Ч����Ԫ��(" & i & "," & j & ")������Ч���ݡ�" & vbCrLf
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
                        result = result & "����: ����ֵ(" & i & "," & j & ")=" & Format(expected, "0.00") & " ��С������������ܲ����á�" & vbCrLf
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

            result = result & "�۲�����:" & vbCrLf
            result = result & "��1: " & table(0, 0) & " | " & table(0, 1) & " | " & table(0, 2) & vbCrLf
            result = result & "��2: " & table(1, 0) & " | " & table(1, 1) & " | " & table(1, 2) & vbCrLf & vbCrLf
            result = result & "����ֵ = " & Format(chiSquare, "0.0000") & vbCrLf
            result = result & "Pֵ = " & Format(pValue, "0.0000") & vbCrLf
            result = result & "���ɶ� = 2"
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
    result = result & "����ʱ����: " & Err.Description & " (�����: " & Err.Number & ")" & vbCrLf
    Resume FinalOutput
End Sub

' ========== ���������������� ==========

' ����Ԫ���ı�����ǿ�棩
Function CleanCellText(text As String) As String
    text = Replace(text, Chr(13), "")         ' �Ƴ����з�
    text = Replace(text, Chr(7), "")          ' �Ƴ������ַ�
    text = Replace(text, ChrW(12288), " ")    ' ȫ�ǿո�ת���
    text = Replace(text, "��", "(")           ' ȫ������ת���
    text = Replace(text, "��", ")")
    text = Replace(text, " ", "")             ' �Ƴ����пո񣨹ؼ�����������ո������ȡ��
    CleanCellText = text
End Function

' ��ȡ����ǰ�����֣�ǿ��У�飩
Function ExtractNumberBeforeParenthesis(text As String) As Double
    Dim pos As Integer
    ' ���������ţ���ǻ�ȫ�ǣ�
    pos = InStr(text, "(")
    If pos = 0 Then pos = InStr(text, "��")
    
    If pos = 0 Then
        Err.Raise 1001, , "δ�ҵ������ţ������ʽ��ӦΪn(�ٷֱ�)��"
    End If
    
    ' ��ȡ����ǰ�����ݲ�ת��Ϊ����
    Dim numText As String
    numText = Left(text, pos - 1)
    If Not IsNumeric(numText) Then
        Err.Raise 1002, , "����ǰ�����ݲ�����Ч���֣�" & numText
    End If
    
    ExtractNumberBeforeParenthesis = CDbl(numText)
End Function

' ����������ȡ������������������ʹ��������ȡ���֣�
Function ExtractTotalFromPercentage(text As String, success As Double) As Double
    Dim pos1 As Integer, pos2 As Integer
    Dim percentText As String
    Dim percent As Double
    Dim regEx As Object, matches As Object
    
    ' ��ʼ���������������ȡ���֣�
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "\d+(\.\d+)?"  ' ƥ��������С������19.10��38.3��
    regEx.Global = False
    
    ' ��������λ��
    pos1 = InStr(text, "(")
    If pos1 = 0 Then pos1 = InStr(text, "��")
    pos2 = InStr(text, ")")
    If pos2 = 0 Then pos2 = InStr(text, "��")
    
    ' У��������Ч��
    If pos1 = 0 Or pos2 = 0 Or pos1 >= pos2 Then
        Err.Raise 1003, , "δ�ҵ���Ч���Ŷԣ������ʽ"
    End If
    
    ' ��ȡ�����ڵ�����
    percentText = Mid(text, pos1 + 1, pos2 - pos1 - 1)
    
    ' ��������ȡ�����ڵ����֣����������ַ���
    Set matches = regEx.Execute(percentText)
    If matches.Count = 0 Then
        Err.Raise 1004, , "������δ�ҵ���Ч���֣�ӦΪ�ٷֱȣ���19.10��"
    End If
    percentText = matches(0).value  ' ��ȡ��һ��ƥ�������
    
    ' ת��Ϊ�ٷֱ���ֵ
    If Not IsNumeric(percentText) Then
        Err.Raise 1005, , "��ȡ�İٷֱȲ�����Ч���֣�" & percentText
    End If
    percent = CDbl(percentText)
    
    ' ����ٷֱȸ�ʽ����19.10 �� 0.1910��
    If percent > 1 Then
        percent = percent / 100
    End If
    
    ' У��ٷֱ���Ч��
    If percent <= 0 Then
        Err.Raise 1006, , "�ٷֱȱ������0����ǰΪ��" & percentText & "��"
    End If
    
    ' ��������������������Ϊ������
    ExtractTotalFromPercentage = Round(success / percent, 0)
    
    ' �������
    Set matches = Nothing
    Set regEx = Nothing
End Function

' �����������������ֲ��䣩
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
