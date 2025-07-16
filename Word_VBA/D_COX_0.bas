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
    
    ' ��ȡѡ�е��ı�
    selectedText = Selection.text
    
    ' �ж��ǵ�����Ԫ���Ƕ����Ԫ��
    If Selection.Cells.Count = 1 Then
        ' ������Ԫ����ȡ�����ڵ����ݣ�����У�
        bracketContent = ExtractBracketContent(selectedText)
        If bracketContent = "" Then
            ' ���û�����ţ���ʹ�������ı�
            bracketContent = selectedText
        End If
        
        ' ��ȡ����
        Set regex = CreateObject("VBScript.RegExp")
        regex.Global = True
        regex.IgnoreCase = True
        regex.Pattern = "-?\d+\.?\d*" ' ƥ������������С��
        
        ' ��������ƥ�������
        Set numMatches = regex.Execute(bracketContent)
        
        ' ����Ƿ��ҵ�������������
        If numMatches.Count < 2 Then
            MsgBox "δ�ҵ��㹻�����֣��޷���ȡ A �� B��", vbExclamation
            Exit Sub
        End If
        
        ' ��ȡ A �� B
        a = Val(numMatches(0).value)
        b = Val(numMatches(1).value)
        
        ' ���� A-B ��ʽ���ı���ȷ�� B Ϊ������
        If InStr(bracketContent, "-") > 0 And numMatches.Count = 2 Then
            b = Abs(b) ' ȡ B �ľ���ֵ
        End If
    ElseIf Selection.Cells.Count = 2 Then
        ' ������Ԫ�񣬷ֱ���ȡ A �� B
        Set leftCell = Selection.Cells(1).Range
        Set rightCell = Selection.Cells(2).Range
        a = Val(leftCell.text)
        b = Val(rightCell.text)
    Else
        MsgBox "��ѡ��һ����Ԫ�񣨰��� A �� B�����������ڵ�Ԫ��A ����B ���ң���", vbExclamation
        Exit Sub
    End If
    
    ' ��֤ A �� B
    If b <= a Then
        MsgBox "B ������� A��", vbExclamation
        Exit Sub
    End If
    
    ' ���� Cox �ع�ϵ�� ��
    Beta = (Log(a) + Log(b)) / 2
    
    ' �����׼�� SE
    SE = (Log(b) - Log(a)) / (2 * 1.96)
    
    ' ������ձ� HR
    HR = Exp(Beta)
    
    ' ���㿨��ֵ (Wald����)
    ChiSquare = (Beta / SE) ^ 2
    
    ' ����Pֵ (ʹ���Ż��Ľ��Ʒ���)
    PValue = ImprovedChiSquarePValue(ChiSquare, 1)
    
    ' ������ע����
    commentText = "HR: " & Format(HR, "0.0000") & vbCrLf & _
                  "95%CI: " & Format(a, "0.0000") & " - " & Format(b, "0.0000") & vbCrLf & _
                  "��: " & Format(Beta, "0.0000") & vbCrLf & _
                  "SE: " & Format(SE, "0.0000") & vbCrLf & _
                  "��2: " & Format(ChiSquare, "0.0000") & vbCrLf & _
                  "Pֵ: " & Format(PValue, "0.0000")
    
    ' �����ע
    With Selection
        ' ɾ��ԭ�е���ע��������ڣ�
        If .Comments.Count > 0 Then
            .Comments(1).Delete
        End If
        ' �������ע
        Set comment = .Range.Comments.Add(Range:=.Range, text:=commentText)
    End With
End Sub

Function ExtractBracketContent(text As String) As String
    ' ��ȡ�����ڵ����ݣ�֧��ȫ�ǺͰ�����ţ�
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "[$��].*?[$��]" ' ƥ�����ż�������
    
    ' ����ƥ�����������
    Dim matches As Object
    Set matches = regex.Execute(text)
    
    If matches.Count > 0 Then
        ' ��ȡ�����ڵ����ݣ�ȥ�����ţ�
        ExtractBracketContent = Mid(matches(0).value, 2, Len(matches(0).value) - 2)
    Else
        ' δ�ҵ�����
        ExtractBracketContent = ""
    End If
End Function

Function ImprovedChiSquarePValue(ChiSquare As Double, df As Integer) As Double
    ' �Ż��Ŀ���Pֵ���㷽����������df=1��
    ' ʹ�÷ֶν��Ʒ���߾���
    
    ' Ԥ�ȼ���Ĺؼ��㣨����ֵ, Pֵ��
    Static keyPoints(1 To 7, 1 To 2) As Double
    keyPoints(1, 1) = 0#:     keyPoints(1, 2) = 1#
    keyPoints(2, 1) = 0.455:  keyPoints(2, 2) = 0.5
    keyPoints(3, 1) = 1.642:  keyPoints(3, 2) = 0.2
    keyPoints(4, 1) = 2.706:  keyPoints(4, 2) = 0.1
    keyPoints(5, 1) = 3.841:  keyPoints(5, 2) = 0.05
    keyPoints(6, 1) = 5.024:  keyPoints(6, 2) = 0.025
    keyPoints(7, 1) = 6.635:  keyPoints(7, 2) = 0.01
    
    ' �߽���
    If ChiSquare <= 0 Then
        ImprovedChiSquarePValue = 1
        Exit Function
    ElseIf ChiSquare >= keyPoints(7, 1) Then
        ImprovedChiSquarePValue = keyPoints(7, 2)
        Exit Function
    End If
    
    ' ���Բ�ֵ
    Dim i As Integer
    For i = 1 To 6
        If ChiSquare >= keyPoints(i, 1) And ChiSquare < keyPoints(i + 1, 1) Then
            ImprovedChiSquarePValue = keyPoints(i, 2) + (keyPoints(i + 1, 2) - keyPoints(i, 2)) * _
                                    (ChiSquare - keyPoints(i, 1)) / (keyPoints(i + 1, 1) - keyPoints(i, 1))
            Exit Function
        End If
    Next i
    
    ' Ĭ�Ϸ�����СPֵ
    ImprovedChiSquarePValue = keyPoints(7, 2)
End Function

