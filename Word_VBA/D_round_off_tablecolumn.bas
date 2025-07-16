Attribute VB_Name = "D_round_off_tablecolumn"
'�Ա���е�ĳһ���������룬��ɺ���ʾ��ɫ����
Sub d_Round_off_TableColumn_Withoutkimi()
    Dim tbl As Table
    Dim cell As cell
    Dim colNum As Integer
    Dim cellValue As String
    Dim numericValue As Double
    Dim i As Integer
    
    ' ����Ƿ���ѡ�еı��
    If Selection.Information(wdWithInTable) = False Then
        MsgBox "����ѡ�б���е�ĳһ�У�", vbExclamation
        Exit Sub
    End If
    
    ' ��ȡѡ�е��к�
    colNum = Selection.Information(wdStartOfRangeColumnNumber)
    
    ' ��ȡѡ�еı��
    Set tbl = Selection.Tables(1)
    
    ' ��������ÿһ��
    For i = 1 To tbl.Rows.Count
        ' ��ȡ��ǰ�е�ָ���еĵ�Ԫ��
        Set cell = tbl.cell(i, colNum)
        
        ' ��ȡ��Ԫ�����ݲ�ȥ��β���Ķ�����
        cellValue = Trim(cell.Range.text)
        If Right(cellValue, 2) = Chr(13) & Chr(7) Then
            cellValue = Left(cellValue, Len(cellValue) - 2)
        End If
        
        ' ����Ƿ�Ϊ����
        If IsNumeric(cellValue) Then
            ' ���ַ���ת��Ϊ����
            numericValue = CDbl(cellValue)
            
            ' �������뱣����λС��
            numericValue = Round(numericValue, 2)
            
            ' �����д�ص�Ԫ�񣬲�����������ɫΪ��ɫ
            With cell.Range
                .text = numericValue
                .Font.Color = wdColorRed
            End With
        End If
    Next i
    
    MsgBox "������ɣ��޸�����������Ϊ��ɫ���壡", vbInformation
End Sub
