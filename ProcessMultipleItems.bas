Attribute VB_Name = "Module1"
Sub ProcessMultipleItems()
    ' ������ �� ã��
    Dim lastRow As Long
    lastRow = Cells(Rows.count, "A").End(xlUp).row
    
    ' Dictionary ��ü ����
    Dim orderDict As Object
    Set orderDict = CreateObject("Scripting.Dictionary")
    
    ' �ֹ���ȣ �� ���
    Dim orderNumberCol As Long
    orderNumberCol = FindColumnIndex("�ֹ���ȣ")
    
    Dim i As Long
    For i = 2 To lastRow ' Assuming there is a header row
        Dim orderNumber As String
        orderNumber = Cells(i, orderNumberCol).Value
        
        ' �ֹ���ȣ�� �̹� ��ųʸ��� ������ ���� ����, ������ ���� �߰�
        If orderDict.Exists(orderNumber) Then
            orderDict(orderNumber) = orderDict(orderNumber) + 1
        Else
            orderDict.Add orderNumber, 1
        End If
    Next i
    
    '���������ֹ���ȣ�ϡ���졡��ǰ���������������Ρ���������Ρ���ġ
    Dim targetRow As Long
    Dim count As Long
    Dim col As Long
    Dim startCol As Long
    startCol = FindColumnIndex("��۸޽���") + 1
    Dim productName As Long
    productName = FindColumnIndex("�ֹ���ǰ��(�ɼ�����)")
    Dim quantity As Long
    quantity = FindColumnIndex("����")
    
    
    For targetRow = 2 To orderDict.count + 1 '��ųʸ��� Ű ����
        count = orderDict(Cells(targetRow, orderNumberCol).Value) ''Ÿ������ �󵵸� count�� ����
        If count >= 2 Then
            For col = startCol To startCol + count - 1 Step 2
                Cells(targetRow, col).Value = Cells(targetRow + 1, productName).Value
                Cells(targetRow, col + 1).Value = Cells(targetRow + 1, quantity).Value
                Rows(targetRow + 1).Delete
            Next col
        End If
    Next targetRow

End Sub

















