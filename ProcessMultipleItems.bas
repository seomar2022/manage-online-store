Attribute VB_Name = "Module1"
Sub ProcessMultipleItems()
    ' 마지막 행 찾기
    Dim lastRow As Long
    lastRow = Cells(Rows.count, "A").End(xlUp).row
    
    ' Dictionary 객체 생성
    Dim orderDict As Object
    Set orderDict = CreateObject("Scripting.Dictionary")
    
    ' 주문번호 빈도 계산
    Dim orderNumberCol As Long
    orderNumberCol = FindColumnIndex("주문번호")
    
    Dim i As Long
    For i = 2 To lastRow ' Assuming there is a header row
        Dim orderNumber As String
        orderNumber = Cells(i, orderNumberCol).Value
        
        ' 주문번호가 이미 딕셔너리에 있으면 값을 증가, 없으면 새로 추가
        If orderDict.Exists(orderNumber) Then
            orderDict(orderNumber) = orderDict(orderNumber) + 1
        Else
            orderDict.Add orderNumber, 1
        End If
    Next i
    
    '＇같은　주문번호일　경우　상품명과　수량을　모두　행방향으로　배치
    Dim targetRow As Long
    Dim count As Long
    Dim col As Long
    Dim startCol As Long
    startCol = FindColumnIndex("배송메시지") + 1
    Dim productName As Long
    productName = FindColumnIndex("주문상품명(옵션포함)")
    Dim quantity As Long
    quantity = FindColumnIndex("수량")
    
    
    For targetRow = 2 To orderDict.count + 1 '딕셔너리의 키 개수
        count = orderDict(Cells(targetRow, orderNumberCol).Value) ''타겟행의 빈도를 count에 넣음
        If count >= 2 Then
            For col = startCol To startCol + count - 1 Step 2
                Cells(targetRow, col).Value = Cells(targetRow + 1, productName).Value
                Cells(targetRow, col + 1).Value = Cells(targetRow + 1, quantity).Value
                Rows(targetRow + 1).Delete
            Next col
        End If
    Next targetRow

End Sub

















