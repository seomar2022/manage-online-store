Attribute VB_Name = "Module21"
Sub 전채널주문리스트()
'
' 전채널주문리스트 매크로
'

'

''잘 모르겠으면 F8로 한줄씩 실행해보기!!

    ''변수선언 ->Dim 문은 변수를 선언하는 데 사용됨. 생략해도 변수를 암시적으로 선언하게 되지만, 이는 가독성이나 코드 유지보수성 측면에서 바람직하지 않습니다.
    Dim lastRow As Long
     
    ''가장 마지막행 찾기.
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.count, "A").End(xlUp).row
    
    
    
    ''''열 해더명 바꾸기
    Cells.Replace What:="주문자명", Replacement:="주문자", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="상품명(한국어 쇼핑몰)", Replacement:="상품명", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="상품옵션", Replacement:="옵션", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="옵션+판매가", Replacement:="가격", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="수령인 우편번호", Replacement:="우편번호", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="수령인 주소(전체)", Replacement:="주소", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="주문서추가항목01_사은품 선택 (공통입력사항)", Replacement:="사은품", _
        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:= _
        False, ReplaceFormat:=False
        
    Columns("I:I").Select
    Selection.Replace What:="강아지용", Replacement:="독", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="고양이용", Replacement:="캣", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    '''' 주문채널별로 주문자 이름 셀에 색채우기. 카페24주문은 채우기 없음.
    Columns("C:C").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=SEARCH(""카카오톡 스토어"", $A1)"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=SEARCH(""스마트스토어"", $A1)"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$A1=""쿠팡"""
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 9408511
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    
    ''''L열 오른쪽에 열 삽입해서 박스열 추가.
    Range("N1").Select
    Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "박스"
    Columns("N:N").Select
    Selection.ColumnWidth = 3
    
    
    ''각 열의 너비 조절
    Columns("A:A").Select
    Selection.ColumnWidth = 0
    Columns("B:B").Select
    Selection.ColumnWidth = 8.13
    Columns("C:C").Select
    Selection.ColumnWidth = 6
    Columns("D:D").Select
    Selection.ColumnWidth = 0
    Columns("E:E").Select
    Selection.ColumnWidth = 30
    Columns("F:F").Select
    Selection.ColumnWidth = 9
    Columns("G:G").Select
    Selection.ColumnWidth = 3 ''세자리수 보이려면 3은 되야함.
    Columns("H:H").Select
    Selection.ColumnWidth = 6
    Columns("I:I").Select
    Selection.ColumnWidth = 4
    Columns("L:L").Select
    Selection.ColumnWidth = 30
    Columns("M:M").Select
    Selection.ColumnWidth = 13
    Columns("J:J").Select
    Selection.ColumnWidth = 8
    
    
    ''가격에 , 붙이기
    Columns("J:J").Select
    Selection.Style = "Comma [0]"
    Cells.Select
    
    ''행높이 자동맞춤
    Range("B1").Activate
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    ''''수량 2개이상인 항목 빨간색으로 채우기
    Range("G2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="2"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
  
    '''2개이상의 상품 주문 시 주문번호 회색으로 채움
    Columns("B:B").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=COUNTIF($B:$B, $B1) > 1"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("B1").Select
    ActiveWindow.View = xlPageLayoutView

    
    ''''시트 이름바꾸기
    ActiveSheet.Name = "주문리스트"
    
    ''''프린트 설정 시작
    Sheets("주문리스트").Select
        Application.CutCopyMode = False
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = "$1:$1" ''1행을 반복해서 프린트하기
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = "&D &T"
        .CenterHeader = "전채널 주문 리스트"
        .RightHeader = "&P/&N"
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintSheetEnd
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    ActiveWindow.View = xlNormalView
    ''''프린트 설정 끝
    
    
    '''' 배송 보류인 주문건의 상품명에 취소선 긋기
    Columns("E:E").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$K1=""배송 보류"""
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Strikethrough = True
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    ''''주문 상태가 적힌 열 숨기기
    Columns("K:K").Select
    Selection.EntireColumn.Hidden = True
    
    
    ''''정기배송건의 주문자명과 상품명을 하늘색으로 채우기.
    Range("D2:F2").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=0
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$P2 >=1"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.599963377788629
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    ActiveWindow.SmallScroll Down:=-30

    
    ''''********************************************************박스통계 시작*********************************************************************'''''
    '''' 행별 총중량 구하기. (옵션에 중량있으면 거기서 숫자만 가져오기. 없으면 중량열에서 가져오기)
    ''Q열 왼쪽쪽에 열 삽입해서 박스열 추가.
    Range("P1").Select
    Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "총중량"
    
    Range("P2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISBLANK(RC15), ""왼쪽에 중량 쓰기"", IF(IFERROR(FIND(""중량"", RC6), 0), MID(RC6, SEARCH(""="", RC6) + 1, SEARCH(""kg"",RC6) - SEARCH(""="", RC6) - 1), RC15)  * RC[-9])"
       ''=IF(IFERROR(FIND("중량", $G2), 0), MID(G2, SEARCH("=", G2) + 1, SEARCH("kg", G2) - SEARCH("=", G2) - 1), $O2)  * I2
       ''RC숫자는 현재 셀을 기준으로 행과열을 몇개나 움직이는지를 알려주는 방식으로 셀의 위치를 표시하는 듯. 근데 양수는 $를 붙이는데 음수는 $가 안붙는거 같다.
     
     
    ''''주문건별 총중량 구하기
    Range("R1").Select
    Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "주문건별 총중량"
    Range("R2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(COUNTIF(R2C[-16]:RC2, RC[-16])=COUNTIF(R2C2:R1000C2, RC[-16]), SUMIF(R2C2:R1000C2, RC[-16], R2C16:R1000C16), """")"
        '=IF(COUNTIF(A$2:$D2, A2)=COUNTIF($D$2:$D$1000, A2), SUMIF($D$2:$D$1000, A2, $Q$2:$Q$1000), "")
     
     
     ''''총 중량별 박스 크기를 지정하는 함수
     Range("N2").Select
   '  ActiveCell.FormulaR1C1 = "=IF(RC15>0, IF(RC15<1, 73, IF(RC15<2, 194, IF(RC15<4, 41, """"))), """")"
     ActiveCell.FormulaR1C1 = "=IF(RC18<1, 73, IF(RC18<2, 194, IF(RC18<3.8, 41, IF(RC18<=4, 420,IF(RC18<=4.3, 104 ,IF(RC18<8, 170,""↓""))))))"
                            '''=IF($S2<1,73,IF($S2<2,194,IF($S2<4,41,IF($S2=4,420,IF($S2<4.3,104,IF($S2<5,170,"-"))))))

     

    ''변수 선언
    Dim criteriaRange As String
    Dim criteria As String
    Dim boxCountCol As String
    
    boxCountCol = "C"
    
     ''박스 번호 입력
    ActiveSheet.Cells(lastRow + 2, "B").Value = "박스"
    ActiveSheet.Cells(lastRow + 2, boxCountCol).Value = "개수"
    ActiveSheet.Cells(lastRow + 3, "B").Value = 73
    ActiveSheet.Cells(lastRow + 4, "B").Value = 194
    ActiveSheet.Cells(lastRow + 5, "B").Value = 41
    ActiveSheet.Cells(lastRow + 6, "B").Value = 420
    ActiveSheet.Cells(lastRow + 7, "B").Value = 104
    ActiveSheet.Cells(lastRow + 8, "B").Value = 170
    ActiveSheet.Cells(lastRow + 9, "B").Value = 58
     
    ''사용할 박스 개수 세기
    ActiveSheet.Cells(lastRow + 3, boxCountCol).Select
   
   
   
    criteriaRange = "N:N" ''M열 전체를 가리킴.
    criteria = "B" & lastRow + 3 & ":B" & lastRow + 9 ''박스 번호가 적힌 범위를 가리킴.
    
    ' COUNTIF 함수 삽입
    ActiveSheet.Cells(lastRow + 3, boxCountCol).Formula = "=COUNTIF(" & criteriaRange & "," & criteria & ")"
        ''->=COUNTIF(M:M,'C20':'C27') 이런식으로 두번째요소들이 작은따옴표에 감싸져서 나옴. 어떻게 고쳐야할지 모르겠음.
        ''->criteriaRange랑 criteria를 수정했더니갑자기 정상적으로 나옴. 도대체 뭘까.
    Selection.AutoFill Destination:=Range(boxCountCol & lastRow + 3 & ":" & boxCountCol & lastRow + 9), Type:=xlFillDefault
    ''''********************************************************박스통계 끝*********************************************************************'''''


    
    
    ''''일련번호 붙이기. 같은 주문번호가 여러개 있더라도 하나의 주문건이기때문에 하나로 카운트
    Range("B1").Select
    Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove ''주문번호열 왼쪽에 열 삽입
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "연번"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(COUNTIF(R2C3:RC[1], RC[1])=1, MAX(R1C2:R[-1]C)+1, IFERROR(VLOOKUP(RC[1], R1C2:R[-1]C3, 2, FALSE), """"))"
       ''=IF(COUNTIF($C$2:C2, C2)=1, MAX($B$1:B1)+1, IFERROR(VLOOKUP(C2, $B$1:$C1, 2, FALSE), ""))
    Columns("B:B").Select
    Selection.ColumnWidth = 2.25 ''열 너비 조정
    
    ''''각 행에 줄긋기
    Range("A1:O1").Select
    Range(Selection, Selection.End(xlDown)).Select ''컨트롤 방향키 아래
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.349986266670736
        .Weight = xlThin
    End With
    
    
    ''''함수가 적용된 열들 맨 밑행까지 자동채우기
    ''자동채우기 적용할 열
    cols = Array("B", "O", "Q", "S")
    
    For i = LBound(cols) To UBound(cols) ' 각 열에 대해 AutoFill 수행
        '' 시작 셀 선택
        Range(cols(i) & 2).Select
        '' 범위 설정 및 AutoFill 수행
        Selection.AutoFill Destination:=Range(cols(i) & 2 & ":" & cols(i) & lastRow), Type:=xlFillDefault
    Next i
    

    
    Range("B1").Select
End Sub


