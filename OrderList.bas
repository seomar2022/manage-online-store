Attribute VB_Name = "Module21"
Sub ��ä���ֹ�����Ʈ()
'
' ��ä���ֹ�����Ʈ ��ũ��
'

'

''�� �𸣰����� F8�� ���پ� �����غ���!!

    ''�������� ->Dim ���� ������ �����ϴ� �� ����. �����ص� ������ �Ͻ������� �����ϰ� ������, �̴� �������̳� �ڵ� ���������� ���鿡�� �ٶ������� �ʽ��ϴ�.
    Dim lastRow As Long
     
    ''���� �������� ã��.
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.count, "A").End(xlUp).row
    
    
    
    ''''�� �ش��� �ٲٱ�
    Cells.Replace What:="�ֹ��ڸ�", Replacement:="�ֹ���", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��ǰ��(�ѱ��� ���θ�)", Replacement:="��ǰ��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="��ǰ�ɼ�", Replacement:="�ɼ�", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="�ɼ�+�ǸŰ�", Replacement:="����", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="������ �����ȣ", Replacement:="�����ȣ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="������ �ּ�(��ü)", Replacement:="�ּ�", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="�ֹ����߰��׸�01_����ǰ ���� (�����Է»���)", Replacement:="����ǰ", _
        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:= _
        False, ReplaceFormat:=False
        
    Columns("I:I").Select
    Selection.Replace What:="��������", Replacement:="��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="����̿�", Replacement:="Ĺ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    '''' �ֹ�ä�κ��� �ֹ��� �̸� ���� ��ä���. ī��24�ֹ��� ä��� ����.
    Columns("C:C").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=SEARCH(""īī���� �����"", $A1)"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=SEARCH(""����Ʈ�����"", $A1)"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$A1=""����"""
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 9408511
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    
    ''''L�� �����ʿ� �� �����ؼ� �ڽ��� �߰�.
    Range("N1").Select
    Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "�ڽ�"
    Columns("N:N").Select
    Selection.ColumnWidth = 3
    
    
    ''�� ���� �ʺ� ����
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
    Selection.ColumnWidth = 3 ''���ڸ��� ���̷��� 3�� �Ǿ���.
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
    
    
    ''���ݿ� , ���̱�
    Columns("J:J").Select
    Selection.Style = "Comma [0]"
    Cells.Select
    
    ''����� �ڵ�����
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
    
    ''''���� 2���̻��� �׸� ���������� ä���
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
  
    '''2���̻��� ��ǰ �ֹ� �� �ֹ���ȣ ȸ������ ä��
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

    
    ''''��Ʈ �̸��ٲٱ�
    ActiveSheet.Name = "�ֹ�����Ʈ"
    
    ''''����Ʈ ���� ����
    Sheets("�ֹ�����Ʈ").Select
        Application.CutCopyMode = False
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = "$1:$1" ''1���� �ݺ��ؼ� ����Ʈ�ϱ�
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = "&D &T"
        .CenterHeader = "��ä�� �ֹ� ����Ʈ"
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
    ''''����Ʈ ���� ��
    
    
    '''' ��� ������ �ֹ����� ��ǰ�� ��Ҽ� �߱�
    Columns("E:E").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$K1=""��� ����"""
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Strikethrough = True
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    ''''�ֹ� ���°� ���� �� �����
    Columns("K:K").Select
    Selection.EntireColumn.Hidden = True
    
    
    ''''�����۰��� �ֹ��ڸ�� ��ǰ���� �ϴû����� ä���.
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

    
    ''''********************************************************�ڽ���� ����*********************************************************************'''''
    '''' �ະ ���߷� ���ϱ�. (�ɼǿ� �߷������� �ű⼭ ���ڸ� ��������. ������ �߷������� ��������)
    ''Q�� �����ʿ� �� �����ؼ� �ڽ��� �߰�.
    Range("P1").Select
    Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "���߷�"
    
    Range("P2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISBLANK(RC15), ""���ʿ� �߷� ����"", IF(IFERROR(FIND(""�߷�"", RC6), 0), MID(RC6, SEARCH(""="", RC6) + 1, SEARCH(""kg"",RC6) - SEARCH(""="", RC6) - 1), RC15)  * RC[-9])"
       ''=IF(IFERROR(FIND("�߷�", $G2), 0), MID(G2, SEARCH("=", G2) + 1, SEARCH("kg", G2) - SEARCH("=", G2) - 1), $O2)  * I2
       ''RC���ڴ� ���� ���� �������� ������� ��� �����̴����� �˷��ִ� ������� ���� ��ġ�� ǥ���ϴ� ��. �ٵ� ����� $�� ���̴µ� ������ $�� �Ⱥٴ°� ����.
     
     
    ''''�ֹ��Ǻ� ���߷� ���ϱ�
    Range("R1").Select
    Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "�ֹ��Ǻ� ���߷�"
    Range("R2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(COUNTIF(R2C[-16]:RC2, RC[-16])=COUNTIF(R2C2:R1000C2, RC[-16]), SUMIF(R2C2:R1000C2, RC[-16], R2C16:R1000C16), """")"
        '=IF(COUNTIF(A$2:$D2, A2)=COUNTIF($D$2:$D$1000, A2), SUMIF($D$2:$D$1000, A2, $Q$2:$Q$1000), "")
     
     
     ''''�� �߷��� �ڽ� ũ�⸦ �����ϴ� �Լ�
     Range("N2").Select
   '  ActiveCell.FormulaR1C1 = "=IF(RC15>0, IF(RC15<1, 73, IF(RC15<2, 194, IF(RC15<4, 41, """"))), """")"
     ActiveCell.FormulaR1C1 = "=IF(RC18<1, 73, IF(RC18<2, 194, IF(RC18<3.8, 41, IF(RC18<=4, 420,IF(RC18<=4.3, 104 ,IF(RC18<8, 170,""��""))))))"
                            '''=IF($S2<1,73,IF($S2<2,194,IF($S2<4,41,IF($S2=4,420,IF($S2<4.3,104,IF($S2<5,170,"-"))))))

     

    ''���� ����
    Dim criteriaRange As String
    Dim criteria As String
    Dim boxCountCol As String
    
    boxCountCol = "C"
    
     ''�ڽ� ��ȣ �Է�
    ActiveSheet.Cells(lastRow + 2, "B").Value = "�ڽ�"
    ActiveSheet.Cells(lastRow + 2, boxCountCol).Value = "����"
    ActiveSheet.Cells(lastRow + 3, "B").Value = 73
    ActiveSheet.Cells(lastRow + 4, "B").Value = 194
    ActiveSheet.Cells(lastRow + 5, "B").Value = 41
    ActiveSheet.Cells(lastRow + 6, "B").Value = 420
    ActiveSheet.Cells(lastRow + 7, "B").Value = 104
    ActiveSheet.Cells(lastRow + 8, "B").Value = 170
    ActiveSheet.Cells(lastRow + 9, "B").Value = 58
     
    ''����� �ڽ� ���� ����
    ActiveSheet.Cells(lastRow + 3, boxCountCol).Select
   
   
   
    criteriaRange = "N:N" ''M�� ��ü�� ����Ŵ.
    criteria = "B" & lastRow + 3 & ":B" & lastRow + 9 ''�ڽ� ��ȣ�� ���� ������ ����Ŵ.
    
    ' COUNTIF �Լ� ����
    ActiveSheet.Cells(lastRow + 3, boxCountCol).Formula = "=COUNTIF(" & criteriaRange & "," & criteria & ")"
        ''->=COUNTIF(M:M,'C20':'C27') �̷������� �ι�°��ҵ��� ��������ǥ�� �������� ����. ��� ���ľ����� �𸣰���.
        ''->criteriaRange�� criteria�� �����ߴ��ϰ��ڱ� ���������� ����. ����ü ����.
    Selection.AutoFill Destination:=Range(boxCountCol & lastRow + 3 & ":" & boxCountCol & lastRow + 9), Type:=xlFillDefault
    ''''********************************************************�ڽ���� ��*********************************************************************'''''


    
    
    ''''�Ϸù�ȣ ���̱�. ���� �ֹ���ȣ�� ������ �ִ��� �ϳ��� �ֹ����̱⶧���� �ϳ��� ī��Ʈ
    Range("B1").Select
    Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove ''�ֹ���ȣ�� ���ʿ� �� ����
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "����"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(COUNTIF(R2C3:RC[1], RC[1])=1, MAX(R1C2:R[-1]C)+1, IFERROR(VLOOKUP(RC[1], R1C2:R[-1]C3, 2, FALSE), """"))"
       ''=IF(COUNTIF($C$2:C2, C2)=1, MAX($B$1:B1)+1, IFERROR(VLOOKUP(C2, $B$1:$C1, 2, FALSE), ""))
    Columns("B:B").Select
    Selection.ColumnWidth = 2.25 ''�� �ʺ� ����
    
    ''''�� �࿡ �ٱ߱�
    Range("A1:O1").Select
    Range(Selection, Selection.End(xlDown)).Select ''��Ʈ�� ����Ű �Ʒ�
    
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
    
    
    ''''�Լ��� ����� ���� �� ������� �ڵ�ä���
    ''�ڵ�ä��� ������ ��
    cols = Array("B", "O", "Q", "S")
    
    For i = LBound(cols) To UBound(cols) ' �� ���� ���� AutoFill ����
        '' ���� �� ����
        Range(cols(i) & 2).Select
        '' ���� ���� �� AutoFill ����
        Selection.AutoFill Destination:=Range(cols(i) & 2 & ":" & cols(i) & lastRow), Type:=xlFillDefault
    Next i
    

    
    Range("B1").Select
End Sub


