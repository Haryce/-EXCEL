Attribute VB_Name = "Module1"
Sub зп()

End Sub
Sub ЗП1()
Attribute ЗП1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ЗП1 Макрос
'

'
    Range("A1:G1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("A2:G2").Select
    Range("G2").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("A1:G1").Select
    ActiveCell.FormulaR1C1 = "ВЕДОМОСТЬ НАЧИСЛЕНИЯ ЗАРАБОТНОЙ ПЛАТЫ"
    Range("A1:G1").Select
    Selection.Font.Bold = True
    With Selection.Font
        .Name = "Calibri"
        .Size = 13
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Range("A2:G2").Select
    ActiveCell.FormulaR1C1 = "за октябрь 20__г."
    Range("A3:G24").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A20:G20").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A20:G20").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("D21:G24").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("C21").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A21:B24").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A21:B21").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("A22:B22").Select
    Range("B22").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("A23:B23").Select
    Range("B23").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("A24:B24").Select
    Range("B24").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("A21:B24").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("D23").Select
    Rows("3:3").RowHeight = 40.5
    Columns("A:A").ColumnWidth = 14.29
    Rows("3:3").RowHeight = 31.5
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Табельный номер"
    Range("A3").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "Ф.И.О."
    Range("B3:G3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "Оклад"
    Range("D3").Select
    ActiveCell.FormulaR1C1 = "Премия"
    Range("E3").Select
    ActiveCell.FormulaR1C1 = "Всего начислено"
    Range("E3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("E:E").ColumnWidth = 11
    Columns("D:D").ColumnWidth = 11.14
    Columns("C:C").ColumnWidth = 13.29
    Columns("C:C").ColumnWidth = 14.57
    Columns("B:B").ColumnWidth = 18.43
    Columns("B:B").ColumnWidth = 19.86

    Range("F3").Select
    ActiveCell.FormulaR1C1 = "Удержания"
    Range("F4").Select
    Columns("F:F").ColumnWidth = 11.57
    Range("G3").Select
    Columns("G:G").ColumnWidth = 11.43
    Range("G3").Select
    ActiveCell.FormulaR1C1 = "К выдаче"
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "27"
    Range("F4").Select
    ActiveCell.FormulaR1C1 = "13"
    Range("D4:G19").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "0.027"
    Range("F4").Select
    ActiveCell.FormulaR1C1 = "0.013"
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "0.27"
    Range("F4").Select
    ActiveCell.FormulaR1C1 = "0.13"
    Range("D4").Select
    Selection.NumberFormat = "0.00%"
    Selection.NumberFormat = "0%"
    Range("F4").Select
    Selection.NumberFormat = "0%"
    Range("A5").Select
    ActiveCell.FormulaR1C1 = "204"
    Range("A6").Select
    ActiveCell.FormulaR1C1 = "210"
    Range("A7").Select
    ActiveCell.FormulaR1C1 = "8/2/2024"
    Range("A8").Select
    ActiveCell.FormulaR1C1 = ""
    Range("A7").Select
    ActiveCell.FormulaR1C1 = "7/26/1900"
    Range("A7").Select
    ActiveCell.FormulaR1C1 = "7/26/1900"
    Range("A7").Select
    Selection.NumberFormat = "0.00"
    ActiveCell.FormulaR1C1 = "208"
    Range("A8").Select
    ActiveCell.FormulaR1C1 = "201"
    Range("A7").Select
    Selection.NumberFormat = "0"
    Range("A9").Select
    ActiveCell.FormulaR1C1 = "206"
    Range("A10").Select
    ActiveCell.FormulaR1C1 = "200"
    Range("A11").Select
    ActiveCell.FormulaR1C1 = "205"
    Range("A12").Select
    ActiveCell.FormulaR1C1 = "213"
    Range("A13").Select
    ActiveCell.FormulaR1C1 = "202"
    Range("A14").Select
    ActiveCell.FormulaR1C1 = "207"
    Range("A15").Select
    ActiveCell.FormulaR1C1 = "209"
    Range("A16").Select
    ActiveCell.FormulaR1C1 = "212"
    Range("A17").Select
    ActiveCell.FormulaR1C1 = "203"
    Range("A18").Select
    ActiveCell.FormulaR1C1 = "21"
    Range("A18").Select
    ActiveCell.FormulaR1C1 = "211"
    Range("B5").Select
    ActiveCell.FormulaR1C1 = "Галкин В.Ж"
    Range("B6").Select
    ActiveCell.FormulaR1C1 = "Дрынкина С.С"
    Range("B7").Select
    ActiveCell.FormulaR1C1 = "Жарова Г.А"
    Range("B8").Select
    ActiveCell.FormulaR1C1 = "Иванова И.Г"
    Range("B9").Select
    ActiveCell.FormulaR1C1 = "Орлова Н.Н"
    Range("B10").Select
    ActiveCell.FormulaR1C1 = "Петров И.Л"
    Range("B11").Select
    ActiveCell.FormulaR1C1 = "Потронов М.Т"
    Range("B12").Select
    ActiveCell.FormulaR1C1 = "Стелков Р.Х"
    Range("B13").Select
    ActiveCell.FormulaR1C1 = "Степанов"
    Range("B14").Select
    ActiveCell.FormulaR1C1 = "Степкин А.В"
    Range("B15").Select
    ActiveCell.FormulaR1C1 = "Стольникова О.Д"
    Range("B16").Select
    ActiveCell.FormulaR1C1 = "Шашкин Р.Н"
    Range("B17").Select
    ActiveCell.FormulaR1C1 = "Шорохо С.М"
    Range("B18").Select
    ActiveCell.FormulaR1C1 = "Шпаро Н.Г"
    Range("B19").Select
    ActiveCell.FormulaR1C1 = "Всего:"
    Range("A3:G3").Select
    Selection.Font.Bold = True
    Range("A5:B19").Select
    Selection.Font.Bold = True
    Range("A5:A18").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A21:B21").Select
    ActiveCell.FormulaR1C1 = "Минимальный доход"
    Range("A22:B22").Select
    ActiveCell.FormulaR1C1 = "Максимальный доход"
    Range("A23:B23").Select
    ActiveCell.FormulaR1C1 = "Средний доход"
    Range("C24").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("C24").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("C5").Select
    ActiveCell.FormulaR1C1 = "5900"
    Range("C6").Select
    ActiveCell.FormulaR1C1 = "8000"
    Range("C7").Select
    ActiveCell.FormulaR1C1 = "7300"
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "4850"
    Range("C9").Select
    ActiveCell.FormulaR1C1 = "6600"
    Range("C10").Select
    ActiveCell.FormulaR1C1 = "4500"
    Range("C11").Select
    ActiveCell.FormulaR1C1 = "6250"
    Range("C12").Select
    ActiveCell.FormulaR1C1 = "9050"
    Range("C13").Select
    ActiveCell.FormulaR1C1 = "5200"
    Range("C14").Select
    ActiveCell.FormulaR1C1 = "6950"
    Range("C15").Select
    ActiveCell.FormulaR1C1 = "7650"
    Range("C16").Select
    ActiveCell.FormulaR1C1 = "8700"
    Range("C17").Select
    ActiveCell.FormulaR1C1 = "5550"
    Range("C18").Select
    ActiveCell.FormulaR1C1 = "8350"
    Range("C5:C18").Select
    Selection.NumberFormat = "#,##0.00 $"
    Range("C5").Select
    ActiveCell.FormulaR1C1 = "5900"
    Range("D5").Select
    ActiveCell.FormulaR1C1 = "=PRODUCT(R[-1]C[-1]:RC)"
    Range("D5").Select
    ActiveCell.FormulaR1C1 = "=PRODUCT(RC[-1],R[-1]C)"
    Range("D5").Select
    Selection.AutoFill Destination:=Range("D5:D18"), Type:=xlFillDefault
    Range("D5:D18").Select
    Range("D7").Select
    ActiveCell.FormulaR1C1 = "=PRODUCT(RC[-1],R[-3]C)"
    Range("D6").Select
    ActiveCell.FormulaR1C1 = "=PRODUCT(RC[-1],R[-2]C)"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "=PRODUCT(RC[-1],R[-4]C)"
    Range("D10").Select
    ActiveCell.FormulaR1C1 = "=PRODUCT(RC[-1],R[-6]C)"
    Range("D9").Select
    ActiveCell.FormulaR1C1 = "=PRODUCT(RC[-1],R[-5]C)"
    Range("D11").Select
    ActiveCell.FormulaR1C1 = "=PRODUCT(RC[-1],R[-7]C)"
    Range("D12").Select
    ActiveCell.FormulaR1C1 = "=PRODUCT(RC[-1],R[-8]C)"
    Range("D14").Select
    ActiveCell.FormulaR1C1 = "=PRODUCT(RC[-1],R[-10]C)"
    Range("D16").Select
    ActiveCell.FormulaR1C1 = "=PRODUCT(RC[-1],R[-12]C)"
    Range("D15").Select
    ActiveCell.FormulaR1C1 = "=PRODUCT(RC[-1],R[-11]C)"
    Range("D13").Select
    ActiveCell.FormulaR1C1 = "=PRODUCT(RC[-1],R[-9]C)"
    Range("D17").Select
    ActiveCell.FormulaR1C1 = "=PRODUCT(RC[-1],R[-13]C)"
    Range("D18").Select
    ActiveCell.FormulaR1C1 = "=PRODUCT(RC[-1],R[-14]C)"
    Range("C19").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-14]C:R[-1]C)"
    Range("C19").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-14]C:R[-1]C)"
    Range("D5").Select
    ActiveCell.FormulaR1C1 = "=PRODUCT(R[14]C[-1],R[-1]C)"
    Range("D5").Select
    ActiveCell.FormulaR1C1 = "=PRODUCT(RC[-1],R[-1]C)"
    Range("D19").Select
    ActiveCell.FormulaR1C1 = "=PRODUCT(RC[-1],R[-15]C)"
    Range("D5:D19").Select
    Selection.NumberFormat = "#,##0.00 $"
    Range("E5").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-2],RC[-1])"
    Range("E5").Select
    Selection.AutoFill Destination:=Range("E5:E19"), Type:=xlFillDefault
    Range("E5:E19").Select
    Range("D5").Select
    ActiveCell.FormulaR1C1 = "=PRODUCT(RC[-1],R[-1]C)"
    Range("F5").Select
    ActiveCell.FormulaR1C1 = "=PRODUCT(RC[-1],R[-1]C)"
    Range("F5").Select
    Range("F6").Select
    ActiveCell.FormulaR1C1 = "=PRODUCT(RC[-1],R[-2]C[-2])"
    Range("F7").Select
    ActiveCell.FormulaR1C1 = ""
    Range("F6").Select
    ActiveCell.FormulaR1C1 = "=PRODUCT(R[1]C[-1],R[-2]C4)"
    Range("F6").Select
    Range("F6").Select
    ActiveCell.FormulaR1C1 = "=PRODUCT(R[1]C[-1],R4C4)"
    Range("F6").Select
    ActiveCell.FormulaR1C1 = "=PRODUCT(RC[-1],R4C4)"
    Range("F6").Select
    Selection.AutoFill Destination:=Range("F6:F19"), Type:=xlFillDefault
    Range("F6:F19").Select
    Range("F5:F19").Select
    Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Range("G5").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]-RC[-1]"
    Range("G5").Select
    Selection.AutoFill Destination:=Range("G5:G19"), Type:=xlFillDefault
    Range("G5:G19").Select
    Range("C21").Select
    ActiveCell.FormulaR1C1 = "=MIN(R[-16]C[4]:R[-2]C[4])"
    Range("C22").Select
    ActiveCell.FormulaR1C1 = "=МАКС"
    Range("C22").Select
    ActiveCell.FormulaR1C1 = "=MAX(R[-17]C[4]:R[-3]C[4])"
    Range("C23").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[-18]C[4]:R[-4]C[4])"
    Range("C24").Select
End Sub
