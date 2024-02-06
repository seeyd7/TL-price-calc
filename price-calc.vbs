Sub Marza()
'
' Makro Marza
'
' Klawisz skrótu: Ctrl+Shift+M
'
    Columns("A:A").Select
    Selection.Replace What:=".", Replacement:=",", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 2), Array(2, 1)), TrailingMinusNumbers:=True
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "EAN"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Cena"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Transport"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(RC[-1]>=40,RC[-1]<=59.99,2.09),2.09,IF(AND(RC[-1]>=60,RC[-1]<=79.99),2.19,IF(AND(RC[-1]>=80,RC[-1]<=199.99),4.99,IF(AND(RC[-1]>=200,RC[-1]<=299.99),7.99,IF(RC[-1]>=300,8.99,""Błąd"")))))"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C5000")
    Range("C2:C5000").Select
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]/1.23"
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D5000")
    Range("D2:D5000").Select
    Selection.Copy
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("D:D").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Cena zakupu"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],'Export iHurt'!C1:C2,2,0)"
    Selection.AutoFill Destination:=Range("D2:D5000")
    Range("D2:D5000").Select
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Cena zakupu allegro"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]*0.095+RC[-2]+RC[-1]"
    Range("E2").Select
    Selection.AutoFill Destination:=Range("E2:E5000")
    Range("E2:E5000").Select
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Marża kwotowa"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=RC[-4]-RC[-1]"
    Range("F2").Select
    Selection.AutoFill Destination:=Range("F2:F5000")
    Range("F2:F5000").Select
    Columns("A:F").Select
    Columns("A:F").EntireColumn.AutoFit
    Range("A1:F1").Select
    Selection.Font.Bold = True
    Range("A1:F5000").Select
    Range("D4").Activate
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
    Columns("B:F").Select
    Selection.NumberFormat = "#,##0.00 $"
    Range("G2").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("Export BL").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Export BL").AutoFilter.Sort.SortFields.Add Key:= _
        Range("F1:F5000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Export BL").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
