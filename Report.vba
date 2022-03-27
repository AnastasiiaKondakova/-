Sub Report()

    Dim lastrow As Integer
    Dim lastrow_pivot1 As Integer
    Dim lastrow_pivot2 As Integer
    Dim lastrow_pivot3 As Integer
    Dim lastrow_pivot4 As Integer
    Dim lastrow_pivot5 As Integer
    
    Application.ScreenUpdating = False
    
    ' додаємо новий аркуш Combined
    With ThisWorkbook
        Set ws = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        ws.Name = "Combined"
    End With
    
    ' копіємо на нього TREATIES
    Sheets("TREATIES").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Combined").Select
    ActiveSheet.Paste
    Range("I1").Select
    Application.CutCopyMode = False
    
    Selection.Interior.ThemeColor = xlThemeColorDark1
    Selection.Interior.TintAndShade = -0.149998474074526
    
    ' додаємо стовпець Name з CLIENTS
    ActiveCell.FormulaR1C1 = "Name"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-6],CLIENTS!C[-8]:C[-2],2,0)"
    Range("I2").Select
    Selection.AutoFill Destination:=Range(Selection, Selection.End(xlDown))
    
    ' додаємо стовпець Type з CLIENTS
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Type"
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-7],CLIENTS!C[-9]:C[-3],7,0)"
    Range("J2").Select
    Selection.AutoFill Destination:=Range(Selection, Selection.End(xlDown))
    Cells.SpecialCells(xlCellTypeFormulas, xlErrors).Clear
    
    ' форматуємо аркуш Combined
    Range("I1:J1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    Selection.Borders(xlInsideVertical).Weight = xlThin
    Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Selection.Borders(xlInsideHorizontal).Weight = xlThin
    Selection.Font.Name = "Arial"
    Selection.Font.Size = 10
    
    Range("I1").Select
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    
    Range("J1").Select
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ThemeColor = xlThemeColorDark1
    Selection.Interior.TintAndShade = -0.149998474074526
    
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("I1").Select
    
    ' додаємо новий аркуш Pivot1
    lastrow = Worksheets("Combined").Range("A" & Rows.Count).End(xlUp).Row
    With ThisWorkbook
        Set ws = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        ws.Name = "Pivot1"
    End With
    
    ' будуємо зведену таблицю на цьому аркуші (Type, Amount)
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Combined!A1:J" & CStr(lastrow), Version:=7).CreatePivotTable TableDestination:= _
        "Pivot1!R4C1", TableName:="Сводная таблица1", DefaultVersion:=7

    Sheets("Pivot1").Select
    Cells(4, 1).Select
    With ActiveSheet.PivotTables("Сводная таблица1").PivotFields("Type")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    ActiveSheet.PivotTables("Сводная таблица1").AddDataField ActiveSheet. _
        PivotTables("Сводная таблица1").PivotFields("Amount"), "Сумма по полю Amount", _
        xlSum
        
    ' додаємо новий аркуш Pivot2
    With ThisWorkbook
        Set ws = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        ws.Name = "Pivot2"
    End With
    
    ' будуємо зведену таблицю на цьому аркуші (Payment terms, Amount)
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Combined!A1:J" & CStr(lastrow), Version:=7).CreatePivotTable TableDestination:= _
        "Pivot2!R4C1", TableName:="Сводная таблица2", DefaultVersion:=7
    
    Sheets("Pivot2").Select
    Cells(4, 1).Select
    With ActiveSheet.PivotTables("Сводная таблица2").PivotFields("Payment terms")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    ActiveSheet.PivotTables("Сводная таблица2").AddDataField ActiveSheet. _
        PivotTables("Сводная таблица2").PivotFields("Amount"), "Сумма по полю Amount", _
        xlSum
        
    ' додаємо новий аркуш Pivot3
    With ThisWorkbook
        Set ws = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        ws.Name = "Pivot3"
    End With
    
    ' будуємо зведену таблицю на цьому аркуші (Name, Amount, фільтр - Closed)
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Combined!A1:J" & CStr(lastrow), Version:=7).CreatePivotTable TableDestination:= _
        "Pivot3!R4C1", TableName:="Сводная таблица3", DefaultVersion:=7
    
    Sheets("Pivot3").Select
    Cells(4, 1).Select
    With ActiveSheet.PivotTables("Сводная таблица3").PivotFields("Closed")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Сводная таблица3").PivotFields("Name")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    ActiveSheet.PivotTables("Сводная таблица3").AddDataField ActiveSheet. _
        PivotTables("Сводная таблица3").PivotFields("Amount"), "Сумма по полю Amount", _
        xlSum
    ActiveSheet.PivotTables("Сводная таблица3").PivotFields("Closed").CurrentPage _
        = "0"
    
    ' додаємо новий аркуш Pivot4
    With ThisWorkbook
        Set ws = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        ws.Name = "Pivot4"
    End With
    
   ' будуємо зведену таблицю на цьому аркуші (Name, Amount, фільтр - Closed)
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Combined!A1:J" & CStr(lastrow), Version:=7).CreatePivotTable TableDestination:= _
        "Pivot4!R4C1", TableName:="Сводная таблица4", DefaultVersion:=7
    
    Sheets("Pivot4").Select
    Cells(4, 1).Select
    With ActiveSheet.PivotTables("Сводная таблица4").PivotFields("Closed")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Сводная таблица4").PivotFields("Name")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    ActiveSheet.PivotTables("Сводная таблица4").AddDataField ActiveSheet. _
        PivotTables("Сводная таблица4").PivotFields("Amount"), "Сумма по полю Amount", _
        xlSum
    ActiveSheet.PivotTables("Сводная таблица4").PivotFields("Closed").CurrentPage _
        = "(All)"
        
    ' додаємо новий аркуш Pivot5
    With ThisWorkbook
        Set ws = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        ws.Name = "Pivot5"
    End With
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Combined!A1:J" & CStr(lastrow), Version:=7).CreatePivotTable TableDestination:= _
        "Pivot5!R4C1", TableName:="Сводная таблица5", DefaultVersion:=7
    
    ' будуємо зведену таблицю на цьому аркуші (Рік (FirstDate), Amount)
    Sheets("Pivot5").Select
    Cells(4, 1).Select
    With ActiveSheet.PivotTables("Сводная таблица5").PivotFields("FirstDate")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    ActiveSheet.PivotTables("Сводная таблица5").PivotFields("FirstDate").AutoGroup
    ActiveSheet.PivotTables("Сводная таблица5").PivotFields("Кварталы"). _
        Orientation = xlHidden
    ActiveSheet.PivotTables("Сводная таблица5").PivotFields("FirstDate"). _
        Orientation = xlHidden
    ActiveSheet.PivotTables("Сводная таблица5").AddDataField ActiveSheet. _
        PivotTables("Сводная таблица5").PivotFields("Amount"), "Сумма по полю Amount", _
        xlSum
        
    ' копіюємо дані із зведеної таблиці на аркуші Pivot1, видаляємо рядок із сумою, сортуємо
    Sheets("Pivot1").Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Range("D4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    lastrow_pivot1 = Worksheets("Pivot1").Range("D" & Rows.Count).End(xlUp).Row
    Range("D" & CStr(lastrow_pivot1) & ":E" & CStr(lastrow_pivot1)).Select
    Selection.ClearContents
    
    Range("E5").Select
    ActiveWorkbook.Worksheets("Pivot1").Sort.SortFields.Add2 Key:=Range(Selection, Selection.End(xlDown)) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    Range("E5:D5").Select
    With ActiveWorkbook.Worksheets("Pivot1").Sort
        .SetRange Range(Selection, Selection.End(xlDown))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "Тип клієнта"
    Range("E4").Select
    ActiveCell.FormulaR1C1 = "Сума договорів (грн)"
    
    'будуємо графік на вкладці Pivot1
    Sheets("Pivot1").Select
    ActiveSheet.Shapes.AddChart2(251, xlDoughnut).Select
    
    lastrow_pivot5 = Worksheets("Pivot1").Range("D" & Rows.Count).End(xlUp).Row
    ActiveChart.SetSourceData Source:=Range("Pivot1!$D$4:$E$" & CStr(lastrow_pivot5))
    ActiveChart.ClearToMatchStyle
    ActiveChart.ChartStyle = 258
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "сума договорів залежно від типу клієнта"
    
    ActiveSheet.ChartObjects("Диаграмма 1").Activate
    ActiveSheet.Shapes("Диаграмма 1").Name = "Chart2"
    Selection.Name = "Chart2"
    
    ' копіюємо дані із зведеної таблиці на аркуші Pivot2, видаляємо рядок із сумою, сортуємо
    Sheets("Pivot2").Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Range("D4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    lastrow_pivot2 = Worksheets("Pivot2").Range("D" & Rows.Count).End(xlUp).Row
    Range("D" & CStr(lastrow_pivot2) & ":E" & CStr(lastrow_pivot2)).Select
    Selection.ClearContents
    
    Range("E5").Select
    ActiveWorkbook.Worksheets("Pivot2").Sort.SortFields.Add2 Key:=Range(Selection, Selection.End(xlDown)) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    Range("E5:D5").Select
    With ActiveWorkbook.Worksheets("Pivot2").Sort
        .SetRange Range(Selection, Selection.End(xlDown))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "Тип оплати"
    Range("E4").Select
    ActiveCell.FormulaR1C1 = "Сума договорів (грн)"
    
    'будуємо графік на вкладці Pivot2
    Sheets("Pivot2").Select
    ActiveSheet.Shapes.AddChart2(251, xlDoughnut).Select
    
    lastrow_pivot5 = Worksheets("Pivot2").Range("D" & Rows.Count).End(xlUp).Row
    ActiveChart.SetSourceData Source:=Range("Pivot2!$D$4:$E$" & CStr(lastrow_pivot5))
    ActiveChart.ClearToMatchStyle
    ActiveChart.ChartStyle = 258
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "сума договорів залежно від типу оплати"
    
    ActiveSheet.ChartObjects("Диаграмма 1").Activate
    ActiveSheet.Shapes("Диаграмма 1").Name = "Chart3"
    Selection.Name = "Chart3"
    
    ' копіюємо дані із зведеної таблиці на аркуші Pivot3, сортуємо
    Sheets("Pivot3").Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Range("D4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    lastrow_pivot3 = Worksheets("Pivot3").Range("D" & Rows.Count).End(xlUp).Row
    Range("D" & CStr(lastrow_pivot3) & ":E" & CStr(lastrow_pivot3)).Select
    Selection.ClearContents
    
    Range("E5").Select
    ActiveWorkbook.Worksheets("Pivot3").Sort.SortFields.Add2 Key:=Range(Selection, Selection.End(xlDown)) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    Range("E5:D5").Select
    With ActiveWorkbook.Worksheets("Pivot3").Sort
        .SetRange Range(Selection, Selection.End(xlDown))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "Клієнт"
    Range("E4").Select
    ActiveCell.FormulaR1C1 = "Сума незакритих договорів (грн)"
    
    ' додаємо стовпець з часткою кожного клієнта
    Range("G5").Select
    ActiveCell.FormulaR1C1 = "=SUM(C[-2])"
    Range("F5").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]/R5C7"
    Selection.AutoFill Destination:=Range(Selection, Selection.End(xlDown))
    
    ' копіюємо дані із зведеної таблиці на аркуші Pivot4, сортуємо
    Sheets("Pivot4").Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Range("D4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    lastrow_pivot4 = Worksheets("Pivot4").Range("D" & Rows.Count).End(xlUp).Row
    Range("D" & CStr(lastrow_pivot4) & ":E" & CStr(lastrow_pivot4)).Select
    Selection.ClearContents
    
    Range("E5").Select
    ActiveWorkbook.Worksheets("Pivot4").Sort.SortFields.Add2 Key:=Range(Selection, Selection.End(xlDown)) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    Range("E5:D5").Select
    With ActiveWorkbook.Worksheets("Pivot4").Sort
        .SetRange Range(Selection, Selection.End(xlDown))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "Клієнт"
    Range("E4").Select
    ActiveCell.FormulaR1C1 = "Сума договорів (грн)"
    
   ' додаємо стовпець з часткою кожного клієнта
    Range("G5").Select
    ActiveCell.FormulaR1C1 = "=SUM(C[-2])"
    Range("F5").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]/R5C7"
    Selection.AutoFill Destination:=Range(Selection, Selection.End(xlDown))
    
    ' копіюємо дані із зведеної таблиці на аркуші Pivot5, видаляємо рядок із сумою
    Sheets("Pivot5").Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Range("D4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    lastrow_pivot5 = Worksheets("Pivot5").Range("D" & Rows.Count).End(xlUp).Row
    Range("D" & CStr(lastrow_pivot5) & ":E" & CStr(lastrow_pivot5)).Select
    Selection.ClearContents
    
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "Рік"
    Range("E4").Select
    ActiveCell.FormulaR1C1 = "Сума договорів (грн)"
    
    'будуємо графік на вкладці Pivot5
    Sheets("Pivot5").Select
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    
    lastrow_pivot5 = Worksheets("Pivot5").Range("D" & Rows.Count).End(xlUp).Row
    ActiveChart.SetSourceData Source:=Range("Pivot5!$D$4:$E$" & CStr(lastrow_pivot5))
    
    ActiveChart.ChartTitle.Select
    Selection.Delete
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    Selection.Delete
    
    ActiveChart.ChartArea.Format.TextFrame2.TextRange.Font.Name = "Arial"
    ActiveChart.ChartArea.Format.TextFrame2.TextRange.Font.Size = 9
    ActiveChart.ChartArea.Font.Color = RGB(0, 0, 0)
    ActiveChart.PlotArea.Select
    ActiveChart.FullSeriesCollection(1).Select
    
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
        .Transparency = 0
    End With
    Selection.MarkerStyle = -4105
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 255)
    End With
    
    ActiveSheet.ChartObjects("Диаграмма 1").Activate
    ActiveSheet.Shapes("Диаграмма 1").Name = "Chart1"
    Selection.Name = "Chart1"
    
    'створюємо, заповнюємо, та форматуємо аркуш DASHBOARD
    With ThisWorkbook
        Set ws = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        ws.Name = "DASHBOARD"
    End With
    
    ActiveWorkbook.Sheets("DASHBOARD").Tab.Color = 192

    Cells.Select
    Selection.Interior.PatternColorIndex = xlAutomatic
    Selection.Font.Name = "Arial"
    Selection.Font.Size = 10

    Range("A1:S2").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.VerticalAlignment = xlCenter
    Selection.MergeCells = True
    Selection.Interior.Pattern = xlSolid
    Selection.Interior.Color = 192

    ActiveCell.FormulaR1C1 = "ГЛОБИНСЬКИЙ М'ЯСОКОМБІНАТ - ПРИКЛАД ЗВІТУ"
    Range("A1:S2").Select
    Selection.Font.Bold = True
    Selection.Font.ThemeColor = xlThemeColorDark1
    Selection.Font.TintAndShade = 0
    Selection.Font.Size = 12
    
    Columns("A:A").ColumnWidth = 1.44
    Columns("S:S").ColumnWidth = 1.44
    Columns("O:O").ColumnWidth = 1.7
    
    Range("B5").Select
    ActiveCell.FormulaR1C1 = "Ключові метрики"
    Selection.HorizontalAlignment = xlCenter
    Selection.Font.Size = 11
    
    Range("B5:N5").Select
    Selection.Merge
    Selection.Font.Bold = True
    Selection.Interior.Pattern = xlSolid
    Selection.Interior.ThemeColor = xlThemeColorAccent2
    Selection.Interior.TintAndShade = 0.799981688894314
    
    Range("B6").Select
    ActiveCell.FormulaR1C1 = "Кількість активних клієнтів"
    Range("B7").Select
    ActiveCell.FormulaR1C1 = "Максимальна сума договору (грн)"
    Range("B8").Select
    ActiveCell.FormulaR1C1 = "Середня сума договору (грн)"
    
    Columns("H:H").ColumnWidth = 2.56
    
    Range("I6").Select
    ActiveCell.FormulaR1C1 = "Кількість договорів"
    Range("I7").Select
    ActiveCell.FormulaR1C1 = "Кількість відкритих договорів"
    Range("I8").Select
    ActiveCell.FormulaR1C1 = "Середня сума відкритого договору"
    
    Range("B6:E6").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlLeft
    Range("B7:E7").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlLeft
    Range("B8:E8").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlLeft
    Range("I6:L6").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlLeft
    Range("I7:L7").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlLeft
    Range("I8:L8").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlLeft
    
    Range("P5").Select
    ActiveCell.FormulaR1C1 = "Динаміка загальної суми договорів (грн)"
    Range("P5:R5").Select
    Selection.Merge
    Selection.Font.Bold = True
    Selection.Interior.Pattern = xlSolid
    Selection.Interior.ThemeColor = xlThemeColorAccent2
    Selection.Interior.TintAndShade = 0.799981688894314
    Selection.HorizontalAlignment = xlCenter
    Selection.Font.Size = 11
    
    Range("P16").Select
    ActiveCell.FormulaR1C1 = "Топ-5 клієнтів (відкриті договори)"
    Range("P16:R16").Select
    Selection.Merge
    Selection.Font.Bold = True
    Selection.Interior.Pattern = xlSolid
    Selection.Interior.ThemeColor = xlThemeColorAccent2
    Selection.Interior.TintAndShade = 0.799981688894314
    Selection.HorizontalAlignment = xlCenter
    Selection.Font.Size = 11
    
    Range("P23").Select
    ActiveCell.FormulaR1C1 = "Топ-5 клієнтів (усі договори)"
    Range("P23:R23").Select
    Selection.Merge
    Selection.Font.Bold = True
    Selection.Interior.Pattern = xlSolid
    Selection.Interior.ThemeColor = xlThemeColorAccent2
    Selection.Interior.TintAndShade = 0.799981688894314
    Selection.HorizontalAlignment = xlCenter
    Selection.Font.Size = 11
    
    Range("B30:S30").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "*будь-які співпадіння з реальністю випадкові; дані, на основі яких побудовано звіт, згенеровано авторкою"
    Range("B30:S30").Select
    Selection.Font.Italic = True
    
    Range("P17").Select
    ActiveCell.FormulaR1C1 = "Клієнт"
    Selection.Interior.Color = 13434879
    Selection.HorizontalAlignment = xlCenter
    Range("P24").Select
    ActiveCell.FormulaR1C1 = "Клієнт"
    Selection.Interior.Color = 13434879
    Selection.HorizontalAlignment = xlCenter
    Range("Q17").Select
    ActiveCell.FormulaR1C1 = "Сума незакритих договорів (грн)"
    Selection.Interior.Color = 13434879
    Selection.HorizontalAlignment = xlCenter
    Range("R17").Select
    ActiveCell.FormulaR1C1 = "Частка в загальній сумі"
    Selection.Interior.Color = 13434879
    Selection.HorizontalAlignment = xlCenter
    Range("Q24").Select
    ActiveCell.FormulaR1C1 = "Сума договорів (грн)"
    Selection.Interior.Color = 13434879
    Selection.HorizontalAlignment = xlCenter
    Range("R24").Select
    ActiveCell.FormulaR1C1 = "Частка в загальній сумі"
    Selection.Interior.Color = 13434879
    Selection.HorizontalAlignment = xlCenter
    
    Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    Selection.Borders(xlInsideVertical).Weight = xlThin
    Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Selection.Borders(xlInsideHorizontal).Weight = xlThin
    
    Sheets("Pivot3").Select
    Range("D5:F9").Select
    Selection.Copy
    Sheets("DASHBOARD").Select
    Range("P18:R22").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("Pivot4").Select
    Range("D5:F9").Select
    Selection.Copy
    Sheets("DASHBOARD").Select
    Range("P25:R29").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("P17:R22").Select
    Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlInsideHorizontal).Weight = xlThin
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    
    Range("P24:R29").Select
    Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlInsideHorizontal).Weight = xlThin
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    
    Columns("P:R").EntireColumn.AutoFit
    Columns("P:P").ColumnWidth = 35.67
    
    Range("Q18:Q22").Select
    Selection.NumberFormat = "#,##0.00"
    Range("R18:R22").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    Range("Q24:Q29").Select
    Selection.NumberFormat = "#,##0.00"
    Range("R24:R29").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    
    Range("A3:S3").Select
    Selection.Interior.ThemeColor = xlThemeColorAccent2
    Selection.Interior.TintAndShade = 0.799981688894314
    
    Range("K3:M3").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Дата побудови звіту:"
    Selection.Font.Underline = xlUnderlineStyleSingle
    Selection.HorizontalAlignment = xlRight
    
    Range("N3:O3").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Selection.Font.Underline = xlUnderlineStyleSingle
    Selection.HorizontalAlignment = xlLeft
    
    Range("M6").Select
    ActiveCell.FormulaR1C1 = "=COUNT(TREATIES!C[-12])"
    Selection.Font.Color = -16750849
    Selection.Font.Bold = True
    
    Range("M7").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(Combined!C[-6],0)"
    Selection.Font.Color = -16750849
    Selection.Font.Bold = True
    
    Range("M8").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(Combined!C[-6],0,Combined!C[-7])/R[-1]C"
    Selection.Font.Color = -16750849
    Selection.Font.Bold = True
    Selection.NumberFormat = "#,##0.00"
    
    Range("F6").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(CLIENTS!C,0)"
    Selection.Font.Color = -16750849
    Selection.Font.Bold = True
    
    Range("F7").Select
    ActiveCell.FormulaR1C1 = "=MAX(Combined!C)"
    Selection.Font.Color = -16750849
    Selection.Font.Bold = True
    Selection.NumberFormat = "#,##0.00"
    
    Range("F8").Select
    ActiveCell.FormulaR1C1 = "=SUM(TREATIES!C)/R[-2]C[7]"
    Selection.Font.Color = -16750849
    Selection.Font.Bold = True
    Selection.NumberFormat = "#,##0.00"
    
    Range("F6:G6").Select
    Selection.Merge
    Range("F7:G7").Select
    Selection.Merge
    Range("F8:G8").Select
    Selection.Merge
    Range("M6:N6").Select
    Selection.Merge
    Range("M7:N7").Select
    Selection.Merge
    Range("M8:N8").Select
    Selection.Merge
    
    Sheets("Pivot5").Select
    ActiveSheet.ChartObjects("Chart1").Activate
    ActiveChart.ChartArea.Copy
    
    Sheets("DASHBOARD").Select
    Range("P6").Select
    ActiveSheet.Paste
    ActiveSheet.ChartObjects("Chart1").Activate
    ActiveSheet.Shapes("Chart1").ScaleHeight 0.6194444444, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Chart1").ScaleWidth 1.3366666667, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Chart1").ScaleWidth 1.0037406484, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Chart1").ScaleHeight 0.9820627803, msoFalse, _
        msoScaleFromTopLeft
        
    Sheets("Pivot1").Select
    ActiveSheet.ChartObjects("Chart2").Activate
    ActiveChart.ChartArea.Copy
    
    Sheets("DASHBOARD").Select
    Range("B9").Select
    ActiveSheet.Paste
    ActiveSheet.ChartObjects("Chart2").Activate
    ActiveSheet.Shapes("Chart2").ScaleWidth 0.8683333333, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Chart2").ScaleHeight 1.2722222222, msoFalse, _
        msoScaleFromTopLeft
    
    Sheets("Pivot2").Select
    ActiveSheet.ChartObjects("Chart3").Activate
    ActiveChart.ChartArea.Copy
    
    Sheets("DASHBOARD").Select
    Range("I9").Select
    ActiveSheet.Paste
    ActiveSheet.ChartObjects("Chart3").Activate
    ActiveSheet.Shapes("Chart3").ScaleWidth 0.8333333333, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Chart3").ScaleHeight 1.2777777778, msoFalse, _
        msoScaleFromTopLeft
        
    Range("A1").Select
    
    Worksheets("Combined").Visible = False
    Worksheets("Pivot1").Visible = False
    Worksheets("Pivot2").Visible = False
    Worksheets("Pivot3").Visible = False
    Worksheets("Pivot4").Visible = False
    Worksheets("Pivot5").Visible = False

    Application.ScreenUpdating = True
    
    MsgBox "Звіт сформовано"
    
End Sub
