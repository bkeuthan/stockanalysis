Sub VBA_SCRIPT()


'worksheet loop
For Each ws In Worksheets
        ws.Activate
        
'grab last row
Dim lastRow As Long

lastRow = Cells(Rows.Count, 1).End(xlUp).Row

'add unique ID for vlookup Ticker_YYYYMMDD
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2").Select
    Application.CutCopyMode = False
   
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=RC[1]&""_""&RC[2]"
    Range("A2").Select
    Selection.AutoFill Destination:=Range("A2:A" & lastRow & "")
    
'copy all tickers & remove duplicates
    Columns("B:B").Select
    Selection.Copy
    Columns("J:J").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.Range("$J$1:$J$" & lastRow & "").RemoveDuplicates Columns:=1, Header:= _
        xlNo
        
'apply vlookup for Yearly Change = EOY close - SOY open.
    Range("K2").Select
    ActiveCell.Formula2R1C1 = _
        "=IF(RC[-1]="""","""",VLOOKUP(RC10&""_""&MAX(VALUE(R2C3:R365C3)),C1:C8,7,0)-VLOOKUP(RC10&""_""&MIN(VALUE(R2C3:R365C3)),C1:C8,4,0))"
    Range("K2").Select
    
'apply vlookup for Percent Change = EOY close / SOY open.
    Range("L2").Select
    ActiveCell.Formula2R1C1 = _
        "=IF(RC[-2]="""","""",VLOOKUP(RC10&""_""&MAX(VALUE(R2C3:R365C3)),C1:C8,7,0)/VLOOKUP(RC10&""_""&MIN(VALUE(R2C3:R365C3)),C1:C8,4,0)-1)"
    Range("L2").Select

'apply sumif formula by unique ticker
    Range("M2").Select
    ActiveCell.Formula2R1C1 = _
        "=IF(RC[-3]="""","""",SUMIF(C[-11],RC[-3],C[-5]))"
    Range("M2").Select
    
'grab last row of unique tickers
Dim lastRow2 As Long

lastRow2 = Cells(Rows.Count, 10).End(xlUp).Row
    
'autofill formulas
    Range("K2:M2").Select
    Selection.AutoFill Destination:=Range("K2:M" & lastRow2 & "")
    
'apply Greatest % increase ticker
    Range("Q2").Select
    ActiveCell.Formula2R1C1 = "=INDEX(C10,MATCH(R2C18,C12,0))"
        Range("R2").Select
        ActiveCell.FormulaR1C1 = "=MAX(C12)"
    
'apply greatest % decrease ticker
    Range("Q3").Select
    ActiveCell.Formula2R1C1 = "=INDEX(C10,MATCH(R3C18,C12,0))"
        Range("R3").Select
        ActiveCell.FormulaR1C1 = "=MIN(C12)"

'apply greatest total volume ticker
    Range("Q4").Select
    ActiveCell.Formula2R1C1 = "=INDEX(C10,MATCH(R4C18,C13,0))"
        Range("R4").Select
        ActiveCell.FormulaR1C1 = "=MAX(C13)"
        
'paste values and remove vlookup
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
   
'define headers
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percentage Change"
    Range("L1") = "Total Stock Volume"
    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest % Decrease"
    Range("O4") = "Greatest Total Volume"
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    Cells.Select
    Cells.EntireColumn.AutoFit
    
'Conditional formatting
    Range("J2:J" & lastRow2 & "").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    
'Formatting
        Columns("K:K").Select
    Selection.NumberFormat = "0.00%"
    Range("Q2:Q3").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    Cells.Select
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

'Next Worksheet in loop
    Next ws
    
End Sub