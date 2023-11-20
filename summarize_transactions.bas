Attribute VB_Name = "Module1"
Sub transformSheet()

'
' Keyboard Shortcut: Ctrl+Shift+O
'
    Dim sheetName As String
    Dim pf As PivotField
    
' 1) Rename sheet to a shorter version

    sheetName = Replace(ActiveSheet.Name, "-TradeActivity", "")
    sheetName = Replace(sheetName, "2023-", "")
    sheetName = Replace(sheetName, "-", "_")
    ActiveSheet.Name = sheetName
    sheetName = "head_" & sheetName
    
' 2) Go to the cell adjacent to the last column header to the left and a new column header "Amount"

    Cells.Find(What:="Filled Orders", After:=ActiveCell, LookIn:=xlFormulas2 _
    , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext).Select

    ' 2.a) Add a range name that will be used later on
    ActiveWorkbook.Names.Add Name:=sheetName, RefersTo:=Selection
    ActiveCell.Offset(1, 14).Select

    ActiveCell.FormulaR1C1 = "Amount"

' 3) Add formula to calculate transaction amounts. Normally only transaction price and quantity are included
' and apply the formula down to the last row

    ActiveCell.Offset(1, 0).Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-2]*RC[-9]"

    Selection.AutoFill Destination:=Range(Selection, Selection.End(xlDown))
    Range(sheetName).Offset(1, 4).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select

' 4) Add a PivotTable that summarizes the transactions for the day
'    Selection.Offset(0, 2).Select
'
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        Selection, Version:=8).CreatePivotTable _
        TableDestination:=ActiveCell.Offset(0, 12), TableName:= _
        sheetName, DefaultVersion:=8

    With ActiveSheet.PivotTables(sheetName)
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables(sheetName).PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables(sheetName).RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables(sheetName).PivotFields("Symbol")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables(sheetName).PivotFields("Side")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables(sheetName).AddDataField ActiveSheet.PivotTables( _
        sheetName).PivotFields("Qty"), "Sum of Qty", xlSum
    'ActiveSheet.PivotTables(sheetName).AddDataField ActiveSheet.PivotTables( _
    '    sheetName).PivotFields("Net Price"), "Sum of Net Price", xlSum
    ActiveSheet.PivotTables(sheetName).AddDataField ActiveSheet.PivotTables( _
        sheetName).PivotFields("Amount"), "Sum of Amount", xlSum
    ActiveSheet.PivotTables(sheetName).CalculatedFields.Add "Calc Ave", _
        "=Amount / ABS(Qty)", True
    ActiveSheet.PivotTables(sheetName).PivotFields("Calc Ave").Orientation = _
        xlDataField

'5) Apply suitable formatting

    'Changes formatting of Value fields
    For Each pf In ActiveSheet.PivotTables(sheetName).DataFields
        pf.NumberFormat = "#,##0.00_);[Red](#,##0.00)"
    Next pf

''6) Copy-paste values of the pivot table -- can be omitted for now
'
'    ActiveCell.Offset(0, 13).Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Range(Selection, Selection.End(xlToRight)).Select
'    Selection.Copy
'    ActiveCell.Offset(0, 3).Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False

'7)

    ActiveSheet.Range(sheetName).Offset(2, 20).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[-2]C=""CALC"",MIN(R[-1]C[-3],-RC[-3]),IF(R[-1]C=""CALC"",RC[-1]+R[1]C[-1],IF(AND(R[1]C[-4]=""BUY"",R[2]C[-4]=""SELL""),""CALC"","""")))"
    Selection.AutoFill Destination:=Range(Selection, Selection.Offset(100, 0))
    ActiveCell.Offset(0, 1).Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]=""CALC"",R[2]C[-1]*-R[1]C[-1],"""")"
    Selection.AutoFill Destination:=Range(Selection, Selection.Offset(100, 0))
    ActiveCell.Offset(-3, 0).Select
    ActiveCell.FormulaR1C1 = "=SUM(R[4]C:R[100]C)"
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Sub


Sub addToSummary()
'
' addToSummary Macro
'
' Keyboard Shortcut: Ctrl+Shift+W

    Dim sheetName, cellAddress As String
    
    sheetName = ActiveSheet.Name
    
    Cells.Find(What:="Filled Orders", After:=ActiveCell, LookIn:=xlFormulas2 _
    , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext).Select
    ActiveCell.Offset(2, 2).Select
    Range(Selection, Selection.Offset(0, 12)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Summary").Select
    Range("A5").Select
    Selection.End(xlDown).Select
    cellAddress = ActiveCell.Address
    If cellAddress = "$A$1048576" Then
    Range("A6").Select
    Else
    Selection.Offset(1, 0).Select
    End If


    ActiveSheet.Paste
    ActiveSheet.Range("A6").Select
   
End Sub

Sub updateSummary()
'
' Macro4 Macro
'
' Keyboard Shortcut: Ctrl+Shift+Q
'
    Dim pf As PivotField
    Sheets("Summary").Columns("O:X").Clear
    Sheets("Summary").Range("A5").Select
    Range(Selection, Selection.Offset(0, 12)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        Selection, Version:=8).CreatePivotTable _
        TableDestination:=ActiveCell.Offset(0, 14), TableName:= _
        "Transactions Summary", DefaultVersion:=8

    With ActiveSheet.PivotTables("Transactions Summary")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("Transactions Summary").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("Transactions Summary").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("Transactions Summary").PivotFields("Symbol")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Transactions Summary").PivotFields("Side")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("Transactions Summary").AddDataField ActiveSheet.PivotTables( _
        "Transactions Summary").PivotFields("Qty"), "Sum of Qty", xlSum
    ActiveSheet.PivotTables("Transactions Summary").AddDataField ActiveSheet.PivotTables( _
        "Transactions Summary").PivotFields("Net Price"), "Sum of Net Price", xlSum
    ActiveSheet.PivotTables("Transactions Summary").AddDataField ActiveSheet.PivotTables( _
        "Transactions Summary").PivotFields("Amount"), "Sum of Amount", xlSum
    With ActiveSheet.PivotTables("Transactions Summary").PivotFields("Sum of Net Price")
        .Caption = "Average of Net Price"
        .Function = xlAverage
    End With
    
    'Changes formatting of Value fields
    For Each pf In ActiveSheet.PivotTables("Transactions Summary").DataFields
        pf.NumberFormat = "#,##0.00_);[Red](#,##0.00)"
    Next pf
    
    Range("O5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("S5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("W1").Select
    ActiveCell.FormulaR1C1 = "=MIN(R[1]C[-3],-R[2]C[-3])"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "=(R[1]C[-5]/R[1]C[-7])-(RC[-5]/RC[-7])"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]*RC[-1]"
    Selection.Interior.ColorIndex = 6
    Selection.Offset(2, 0).FormulaR1C1 = "=(Sum(R[2]C:R[306]C))"

End Sub




