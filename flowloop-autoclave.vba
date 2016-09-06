Option Explicit
Sub AutoclaveDataImport()
'
' AutoclaveDataImport Macro
' Utility to import all of the files associated with an autoclave run
'

'
    Dim sttime As Double, FilePaths As Collection, filePath As String
    Dim fd As Object
    Dim DataRange As Range
    Dim Data() As Variant, PasteData() As Variant
    Dim xstep As Double, xmin As Double, xmax As Double, xval As Double, ymax As Double
    Dim i As Long, j As Long, k As Long, MULTIFILE As Boolean, CANIMPORT As Boolean
    
    Set FilePaths = New Collection
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .InitialView = msoFileDialogViewDetails
        .Title = "Select the data file(s)"
        .Filters.Add "CSV files", "*.csv", 1
        .Filters.Add "TXT files", "*.txt", 2
        .AllowMultiSelect = True
        If .Show = -1 Then
            If fd.SelectedItems.Count > 1 Then
                MULTIFILE = True
                For i = 1 To fd.SelectedItems.Count
                    ' Auto sort will work fine if filenames begin with full dates
                    ' e.g. MM-DD-YY but not M-D-YY (need 0's for single digits)
                    FilePaths.Add fd.SelectedItems(i)
                Next i
            Else
                MULTIFILE = False
                FilePaths.Add fd.SelectedItems(1)
                filePath = FilePaths(1)
            End If
        Else
            Exit Sub
        End If
    End With
    Set fd = Nothing
    
    ' Data Import steps
    sttime = Timer
    Application.ScreenUpdating = False
    i = 1: CANIMPORT = False
    ' Get blank sheet
    Do While Not CANIMPORT
        If ((Sheets(i).UsedRange.Address = vbNullString) _
            Or _
           (Sheets(i).UsedRange.Address = "$A$1" And Sheets(i).Range("A1").Value = vbNullString)) Then
            CANIMPORT = True
            Sheets(i).Activate
            Exit Do
        End If
        i = i + 1
        If i > Sheets.Count Then
            Sheets.Add After:=Sheets(Sheets.Count)
            CANIMPORT = True
        End If
    Loop
    ActiveSheet.Range("A1:AB1").Value = Array("Date", "Time", "Datetime", "HTRC", "HT1", "HT2", "HT3", "TSIN", "TS1", "TS2", "TS3", "TS4", "TS5", "TS6", "TS7", "TSX", "CW", "SMPL", "O2IN", "O2X", "COND", "LVL", "TSP", "TSC", "flow rate", "pump", "Column1", "Column2")
    If Not MULTIFILE Then
        ' Single file given
        Call ReadData(filePath, True)
    Else
        ' Multiple files given, loop over them
        filePath = FilePaths.Item(1)
        Call ReadData(filePath, True)
        For k = 2 To FilePaths.Count
            filePath = FilePaths.Item(k)
            Call ReadData(filePath, False)
        Next k
    End If
    Call DeleteTextRows(1, "Date")
    Columns("A:A").NumberFormat = "yyyy-mm-dd"
    Columns("B:B").NumberFormat = "hh:mm:ss;@"
    Columns("C:C").NumberFormat = "yyyy-mm-dd hh:mm:ss;@"
    [A1].Select
    Call CreateDateTimeColumn
    Do While ActiveWorkbook.Connections.Count > 0
        ActiveWorkbook.Connections(ActiveWorkbook.Connections.Count).Delete
    Loop
    Do While ActiveWorkbook.Names.Count > 0
        ActiveWorkbook.Names(ActiveWorkbook.Names.Count).Delete
    Loop
    Do While ActiveSheet.QueryTables.Count > 0
        ActiveSheet.QueryTables(ActiveSheet.QueryTables.Count).Delete
    Loop
    ActiveSheet.ListObjects.Add(xlSrcRange, ActiveSheet.Range(Cells(1, 1), Cells(1, 1).End(xlDown).End(xlToRight)), , xlYes) _
                         .Name = "Autoclave_Import_" & Format(Date, "mmddyy") & "_" & ActiveSheet.Name
    Columns.AutoFit
    Application.ScreenUpdating = True
    Debug.Print Timer - sttime
    End Sub
Sub ReadData(path As String, first As Boolean)
    ' Import the file(s)
    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + path, Destination _
        :=Cells(1 - first + ActiveSheet.UsedRange.Rows.Count, 1))
        .Name = "0b555749424c"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = True
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
End Sub
Sub DeleteTextRows(col As Integer, testVal As Variant)
    Dim i As Long
    
    With ActiveSheet
        For i = .UsedRange.Rows.Count To 3 Step -1
            If .Cells(i, col) = testVal Then
                .Rows(i).Delete
            End If
        Next i
    End With
End Sub
Sub CreateDateTimeColumn()
    Dim maxRow As Long
    
    maxRow = ActiveSheet.UsedRange.Rows.Count
    ActiveSheet.Range(Cells(2, 3), Cells(maxRow, 3)).FormulaR1C1 = "=R[0]C1+R[0]C2"
    Application.Calculate
    ActiveSheet.Range(Cells(2, 3), Cells(maxRow, 3)).Value = ActiveSheet.Range(Cells(2, 3), Cells(maxRow, 3)).Value
    ActiveSheet.Range("C1").Value = "Datetime"
End Sub
Sub CreatePivot()
    
End Sub
Sub CreateAutoClavePivot()
'
' CreateAutoClavePivot Macro
'

'
    Dim aSubtotals As Variant
    
    aSubtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        Sheets(2).ListObjects(1).Name, Version:=xlPivotTableVersion14). _
        CreatePivotTable TableDestination:=Sheets(1).Cells(3, 1), TableName:="PivotTable1" _
        , DefaultVersion:=xlPivotTableVersion14
    Sheets(1).Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable1")
        With .PivotFields("Datetime")
            .Orientation = xlRowField
            .Position = 1
            .Subtotals = aSubtotals
        End With
        .AddDataField ActiveSheet.PivotTables("PivotTable1").PivotFields("TSC"), "Avg of TSC", xlAverage
        .AddDataField ActiveSheet.PivotTables("PivotTable1").PivotFields("TSC"), "StdDev of TSC", xlStDev
        .AddDataField ActiveSheet.PivotTables("PivotTable1").PivotFields("TSP"), "Avg of TSP", xlAverage
        .AddDataField ActiveSheet.PivotTables("PivotTable1").PivotFields("TSP"), "StdDev of TSP", xlStDev
        .RowAxisLayout xlTabularRow
    Range("A6").Group Start:=True, End:=True, Periods:=Array(False, True, True, True, True, False, False)
        With .PivotFields("Hours")
            .PivotItems("1 PM").ShowDetail = False
            .PivotItems("2 PM").ShowDetail = False
            .PivotItems("3 PM").ShowDetail = False
            .PivotItems("4 PM").ShowDetail = False
            .PivotItems("5 PM").ShowDetail = False
            .PivotItems("6 PM").ShowDetail = False
            .PivotItems("7 PM").ShowDetail = False
            .PivotItems("8 PM").ShowDetail = False
            .PivotItems("9 PM").ShowDetail = False
            .PivotItems("10 PM").ShowDetail = False
            .PivotItems("11 PM").ShowDetail = False
            .PivotItems("12 PM").ShowDetail = False
            .PivotItems("12 AM").ShowDetail = False
            .PivotItems("1 AM").ShowDetail = False
            .PivotItems("2 AM").ShowDetail = False
            .PivotItems("3 AM").ShowDetail = False
            .PivotItems("4 AM").ShowDetail = False
            .PivotItems("5 AM").ShowDetail = False
            .PivotItems("6 AM").ShowDetail = False
            .PivotItems("7 AM").ShowDetail = False
            .PivotItems("8 AM").ShowDetail = False
            .PivotItems("9 AM").ShowDetail = False
            .PivotItems("10 AM").ShowDetail = False
            .PivotItems("11 AM").ShowDetail = False
        End With
        .ColumnGrand = False
        .RowGrand = False
        .PivotFields("Date").Subtotals = aSubtotals
        .PivotFields("Time").Subtotals = aSubtotals
        .PivotFields("HTRC").Subtotals = aSubtotals
        .PivotFields("HT1").Subtotals = aSubtotals
        .PivotFields("HT2").Subtotals = aSubtotals
        .PivotFields("HT3").Subtotals = aSubtotals
        .PivotFields("TSIN").Subtotals = aSubtotals
        .PivotFields("TS1").Subtotals = aSubtotals
        .PivotFields("TS2").Subtotals = aSubtotals
        .PivotFields("TS3").Subtotals = aSubtotals
        .PivotFields("TS4").Subtotals = aSubtotals
        .PivotFields("TS5").Subtotals = aSubtotals
        .PivotFields("TS6").Subtotals = aSubtotals
        .PivotFields("TS7").Subtotals = aSubtotals
        .PivotFields("TSX").Subtotals = aSubtotals
        .PivotFields("CW").Subtotals = aSubtotals
        .PivotFields("SMPL").Subtotals = aSubtotals
        .PivotFields("O2IN").Subtotals = aSubtotals
        .PivotFields("O2X").Subtotals = aSubtotals
        .PivotFields("COND").Subtotals = aSubtotals
        .PivotFields("LVL").Subtotals = aSubtotals
        .PivotFields("TSP").Subtotals = aSubtotals
        .PivotFields("TSC").Subtotals = aSubtotals
        .PivotFields("flow rate").Subtotals = aSubtotals
        .PivotFields("pump").Subtotals = aSubtotals
        .PivotFields("Column1").Subtotals = aSubtotals
        .PivotFields("Column2").Subtotals = aSubtotals
    End With
    Columns("D:D").EntireColumn.Hidden = True
    Columns("E:H").NumberFormat = "0.000"
    Columns.AutoFit
    With Range("E3:H3")
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
    Columns("I:I").ColumnWidth = 0.92
    
    ' insert 3 plots here at J2, Q2, and J17, all with default sizing
    
    Range("J2").Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlXYScatterLinesNoMarkers
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(1).Name = "=""Sample Temperature"""
    ActiveChart.SeriesCollection(1).Values = "=Sheet7!$E$4:$E$147"
    ActiveWindow.SmallScroll Down:=-144
    Columns("C:C").EntireColumn.AutoFit
    Range("B3").Select
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.Axes(xlValue).MajorUnit = 50
    'Selection.TickLabels.NumberFormat = "General"
    ActiveChart.Axes(xlCategory).TickLabelSpacing = 50
    ActiveChart.Legend.Delete
    ActiveChart.Axes(xlValue).MajorGridlines.Delete
    ActiveSheet.ChartObjects("Chart 1").Activate
    With ActiveChart.Axes(xlValue).Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With
    With ActiveChart.Axes(xlCategory).Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
    End With
    With ActiveChart.PlotArea.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
    End With
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes("Chart 1").IncrementLeft 288.3750393701
    ActiveSheet.Shapes("Chart 1").IncrementTop -163.5
    ActiveChart.SeriesCollection(1).HasErrorBars = True
    ActiveChart.SeriesCollection(1).ErrorBar Direction:=xlY, Include:=xlBoth, _
        Type:=xlErrorBarTypeCustom, Amount:="=Sheet7!$F$4:$F$147", MinusValues:="=Sheet7!$F$4:$F$147"
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveSheet.Shapes("Chart 1").Line.Visible = msoFalse
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.ChartArea.Copy
    Range("J17").Paste
    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.SeriesCollection(1).Name = "Sample Pressure"
    ActiveChart.SeriesCollection(1).Values = "=Sheet7!$G$4:$G$147"
    ActiveChart.SeriesCollection(1).ErrorBar Direction:=xlY, Include:=xlBoth, _
        Type:=xlErrorBarTypeCustom, Amount:="=Sheet7!$H$4:$H$147", MinusValues:="=Sheet7!$H$4:$H$147"
    ActiveChart.Axes(xlValue).Select
    ActiveSheet.ChartObjects("Chart 2").Activate
    
End Sub
