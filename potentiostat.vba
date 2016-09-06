Private Function GetBlankSheet()
' Converted to private function 11/21/2012 by Ben Hauch
' Returns an empty sheet object

' Calling code usage:
' Set ____ = GetBlankSheet

    Dim i As Integer, _
    blCanImport As Boolean, _
    ojSheet As Worksheet
    
    i = 1: CANIMPORT = False
    Do While Not CANIMPORT
        If TypeName(Sheets(i)) <> "Chart" Then
            If ((Sheets(i).UsedRange.Address = vbNullString) _
                Or _
               (Sheets(i).UsedRange.Address = "$A$1" And Sheets(i).Range("A1").Value = vbNullString)) Then
                CANIMPORT = True
                Set ojSheet = Sheets(i)
                ojSheet.Activate
                Exit Do
            End If
        End If
        i = i + 1
        If i > Sheets.Count Then
            Set ojSheet = Sheets.Add(After:=Sheets(Sheets.Count))
            CANIMPORT = True
        End If
    Loop
    Set GetBlankSheet = ojSheet
End Function
Sub PStat_Import()
' Written 9/24/2012 by Ben Hauch
' Last updated 11/21/2012 by Ben Hauch
' Load in a .cor file from the Solartron Potentiostat

Dim SS As Worksheet
Dim fileName As String, _
    fd As Object, _
    iCommentOffset As Integer, _
    sScantype As String, _
    iScantypeOffset As Integer, _
    vVoltData As Variant, _
    vTimeData As Variant, _
    vCurrData As Variant, _
    vThresholdData() As Variant
    
' NEEDS VERIFICATION:
'   Verification: multiple lines of comments will appear in column E as well, pushing those items down

' NEEDS IMPLEMENTATION:
'   Context-based plotting
'   Automatic averaging/stability of Open Circuit type files

' AND THEN DO FOR LOOPS TO PROPERLY DO LTRIM()
' Works for OC files
' Need to set up context-dependent plotting (e.g. Evs I axes or linear axes, etc)

    ' Require a blank sheet to do data import
    Set SS = GetBlankSheet

    ' File browser for the input datafile
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .InitialView = msoFileDialogViewDetails
        .Title = "Select the CorrWare .cor ASCII datafile"
        .AllowMultiSelect = False
        .Filters.Add "COR Files", "*.cor", 1
        .Show
    End With
    fileName = fd.SelectedItems.Item(1)
    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & fileName & "", Destination:=Range("$A$1"))
        .Name = fileName
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1252
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileOtherDelimiter = ":"
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    If ActiveWorkbook.Connections.Count > 0 Then ActiveWorkbook.Connections(1).Delete
    With SS
    ' Get the number of comments in the file
        iCommentOffset = .Range("B42").Value
    ' Clean up file import formatting
        .Range("B4").Value = Left(LTrim(.Range("B4").Value), 10) ' Expt Date
        ' Experiment time
        .Range("D4").Value = .Range("C4").Value & ":" & .Range("D4").Value & ":" & .Range("E4").Value
        .Range("C4").Value = "Time"
        ' Remove extraneous date/time values & legacy formatting
        .Columns("E:H").Delete Shift:=xlToLeft
        .Columns("B:B").ColumnWidth = 16.86
        ' Remove prefixed spaces around numbers, titles
        .Range("B8").Value = LTrim(.Range("B8").Value)
        .Range("B19").Value = LTrim(.Range("B19").Value)
        .Range("B21").Value = LTrim(.Range("B21").Value)
        ' Experiment detail type
        .Range("B39").Value = LTrim(.Range("B39").Value) & ":" & .Range("C39").Value
        ' File Path
        .Range("B40").Value = LTrim(.Range("B40")) & ":" & .Range("C40").Value
        If iCommentOffset > 0 Then .Cells(42 + iCommentOffset, 2).Value = LTrim(.Cells(42 + iCommentOffset, 2).Value)
    ' Depending on the scan, there are different data positions (also on number of comments)
        sScantype = LTrim(.Range("A3").Value)
        Select Case sScantype
            Case "Open Circuit", "Galvanostatic"
                iScantypeOffset = 0
            Case "Potentiostatic"
                iScantypeOffset = 1
                
            Case "Potentiodynamic"
                iScantypeOffset = 9
                
            Case "Potential Stair-Step"
                iScantypeOffset = 8
                
            Case Else
                MsgBox Prompt:="No case for scan type" & .Range("A3").Value
                Exit Sub
                ' Could launch a pickwindow instructing the user to identify the start cell and use the .Row property to compute iScantypeOffset
        End Select
    ' Write the detailed Experiment info
        .Cells(59 + iCommentOffset + iScantypeOffset, 2).Value = LTrim(.Cells(59 + iCommentOffset + iScantypeOffset, 2).Value) & ":" & .Cells(59 + iCommentOffset + iScantypeOffset, 3).Value
    ' Remove old info sources
        .Range("C39:D39,C40").ClearContents
        .Range(Cells(59 + iCommentOffset + iScantypeOffset, 3), Cells(59 + iCommentOffset + iScantypeOffset, 4)).ClearContents
    ' Move Experiment info to make room for plot
        .Range(Cells(39, 1), Cells(59 + iCommentOffset + iScantypeOffset, 2)).Cut Destination:=.Range(Cells(7, 5), Cells(27 + iCommentOffset + iScantypeOffset, 6))
    ' Move Potentiostat info to make room for plot
        .Range("A20:B38").Cut Destination:=.Range(Cells(29 + iCommentOffset + iScantypeOffset, 5), Cells(29 + iCommentOffset + iScantypeOffset + 18, 6))
    ' Move Corrosion cell info to make room for plot
        .Range("A7:B19").Cut Destination:=.Range(Cells(49 + iCommentOffset + iScantypeOffset, 5), Cells(49 + iCommentOffset + iScantypeOffset + 12, 6))
    ' Move the experiment data up to square off the plot area
        .Range(Cells(26, 1), Cells(58 + iCommentOffset + iScantypeOffset, 3)).Delete Shift:=xlUp
    ' Set Widths, etc
        .Columns("A:A").ColumnWidth = 22.43
        .Columns("C:C").ColumnWidth = 11.14
        .Columns("D:D").EntireColumn.AutoFit
        .Columns("E:E").EntireColumn.AutoFit
    ' Contextual Plot
        PStatPlottype.Show
        Select Case PStatPlottype.Tag
            Case "axesEvsTime"
                ' Want to plot E vs. Time
                .Range("A28:D28").Value = Array("E [V]", "I [A/cm²]", "T [sec]", "T [hr]")
                .Range(Cells(30, 4), Cells(Cells(30, 2).End(xlDown).Row, 4)).FormulaR1C1 = "=RC[-1]/3600"
                ' Plot format code
                .Range("A10").Select
                .Shapes.AddChart(XlChartType:=xlXYScatterLinesNoMarkers, Left:=1, Top:=90, Width:=320, Height:=290).Select
                With ActiveChart
                    .SeriesCollection.NewSeries
                    With .SeriesCollection(1)
                        .XValues = SS.Range(Cells(30, 4), Cells(30, 4).End(xlDown))
                        .Values = SS.Range(Cells(30, 1), Cells(30, 1).End(xlDown))
                        .Name = SS.Range("C3").Value
                        .Format.Line.Weight = 2
                    End With
                    With .Axes(xlValue, xlPrimary)
                        .MajorGridlines.Delete
                        .HasTitle = True
                        .Format.Line.Visible = msoCTrue
                        .AxisTitle.Text = "Potential [V]"
                        .CrossesAt = -9999
                    End With
                    With .Axes(xlCategory)
                        .MinimumScale = 0
                        .HasTitle = True
                        .MinorTickMark = xlTickMarkInside
                        .Format.Line.Visible = msoCTrue
                        .AxisTitle.Text = "Time [hr]"
                        .TickLabels.NumberFormat = "#,##0.00"
                    End With
                    With .PlotArea.Format.Line
                        .Visible = msoCTrue
                        .Style = msoLineSingle
                    End With
                    .Legend.Delete
                End With
            Case "axesEIvsTime"
                ' Want to plot E vs. Time, and I vs time to verify constancy/disruptions
                .Range("A28:D28").Value = Array("E [V]", "I [A/cm²]", "T [sec]", "T [hr]")
                .Range(Cells(30, 4), Cells(Cells(30, 2).End(xlDown).Row, 4)).FormulaR1C1 = "=RC[-1]/3600"
                ' Plot format code
                .Range("A10").Select
                .Shapes.AddChart(XlChartType:=xlXYScatterLinesNoMarkers, Left:=1, Top:=90, Width:=320, Height:=290).Select
                With ActiveChart
                    .SeriesCollection.NewSeries
                    With .SeriesCollection(1)
                        .XValues = SS.Range(Cells(30, 4), Cells(30, 4).End(xlDown))
                        .Values = SS.Range(Cells(30, 1), Cells(30, 1).End(xlDown))
                        .Name = "Potential [V]"
                        .Format.Line.Weight = 2
                    End With
                    .SeriesCollection.NewSeries
                    With .SeriesCollection(2)
                        .AxisGroup = xlSecondary
                        .XValues = SS.Range(Cells(30, 4), Cells(30, 4).End(xlDown))
                        .Values = SS.Range(Cells(30, 2), Cells(30, 2).End(xlDown))
                        .Name = "Current [A/cm²]"
                        .Format.Line.Weight = 2
                    End With
                    With .Axes(xlValue, xlPrimary)
                        .MajorGridlines.Delete
                        .HasTitle = True
                        .Format.Line.Visible = msoCTrue
                        .AxisTitle.Text = "Potential [V]"
                        .CrossesAt = -9999
                    End With
                    With .Axes(xlValue, xlSecondary)
                        .MajorGridlines.Delete
                        .HasTitle = True
                        .Format.Line.Visible = msoCTrue
                        .AxisTitle.Text = "Current [A/cm²]"
                    End With
                    .Legend.Position = xlLegendPositionTop
                    With .Axes(xlCategory)
                        .MinimumScale = 0
                        .HasTitle = True
                        .MinorTickMark = xlTickMarkInside
                        .Format.Line.Visible = msoCTrue
                        .AxisTitle.Text = "Time [hr]"
                        .TickLabels.NumberFormat = "#,##0.00"
                    End With
                    With .PlotArea.Format.Line
                        .Visible = msoCTrue
                        .Style = msoLineSingle
                    End With
                    .HasTitle = True
                    .ChartTitle.Text = SS.Range("C3").Value
                End With
            Case "axesIvsTime"
                ' Want to plot I vs. Time
                .Range("A28:D28").Value = Array("E [V]", "I [A/cm²]", "T [sec]", "T [hr]")
                .Range(Cells(30, 4), Cells(Cells(30, 2).End(xlDown).Row, 4)).FormulaR1C1 = "=RC[-1]/3600"
                ' Plot format code
                .Range("A10").Select
                .Shapes.AddChart(XlChartType:=xlXYScatterLinesNoMarkers, Left:=1, Top:=90, Width:=320, Height:=290).Select
                With ActiveChart
                    .SeriesCollection.NewSeries
                    With .SeriesCollection(1)
                        .XValues = SS.Range(Cells(30, 4), Cells(30, 4).End(xlDown))
                        .Values = SS.Range(Cells(30, 2), Cells(30, 2).End(xlDown))
                        .Name = SS.Range("C3").Value
                        .Format.Line.Weight = 2
                    End With
                    With .Axes(xlValue, xlPrimary)
                        .MajorGridlines.Delete
                        .HasTitle = True
                        .Format.Line.Visible = msoCTrue
                        .AxisTitle.Text = "Current [A/cm²]"
                        .CrossesAt = -9999
                    End With
                    With .Axes(xlCategory)
                        .MinimumScale = 0
                        .HasTitle = True
                        .MinorTickMark = xlTickMarkInside
                        .Format.Line.Visible = msoCTrue
                        .AxisTitle.Text = "Time [hr]"
                        .TickLabels.NumberFormat = "#,##0.00"
                    End With
                    With .PlotArea.Format.Line
                        .Visible = msoCTrue
                        .Style = msoLineSingle
                    End With
                    .Legend.Delete
                End With
            Case "axesEvslogI"
                ' Want to plot E vs log(abs(I))
                .Range("A28:D28").Value = Array("E [V]", "I [A/cm²]", "T [sec]", "Abs(I) [A/cm²]")
                .Range(Cells(30, 4), Cells(Cells(30, 2).End(xlDown).Row, 4)).FormulaR1C1 = "=ABS(RC[-2])"
                ' Plot format code
                .Range("A10").Select
                .Shapes.AddChart(XlChartType:=xlXYScatterLinesNoMarkers, Left:=1, Top:=90, Width:=320, Height:=290).Select
                With ActiveChart
                    .SeriesCollection.NewSeries
                    With .SeriesCollection(1)
                        .XValues = SS.Range(Cells(30, 4), Cells(30, 4).End(xlDown))
                        .Values = SS.Range(Cells(30, 1), Cells(30, 1).End(xlDown))
                        .Name = "Forward"
                        .Format.Line.Weight = 2
                    End With
                    .SeriesCollection.NewSeries
                    With .SeriesCollection(2)
                        .XValues = SS.Range(Cells(30, 4), Cells(30, 4).End(xlDown))
                        .Values = SS.Range(Cells(30, 1), Cells(30, 1).End(xlDown))
                        .Name = "Reverse"
                        .Format.Line.Weight = 2
                    End With
                    With .Axes(xlValue, xlPrimary)
                        .MajorGridlines.Delete
                        .HasTitle = True
                        .Format.Line.Visible = msoCTrue
                        .AxisTitle.Text = "Potential [V]"
                        .CrossesAt = -9999
                    End With
                    With .Axes(xlCategory)
                        .HasTitle = True
                        .ScaleType = xlScaleLogarithmic
                        .MinorTickMark = xlTickMarkInside
                        .Format.Line.Visible = msoCTrue
                        .AxisTitle.Text = "Current [A/cm²]"
                        .Crosses = xlMinimum
                        .TickLabels.NumberFormat = "0.00E+00"
                    End With
                    With .PlotArea.Format.Line
                        .Visible = msoCTrue
                        .Style = msoLineSingle
                    End With
                    .Legend.IncludeInLayout = False
                End With
                
            Case "axesEvsI"
                ' Want to plot E vs I
                
        End Select
        Unload PStatPlottype
        Select Case sScantype
            Case "Open Circuit"
                ' Perform automatic averaging of potential from 1h -> end
                vCurrData = SS.Range("B27").Value + 29
                SS.Range(Cells(5 + iCommentOffset, 1), Cells(5 + iCommentOffset, 2)).Value = Array("t > 1hr Avg([V])", "±")
                SS.Cells(6 + iCommentOffset, 1).Formula = _
                    "=Averageif($C$30:$C$" & vCurrData & "," & Chr(34) & ">3600" & Chr(34) & ",$A$30:$A$" & vCurrData & ")"
                For Each itm In SS.Range("$C$30:$C$100000")
                    If itm.Value >= 3600 Then
                        Set vVoltData = itm
                        Exit For
                    End If
                Next itm
                SS.Cells(6 + iCommentOffset, 2).Formula = _
                    "=STDEV($A$" & vVoltData.Row & ":$A$" & vCurrData & ")"
        End Select
    End With
End Sub
Sub PStat_Plot()
Dim SS
Dim Sheets2plot As New Collection
Dim DataArr As Variant
Dim UnitsCol As Variant
Dim Sname As String
Dim j As Long
Dim i As Long
Set SS = Sheets.Add(After:=Sheets(Sheets.Count))
Sname = "Cmbd"
j = 0
Do While WorksheetExists(Sname)
    j = j + 1
    Sname = "Cmbd" & " " & j
Loop
SS.Name = Sname
' Get Sheets to plot
For i = 1 To Sheets.Count
    If Sheets(i).Range("A1").Value = "CORRW ASCII" Then Sheets2plot.Add i, Key:="I" & Sheets2plot.Count + 1
Next i
If Sheets2plot.Count = 0 Then
    MsgBox Prompt:="No data to plot", Buttons:=vbOKOnly
Else
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    SS.Range("A1:E1").Value = Array("Potential [V]", "I [A/cm²]", "Time [sec]", "I [mA/cm²]", "Time [hr]")
    ' Data copy code
    With Sheets(Sheets2plot.Item("I1"))
        .Activate
        DataArr = .Range(.Range("A20").End(xlDown).Offset(3, 0), .Range("A20").End(xlDown).Offset(4, 2).End(xlDown)).Value
    End With
    With SS
        .Activate
        .Range(Cells(2, 1), Cells(1 + UBound(DataArr, 1), 3)).Value = DataArr
    End With
    DataArr = vbNullString
    If Sheets2plot.Count > 1 Then
        For j = 2 To Sheets2plot.Count
            With Sheets(Sheets2plot.Item("I" & j))
                .Activate
                DataArr = .Range(.Range("A20").End(xlDown).Offset(3, 0), .Range("A20").End(xlDown).Offset(4, 2).End(xlDown)).Value
            End With
            With SS
                .Activate
                .Range(Cells(2, 1).End(xlDown).Offset(1, 0), Cells(Cells(2, 1).End(xlDown).Row + UBound(DataArr, 1), 3)).Value = DataArr
            End With
            DataArr = vbNullString
        Next j
    End If
    ' Units formatting
    With SS
        UnitsCol = .Range(.Cells(2, 1), .Cells(2, 1).End(xlDown).End(xlToRight)).Value
        ReDim DataArr(1 To UBound(UnitsCol, 1), 1 To 2)
        DataArr(1, 1) = UnitsCol(1, 2) * 1000
        DataArr(1, 2) = UnitsCol(1, 3) / 3600
        For i = 2 To UBound(DataArr)
            DataArr(i, 1) = UnitsCol(i, 2) * 1000 'Convert to mA
            ' Incremental time over sequential experiments, in hours
            If UnitsCol(i, 3) = 0 Then
                DataArr(i, 2) = DataArr(i - 1, 2)
            Else
                DataArr(i, 2) = (UnitsCol(i, 3) - UnitsCol(i - 1, 3)) / 3600 + DataArr(i - 1, 2)
            End If
        Next i
        .Range(Cells(2, 4), Cells(1 + UBound(DataArr), 5)).Value = DataArr
        .Columns("A:A").Cut Destination:=.Columns("F:F")
        .Columns("D:D").Cut Destination:=.Columns("G:G")
        .Columns("A:D").Delete
        ' Plot format code
        .Range("A1").Select
        .Shapes.AddChart(XlChartType:=xlXYScatterLinesNoMarkers, Left:=(Range("D1").Left + Range("E1").Left) / 2, Top:=0, Width:=Range("I1").Left, Height:=Range("G20").Top).Select
        With ActiveChart
            With .SeriesCollection(2)
                .AxisGroup = 2
            End With
            With .Axes(xlValue, xlPrimary)
                .MajorGridlines.Delete
                .HasTitle = True
                .Format.Line.Visible = msoCTrue
                .AxisTitle.Text = "Potential (V,SCE)"
            End With
            With .Axes(xlValue, xlSecondary)
                .HasTitle = True
                .Format.Line.Visible = msoCTrue
                .AxisTitle.Text = "Current (mA/cm²)"
            End With
            .Legend.Position = xlLegendPositionTop
            With .Axes(xlCategory)
                .MinimumScale = 0
                .HasTitle = True
                .Format.Line.Visible = msoCTrue
                .AxisTitle.Text = "Time (hr)"
            End With
            With .PlotArea.Format.Line
                .Visible = msoCTrue
                .Style = msoLineSingle
            End With
        End With
    End With
End If
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Set Sheets2plot = Nothing
End Sub
Private Function WorksheetExists(ByVal sheetName As String) As Boolean
' Determine if a Worksheet with that name exists in the workbook
' Found at MrExcel.com: http://www.mrexcel.com/forum/showthread.php?t=3228
    On Error Resume Next
    WorksheetExists = (Sheets(sheetName).Name <> "")
    On Error GoTo 0
End Function

