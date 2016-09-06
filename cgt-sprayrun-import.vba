Option Explicit
Sub ColdSprayRpts()
' Written by Ben Hauch for Excel07+ in June 2011
' Last updated by Ben Hauch 5/07/2014
'
' This macro will split apart the imported csv data and place
' each run on a separate sheet to facilitate closer analysis
'

' General Structure
    ' read in used area as array
    ' comb through Time column looking for a difference of more than 10s
    ' flag that section as requiring a new sheet
    ' paste relevant sections to new sheet
    ' process each run to highlight noteworthy statistics
If Right(ActiveWorkbook.Name, 3) = "csv" Then
    MsgBox Prompt:="Run from a blank workbook"
    Exit Sub
End If
Dim SS As Excel.Worksheet
Dim rownum As Integer                                   ' outer step counter
Dim rg As Range                                         'for pasting
Dim nssname As String
Dim rawdata() As Variant                                'dynamic array for copying
Dim rowstart As Integer, rowend As Integer              'for splitting
Dim NRows As Integer
Dim i As Integer                                        ' inner step counter
Dim startflag As Boolean
Dim GasSetpointVal As Double
Dim GasSetpointDev As Double
Dim gasSteadyStateAchieved As Boolean

' File browser for the input datafile
Dim fileName$
Dim fd As Object
If Sheets(1).Range("B10").Value = vbNullString Then
    ' Require a blank sheet to do data import
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .Filters.Add "CSV Files", "*.csv", 1
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
        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
End If
If ActiveWorkbook.Connections.Count > 0 Then ActiveWorkbook.Connections(1).Delete        ' Remove the connection to the file

Application.ScreenUpdating = False
If Not WorksheetExists("AllData") Then      ' Have we already created AllData?
    Set SS = Sheets(1)                          ' No we haven't
    Application.DisplayAlerts = False
'    Sheets("Sheet2").Delete
'    Sheets("Sheet3").Delete
    Application.DisplayAlerts = True
Else
    Set SS = Sheets("AllData")                  ' Yes we have
End If
rawdata = SS.UsedRange.Value                ' Import the whole dataset to a VBA array
rawdata(1, 1) = "El. Time (s)"              ' Overwrite some obnoxious/junk header names
rawdata(1, 5) = "T_Gun_Out"                 ' with more readable/understandable text
rawdata(1, 6) = "T_Gun_In"
rawdata(1, 7) = "T_PHeat_Out"
rawdata(1, 8) = "T_PHeat_In"

SS.Name = "AllData"
Set SS = Nothing                            ' All done processing the AllData sheet
rowstart = 2
For rownum = 3 To UBound(rawdata, 1)        ' Iterate through all rows of the VBA array (1 = header row)
    If (rownum = UBound(rawdata, 1) Or _
       (rawdata(rownum, 2) - rawdata(rownum - 1, 2) > (10 / 86400))) Then
        ' Exit conditions are
            ' Reached the end of the VBA array (no more rows to process)
            ' Time difference between this row and previous row > 10s (data logging restarted)
            
        rowend = rownum                     ' Store the row that triggered execution
        NRows = rowend - rowstart + 1       ' Determine size of data to transfer
        
        ' Confirm that this is a single experiment and not a "continuous" experiment by analyzing
        ' the heater and gas flow presets
            ' Cannot detect differences in gun offset, since that is a robot parameter and not a
            ' Cold Spray Controller parameter
            ' Gas Flow setpoint is an 'exact' parameter, whereas Heater setpoint fluctuates, and
            ' thus we use a 10-pt floating average
        If (rawdata(rowend, 2) - rawdata(rowstart, 2) >= 120 / 86400) Then
            ' At least two minutes worth of data collected in this rawdata() section
            gasSteadyStateAchieved = False
            For i = rowstart + 10 To rowend     ' Use i since rownum is used in the outer loop
                ' Compute flow setpoint rolling average
                GasSetpointVal = (rawdata(i - 1, 12) + rawdata(i - 2, 12) + rawdata(i - 3, 12) + _
                  rawdata(i - 4, 12) + rawdata(i - 5, 12) + rawdata(i - 6, 12) + rawdata(i - 7, 12) + _
                  rawdata(i - 8, 12) + rawdata(i - 9, 12) + rawdata(i - 10, 12)) / 10
                
                ' Compute the current deviation between this row's value and the average
                GasSetpointDev = rawdata(i, 12) - GasSetpointVal
                If gasSteadyStateAchieved Then
                    ' Have already achieved steady-state on the flow side of things
                    If GasSetpointDev <> 0 Then
                        ' Transition between pressure parameters detected
                    End If
                Else
                    ' Nothing to be done -- still have gas pressure steady-state
                End If
                ' If deviation is 0, set pressure SS flag (exclude preflow value's setpoint)
                If ((GasSetpointDev = 0) And (GasSetpointVal <> 2.631579)) Then gasSteadyStateAchieved = True
                
                ' If flow setpoint deviates from the first steadystate, create new experiment from
                    ' the first part of the triggered data
                ' or if heater setpoint deviates from the first steadystate, create new experiment
                    ' from the first part of the triggered data
                '
                
            Next i
        
            ' Down to one experiment now, so put it on its own sheet
            Set SS = Sheets.Add(After:=Sheets(Sheets.Count))    ' Create sheet for experiment
            SS.Range(Cells(1, 1), Cells(UBound(rawdata, 1), UBound(rawdata, 2))).Value = rawdata
        
        
            ' Delete all experiments not related to this one
            SS.Rows("" & rownum & ":" & UBound(rawdata, 1) & "").Delete ' First delete trailing rows
            If Not rowstart = 1 Then
                SS.Rows("2:" & rowstart & "").Delete                    ' Then preceding rows (leave headers)
            End If
        Else
            ' Less than two minutes of data -- this is not a real experiment, likely an error
            ' Thus do not create a new sheet
        End If
        ' Update rowstart and continue parsing timestamps in rawdata() to find
        ' new experiments/multi-experiments
        rowstart = rownum - 1
        Set SS = Nothing
    End If
Next rownum

' Miscellaneous post-processing steps to operate on confirmed experiments
For i = 1 To Sheets.Count - 1
    startflag = False
    Set SS = Sheets(i + 1)
    SS.Columns("B").NumberFormat = "h:mm:ss AM/PM"
    SS.Name = "Expt" & i
    SS.Rows("1:10").Insert
    Set rg = SS.UsedRange.SpecialCells(xlCellTypeLastCell)
    SS.Range("A1:A10").Value = Application.Transpose(Array("Sample Name", "", "Start Time", "End Time", "", "Steady State", "Stop Pushed", "", "Row Start", "Row Stop"))
    SS.Range("B1:B4").Value = Application.Transpose(Array("", "", SS.Range("B12"), SS.Range("B12").End(xlDown)))
    SS.Range("D1:G1").Value = Array("Spray-time…", "Min", "Average", "Max")
    SS.Range("D2:D6").Value = Application.Transpose(Array("…Nozzle flow", "…Nozzle pressure", "…Gun temp", "…Preheater temp", "…Carrier gas flow"))
    
    ' Need to determine the steady-state region for each process value
        ' The XA Carrier gas (column P) can be used to determine when "Powder ON" is issued
        ' The Heater XA (column I,J) can be used to determine when "Spray OFF" is issued
        ' All reported averages/statistics will come from the region between those two commands
    
    ' Do looping on VBA array instead of ranges, change consecutive relative counts to absolute counts
    rawdata = SS.Range("$A$12:" & rg.Address & "").Value
    For rownum = 1 To UBound(rawdata, 1)
        rawdata(rownum, 1) = rownum
    Next rownum
    SS.Range("$A$12:" & rg.Address & "").Value = rawdata
    
    ' Determine the start & end points of interest
    For rownum = 2 To UBound(rawdata, 1)
        ' Has the XA_CG1_Flow value decreased? A decrease
        ' indicates the end of PreFlow.  There can be more
        ' than one decrease, so only get the first instance
        If rawdata(rownum, 16) < rawdata(rownum - 1, 16) Then
            If Not startflag Then
                rowstart = rownum
                startflag = True
            End If
        End If
        ' Has the XA_Heater_Gun value has decreased to 0?
        If ((rawdata(rownum, 9) = 0) And (rawdata(rownum - 1, 9) > 0)) Then rowend = rownum
    Next rownum

    ' Print min, max, and averages of interesting quantities
    On Error Resume Next
    With SS.Range("B9:B10")
        .Value = Application.Transpose(Array(rowstart + 11, rowend + 11))
        .NumberFormat = "###0"
    End With
    SS.Range("B6:B7").Value = Application.Transpose(Array(Format(rawdata(rowstart, 2), "h:mm:ss AM/PM"), Format(rawdata(rowend, 2), "h:mm:ss AM/PM")))
    SS.Range("E2:G2").FormulaR1C1 = Array("=Min(INDIRECT(" & Chr(34) & "K" & Chr(34) & "&R9C2&" & Chr(34) & ":K" & Chr(34) & "&R10C2))", "=average(INDIRECT(" & Chr(34) & "K" & Chr(34) & "&R9C2&" & Chr(34) & ":K" & Chr(34) & "&R10C2))", "=Max(INDIRECT(" & Chr(34) & "K" & Chr(34) & "&R9C2&" & Chr(34) & ":K" & Chr(34) & "&R10C2))")
    SS.Range("E3:G3").FormulaR1C1 = Array("=Min(INDIRECT(" & Chr(34) & "C" & Chr(34) & "&R9C2&" & Chr(34) & ":C" & Chr(34) & "&R10C2))", "=average(INDIRECT(" & Chr(34) & "C" & Chr(34) & "&R9C2&" & Chr(34) & ":C" & Chr(34) & "&R10C2))", "=Max(INDIRECT(" & Chr(34) & "C" & Chr(34) & "&R9C2&" & Chr(34) & ":C" & Chr(34) & "&R10C2))")
    SS.Range("E4:G4").FormulaR1C1 = Array("=Min(INDIRECT(" & Chr(34) & "E" & Chr(34) & "&R9C2&" & Chr(34) & ":E" & Chr(34) & "&R10C2))", "=average(INDIRECT(" & Chr(34) & "E" & Chr(34) & "&R9C2&" & Chr(34) & ":E" & Chr(34) & "&R10C2))", "=Max(INDIRECT(" & Chr(34) & "E" & Chr(34) & "&R9C2&" & Chr(34) & ":E" & Chr(34) & "&R10C2))")
    SS.Range("E5:G5").FormulaR1C1 = Array("=Min(INDIRECT(" & Chr(34) & "G" & Chr(34) & "&R9C2&" & Chr(34) & ":G" & Chr(34) & "&R10C2))", "=average(INDIRECT(" & Chr(34) & "G" & Chr(34) & "&R9C2&" & Chr(34) & ":G" & Chr(34) & "&R10C2))", "=Max(INDIRECT(" & Chr(34) & "G" & Chr(34) & "&R9C2&" & Chr(34) & ":G" & Chr(34) & "&R10C2))")
    SS.Range("E6:G6").FormulaR1C1 = Array("=Min(INDIRECT(" & Chr(34) & "O" & Chr(34) & "&R9C2&" & Chr(34) & ":O" & Chr(34) & "&R10C2))", "=average(INDIRECT(" & Chr(34) & "O" & Chr(34) & "&R9C2&" & Chr(34) & ":O" & Chr(34) & "&R10C2))", "=Max(INDIRECT(" & Chr(34) & "O" & Chr(34) & "&R9C2&" & Chr(34) & ":O" & Chr(34) & "&R10C2))")
    On Error GoTo 0
    
    ' Format the sheet text
    SS.Columns("A").AutoFit
    SS.Range("D1").Font.Italic = True
    SS.Range("D1:D6").HorizontalAlignment = xlRight
    SS.Range("E2:G6").NumberFormat = "#,##0.00"
    SS.Range("E1:G1").Font.Bold = True
    SS.Columns("E:G").AutoFit
    
    ' Make graphs of interesting quantities
    SS.Activate
    SS.Range("P2").Select
    ' Gun Pressure and Nitrogen Flow
    SS.Shapes.AddChart(XlChartType:=xlXYScatterLinesNoMarkers, Left:=(Range("I1").Left + Range("H1").Left) / 2, Top:=0, Width:=Range("E1").Left, Height:=Range("G11").Top).Select
    With ActiveChart
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .Name = "Pressure"
            .XValues = SS.Range(Cells(12, 1), Cells(rg.Row, 1))
            .Values = SS.Range(Cells(12, 3), Cells(rg.Row, 3))
        End With
        .SeriesCollection.NewSeries
        With .SeriesCollection(2)
            .AxisGroup = 2
            .Name = "N2 Flow"
            .XValues = SS.Range(Cells(12, 1), Cells(rg.Row, 1))
            .Values = SS.Range(Cells(12, 11), Cells(rg.Row, 11))
        End With
        ' Logic for plotting He flow
        If Application.Max(Application.Index(rawdata, 0, 14)) > 1 Then
            .SeriesCollection.NewSeries
            With .SeriesCollection(3)
                .AxisGroup = 2
                .Name = "He Flow"
                .XValues = SS.Range(Cells(12, 1), Cells(rg.Row, 1))
                .Values = SS.Range(Cells(12, 13), Cells(rg.Row, 13))
            End With
        End If
        ' End He flow plot logic
        With .Axes(xlValue, xlPrimary)
            .MinimumScale = 0
            .MajorGridlines.Delete
            .HasTitle = True
            .Format.Line.Visible = msoCTrue
            .AxisTitle.Text = "Pressure (bar)"
        End With
        With .Axes(xlValue, xlSecondary)
            .MinimumScale = 0
            .HasTitle = True
            .Format.Line.Visible = msoCTrue
            .AxisTitle.Text = "Gas Flow (m³/h)"
        End With
        .Legend.Position = xlLegendPositionTop
        With .Axes(xlCategory)
            .MinimumScale = 0
            .HasTitle = True
            .Format.Line.Visible = msoCTrue
            .AxisTitle.Text = "Spray Run Time (s)"
        End With
        With .PlotArea.Format.Line
            .Visible = msoCTrue
            .Style = msoLineSingle
        End With
    End With
    ' Gun Temp & Preheater Temp
    SS.Shapes.AddChart(XlChartType:=xlXYScatterLinesNoMarkers, Left:=Range("M1").Left, Top:=0, Width:=Range("e1").Left, Height:=Range("G11").Top).Select
    With ActiveChart
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .Name = "Gun"
            .XValues = SS.Range(Cells(12, 1), Cells(rg.Row, 1))
            .Values = SS.Range(Cells(12, 5), Cells(rg.Row, 5))
        End With
        .SeriesCollection.NewSeries
        With .SeriesCollection(2)
            .AxisGroup = 2
            .Name = "Preheater"
            .XValues = SS.Range(Cells(12, 1), Cells(rg.Row, 1))
            .Values = SS.Range(Cells(12, 7), Cells(rg.Row, 7))
        End With
        With .Axes(xlValue, xlPrimary)
            .MinimumScale = 0
            .MajorGridlines.Delete
            .HasTitle = True
            .Format.Line.Visible = msoCTrue
            .AxisTitle.Text = "Gun T (°C)"
        End With
        With .Axes(xlValue, xlSecondary)
            .MinimumScale = 0
            .HasTitle = True
            .Format.Line.Visible = msoCTrue
            .AxisTitle.Text = "Preheater T (°C)"
        End With
        .Legend.Position = xlLegendPositionTop
        With .Axes(xlCategory)
            .MinimumScale = 0
            .HasTitle = True
            .Format.Line.Visible = msoCTrue
            .AxisTitle.Text = "Spray Run Time (s)"
        End With
        With .PlotArea.Format.Line
            .Visible = msoCTrue
            .Style = msoLineSingle
        End With
    End With
    SS.Columns("U").Delete
    SS.Range("B1").Select
Next i
Application.ScreenUpdating = True
End Sub
Private Function WorksheetExists(ByVal sheetName As String) As Boolean
' Determine if a Worksheet with that name exists in the workbook
' Found at MrExcel.com: http://www.mrexcel.com/forum/showthread.php?t=3228
    On Error Resume Next
    WorksheetExists = (Sheets(sheetName).Name <> "")
    On Error GoTo 0
End Function
Sub CS_LoadProfileFile()
'
' CS_LoadProfileFile Macro
'

Dim fd As Object
Dim i As Integer, j As Integer
Dim fileName As String
Dim hData As Variant, curmin As Double
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.AllowMultiSelect = True
    fd.Show
    Application.ScreenUpdating = False
    Dim df As New Collection
    For i = 2 To fd.SelectedItems.Count
        df.Add fd.SelectedItems.Item(i)
    Next i
    df.Add fd.SelectedItems.Item(1)
    
    For i = 1 To df.Count
        fileName = df.Item(i)
        With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & fileName & "", _
            Destination:=Cells(1, i * 2 - 1))
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
            .TextFileCommaDelimiter = True
            .TextFileSpaceDelimiter = False
            .TextFileColumnDataTypes = Array(1, 1)
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False
        End With
        ActiveWorkbook.Connections(ActiveWorkbook.Connections.Count).Delete
        Cells(1, i * 2 - 1).Value = Dir(fileName)
        Range(Cells(1, i * 2 - 1), Cells(1, i * 2)).Merge
    Next i
    Application.ScreenUpdating = True
    MsgBox ("Imported " & fd.SelectedItems.Count & " files")
    Exit Sub
    MsgBox ("Shifting all reported Heights such that minimum height = 0")
    For i = 1 To fd.SelectedItems.Count
        hData = Range(Cells(4, i * 2), Cells(4, i * 2).End(xlDown)).Value
        curmin = Application.Min(hData)
        For j = 1 To UBound(hData, 1)
            hData(j, 1) = hData(j, 1) - curmin
        Next j
        Range(Cells(4, i * 2), Cells(4, i * 2).End(xlDown)).Value = hData
    Next i
    
End Sub
