Sub FretWearSingle()
' Written by Ben Hauch on 06/21/2011
'   Last updated by Ben Hauch on 06/21/2011
'
' FretWearSingle reads in a user-selected file and creates a
' formatted plot of the  friction response
Dim fileName$
Dim fd As Object
Set fd = Application.FileDialog(msoFileDialogFilePicker)
fd.AllowMultiSelect = False
fd.Show
fileName = fd.SelectedItems.Item(1)
Call FretWearProcessing(fileName)
End Sub
Sub FretWearMany()
' Written by Ben Hauch on 06/21/2011
'   Last updated by Ben Hauch on 06/21/2011
'
' FretWearMany reads in all .txt files in a user-selected directory
' and runs FretWearSingle for each file. Alphabetical order is not preserved,
' but the filename is printed for each file
Dim sttime#
Dim fileName$
Dim FolderName$
Dim AllFiles As New Collection
Dim fd As Object
sttime = Timer
Set fd = Application.FileDialog(msoFileDialogFolderPicker)
fd.AllowMultiSelect = False
fd.Show
FolderName = fd.SelectedItems.Item(1)
Application.ScreenUpdating = False
fileName = Dir(FolderName & "*.txt")
While fileName <> ""
    AllFiles.Add fileName
    fileName = Dir
Wend
For i = 1 To AllFiles.Count
    Call FretWearProcessing(AllFiles(i))
Next i
Application.ScreenUpdating = True
Debug.Print "Exe time: " & Format(Timer - sttime, "##0.00") & "s"
End Sub
Private Sub FretWearProcessing(fileName$)
Dim SS As Worksheet
Dim DataRange As Range
Set SS = Sheets.Add(After:=Sheets(Sheets.Count))
SS.Range("A1:B1").Value = Array("File Name", fileName)
    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & fileName & "", _
        Destination:=SS.Range("$A$2"))
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
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileOtherDelimiter = "="
        .TextFileColumnDataTypes = Array(1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    ActiveWorkbook.Connections(ActiveWorkbook.Connections.Count).Delete
    SS.Range(SS.Range("A11"), SS.Range("A11").End(xlDown).Offset(-2, 0)).Cut Destination:=SS.Range("B11")
    Set DataRange = SS.Range(SS.Range("B11"), SS.Range("B11").End(xlDown))
    ' Need to find limits of actual data collection
    arData = DataRange.Value
    For i = 1 To UBound(arData)
        If arData(i, 1) > -50 Then Exit For
    Next i
    For j = Round(UBound(arData) / 2, 0) To UBound(arData)
        If arData(j, 1) < -50 Then Exit For
    Next j
    Set DataRange = Nothing
    Set DataRange = SS.Range(Cells(i + 9, 2), Cells(j + 10, 2))
    SS.Cells(i + 9, 1).FormulaR1C1 = "0"
    SS.Cells(i + 10, 1).FormulaR1C1 = "=R[-1]C+R9C2*R5C2"
    SS.Cells(i + 10, 1).AutoFill Destination:=SS.Range(SS.Cells(i + 10, 1), SS.Cells(i + 10, 1).End(xlDown).Offset(-1, 0)), Type:=xlFillDefault
    SS.Shapes.AddChart(XlChartType:=xlXYScatterLinesNoMarkers, Left:=SS.Range("D1").Left, Top:=SS.Range("D1").Top, Width:=SS.Range("B1").Left).Select
    With ActiveChart.SeriesCollection(1)
        .Name = SS.Range("B4").Value & " -" & SS.Range("B3").Value & " tip"
        .XValues = "='" & SS.Name & "'!" & DataRange.Offset(0, -1).Address & ""
        .Values = "='" & SS.Name & "'!" & DataRange.Address & ""
    End With
    SS.ChartObjects(1).Activate
    With ActiveChart
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "Friction Response (arb. units)"
            .MinimumScale = -5
            .MinorTickMark = xlInside
            .CrossesAt = -10
        End With
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "Approx. # Fretting Cycles"
            .MinimumScale = 0
            .MinorTickMark = xlInside
        End With
        .PlotArea.Select
        .SetElement (msoElementPrimaryCategoryGridLinesMajor)
        .HasLegend = False
    End With
End Sub
