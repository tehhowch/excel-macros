Option Explicit
Sub GetXRDFile()
' Written by Ben Hauch February 2012
' Last updated by Ben Hauch, 6/7/2012
'
' GetXRDFile Macro
'   Imports a .pd3 file, which has a strange-ish data format,
'   converts it to a deg-count-relcount format, and plots
'

Dim fileName As String, sttime As Double
Dim fd As Object
Dim SS As Worksheet
Dim DataRange As Range
Dim Data() As Variant, PasteData() As Variant
Dim xstep As Double, xmin As Double, xmax As Double, xval As Double
Dim i As Long, j As Long, ymax As Double

Set fd = Application.FileDialog(msoFileDialogFilePicker)
With fd
    .AllowMultiSelect = False
    .Filters.Add "XRD Files", "*.pd3", 1
    .Show
End With
fileName = fd.SelectedItems.Item(1)
Set fd = Nothing
If Right(fileName, 4) <> ".pd3" Then MsgBox Prompt:="Wrong Data Format: .PD3 required"
sttime = Timer
If (ActiveSheet.UsedRange.Address <> "$A$1") Or (ActiveSheet.Range("A1").Value <> "") Then Set SS = Sheets.Add(After:=Sheets(Sheets.Count)) Else Set SS = ActiveSheet
Application.ScreenUpdating = False
With SS.QueryTables.Add(Connection:="TEXT;" & fileName & "", _
    Destination:=Range("$A$1"))
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
    .TextFileSemicolonDelimiter = False
    .TextFileCommaDelimiter = False
    .TextFileSpaceDelimiter = False
    .TextFileOtherDelimiter = "="
    .TextFileColumnDataTypes = Array(1, 1)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=False
End With
' Remove the created data connection
ActiveWorkbook.Connections(ActiveWorkbook.Connections.Count).Delete

' Convert header data from single column to multiple columns
SS.Range("A1:A20").TextToColumns Destination:=SS.Range("A1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
    Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
    :="&", FieldInfo:=Array(Array(1, 9), Array(2, 1)), TrailingMinusNumbers:=True

' Convert collected data from single column to multiple columns
SS.Range("A21").Select
SS.Range(Selection, Selection.End(xlDown).Offset(-1, 0)).Select
Selection.TextToColumns Destination:=SS.Range("A21"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
    Semicolon:=False, Comma:=False, Space:=True, Other:=True, OtherChar:= _
    "&", FieldInfo:=Array(Array(1, 9), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), _
    Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1)), TrailingMinusNumbers _
    :=True
Columns("A:A").EntireColumn.AutoFit

' Fill in "SAMPLE IDENT" and "DATE" fields correctly
fileName = SS.Range("B4").Value
If IsEmpty(SS.Range("B1")) Then
    SS.Range("B1").Value = StrConv(Right(Left(fileName, Len(fileName) - 4), Len(fileName) - 4 - 9), vbProperCase)
End If
If (Left(fileName, 8) Like "20*") And (Left(fileName, 9) Like "*_") Then
    ' assuming a date coded in YYYYMMDD format
    SS.Range("B3").Value = Right(Left(fileName, 6), 2) & "/" & Right(Left(fileName, 8), 2) & "/" & Left(fileName, 4)
End If


' Need to reconstruct the data for plotting
xstep = SS.Range("B12").Value
Data = SS.Range(Cells(21, 1), Cells(21, 8).End(xlDown).Offset(0, 1)).Value
xmin = SS.Range("B16").Value
xmax = SS.Range("B17").Value
ymax = SS.Range("B18").Value
SS.Range("B20").Value = "Count"
SS.Range("C20").Value = "Rel. Intensity"
ReDim PasteData(1 To SS.Range("B19").Value, 1 To 3)
xval = xmin
i = 1
On Error Resume Next
Do While i <= UBound(PasteData, 1)
    For j = 2 To 9  ' 8 columns of y data
        PasteData(i + j - 2, 1) = xval
        PasteData(i + j - 2, 2) = Data((i \ 8) + 1, j)
        PasteData(i + j - 2, 3) = PasteData(i + j - 2, 2) / ymax
        xval = xval + xstep
    Next j
    i = i + 8
    If xval >= 100 Then
        ' Formatting issues if we go over 100° in 2Theta
        Do While i <= UBound(PasteData, 1)
            For j = 2 To 9
                PasteData(i + j - 2, 1) = xval
                PasteData(i + j - 2, 2) = Data((i \ 8) + 1, j - 1)
                PasteData(i + j - 2, 3) = PasteData(i + j - 2, 2) / ymax
                xval = xval + xstep
            Next j
            i = i + 8
        Loop
        Exit Do
    End If
Loop
On Error GoTo 0
' Paste data into sheet
SS.Range(Cells(21, 1), Cells(21 + UBound(PasteData, 1), 9)).Delete (xlUp)
SS.Range(Cells(21, 1), Cells(20 + UBound(PasteData, 1), 3)).Value = PasteData
SS.Range(Cells(21, 3), Cells(20 + UBound(PasteData, 1), 3)).NumberFormat = "0.00%"
SS.Cells(21 + UBound(PasteData, 1), 1).Value = "&END"

' Plot the data
SS.Range("D1").Select
SS.Shapes.AddChart(XlChartType:=xlXYScatterLinesNoMarkers, Left:=Range("D1").Left, Top:=0, Width:=Range("E1").Left, Height:=Range("G20").Top).Select
With ActiveChart
    .SeriesCollection.NewSeries
    With .SeriesCollection(1)
        .Name = "Rel. Intensity"
        .XValues = SS.Range(Cells(21, 1), Cells(20 + UBound(PasteData, 1), 1))
        .Values = SS.Range(Cells(21, 3), Cells(20 + UBound(PasteData, 1), 3))
        .Format.Line.Weight = 0.1
    End With
    With .Axes(xlValue, xlPrimary)
        .MinimumScale = 0
        .MaximumScale = 1
        .MajorGridlines.Delete
        .HasTitle = True
        .Format.Line.Visible = msoCTrue
        .AxisTitle.Text = "Rel. Intensity"
        .TickLabels.NumberFormat = "0%"
    End With
    With .Axes(xlCategory)
        .HasTitle = True
        .Format.Line.Visible = msoCTrue
        .AxisTitle.Text = "2Theta (°)"
    End With
    With .PlotArea.Format.Line
        .Visible = msoCTrue
        .Style = msoLineSingle
    End With
    .HasTitle = False
    .HasLegend = False
End With
Application.ScreenUpdating = True
Debug.Print Format(Timer - sttime, "0.000") & " sec"
End Sub
Sub BatchXRDCompare()
' Run this sub on already-opened files and then plot their relevant spectra together
If Sheets.Count < 2 Then
    MsgBox Prompt:="Please run GetXRDFile to set up several files first", Title:="Not enough information..."
    Exit Sub
End If
Dim i As Long           ' Sheet/series iterator
Dim j As Long           ' Counter
Dim strln As Integer    ' string length
Dim SS As Worksheet
Dim xData() As Variant, yData() As Variant
Dim chtmin As Double, chtmax As Double, plotstyle As Integer
Dim AvailData As New Collection, SeriesNames As New Collection, newname As String
Application.ScreenUpdating = False
j = 0
chtmin = 1000
chtmax = 1
Set AvailData = New Collection
Set SeriesNames = New Collection
For i = 1 To Sheets.Count
    ' Only prompt on sheets that are macro-imported X-Ray data
    If (Sheets(i).Range("A1").Value = "SAMPLE IDENT") Then
        If MsgBox("Sheet: " & Sheets(i).Name & Chr(10) & "Series Name: " & Sheets(i).Range("B1").Value & Chr(10) & Chr(10) & "Plot data?", vbYesNo, "Select data to plot") = vbYes Then
            AvailData.Add Item:=Sheets(i).Name
            SeriesNames.Add Item:=Sheets(i).Range("B1").Value, Key:=Sheets(i).Name
        End If
    End If
Next i
If Not AvailData.Count = 0 Then
    For i = 1 To AvailData.Count
        ' Get User-spec name for the series
        If MsgBox("Sheet: " & AvailData.Item(i) & Chr(10) & "Use '" & SeriesNames.Item(AvailData(i)) & "' as the series name?", vbYesNo, "Change the default series name?") = vbNo Then
            SeriesNames.Remove (AvailData(i))
            newname = InputBox("Type the new series name for sheet " & AvailData(i), "Enter new Series name", Sheets(AvailData(i)).Range("B1").Value)
            If newname = vbNullString Then newname = Sheets(AvailData(i)).Range("B1").Value
            SeriesNames.Add newname, Key:=AvailData(i)
        End If
    Next i
    plotstyle = MsgBox("How to plot the data?  Yes = Stack separately, No = Overlay", vbYesNoCancel, "Stack sample spectra instead of overlay?")
    If Not plotstyle = vbCancel Then
        ' Prepare the data table
        Set SS = Sheets.Add(before:=Sheets(1))
        j = 1
        If plotstyle = vbYes Then
            newname = "Stack"
            Do While WorksheetExists(newname)
                newname = "Stack" & " " & j
                j = j + 1
            Loop
        ElseIf plotstyle = vbNo Then
            newname = "Overlay"
            Do While WorksheetExists(newname)
                newname = "Overlay" & " " & j
                j = j + 1
            Loop
        End If
        SS.Name = newname
        j = 0
        For i = 1 To AvailData.Count
            Sheets(AvailData(i)).Activate
            xData = Sheets(AvailData(i)).Range(Cells(21, 1), Cells(21, 1).End(xlDown).Offset(-1, 0)).Value
            yData = Sheets(AvailData(i)).Range(Cells(21, 3), Cells(21, 3).End(xlDown)).Value
            SS.Cells(3, j + 2).Value = SeriesNames(AvailData(i))
            SS.Activate
            SS.Range(Cells(4, 1 + j), Cells(3 + UBound(yData), 1 + j)).Value = xData
            If plotstyle = 6 Then Call XRDDataStack(yData, AvailData.Count - i)
            SS.Range(Cells(4, 2 + j), Cells(3 + UBound(yData), 2 + j)).Value = yData
            j = j + 2
            If xData(1, 1) < chtmin Then chtmin = xData(1, 1)
            If xData(UBound(xData), 1) > chtmax Then chtmax = xData(UBound(xData), 1)
        Next i
        j = 0
        
        ' Plot!
        SS.Range("A1").Select
        SS.Shapes.AddChart(XlChartType:=xlXYScatterLinesNoMarkers, Left:=(Range("A1").Left + Range("B1").Left) / 2, Top:=Range("B2").Top, Width:=Range("I1").Left, Height:=Range("G28").Top).Select
        With ActiveChart
            For i = 1 To (AvailData.Count)
                .SeriesCollection.NewSeries
                With .SeriesCollection(i)
                    .Name = SeriesNames(AvailData(i))
                    .XValues = SS.Range(Cells(4, 1 + j), Cells(4, 1 + j).End(xlDown))
                    .Values = SS.Range(Cells(4, 2 + j), Cells(4, 2 + j).End(xlDown))
                    .Format.Line.Weight = 0.1
                End With
                j = j + 2
            Next i
            With .Axes(xlValue, xlPrimary)
                .MinimumScale = 0
                .MajorGridlines.Delete
                .Format.Line.Visible = msoCTrue
                If plotstyle = 7 Then
                    .HasTitle = True
                    .AxisTitle.Text = "Rel. Intensity (%)"
                Else
                    .HasTitle = False
                    .MajorTickMark = xlTickMarkNone
                    .TickLabels.Font.ColorIndex = 2
                    .MaximumScale = AvailData.Count
                End If
            End With
            If plotstyle = 7 Then .Legend.Position = xlLegendPositionTop Else .Legend.Position = xlLegendPositionRight
            With .Axes(xlCategory)
                .MinimumScale = chtmin
                .MaximumScale = chtmax
                .HasTitle = True
                .Format.Line.Visible = msoCTrue
                .AxisTitle.Text = "2Theta (°)"
            End With
            With .PlotArea.Format.Line
                .Visible = msoCTrue
                .Style = msoLineSingle
            End With
        End With
    End If
End If
Application.ScreenUpdating = True
End Sub
Private Sub XRDDataStack(ByRef InputYdata(), ByVal IncrementConstant As Integer)
' Used to help stack XRD datafiles when comparing multiple samples
' Element by element addition
Dim i As Integer
For i = 1 To UBound(InputYdata)
    InputYdata(i, 1) = InputYdata(i, 1) + IncrementConstant
Next i
End Sub
Private Function WorksheetExists(ByVal sheetName As String) As Boolean
' Determine if a Worksheet with that name exists in the workbook
' Found at MrExcel.com: http://www.mrexcel.com/forum/showthread.php?t=3228
    On Error Resume Next
    WorksheetExists = (Sheets(sheetName).Name <> "")
    On Error GoTo 0
End Function
