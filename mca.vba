Sub MCAImport2()
' Written 2015-02-10 by Ben Hauch
' Last updated 2015-02-10 by Ben Hauch
' Load in a .txt or .asc file from the PGT Multichannel Analyzer, place dataset start at cursor position.
Dim SS As Worksheet, FileName As String, fd As Object, originCell As String, colData As Variant, nBins As Integer
' File browser for the input datafile
    originCell = ActiveCell.Address
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .InitialView = msoFileDialogViewDetails
        .Title = "Select the PGT .ASC datafile, or the renamed .txt datafile"
        .AllowMultiSelect = False
        .Filters.Add "TXT/ASC Files", "*.txt; *.asc", 1
        .Show
    End With
    FileName = fd.SelectedItems.Item(1)
    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & FileName & "", Destination:=Range(originCell))
        .Name = FileName
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = False
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = False
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
        .TextFileSpaceDelimiter = True
        .TextFileColumnDataTypes = Array(1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    If ActiveWorkbook.Connections.Count > 0 Then ActiveWorkbook.Connections(1).Delete
    Range(originCell).Offset(0, 1).Value = FileName
    Range(Range(originCell).Offset(2, 0), Range(originCell).Offset(11, 1)).Clear
    Range(Range(originCell).Offset(3, 2), Range(originCell).Offset(4, 2)).Clear
    Range(Range(originCell).Offset(2, 0), Range(originCell).Offset(8, 0)).Value = Application.Transpose(Array("Acquisition Date:", "Elapsed Real Time:", "Elapsed Live Time:", "Conversion Gain:", "High Voltage:", "Coarse Gain:", "Fine Gain:"))
    Range(Range(originCell).Offset(11, 0), Range(originCell).Offset(11, 3)).Value = Array("MCA Bin Voltage [V]", "Channel #", "Counts", "Rel. Freq.")
    nBins = Range(originCell).Offset(5, 2).Value
    If nBins = 1024 Then
        ' need to shift channels 1000-1023 right by one
        colData = Range(Range(originCell).Offset(1012, 0), Range(originCell).Offset(1035, 1)).Value
        Range(Range(originCell).Offset(1012, 0), Range(originCell).Offset(1035, 1)).Clear
        Range(Range(originCell).Offset(1012, 1), Range(originCell).Offset(1035, 2)).Value = colData
    End If
    Range(Range(originCell).Offset(12, 0), Range(originCell).Offset(11 + nBins, 0)).FormulaR1C1 = "=10/" & nBins & "*R[0]C[1]"
    Range(Range(originCell).Offset(10, 1), Range(originCell).Offset(10, 2)).FormulaR1C1 = Array("Integral", "=SUM(R[2]C[0]:R[" & 1 + nBins & "]C[0])")
    Range(Range(originCell).Offset(12, 3), Range(originCell).Offset(11 + nBins, 3)).FormulaR1C1 = "=R[0]C[-1]/R11C[-1]"
    Range(Range(originCell).Offset(12, 3), Range(originCell).Offset(11 + nBins, 3)).NumberFormat = "0.00%"
    Columns(Range(originCell).Offset(0, 4).Column).Delete
    Range(originCell).Offset(0, 6).Select
    
End Sub
