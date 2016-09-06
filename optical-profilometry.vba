Option Explicit
Private Sub SummarizeSheets()
' Ask for confirmation once
'

End Sub
Private Sub ReadProfile(fileName As String, sheet As Excel.Worksheet)
Dim fileProps As Collection, numPoints As Long, leftStart As Long, rightStart As Long, leftEnd As Long, rightEnd As Long, divotStart As Long, divotEnd As Long, divotMid As Long
' Reads in the specified filename and puts the data on the specified sheet
    With sheet.QueryTables.Add(Connection:="TEXT;" & fileName & "", Destination:=Range("$A$4"))
        .Name = fileName
        .FieldNames = False
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = False
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 437
        .TextFileStartRow = 3
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    With ActiveWorkbook
        If .Connections.Count > 0 Then .Connections(1).Delete
        If .Names.Count > 0 Then .Names(.Names.Count).Delete
    End With
    Set fileProps = extractParameters(fileName)
    With sheet
        numPoints = .Cells.SpecialCells(xlCellTypeLastCell).Row - 4
        leftStart = 5
        rightEnd = numPoints + 4
        ' add/subtract 20% of the points gathered as the edges
        leftEnd = leftStart + Round(numPoints * 0.2, 0)
        rightStart = rightEnd - Round(numPoints * 0.2, 0)
        
        ' blanket divot - start in middle and work outward ±10%, assuming the trace was centered on the divot
        divotMid = Round(numPoints / 2, 0) + 4
        divotStart = divotMid - Round(numPoints * 0.1, 0)
        divotEnd = divotMid + Round(numPoints * 0.1, 0)
        
        
        .Range("A1:L1").Value = Array(fileProps("sampleName"), fileProps("energy"), fileProps("millTime"), "", "Left", "Right", "", "", "Indent", "Distance", "Floor Span", "Floor Val")
        .Range("A2:M2").Formula = Array("Ind#" & fileProps("indentNumber"), "Trace#" & fileProps("traceNumber"), "", "RowStart", leftStart, rightStart, "", "Est. Start", divotStart, "=INDIRECT(" & Chr(34) & "$A" & Chr(34) & "&$I2)", "=$I$5-0.1*$I$5", "=$I$4", "µm")
        .Range("A3:M3").Formula = Array("Num Points", numPoints, "", "Avg. Stop", leftEnd, rightEnd, "", "Est. End", divotEnd, "=INDIRECT(" & Chr(34) & "$A" & Chr(34) & "&$I3)", "=$I$5+0.1*$I$5", "=$I$4", "µm")
        .Range("A4:J4").Formula = Array("Distance (µm)", "Height (µm)", "", "Dist-Start", "=INDIRECT(" & Chr(34) & "$A" & Chr(34) & "&$E2)", "=INDIRECT(" & Chr(34) & "$A" & Chr(34) & "&$F2)", "µm", "Local Min.", "=MIN(INDIRECT(" & Chr(34) & "B" & Chr(34) & "&$I$2&" & Chr(34) & ":B" & Chr(34) & "&$I$3))", "µm")
        .Range("D5:J5").Formula = Array("Dist-Stop", "=INDIRECT(" & Chr(34) & "$A" & Chr(34) & "&$E3)", "=INDIRECT(" & Chr(34) & "$A" & Chr(34) & "&$F3)", "µm", "'Dist@Min", "=INDEX(INDIRECT(" & Chr(34) & "A" & Chr(34) & "&$I$2&" & Chr(34) & ":A" & Chr(34) & "&$I$3),MATCH($I$4,INDIRECT(" & Chr(34) & "B" & Chr(34) & "&$I$2&" & Chr(34) & ":B" & Chr(34) & "&$I$3),0))", "µm")
        .Range("D6:G6").Formula = Array("Avg. Height", "=AVERAGE(INDIRECT(" & Chr(34) & "B" & Chr(34) & "&$E$2&" & Chr(34) & ":B" & Chr(34) & "&$E$3))", "=AVERAGE(INDIRECT(" & Chr(34) & "B" & Chr(34) & "&$F$2&" & Chr(34) & ":B" & Chr(34) & "&$F$3))", "µm")
        .Range("D7:J7").Formula = Array("Avg. Height", "=AVERAGE(INDIRECT(" & Chr(34) & "B" & Chr(34) & "&$E$2&" & Chr(34) & ":B" & Chr(34) & "&$E$3))", "=AVERAGE(INDIRECT(" & Chr(34) & "B" & Chr(34) & "&$F$2&" & Chr(34) & ":B" & Chr(34) & "&$F$3))", "µm", "Calc. Depth", "=AVERAGE($E$6:$E$7)-$I$4", "µm")
        .Range("E2:F3,I2:I3").Interior.ColorIndex = 24 ' color the cells to be manually edited
        .Range("L6:L7").Formula = Application.Transpose(Array("Summarize", "False"))
        .Range("E4:F7,I4:I7,J2:L3").NumberFormat = "0.000"
        .Columns("A:F").AutoFit

        ' insert chart plotting code
        Call PlotSingleDataSeries(xlXYScatterLinesNoMarkers, .Range(.Cells(5, 1), .Cells(numPoints + 4, 1)), .Range(.Cells(5, 2), .Cells(numPoints + 4, 2)), _
                   "Profile", Range("A9").Top, Range("D1").Left, Range("G20").Top * 0.99, Range("K1").Left - Range("D1").Left, _
                    VertAxTitle:="Height (µm)", VertAxMin:=-3, HorzAxMin:=0, LegendPos:=xlLegendPositionTop, _
                    HorzAxTitle:="Distance (µm)")
        Call PlotAddSingleDataSeries(.Range(.Cells(4, 5), Cells(5, 5)), .Range(.Cells(6, 5), Cells(7, 5)), "Left Bkgd")
        Call PlotAddSingleDataSeries(.Range(.Cells(4, 6), Cells(5, 6)), .Range(.Cells(6, 6), Cells(7, 6)), "Right Bkgd")
        Call PlotAddSingleDataSeries(.Range(.Cells(2, 11), Cells(3, 11)), .Range(.Cells(2, 12), .Cells(3, 12)), "Indent Floor")
        
            
            
    End With
End Sub
Public Sub IndentUpdate()
' Select files, run import tool to arrange them -
' one indent per sheet.
Dim wb As Excel.Workbook, SS As Excel.Worksheet, fd As FileDialog, fileList As FileDialogSelectedItems, fileName As String, i As Long, filePath As String, _
    procd As Collection, wData As Collection, fData As Variant
Set wb = ActiveWorkbook

' Grab list of previously imported files & assemble collection
Set procd = New Collection
On Error GoTo DataImportStepBegin
fData = wb.Sheets("imported").Range(wb.Sheets("imported").Cells(1, 1), wb.Sheets("imported").Cells(wb.Sheets("imported").Cells.SpecialCells(xlCellTypeLastCell).Row, 1)).Value
For i = 1 To UBound(fData, 1)
    procd.Add True, Key:=fData(i, 1)
Next i
On Error GoTo 0

DataImportStepBegin:
Set fd = Application.FileDialog(msoFileDialogFilePicker)
With fd
    .AllowMultiSelect = True
    .Filters.Add "CSV Files", "*.csv", 1
    .Title = "Select the Zygo Scan Profiles to import"
    .Show
End With

If fd.SelectedItems.Count > 0 Then
    Set wData = New Collection
    Set fileList = fd.SelectedItems
    For i = 1 To fileList.Count
        filePath = fileList(i)
        fileName = extractParameters(filePath).Item("sheetName")
        ' check if it has been imported already
        If (isImportedAlready(fileName, procd) = True) Then
            Debug.Print "imported previously"
        Else
        ' not yet imported, so import
            ReadProfile filePath, createNewSheet(fileName)
            procd.Add True, Key:=fileName
            wData.Add fileName, Key:=fileName
       End If
    Next i
    If wData.Count > 0 Then writeImportedList wData, wb
    
    ' insert data summarization code
    
End If
    


End Sub
Private Function getFileName(filePath As String) As String
Dim pathStr As String, fileName As String, newFileName As String, p As Long
' ex filePath = C:\Users\BJH\Box Sync\NEET\Data\HT9\Zygo\HT9r01s07\HT9r01s07 6kV 0hr\HT9r01s07.indent.01.04.csv
' ex pathStr == C:\Users\BJH\Box Sync\NEET\Data\HT9\Zygo\HT9r01s07\HT9r01s07 6kV 0hr\
' ex fileName = HT9r01s07.indent.01.04.csv
' ex newFileName = HT9r01s07_id01_tr04
p = InStrRev(filePath, "\")
pathStr = Left(filePath, p - 1)
fileName = Replace(Right(filePath, Len(filePath) - p), "indent", "ind")
newFileName = Left(Replace(fileName, ".", "_"), Len(fileName) - 4)
getFileName = newFileName
End Function
Private Function createNewSheet(sheetName As String) As Excel.Worksheet
' Creates a sheet for holding the CSV data from indent traces
Dim newSheet As Excel.Worksheet
On Error GoTo YesExist
Set newSheet = Sheets.Add(After:=Sheets(Sheets.Count))
newSheet.Name = sheetName
On Error GoTo 0
Set createNewSheet = newSheet
Exit Function
YesExist:
    Debug.Print "A sheet with that name cannot exist"
Set createNewSheet = newSheet
End Function
Private Function isImportedAlready(fileName As String, importedlist As Collection) As Boolean
' returns True if a filename has already been imported, and false if not
    On Error GoTo FileNotFound
        isImportedAlready = importedlist.Item(fileName)
        Exit Function
    On Error GoTo 0
FileNotFound:
    isImportedAlready = False
End Function
Private Function writeImportedList(newImports As Collection, wb As Excel.Workbook) As Boolean
' writes all newly-imported files to the imported sheet
    Dim sheet As Excel.Worksheet, i As Long, pData() As Variant
    On Error GoTo MakeImportedSheet
    Set sheet = wb.Sheets("imported")
    On Error GoTo 0
DoWriteOfList:
    ReDim pData(1 To newImports.Count, 1 To 1)
    For i = 1 To newImports.Count
        pData(i, 1) = newImports(i)
    Next i
    sheet.Range(sheet.Cells(sheet.Cells.SpecialCells(xlCellTypeLastCell).Row + LBound(pData), 1), sheet.Cells(sheet.Cells.SpecialCells(xlCellTypeLastCell).Row + UBound(pData), 1)).Value = pData
    Exit Function
MakeImportedSheet:
    Set sheet = wb.Sheets.Add
    sheet.Visible = xlSheetHidden
    sheet.Name = "imported"
    GoTo DoWriteOfList
End Function
Private Function extractParameters(filePath As String) As Collection
Dim p As Long, pathStr As String, fileName As String, sampleName As String, energy As String, indentNumber As Integer, traceNumber As Integer, millTime As String, sheetName As String
' ex filePath = C:\Users\BJH\Box Sync\NEET\Data\HT9\Zygo\HT9r01s07\HT9r01s07 6kV 0hr\HT9r01s07.repolish.00.6kV.00h.indent.01.04.csv
' ex pathStr == C:\Users\BJH\Box Sync\NEET\Data\HT9\Zygo\HT9r01s07\HT9r01s07 6kV 0hr\
' ex fileName = HT9r01s07.repolish.00.6kV.00h.indent.01.04.csv
' ex newFileName = HT9r01s07_id01_tr04
Set extractParameters = New Collection
p = InStrRev(filePath, "\")
pathStr = Left(filePath, p - 1)
fileName = Right(filePath, Len(filePath) - p)
p = InStr(1, fileName, ".")
sampleName = Left(fileName, p - 1)
p = InStr(p, fileName, "kV") ' search for energy units
If (p > 0) Then
    energy = Mid(fileName, InStrRev(fileName, ".", p) + 1, 3) ' kV = 1 / 2 / 3 / 4 / 5 / 6 as values
Else
    p = InStr(InStr(1, fileName, "."), fileName, "eV.") ' search for other energy units
    If (p > 0) Then
        energy = Mid(fileName, InStrRev(fileName, ".", p) + 1, 5) ' eV = 100 / 200 / 300 ... 999 as values
    Else
        energy = "u"
        p = 1
    End If
End If
p = InStr(p, fileName, "h.") ' search for time units
If (p > 0) Then
    millTime = Mid(fileName, InStrRev(fileName, ".", p) + 1, p - InStrRev(fileName, ".", p)) ' hours units
Else
    p = InStr(InStr(1, fileName, "."), fileName, "m.")
    If (p > 0) Then
        millTime = Mid(fileName, InStrRev(fileName, ".", p) + 1, p - InStrRev(fileName, ".", p)) ' minutes units
    Else
        millTime = "u"
    End If
End If
p = InStr(1, fileName, "indent.")
indentNumber = Mid(fileName, p + 7, 2)
traceNumber = Mid(fileName, p + 10, 2)
sheetName = sampleName & "_" & energy & millTime & "_ind" & indentNumber & "tr" & traceNumber
extractParameters.Add pathStr, Key:="pathStr"
extractParameters.Add fileName, Key:="fileName"
extractParameters.Add sampleName, Key:="sampleName"
extractParameters.Add energy, Key:="energy"
extractParameters.Add millTime, Key:="millTime"
extractParameters.Add indentNumber, Key:="indentNumber"
extractParameters.Add traceNumber, Key:="traceNumber"
extractParameters.Add sheetName, Key:="sheetName"
End Function
