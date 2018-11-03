Attribute VB_Name = "PoreSizeMasterFileCreator"
Option Explicit
Dim xlBook1 As Excel.Workbook
Dim xlBook2 As Excel.Workbook
Dim xlSheet1 As Excel.Worksheet
Dim xlSheet2 As Excel.Worksheet
Dim CutPoint As Long, j As Long, k As Long

Sub PoreSizeMaster()
Attribute PoreSizeMaster.VB_Description = "Working Draft. For Gautam.  Compiles Pore diameter information from PaxIt reported files.  Run after Renaming macro to minimize errors."
Attribute PoreSizeMaster.VB_ProcData.VB_Invoke_Func = " \n14"
'This Macro should be run on a blank Excel file only!!!
'Will delete ALL prior data!
'Written by Ben Hauch in April 2010

    Set xlBook1 = ActiveWorkbook
'This will generate the master file's name
    Dim WkbkName1 As String, SampName As String, ContName As String, NumMax As Integer, strErrFiles(1 To 60000) As String, blErrFlag As Boolean, strErrOutput As String, blPasteFlag As Boolean
    WkbkName1 = InputBox("What is the test request designation? E.g. MT XXXX", "Workbook Name")
    

    SampName = InputBox("What is the sample type? E.g. 'BMPPS' or 'Ti'.", "Sample Name")
    ContName = InputBox("What is the control type? E.g. PPS", "Control Name")
    NumMax = InputBox("What is the maximum number of either samples or controls?")
'Samp & Cont are also used to generate the Sheet Names

'To speed analysis, we turn off screen updating and automatic calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

'We want a total of 12 clean sheets
    Call Shtmkr(12)
 

'Data Transfer Loop Section
'   In this section all relevant excel files will be indexed
'   and then individually opened and analyzed
    
'This code block generates an array of file paths that can then be opened or sorted
    Dim fso As Object
    Dim szMasterBook As String, szMasterPath As String, pthlngth As String
    Dim strArr(1 To 1000, 1 To 1) As String, i As Long, strName As String
    szMasterBook = xlBook1.Name                              'This is so the right wkbk is pasted into
    szMasterPath = InputBox("Please find & paste the directory path to the slice folders, e.g.: Z:\Gautam\Ben Hauch (Jan-Aug 2010)\Thickness Porosity Pore Size Analyses\MT 5601 - CoCr BMPPS & PPS (Pore Size)\CD Data")
    pthlngth = Len(szMasterPath)
    Let strName = Dir$(szMasterPath & "\*" & "*.xlsx")
    Do While strName <> vbNullString
        Let i = i + 1
        Let strArr(i, 1) = szMasterPath & "\" & strName
        Let strName = Dir$()
    Loop
    Set fso = CreateObject("Scripting.FileSystemObject")
    Call recurseSubFolders(fso.GetFolder(szMasterPath), strArr(), i) 'This line calls the Private Sub defined below
    Set fso = Nothing
    Dim LotNum As String
    Dim SampID As String, FileID As String, ColTitle As String
    Dim RgResume() As String                                        'Where to copy to next
    ReDim Preserve RgResume(1 To NumMax, 2 To 11)
        
    Let i = 1
    Let j = 1
    Do While strArr(i, 1) <> vbNullString
        'Each sample is ID'd by the combo of two pieces
        'of data: the lot number and the sample ID
        Workbooks.Open strArr(i, 1), ReadOnly:=True                 'We don't want to change raw data
        Set xlBook2 = Workbooks(Workbooks.Count)
        Set xlSheet2 = xlBook2.Sheets("Lines")
        LotNum = Range("C21")
        FileID = Trim(Range("C19"))
        SampID = RTrim(Left(FileID, Len(FileID) - 3))               'Get the only the sample name
        ColTitle = LotNum & " - " & SampID                          'Each sample has its own column, this
                                                                    'helps ID where to place the copied data
        'Copy the relevant data
        If xlSheet2.Range("H7:h8").Count < 2 Then                            'Check to see if there is more than one data point
            Range("H7").Copy                                        'Only one data point? then select only it.
        Else
            Range(xlSheet2.Range("H7"), xlSheet2.Range("H7").End(xlDown)).Copy            'This copies from h7 to the final entry -
        End If                                                      'only one entry would cause an error
        
        'Figure out what sheet to activate
        CutPoint = InStr(1, strArr(i, 1), "micron", vbTextCompare)
        If xlSheet2.Range("C22") = "" Then          'Check Operator # for depth info
            If Left(xlBook2.Name, 7) = SampName Then
                Select Case Right(Left(strArr(i, 1), CutPoint - 2), 1) 'The case is chosen by chopping off all characters
                    Case 0                                              'except for the first character of the subfolders,
                    Set xlSheet1 = xlBook1.Sheets(2)                    'i.e. "1" from 127 or "5" from 508
                    Case 7
                    Set xlSheet1 = xlBook1.Sheets(3)
                    Case 4
                    Set xlSheet1 = xlBook1.Sheets(4)
                    Case 1
                    Set xlSheet1 = xlBook1.Sheets(5)
                    Case 8
                    Set xlSheet1 = xlBook1.Sheets(6)
                    Case Else
                        strErrFiles(j) = xlBook2.Name
                        Let j = j + 1
                        blErrFlag = True
                        Workbooks(xlBook2.Name).Close
                        GoTo NextFile
                End Select
            Else
                Select Case Right(Left(strArr(i, 1), CutPoint - 2), 1)
                    Case 0
                    Set xlSheet1 = xlBook1.Sheets(7)
                    Case 7
                    Set xlSheet1 = xlBook1.Sheets(8)
                    Case 4
                    Set xlSheet1 = xlBook1.Sheets(9)
                    Case 1
                    Set xlSheet1 = xlBook1.Sheets(10)
                    Case 8
                    Set xlSheet1 = xlBook1.Sheets(11)
                    Case Else
                        strErrFiles(j) = xlBook2.Name
                        Let j = j + 1
                        blErrFlag = True
                        Workbooks(xlBook2.Name).Close
                        GoTo NextFile
                End Select
            End If
        Else                                                'A non-blank depth field, so use that
            If Right(Left(strArr(i, 1), CutPoint - 2), 1) = Right(xlSheet2.Range("C22").Formula, 1) Then 'but only if the value matches the directory or filename cues
                If Left(xlBook2.Name, 7) = SampName Then
                    Select Case Left(xlSheet2.Range("C22").Formula, 1)
                        Case 0
                        Set xlSheet1 = xlBook1.Sheets(2)
                        Case 1
                        Set xlSheet1 = xlBook1.Sheets(3)
                        Case 2
                        Set xlSheet1 = xlBook1.Sheets(4)
                        Case 3
                        Set xlSheet1 = xlBook1.Sheets(5)
                        Case 5
                        Set xlSheet1 = xlBook1.Sheets(6)
                        Case Else
                            strErrFiles(j) = xlBook2.Name
                            Let j = j + 1
                            blErrFlag = True
                            Workbooks(xlBook2.Name).Close
                            GoTo NextFile
                    End Select
                Else
                    Select Case Left(xlSheet2.Range("C22").Formula, 1)
                        Case 0
                        Set xlSheet1 = xlBook1.Sheets(7)
                        Case 1
                        Set xlSheet1 = xlBook1.Sheets(8)
                        Case 2
                        Set xlSheet1 = xlBook1.Sheets(9)
                        Case 3
                        Set xlSheet1 = xlBook1.Sheets(10)
                        Case 5
                        Set xlSheet1 = xlBook1.Sheets(11)
                        Case Else
                            strErrFiles(j) = xlBook2.Name
                            Let j = j + 1
                            blErrFlag = True
                            Workbooks(xlBook2.Name).Close
                            GoTo NextFile
                    End Select
                End If
            Else
                strErrFiles(j) = xlBook2.Name
                Let j = j + 1
                blErrFlag = True
                Workbooks(xlBook2.Name).Close
                GoTo NextFile
            End If
        End If
        
        xlSheet1.Activate
        'Determine the column to start in
        Let k = 1
        blPasteFlag = False
        Do While blPasteFlag = False
            If xlSheet1.Cells(14, k + 1).Formula = "" Then              'Is this column empty?
                xlSheet1.Cells(14, k + 1).Formula = ColTitle            'If so, assign this group to it and store
                RgResume(k, xlSheet1.Index) = xlSheet1.Cells(15, k + 1).Address     'the first cell for data, then
                blPasteFlag = True                                      'Signal for loop exit
            ElseIf ColTitle = xlSheet1.Cells(14, k + 1).Formula Then    'Is this column already assigned for this group?
                blPasteFlag = True                                      'If so, signal for loop exit
            Else                                                        'This column is not for this group
                k = k + 1                                               'so pick new column & repeat
            End If
        Loop
        'With a column selected, now we just select the resume point, paste and store the new resume point
        xlSheet1.Range(RgResume(k, xlSheet1.Index)).PasteSpecial xlPasteValues
        RgResume(k, xlSheet1.Index) = xlSheet1.Cells(14, k + 1).End(xlDown).Offset(1, 0).Address
        Workbooks(xlBook2.Name).Close
        
        
        
NextFile:
        Set xlBook2 = Nothing
        Set xlSheet2 = Nothing
        Let i = i + 1
    Loop
FormattingSection:
    Workbooks(szMasterBook).Activate

'To format the sheets, we call the Private Sub ShtsLblng and pass it the user-picked names
    Call ShtsLblng(SampName, ContName, NumMax, xlBook1)

'Now we re-enable the screen update and the automatic calculations
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

'Finally, the macro saves the file according to the given test request designation
    xlBook1.SaveAs (WkbkName1 & " - Pore Size Master File - " & SampName & " vs. " & ContName & ".xlsx")
    MsgBox ("The macro has finished running.  Don't forget to calculate the group averages and standard deviations for this page from the data on the Statistics page.")
    If blErrFlag = True Then GoTo OnErr2
    Exit Sub
OnErr2:
    Let j = 1
    Do While strErrFiles(j) <> vbNullString
        strErrOutput = strErrOutput & strErrFiles(j) & Chr(13)
        j = j + 1
    Loop
    MsgBox ("The following files were not analyzed properly:" & Chr(13) & _
        strErrOutput)
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    strErrOutput = ""
End Sub


Private Sub ShtsLblng(SampName, ContName, NumMax, xlBook)
'Now we name the sheets according to the slice depths of 0, 127, 254, 381, and 508,
'and sample & control names
    xlBook.Sheets(1).Name = "Net"
    xlBook.Sheets(2).Name = SampName & " 000"
    xlBook.Sheets(3).Name = SampName & " 127"
    xlBook.Sheets(4).Name = SampName & " 254"
    xlBook.Sheets(5).Name = SampName & " 381"
    xlBook.Sheets(6).Name = SampName & " 508"
    xlBook.Sheets(7).Name = ContName & " 000"
    xlBook.Sheets(8).Name = ContName & " 127"
    xlBook.Sheets(9).Name = ContName & " 254"
    xlBook.Sheets(10).Name = ContName & " 381"
    xlBook.Sheets(11).Name = ContName & " 508"
    xlBook.Sheets(12).Name = "Statistics"
'Now we fill in the box descriptors
    Dim i As Integer, xlSheet As Excel.Worksheet
    xlBook.Sheets("Statistics").Activate
    xlBook.Sheets("Statistics").Range("A1").Select
    For i = 2 To 11 'The other sheets are the summary and the statistics pages and have special formatting
        Set xlSheet = Sheets(i) 'Select the i-th sheet
        
        'Begin cell labeling
        xlSheet.Activate
        xlSheet.Range("A1").FormulaR1C1 = "Pore Diameter - " & Right(ActiveSheet.Name, 3) & " µm"
        xlSheet.Range("A3").FormulaR1C1 = "Average Pore Diameter"
        xlSheet.Range("A4").FormulaR1C1 = "Std. Dev."
        xlSheet.Range("A5").FormulaR1C1 = "Max"
        xlSheet.Range("A6").FormulaR1C1 = "Min"
        xlSheet.Range("A7").FormulaR1C1 = "Pore Count"
        xlSheet.Range("A10").FormulaR1C1 = "Sample Avg. Pore Diam."
        xlSheet.Range("A11").FormulaR1C1 = "Sample Std. Dev."
        xlSheet.Range("A12").FormulaR1C1 = "Sample Max"
        xlSheet.Range("A13").FormulaR1C1 = "Sample Min"
        xlSheet.Range("A14").FormulaR1C1 = "Sample ID"
        xlSheet.Range("A1:A14").Font.Bold = True
        xlSheet.Range("A1:A14").HorizontalAlignment = xlRight
        xlSheet.Range("A1:A14").VerticalAlignment = xlCenter
        xlSheet.Range("A1:A14").Orientation = 0
        xlSheet.Range("A1:A14").AddIndent = False
        xlSheet.Range("A1:A14").IndentLevel = 0
        xlSheet.Range("A1:A14").ShrinkToFit = False
        xlSheet.Range("A1:A14").ReadingOrder = xlContext
        xlSheet.Range("A1:A14").MergeCells = False
        xlSheet.Rows("14:14").Font.Bold = True
        xlSheet.Rows("14:14").HorizontalAlignment = xlCenter
        xlSheet.Rows("14:14").VerticalAlignment = xlCenter
        xlSheet.Rows("14:14").WrapText = True
        xlSheet.Rows("14:14").Orientation = 0
        xlSheet.Rows("14:14").AddIndent = False
        xlSheet.Rows("14:14").IndentLevel = 0
        xlSheet.Rows("14:14").ShrinkToFit = False
        xlSheet.Rows("14:14").ReadingOrder = xlContext
        xlSheet.Rows("14:14").MergeCells = False
        xlSheet.Rows("14:14").EntireRow.AutoFit
        xlSheet.Range("C3:C6").FormulaR1C1 = "µm"
        xlSheet.Range("C7").FormulaR1C1 = "pores"
        xlSheet.Columns("A").EntireColumn.AutoFit
        xlSheet.Columns("B:ZZ").ColumnWidth = 8.43
        
        'All the labeling and formatting is now done,
        'and calculation formulas must be entered.
        xlSheet.Range("B13").FormulaR1C1 = "=MIN(R15C:R60000C)"
        xlSheet.Range("B12").FormulaR1C1 = "=MAX(R15C:R60000C)"
        xlSheet.Range("B11").FormulaR1C1 = "=STDEV(R15C:R60000C)"
        xlSheet.Range("B10").FormulaR1C1 = "=AVERAGE(R15C:R60000C)"
        xlSheet.Range("B10:B13").AutoFill Destination:=xlSheet.Range(Cells(10, 2), Cells(13, NumMax + 1))
        xlSheet.Range("B7").FormulaR1C1 = "=COUNT(R15C2:R60000C[1700])"     'Num datapoints
        xlSheet.Range("B6").FormulaR1C1 = "=MIN(R13C2:R13C1702)"            'Global minimum pore size
        xlSheet.Range("B5").FormulaR1C1 = "=MAX(R12C2:R12C1702)"            'Global maximum pore size
        xlSheet.Range("B4").FormulaR1C1 = "=STDEV(R15C2:R60000C1702)"       'Global std. deviation
        xlSheet.Range("B3").FormulaR1C1 = "=AVERAGE(R15C2:R60000C1702)"     'Global average
        xlSheet.Cells.FormatConditions.Delete
        'Sort the input data according to the column title - this will separate the lot numbers
        Call LotSort(i)
        xlSheet.Range(Cells(14, 2), Cells(14, 2).SpecialCells(xlLastCell)).Copy
        Sheets(12).Select                                 'Agglomeration of the data for statistics
        ActiveCell.FormulaR1C1 = xlSheet.Name
        ActiveCell.Offset(1, 0).Select
        ActiveSheet.paste
        ActiveCell.Offset(-1, NumMax).Select                        'Do not overwrite
    Next i
    
    'Format the first page for the summary sheet
    Call netpgformat(SampName, ContName, NumMax, xlBook)
End Sub


Private Sub recurseSubFolders(ByRef Folder As Object, ByRef strArr() As String, ByRef i As Long)
'This is the grunt work code for finding all the excel07 files needed for analysis

Dim SubFolder As Object
Dim strName As String
For Each SubFolder In Folder.SubFolders
    Let strName = Dir$(SubFolder.Path & "\*" & "*.xlsx")
    Do While strName <> vbNullString
        Let i = i + 1
        Let strArr(i, 1) = SubFolder.Path & "\" & strName
        Let strName = Dir$()
    Loop
    Call recurseSubFolders(SubFolder, strArr(), i)
Next
End Sub


Private Sub netpgformat(SampName, ContName, NumMax, xlBook)
'Format the 'Net' page
    
    'Enter cell values
    xlBook.Sheets(1).Range("A1").FormulaR1C1 = "Per Slice"
    xlBook.Sheets(1).Range("A1").Font.Bold = True
    xlBook.Sheets(1).Range("A2").FormulaR1C1 = "Slice Depth (µm)"
    xlBook.Sheets(1).Range("A3").FormulaR1C1 = "Average Pore Diameter (µm)"
    xlBook.Sheets(1).Range("A4").FormulaR1C1 = "Standard Deviation (µm)"
    xlBook.Sheets(1).Range("A5").FormulaR1C1 = "Maximum Diameter (µm)"
    xlBook.Sheets(1).Range("A6").FormulaR1C1 = "Minimum Diameter (µm)"
    xlBook.Sheets(1).Range("A7").FormulaR1C1 = "Pore Count"
    xlBook.Sheets(1).Range("A12").FormulaR1C1 = "Overall"
    xlBook.Sheets(1).Range("A13").FormulaR1C1 = "Ave. Pore Diameter (µm)"
    xlBook.Sheets(1).Range("A14").FormulaR1C1 = "Std. Dev."
    xlBook.Sheets(1).Range("A15").FormulaR1C1 = "Max "
    xlBook.Sheets(1).Range("A16").FormulaR1C1 = "Min"
    xlBook.Sheets(1).Range("A17").FormulaR1C1 = "Pore Count"
    xlBook.Sheets(1).Range("B2").FormulaR1C1 = "'000"
    xlBook.Sheets(1).Range("C2").FormulaR1C1 = "127"
    xlBook.Sheets(1).Range("D2").FormulaR1C1 = "254"
    xlBook.Sheets(1).Range("E2").FormulaR1C1 = "381"
    xlBook.Sheets(1).Range("F2").FormulaR1C1 = "508"
    xlBook.Sheets(1).Range("g2").FormulaR1C1 = "'000"
    xlBook.Sheets(1).Range("H2").FormulaR1C1 = "127"
    xlBook.Sheets(1).Range("I2").FormulaR1C1 = "254"
    xlBook.Sheets(1).Range("J2").FormulaR1C1 = "381"
    xlBook.Sheets(1).Range("K2").FormulaR1C1 = "508"
    xlBook.Sheets(1).Range("B3:B7").FormulaR1C1 = "='" & SampName & " 000'!RC"
    xlBook.Sheets(1).Range("C3:c7").FormulaR1C1 = "='" & SampName & " 127'!RC[-1]"
    xlBook.Sheets(1).Range("d3:d7").FormulaR1C1 = "='" & SampName & " 254'!RC[-2]"
    xlBook.Sheets(1).Range("e3:e7").FormulaR1C1 = "='" & SampName & " 381'!RC[-3]"
    xlBook.Sheets(1).Range("f3:f7").FormulaR1C1 = "='" & SampName & " 508'!RC[-4]"
    xlBook.Sheets(1).Range("g3:g7").FormulaR1C1 = "='" & ContName & " 000'!RC[-5]"
    xlBook.Sheets(1).Range("h3:h7").FormulaR1C1 = "='" & ContName & " 127'!RC[-6]"
    xlBook.Sheets(1).Range("i3:i7").FormulaR1C1 = "='" & ContName & " 254'!RC[-7]"
    xlBook.Sheets(1).Range("j3:j7").FormulaR1C1 = "='" & ContName & " 381'!RC[-8]"
    xlBook.Sheets(1).Range("k3:k7").FormulaR1C1 = "='" & ContName & " 508'!RC[-9]"
    xlBook.Sheets(1).Range("B12").FormulaR1C1 = SampName
    xlBook.Sheets(1).Range("B12").Font.Bold = True
    xlBook.Sheets(1).Range("C12").FormulaR1C1 = ContName
    xlBook.Sheets(1).Range("C12").Font.Bold = True
    xlBook.Sheets(1).Range("B13").FormulaR1C1 = "=AVERAGE('Statistics'!R3C1:R60000C" & NumMax * 5 & ")"
    xlBook.Sheets(1).Range("B14").FormulaR1C1 = "=STDEV('Statistics'!R3C1:R60000C" & NumMax * 5 & ")"
    xlBook.Sheets(1).Range("C13").FormulaR1C1 = "=AVERAGE('Statistics'!R3C" & NumMax * 5 + 1 & ":R60000C" & NumMax * 10 + 1 & ")"
    xlBook.Sheets(1).Range("C14").FormulaR1C1 = "=STDEV('Statistics'!R3C" & NumMax * 5 + 1 & ":R60000C" & NumMax * 10 + 1 & ")"
    xlBook.Sheets(1).Range("b15").FormulaR1C1 = "=Max(R[-10]C:R[-10]C[4])"
    xlBook.Sheets(1).Range("c15").FormulaR1C1 = "=max(R[-10]C[4]:R[-10]C[8])"
    xlBook.Sheets(1).Range("b16").FormulaR1C1 = "=min(R[-10]C:R[-10]C[4])"
    xlBook.Sheets(1).Range("c16").FormulaR1C1 = "=min(R[-10]C[4]:R[-10]C[8])"
    xlBook.Sheets(1).Range("b17").FormulaR1C1 = "=Sum(R[-10]C:R[-10]C[4])"
    xlBook.Sheets(1).Range("C17").FormulaR1C1 = "=Sum(R[-10]C[4]:R[-10]C[8])"
    xlBook.Sheets(1).Columns("A").EntireColumn.AutoFit
    
    'Text and Cell Formatting
    xlBook.Sheets(1).Range("B1:F1").Font.Bold = True
    xlBook.Sheets(1).Range("B1:F1").HorizontalAlignment = xlCenter
    xlBook.Sheets(1).Range("B1:F1").VerticalAlignment = xlCenter
    xlBook.Sheets(1).Range("B1:F1").WrapText = True
    xlBook.Sheets(1).Range("B1:F1").Orientation = 0
    xlBook.Sheets(1).Range("B1:F1").AddIndent = False
    xlBook.Sheets(1).Range("B1:F1").IndentLevel = 0
    xlBook.Sheets(1).Range("B1:F1").ShrinkToFit = False
    xlBook.Sheets(1).Range("B1:F1").ReadingOrder = xlContext
    xlBook.Sheets(1).Range("B1:F1").MergeCells = True
    xlBook.Sheets(1).Range("B1:F1").FormulaR1C1 = SampName
    xlBook.Sheets(1).Range("G1:K1").Font.Bold = True
    xlBook.Sheets(1).Range("G1:K1").HorizontalAlignment = xlCenter
    xlBook.Sheets(1).Range("G1:K1").VerticalAlignment = xlCenter
    xlBook.Sheets(1).Range("G1:K1").WrapText = True
    xlBook.Sheets(1).Range("G1:K1").Orientation = 0
    xlBook.Sheets(1).Range("G1:K1").AddIndent = False
    xlBook.Sheets(1).Range("G1:K1").IndentLevel = 0
    xlBook.Sheets(1).Range("G1:K1").ShrinkToFit = False
    xlBook.Sheets(1).Range("G1:K1").ReadingOrder = xlContext
    xlBook.Sheets(1).Range("G1:K1").MergeCells = True
    xlBook.Sheets(1).Range("G1:K1").FormulaR1C1 = ContName
    
    'Cell Borders
    xlBook.Sheets(1).Range("B1:K1").Borders(xlDiagonalDown).LineStyle = xlNone
    xlBook.Sheets(1).Range("B1:K1").Borders(xlDiagonalUp).LineStyle = xlNone
    With xlBook.Sheets(1).Range("B1:K1").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With xlBook.Sheets(1).Range("B1:K1").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With xlBook.Sheets(1).Range("B1:K1").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With xlBook.Sheets(1).Range("B1:K1").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With xlBook.Sheets(1).Range("B1:K1").Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With xlBook.Sheets(1).Range("B1:K1").Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    xlBook.Sheets(1).Range("a2:k2").Font.Italic = True
    xlBook.Sheets(1).Range("A12:A17").Font.Bold = True
    
    xlBook.Sheets(1).Range("A2:K7").Borders(xlDiagonalDown).LineStyle = xlNone
    xlBook.Sheets(1).Range("A2:K7").Borders(xlDiagonalUp).LineStyle = xlNone
    With xlBook.Sheets(1).Range("A2:K7").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With xlBook.Sheets(1).Range("A2:K7").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With xlBook.Sheets(1).Range("A2:K7").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With xlBook.Sheets(1).Range("A2:K7").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With xlBook.Sheets(1).Range("A2:K7").Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With xlBook.Sheets(1).Range("A2:K7").Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    xlBook.Sheets(1).Range("A13:C17").Borders(xlDiagonalDown).LineStyle = xlNone
    xlBook.Sheets(1).Range("A13:C17").Borders(xlDiagonalUp).LineStyle = xlNone
    With xlBook.Sheets(1).Range("A13:C17").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With xlBook.Sheets(1).Range("A13:C17").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With xlBook.Sheets(1).Range("A13:C17").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With xlBook.Sheets(1).Range("A13:C17").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With xlBook.Sheets(1).Range("A13:C17").Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With xlBook.Sheets(1).Range("A13:C17").Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    xlBook.Sheets(1).Range("B12:C12").Borders(xlDiagonalDown).LineStyle = xlNone
    xlBook.Sheets(1).Range("B12:C12").Borders(xlDiagonalUp).LineStyle = xlNone
    With xlBook.Sheets(1).Range("B12:C12").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With xlBook.Sheets(1).Range("B12:C12").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With xlBook.Sheets(1).Range("B12:C12").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With xlBook.Sheets(1).Range("B12:C12").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With xlBook.Sheets(1).Range("B12:C12").Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With xlBook.Sheets(1).Range("B12:C12").Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub


Private Sub Shtmkr(NumSheets)
'This sub makes the specified number of clean sheets

'In case this is not a new wkbk
    Do While Worksheets.Count > NumSheets
        Sheets(1).Delete
    Loop
'If it is 'new'
    Do While Worksheets.Count < NumSheets
        ActiveWorkbook.Worksheets.Add After:=Worksheets(Worksheets.Count)
    Loop
'To ensure there is no extra data
    Dim i As Integer
    For i = 1 To NumSheets
        Sheets(i).Cells.ClearContents
    Next
End Sub


Private Sub LotSort(i)
'This sub will sort the pasted data by lot # and then by sample number
'This effectively negates the need to grab only one lot number first
'and then only the other
    Range("B14").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    ActiveWorkbook.Worksheets(i).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(i).Sort.SortFields.Add Key:=Range("b14:xfd14"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(i).Sort
        .SetRange Range("b14:xfd60000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlLeftToRight
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
