Attribute VB_Name = "PoreDiameterDataCompiler"
Option Explicit
Dim xlBook1 As Excel.Workbook
Dim xlBook2 As Excel.Workbook
Dim xlSheet1 As Excel.Worksheet
Dim xlSheet2 As Excel.Worksheet


Sub PoreDiameterMaster()
Attribute PoreSizeMaster.VB_Description = "Working Draft. For Gautam.  Compiles Pore diameter information from PaxIt reported files.  Run after Renaming macro to minimize errors."
Attribute PoreSizeMaster.VB_ProcData.VB_Invoke_Func = " \n14"
'This Macro should be run on a blank Excel file only!!!
'Will delete ALL prior data!
'Written by Ben Hauch in April 2010
'Updated by Ben Hauch in August 2010
    'Rewritten to handle from 1 to 5 test groups
    'Require input file directories organized by group
    'Sheet selection based on iteration (i.e. group) and depth
    'Error handling
    'Improved commenting
    'Future work: netpgformat update for more than 2 groups

Set xlBook1 = ActiveWorkbook
Dim strErrFiles(1 To 65000) As String
Dim blErrFlag As Boolean
Dim strErrOutput As String
Dim blPasteFlag As Boolean
Dim i As Long
Dim j As Long
Dim k As Long
Dim m As Long

'The general structure of this macro is as follows _
    A) Determine applicable test request information _
        i)MT number _
        ii)Number of test groups _
        iii)Name of each test group _
        iv)Maximum test group size _
    B) With each test group _
        i)Get user-provided location of excel files _
        ii)For each file _
            a)Determine sample number _
            b)Determine lot number _
            c)Use file data to determine where to paste _
            d)Copy from file _
            e)Paste into "master" file _
            f)Close file _
        iii)Compile stats for each sample _
    C) Compile stats for each group _
    D) Summarize results
    
'For any given section, we first declare intuitive variables _
 and then we set their value in following lines.
    
'Section A: Determine save name & set key internal variables
    Dim WkbkName1 As String
    Dim NumGroups
    Dim strArrGroupNames() As String
    Dim NumMax
    
    WkbkName1 = InputBox("What is the test request designation?" & Chr(13) & "E.g. MTXXXX", "Workbook Name")
If WkbkName1 = vbNullString Then GoTo OnErr1                            'If user clicks "Cancel" InputBox assigns a value of vbNullString ("") to the variable
    NumGroups = InputBox("How many groups are there?" & Chr(13) & "E.g. 1 or 2 or ...", "Number of Test Groups in Test Request")
If NumGroups = vbNullString Then GoTo OnErr1
    ReDim Preserve strArrGroupNames(1 To NumGroups)
    For i = 1 To NumGroups
        strArrGroupNames(i) = InputBox("What is the unique name of " & WkbkName1 & " group #" & i & Chr(13) & "E.g. A or Ti or CoCr or PPS...", "Group Name")
    Next i
If strArrGroupNames(1) = vbNullString Then GoTo OnErr1
    NumMax = InputBox("What is the maximum number of test samples in any single group?" & Chr(13) & "E.g. 5 or 15 or ...")
If NumMax = vbNullString Then GoTo OnErr1

'Section B: Data Transfer & Agglomeration

    'Section B.i: Obtain necessary user inputs
    Dim strArrGroupFilePaths() As String
    ReDim Preserve strArrGroupFilePaths(1 To NumGroups)
    For i = 1 To NumGroups
        strArrGroupFilePaths(i) = InputBox("Please find & paste the directory path to" & Chr(13) & _
                                "the excel files for group #" & i & Chr(13) & "E.g. A:\Gautam\YourName\PoreDiameterFiles\MTXXXX\Group" & i, strArrGroupNames(i))
    Next i
If strArrGroupFilePaths(1) = vbNullString Then GoTo OnErr1

'To speed analysis, we turn off screen updating and automatic calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
'We want a total of 5 clean sheets for each group, plus _
 1 sheet for the data agglomeration and 1 sheet for the _
 overall data summary. We call a separate sub defined _
 below to do this job, and we pass it the number of sheets _
 we need.
    Call Shtmkr(2 + NumGroups * 5)

'We require named sheets to eliminate any chance of placing _
 the data in the wrong sheet, so we call another sub defined _
 below
    Call ShtsNaming(strArrGroupNames, NumMax, xlBook1)
 

    'Section B.ii: Data Transfer
    Dim fso As Object
    Dim strPasteBook As String
    Dim pthlngth As String
    Dim strArr(1 To 65000) As String, strXLPath As String
    Dim Cutpoint As Integer
    Dim LotNum As String
    Dim SampID As String, FileID As String, ColTitle As String
    Dim RgResume() As String
        
    strPasteBook = xlBook1.Name                                     'This is so the right wkbk is pasted into
        
    For m = 1 To NumGroups                                          'FOR EACH GROUP
        pthlngth = Len(strArrGroupFilePaths(m))
        Let strXLPath = Dir$(strArrGroupFilePaths(m) & "\*" & "*.xlsx")  'Note that ONLY Excel 2007 files will be opened!
        Do While strXLPath <> vbNullString
            Let i = i + 1
            Let strArr(i) = strArrGroupFilePaths(m) & "\" & strXLPath
            Let strXLPath = Dir$()
        Loop
        Set fso = CreateObject("Scripting.FileSystemObject")
        Call recurseSubFolders(fso.GetFolder(strArrGroupFilePaths(m)), strArr(), i) 'This line calls the Private Sub defined below
        Set fso = Nothing
        ReDim Preserve RgResume(1 To NumMax, 3 To xlBook1.Worksheets.Count)
        
        Let i = 1                                                       'File counter
        Let j = 1                                                       'Error line counter
        Do While strArr(i) <> vbNullString                              'FOR EACH FILE
            Workbooks.Open strArr(i), ReadOnly:=True                    'Each sample is ID'd by the combo of two pieces
            Set xlBook2 = Workbooks(Workbooks.Count)                    'of data: the lot number and the sample ID
            Set xlSheet2 = xlBook2.Sheets("Lines")
            LotNum = Range("C21")
            FileID = Trim(Range("C19"))
            SampID = RTrim(Left(FileID, Len(FileID) - 3))               'Get only the sample name
            ColTitle = LotNum & " - " & SampID                          'Each sample has its own column, this
                                                                        'helps ID where to place the copied data
            'Copy the relevant data
            If xlSheet2.Range("H7:H8").Count < 2 Then                            'Check to see if there is more than one data point
                Range("H7").Copy                                        'Only one data point?   Then select only it.
            Else
                Range(xlSheet2.Range("H7"), xlSheet2.Range("H7").End(xlDown)).Copy            'This copies from h7 to the final entry -
            End If                                                      'only one entry would cause an error
            
            'Figure out what sheet to activate and then paste into
            Cutpoint = InStr(1, strArr(i), "micron", vbTextCompare)
            If xlSheet2.Range("C22") = "" Then                          'Check Operator # for depth info
                Select Case Right(Left(strArr(i), Cutpoint - 2), 1) 'The case is chosen by inspecting the filename
                    Case 0
                    Set xlSheet1 = xlBook1.Sheets(strArrGroupNames(m) & " 000")
                    Case 7
                    Set xlSheet1 = xlBook1.Sheets(strArrGroupNames(m) & " 127")
                    Case 4
                    Set xlSheet1 = xlBook1.Sheets(strArrGroupNames(m) & " 254")
                    Case 1
                    Set xlSheet1 = xlBook1.Sheets(strArrGroupNames(m) & " 381")
                    Case 8
                    Set xlSheet1 = xlBook1.Sheets(strArrGroupNames(m) & " 508")
                    Case Else
                        strErrFiles(j) = xlBook2.Name
                        Let j = j + 1
                        blErrFlag = True
                        Workbooks(xlBook2.Name).Close
                        GoTo NextFile
                End Select
            Else                                                'A non-blank depth field, so use that
                If Right(Left(strArr(i), Cutpoint - 2), 1) = Right(xlSheet2.Range("C22").Formula, 1) Then 'but only if the value matches the directory or filename cues
                    Select Case Left(xlSheet2.Range("C22").Formula, 1)
                        Case 0
                        Set xlSheet1 = xlBook1.Sheets(strArrGroupNames(m) & " 000")
                        Case 1
                        Set xlSheet1 = xlBook1.Sheets(strArrGroupNames(m) & " 127")
                        Case 2
                        Set xlSheet1 = xlBook1.Sheets(strArrGroupNames(m) & " 254")
                        Case 3
                        Set xlSheet1 = xlBook1.Sheets(strArrGroupNames(m) & " 381")
                        Case 5
                        Set xlSheet1 = xlBook1.Sheets(strArrGroupNames(m) & " 508")
                        Case Else
                            strErrFiles(j) = xlBook2.Name
                            Let j = j + 1
                            blErrFlag = True
                            Workbooks(xlBook2.Name).Close
                            GoTo NextFile
                    End Select
                Else                                                                'If depth doesn't match cues, the Excel file needs to be inspected
                strErrFiles(j) = xlBook2.Name                                       'and have the depth field corrected manually
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
    Next m
    
'Section B.iii: Sample formatting & stats analysis
    Workbooks(strPasteBook).Activate
    Call ShtsFormat(NumMax, xlBook1)
    
'Section C: Group formatting & stats analysis & _
 Section D: Summary of results
    Call netpgformat(strArrGroupNames, NumMax, NumGroups, xlBook1)
    
'Now we re-enable screen updating and automatic calculations
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

'Finally, the macro saves the file according to the given test request designation
    xlBook1.SaveAs (WkbkName1 & " - Pore Size Master File.xlsx")
    MsgBox ("The macro has finished running.  Don't forget to doublecheck the group formulas for this page from the data on the Statistics page.")
    If blErrFlag = True Then GoTo OnErr2
    Exit Sub

OnErr1:
    MsgBox ("Macro canceled by user.")
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
OnErr2:
    ActiveWorkbook.Worksheets.Add After:=Worksheets(Worksheets.Count)
    Let j = 1
    Do While strErrFiles(j) <> vbNullString
        Worksheets(Worksheets.Count).Cells(j + 3, 1).FormulaR1C1 = strErrFiles(j)
        j = j + 1
    Loop
    Worksheets(Worksheets.Count).Range("A1").FormulaR1C1 = "The following files were not analyzed properly;"
    Worksheets(Worksheets.Count).Range("A2").FormulaR1C1 = "Open each and add data manually"
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
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
                    
                    
Private Sub ShtsNaming(GroupNameArray, NumMax, xlBook)
'There is not a well-defined number of sheets to expect, but it is _
 expected that any reasonable test will have less than 5 groups. _
 If there is ever a case with more than five groups, simply expand _
 this section.
    
    Dim i As Integer
    Dim xlsheet As Excel.Worksheet
    
    xlBook.Sheets(1).Name = "Net"
    xlBook.Sheets(2).Name = "Statistics"
    xlBook.Sheets(3).Name = GroupNameArray(1) & " 000"
    xlBook.Sheets(4).Name = GroupNameArray(1) & " 127"
    xlBook.Sheets(5).Name = GroupNameArray(1) & " 254"
    xlBook.Sheets(6).Name = GroupNameArray(1) & " 381"
    xlBook.Sheets(7).Name = GroupNameArray(1) & " 508"       'There are always at least 7 sheets
    
    If Worksheets.Count > 7 Then
        xlBook.Sheets(8).Name = GroupNameArray(2) & " 000"
        xlBook.Sheets(9).Name = GroupNameArray(2) & " 127"
        xlBook.Sheets(10).Name = GroupNameArray(2) & " 254"
        xlBook.Sheets(11).Name = GroupNameArray(2) & " 381"
        xlBook.Sheets(12).Name = GroupNameArray(2) & " 508"
    End If
    If Worksheets.Count > 13 Then
        xlBook.Sheets(13).Name = GroupNameArray(3) & " 000"
        xlBook.Sheets(14).Name = GroupNameArray(3) & " 127"
        xlBook.Sheets(15).Name = GroupNameArray(3) & " 254"
        xlBook.Sheets(16).Name = GroupNameArray(3) & " 381"
        xlBook.Sheets(17).Name = GroupNameArray(3) & " 508"
    End If
    If Worksheets.Count > 17 Then
        xlBook.Sheets(18).Name = GroupNameArray(4) & " 000"
        xlBook.Sheets(19).Name = GroupNameArray(4) & " 127"
        xlBook.Sheets(20).Name = GroupNameArray(4) & " 254"
        xlBook.Sheets(21).Name = GroupNameArray(4) & " 381"
        xlBook.Sheets(22).Name = GroupNameArray(4) & " 508"
    End If
    If Worksheets.Count > 22 Then
        xlBook.Sheets(23).Name = GroupNameArray(5) & " 000"
        xlBook.Sheets(24).Name = GroupNameArray(5) & " 127"
        xlBook.Sheets(25).Name = GroupNameArray(5) & " 254"
        xlBook.Sheets(26).Name = GroupNameArray(5) & " 381"
        xlBook.Sheets(27).Name = GroupNameArray(5) & " 508"
    End If
End Sub
                        
                        
Private Sub recurseSubFolders(ByRef Folder As Object, ByRef strArr() As String, ByRef i As Long)
'This is the grunt work code for finding all the excel07 files needed for analysis

Dim SubFolder As Object
Dim strName As String
For Each SubFolder In Folder.SubFolders
    Let strName = Dir$(SubFolder.Path & "\*" & "*.xlsx")
    Do While strName <> vbNullString
        Let i = i + 1
        Let strArr(i) = SubFolder.Path & "\" & strName
        Let strName = Dir$()
    Loop
    Call recurseSubFolders(SubFolder, strArr(), i)
Next
End Sub
                    
                    
Private Sub ShtsFormat(NumMax, xlBook)
'Now we fill in the box descriptors
    Dim xlsheet As Excel.Worksheet
    Dim i As Long
    
    xlBook.Sheets("Statistics").Activate
    xlBook.Sheets("Statistics").Range("A1").Select      'If this is not done, issues can arise with the agglomeration steps
    For i = 3 To xlBook.Worksheets.Count 'The other sheets are the summary and the statistics pages and have special formatting
        Set xlsheet = Sheets(i) 'Select the i-th sheet
        
        'Begin cell labeling
        xlsheet.Range("A1").FormulaR1C1 = "Pore Diameter - " & Right(ActiveSheet.Name, 3) & " µm"
        xlsheet.Range("A3").FormulaR1C1 = "Average Pore Diameter"
        xlsheet.Range("A4").FormulaR1C1 = "Std. Dev."
        xlsheet.Range("A5").FormulaR1C1 = "Max"
        xlsheet.Range("A6").FormulaR1C1 = "Min"
        xlsheet.Range("A7").FormulaR1C1 = "Pore Count"
        xlsheet.Range("A10").FormulaR1C1 = "Sample Avg. Pore Diam."
        xlsheet.Range("A11").FormulaR1C1 = "Sample Std. Dev."
        xlsheet.Range("A12").FormulaR1C1 = "Sample Max"
        xlsheet.Range("A13").FormulaR1C1 = "Sample Min"
        xlsheet.Range("A14").FormulaR1C1 = "Sample ID"
        xlsheet.Range("A1:A14").Font.Bold = True
        xlsheet.Range("A1:A14").HorizontalAlignment = xlRight
        xlsheet.Range("A1:A14").VerticalAlignment = xlCenter
        xlsheet.Range("A1:A14").Orientation = 0
        xlsheet.Range("A1:A14").AddIndent = False
        xlsheet.Range("A1:A14").IndentLevel = 0
        xlsheet.Range("A1:A14").ShrinkToFit = False
        xlsheet.Range("A1:A14").ReadingOrder = xlContext
        xlsheet.Range("A1:A14").MergeCells = False
        xlsheet.Rows("14:14").Font.Bold = True
        xlsheet.Rows("14:14").HorizontalAlignment = xlCenter
        xlsheet.Rows("14:14").VerticalAlignment = xlCenter
        xlsheet.Rows("14:14").WrapText = True
        xlsheet.Rows("14:14").Orientation = 0
        xlsheet.Rows("14:14").AddIndent = False
        xlsheet.Rows("14:14").IndentLevel = 0
        xlsheet.Rows("14:14").ShrinkToFit = False
        xlsheet.Rows("14:14").ReadingOrder = xlContext
        xlsheet.Rows("14:14").MergeCells = False
        xlsheet.Rows("14:14").EntireRow.AutoFit
        xlsheet.Range("C3:C6").FormulaR1C1 = "µm"
        xlsheet.Range("C7").FormulaR1C1 = "pores"
        xlsheet.Columns("A").EntireColumn.AutoFit
        xlsheet.Columns("B:ZZ").ColumnWidth = 8.43
        
        'All the labeling and formatting is now done,
        'and calculation formulas must be entered.
        xlsheet.Range("B13").FormulaR1C1 = "=MIN(R15C:R60000C)"
        xlsheet.Range("B12").FormulaR1C1 = "=MAX(R15C:R60000C)"
        xlsheet.Range("B11").FormulaR1C1 = "=STDEV(R15C:R60000C)"
        xlsheet.Range("B10").FormulaR1C1 = "=AVERAGE(R15C:R60000C)"
        xlsheet.Range("B10:B13").AutoFill Destination:=xlsheet.Range(Cells(10, 2), Cells(13, NumMax + 1))
        xlsheet.Range("B7").FormulaR1C1 = "=COUNT(R15C2:R60000C[1700])"     'Num datapoints
        xlsheet.Range("B6").FormulaR1C1 = "=MIN(R13C2:R13C1702)"            'Global minimum pore size
        xlsheet.Range("B5").FormulaR1C1 = "=MAX(R12C2:R12C1702)"            'Global maximum pore size
        xlsheet.Range("B4").FormulaR1C1 = "=STDEV(R15C2:R60000C1702)"       'Global std. deviation
        xlsheet.Range("B3").FormulaR1C1 = "=AVERAGE(R15C2:R60000C1702)"     'Global average
        xlsheet.Cells.FormatConditions.Delete
        'Sort the input data according to the column title - this will separate the lot numbers
        Call LotSort(i)
        xlsheet.Range(Cells(14, 2), Cells(14, 2).SpecialCells(xlLastCell)).Copy
        Sheets("Statistics").Select                                 'Agglomeration of the data for statistics
        ActiveCell.FormulaR1C1 = xlsheet.Name
        ActiveCell.Offset(1, 0).Select
        ActiveSheet.paste
        ActiveCell.Offset(-1, NumMax).Select                        'Do not overwrite
    Next i
End Sub

                                    
Private Sub netpgformat(GroupNameArray, NumMax, NumGroups, xlBook)
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
    xlBook.Sheets(1).Range("B3:B7").FormulaR1C1 = "='" & GroupNameArray(1) & " 000'!RC"
    xlBook.Sheets(1).Range("C3:C7").FormulaR1C1 = "='" & GroupNameArray(1) & " 127'!RC[-1]"
    xlBook.Sheets(1).Range("D3:D7").FormulaR1C1 = "='" & GroupNameArray(1) & " 254'!RC[-2]"
    xlBook.Sheets(1).Range("E3:E7").FormulaR1C1 = "='" & GroupNameArray(1) & " 381'!RC[-3]"
    xlBook.Sheets(1).Range("F3:F7").FormulaR1C1 = "='" & GroupNameArray(1) & " 508'!RC[-4]"
    xlBook.Sheets(1).Range("B12").FormulaR1C1 = GroupNameArray(1)
    xlBook.Sheets(1).Range("B12").Font.Bold = True
    xlBook.Sheets(1).Range("B13").FormulaR1C1 = "=AVERAGE('Statistics'!R3C1:R60000C" & NumMax * 5 & ")"
    xlBook.Sheets(1).Range("B14").FormulaR1C1 = "=STDEV('Statistics'!R3C1:R60000C" & NumMax * 5 & ")"
    xlBook.Sheets(1).Range("B15").FormulaR1C1 = "=Max(R[-10]C:R[-10]C[4])"
    xlBook.Sheets(1).Range("B16").FormulaR1C1 = "=min(R[-10]C:R[-10]C[4])"
    xlBook.Sheets(1).Range("B17").FormulaR1C1 = "=Sum(R[-10]C:R[-10]C[4])"
    xlBook.Sheets(1).Range("B1:F1").MergeCells = True
    xlBook.Sheets(1).Range("B1:F1").FormulaR1C1 = GroupNameArray(1)
    If NumGroups > 1 Then
        xlBook.Sheets(1).Range("G2").FormulaR1C1 = "'000"
        xlBook.Sheets(1).Range("H2").FormulaR1C1 = "127"
        xlBook.Sheets(1).Range("I2").FormulaR1C1 = "254"
        xlBook.Sheets(1).Range("J2").FormulaR1C1 = "381"
        xlBook.Sheets(1).Range("K2").FormulaR1C1 = "508"
        xlBook.Sheets(1).Range("G3:G7").FormulaR1C1 = "='" & GroupNameArray(2) & " 000'!RC[-5]"
        xlBook.Sheets(1).Range("H3:H7").FormulaR1C1 = "='" & GroupNameArray(2) & " 127'!RC[-6]"
        xlBook.Sheets(1).Range("I3:I7").FormulaR1C1 = "='" & GroupNameArray(2) & " 254'!RC[-7]"
        xlBook.Sheets(1).Range("J3:J7").FormulaR1C1 = "='" & GroupNameArray(2) & " 381'!RC[-8]"
        xlBook.Sheets(1).Range("K3:K7").FormulaR1C1 = "='" & GroupNameArray(2) & " 508'!RC[-9]"
        xlBook.Sheets(1).Range("C12").FormulaR1C1 = GroupNameArray(2)
        xlBook.Sheets(1).Range("C12").Font.Bold = True
        xlBook.Sheets(1).Range("C13").FormulaR1C1 = "=AVERAGE('Statistics'!R3C" & NumMax * 5 + 1 & ":R60000C" & NumMax * 10 + 1 & ")"
        xlBook.Sheets(1).Range("C14").FormulaR1C1 = "=STDEV('Statistics'!R3C" & NumMax * 5 + 1 & ":R60000C" & NumMax * 10 + 1 & ")"
        xlBook.Sheets(1).Range("C15").FormulaR1C1 = "=max(R[-10]C[4]:R[-10]C[8])"
        xlBook.Sheets(1).Range("C16").FormulaR1C1 = "=min(R[-10]C[4]:R[-10]C[8])"
        xlBook.Sheets(1).Range("C17").FormulaR1C1 = "=Sum(R[-10]C[4]:R[-10]C[8])"
        xlBook.Sheets(1).Range("G1:K1").MergeCells = True
        xlBook.Sheets(1).Range("G1:K1").FormulaR1C1 = GroupNameArray(2)
    End If
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
    xlBook.Sheets(1).Range("G1:K1").Font.Bold = True
    xlBook.Sheets(1).Range("G1:K1").HorizontalAlignment = xlCenter
    xlBook.Sheets(1).Range("G1:K1").VerticalAlignment = xlCenter
    xlBook.Sheets(1).Range("G1:K1").WrapText = True
    xlBook.Sheets(1).Range("G1:K1").Orientation = 0
    xlBook.Sheets(1).Range("G1:K1").AddIndent = False
    xlBook.Sheets(1).Range("G1:K1").IndentLevel = 0
    xlBook.Sheets(1).Range("G1:K1").ShrinkToFit = False
    xlBook.Sheets(1).Range("G1:K1").ReadingOrder = xlContext
    
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
    xlBook.Sheets(1).Range("A2:K2").Font.Italic = True
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
