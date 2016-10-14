Sub science()
'
' science Macro
'
' hide the work and hide alerts
Application.DisplayAlerts = False
Application.ScreenUpdating = False

' select science sheet and delete unnecessary columns
Sheets("sci").Select
    Columns("A:B").Select
    Selection.Delete Shift:=xlToLeft
    
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:J").Select
    Selection.Delete Shift:=xlToLeft
    Columns("I:S").Select
    Selection.Delete Shift:=xlToLeft
' vlookup password from passwords sheet
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Password"
    
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],passwords!C[9]:C[10],2,FALSE)"
    Range("B2").AutoFill Destination:=Range("B2:B" & lastRow)

    Columns("B:B").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
' turn section column into course column
    Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "SIS Primary Key"
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("H:H").Select
    Selection.Copy
    Columns("I:I").Select
    ActiveSheet.Paste
    Selection.Replace what:=" section", Replacement:="$", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("I1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="$", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
' filter course column to just the classes we want to import
    Columns("J:J").Select
    Selection.Delete Shift:=xlToLeft
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "course"
    Columns("I:I").Select
    Selection.AutoFilter
    ActiveSheet.Range("$I$1:$I$15000").AutoFilter Field:=1, Criteria1:=Array( _
        "AP Biology", "Biology", "Biology LS", "CR Biology", "College in High School Principles of Biology", _
        "Science 3", "Science 4", "Science 5", "Science 6", "Science 6 LS", "Science 7", _
        "Science 7 LS", "Science 8"), Operator:=xlFilterValues

' This bit of code removes the hidden lines
For lp = 256 To 1 Step -1
If Columns(lp).EntireColumn.Hidden = True Then Columns(lp).EntireColumn.Delete Else
Next
For lp = 65536 To 1 Step -1
If Rows(lp).EntireRow.Hidden = True Then Rows(lp).EntireRow.Delete Else
Next

    Columns("I:I").Select
    Selection.Delete Shift:=xlToLeft
    
    Columns("H:I").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
' divide the teacher name column by the space to create a teacher first and teacher last column
    Columns("G:G").Select
    Selection.TextToColumns Destination:=Range("G1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, OtherChar _
        :="$", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Columns("I:I").Select
    Selection.Delete Shift:=xlToLeft
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Teacher Fname"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Teacher Lname"
    Range("H2").Select
    
' crete the homeroom column with this model 16-17_math sec vc8 Barry
    Columns("J:J").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Homeroom"
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "=""16-17_""&RC[-1]&""_""&RC[-2]"
    Range("J2").AutoFill Destination:=Range("J2:J" & lastRow)
    Columns("J:J").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
             
' use find and replace to change section into Sec and also to shorten any classes that have long names

    Columns("J:J").Select
    Selection.Replace what:="Section", Replacement:="Sec", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Columns("J:J").Select
    Selection.Replace what:="College in High School Principles of Biology", Replacement:="College in HS Princip of Bio", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    

' sort by SIS Primary key and then use an if then statement to identify duplicate primary keys and move them to another sheet
    Cells.Select
    ActiveWorkbook.Worksheets("sci").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("sci").Sort.SortFields.Add Key:=Range("A2:A15000") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("sci").Sort
        .SetRange Range("A1:AI15000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]=R[-1]C[-1],TRUE,FALSE)"
    Range("B2").AutoFill Destination:=Range("B2:B" & lastRow)

    Columns("B:B").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "truefalse"
    
    Cells.Select
    Range("A15000").Activate
    ActiveWorkbook.Worksheets("sci").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("sci").Sort.SortFields.Add Key:=Range("B2:B11497") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("sci").Sort
        .SetRange Range("A1:AJ15000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
 

    ActiveWorkbook.RunAutoMacros Which:=xlAutoClose
' filter cut and paste the duplicates to the duplicate sheet
    Sheets("sci dup").Select
    Rows("2:300").Select
    Selection.Delete Shift:=xlUp
    Sheets("sci").Select
    
    If Range("B2").Value = "True" Then
    Columns("B:B").Select
    Selection.AutoFilter
    ActiveSheet.Range("$B$1:$B$15000").AutoFilter Field:=1, Criteria1:="TRUE"
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Sheets("sci dup").Select
    Range("A2").Select
    ActiveSheet.Paste
    Sheets("sci").Select
    
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp

    
    Columns("B:B").Select
    Selection.AutoFilter
    Cells.Select
    ActiveWorkbook.Worksheets("sci").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("sci").Sort.SortFields.Add Key:=Range("J2:J15000"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("sci").Sort
        .SetRange Range("A1:O15000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A1").Select
    End If
    
' hide the work and hide alerts
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
'
End Sub


