Sub Keystone()
'
' Keystone Macro
'
' hide the work and hide alerts
Application.DisplayAlerts = False
Application.ScreenUpdating = False

Sheets("keystone").Select
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
' vlookup password from passwords sheet
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
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("H:H").Select
    Selection.Copy
    Range("J1").Select
    ActiveSheet.Paste
    Columns("J:J").Select
    Selection.Replace what:=" section", Replacement:="$", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("J:J").Select
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("J1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="$", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Columns("K:L").Select
    Selection.Delete Shift:=xlToLeft
    
' filter course column to just the classes we want to import
    Columns("J:J").Select
    Selection.AutoFilter
    ActiveSheet.Range("$J$1:$J$15000").AutoFilter Field:=1, Criteria1:=Array( _
        "Keystone Algebra I Fall 2016", "Keystone Algebra I Spring 2017", _
        "Keystone Biology Fall 2016", "Keystone Biology Spring 2017", _
        "Keystone English Literature Fall 2016", "Keystone English Literature Spring 2017"), Operator:=xlFilterValues
        
' This bit of code removes the hidden lines
For lp = 256 To 1 Step -1
If Columns(lp).EntireColumn.Hidden = True Then Columns(lp).EntireColumn.Delete Else
Next
For lp = 65536 To 1 Step -1
If Rows(lp).EntireRow.Hidden = True Then Rows(lp).EntireRow.Delete Else
Next


' divide the teacher name column by the space to create a teacher first and teacher last column
    Columns("F:F").Select
    Selection.TextToColumns Destination:=Range("F1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, OtherChar _
        :="$", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Teacher Fname"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Teacher Lname"

    
' crete the homeroom column with this model 16-17_math sec vc8 Barry
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Homeroom"
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "=""16-17_""&RC[-3]&""_""&RC[-4]"
    Range("K2").AutoFill Destination:=Range("K2:K" & lastRow)
    Columns("K:K").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
             
' use find and replace to change section into Sec and also to shorten any classes that have long names

    Columns("K:K").Select
    Selection.Replace what:="Section", Replacement:="Sec", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
          
    Columns("K:K").Select
    Selection.Replace what:="Algebra", Replacement:="Alg", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Columns("K:K").Select
    Selection.Replace what:="English Literature", Replacement:="Eng Lit", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    Columns("J:J").Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "course"


' delete the rows in the keystone subjects sheets

    Sheets("alg").Select
    Rows("2:1000").Select
    Selection.Delete Shift:=xlUp
    
    Sheets("bio").Select
    Rows("2:1000").Select
    Selection.Delete Shift:=xlUp
    
    Sheets("lit").Select
    Rows("2:1000").Select
    Selection.Delete Shift:=xlUp
    
    Sheets("keystone").Select
    Columns("J:J").Select
    Selection.AutoFilter
    
'filter the keystone file and distribute it to the different subjects
    
    ActiveSheet.Range("$J$1:$J$15000").AutoFilter Field:=1, Criteria1:= _
        "=Keystone Algebra I Fall 2016", Operator:=xlOr, Criteria2:= _
        "=Keystone Algebra I Spring 2017"
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("alg").Select
    Range("A2").Select
    ActiveSheet.Paste
    Sheets("keystone").Select

    ActiveSheet.Range("$J$1:$J$15000").AutoFilter Field:=1, Criteria1:= _
        "=Keystone Biology Fall 2016", Operator:=xlOr, Criteria2:= _
        "=Keystone Biology Spring 2017"

    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("bio").Select
    Range("A2").Select
    ActiveSheet.Paste
    
    
    Sheets("keystone").Select

    ActiveSheet.Range("$J$1:$J$15000").AutoFilter Field:=1, Criteria1:= _
        "Keystone English Literature Fall 2016", Operator:=xlOr, Criteria2:= _
        "=Keystone English Spring 2017"
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("lit").Select
    Range("A2").Select
    ActiveSheet.Paste
    Sheets("keystone").Select
    
    Cells.Select
    Range("E1").Activate
    Application.CutCopyMode = False
    Selection.AutoFilter
    
' hide the work and hide alerts
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub
