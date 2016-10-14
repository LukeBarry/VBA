Sub format_data()
'
' format_data Macro
' format data
'

'
    Sheets("paste data").Select
    Cells.Select
    Columns("D:T").Select
    Range("T1").Activate
    Selection.Replace What:="null", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("G:G").Select
    Selection.Copy
    Columns("U:U").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("U1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="/", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    Columns("V:V").Select
    Selection.Delete Shift:=xlToLeft
    Columns("I:I").Select
    Selection.Copy
    Columns("W:W").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("W1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="/", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    Columns("X:X").Select
    Selection.Delete Shift:=xlToLeft
    Columns("U:X").Select
    Range("X1").Activate
    Selection.NumberFormat = "0"
    ActiveWindow.SmallScroll ToRight:=2
    Range("Y2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-18]="""",TODAY(),RC[-18])"
    Range("Z2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-17]="""",TODAY(),RC[-17])"
    Range("Z3").Select
    ActiveWindow.SmallScroll ToRight:=3
    Columns("Y:Z").Select
    Selection.NumberFormat = "m/d/yyyy"
    Range("AA2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]-RC[-2]"
    Range("AB2").Select
    Sheets("Cover").Select

    Range("AA1").Select
    Selection.Copy
    Sheets("paste data").Select
    Range("AB2").Select
    ActiveSheet.Paste
    Range("Y2:AB2").Select
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("Y2:AB61443")
    Range("Y2:AB61443").Select
 
    Rows("2:2").Select
    Range("F2").Activate
    Range(Selection, Selection.End(xlDown)).Select
    Rows("2:270").Select
    Range("F2").Activate
    ActiveWindow.SmallScroll Down:=6
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
 
    Rows("2:2").Select
    Range("F2").Activate
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    Sheets("PT Data").Select
    Rows("2:2").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Sheets("paste data").Select
    ActiveWindow.ScrollRow = 61282
    ActiveWindow.ScrollRow = 60026
    ActiveWindow.ScrollRow = 57653
    ActiveWindow.ScrollRow = 33922
    ActiveWindow.ScrollRow = 17032
    ActiveWindow.ScrollRow = 5725
    ActiveWindow.ScrollRow = 1677
    ActiveWindow.ScrollRow = 560
    ActiveWindow.ScrollRow = 2
    Rows("2:2").Select
    Range("F2").Activate
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    Sheets("PT Data").Select
    Range("A2").Select
    ActiveSheet.Paste

    Application.CutCopyMode = False
    ActiveWorkbook.RefreshAll
    Sheets("paste data").Select
    Cells.Select
    Selection.ClearContents
    Sheets("Cover").Select
    Range("A1").Select
End Sub
