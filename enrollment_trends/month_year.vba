Sub By_Year()
'
' By_Year Macro
' Collapse all fields to show yearly data
'

'
    Sheets("E State").Select
    Range("C1").Select
    Range("B3").Select
    ActiveSheet.PivotTables("PivotTable6").PivotFields("eyear").ShowDetail = False
    Range("B18").Select
    ActiveSheet.PivotTables("PivotTable7").PivotFields("eyear").ShowDetail = False
    Sheets("W State").Select
    Range("B3").Select
    ActiveSheet.PivotTables("PivotTable2").PivotFields("wyear").ShowDetail = False
    Range("B22").Select
    ActiveSheet.PivotTables("PivotTable3").PivotFields("wyear").ShowDetail = False
    Sheets("W Internal").Select
    Range("B4").Select
    ActiveSheet.PivotTables("PivotTable4").PivotFields("wyear").ShowDetail = False
    ActiveWindow.SmallScroll Down:=39
    Range("B46").Select
    ActiveSheet.PivotTables("PivotTable5").PivotFields("wyear").ShowDetail = False
    Sheets("Length of Stay").Select
    Range("B3").Select
    ActiveSheet.PivotTables("PivotTable3").PivotFields("wyear").ShowDetail = False
    Range("B16").Select
    ActiveSheet.PivotTables("PivotTable1").PivotFields("wyear").ShowDetail = False
    Sheets("E State").Select
    Range("B1").Select
End Sub
Sub By_Month()
'
' By_Month Macro
' Expand entire field to show monthly data
'

'
    Sheets("E State").Select
    Range("C1").Select
    Range("B3").Select
    ActiveSheet.PivotTables("PivotTable6").PivotFields("eyear").ShowDetail = True
    Range("B18").Select
    ActiveSheet.PivotTables("PivotTable7").PivotFields("eyear").ShowDetail = True
    Sheets("W State").Select
    Range("B3").Select
    ActiveSheet.PivotTables("PivotTable2").PivotFields("wyear").ShowDetail = True
    Range("B22").Select
    ActiveSheet.PivotTables("PivotTable3").PivotFields("wyear").ShowDetail = True
    Sheets("W Internal").Select
    Range("B4").Select
    ActiveSheet.PivotTables("PivotTable4").PivotFields("wyear").ShowDetail = True
    Range("B46").Select
    ActiveSheet.PivotTables("PivotTable5").PivotFields("wyear").ShowDetail = True
    ActiveWindow.SmallScroll Down:=-69
    Sheets("Length of Stay").Select
    Range("B3").Select
    ActiveSheet.PivotTables("PivotTable3").PivotFields("wyear").ShowDetail = True
    Range("B16").Select
    ActiveSheet.PivotTables("PivotTable1").PivotFields("wyear").ShowDetail = True
    Sheets("E State").Select
    Range("B1").Select
End Sub
