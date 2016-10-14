Sub quote()
'
' quote Macro
'

'
    Range("E6:K14").Select
    ActiveCell.FormulaR1C1 = "=INDEX(R[-5]C[74]:R[4]C[74], RANDBETWEEN(1,10))"
    Range("E6:K14").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False


End Sub
