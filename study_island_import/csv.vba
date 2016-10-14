Sub csv_save()
'
' csv_save Macro

' hide the work and hide alerts
Application.DisplayAlerts = False
Application.ScreenUpdating = False

ActiveWorkbook.Save

Dim WS As Excel.Worksheet
Dim SaveToDirectory As String

    SaveToDirectory = "C:\Users\lbarry\Desktop\New folder\"

    For Each WS In ThisWorkbook.Worksheets
        WS.SaveAs SaveToDirectory & WS.Name, xlCSV
    Next

' hide the work and hide alerts
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

