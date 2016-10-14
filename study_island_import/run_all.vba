Sub run_all()
'
' run_all Macro
'
' hide the work and hide alerts
Application.DisplayAlerts = False
Application.ScreenUpdating = False


Call english
Call mathematics
Call science
Call social_studies
Call Keystone
Call csv_save

' hide the work and hide alerts
Application.DisplayAlerts = True
Application.ScreenUpdating = True


End Sub
