Attribute VB_Name = "FULL"
Sub FULL()

Application.Calculate
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Call DATA.DATA
Call Pagedegarde.Pagedegarde
Call RESULTATS.RESULTATS
Call EVOLUTIONEFFECTIFS.EVOLUTIONEFFECTIFS
Call DEMOGRAPHIE.DEMOGRAPHIE
Call PRESTATIONSREGLEES_NEW.PRESTATIONSREGLEES_NEW
Call PRESTATIONSREGLEES_NEW_COMPAR.PRESTATIONSREGLEES_NEW_COMPAR
Call PRESTATIONSREGLEESGRAPH.PRESTATIONSREGLEESGRAPH
Call Pagedegarde.Pagedegarde

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic


End Sub

