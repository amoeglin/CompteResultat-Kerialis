Attribute VB_Name = "DEMOGRAPHIE"
Sub DEMOGRAPHIE()

Sheets("Démographie").Select

'VARIABLES

'RAZ
        Sheets("Démographie").Select

        For i = 1 To 3
        For j = 1 To 5
        Cells(13 + i, j + 3) = ""
        Next j
        Next i
        
        For i = 1 To 5
        For j = 1 To 2
        Cells(36 + i, j + 3) = ""
        Next j
        Next i
        
        For i = 1 To 10
        For j = 1 To 3
        Cells(13 + i, j + 11) = ""
        Next j
        Next i
      
    
'REMPLISSAGE


        Sheets("DATA PREST").Select
        
        i = 2
        If Cells(2, 4) = "" Then
        Exit Sub
        End If
        ANNEE2 = Cells(2, 4)
        While Cells(i, 4) = ANNEE2
        ANNEE1 = Cells(i + 1, 4)
        i = i + 1
        Wend
        
        If ANNEE1 = "" Then
        ANNEE1 = ANNEE2
        ANNEE2 = ""
        End If
        
        
        Sheets("DATA DEMO").Select

        FDEMO = "DATA DEMO"
        ANNEEDEMO = "A:A"
        SOCIETEDEMO = "B:B"
        SEXEDEMO = "D:D"
        AGEDEMO = "H:H"
        DEMO = "G:G"
        LIENDEMO = "E:E"
        TRANCHEDEMO = "F:F"
        COLDEMO = "J:J"
        
   
        Sheets("Démographie").Select

        Cells(14, 4) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE2, Worksheets(FDEMO).Range(LIENDEMO), "Assuré", Worksheets(FDEMO).Range(SEXEDEMO), "Masculin")
        Cells(15, 4) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE2, Worksheets(FDEMO).Range(LIENDEMO), "Assuré", Worksheets(FDEMO).Range(SEXEDEMO), "Féminin")
        Cells(16, 4) = Cells(14, 4) + Cells(15, 4)
        
        If Cells(14, 4) > 0 Then
        Cells(14, 5) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(AGEDEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE2, Worksheets(FDEMO).Range(LIENDEMO), "Assuré", Worksheets(FDEMO).Range(SEXEDEMO), "Masculin") / Cells(14, 4)
        End If
        
        If Cells(15, 4) > 0 Then
        Cells(15, 5) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(AGEDEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE2, Worksheets(FDEMO).Range(LIENDEMO), "Assuré", Worksheets(FDEMO).Range(SEXEDEMO), "Féminin") / Cells(15, 4)
        End If
        
        If Cells(16, 4) > 0 Then
        Cells(16, 5) = (Cells(14, 4) * Cells(14, 5) + Cells(15, 4) * Cells(15, 5)) / Cells(16, 4)
        End If
        
        
        Cells(14, 6) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE1, Worksheets(FDEMO).Range(LIENDEMO), "Assuré", Worksheets(FDEMO).Range(SEXEDEMO), "Masculin")
        Cells(15, 6) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE1, Worksheets(FDEMO).Range(LIENDEMO), "Assuré", Worksheets(FDEMO).Range(SEXEDEMO), "Féminin")
        Cells(16, 6) = Cells(14, 6) + Cells(15, 6)
        
        If Cells(14, 6) > 0 Then
        Cells(14, 7) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(AGEDEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE1, Worksheets(FDEMO).Range(LIENDEMO), "Assuré", Worksheets(FDEMO).Range(SEXEDEMO), "Masculin") / Cells(14, 6)
        End If
        
        If Cells(15, 6) > 0 Then
        Cells(15, 7) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(AGEDEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE1, Worksheets(FDEMO).Range(LIENDEMO), "Assuré", Worksheets(FDEMO).Range(SEXEDEMO), "Féminin") / Cells(15, 6)
        End If
        
        If Cells(16, 6) > 0 Then
        Cells(16, 7) = (Cells(14, 6) * Cells(14, 7) + Cells(15, 6) * Cells(15, 7)) / Cells(16, 6)
        End If
        
    
        If Cells(14, 4) > 0 Then
        Cells(14, 8) = Cells(14, 6) / Cells(14, 4) - 1
        End If
        
        If Cells(15, 4) > 0 Then
        Cells(15, 8) = Cells(15, 6) / Cells(15, 4) - 1
        End If
        
        If Cells(16, 4) > 0 Then
        Cells(16, 8) = Cells(16, 6) / Cells(16, 4) - 1
        End If
        
       

     
        
        Cells(37, 4) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE1, Worksheets(FDEMO).Range(LIENDEMO), "Assuré")
        Cells(38, 4) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE1, Worksheets(FDEMO).Range(LIENDEMO), "Conjoint")
        Cells(39, 4) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE1, Worksheets(FDEMO).Range(LIENDEMO), "Enfant")
        
        Cells(40, 4) = Cells(37, 4) + Cells(38, 4) + Cells(39, 4)
        
        If Cells(37, 4) > 0 Then
        Cells(37, 5) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(AGEDEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE1, Worksheets(FDEMO).Range(LIENDEMO), "Assuré") / Cells(37, 4)
        End If
        
        If Cells(38, 4) > 0 Then
        Cells(38, 5) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(AGEDEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE1, Worksheets(FDEMO).Range(LIENDEMO), "Conjoint") / Cells(38, 4)
        End If
        
        If Cells(39, 4) > 0 Then
        Cells(39, 5) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(AGEDEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE1, Worksheets(FDEMO).Range(LIENDEMO), "Enfant") / Cells(39, 4)
        End If
        

        
        
        For i = 1 To 9
        Cells(13 + i, 12) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE1, Worksheets(FDEMO).Range(LIENDEMO), "Assuré", Worksheets(FDEMO).Range(SEXEDEMO), "Masculin", Worksheets(FDEMO).Range(TRANCHEDEMO), Cells(13 + i, 11))
        Cells(13 + i, 13) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE1, Worksheets(FDEMO).Range(LIENDEMO), "Assuré", Worksheets(FDEMO).Range(SEXEDEMO), "Féminin", Worksheets(FDEMO).Range(TRANCHEDEMO), Cells(13 + i, 11))
        Cells(13 + i, 14) = Cells(13 + i, 12) + Cells(13 + i, 13)
        Next i
        
        Cells(23, 12) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE1, Worksheets(FDEMO).Range(LIENDEMO), "Assuré", Worksheets(FDEMO).Range(SEXEDEMO), "Masculin")
        Cells(23, 13) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE1, Worksheets(FDEMO).Range(LIENDEMO), "Assuré", Worksheets(FDEMO).Range(SEXEDEMO), "Féminin")
        Cells(23, 14) = Cells(23, 12) + Cells(23, 13)
   
    Calculate


End Sub
