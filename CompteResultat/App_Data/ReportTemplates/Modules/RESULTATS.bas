Attribute VB_Name = "RESULTATS"
Sub RESULTATS()

'VARIABLE

Dim college(100) As String
Dim LigneAnnee2 As Double
Dim LigneAnnee1 As Double

'RAZ

LigneAnnee2 = 20  ' numéro de ligne pour l'Année 1
LigneAnnee1 = 25  ' numéro de ligne pour l'Année 2

    Sheets("DATA DEMO").Select

    FDEMO = "DATA DEMO"
    ANNEEDEMO = "A:A"
    SOCIETEDEMO = "B:B"
    LIENDEMO = "E:E"
    DEMO = "G:G"
    COLDEMO = "J:J"
    
    
        Range("J:J").ClearContents
        
        Cells(1, 10) = "FAMILLE COLLEGE"
        
        
        FCOL = "COLLEGE"
        FAMCOL = "B:C"
        
        j = 2
        While Cells(j, 1) <> ""
        Cells(j, 10) = Application.VLookup(Cells(j, 9), Sheets(FCOL).Range(FAMCOL), 2, False)
        j = j + 1
        Wend
 
    Sheets("Résultats").Select
        
' ANNEE2
        Cells(LigneAnnee2, 3) = ""
        Cells(LigneAnnee2, 3) = ""
        Cells(LigneAnnee2, 7) = ""
        Cells(LigneAnnee2, 9) = ""
        Cells(LigneAnnee2, 11) = ""
        Cells(LigneAnnee2, 13) = ""
        Cells(LigneAnnee2, 15) = ""
        Cells(LigneAnnee2, 17) = ""

' ANNEE1
        Cells(LigneAnnee1, 3) = ""
        Cells(LigneAnnee1, 3) = ""
        Cells(LigneAnnee1, 7) = ""
        Cells(LigneAnnee1, 9) = ""
        Cells(LigneAnnee1, 11) = ""
        Cells(LigneAnnee1, 13) = ""
        Cells(LigneAnnee1, 15) = ""
        Cells(LigneAnnee1, 17) = ""

'DATA
    
    Sheets("AFFICHAGE").Select
    Taxe = Cells(2, 17)
    TaxeActif = Cells(2, 18)
    TaxePerif = Cells(2, 19)
    
    
    Sheets("DATA PREST").Select
        
        i = 2
        If Cells(2, 4) = "" Then
        Exit Sub
        End If
        ANNEE1 = Cells(2, 4)
        While Cells(i, 4) = ANNEE1
        ANNEE2 = Cells(i + 1, 4)
        i = i + 1
        Wend
        
        If ANNEE2 = "" Then
        ANNEE2 = ANNEE1
        ANNEE1 = ""
        End If
        
        
        Range("R:R").ClearContents
        
        Cells(1, 18) = "FAMILLE COLLEGE"
        
        FCOL = "COLLEGE"
        FAMCOL = "B:C"
        
        j = 2
        While Cells(j, 1) <> ""
        Cells(j, 18) = Application.VLookup(Cells(j, 3), Sheets(FCOL).Range(FAMCOL), 2, False)
        j = j + 1
        Wend
    
        FPREST = "DATA PREST"
        PRESTATION = "L:L"
        PRESTATIONANNEE = "D:D"
        PRESTATIONCOL = "R:R"
        
        Sheets("DATA PROV").Select
    
        Range("H:H").ClearContents
        
        Cells(1, 8) = "FAMILLE COLLEGE"
        
        
        FCOL = "COLLEGE"
        FAMCOL = "B:C"
        
        j = 2
        While Cells(j, 1) <> ""
        Cells(j, 8) = Application.VLookup(Cells(j, 3), Sheets(FCOL).Range(FAMCOL), 2, False)
        j = j + 1
        Wend
    
        
        FPROV = "DATA PROV"
        PROVISION = "G:G"
        PROVISIONANNEE = "D:D"
        PROVISIONCOL = "H:H"
        
    Sheets("DATA COT").Select
    
    
        Range("G:G").ClearContents
        
        Cells(1, 7) = "FAMILLE COLLEGE"
        
        
        FCOL = "COLLEGE"
        FAMCOL = "B:C"
        
        j = 2
        While Cells(j, 1) <> ""
        Cells(j, 7) = Application.VLookup(Cells(j, 4), Sheets(FCOL).Range(FAMCOL), 2, False)
        j = j + 1
        Wend
    
        FCOT = "DATA COT"
        COTISATION_NETTE = "F:F"
        COTISATION_BRUTE = "H:H"
        COTISATIONANNEE = "E:E"
        COTISATIONCOL = "G:G"
    
    
 'REMPLISSAGE
    
    
    Sheets("Résultats").Select
    
 '***
 '*** ANNEE 2 ***
 '***
        Cells(LigneAnnee2, 3) = ANNEE2

        ' cotisations Brutes
        Cells(LigneAnnee2, 11) = Application.WorksheetFunction.SumIfs(Sheets(FCOT).Range(COTISATION_BRUTE), Sheets(FCOT).Range(COTISATIONANNEE), ANNEE2, Sheets(FCOT).Range(COTISATIONCOL), "ACTIFS")
        
        ' cotisations nettes
        Cells(LigneAnnee2, 15) = Application.WorksheetFunction.SumIfs(Sheets(FCOT).Range(COTISATION_NETTE), Sheets(FCOT).Range(COTISATIONANNEE), ANNEE2, Sheets(FCOT).Range(COTISATIONCOL), "ACTIFS")
        
        ' chargements
        If Cells(LigneAnnee2, 11) <> 0 Then
        Cells(LigneAnnee2, 13) = Round(1 - Cells(LigneAnnee2, 15) / Cells(LigneAnnee2, 11), 4)
        Else
        Cells(LigneAnnee2, 13) = 0
        End If
        
        ' prestations
        Cells(LigneAnnee2, 7) = Application.WorksheetFunction.SumIfs(Sheets(FPREST).Range(PRESTATION), Sheets(FPREST).Range(PRESTATIONANNEE), ANNEE2, Sheets(FPREST).Range(PRESTATIONCOL), "ACTIFS")
        
        'provisions
        Cells(LigneAnnee2, 9) = Application.WorksheetFunction.SumIfs(Sheets(FPROV).Range(PROVISION), Sheets(FPROV).Range(PROVISIONANNEE), ANNEE2, Sheets(FPROV).Range(PROVISIONCOL), "ACTIFS")
        
        ' ratio
        If Cells(LigneAnnee2, 15) > 0 Then
        Cells(LigneAnnee2, 17) = (Cells(LigneAnnee2, 7) + Cells(LigneAnnee2, 9)) / Cells(LigneAnnee2, 15)
        End If
        

 '***
 '*** ANNEE 1 ***
 '***
        Cells(LigneAnnee1, 3) = ANNEE1
               
               
        ' cotisations Brutes
        Cells(LigneAnnee1, 11) = Application.WorksheetFunction.SumIfs(Sheets(FCOT).Range(COTISATION_BRUTE), Sheets(FCOT).Range(COTISATIONANNEE), ANNEE1, Sheets(FCOT).Range(COTISATIONCOL), "ACTIFS")
        
        ' cotisations nettes
        Cells(LigneAnnee1, 15) = Application.WorksheetFunction.SumIfs(Sheets(FCOT).Range(COTISATION_NETTE), Sheets(FCOT).Range(COTISATIONANNEE), ANNEE1, Sheets(FCOT).Range(COTISATIONCOL), "ACTIFS")
        
        ' chargements
        If Cells(LigneAnnee1, 11) <> 0 Then
        Cells(LigneAnnee1, 13) = Round(1 - Cells(LigneAnnee1, 15) / Cells(LigneAnnee1, 11), 4)
        Else
        Cells(LigneAnnee1, 13) = 0
        End If
        
        ' prestations
        Cells(LigneAnnee1, 7) = Application.WorksheetFunction.SumIfs(Sheets(FPREST).Range(PRESTATION), Sheets(FPREST).Range(PRESTATIONANNEE), ANNEE1, Sheets(FPREST).Range(PRESTATIONCOL), "ACTIFS")
        
        'provisions
        Cells(LigneAnnee1, 9) = Application.WorksheetFunction.SumIfs(Sheets(FPROV).Range(PROVISION), Sheets(FPROV).Range(PROVISIONANNEE), ANNEE1, Sheets(FPROV).Range(PROVISIONCOL), "ACTIFS")
        
        ' ratio
        If Cells(LigneAnnee1, 15) > 0 Then
        Cells(LigneAnnee1, 17) = (Cells(LigneAnnee1, 7) + Cells(LigneAnnee1, 9)) / Cells(LigneAnnee1, 15)
        End If
               
               
               
End Sub
