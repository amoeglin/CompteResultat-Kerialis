Attribute VB_Name = "PRESTATIONSREGLEESGRAPH"
Sub PRESTATIONSREGLEESGRAPH()

    Dim Famille(100) As String
    Dim Acte(100) As String
    Dim Famille_Bis(100) As String
    

' RAZ


    Sheets("Prestations Réglées Graph").Select
    
    p = 1
    
    While Cells(66 + p, 5) <> ""
    p = p + 1
    Wend
    
    For k = 1 To p
    Cells(66 + k, 5) = ""
    Cells(66 + k, 6) = ""
    Cells(66 + k, 7) = ""
    Cells(66 + k, 8) = ""
    Cells(66 + k, 9) = ""
    Cells(66 + k, 10) = ""
    Cells(66 + k, 11) = ""
    Cells(66 + k, 12) = ""
    
    Next k

' DATA

    Sheets("DATA PREST").Select
        
        i = 2
        If Cells(2, 4) = "" Then
        Exit Sub
        End If
        ANNEE_1 = Cells(2, 4)
        While Cells(i, 4) = ANNEE_1
        ANNEE_2 = Cells(i + 1, 4)
        i = i + 1
        Wend
        
        If ANNEE_2 = "" Then
        ANNEE_2 = ANNEE_1
        ANNEE_1 = ""
        End If
        
    Sheets("AFFICHAGE").Select
    i = 2
    s = 1
    While Cells(i, 1) <> ""
    
    If Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("H:H"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("F:F"), Cells(i, 2)) > 0 Then
            If s = 1 Then
            Famille(s) = Cells(i, 2)
            s = s + 1
            ElseIf Famille(s - 1) <> Cells(i, 2) Then
            Famille(s) = Cells(i, 2)
            s = s + 1
            End If
    End If
    
    i = i + 1
    Wend
    
    nbfamille = s - 1
    
    i = 2
    s = 1
    While Cells(i, 1) <> ""
    If Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("H:H"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("F:F"), Cells(i, 2)) > 0 Then
            If s = 1 Then
            Acte(s) = Cells(i, 3)
            Famille_Bis(s) = Cells(i, 2)
            s = s + 1
            ElseIf Acte(s - 1) <> Cells(i, 3) Then
            Acte(s) = Cells(i, 3)
            Famille_Bis(s) = Cells(i, 2)
            s = s + 1
            End If
    End If
    i = i + 1
    Wend
    
    NbActe = s - 1






' Prestations Réglées Graph
 
 
    Sheets("Prestations Réglées Graph").Select
    
    For k = 1 To nbfamille
    Cells(66 + k, 5) = Famille(k)
    Cells(66 + k, 6) = Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("L:L"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("F:F"), Cells(66 + k, 5))
    Cells(66 + k, 7) = Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("L:L"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2, Worksheets("DATA EXP").Range("F:F"), Cells(66 + k, 5))
    Cells(66 + k, 8) = Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("L:L"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("F:F"), Cells(66 + k, 5))
    Cells(66 + k, 9) = Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("J:J"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("F:F"), Cells(66 + k, 5))
    Cells(66 + k, 10) = Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("K:K"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("F:F"), Cells(66 + k, 5))
    Cells(66 + k, 11) = Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("L:L"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("F:F"), Cells(66 + k, 5))
    Cells(66 + k, 12) = Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("I:I"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("F:F"), Cells(66 + k, 5)) - Cells(66 + k, 10) - Cells(66 + k, 9) - Cells(66 + k, 8)
      
    Cells(66 + k, 13) = Cells(66 + k, 6) / Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("H:H"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("F:F"), Cells(66 + k, 5))
     
    If Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("H:H"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2, Worksheets("DATA EXP").Range("F:F"), Cells(66 + k, 5)) > 0 Then
    Cells(66 + k, 14) = Cells(66 + k, 7) / Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("H:H"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2, Worksheets("DATA EXP").Range("F:F"), Cells(66 + k, 5))
    End If
      
    
    If Cells(66 + k, 5) <> "MATERNITE" Then
   
    If Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("I:I"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("F:F"), Cells(66 + k, 5)) > 0 Then
    Cells(34 + nbfamille - k, 10) = (Cells(66 + k, 9) + Cells(66 + k, 10) + Cells(66 + k, 11)) / Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("I:I"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("F:F"), Cells(66 + k, 5))
    End If
    
    If Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("I:I"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2, Worksheets("DATA EXP").Range("F:F"), Cells(66 + k, 5)) > 0 Then
    Cells(34 + nbfamille - k, 11) = (Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("J:J"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2, Worksheets("DATA EXP").Range("F:F"), Cells(66 + k, 5)) + Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("K:K"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2, Worksheets("DATA EXP").Range("F:F"), Cells(66 + k, 5)) + Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("L:L"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2, Worksheets("DATA EXP").Range("F:F"), Cells(66 + k, 5))) / Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("I:I"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2, Worksheets("DATA EXP").Range("F:F"), Cells(66 + k, 5))
    End If
    
    End If
    
    Next k

    zone1 = "$E$66:$E$" & 66 + nbfamille
    zone2 = "$F$66:$F$" & 66 + nbfamille
    zone3 = "$I$66:$L$" & 66 + nbfamille
    zone4 = "$M$66:$N$" & 66 + nbfamille
    zone5 = "$G$66:$G$" & 66 + nbfamille
    
    ActiveSheet.ChartObjects("Prest1").Activate
    ActiveChart.PlotArea.Select
    ActiveChart.SetSourceData Source:=Range( _
        "'Prestations Réglées Graph'!" & zone1 & ",'Prestations Réglées Graph'!" & zone2)
        
    ActiveSheet.ChartObjects("Prest4").Activate
    ActiveChart.PlotArea.Select
    ActiveChart.SetSourceData Source:=Range( _
        "'Prestations Réglées Graph'!" & zone1 & ",'Prestations Réglées Graph'!" & zone5)
        
    ActiveSheet.ChartObjects("Prest2").Activate
    ActiveChart.PlotArea.Select
    ActiveChart.SetSourceData Source:=Range( _
        "'Prestations Réglées Graph'!" & zone1 & ",'Prestations Réglées Graph'!" & zone3)
        
    ActiveSheet.ChartObjects("Prest3").Activate
    ActiveChart.PlotArea.Select
    ActiveChart.SetSourceData Source:=Range( _
        "'Prestations Réglées Graph'!" & zone1 & ",'Prestations Réglées Graph'!" & zone4)

    Calculate

End Sub




