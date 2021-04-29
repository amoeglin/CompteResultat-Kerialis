Attribute VB_Name = "PRESTATIONSREGLEES"
Sub PRESTATIONSREGLEES()

   
    Dim Famille(100) As String
    Dim Acte(100) As String
    Dim Famille_Bis(100) As String
    
    
'RAZ

    Sheets("Prestations Réglées").Select
    
    p = 1
    While Cells(p, 3) <> "Total général"
    p = p + 1
    Wend
    p = p - 15
    
    solde = 0
    
    For Z = 1 To p - 1
    Rows("16:16").Select
    Selection.Delete Shift:=xlDown
    solde = solde - 1
    Next Z


    For i = 3 To 18
    Cells(15, i) = ""
    Cells(16, i) = ""
    Next i

    Cells(16, 3) = "Total général"
    
    
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
    
    If Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("H:H"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("F:F"), Cells(i, 2)) <> 0 Then
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
    If Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("H:H"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("F:F"), Cells(i, 2)) <> 0 Then

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
    
    nbacte = s - 1

    
' REMPLISSAGE

    Sheets("Prestations Réglées").Select
    
    Cells(15, 3) = Famille(1)
    For k = 2 To (nbfamille)
    rangeinsert = k + 14 & ":" & k + 14
    Rows(rangeinsert).Select
    Selection.Insert Shift:=xlDown
    Cells(14 + k, 3) = Famille(k)
    solde = solde + 1
    Next k

    t = 0
    For k = 1 To (nbacte)
    rangeinsert = t + k + 15 & ":" & t + k + 15
    Rows(rangeinsert).Select
    Selection.Insert Shift:=xlDown
    Cells(15 + k + t, 4) = Acte(k)
    solde = solde + 1
    If Famille_Bis(k) <> Famille_Bis(k + 1) Then
    t = t + 1
    End If
    
       
    If Famille_Bis(k) = Acte(k) Then
    Rows(rangeinsert).Select
    Selection.Delete Shift:=xlDown
    t = t - 1
    solde = solde - 1
    End If
    Next k
    
    
    f = 15
    While Cells(f, 3) <> "Total général"
    
    If Cells(f, 4) <> "" Then
    Range("C" & f & ":P" & f).Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    End If
    
    f = f + 1
    Wend
  
        
    
    
    Total_FR = 0
    Total_SS = 0
    Total_AUTRES = 0
    Total_NOUS = 0
    Total_NB = 0
    
    p = 1
    While Cells(p, 3) <> "Total général"
    p = p + 1
    Wend
    p = p - 15
    
        FDEMO = "DATA DEMO"
        ANNEEDEMO = "A:A"
        SOCIETEDEMO = "B:B"
        SEXEDEMO = "D:D"
        AGEDEMO = "H:H"
        DEMO = "G:G"
        LIENDEMO = "E:E"
        TRANCHEDEMO = "F:F"
        COLDEMO = "J:J"
   
    Exposition = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE_2)
    
    
    For k = 1 To p
    If Cells(14 + k, 3) = "" Then
    If Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("H:H"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("E:E"), Cells(14 + k, 4)) <> 0 Then
    Cells(14 + k, 5) = Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("H:H"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("E:E"), Cells(14 + k, 4))
    If Exposition > 0 Then
    Cells(14 + k, 6) = Cells(14 + k, 5) / Exposition
    End If
    Cells(14 + k, 7) = Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("I:I"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("E:E"), Cells(14 + k, 4))
    Cells(14 + k, 8) = Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("J:J"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("E:E"), Cells(14 + k, 4))
    Cells(14 + k, 9) = Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("K:K"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("E:E"), Cells(14 + k, 4))
    Cells(14 + k, 10) = Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("L:L"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("E:E"), Cells(14 + k, 4))
    Cells(14 + k, 11) = Cells(14 + k, 10) / Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("H:H"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("E:E"), Cells(14 + k, 4))
    
    If Cells(14 + k, 3) <> "MATERNITE" Then
    Cells(14 + k, 13) = (Cells(14 + k, 8) + Cells(14 + k, 10) + Cells(14 + k, 9)) / Cells(14 + k, 7)
    End If
    
    End If
    Else
    If Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("H:H"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("F:F"), Cells(14 + k, 3)) <> 0 Then
    
    Cells(14 + k, 5) = Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("H:H"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("F:F"), Cells(14 + k, 3))
    Total_NB = Total_NB + Cells(14 + k, 5)
    If Exposition > 0 Then
    Cells(14 + k, 6) = Cells(14 + k, 5) / Exposition
    End If
    Cells(14 + k, 7) = Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("I:I"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("F:F"), Cells(14 + k, 3))
    Total_FR = Total_FR + Cells(14 + k, 7)
    Cells(14 + k, 8) = Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("J:J"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("F:F"), Cells(14 + k, 3))
    Total_SS = Total_SS + Cells(14 + k, 8)
    Cells(14 + k, 9) = Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("K:K"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("F:F"), Cells(14 + k, 3))
    Total_AUTRES = Total_AUTRES + Cells(14 + k, 9)
    Cells(14 + k, 10) = Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("L:L"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("F:F"), Cells(14 + k, 3))
    Total_NOUS = Total_NOUS + Cells(14 + k, 10)
    Cells(14 + k, 11) = Cells(14 + k, 10) / Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("H:H"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("F:F"), Cells(14 + k, 3))
    If Cells(14 + k, 7) <> 0 Then
    If Cells(14 + k, 3) <> "MATERNITE" Then
    Cells(14 + k, 13) = (Cells(14 + k, 10) + Cells(14 + k, 9) + Cells(14 + k, 8)) / Cells(14 + k, 7)
    End If
    End If
    End If
    End If
    Next k

    Cells(14 + p + 1, 5) = Total_NB
    
    If Exposition > 0 Then
    Cells(14 + p + 1, 6) = Total_NB / Exposition
    End If
    
    Cells(14 + p + 1, 7) = Total_FR
    Cells(14 + p + 1, 8) = Total_SS
    Cells(14 + p + 1, 9) = Total_AUTRES
    Cells(14 + p + 1, 10) = Total_NOUS
    If Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("H:H"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2) <> 0 Then
    Cells(14 + p + 1, 11) = Total_NOUS / Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("H:H"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2)
    End If
    
    If Total_FR <> 0 Then
    Cells(14 + p + 1, 13) = (Total_NOUS + Total_SS + Total_AUTRES) / Total_FR
    End If
    
    Total_FR = 0
    Total_SS = 0
    Total_AUTRES = 0
    Total_NOUS = 0
    
    
    
    For k = 1 To p
    If Cells(14 + k, 3) = "" Then
    If Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("H:H"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("E:E"), Cells(14 + k, 4)) <> 0 Then
    Cells(14 + k, 12) = Cells(14 + k, 10) / Cells(14 + p + 1, 10)
    End If
    Else
    If Application.WorksheetFunction.SumIfs(Worksheets("DATA PREST").Range("H:H"), Worksheets("DATA PREST").Range("D:D"), ANNEE_2, Worksheets("DATA PREST").Range("F:F"), Cells(14 + k, 3)) <> 0 Then
    Cells(14 + k, 12) = Cells(14 + k, 10) / Cells(14 + p + 1, 10)
    End If
    End If
    Next k
    
    
    For k = 1 To p
    If Cells(14 + k, 3) = "" Then
    If Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("H:H"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2, Worksheets("DATA EXP").Range("E:E"), Cells(14 + k, 4)) <> 0 Then
    T_5 = Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("I:I"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2, Worksheets("DATA EXP").Range("E:E"), Cells(14 + k, 4))
    T_6 = Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("J:J"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2, Worksheets("DATA EXP").Range("E:E"), Cells(14 + k, 4))
    T_7 = Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("K:K"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2, Worksheets("DATA EXP").Range("E:E"), Cells(14 + k, 4))
    T_8 = Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("L:L"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2, Worksheets("DATA EXP").Range("E:E"), Cells(14 + k, 4))
    Cells(14 + k, 15) = T_8 / Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("H:H"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2, Worksheets("DATA EXP").Range("E:E"), Cells(14 + k, 4))
    Cells(14 + k, 16) = T_8 / T_5
    
    If Cells(14 + k, 3) <> "MATERNITE" Then
    Cells(14 + k, 17) = (T_6 + T_7 + T_8) / T_5
    End If
    
    If Cells(14 + p + 1, 10) <> 0 Then
    Cells(14 + k, 12) = Cells(14 + k, 10) / Cells(14 + p + 1, 10)
    End If
    End If
    Else
    If Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("H:H"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2, Worksheets("DATA EXP").Range("F:F"), Cells(14 + k, 3)) <> 0 Then
    If Cells(14 + p + 1, 10) <> 0 Then
    Cells(14 + k, 12) = Cells(14 + k, 10) / Cells(14 + p + 1, 10)
    End If
    Total_R = Total_R + Cells(14 + k, 12)
    
    Cells(14 + p + 1, 12) = Total_R

    
    
    T_5 = Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("I:I"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2, Worksheets("DATA EXP").Range("F:F"), Cells(14 + k, 3))
    Total_FR = Total_FR + T_5
    T_6 = Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("J:J"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2, Worksheets("DATA EXP").Range("F:F"), Cells(14 + k, 3))
    Total_SS = Total_SS + T_6
    T_7 = Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("K:K"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2, Worksheets("DATA EXP").Range("F:F"), Cells(14 + k, 3))
    Total_AUTRES = Total_AUTRES + T_7
    T_8 = Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("L:L"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2, Worksheets("DATA EXP").Range("F:F"), Cells(14 + k, 3))
    Total_NOUS = Total_NOUS + T_8
    If Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("H:H"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2, Worksheets("DATA EXP").Range("F:F"), Cells(14 + k, 3)) <> 0 Then
    Cells(14 + k, 15) = T_8 / Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("H:H"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2, Worksheets("DATA EXP").Range("F:F"), Cells(14 + k, 3))
    End If
    If T_5 <> 0 Then
    Cells(14 + k, 16) = T_8 / T_5
    
    If Cells(14 + k, 3) <> "MATERNITE" Then
    Cells(14 + k, 17) = (T_6 + T_7 + T_8) / T_5
    End If
    
    End If
    End If
    End If
    Next k
    
    If Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("H:H"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2) <> 0 Then
    Cells(14 + p + 1, 15) = Total_NOUS / Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("H:H"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2)
    End If
    If Total_FR <> 0 Then
    Cells(14 + p + 1, 17) = (Total_NOUS + Total_SS + Total_AUTRES) / Total_FR
    End If
    
    Cells(14 + p + 1, 12) = Total_R


    For k = 1 To p
    If Total_NOUS <> 0 Then
    If Cells(14 + k, 3) = "" Then
    If Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("H:H"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2, Worksheets("DATA EXP").Range("E:E"), Cells(14 + k, 4)) <> 0 Then
    Cells(14 + k, 16) = Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("L:L"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2, Worksheets("DATA EXP").Range("E:E"), Cells(14 + k, 4)) / Total_NOUS
    End If
    Else
    If Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("H:H"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2, Worksheets("DATA EXP").Range("F:F"), Cells(14 + k, 3)) <> 0 Then
    Cells(14 + k, 16) = Application.WorksheetFunction.SumIfs(Worksheets("DATA EXP").Range("L:L"), Worksheets("DATA EXP").Range("D:D"), ANNEE_2, Worksheets("DATA EXP").Range("F:F"), Cells(14 + k, 3)) / Total_NOUS
    End If
    End If
    End If
    Next k

    Cells(14 + p + 1, 16) = Total_R

    If Exposition > 0 Then
    Cells(14 + p + 3, 7) = Cells(14 + p + 1, 7) / Exposition
    Cells(14 + p + 3, 10) = Cells(14 + p + 1, 10) / Exposition
    End If
    
    
    TotalRemboursement = Cells(14 + p + 1, 8) + Cells(14 + p + 1, 9) + Cells(14 + p + 1, 10)
    
    
    If TotalRemboursement > 0 Then
    Cells(14 + p + 5, 8) = Cells(14 + p + 1, 8) / TotalRemboursement
    Cells(14 + p + 5, 9) = Cells(14 + p + 1, 9) / TotalRemboursement
    Cells(14 + p + 5, 10) = Cells(14 + p + 1, 10) / TotalRemboursement
    End If


    t = 0
    For k = 1 To p
    If Cells(14 + k - t, 5) = "" Then
    Rows(14 + k - t & ":" & 14 + k - t).Select
    Selection.Delete Shift:=xlDown
    solde = solde - 1
    t = t + 1
    End If
    Next k
    

     If solde > 0 Then
    
        For i = 1 To solde
        rangeinsert = p + 21 & ":" & p + 21
        Rows(rangeinsert).Select
        Selection.Delete Shift:=xlDown
        Next i
    
    Else
        
        For i = 1 To -solde
        rangeinsert = p + 21 & ":" & p + 21
        Rows(rangeinsert).Select
        Selection.Insert Shift:=xlDown
        Next i
    
    End If
    
    
    
End Sub


