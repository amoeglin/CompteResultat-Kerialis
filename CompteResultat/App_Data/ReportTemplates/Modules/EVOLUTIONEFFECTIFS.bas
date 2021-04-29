Attribute VB_Name = "EVOLUTIONEFFECTIFS"
Sub EVOLUTIONEFFECTIFS()


 'VARIABLES
    Dim Societe(100) As String
 'RAZ
     
    solde = 0
     
    Sheets("Evolution effectifs").Select
    
    p = 1
    While Cells(p, 3) <> "Total général"
    p = p + 1
    Wend
    p = p - 19
    
    For Z = 0 To p - 1
    Rows("15:15").Select
    Selection.Delete Shift:=xlDown
    solde = solde - 1
    Next Z
    
    For i = 3 To 11
    Cells(15, i) = ""
    Cells(16, i) = ""
    Cells(17, i) = ""
    Cells(18, i) = ""
    Cells(19, i) = ""
    Next i
    
    
    Cells(16, 3) = "assuré"
    Cells(17, 3) = "conjoint"
    Cells(18, 3) = "enfant"
    Cells(19, 3) = "Total général"

    For i = 15 To 40
    Cells(i, 27) = ""
    Cells(i, 28) = ""
    Cells(i, 29) = ""
    Next i


'REMPLISSAGE
   
     
     
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
    
    

        
    nbl = 1
    While Cells(nbl, 1) <> ""
    nbl = nbl + 1
    Wend
    
    j = 1
    Societe(j) = Cells(2, 2)
    For i = 2 To (nbl - 1)
        If Societe(j) <> Cells(i, 2) Then
        j = j + 1
        Societe(j) = Cells(i, 2)
        End If
    Next i
    
    Sheets("Evolution effectifs").Select
    
    totalactif1 = 0
    totalperif1 = 0
    totalactif2 = 0
    totalperif2 = 0
    
    For k = 1 To j
    
    If k > 1 Then
    rangeinsert = k * 4 + 11 & ":" & k * 4 + 14
    Rows(rangeinsert).Select
    Selection.Insert Shift:=xlDown
    Rows("15:18").Select
    Selection.Copy
    Rows(rangeinsert).Select
    ActiveSheet.Paste
    solde = solde + 4
    End If
    
    Cells(k * 4 + 11, 3) = ""
    Cells(k * 4 + 11, 3) = Societe(k)
    
    
    
    Cells(k * 4 + 11, 4) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(SOCIETEDEMO), Societe(k), Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE1)
    Cells(k * 4 + 12, 4) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(SOCIETEDEMO), Societe(k), Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE1, Worksheets(FDEMO).Range(LIENDEMO), "assuré")
    Cells(k * 4 + 13, 4) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(SOCIETEDEMO), Societe(k), Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE1, Worksheets(FDEMO).Range(LIENDEMO), "conjoint")
    Cells(k * 4 + 14, 4) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(SOCIETEDEMO), Societe(k), Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE1, Worksheets(FDEMO).Range(LIENDEMO), "enfant")
    
    Cells(k * 4 + 11, 5) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(SOCIETEDEMO), Societe(k), Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE2)
    Cells(k * 4 + 12, 5) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(SOCIETEDEMO), Societe(k), Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE2, Worksheets(FDEMO).Range(LIENDEMO), "assuré")
    Cells(k * 4 + 13, 5) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(SOCIETEDEMO), Societe(k), Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE2, Worksheets(FDEMO).Range(LIENDEMO), "conjoint")
    Cells(k * 4 + 14, 5) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(SOCIETEDEMO), Societe(k), Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE2, Worksheets(FDEMO).Range(LIENDEMO), "enfant")

    
    'Cells(k * 4 + 11, 8) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "PERIPHERIQUES", Worksheets(FDEMO).Range(SOCIETEDEMO), Societe(k), Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE1)
    'Cells(k * 4 + 12, 8) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "PERIPHERIQUES", Worksheets(FDEMO).Range(SOCIETEDEMO), Societe(k), Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE1, Worksheets(FDEMO).Range(LIENDEMO), "assuré")
    'Cells(k * 4 + 13, 8) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "PERIPHERIQUES", Worksheets(FDEMO).Range(SOCIETEDEMO), Societe(k), Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE1, Worksheets(FDEMO).Range(LIENDEMO), "conjoint")
    'Cells(k * 4 + 14, 8) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "PERIPHERIQUES", Worksheets(FDEMO).Range(SOCIETEDEMO), Societe(k), Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE1, Worksheets(FDEMO).Range(LIENDEMO), "enfant")
   
    'Cells(k * 4 + 11, 9) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "PERIPHERIQUES", Worksheets(FDEMO).Range(SOCIETEDEMO), Societe(k), Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE2)
    'Cells(k * 4 + 12, 9) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "PERIPHERIQUES", Worksheets(FDEMO).Range(SOCIETEDEMO), Societe(k), Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE2, Worksheets(FDEMO).Range(LIENDEMO), "assuré")
    'Cells(k * 4 + 13, 9) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "PERIPHERIQUES", Worksheets(FDEMO).Range(SOCIETEDEMO), Societe(k), Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE2, Worksheets(FDEMO).Range(LIENDEMO), "conjoint")
    'Cells(k * 4 + 14, 9) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "PERIPHERIQUES", Worksheets(FDEMO).Range(SOCIETEDEMO), Societe(k), Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE2, Worksheets(FDEMO).Range(LIENDEMO), "enfant")
      
    
    
    totalactif1 = totalactif1 + Cells(k * 4 + 11, 4)
    'totalperif1 = totalperif1 + Cells(k * 4 + 11, 8)
    totalactif2 = totalactif2 + Cells(k * 4 + 11, 5)
    'totalperif2 = totalperif2 + Cells(k * 4 + 11, 9)
    
    Next k
    
    
    
    Cells(11 + j * 4 + 4, 4) = totalactif1
    'Cells(11 + j * 4 + 4, 8) = totalperif1
    Cells(11 + j * 4 + 4, 5) = totalactif2
    'Cells(11 + j * 4 + 4, 9) = totalperif2
    
    
    

    If totalactif2 > 0 Then
    
    Cells(11 + j * 4 + 4, 7) = totalactif2 / totalactif2
    
    For k = 1 To j
    Cells(k * 4 + 11, 7) = Cells(k * 4 + 11, 5) / totalactif2
    Cells(k * 4 + 12, 7) = Cells(k * 4 + 12, 5) / totalactif2
    Cells(k * 4 + 13, 7) = Cells(k * 4 + 13, 5) / totalactif2
    Cells(k * 4 + 14, 7) = Cells(k * 4 + 14, 5) / totalactif2
    Next k
    
    End If
    
    'If totalperif2 > 0 Then
    
    'Cells(11 + j * 4 + 4, 11) = totalperif2 / totalperif2
    
    'For k = 1 To j
    'Cells(k * 4 + 11, 11) = Cells(k * 4 + 11, 9) / totalperif2
    'Cells(k * 4 + 12, 11) = Cells(k * 4 + 12, 9) / totalperif2
    'Cells(k * 4 + 13, 11) = Cells(k * 4 + 13, 9) / totalperif2
    'Cells(k * 4 + 14, 11) = Cells(k * 4 + 14, 9) / totalperif2
    'Next k
    
    'End If
    
    
    If totalactif1 > 0 Then
    
    Cells(11 + j * 4 + 4, 6) = totalactif2 / totalactif1 - 1
    
    For k = 1 To j
    
    For i = 11 To 14
    If Cells(k * 4 + i, 4) > 0 Then
    Cells(k * 4 + i, 6) = Cells(k * 4 + i, 5) / Cells(k * 4 + i, 4) - 1
    End If
    Next i
    
    Next k
    
    End If
    
    'If totalperif1 > 0 Then
    
    'Cells(11 + j * 4 + 4, 10) = totalperif2 / totalperif1 - 1
    
    'For k = 1 To j
    'For i = 11 To 14
    'If Cells(k * 4 + i, 8) > 0 Then
    'Cells(k * 4 + i, 10) = Cells(k * 4 + i, 9) / Cells(k * 4 + i, 8) - 1
    'End If
    'Next i
    'Next k
    
    'End If
    



    
    
    For i = 1 To j
    Cells(14 + i, 27) = Societe(i)
    Cells(14 + i, 28) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(SOCIETEDEMO), Societe(i), Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE1)
    Cells(14 + i, 29) = Application.WorksheetFunction.SumIfs(Worksheets(FDEMO).Range(DEMO), Worksheets(FDEMO).Range(COLDEMO), "ACTIFS", Worksheets(FDEMO).Range(SOCIETEDEMO), Societe(i), Worksheets(FDEMO).Range(ANNEEDEMO), ANNEE2)
    Next i
    
    
    zone1 = "$AA$14:$AC$" & 14 + j
   
    
    
    'ActiveSheet.ChartObjects("EVOLUTION").Activate
    'ActiveChart.PlotArea.Select
    'ActiveChart.SetSourceData Source:=Range("'Evolution effectifs'!" & zone1 _
        )
    

    If j > 1 Then

    ActiveSheet.ChartObjects("EVOLUTION").Activate
    ActiveChart.SeriesCollection(2).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent6
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.400000006
        .Transparency = 0
        .Solid
    End With
    ActiveChart.SeriesCollection(1).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.25
        .Transparency = 0
        .Solid
    End With
    
    
    
    End If
    
    
        If solde > 0 Then
    
        For i = 1 To solde
        rangeinsert = 11 + j * 4 + 5 & ":" & 11 + j * 4 + 5
        Rows(rangeinsert).Select
        Selection.Delete Shift:=xlDown
        Next i
    
    Else
        
        For i = 1 To -solde
        rangeinsert = 11 + j * 4 + 5 & ":" & 11 + j * 4 + 5
        Rows(rangeinsert).Select
        Selection.Insert Shift:=xlDown
        Next i
    
    End If

End Sub



