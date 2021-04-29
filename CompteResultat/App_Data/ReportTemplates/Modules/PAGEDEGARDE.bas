Attribute VB_Name = "PAGEDEGARDE"
Sub Pagedegarde()

'VARIABLE
Dim contrat(1000) As String
Dim college(100) As String
Dim client(100) As String


'RAZ
 
    Sheets("Page de garde").Select
        
        Cells(13, 3) = ""
        Cells(19, 7) = ""
        Cells(20, 7) = ""
        
        Cells(22, 8) = ""
        Cells(23, 10) = ""

 
 
'REMPLISSAGE

    Sheets("DATA PREST").Select
    
    ActiveWorkbook.Worksheets("DATA PREST").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("DATA PREST").Sort.SortFields.Add Key:=Range( _
        "B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("DATA PREST").Sort
        .SetRange Range("A:Z")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
        i = 3
        nbcontrat = 1
        contrat(nbcontrat) = Cells(2, 2)
     nbcontrat = 2
        While Cells(i, 1) <> ""
        
            If Cells(i, 2) <> contrat(nbcontrat - 1) Then
            contrat(nbcontrat) = Cells(i, 2)
            nbcontrat = nbcontrat + 1
            End If
   
        i = i + 1
        Wend
        
    nbcontrat = nbcontrat - 1
        
    ActiveWorkbook.Worksheets("DATA PREST").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("DATA PREST").Sort.SortFields.Add Key:=Range( _
        "C:C"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("DATA PREST").Sort
        .SetRange Range("A:Z")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
        i = 3
        nbcollege = 1
        college(nbcollege) = Cells(2, 3)
        nbcollege = 2
        While Cells(i, 1) <> ""
            
            If Cells(i, 3) <> college(nbcollege - 1) Then
            college(nbcollege) = Cells(i, 3)
            nbcollege = nbcollege + 1
            End If
   
        i = i + 1
        Wend
        
    nbcollege = nbcollege - 1
        
    Sheets("AFFICHAGE").Select
        
    Date_arrete = Cells(2, 15)
    Date_fin = Cells(2, 14)
    Date_debut = Cells(2, 13)
    
    Sheets("DATA DEMO").Select
    
        
    ActiveWorkbook.Worksheets("DATA DEMO").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("DATA DEMO").Sort.SortFields.Add Key:=Range( _
        "B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("DATA DEMO").Sort
        .SetRange Range("A:Z")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
        i = 3
        nbclient = 1
        client(nbclient) = Cells(2, 2)
        nbclient = 2
        While Cells(i, 1) <> ""
            
            If Cells(i, 2) <> client(nbclient - 1) Then
            client(nbclient) = Cells(i, 2)
            nbclient = nbclient + 1
            End If
   
        i = i + 1
        Wend
        
    nbclient = nbclient - 1
         
         
    Call DATA.DATA
    
    
    Sheets("Page de garde").Select
    
        For i = 1 To nbclient
            If i = 1 Then
            Cells(13, 3) = client(i)
            Else
            Cells(13, 3) = Cells(13, 3) & " - " & client(i)
            End If
        Next i
        
        For i = 1 To nbcontrat
            If i = 1 Then
            Cells(19, 7) = contrat(i)
            Else
            Cells(19, 7) = Cells(19, 7) & ", " & contrat(i)
            End If
        Next i
        
        For i = 1 To nbcollege
            If i = 1 Then
            Cells(20, 7) = college(i)
            Else
            Cells(20, 7) = Cells(20, 7) & ", " & college(i)
            End If
        Next i
        
        Cells(22, 8) = Format(Date_debut, "d mmmm yyyy") & " au " & Format(Date_fin, "d mmmm yyyy")
        Cells(23, 10) = Format(Date_arrete, "d mmmm yyyy")


End Sub

