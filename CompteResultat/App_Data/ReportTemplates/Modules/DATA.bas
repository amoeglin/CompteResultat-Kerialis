Attribute VB_Name = "DATA"
Sub DATA()
 
    Sheets("DATA PREST").Select

        ActiveWorkbook.Worksheets("DATA PREST").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("DATA PREST").Sort.SortFields.Add Key:=Range( _
        "D:D"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("DATA PREST").Sort.SortFields.Add Key:=Range( _
        "F:F"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("DATA PREST").Sort.SortFields.Add Key:=Range( _
        "E:E"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("DATA PREST").Sort
        .SetRange Range("A:Z")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
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
    

    Sheets("DATA COT").Select
    Columns("A:B").Select
    Selection.NumberFormat = "m/d/yyyy"
    Sheets("DATA PREST").Select
    Columns("A:A").Select
    Selection.NumberFormat = "m/d/yyyy"
    Sheets("DATA EXP").Select
    Columns("A:A").Select
    Selection.NumberFormat = "m/d/yyyy"
    Sheets("DATA PROV").Select
    Columns("A:A").Select
    Selection.NumberFormat = "m/d/yyyy"
    Sheets("DATA DEMO").Select
    Sheets("Page de garde").Select


    End Sub
   


