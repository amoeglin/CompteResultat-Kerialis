Attribute VB_Name = "OUVERTURE"
Private Sub Auto_Open()

If Sheets("Page de garde").Cells(2, 30) = "OUI" Then
Sheets("Page de garde").Cells(2, 30) = "NON"
FULL.FULL
EXPORT.exportFile


'While Sheets("Page de garde").Cells(2, 31) = "NON"
'waitTime = TimeSerial(0, 0, 30)
'Application.Wait waitTime
'Wend

ThisWorkbook.Save
'Application.Quit

End If

End Sub
