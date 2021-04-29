Attribute VB_Name = "EXPORT"
Sub exportFile()
Dim myWB As Workbook
Dim shtParams As Worksheet
Dim APPppt As Variant

Dim cheminDestXls As String
Dim cheminDestPPT As String
Dim cheminSourcePPT As String
Dim oPPTApp As PowerPoint.Application
Dim oPPTPres As PowerPoint.Presentation
Dim sPresentationFile As String
Dim param1 As String
Dim param2 As String
Dim companyLogo As String
Dim assureurLogo As String
Dim Rslt As Variant


Nomxls = ThisWorkbook.Name
Nomppt = Replace(Nomxls, "xlsm", "pptm")

'récupération des chemins complets des fichiers d'export
cheminDestXls = Application.ThisWorkbook.Path & "\" & Nomxls
cheminDestPPT = Application.ThisWorkbook.Path & "\" & Nomppt
'récupération des chemins complets du fichier modèle ppt
cheminSourcePPT = Application.ThisWorkbook.Path


sPresentationFile = cheminSourcePPT

Set oPPTApp = CreateObject("PowerPoint.application")
oPPTApp.Visible = True

Set oPPTPres = oPPTApp.Presentations.Open(cheminDestPPT)



param1 = cheminDestPPT & "!Module1.M2"


param2 = cheminDestXls

Rslt = oPPTApp.Run(param1)

Set oPPTPres = Nothing
Set oPPTApp = Nothing

Exit Sub

End Sub
