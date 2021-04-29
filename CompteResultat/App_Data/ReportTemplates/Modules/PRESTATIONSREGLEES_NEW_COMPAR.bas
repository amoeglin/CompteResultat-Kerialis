Attribute VB_Name = "PRESTATIONSREGLEES_NEW_COMPAR"
Sub PRESTATIONSREGLEES_NEW_COMPAR()

'Option Explicit
'
Dim bValue As Boolean
Dim Famille_A_Traiter As String
Dim NomModule As String

' déclaration des feuilles

Dim shDonnees As Worksheet
Dim shDonneesAffichage As Worksheet
Dim shDonnéesDemo As Worksheet
Dim shDonneesExperience As Worksheet
Dim shResultats As Worksheet
Dim shErreurs As Worksheet
Dim shResultats_Annee_N As Worksheet


' CHARGEMENT des feuilles

Set shDonnees = Worksheets("DATA PREST")
Set shDonneesAffichage = Worksheets("AFFICHAGE")
Set shDonnéesDemo = Worksheets("DATA DEMO")
Set shDonneesExperience = Worksheets("DATA EXP")
'Set shResultats = Worksheets("Prestations Réglées_NEW")
Set shErreurs = Worksheets("Erreurs")

'*** RAZ des lignes dans la feuille Erreurs  ***
'bValue = RAZ_shErreurs(shErreurs) ' désactivé
    
' TRAITEMENT de TOUTES LES PRESTATIONS - Onglet Prestations Réglées_NEW
Set shResultats = Worksheets("Prestations Réglées N et N-1")
NomModule = "Onglet Prestations Réglées N et N-1"
Famille_A_Traiter = "TOUTES" ' traitement de toutes les familles
bValue = PRESTATIONSREGLEES_GENERIQUE_COMPARAISON(Famille_A_Traiter, NomModule, shDonnees, shDonneesAffichage, shDonnéesDemo, shDonneesExperience, shResultats, shErreurs)
    
    
End Sub

Function PRESTATIONSREGLEES_GENERIQUE_COMPARAISON(Famille_A_Traiter As String, NomModule As String, shDonnees As Worksheet, shDonneesAffichage As Worksheet, shDonnéesDemo As Worksheet, shDonneesExperience As Worksheet, shResultats As Worksheet, shErreurs As Worksheet) As Boolean

'Option Explicit
'Version 1-0 le 29/03/2021
' Mot de passe de protection des feuilles XL
'Public Const PROTECT_PASSWORD As String = "CMPASS" 'déclaration des feuilles

Dim Famille(200) As String
Dim Acte(200) As String
Dim Famille_Bis(200) As String
Dim bValue As Boolean

Dim annee As Double
Dim annee_N_1 As Double
Dim LibelleFamille As String
Dim LibelleActe As String
Dim Tmax As Double
Tmax = 200 ' nombre maximum de valeurs dans un tableau

'shResultats.Unprotect PROTECT_PASSWORD
   
Dim LigneDebut As Double  ' = PREMIERE LIGNE PREMIERE LIGNE DU TABLEAU à compléter dans la feuille shDonnees
Dim MessageErreur As String
Dim Famille_Acte_Annnee As String
Dim NoLigneEnErreur As Double

  On Error GoTo err_Chargement_PRESTATIONS_REGLEES
 
' Module
    NoLigneEnErreur = 0
    
    ' sélection de la feuille shResultats
    shResultats.Select
    
    'PREMIERE LIGNE PREMIERE LIGNE DU TABLEAU à compléter dans la feuille shDonnees
    LigneDebut = 14
    
    '*** TRAITEMENT DES LIGNES FAMILLES ET DES ACTES DANS LA FEUILLE RESULTATS  ***
    
    ' recherche du numéro de la ligne "Total général" pour permettre de supprimer toutes les lignes avant d'insérer les nouvelles
    p = 1
    While Cells(p, 3) <> "Total général"
    '''While Cells(P, 3) <> "Total générall"
    p = p + 1
    Wend
    p = p - (LigneDebut + 1)
    
    solde = 0
    
    ' SUPPRESSSION DE TOUTES LES LIGNES dans le tableau jusqu'à la ligne avec "Total général"
    For Z = 1 To p - 1
    
    RangeDelete = k + LigneDebut + 2 & ":" & k + LigneDebut + 2 ' chargement du numéro de la ligne suivante à insérer
    Rows(RangeDelete).Select
    'Rows("16:16").Select
    Selection.Delete Shift:=xlDown
    solde = solde - 1
    Next Z

    For i = 3 To 18
    Cells(LigneDebut + 1, i) = ""
    Cells(LigneDebut + 2, i) = ""
    Next i

' CHARGEMENT DE LA LIGNE (LigneDebut + 2) avec le libellé "Total général"
    Cells(LigneDebut + 2, 3) = "Total général"
    
' DATA

    ' sélection de la feuille shDonnees
    shDonnees.Select

    ' chargement des 2années (ANNEE_1 et ANNEE_2) depuis les données Source Prestations
        i = 2
        If Cells(2, 4) = "" Then
        Exit Function
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
        
        
' sélection de la feuille shDonneesAffichage
shDonneesAffichage.Select

    ' il est IMPORTANT DE CLASSER LES FAMILLES PAR ORDRE APHABETIQUE DANS shDonneesAffichage
    
    i = 2  ' numéro de ligne = 2
    s = 1  ' nombre de familles chargées
    
If Famille_A_Traiter = "TOUTES" Then ' famille lue dans shDonneesAffichage 'traitement DE TOUTES LES FAMILLES
    
    While Cells(i, 2) <> ""  ' famille lue Cells(i, 2) Dans shDonneesAffichage
    
    LibelleFamille = ""
    LibelleActe = ""
    annee = ANNEE_2
            If s = 1 Then
            Famille(s) = Cells(i, 2)
            LibelleFamille = Cells(i, 2)
            s = s + 1
            ElseIf Famille(s - 1) <> Cells(i, 2) Then
            Famille(s) = Cells(i, 2)
            LibelleFamille = Cells(i, 2)
            s = s + 1
            End If
    
    ' controle Tmax atteint tableau
    If s >= Tmax Then
    NomModule = NomModule & "_" & " AUGMENTER le stockage Tmax pour les FAMILLES Tmax= " & Tmax & " atteint"
    GoTo err_Chargement_PRESTATIONS_REGLEES
    End If
    
    i = i + 1
    
    Wend
    ' nombre de lignes familles chargées
    nbfamille = s - 1

Else 'traitement D'UNE SEULE FAMILLE

    Famille(s) = Famille_A_Traiter
    nbfamille = 1
    
End If

    
    '2ème chargement du libellé FAMILLE de shDonneesAffichage et cellule Cells(i, 2) ET CHARGEMENTS DES LIBELLES ACTES POUR CHAQUE FAMILLE
        ' il est IMPORTANT DE CLASSER LES CODES ACTES PAR ORDRE APHABETIQUE (FAMILLE ET COED ACTES) DANS shDonneesAffichage

    i = 2  ' numéro de ligne = 2
    s = 1
    
If Famille_A_Traiter = "TOUTES" Then ' famille lue dans shDonneesAffichage 'traitement DE TOUTES LES FAMILLES

    While Cells(i, 2) <> ""  ' famille lue Cells(i, 2) Dans shDonneesAffichage
    ' Pour la famille lue Cells(i, 2)
    ' 2 ème calcul avec la feuille shDonnees la somme de tous les actes (somme de la colonne H) si la famille en colonne H = la famille traitée Cells(i, 2) et l'année (colonne F) = ANNEE_2
    
    LibelleFamille = ""
    LibelleActe = ""
    annee = ANNEE_2

            If s = 1 Then
            Acte(s) = Cells(i, 3)                   ' chargement du libellé acte
            LibelleActe = Cells(i, 3)
            Famille_Bis(s) = Cells(i, 2)
            LibelleFamille = Cells(i, 2)
            ' chargement de la famille lue dans la table FAMILLE_BIS
            s = s + 1
            ElseIf Acte(s - 1) <> Cells(i, 3) Then  ' chargement du libellé acte si l'acte précédent est différent de l'acte lu
            Acte(s) = Cells(i, 3)                   ' chargement du libellé acte
            LibelleActe = Cells(i, 3)
            Famille_Bis(s) = Cells(i, 2)
            LibelleFamille = Cells(i, 2)
            ' chargement de la famille lue dans la table FAMILLE_BIS (il y a autant de famille que de code actes)
            s = s + 1
                    
            End If
            
            ' controle Tmax atteint tableau
            If s >= Tmax Then
            NomModule = NomModule & "_" & " AUGMENTER le stockage Tmax pour les FAMILLES Tmax= " & Tmax & " atteint"
            GoTo err_Chargement_PRESTATIONS_REGLEES
            End If

    i = i + 1
    

    Wend
    ' nombre de lignes actes chargées
    NbActe = s - 1

Else 'traitement D'UNE SEULE FAMILLE

    While Cells(i, 2) <> ""   ' famille lue dans shDonneesAffichage
    
    If Cells(i, 2) = Famille_A_Traiter Then ' on charge les codes actes
    
    ' Pour la famille lue Cells(i, 2)
    ' 2 ème calcul avec la feuille shDonnees la somme de tous les actes (somme de la colonne H) si la famille en colonne H = la famille traitée Cells(i, 2) et l'année (colonne F) = ANNEE_2
    
    LibelleFamille = ""
    LibelleActe = ""
    annee = ANNEE_2

            If s = 1 Then
            Acte(s) = Cells(i, 3)                   ' chargement du libellé acte
            LibelleActe = Cells(i, 3)
            Famille_Bis(s) = Cells(i, 2)
            LibelleFamille = Cells(i, 2)
            s = s + 1
            ElseIf Acte(s - 1) <> Cells(i, 3) Then  ' chargement du libellé acte si l'acte précédent est différent de l'acte lu
            Acte(s) = Cells(i, 3)                   ' chargement du libellé acte
            LibelleActe = Cells(i, 3)
            Famille_Bis(s) = Cells(i, 2)
            LibelleFamille = Cells(i, 2)
            s = s + 1
            
            ' controle Tmax atteint tableau
            If s >= Tmax Then
            NomModule = NomModule & "_" & " AUGMENTER le stockage Tmax pour les FAMILLES Tmax= " & Tmax & " atteint"
            GoTo err_Chargement_PRESTATIONS_REGLEES
            End If

            End If
    Else
    End If
    i = i + 1
    
    Wend
    ' nombre de lignes actes chargées
    NbActe = s - 1

End If
    
'***  REMPLISSAGE de la feuille shResultats avec les LIBELLES FAMILLES et les LIBELLES ACTES   ***

    shResultats.Select
    
    ' INSERTION DES LIGNES FAMILLES à partir de la ligne 15 (la famille est en colonne 3)
    Cells(LigneDebut + 1, 3) = Famille(1) ' chargement de la 1 ère famille en ligne 15 qui est vide et en colonne 3
    For k = 2 To (nbfamille)
    rangeinsert = k + LigneDebut & ":" & k + LigneDebut  ' chargement du numéro de la ligne suivante à insérer
    Rows(rangeinsert).Select                            ' selection de la ligne à insérer
    Selection.Insert Shift:=xlDown                      ' insertion de la ligne à insérer
    Cells(LigneDebut + k, 3) = Famille(k)        ' chargement de la famille suivante dans la ligne qui vient d'être insérée
    solde = solde + 1
    Next k
    
    
    ' INSERTION DES LIGNES ACTES à partir de la ligne 15 (la famille est en colonne 3) pour chaque FAMILLE_BIS
    t = 0
    For k = 1 To (NbActe)
    rangeinsert = t + k + LigneDebut + 1 & ":" & t + k + LigneDebut + 1 ' chargement du numéro de la ligne à insérer (1 er acte pour la 1 ère famille en ligne 15+1 en colonne 4)
    Rows(rangeinsert).Select                     ' selection de la ligne à insérer
    Selection.Insert Shift:=xlDown               ' insertion de la ligne à insérer
    Cells(LigneDebut + 1 + k + t, 4) = Acte(k)              ' chargement du code acte en colonne 4 dans la ligne qui vient d'être insérée
    solde = solde + 1
    If Famille_Bis(k) <> Famille_Bis(k + 1) Then
    t = t + 1
    End If
    
    Next k
    
    ' à partir de la ligne 15 on recherche le libellé "Total général" dans la cellule famille Cells(f, 3)
    ' on définit un format pour les colonnes (C à Q) uniquement pour les lignes ACTES (Cells(f, 4) <> "") la ligne FAMILLE n'est pas traitée
    
    f = LigneDebut + 1
    
    While Cells(f, 3) <> "Total général"
    
    If Cells(f, 4) <> "" Then
    Range("C" & f & ":Q" & f).Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    End If
    
    f = f + 1
    Wend
    
' recherche de la ligne p = p - (nblignes entete = 15)  pour la ligne "Total général"
    p = 1
    While Cells(p, 3) <> "Total général"
    p = p + 1
    Wend
    p = p - (LigneDebut + 1)
    

'*** CALCUL DES CUMULS du Total général  ****

' RAZ des totaux pour toutes les familles
    
    Total_NOUS = 0          ' montant total REMBNOUS
    Total_NOUS_ANNEE_1 = 0  ' montant total REMBNOUS ANNEE_1
    Total_R = 0

' itération sur les FAMILLES
    
    LibelleFamille = ""
    LibelleActe = ""
    annee = ANNEE_2
    annee_N_1 = ANNEE_1
    
    For k = 1 To p
    If Cells(LigneDebut + k, 3) <> "" Then  ' traitement DES FAMILLE (la cellule Cells(LigneDebut + k, 3)="")
    
        LibelleFamille = Cells(LigneDebut + k, 3) ' libellé Famille
            
            Total_NOUS = Total_NOUS + Remboursement_Kerialis_Famille(shDonnees, annee, LibelleFamille)   ' Remboursement KERIALIS
            Total_NOUS_ANNEE_1 = Total_NOUS_ANNEE_1 + Remboursement_Kerialis_Famille(shDonnees, annee_N_1, LibelleFamille)   ' Remboursement KERIALIS ANNEE_1
        
    End If
    
    Next k


'****
'*** REMPLISSAGE DE LA LIGNE FAMILLE OU ACTE  dans les colonnes 10
'*** et calcul des CUMULS ***

    LibelleFamille = ""
    LibelleActe = ""
    annee = ANNEE_2
    annee_N_1 = ANNEE_1
    
    ' ajout des années dans le tableau
    Cells(8, 3) = Cells(8, 3) & " - années " & annee_N_1 & " et " & annee
    Cells(LigneDebut, 5) = annee_N_1
    Cells(LigneDebut, 6) = annee
    Cells(LigneDebut, 7) = " Variation " & annee & " / " & annee_N_1
    
    For k = 1 To p
    
    If Cells(LigneDebut + k, 3) = "" Then  ' traitement de la ligne ACTE (la cellule Cells(LigneDebut + k, 3)="")
    
    '*** Traitement ligne ACTE ***
    LibelleActe = Cells(LigneDebut + k, 4) ' libellé Acte
            Cells(LigneDebut + k, 5) = Remboursement_Kerialis_Acte(shDonnees, annee_N_1, LibelleFamille, LibelleActe) ' Remboursement KERIALIS annee_N_1
            Cells(LigneDebut + k, 6) = Remboursement_Kerialis_Acte(shDonnees, annee, LibelleFamille, LibelleActe)    ' Remboursement KERIALIS
            
            If Cells(LigneDebut + k, 5) <> 0 Then
            Cells(LigneDebut + k, 7) = Cells(LigneDebut + k, 6) / Cells(LigneDebut + k, 5) - 1
            Else
            End If
            
    Else ' traitement de la ligne FAMILLE (elle contient le CUMUL des données actes de la famille)
    
    '*** Traitement ligne FAMILLE ***
    LibelleFamille = Cells(LigneDebut + k, 3) ' libellé Famille
            Cells(LigneDebut + k, 5) = Remboursement_Kerialis_Famille(shDonnees, annee_N_1, LibelleFamille)     ' Remboursement KERIALIS
            Cells(LigneDebut + k, 6) = Remboursement_Kerialis_Famille(shDonnees, annee, LibelleFamille)        ' Remboursement KERIALIS
            
            If Cells(LigneDebut + k, 5) <> 0 Then
            Cells(LigneDebut + k, 7) = Cells(LigneDebut + k, 6) / Cells(LigneDebut + k, 5) - 1
            Else
            End If
            
    End If
    
    Next k

'*** REMPLISSAGE DE LA LIGNE TOTAL GENERAL colonnes 6 7 8 9 10 11 12 13 *** *** EXPERIENCE 15 16 17 ***

    
    ' Remboursement KERIALIS
    Cells(LigneDebut + p + 1, 5) = Total_NOUS_ANNEE_1
    Cells(LigneDebut + p + 1, 6) = Total_NOUS

    If Cells(LigneDebut + p + 1, 5) <> 0 Then
    Cells(LigneDebut + p + 1, 7) = Cells(LigneDebut + p + 1, 6) / Cells(LigneDebut + p + 1, 5) - 1
    Else
    End If

'*** SUPPRESSION DE TOUTES LES LIGNES SI le taux de couverture colonnes (13 ou 17) = 0
    t = 0
    For k = 1 To p
    'If Cells(LigneDebut + k - t, 5) = "" Then
    
    'Calcul cumul des lignes *** colonnes 5 6 7 8 9 10 11 12 13 *** EXPERIENCE 15 16 17 ***
    SommeLignesPresta = Cells(LigneDebut + k - t, 5) + Cells(LigneDebut + k - t, 6)
    
    'If Cells(LigneDebut + k - t, 3) = "TEST_1" Then
    'Stop
    'End If
    
    'Suppression ligne si montants = 0
    If SommeLignesPresta = 0 Then
    'Rows(14 + k - t & ":" & 14 + k - t).Select
    Rows(LigneDebut + k - t & ":" & LigneDebut + k - t).Select
    Selection.Delete Shift:=xlDown
    solde = solde - 1
    t = t + 1
    End If
    Next k
    
'*** en FIN DE TABLEAU on ajoute ou en supprime le nombre de lignes = solde (solde étant le nombre de lignes ajoutées ou supprimées précédemment)
     If solde > 0 Then  ' on SUPPRIME un nombre s de lignes à partir de (p+21)
        For i = 1 To solde
        rangeinsert = p + 21 & ":" & p + 21
        Rows(rangeinsert).Select
        Selection.Delete Shift:=xlDown
        Next i
    Else
        For i = 1 To -solde ' on AJOUTE s de lignes à partir de (p+21)
        rangeinsert = p + 21 & ":" & p + 21
        Rows(rangeinsert).Select
        Selection.Insert Shift:=xlDown
        Next i
    
    End If
    
Application.Cursor = xlDefault

'*** test ecriture dans Erreurs
'erreur = 1 / 0

  Exit Function


' erreur
err_Chargement_PRESTATIONS_REGLEES:
  
  Application.Cursor = xlDefault
  
 ' recherche du dernier numéro de la ligne <> "" pour permettre d'ajouter un message à la ligne suivante
    p = 1
    While shErreurs.Cells(p, 1) <> ""
    p = p + 1
    Wend
    p = p - 1
    NoLigneEnErreur = p - 1
  
  MsgBox "Erreur dans Chargement_PRESTATIONS_REGLEES() : " & Err.Number & vbLf & Err.Description, vbCritical
  
  ' affichage du message d'erreur dans la feuille "Erreurs"
  MessageErreur = "Chargement_PRESTATIONS_REGLEES() : " & Err.Number & " - " & Err.Description

  NoLigneEnErreur = NoLigneEnErreur + 1
  'shErreurs.Range("A1").Offset(NoLigneEnErreur, 0).Value = shResultats.Range("C2").Offset(NoLigneEnErreur, 0).Value
  shErreurs.Range("A1").Offset(NoLigneEnErreur, 0).Value = NoLigneEnErreur
  shErreurs.Range("B1").Offset(NoLigneEnErreur, 0).Value = NomModule
  shErreurs.Range("C1").Offset(NoLigneEnErreur, 0).Value = LibelleFamille
  shErreurs.Range("D1").Offset(NoLigneEnErreur, 0).Value = LibelleActe
  shErreurs.Range("E1").Offset(NoLigneEnErreur, 0).Font.Color = vbRed
  shErreurs.Range("E1").Offset(NoLigneEnErreur, 0).Value = "Erreur : " & Err.Number & " - " & Err.Description
    
  'shResultats.Protect PROTECT_PASSWORD
  
  'Resume Next
    
End Function
