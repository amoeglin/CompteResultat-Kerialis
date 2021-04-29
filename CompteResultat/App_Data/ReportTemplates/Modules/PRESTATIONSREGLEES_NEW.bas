Attribute VB_Name = "PRESTATIONSREGLEES_NEW"
Sub PRESTATIONSREGLEES_NEW()
Attribute PRESTATIONSREGLEES_NEW.VB_Description = "test"
Attribute PRESTATIONSREGLEES_NEW.VB_ProcData.VB_Invoke_Func = " \n14"

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


' CHARGEMENT des feuilles

Set shDonnees = Worksheets("DATA PREST")
Set shDonneesAffichage = Worksheets("AFFICHAGE")
Set shDonnéesDemo = Worksheets("DATA DEMO")
Set shDonneesExperience = Worksheets("DATA EXP")
'Set shResultats à charger
Set shErreurs = Worksheets("Erreurs")


'*** RAZ des lignes dans la feuille Erreurs  ***
bValue = RAZ_shErreurs(shErreurs)
    
' TRAITEMENT de TOUTES LES PRESTATIONS - Onglet Prestations Réglées
Set shResultats = Worksheets("Prestations Réglées")
NomModule = "Onglet Prestations Réglées"
Famille_A_Traiter = "TOUTES" ' traitement de toutes les familles
bValue = PRESTATIONSREGLEES_GENERIQUE(Famille_A_Traiter, NomModule, shDonnees, shDonneesAffichage, shDonnéesDemo, shDonneesExperience, shResultats, shErreurs)
    
' TRAITEMENT OPTIQUE - Onglet Prestations Réglées_OPTIQUE
Set shResultats = Worksheets("Prestations Réglées_OPTIQUE")
NomModule = "Prestations Réglées_OPTIQUE"
Sheets("AFFICHAGE").Select
Famille_A_Traiter = Cells(5, 13) ' optique
If Famille_A_Traiter <> "" Then
bValue = PRESTATIONSREGLEES_GENERIQUE(Famille_A_Traiter, NomModule, shDonnees, shDonneesAffichage, shDonnéesDemo, shDonneesExperience, shResultats, shErreurs)
Else
End If

' TRAITEMENT DENTAIRE - Onglet Prestations Réglées_DENTAIRE
Set shResultats = Worksheets("Prestations Réglées_DENTAIRE")
NomModule = "Prestations Réglées_DENTAIRE"
Sheets("AFFICHAGE").Select
Famille_A_Traiter = Cells(6, 13) ' dentaire
If Famille_A_Traiter <> "" Then
bValue = PRESTATIONSREGLEES_GENERIQUE(Famille_A_Traiter, NomModule, shDonnees, shDonneesAffichage, shDonnéesDemo, shDonneesExperience, shResultats, shErreurs)
Else
End If
    
End Sub

Function PRESTATIONSREGLEES_GENERIQUE(Famille_A_Traiter As String, NomModule As String, shDonnees As Worksheet, shDonneesAffichage As Worksheet, shDonnéesDemo As Worksheet, shDonneesExperience As Worksheet, shResultats As Worksheet, shErreurs As Worksheet) As Boolean

'Option Explicit
'Version 1-0 le 29/03/2021
' Mot de passe de protection des feuilles XL
'Public Const PROTECT_PASSWORD As String = "CMPASS" 'déclaration des feuilles

Dim NomOnglet As String

Dim Famille(200) As String
Dim Acte(200) As String
Dim Famille_Bis(200) As String
Dim bValue As Boolean

Dim annee As Double
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
            
            If LibelleFamille = TEST_2 Then
            Stop
            End If

            
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
    
' définition des zones dans la feuille shDonnéesDemo
        'FDEMO = "DATA DEMO"
        ANNEEDEMO = "A:A"
        SOCIETEDEMO = "B:B"
        SEXEDEMO = "D:D"
        AGEDEMO = "H:H"
        DEMO = "G:G"
        LIENDEMO = "E:E"
        TRANCHEDEMO = "F:F"
        COLDEMO = "J:J"
   
' calcul de la durée d'exposition = nombre d'assurés dans la démo pour ANNEE_2
    Exposition = Application.WorksheetFunction.SumIfs(shDonnéesDemo.Range(DEMO), shDonnéesDemo.Range(ANNEEDEMO), ANNEE_2)
    

'*** CALCUL DES CUMULS du Total général  ****

' RAZ des totaux pour toutes les familles
    
    Total_NB = 0    ' nombre d'actes
    Total_FR = 0    ' montant total FRAISREELS
    Total_SS = 0    ' montant total REMBSS
    Total_AUTRES = 0 ' montant total REMBANNEXE
    Total_NOUS = 0  ' montant total REMBNOUS
    Total_R = 0

    Total_NB_Experience = 0    ' nombre d'actes
    Total_FR_Experience = 0    ' montant total FRAISREELS
    Total_SS_Experience = 0    ' montant total REMBSS
    Total_AUTRES_Experience = 0 ' montant total REMBANNEXE
    Total_NOUS_Experience = 0  ' montant total REMBNOUS
    Total_R_Experience = 0


' itération sur les FAMILLES
    
    LibelleFamille = ""
    LibelleActe = ""
    annee = ANNEE_2
    
    For k = 1 To p
    If Cells(LigneDebut + k, 3) <> "" Then  ' traitement DES FAMILLE (la cellule Cells(LigneDebut + k, 3)="")
    
        LibelleFamille = Cells(LigneDebut + k, 3) ' libellé Famille
        
        If Nb_Acte_Famille(shDonnees, annee, LibelleFamille) <> 0 Or Nb_Acte_Famille(shDonneesExperience, annee, LibelleFamille) <> 0 Then  ' nb d'actes pour l'Famille traité
            
            ' Pour les Prestations
            Total_NB = Total_NB + Nb_Acte_Famille(shDonnees, annee, LibelleFamille)                     ' nb d'actes pour l'Famille traité
            Total_FR = Total_FR + Frais_Reels_Famille(shDonnees, annee, LibelleFamille)                 ' Frais réels pour l'Famille traité
            Total_SS = Total_SS + Remboursement_SS_Famille(shDonnees, annee, LibelleFamille)            ' Remboursement SS
            Total_AUTRES = Total_AUTRES + Remboursement_Autres_Regimes_Famille(shDonnees, annee, LibelleFamille) ' Remboursement Autres régimes
            Total_NOUS = Total_NOUS + Remboursement_Kerialis_Famille(shDonnees, annee, LibelleFamille)   ' Remboursement KERIALIS
        
            ' Pour les Prestations EXPERIENCE
            Total_NB = Total_NB + Nb_Acte_Famille(shDonneesExperience, annee, LibelleFamille)             ' nb d'actes pour l'Famille traité
            Total_FR = Total_FR + Frais_Reels_Famille(shDonneesExperience, annee, LibelleFamille)         ' Frais réels pour l'Famille traité
            Total_SS = Total_SS + Remboursement_SS_Famille(shDonneesExperience, annee, LibelleFamille)    ' Remboursement SS
            Total_AUTRES = Total_AUTRES + Remboursement_Autres_Regimes_Famille(shDonneesExperience, annee, LibelleFamille) ' Remboursement Autres régimes
            Total_NOUS = Total_NOUS + Remboursement_Kerialis_Famille(shDonneesExperience, annee, LibelleFamille)   ' Remboursement KERIALIS
        
        End If
    End If
    
    Next k
    
    
'*** CALCUL DES CUMULS du Total général  ****

' RAZ des totaux pour toutes les familles
    
    Total_NB = 0    ' nombre d'actes
    Total_FR = 0    ' montant total FRAISREELS
    Total_SS = 0    ' montant total REMBSS
    Total_AUTRES = 0 ' montant total REMBANNEXE
    Total_NOUS = 0  ' montant total REMBNOUS
    Total_R = 0
    
    Total_NB_Experience = 0    ' nombre d'actes
    Total_FR_Experience = 0    ' montant total FRAISREELS
    Total_SS_Experience = 0    ' montant total REMBSS
    Total_AUTRES_Experience = 0 ' montant total REMBANNEXE
    Total_NOUS_Experience = 0  ' montant total REMBNOUS
    Total_R_Experience = 0


' itération sur les FAMILLES
    
    LibelleFamille = ""
    LibelleActe = ""
    annee = ANNEE_2
    
    For k = 1 To p
    If Cells(LigneDebut + k, 3) <> "" Then  ' traitement DES FAMILLE (la cellule Cells(LigneDebut + k, 3)="")
    
        LibelleFamille = Cells(LigneDebut + k, 3) ' libellé Famille
        
        If Nb_Acte_Famille(shDonnees, annee, LibelleFamille) <> 0 Or Nb_Acte_Famille(shDonneesExperience, annee, LibelleFamille) <> 0 Then ' nb d'actes pour l'Famille traité
            
            Total_NB = Total_NB + Nb_Acte_Famille(shDonnees, annee, LibelleFamille)                     ' nb d'actes pour l'Famille traité
            Total_FR = Total_FR + Frais_Reels_Famille(shDonnees, annee, LibelleFamille)                 ' Frais réels pour l'Famille traité
            Total_SS = Total_SS + Remboursement_SS_Famille(shDonnees, annee, LibelleFamille)            ' Remboursement SS
            Total_AUTRES = Total_AUTRES + Remboursement_Autres_Regimes_Famille(shDonnees, annee, LibelleFamille) ' Remboursement Autres régimes
            Total_NOUS = Total_NOUS + Remboursement_Kerialis_Famille(shDonnees, annee, LibelleFamille)   ' Remboursement KERIALIS
        
            ' Pour les Prestations EXPERIENCE
            Total_NB_Experience = Total_NB_Experience + Nb_Acte_Famille(shDonneesExperience, annee, LibelleFamille)             ' nb d'actes pour l'Famille traité
            Total_FR_Experience = Total_FR_Experience + Frais_Reels_Famille(shDonneesExperience, annee, LibelleFamille)         ' Frais réels pour l'Famille traité
            Total_SS_Experience = Total_SS_Experience + Remboursement_SS_Famille(shDonneesExperience, annee, LibelleFamille)    ' Remboursement SS
            Total_AUTRES_Experience = Total_AUTRES_Experience + Remboursement_Autres_Regimes_Famille(shDonneesExperience, annee, LibelleFamille) ' Remboursement Autres régimes
            Total_NOUS_Experience = Total_NOUS_Experience + Remboursement_Kerialis_Famille(shDonneesExperience, annee, LibelleFamille)   ' Remboursement KERIALIS
       
        End If
    End If
    
    Next k


'****
'*** REMPLISSAGE DE LA LIGNE FAMILLE OU ACTE  dans les colonnes 3 5 6 7 8 9 10 11 12 13 *** EXPERIENCE 15 16 17 ***
'*** et calcul des CUMULS ***

    LibelleFamille = ""
    LibelleActe = ""
    annee = ANNEE_2
    
    For k = 1 To p
    If Cells(LigneDebut + k, 3) = "" Then  ' traitement de la ligne ACTE (la cellule Cells(LigneDebut + k, 3)="")
    
    '*** Traitement ligne ACTE ***
    'LibelleFamille déjà chargé dans la partie traitement de la ligne FAMILLE
    LibelleActe = Cells(LigneDebut + k, 4) ' libellé Acte
    
    If Nb_Acte_Famille(shDonnees, annee, LibelleFamille) <> 0 Or Nb_Acte_Famille(shDonneesExperience, annee, LibelleFamille) <> 0 Then ' nb d'actes pour l'Acte traité
    
        Cells(LigneDebut + k, 5) = Nb_Acte_Acte(shDonnees, annee, LibelleFamille, LibelleActe) ' nb d'actes pour l'Acte traité
    
        If Exposition <> 0 Then
        Cells(LigneDebut + k, 6) = Cells(LigneDebut + k, 5) / Exposition 'Fréquence
        End If
        
        Cells(LigneDebut + k, 7) = Frais_Reels_Acte(shDonnees, annee, LibelleFamille, LibelleActe)                ' Frais réels pour l'acte traité
        Cells(LigneDebut + k, 8) = Remboursement_SS_Acte(shDonnees, annee, LibelleFamille, LibelleActe)           ' Remboursement SS
        Cells(LigneDebut + k, 9) = Remboursement_Autres_Regimes_Acte(shDonnees, annee, LibelleFamille, LibelleActe) ' Remboursement Autres régimes
        Cells(LigneDebut + k, 10) = Remboursement_Kerialis_Acte(shDonnees, annee, LibelleFamille, LibelleActe)    ' Remboursement KERIALIS
        
        If Cells(LigneDebut + k, 5) <> 0 Then
        Cells(LigneDebut + k, 11) = Cells(LigneDebut + k, 10) / Cells(LigneDebut + k, 5)            ' Remboursement moyen par acte
        End If
        
        If Total_NOUS <> 0 Then
        Cells(LigneDebut + k, 12) = Cells(LigneDebut + k, 10) / Total_NOUS                          ' Répartition du remboursements KERIALIS en % du total remboursements KERIALIS
        Else
        Cells(LigneDebut + k, 12) = 0
        End If
        
        ' taux de couverture
        If Cells(LigneDebut + k, 7) <> 0 Then
        Cells(LigneDebut + k, 13) = (Cells(LigneDebut + k, 8) + Cells(LigneDebut + k, 10) + Cells(LigneDebut + k, 9)) / Cells(LigneDebut + k, 7)
        Else
        Cells(LigneDebut + k, 13) = 0
        End If
    
       '*** ligne ACTE pour L'EXPERIENCE  ***
       
       T_5 = Frais_Reels_Acte(shDonneesExperience, annee, LibelleFamille, LibelleActe)                   'FRAISREELS
       T_6 = Remboursement_SS_Acte(shDonneesExperience, annee, LibelleFamille, LibelleActe)              'REMBSS
       T_7 = Remboursement_Autres_Regimes_Acte(shDonneesExperience, annee, LibelleFamille, LibelleActe)  'REMBANNEXE
       T_8 = Remboursement_Kerialis_Acte(shDonneesExperience, annee, LibelleFamille, LibelleActe)        'REMBNOUS
       
       'Remboursement moyen par acte EXPERIENCE
       
    'If LibelleActe = "LUNETTES" Then
    'Stop
    'End If

       
        Nb_Acte_Acte_Experience = Nb_Acte_Acte(shDonneesExperience, annee, LibelleFamille, LibelleActe) ' nb d'actes pour la Famille traité EXPERIENCE
        If Nb_Acte_Acte_Experience <> 0 Then
        Cells(LigneDebut + k, 15) = T_8 / Nb_Acte_Acte_Experience
        Else
        Cells(LigneDebut + k, 15) = 0
        End If
        
        'Répartition des Remboursements KERIALIS EXPERIENCE en % du total des Remboursements KERIALIS   T8 /Total_NOUS_Experience
        If Total_NOUS_Experience <> 0 Then
        Cells(LigneDebut + k, 16) = T_8 / Total_NOUS_Experience
        Else
        Cells(LigneDebut + k, 16) = 0
        End If
        
        'Taux de couverture EXPERIENCE = (T_6 + T_7 + T_8) / T_5
        If T_5 <> 0 Then
        Cells(LigneDebut + k, 17) = (T_6 + T_7 + T_8) / T_5
        Else
        Cells(LigneDebut + k, 17) = 0
        End If
    
    
    
    End If
    
    Else ' traitement de la ligne FAMILLE (elle contient le CUMUL des données actes de la famille)
    
    '*** Traitement ligne FAMILLE ***
    
    LibelleFamille = Cells(LigneDebut + k, 3) ' libellé Famille
    
    'If LibelleFamille = "TEST_2" Then
    'Stop
    'End If
     
    If Nb_Acte_Famille(shDonnees, annee, LibelleFamille) <> 0 Or Nb_Acte_Famille(shDonneesExperience, annee, LibelleFamille) <> 0 Then ' nb d'actes pour l'Famille traité
    
        Cells(LigneDebut + k, 5) = Nb_Acte_Famille(shDonnees, annee, LibelleFamille) ' nb d'actes pour l'Famille traité
    
        If Exposition <> 0 Then
        Cells(LigneDebut + k, 6) = Cells(LigneDebut + k, 5) / Exposition 'Fréquence
        End If
        
        Cells(LigneDebut + k, 7) = Frais_Reels_Famille(shDonnees, annee, LibelleFamille)                 ' Frais réels pour l'Famille traité
        Cells(LigneDebut + k, 8) = Remboursement_SS_Famille(shDonnees, annee, LibelleFamille)            ' Remboursement SS
        Cells(LigneDebut + k, 9) = Remboursement_Autres_Regimes_Famille(shDonnees, annee, LibelleFamille) ' Remboursement Autres régimes
        Cells(LigneDebut + k, 10) = Remboursement_Kerialis_Famille(shDonnees, annee, LibelleFamille)     ' Remboursement KERIALIS
        
        If Cells(LigneDebut + k, 5) <> 0 Then
        Cells(LigneDebut + k, 11) = Cells(LigneDebut + k, 10) / Cells(LigneDebut + k, 5)            ' Remboursement moyen par Famille
        End If
        
        If Total_NOUS <> 0 Then
        Cells(LigneDebut + k, 12) = Cells(LigneDebut + k, 10) / Total_NOUS                          ' Répartition du remboursements KERIALIS en % du total remboursements KERIALIS
        Else
        Cells(LigneDebut + k, 12) = 0
        End If
        
        Total_R = Total_R + Cells(LigneDebut + k, 12) ' calcul du Total_R
        
        ' taux de couverture
        If Cells(LigneDebut + k, 7) <> 0 Then
        Cells(LigneDebut + k, 13) = (Cells(LigneDebut + k, 8) + Cells(LigneDebut + k, 10) + Cells(LigneDebut + k, 9)) / Cells(LigneDebut + k, 7)
        Else
        Cells(LigneDebut + k, 13) = 0
        End If
    
       '*** ligne FAMILLE pour L'EXPERIENCE  ***
       
       T_5 = Frais_Reels_Famille(shDonneesExperience, annee, LibelleFamille)                    'FRAISREELS
       T_6 = Remboursement_SS_Famille(shDonneesExperience, annee, LibelleFamille)               'REMBSS
       T_7 = Remboursement_Autres_Regimes_Famille(shDonneesExperience, annee, LibelleFamille)   'REMBANNEXE
       T_8 = Remboursement_Kerialis_Famille(shDonneesExperience, annee, LibelleFamille)         'REMBNOUS
       
       'Remboursement moyen par acte EXPERIENCE
       
        Nb_Acte_Famille_EXPERENCE = Nb_Acte_Famille(shDonneesExperience, annee, LibelleFamille) ' nb d'actes pour la Famille traité EXPERIENCE
        If Nb_Acte_Famille_EXPERENCE <> 0 Then
        Cells(LigneDebut + k, 15) = T_8 / Nb_Acte_Famille_EXPERENCE
        Else
        Cells(LigneDebut + k, 15) = 0
        End If
        
        'Répartition des Remboursements KERIALIS EXPERIENCE en % du total des Remboursements KERIALIS   T8 /Total_NOUS_Experience
        If Total_NOUS_Experience <> 0 Then
        Cells(LigneDebut + k, 16) = T_8 / Total_NOUS_Experience
        Else
        Cells(LigneDebut + k, 16) = 0
        End If
        
        Total_R_Experience = Total_R_Experience + Cells(LigneDebut + k, 16)   ' calcul du Total_R_Experience
        
        'Taux de couverture EXPERIENCE = (T_6 + T_7 + T_8) / T_5
        If T_5 <> 0 Then
        Cells(LigneDebut + k, 17) = (T_6 + T_7 + T_8) / T_5
        Else
        Cells(LigneDebut + k, 17) = 0
        End If
    
    
    End If
    End If
    
    Next k

'*** REMPLISSAGE DE LA LIGNE TOTAL GENERAL colonnes 6 7 8 9 10 11 12 13 *** *** EXPERIENCE 15 16 17 ***

    ' quantité nombres d'actes
    Cells(LigneDebut + p + 1, 5) = Total_NB
    
    'Fréquence
    If Exposition <> 0 Then
    Cells(LigneDebut + p + 1, 6) = Total_NB / Exposition
    End If
    
    ' Frais réels
    Cells(LigneDebut + p + 1, 7) = Total_FR
    
    ' Remboursement SS
    Cells(LigneDebut + p + 1, 8) = Total_SS
    
    ' Remboursement Autres régimes
    Cells(LigneDebut + p + 1, 9) = Total_AUTRES
    
    ' Remboursement KERIALIS
    Cells(LigneDebut + p + 1, 10) = Total_NOUS
    
    ' Remboursement moyen par acte
    If Total_NB <> 0 Then
    Cells(LigneDebut + p + 1, 11) = Total_NOUS / Total_NB
    End If
    
    ' Répartition des Remboursements KERIALIS (en % du total des Remb kerialis)
    Cells(LigneDebut + p + 1, 12) = Total_R
    
    ' taux de couverture
    If Total_FR <> 0 Then
    Cells(LigneDebut + p + 1, 13) = (Total_SS + Total_AUTRES + Total_NOUS) / Total_FR  ' 13=(8+9+10)/7
    End If
    
    '*** ligne pour L'EXPERIENCE  ***
    
    ' Remboursement moyen par acte
    If Total_NB_Experience <> 0 Then
    Cells(LigneDebut + p + 1, 15) = Total_NOUS_Experience / Total_NB_Experience
    End If
    
    ' Répartition des Remboursements KERIALIS (en % du total des Remb kerialis)
    Cells(LigneDebut + p + 1, 16) = Total_R_Experience
    
    ' taux de couverture
    If Total_FR_Experience <> 0 Then
    Cells(LigneDebut + p + 1, 17) = (Total_SS_Experience + Total_AUTRES_Experience + Total_NOUS_Experience) / Total_FR_Experience  ' 13=(8+9+10)/7
    End If

'*** SUPPRESSION DE TOUTES LES LIGNES SI le taux de couverture colonnes (13 ou 17) = 0
    t = 0
    For k = 1 To p
    'If Cells(LigneDebut + k - t, 5) = "" Then
    
    'Calcul cumul des lignes *** colonnes 5 6 7 8 9 10 11 12 13 *** EXPERIENCE 15 16 17 ***
    SommeLignesPresta = 0
    SommeLignesPresta = Cells(LigneDebut + k - t, 5) + Cells(LigneDebut + k - t, 6) + Cells(LigneDebut + k - t, 7)
    SommeLignesPresta = SommeLignesPresta + Cells(LigneDebut + k - t, 8) + Cells(LigneDebut + k - t, 9) + Cells(LigneDebut + k - t, 10) + Cells(LigneDebut + k - t, 11)
    SommeLignesPresta = SommeLignesPresta + Cells(LigneDebut + k - t, 12) + Cells(LigneDebut + k - t, 13)
    
    SommeLignesPrestaExperience = 0
    SommeLignesPrestaExperience = Cells(LigneDebut + k - t, 15) + Cells(LigneDebut + k - t, 16) + Cells(LigneDebut + k - t, 17)
    
    'If Cells(LigneDebut + k - t, 3) = "TEST_1" Then
    'Stop
    'End If
    
    'Suppression ligne si montants = 0
    If SommeLignesPresta = 0 And SommeLignesPrestaExperience = 0 Then
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

' recherche de la ligne p = p - (nblignes entete = 15)  pour la ligne "Total général"
    p = 1
    While Cells(p, 3) <> "Total général"
    p = p + 1
    Wend
    p = p - (LigneDebut + 1)
    
    zone1 = "$D$" & LigneDebut + 2 & ":$D$" & LigneDebut + p
    zone2 = "$J$" & LigneDebut + 2 & ":$J$" & LigneDebut + p

'*** Graphique  OPTIQUE

If NomModule = "Prestations Réglées_OPTIQUE" Then
    ' recherche nom Onglet
    NomOnglet = ActiveSheet.Name
    NomOnglet = "'" & NomOnglet & "'!"

    ActiveSheet.ChartObjects("Graphique 2").Activate
    ActiveChart.PlotArea.Select
    'ActiveChart.SetSourceData Source:=Range(
    '   "'Prestations Réglées_OPTIQUE'!" & zone1 & ",'Prestations Réglées_OPTIQUE'!" & zone2)
    ActiveChart.SetSourceData Source:=Range(NomOnglet & zone1 & "," & NomOnglet & zone2)
    
    
Calculate
Else
End If

'*** Graphique  DENTAIRE

If NomModule = "Prestations Réglées_DENTAIRE" Then
    ' recherche nom Onglet
    NomOnglet = ActiveSheet.Name
    NomOnglet = "'" & NomOnglet & "'!"

    ActiveSheet.ChartObjects("Graphique 2").Activate
    ActiveChart.PlotArea.Select
    'ActiveChart.SetSourceData Source:=Range( _
    '   "'Prestations Réglées_DENTAIRE'!" & zone1 & ",'Prestations Réglées_DENTAIRE'!" & zone2)
    ActiveChart.SetSourceData Source:=Range(NomOnglet & zone1 & "," & NomOnglet & zone2)

Calculate
Else
End If

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

    

Function RAZ_shErreurs(shErreurs As Worksheet) As Boolean

' RAZ des lignes dans la feuille Erreurs
    
    shErreurs.Select
   ' recherche du dernier numéro de la ligne <> "" pour permettre de supprimer toutes les lignes
    p = 1 ' à partir de la 1 ème ligne
    While Cells(p, 1) <> ""
    p = p + 1
    Wend
    
    RangeDelete = 1 & ":" & p ' chargement des no de lignes à supprimer
    Rows(RangeDelete).Select
    Selection.Delete Shift:=xlDown
    
    ' titre
    Cells(1, 1) = "Numéro chronologique"
    Cells(1, 2) = "Traitement concerné"
    Cells(1, 3) = "Libellé Famille"
    Cells(1, 4) = "Libellé Acte"
    Cells(1, 5) = "Libellé de l'erreur"

End Function
'
'** ACTES PAR FAMILLE ***
'
Function Nb_Acte_Acte(shDonnees As Worksheet, annee As Double, LibelleFamille As String, LibelleActe As String) As Double
' quantité nombres d'actes
Nb_Acte_Acte = 0
Nb_Acte_Acte = Application.WorksheetFunction.SumIfs(shDonnees.Range("H:H"), shDonnees.Range("D:D"), annee, shDonnees.Range("F:F"), LibelleFamille, shDonnees.Range("E:E"), LibelleActe)
End Function

Function Frais_Reels_Acte(shDonnees As Worksheet, annee As Double, LibelleFamille As String, LibelleActe As String) As Double
' Frais_Reels_Acte
Frais_Reels_Acte = 0
Frais_Reels_Acte = Application.WorksheetFunction.SumIfs(shDonnees.Range("I:I"), shDonnees.Range("D:D"), annee, shDonnees.Range("F:F"), LibelleFamille, shDonnees.Range("E:E"), LibelleActe)
End Function
Function Remboursement_SS_Acte(shDonnees As Worksheet, annee As Double, LibelleFamille As String, LibelleActe As String) As Double
' Remboursement_SS_Acte
Remboursement_SS_Acte = 0
Remboursement_SS_Acte = Application.WorksheetFunction.SumIfs(shDonnees.Range("J:J"), shDonnees.Range("D:D"), annee, shDonnees.Range("F:F"), LibelleFamille, shDonnees.Range("E:E"), LibelleActe)
End Function

Function Remboursement_Autres_Regimes_Acte(shDonnees As Worksheet, annee As Double, LibelleFamille As String, LibelleActe As String) As Double
' Remboursement_Autres_Regimes_Acte
Remboursement_Autres_Regimes_Acte = 0
Remboursement_Autres_Regimes_Acte = Application.WorksheetFunction.SumIfs(shDonnees.Range("K:K"), shDonnees.Range("D:D"), annee, shDonnees.Range("F:F"), LibelleFamille, shDonnees.Range("E:E"), LibelleActe)
End Function
    
Function Remboursement_Kerialis_Acte(shDonnees As Worksheet, annee As Double, LibelleFamille As String, LibelleActe As String) As Double
' Remboursement_Kerialis_Acte
Remboursement_Kerialis_Acte = 0
Remboursement_Kerialis_Acte = Application.WorksheetFunction.SumIfs(shDonnees.Range("L:L"), shDonnees.Range("D:D"), annee, shDonnees.Range("F:F"), LibelleFamille, shDonnees.Range("E:E"), LibelleActe)
End Function
'
'** FAMILLE ***
'
Function Nb_Acte_Famille(shDonnees As Worksheet, annee As Double, LibelleFamille As String) As Double
' quantité nombres d'actes
Nb_Acte_Famille = 0
Nb_Acte_Famille = Application.WorksheetFunction.SumIfs(shDonnees.Range("H:H"), shDonnees.Range("D:D"), annee, shDonnees.Range("F:F"), LibelleFamille)
End Function

Function Frais_Reels_Famille(shDonnees As Worksheet, annee As Double, LibelleFamille As String) As Double
' Frais_Reels_Famille
Frais_Reels_Famille = 0
Frais_Reels_Famille = Application.WorksheetFunction.SumIfs(shDonnees.Range("I:I"), shDonnees.Range("D:D"), annee, shDonnees.Range("F:F"), LibelleFamille)
End Function

Function Remboursement_SS_Famille(shDonnees As Worksheet, annee As Double, LibelleFamille As String) As Double
' Remboursement_SS_Famille
Remboursement_SS_Famille = 0
Remboursement_SS_Famille = Application.WorksheetFunction.SumIfs(shDonnees.Range("J:J"), shDonnees.Range("D:D"), annee, shDonnees.Range("F:F"), LibelleFamille)
End Function

Function Remboursement_Autres_Regimes_Famille(shDonnees As Worksheet, annee As Double, LibelleFamille As String) As Double
' Remboursement_Autres_Regimes_Famille
Remboursement_Autres_Regimes_Famille = 0
Remboursement_Autres_Regimes_Famille = Application.WorksheetFunction.SumIfs(shDonnees.Range("K:K"), shDonnees.Range("D:D"), annee, shDonnees.Range("F:F"), LibelleFamille)
End Function
    
Function Remboursement_Kerialis_Famille(shDonnees As Worksheet, annee As Double, LibelleFamille As String) As Double
' Remboursement_Kerialis_Famille
Remboursement_Kerialis_Famille = 0
Remboursement_Kerialis_Famille = Application.WorksheetFunction.SumIfs(shDonnees.Range("L:L"), shDonnees.Range("D:D"), annee, shDonnees.Range("F:F"), LibelleFamille)
End Function

