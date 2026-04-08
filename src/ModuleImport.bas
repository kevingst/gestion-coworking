Attribute VB_Name = "ModuleImport"
' =============================================================================
' Module : ModuleImport
' Description : Import de données existantes depuis la feuille IMPORT
'
' USAGE FUTUR :
' Ce module est préparé pour permettre l'import de participants depuis un
' fichier Excel existant. La feuille IMPORT doit contenir les données à
' importer avec les mêmes colonnes que PARTICIPANTS :
'   ID_Participant | Nom | Prenom | Statut | Date_Premier_Contact |
'   Nom_Entreprise | Commune | Code_Postal | Mail | Telephone | Activite
'
' Pour importer :
' 1. Collez vos données dans la feuille IMPORT (sans la ligne d'en-tête)
' 2. Cliquez sur le bouton "Importer les données"
' 3. Les participants seront ajoutés à la feuille PARTICIPANTS
'    en conservant leurs IDs existants
'
' ATTENTION : L'import ne duplique pas les participants déjà existants
' (vérification par ID). Assurez-vous que les IDs dans IMPORT sont uniques.
' =============================================================================

' -----------------------------------------------------------------------------
' ImporterDonnees : Importe les données de la feuille IMPORT vers PARTICIPANTS
' Cette macro est assignée au bouton "Importer les données" de la feuille IMPORT
'
' TODO : Compléter cette macro pour gérer les cas suivants :
'   - Validation des données avant import (format de date, statut valide, etc.)
'   - Gestion des conflits d'ID (doublon avec participant existant)
'   - Rapport d'import (nombre de lignes importées, ignorées, en erreur)
'   - Option de mise à jour des participants existants (par opposition à ignorer)
' -----------------------------------------------------------------------------
Public Sub ImporterDonnees()
    ' =========================================================================
    ' MACRO PRÉPARÉE — À COMPLÉTER
    ' =========================================================================
    '
    ' Le code ci-dessous est une ébauche fonctionnelle de base.
    ' Pour une utilisation en production, il est recommandé d'ajouter :
    '   - Une validation complète des données
    '   - Une gestion des erreurs ligne par ligne
    '   - Un rapport d'import détaillé
    ' =========================================================================
    
    Dim wsImport As Worksheet
    Dim wsParticipants As Worksheet
    Dim tblImport As ListObject
    Dim tblParticipants As ListObject
    Dim ligneImport As ListRow
    Dim nouvelleLigne As ListRow
    Dim nbImportes As Long
    Dim nbIgnores As Long
    Dim idImport As Long
    
    nbImportes = 0
    nbIgnores = 0
    
    ' Confirmation avant l'import
    Dim reponse As Integer
    reponse = MsgBox("Voulez-vous importer les données de la feuille IMPORT vers PARTICIPANTS ?" & vbCrLf & _
                     "Les participants avec un ID déjà existant seront ignorés.", _
                     vbYesNo + vbQuestion, "Confirmation d'import")
    
    If reponse <> vbYes Then Exit Sub
    
    ' Accès aux feuilles
    On Error GoTo ErrImport
    Set wsImport = ThisWorkbook.Sheets("IMPORT")
    Set wsParticipants = ThisWorkbook.Sheets("PARTICIPANTS")
    Set tblImport = wsImport.ListObjects("TblImport")
    Set tblParticipants = wsParticipants.ListObjects("TblParticipants")
    On Error GoTo 0
    
    ' Vérifier qu'il y a des données à importer
    If tblImport.DataBodyRange Is Nothing Then
        MsgBox "La feuille IMPORT ne contient aucune donnée à importer.", _
               vbInformation, "Aucune donnée"
        Exit Sub
    End If
    
    ' Désprotéger la feuille PARTICIPANTS
    wsParticipants.Unprotect Password:=MOT_DE_PASSE
    
    ' Parcourir chaque ligne à importer
    For Each ligneImport In tblImport.ListRows
        
        ' Récupérer l'ID de la ligne à importer
        If Not IsNumeric(ligneImport.Range.Cells(1, 1).Value) Then
            nbIgnores = nbIgnores + 1
            GoTo LigneSuivante
        End If
        
        idImport = CLng(ligneImport.Range.Cells(1, 1).Value)
        
        ' Vérifier si cet ID existe déjà dans PARTICIPANTS
        If IDParticipantExiste(tblParticipants, idImport) Then
            nbIgnores = nbIgnores + 1
            GoTo LigneSuivante
        End If
        
        ' Ajouter la ligne dans PARTICIPANTS
        Set nouvelleLigne = tblParticipants.ListRows.Add
        
        Dim j As Integer
        For j = 1 To 11
            nouvelleLigne.Range.Cells(1, j).Value = ligneImport.Range.Cells(1, j).Value
        Next j
        
        ' Reformater la date si nécessaire
        If nouvelleLigne.Range.Cells(1, 5).Value <> "" Then
            nouvelleLigne.Range.Cells(1, 5).NumberFormat = "DD/MM/YYYY"
        End If
        
        nbImportes = nbImportes + 1
        
LigneSuivante:
    Next ligneImport
    
    ' Reprotéger la feuille PARTICIPANTS
    wsParticipants.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True
    
    ' Rapport d'import
    MsgBox "Import terminé !" & vbCrLf & vbCrLf & _
           "Participants importés : " & nbImportes & vbCrLf & _
           "Lignes ignorées (ID existant ou invalide) : " & nbIgnores, _
           vbInformation, "Import terminé"
    
    Exit Sub
    
ErrImport:
    wsParticipants.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True
    MsgBox "Erreur lors de l'import : " & Err.Description & vbCrLf & _
           "Vérifiez que les feuilles IMPORT et PARTICIPANTS existent et contiennent les bons tableaux.", _
           vbCritical, "Erreur d'import"
End Sub

' -----------------------------------------------------------------------------
' IDParticipantExiste : Vérifie si un ID participant existe déjà dans le tableau
' Paramètres :
'   tblParticipants : Le tableau des participants
'   id              : L'ID à vérifier
' Retourne True si l'ID existe déjà
' -----------------------------------------------------------------------------
Private Function IDParticipantExiste(tblParticipants As ListObject, id As Long) As Boolean
    Dim ligne As ListRow
    IDParticipantExiste = False
    
    If tblParticipants.DataBodyRange Is Nothing Then Exit Function
    
    For Each ligne In tblParticipants.ListRows
        If IsNumeric(ligne.Range.Cells(1, 1).Value) Then
            If CLng(ligne.Range.Cells(1, 1).Value) = id Then
                IDParticipantExiste = True
                Exit Function
            End If
        End If
    Next ligne
End Function
