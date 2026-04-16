' =============================================================================
' Module : ModuleImport
' Description : Import de donnÃ©es existantes depuis la feuille IMPORT
'
' USAGE FUTUR :
' Ce module est prÃ©parÃ© pour permettre l'import de participants depuis un
' fichier Excel existant. La feuille IMPORT doit contenir les donnÃ©es Ã 
' importer avec les mÃªmes colonnes que PARTICIPANTS :
'   ID_Participant | Nom | Prenom | Statut | Date_Premier_Contact |
'   Nom_Entreprise | Commune | Code_Postal | Mail | Telephone | Activite
'
' Pour importer :
' 1. Collez vos donnÃ©es dans la feuille IMPORT (sans la ligne d'en-tÃªte)
' 2. Cliquez sur le bouton "Importer les donnÃ©es"
' 3. Les participants seront ajoutÃ©s Ã  la feuille PARTICIPANTS
'    en conservant leurs IDs existants
'
' ATTENTION : L'import ne duplique pas les participants dÃ©jÃ  existants
' (vÃ©rification par ID). Assurez-vous que les IDs dans IMPORT sont uniques.
' =============================================================================

' -----------------------------------------------------------------------------
' ImporterDonnees : Importe les donnÃ©es de la feuille IMPORT vers PARTICIPANTS
' Cette macro est assignÃ©e au bouton "Importer les donnÃ©es" de la feuille IMPORT
'
' TODO : ComplÃ©ter cette macro pour gÃ©rer les cas suivants :
'   - Validation des donnÃ©es avant import (format de date, statut valide, etc.)
'   - Gestion des conflits d'ID (doublon avec participant existant)
'   - Rapport d'import (nombre de lignes importÃ©es, ignorÃ©es, en erreur)
'   - Option de mise Ã  jour des participants existants (par opposition Ã  ignorer)
' -----------------------------------------------------------------------------
Public Sub ImporterDonnees()
    ' =========================================================================
    ' MACRO PRÃ‰PARÃ‰E â€” Ã€ COMPLÃ‰TER
    ' =========================================================================
    '
    ' Le code ci-dessous est une Ã©bauche fonctionnelle de base.
    ' Pour une utilisation en production, il est recommandÃ© d'ajouter :
    '   - Une validation complÃ¨te des donnÃ©es
    '   - Une gestion des erreurs ligne par ligne
    '   - Un rapport d'import dÃ©taillÃ©
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
    reponse = MsgBox("Voulez-vous importer les donnÃ©es de la feuille IMPORT vers PARTICIPANTS ?" & vbCrLf & _
                     "Les participants avec un ID dÃ©jÃ  existant seront ignorÃ©s.", _
                     vbYesNo + vbQuestion, "Confirmation d'import")
    
    If reponse <> vbYes Then Exit Sub
    
    ' AccÃ¨s aux feuilles
    On Error GoTo ErrImport
    Set wsImport = ThisWorkbook.Sheets("IMPORT")
    Set wsParticipants = ThisWorkbook.Sheets("PARTICIPANTS")
    Set tblImport = wsImport.ListObjects("TblImport")
    Set tblParticipants = wsParticipants.ListObjects("TblParticipants")
    On Error GoTo 0
    
    ' VÃ©rifier qu'il y a des donnÃ©es Ã  importer
    If tblImport.DataBodyRange Is Nothing Then
        MsgBox "La feuille IMPORT ne contient aucune donnÃ©e Ã  importer.", _
               vbInformation, "Aucune donnÃ©e"
        Exit Sub
    End If
    
    ' DÃ©sprotÃ©ger la feuille PARTICIPANTS
    wsParticipants.Unprotect Password:=MOT_DE_PASSE
    
    ' Parcourir chaque ligne Ã  importer
    For Each ligneImport In tblImport.ListRows
        
        ' RÃ©cupÃ©rer l'ID de la ligne Ã  importer
        If Not IsNumeric(ligneImport.Range.Cells(1, 1).Value) Then
            nbIgnores = nbIgnores + 1
            GoTo LigneSuivante
        End If
        
        idImport = CLng(ligneImport.Range.Cells(1, 1).Value)
        
        ' VÃ©rifier si cet ID existe dÃ©jÃ  dans PARTICIPANTS
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
        
        ' Reformater la date si nÃ©cessaire
        If nouvelleLigne.Range.Cells(1, 5).Value <> "" Then
            nouvelleLigne.Range.Cells(1, 5).NumberFormat = "DD/MM/YYYY"
        End If
        
        nbImportes = nbImportes + 1
        
LigneSuivante:
    Next ligneImport
    
    ' ReprotÃ©ger la feuille PARTICIPANTS
    wsParticipants.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True
    
    ' Rapport d'import
    MsgBox "Import terminÃ© !" & vbCrLf & vbCrLf & _
           "Participants importÃ©s : " & nbImportes & vbCrLf & _
           "Lignes ignorÃ©es (ID existant ou invalide) : " & nbIgnores, _
           vbInformation, "Import terminÃ©"
    
    Exit Sub
    
ErrImport:
    wsParticipants.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True
    MsgBox "Erreur lors de l'import : " & Err.Description & vbCrLf & _
           "VÃ©rifiez que les feuilles IMPORT et PARTICIPANTS existent et contiennent les bons tableaux.", _
           vbCritical, "Erreur d'import"
End Sub

' -----------------------------------------------------------------------------
' IDParticipantExiste : VÃ©rifie si un ID participant existe dÃ©jÃ  dans le tableau
' ParamÃ¨tres :
'   tblParticipants : Le tableau des participants
'   id              : L'ID Ã  vÃ©rifier
' Retourne True si l'ID existe dÃ©jÃ 
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
