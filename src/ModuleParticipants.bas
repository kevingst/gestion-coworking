Attribute VB_Name = "ModuleParticipants"
' =============================================================================
' Module : ModuleParticipants
' Description : Gestion des participants — création et modification
' =============================================================================

' Statuts disponibles pour les participants
Public Const STATUTS_PARTICIPANTS As String = "Projet pro,Lancé"

' -----------------------------------------------------------------------------
' EnregistrerParticipant : Enregistre un nouveau participant dans PARTICIPANTS
' Paramètres :
'   nom           : Nom du participant (obligatoire)
'   prenom        : Prénom du participant
'   statut        : Statut ("Projet pro" ou "Lancé")
'   dateContact   : Date du premier contact au format JJ/MM/AAAA
'   nomEntreprise : Nom de l'entreprise
'   commune       : Commune
'   codePostal    : Code postal
'   mail          : Adresse mail
'   telephone     : Numéro de téléphone
'   activite      : Description de l'activité (texte libre)
' Retourne True si succès, False si erreur
' -----------------------------------------------------------------------------
Public Function EnregistrerParticipant(nom As String, prenom As String, _
                                       statut As String, dateContact As String, _
                                       nomEntreprise As String, commune As String, _
                                       codePostal As String, mail As String, _
                                       telephone As String, activite As String) As Boolean
    Dim wsParticipants As Worksheet
    Dim tblParticipants As ListObject
    Dim nouvelleDate As Date
    Dim nouvelID As Long
    Dim nouvelleLigne As ListRow
    
    EnregistrerParticipant = False
    
    ' Validation du champ obligatoire
    If Trim(nom) = "" Then
        MsgBox "Le nom du participant est obligatoire.", vbExclamation, "Champ obligatoire"
        Exit Function
    End If
    
    ' Validation et conversion de la date si fournie
    If Trim(dateContact) <> "" Then
        On Error GoTo ErrDate
        nouvelleDate = CDate(dateContact)
        On Error GoTo 0
    End If
    
    ' Accès à la feuille et au tableau
    On Error GoTo ErrFeuille
    Set wsParticipants = ThisWorkbook.Sheets("PARTICIPANTS")
    Set tblParticipants = wsParticipants.ListObjects("TblParticipants")
    On Error GoTo 0
    
    ' Désprotéger la feuille pour écrire
    wsParticipants.Unprotect Password:=MOT_DE_PASSE
    
    ' Calcul du prochain ID (en continuité avec les IDs existants)
    nouvelID = 1
    If Not tblParticipants.DataBodyRange Is Nothing Then
        Dim ligne As ListRow
        For Each ligne In tblParticipants.ListRows
            If IsNumeric(ligne.Range.Cells(1, 1).Value) Then
                If CLng(ligne.Range.Cells(1, 1).Value) >= nouvelID Then
                    nouvelID = CLng(ligne.Range.Cells(1, 1).Value) + 1
                End If
            End If
        Next ligne
    End If
    
    ' Ajout d'une nouvelle ligne dans le tableau
    Set nouvelleLigne = tblParticipants.ListRows.Add
    
    With nouvelleLigne.Range
        .Cells(1, 1).Value = nouvelID              ' ID_Participant
        .Cells(1, 2).Value = Trim(nom)             ' Nom
        .Cells(1, 3).Value = Trim(prenom)          ' Prenom
        .Cells(1, 4).Value = Trim(statut)          ' Statut
        
        ' Date_Premier_Contact
        If Trim(dateContact) <> "" Then
            .Cells(1, 5).Value = nouvelleDate
            .Cells(1, 5).NumberFormat = "DD/MM/YYYY"
        End If
        
        .Cells(1, 6).Value = Trim(nomEntreprise)   ' Nom_Entreprise
        .Cells(1, 7).Value = Trim(commune)         ' Commune
        .Cells(1, 8).Value = Trim(codePostal)      ' Code_Postal
        .Cells(1, 9).Value = Trim(mail)            ' Mail
        .Cells(1, 10).Value = Trim(telephone)      ' Telephone
        .Cells(1, 11).Value = Trim(activite)       ' Activite
    End With
    
    ' Reprotéger la feuille
    wsParticipants.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True
    
    EnregistrerParticipant = True
    Exit Function
    
ErrDate:
    MsgBox "Format de date invalide. Utilisez le format JJ/MM/AAAA." & vbCrLf & _
           "Exemple : 25/03/2025", vbExclamation, "Date invalide"
    wsParticipants.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True
    Exit Function

ErrFeuille:
    MsgBox "Erreur d'accès à la feuille PARTICIPANTS ou au tableau TblParticipants.", _
           vbCritical, "Erreur"
    Exit Function
End Function

' -----------------------------------------------------------------------------
' ModifierParticipant : Modifie les informations d'un participant existant
' Paramètres :
'   idParticipant : L'ID du participant à modifier
'   (autres paramètres : voir EnregistrerParticipant)
' Retourne True si succès, False si erreur
' -----------------------------------------------------------------------------
Public Function ModifierParticipant(idParticipant As Long, nom As String, _
                                    prenom As String, statut As String, _
                                    dateContact As String, nomEntreprise As String, _
                                    commune As String, codePostal As String, _
                                    mail As String, telephone As String, _
                                    activite As String) As Boolean
    Dim wsParticipants As Worksheet
    Dim tblParticipants As ListObject
    Dim ligneParticipant As ListRow
    Dim nouvelleDate As Date
    Dim trouve As Boolean
    
    ModifierParticipant = False
    
    ' Validation du champ obligatoire
    If Trim(nom) = "" Then
        MsgBox "Le nom du participant est obligatoire.", vbExclamation, "Champ obligatoire"
        Exit Function
    End If
    
    ' Validation et conversion de la date si fournie
    If Trim(dateContact) <> "" Then
        On Error GoTo ErrDate
        nouvelleDate = CDate(dateContact)
        On Error GoTo 0
    End If
    
    ' Accès à la feuille et au tableau
    On Error GoTo ErrFeuille
    Set wsParticipants = ThisWorkbook.Sheets("PARTICIPANTS")
    Set tblParticipants = wsParticipants.ListObjects("TblParticipants")
    On Error GoTo 0
    
    ' Désprotéger la feuille
    wsParticipants.Unprotect Password:=MOT_DE_PASSE
    
    ' Rechercher le participant par son ID
    trouve = False
    If Not tblParticipants.DataBodyRange Is Nothing Then
        For Each ligneParticipant In tblParticipants.ListRows
            If IsNumeric(ligneParticipant.Range.Cells(1, 1).Value) Then
                If CLng(ligneParticipant.Range.Cells(1, 1).Value) = idParticipant Then
                    ' Mettre à jour les informations
                    With ligneParticipant.Range
                        .Cells(1, 2).Value = Trim(nom)
                        .Cells(1, 3).Value = Trim(prenom)
                        .Cells(1, 4).Value = Trim(statut)
                        
                        If Trim(dateContact) <> "" Then
                            .Cells(1, 5).Value = nouvelleDate
                            .Cells(1, 5).NumberFormat = "DD/MM/YYYY"
                        End If
                        
                        .Cells(1, 6).Value = Trim(nomEntreprise)
                        .Cells(1, 7).Value = Trim(commune)
                        .Cells(1, 8).Value = Trim(codePostal)
                        .Cells(1, 9).Value = Trim(mail)
                        .Cells(1, 10).Value = Trim(telephone)
                        .Cells(1, 11).Value = Trim(activite)
                    End With
                    
                    trouve = True
                    Exit For
                End If
            End If
        Next ligneParticipant
    End If
    
    ' Reprotéger la feuille
    wsParticipants.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True
    
    If Not trouve Then
        MsgBox "Participant avec l'ID " & idParticipant & " introuvable.", _
               vbExclamation, "Participant non trouvé"
        Exit Function
    End If
    
    ' Mettre à jour aussi les informations dénormalisées dans PRESENCES
    Call MettreAJourStatutPresences(idParticipant, Trim(statut))
    
    ModifierParticipant = True
    Exit Function
    
ErrDate:
    MsgBox "Format de date invalide. Utilisez le format JJ/MM/AAAA." & vbCrLf & _
           "Exemple : 25/03/2025", vbExclamation, "Date invalide"
    wsParticipants.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True
    Exit Function

ErrFeuille:
    MsgBox "Erreur d'accès à la feuille PARTICIPANTS ou au tableau TblParticipants.", _
           vbCritical, "Erreur"
    Exit Function
End Function

' -----------------------------------------------------------------------------
' MettreAJourStatutPresences : Met à jour le statut dans PRESENCES quand
' le statut d'un participant change (pour la cohérence des données)
' Paramètres :
'   idParticipant : L'ID du participant
'   nouveauStatut : Le nouveau statut
' -----------------------------------------------------------------------------
Private Sub MettreAJourStatutPresences(idParticipant As Long, nouveauStatut As String)
    Dim wsPresences As Worksheet
    Dim tblPresences As ListObject
    Dim lignePresence As ListRow
    
    On Error GoTo ErrPresences
    Set wsPresences = ThisWorkbook.Sheets("PRESENCES")
    Set tblPresences = wsPresences.ListObjects("TblPresences")
    
    wsPresences.Unprotect Password:=MOT_DE_PASSE
    
    ' Mettre à jour toutes les lignes de présence de ce participant
    If Not tblPresences.DataBodyRange Is Nothing Then
        For Each lignePresence In tblPresences.ListRows
            ' Colonne 3 = ID_Participant
            If IsNumeric(lignePresence.Range.Cells(1, 3).Value) Then
                If CLng(lignePresence.Range.Cells(1, 3).Value) = idParticipant Then
                    lignePresence.Range.Cells(1, 6).Value = nouveauStatut ' Statut_Participant
                End If
            End If
        Next lignePresence
    End If
    
    wsPresences.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True
    
    ' Recalculer toutes les statistiques (le statut a changé)
    Call MettreAJourStats
    
    Exit Sub
    
ErrPresences:
    wsPresences.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------
' ObtenirListeStatuts : Retourne le tableau des statuts disponibles
' -----------------------------------------------------------------------------
Public Function ObtenirListeStatuts() As String()
    ObtenirListeStatuts = Split(STATUTS_PARTICIPANTS, ",")
End Function

' -----------------------------------------------------------------------------
' RechercherParticipants : Recherche des participants par nom ou prénom
' Paramètre :
'   critere : Le texte à rechercher
' Retourne un tableau 2D (ID, Nom, Prenom, Statut) ou un tableau vide
' -----------------------------------------------------------------------------
Public Function RechercherParticipants(critere As String) As Variant
    Dim wsParticipants As Worksheet
    Dim tblParticipants As ListObject
    Dim ligneParticipant As ListRow
    Dim resultats() As Variant
    Dim nb As Integer
    
    nb = 0
    ReDim resultats(0 To 0, 0 To 3)
    
    On Error GoTo ErrRecherche
    Set wsParticipants = ThisWorkbook.Sheets("PARTICIPANTS")
    Set tblParticipants = wsParticipants.ListObjects("TblParticipants")
    On Error GoTo 0
    
    If tblParticipants.DataBodyRange Is Nothing Then
        RechercherParticipants = resultats
        Exit Function
    End If
    
    ' Compter les résultats d'abord
    For Each ligneParticipant In tblParticipants.ListRows
        Dim nomP As String
        Dim prenomP As String
        nomP = LCase(CStr(ligneParticipant.Range.Cells(1, 2).Value))
        prenomP = LCase(CStr(ligneParticipant.Range.Cells(1, 3).Value))
        
        If Trim(critere) = "" Or _
           InStr(nomP, LCase(Trim(critere))) > 0 Or _
           InStr(prenomP, LCase(Trim(critere))) > 0 Then
            nb = nb + 1
        End If
    Next ligneParticipant
    
    If nb = 0 Then
        RechercherParticipants = resultats
        Exit Function
    End If
    
    ' Remplir le tableau de résultats
    ReDim resultats(0 To nb - 1, 0 To 3)
    Dim i As Integer
    i = 0
    
    For Each ligneParticipant In tblParticipants.ListRows
        nomP = LCase(CStr(ligneParticipant.Range.Cells(1, 2).Value))
        prenomP = LCase(CStr(ligneParticipant.Range.Cells(1, 3).Value))
        
        If Trim(critere) = "" Or _
           InStr(nomP, LCase(Trim(critere))) > 0 Or _
           InStr(prenomP, LCase(Trim(critere))) > 0 Then
            resultats(i, 0) = ligneParticipant.Range.Cells(1, 1).Value  ' ID
            resultats(i, 1) = ligneParticipant.Range.Cells(1, 2).Value  ' Nom
            resultats(i, 2) = ligneParticipant.Range.Cells(1, 3).Value  ' Prenom
            resultats(i, 3) = ligneParticipant.Range.Cells(1, 4).Value  ' Statut
            i = i + 1
        End If
    Next ligneParticipant
    
    RechercherParticipants = resultats
    Exit Function
    
ErrRecherche:
    RechercherParticipants = resultats
End Function
