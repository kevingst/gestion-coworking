' =============================================================================
' UserForm : FrmGererAteliers
' Description : Formulaire de recherche, modification et suppression d'ateliers,
'               et gestion des présences d'un atelier sélectionné
'
' Contrôles requis (à créer dans l'éditeur VBA) :
'   Zone de recherche / liste des ateliers :
'   - TxtRecherche        : TextBox — Filtre de recherche par nom d'atelier
'   - LstAteliers         : ListBox — Liste des ateliers (ColumnCount=4 : ID, Nom, Date, Thème)
'
'   Zone de détail de l'atelier (Enabled=False par défaut) :
'   - TxtNom              : TextBox — Nom de l'atelier (éditable)
'   - TxtDate             : TextBox — Date (JJ/MM/AAAA, éditable)
'   - TxtHeureDebut       : TextBox — Heure début (HH:MM, éditable)
'   - TxtHeureFin         : TextBox — Heure fin (HH:MM, éditable)
'   - TxtDuree            : TextBox — Durée calculée (lecture seule)
'   - CboTheme            : ComboBox — Thème de l'atelier (éditable)
'   - TxtNbParticipants   : TextBox — Nombre de participants (lecture seule)
'   - TxtNbParticipantsPro: TextBox — Nombre de participants pro (lecture seule)
'   - TxtAnimePar        : TextBox — Animé par (saisie libre)
'
'   Zone des présences :
'   - LstPresences        : ListBox — Participants présents (ColumnCount=5 : ID, Nom, Prénom, Statut, Mail)
'
'   Boutons d'action :
'   - BtnSauvegarder      : CommandButton — Sauvegarder les modifications de l'atelier
'   - BtnSupprimerAtelier : CommandButton — Supprimer l'atelier entier
'   - BtnSupprimerPresence: CommandButton — Supprimer la présence sélectionnée
'   - BtnFermer           : CommandButton — Fermer le formulaire
'
' Voir formulaires.md pour les propriétés détaillées de chaque contrôle.
' =============================================================================

' Variable pour stocker l'ID de l'atelier en cours de modification
Private idAtelierSelectionne As Long

' -----------------------------------------------------------------------------
' UserForm_Initialize : Initialisation du formulaire à l'ouverture
' -----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    idAtelierSelectionne = 0

    ' Configurer la ListBox des ateliers
    LstAteliers.ColumnCount = 4
    LstAteliers.ColumnWidths = "40;200;80;120"

    ' Configurer la ListBox des présences
    LstPresences.ColumnCount = 5
    LstPresences.ColumnWidths = "40;150;120;100;180"

    ' Remplir le ComboBox des thèmes
    Dim themes As Variant
    Dim theme As Variant
    themes = ObtenirListeThemes()
    CboTheme.Clear
    For Each theme In themes
        CboTheme.AddItem theme
    Next theme

    ' Désactiver les champs de détail jusqu'à sélection d'un atelier
    Call DefinirEtatChamps(False)

    ' Charger tous les ateliers (tri par date décroissante)
    Call ChargerAteliers("")
End Sub

' -----------------------------------------------------------------------------
' TxtRecherche_Change : Filtrage en temps réel de la liste des ateliers
' -----------------------------------------------------------------------------
Private Sub TxtRecherche_Change()
    Call ChargerAteliers(TxtRecherche.Value)
End Sub

' -----------------------------------------------------------------------------
' LstAteliers_Click : Sélection d'un atelier dans la liste
' -----------------------------------------------------------------------------
Private Sub LstAteliers_Click()
    If LstAteliers.ListIndex < 0 Then Exit Sub

    On Error Resume Next
    idAtelierSelectionne = CLng(LstAteliers.List(LstAteliers.ListIndex, 0))
    On Error GoTo 0

    If idAtelierSelectionne <= 0 Then Exit Sub

    ' Charger les détails de l'atelier dans les champs
    Call ChargerDetailsAtelier(idAtelierSelectionne)

    ' Charger les présences de l'atelier
    Call ChargerPresencesAtelier(idAtelierSelectionne)

    ' Activer les champs et boutons
    Call DefinirEtatChamps(True)
End Sub

' -----------------------------------------------------------------------------
' BtnSauvegarder_Click : Sauvegarde les modifications de l'atelier
' -----------------------------------------------------------------------------
Private Sub BtnSauvegarder_Click()
    If idAtelierSelectionne <= 0 Then
        MsgBox "Aucun atelier sélectionné.", vbExclamation, "Erreur"
        Exit Sub
    End If

    ' Validation des champs obligatoires
    If Trim(TxtNom.Value) = "" Then
        MsgBox "Le nom de l'atelier est obligatoire.", vbExclamation, "Champ obligatoire"
        TxtNom.SetFocus
        Exit Sub
    End If

    If Trim(TxtDate.Value) = "" Then
        MsgBox "La date de l'atelier est obligatoire.", vbExclamation, "Champ obligatoire"
        TxtDate.SetFocus
        Exit Sub
    End If

    ' Validation du format de date
    Dim nouvelleDate As Date
    On Error GoTo ErrDate
    nouvelleDate = CDate(TxtDate.Value)
    On Error GoTo 0

    ' Validation des heures (si renseignées)
    Dim heureDebutVal As Date
    Dim heureFinVal As Date
    Dim dureeMinutes As Long
    Dim dureeFormatee As String

    If Trim(TxtHeureDebut.Value) <> "" And Trim(TxtHeureFin.Value) <> "" Then
        On Error GoTo ErrHeure
        heureDebutVal = CDate(TxtHeureDebut.Value)
        heureFinVal = CDate(TxtHeureFin.Value)
        On Error GoTo 0

        If heureFinVal <= heureDebutVal Then
            MsgBox "L'heure de fin doit être postérieure à l'heure de début.", vbExclamation, "Heure invalide"
            TxtHeureFin.SetFocus
            Exit Sub
        End If

        dureeMinutes = DateDiff("n", heureDebutVal, heureFinVal)
        dureeFormatee = Format(dureeMinutes \ 60, "00") & ":" & Format(dureeMinutes Mod 60, "00")
    Else
        dureeFormatee = "00:00"
    End If

    ' Mise à jour dans la feuille ATELIERS
    Dim wsAteliers As Worksheet
    Dim tblAteliers As ListObject
    Dim ligneAtelier As ListRow

    On Error GoTo ErrFeuille
    Set wsAteliers = ThisWorkbook.Sheets("ATELIERS")
    Set tblAteliers = wsAteliers.ListObjects("TblAteliers")
    On Error GoTo 0

    wsAteliers.Unprotect Password:=MOT_DE_PASSE

    Dim trouve As Boolean
    trouve = False

    If Not tblAteliers.DataBodyRange Is Nothing Then
        For Each ligneAtelier In tblAteliers.ListRows
            If IsNumeric(ligneAtelier.Range.Cells(1, 1).Value) Then
                If CLng(ligneAtelier.Range.Cells(1, 1).Value) = idAtelierSelectionne Then
                    ligneAtelier.Range.Cells(1, 2).Value = Trim(TxtNom.Value)          ' Nom
                    ligneAtelier.Range.Cells(1, 3).Value = nouvelleDate                ' Date
                    ligneAtelier.Range.Cells(1, 3).NumberFormat = "DD/MM/YYYY"
                    ligneAtelier.Range.Cells(1, 4).Value = TxtHeureDebut.Value         ' Heure_Debut
                    ligneAtelier.Range.Cells(1, 4).NumberFormat = "HH:MM"
                    ligneAtelier.Range.Cells(1, 5).Value = TxtHeureFin.Value           ' Heure_Fin
                    ligneAtelier.Range.Cells(1, 5).NumberFormat = "HH:MM"
                    ligneAtelier.Range.Cells(1, 6).Value = dureeFormatee               ' Duree
                    ligneAtelier.Range.Cells(1, 6).NumberFormat = "@"  ' Texte pour forcer le stockage en HH:MM texte
                    ligneAtelier.Range.Cells(1, 7).Value = CboTheme.Value              ' Theme
                    ligneAtelier.Range.Cells(1, 10).Value = Trim(TxtAnimePar.Value) ' Anime_Par
                    trouve = True
                    Exit For
                End If
            End If
        Next ligneAtelier
    End If

    wsAteliers.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True

    If Not trouve Then
        MsgBox "Atelier introuvable dans la feuille.", vbCritical, "Erreur"
        Exit Sub
    End If

    ' Mettre à jour la durée affichée
    TxtDuree.Value = dureeFormatee

    ' Mettre à jour les statistiques
    Call MettreAJourStats

    ' Rafraîchir la liste des ateliers
    Call ChargerAteliers(TxtRecherche.Value)

    MsgBox "Les modifications ont été sauvegardées avec succès !", vbInformation, "Succès"
    Exit Sub

ErrDate:
    MsgBox "Format de date invalide. Utilisez le format JJ/MM/AAAA." & vbCrLf & _
           "Exemple : 25/03/2025", vbExclamation, "Date invalide"
    TxtDate.SetFocus
    Exit Sub

ErrHeure:
    MsgBox "Format d'heure invalide. Utilisez le format HH:MM." & vbCrLf & _
           "Exemple : 09:30", vbExclamation, "Heure invalide"
    TxtHeureDebut.SetFocus
    Exit Sub

ErrFeuille:
    MsgBox "Erreur d'accès à la feuille ATELIERS ou au tableau TblAteliers.", _
           vbCritical, "Erreur"
    Exit Sub
End Sub

' -----------------------------------------------------------------------------
' BtnSupprimerAtelier_Click : Supprime l'atelier sélectionné et ses présences
' -----------------------------------------------------------------------------
Private Sub BtnSupprimerAtelier_Click()
    If idAtelierSelectionne <= 0 Then
        MsgBox "Aucun atelier sélectionné.", vbExclamation, "Erreur"
        Exit Sub
    End If

    ' Demander confirmation
    Dim reponse As Integer
    reponse = MsgBox("Êtes-vous sûr de vouloir supprimer cet atelier ?" & vbCrLf & _
                     "Toutes les présences associées seront également supprimées." & vbCrLf & vbCrLf & _
                     "Cette action est irréversible.", _
                     vbYesNo + vbCritical, "Confirmation de suppression")

    If reponse <> vbYes Then Exit Sub

    Dim wsAteliers As Worksheet
    Dim wsPresences As Worksheet
    Dim tblAteliers As ListObject
    Dim tblPresences As ListObject
    Dim ligneAtelier As ListRow
    Dim lignePresence As ListRow
    Dim idAtelier As Long

    idAtelier = idAtelierSelectionne

    On Error GoTo ErrSuppression
    Set wsAteliers = ThisWorkbook.Sheets("ATELIERS")
    Set wsPresences = ThisWorkbook.Sheets("PRESENCES")
    Set tblAteliers = wsAteliers.ListObjects("TblAteliers")
    Set tblPresences = wsPresences.ListObjects("TblPresences")
    On Error GoTo 0

    ' Supprimer toutes les présences de cet atelier (parcourir à rebours)
    wsPresences.Unprotect Password:=MOT_DE_PASSE

    If Not tblPresences.DataBodyRange Is Nothing Then
        Dim i As Long
        For i = tblPresences.ListRows.Count To 1 Step -1
            Set lignePresence = tblPresences.ListRows(i)
            If IsNumeric(lignePresence.Range.Cells(1, 2).Value) Then
                If CLng(lignePresence.Range.Cells(1, 2).Value) = idAtelier Then
                    lignePresence.Delete
                End If
            End If
        Next i
    End If

    wsPresences.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True

    ' Supprimer la ligne de l'atelier dans TblAteliers
    wsAteliers.Unprotect Password:=MOT_DE_PASSE

    If Not tblAteliers.DataBodyRange Is Nothing Then
        For Each ligneAtelier In tblAteliers.ListRows
            If IsNumeric(ligneAtelier.Range.Cells(1, 1).Value) Then
                If CLng(ligneAtelier.Range.Cells(1, 1).Value) = idAtelier Then
                    ligneAtelier.Delete
                    Exit For
                End If
            End If
        Next ligneAtelier
    End If

    wsAteliers.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True

    ' Mettre à jour les statistiques
    Call MettreAJourStats

    ' Réinitialiser la sélection et les champs
    idAtelierSelectionne = 0
    Call ViderChampsDetail
    Call DefinirEtatChamps(False)
    LstPresences.Clear

    ' Rafraîchir la liste des ateliers
    Call ChargerAteliers(TxtRecherche.Value)

    MsgBox "L'atelier a été supprimé avec succès.", vbInformation, "Suppression effectuée"
    Exit Sub

ErrSuppression:
    On Error Resume Next
    wsPresences.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True
    wsAteliers.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True
    On Error GoTo 0
    MsgBox "Erreur lors de la suppression de l'atelier.", vbCritical, "Erreur"
End Sub

' -----------------------------------------------------------------------------
' BtnSupprimerPresence_Click : Supprime la présence sélectionnée dans LstPresences
' -----------------------------------------------------------------------------
Private Sub BtnSupprimerPresence_Click()
    If idAtelierSelectionne <= 0 Then
        MsgBox "Aucun atelier sélectionné.", vbExclamation, "Erreur"
        Exit Sub
    End If

    If LstPresences.ListIndex < 0 Then
        MsgBox "Veuillez sélectionner un participant dans la liste des présences.", _
               vbExclamation, "Sélection manquante"
        Exit Sub
    End If

    ' Récupérer l'ID du participant sélectionné
    Dim idPresenceParticipant As Long
    On Error Resume Next
    idPresenceParticipant = CLng(LstPresences.List(LstPresences.ListIndex, 0))
    On Error GoTo 0

    If idPresenceParticipant <= 0 Then
        MsgBox "Impossible d'identifier le participant sélectionné.", vbExclamation, "Erreur"
        Exit Sub
    End If

    ' Récupérer le nom pour le message de confirmation
    Dim nomAffiche As String
    nomAffiche = LstPresences.List(LstPresences.ListIndex, 1) & " " & _
                 LstPresences.List(LstPresences.ListIndex, 2)

    ' Demander confirmation
    Dim reponse As Integer
    reponse = MsgBox("Êtes-vous sûr de vouloir supprimer la présence de " & Trim(nomAffiche) & " ?" & vbCrLf & _
                     "Cette action est irréversible.", _
                     vbYesNo + vbQuestion, "Confirmation de suppression")

    If reponse <> vbYes Then Exit Sub

    ' Supprimer la ligne correspondante dans TblPresences
    Dim wsPresences As Worksheet
    Dim tblPresences As ListObject
    Dim lignePresence As ListRow
    Dim idAtelier As Long

    idAtelier = idAtelierSelectionne

    On Error GoTo ErrSuppPresence
    Set wsPresences = ThisWorkbook.Sheets("PRESENCES")
    Set tblPresences = wsPresences.ListObjects("TblPresences")
    On Error GoTo 0

    wsPresences.Unprotect Password:=MOT_DE_PASSE

    Dim j As Long
    For j = tblPresences.ListRows.Count To 1 Step -1
        Set lignePresence = tblPresences.ListRows(j)
        If IsNumeric(lignePresence.Range.Cells(1, 2).Value) And _
           IsNumeric(lignePresence.Range.Cells(1, 3).Value) Then
            If CLng(lignePresence.Range.Cells(1, 2).Value) = idAtelier And _
               CLng(lignePresence.Range.Cells(1, 3).Value) = idPresenceParticipant Then
                lignePresence.Delete
                Exit For
            End If
        End If
    Next j

    wsPresences.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True

    ' Recalculer les compteurs de l'atelier
    Call RecalculerNbParticipants(idAtelier)

    ' Mettre à jour les statistiques
    Call MettreAJourStats

    ' Rafraîchir la liste des présences et les compteurs
    Call ChargerPresencesAtelier(idAtelier)
    Call MettreAJourCompteurs(idAtelier)

    MsgBox "La présence a été supprimée avec succès.", vbInformation, "Suppression effectuée"
    Exit Sub

ErrSuppPresence:
    On Error Resume Next
    wsPresences.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True
    On Error GoTo 0
    MsgBox "Erreur lors de la suppression de la présence.", vbCritical, "Erreur"
End Sub

' -----------------------------------------------------------------------------
' BtnFermer_Click : Fermeture du formulaire
' -----------------------------------------------------------------------------
Private Sub BtnFermer_Click()
    Unload Me
End Sub

' =============================================================================
' PROCÉDURES INTERNES
' =============================================================================

' -----------------------------------------------------------------------------
' ChargerAteliers : Charge les ateliers dans LstAteliers (tri par date décroissante)
' Paramètre :
'   filtre : Texte de filtre sur le nom (vide = tous)
' -----------------------------------------------------------------------------
Private Sub ChargerAteliers(filtre As String)
    Dim wsAteliers As Worksheet
    Dim tblAteliers As ListObject
    Dim ligneAtelier As ListRow

    LstAteliers.Clear

    On Error GoTo ErrChargement
    Set wsAteliers = ThisWorkbook.Sheets("ATELIERS")
    Set tblAteliers = wsAteliers.ListObjects("TblAteliers")
    On Error GoTo 0

    If tblAteliers.DataBodyRange Is Nothing Then Exit Sub

    ' Collecter les lignes qui correspondent au filtre
    Dim nb As Long
    nb = 0
    For Each ligneAtelier In tblAteliers.ListRows
        Dim nomAtelier As String
        nomAtelier = CStr(ligneAtelier.Range.Cells(1, 2).Value)
        If filtre = "" Or InStr(1, nomAtelier, filtre, vbTextCompare) > 0 Then
            nb = nb + 1
        End If
    Next ligneAtelier

    If nb = 0 Then Exit Sub

    ' Construire un tableau temporaire pour le tri
    Dim donnees() As Variant
    ReDim donnees(0 To nb - 1, 0 To 3)
    Dim dates() As Date
    ReDim dates(0 To nb - 1)

    Dim idx As Long
    idx = 0
    For Each ligneAtelier In tblAteliers.ListRows
        nomAtelier = CStr(ligneAtelier.Range.Cells(1, 2).Value)
        If filtre = "" Or InStr(1, nomAtelier, filtre, vbTextCompare) > 0 Then
            donnees(idx, 0) = ligneAtelier.Range.Cells(1, 1).Value  ' ID
            donnees(idx, 1) = nomAtelier                             ' Nom
            ' Date formatée en JJ/MM/AAAA
            Dim dateVal As String
            dateVal = ""
            On Error Resume Next
            dateVal = Format(CDate(ligneAtelier.Range.Cells(1, 3).Value), "DD/MM/YYYY")
            On Error GoTo 0
            donnees(idx, 2) = dateVal                                ' Date
            donnees(idx, 3) = CStr(ligneAtelier.Range.Cells(1, 7).Value)  ' Thème

            ' Stocker la date pour le tri
            On Error Resume Next
            dates(idx) = CDate(ligneAtelier.Range.Cells(1, 3).Value)
            On Error GoTo 0

            idx = idx + 1
        End If
    Next ligneAtelier

    ' Tri à bulles par date décroissante
    Dim tempDate As Date
    Dim tempDonnees(0 To 3) As Variant
    Dim k As Long, m As Long
    For k = 0 To nb - 2
        For m = 0 To nb - 2 - k
            If dates(m) < dates(m + 1) Then
                ' Échanger les dates
                tempDate = dates(m)
                dates(m) = dates(m + 1)
                dates(m + 1) = tempDate
                ' Échanger les données
                Dim col As Long
                For col = 0 To 3
                    tempDonnees(col) = donnees(m, col)
                    donnees(m, col) = donnees(m + 1, col)
                    donnees(m + 1, col) = tempDonnees(col)
                Next col
            End If
        Next m
    Next k

    ' Remplir la ListBox
    For idx = 0 To nb - 1
        LstAteliers.AddItem donnees(idx, 0)
        LstAteliers.List(LstAteliers.ListCount - 1, 1) = donnees(idx, 1)
        LstAteliers.List(LstAteliers.ListCount - 1, 2) = donnees(idx, 2)
        LstAteliers.List(LstAteliers.ListCount - 1, 3) = donnees(idx, 3)
    Next idx

    Exit Sub

ErrChargement:
    ' Feuille ou tableau indisponible — liste reste vide
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------
' ChargerDetailsAtelier : Charge les infos d'un atelier dans les champs
' Paramètre :
'   idAtelier : L'ID de l'atelier à charger
' -----------------------------------------------------------------------------
Private Sub ChargerDetailsAtelier(idAtelier As Long)
    Dim wsAteliers As Worksheet
    Dim tblAteliers As ListObject
    Dim ligneAtelier As ListRow

    On Error GoTo ErrDetail
    Set wsAteliers = ThisWorkbook.Sheets("ATELIERS")
    Set tblAteliers = wsAteliers.ListObjects("TblAteliers")
    On Error GoTo 0

    If tblAteliers.DataBodyRange Is Nothing Then Exit Sub

    For Each ligneAtelier In tblAteliers.ListRows
        If IsNumeric(ligneAtelier.Range.Cells(1, 1).Value) Then
            If CLng(ligneAtelier.Range.Cells(1, 1).Value) = idAtelier Then
                TxtNom.Value = CStr(ligneAtelier.Range.Cells(1, 2).Value)

                ' Date au format JJ/MM/AAAA
                Dim dateStr As String
                dateStr = ""
                On Error Resume Next
                dateStr = Format(CDate(ligneAtelier.Range.Cells(1, 3).Value), "DD/MM/YYYY")
                On Error GoTo 0
                TxtDate.Value = dateStr

                ' Heure début — conversion du décimal Excel en format HH:MM
                Dim heureDebutStr As String
                heureDebutStr = ""
                On Error Resume Next
                If Not IsEmpty(ligneAtelier.Range.Cells(1, 4).Value) And IsNumeric(ligneAtelier.Range.Cells(1, 4).Value) Then
                    heureDebutStr = Format(CDate(ligneAtelier.Range.Cells(1, 4).Value), "HH:MM")
                ElseIf Not IsEmpty(ligneAtelier.Range.Cells(1, 4).Value) Then
                    heureDebutStr = CStr(ligneAtelier.Range.Cells(1, 4).Value)
                End If
                On Error GoTo 0
                TxtHeureDebut.Value = heureDebutStr

                ' Heure fin — conversion du décimal Excel en format HH:MM
                Dim heureFinStr As String
                heureFinStr = ""
                On Error Resume Next
                If Not IsEmpty(ligneAtelier.Range.Cells(1, 5).Value) And IsNumeric(ligneAtelier.Range.Cells(1, 5).Value) Then
                    heureFinStr = Format(CDate(ligneAtelier.Range.Cells(1, 5).Value), "HH:MM")
                ElseIf Not IsEmpty(ligneAtelier.Range.Cells(1, 5).Value) Then
                    heureFinStr = CStr(ligneAtelier.Range.Cells(1, 5).Value)
                End If
                On Error GoTo 0
                TxtHeureFin.Value = heureFinStr

                ' Durée — conversion du décimal Excel en format HH:MM
                Dim dureeStr As String
                dureeStr = ""
                On Error Resume Next
                Dim dureeRaw As Variant
                dureeRaw = ligneAtelier.Range.Cells(1, 6).Value
                If Not IsEmpty(dureeRaw) And IsNumeric(dureeRaw) Then
                    ' Valeur décimale Excel ? convertir en minutes
                    Dim dureeMin As Long
                    dureeMin = CLng(CDbl(dureeRaw) * 24 * 60)
                    dureeStr = Format(dureeMin \ 60, "00") & ":" & Format(dureeMin Mod 60, "00")
                ElseIf InStr(CStr(dureeRaw), ":") > 0 Then
                    dureeStr = CStr(dureeRaw)
                End If
                On Error GoTo 0
                TxtDuree.Value = dureeStr

                ' Sélectionner le thème dans le ComboBox
                Dim themeVal As String
                themeVal = CStr(ligneAtelier.Range.Cells(1, 7).Value)
                Dim t As Long
                For t = 0 To CboTheme.ListCount - 1
                    If CboTheme.List(t) = themeVal Then
                        CboTheme.ListIndex = t
                        Exit For
                    End If
                Next t

                TxtNbParticipants.Value = CStr(ligneAtelier.Range.Cells(1, 8).Value)
                TxtNbParticipantsPro.Value = CStr(ligneAtelier.Range.Cells(1, 9).Value)
                TxtAnimePar.Value = CStr(ligneAtelier.Range.Cells(1, 10).Value)
                Exit For
            End If
        End If
    Next ligneAtelier

    Exit Sub

ErrDetail:
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------
' ChargerPresencesAtelier : Charge la liste des présences dans LstPresences
' Paramètre :
'   idAtelier : L'ID de l'atelier
' -----------------------------------------------------------------------------
Private Sub ChargerPresencesAtelier(idAtelier As Long)
    Dim wsPresences As Worksheet
    Dim wsParticipants As Worksheet
    Dim tblPresences As ListObject
    Dim tblParticipants As ListObject
    Dim lignePresence As ListRow
    Dim ligneParticipant As ListRow
    Dim dictMails As Object
    Dim mailPart As String
    Dim compteur As Long

    LstPresences.Clear

    On Error GoTo ErrPresences
    Set wsPresences = ThisWorkbook.Sheets("PRESENCES")
    Set wsParticipants = ThisWorkbook.Sheets("PARTICIPANTS")
    Set tblPresences = wsPresences.ListObjects("TblPresences")
    Set tblParticipants = wsParticipants.ListObjects("TblParticipants")
    On Error GoTo 0

    If tblPresences.DataBodyRange Is Nothing Then Exit Sub

    ' Construire un dictionnaire ID_Participant -> Mail pour une recherche en O(1)
    Set dictMails = CreateObject("Scripting.Dictionary")
    If Not tblParticipants.DataBodyRange Is Nothing Then
        For Each ligneParticipant In tblParticipants.ListRows
            If IsNumeric(ligneParticipant.Range.Cells(1, 1).Value) Then
                dictMails(CLng(ligneParticipant.Range.Cells(1, 1).Value)) = _
                    CStr(ligneParticipant.Range.Cells(1, 9).Value)
            End If
        Next ligneParticipant
    End If

    compteur = 0
    For Each lignePresence In tblPresences.ListRows
        If IsNumeric(lignePresence.Range.Cells(1, 2).Value) Then
            If CLng(lignePresence.Range.Cells(1, 2).Value) = idAtelier Then
                ' Colonne 3 = ID_Participant
                LstPresences.AddItem CStr(lignePresence.Range.Cells(1, 3).Value)  ' col0 : ID
                LstPresences.List(compteur, 1) = _
                    CStr(lignePresence.Range.Cells(1, 4).Value)  ' col1 : Nom_Participant
                LstPresences.List(compteur, 2) = _
                    CStr(lignePresence.Range.Cells(1, 5).Value)  ' col2 : Prenom_Participant
                LstPresences.List(compteur, 3) = _
                    CStr(lignePresence.Range.Cells(1, 6).Value)  ' col3 : Statut_Participant

                ' Récupérer le mail depuis le dictionnaire (O(1))
                Dim idPart As Long
                idPart = CLng(lignePresence.Range.Cells(1, 3).Value)
                If dictMails.Exists(idPart) Then
                    mailPart = CStr(dictMails(idPart))
                Else
                    mailPart = ""
                End If
                LstPresences.List(compteur, 4) = mailPart  ' col4 : Mail

                compteur = compteur + 1
            End If
        End If
    Next lignePresence

    Exit Sub

ErrPresences:
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------
' MettreAJourCompteurs : Met à jour TxtNbParticipants et TxtNbParticipantsPro
' Paramètre :
'   idAtelier : L'ID de l'atelier
' -----------------------------------------------------------------------------
Private Sub MettreAJourCompteurs(idAtelier As Long)
    Dim wsAteliers As Worksheet
    Dim tblAteliers As ListObject
    Dim ligneAtelier As ListRow

    On Error GoTo ErrCompteurs
    Set wsAteliers = ThisWorkbook.Sheets("ATELIERS")
    Set tblAteliers = wsAteliers.ListObjects("TblAteliers")
    On Error GoTo 0

    If tblAteliers.DataBodyRange Is Nothing Then Exit Sub

    For Each ligneAtelier In tblAteliers.ListRows
        If IsNumeric(ligneAtelier.Range.Cells(1, 1).Value) Then
            If CLng(ligneAtelier.Range.Cells(1, 1).Value) = idAtelier Then
                TxtNbParticipants.Value = CStr(ligneAtelier.Range.Cells(1, 8).Value)
                TxtNbParticipantsPro.Value = CStr(ligneAtelier.Range.Cells(1, 9).Value)
                Exit For
            End If
        End If
    Next ligneAtelier

    Exit Sub

ErrCompteurs:
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------
' DefinirEtatChamps : Active ou désactive les champs et boutons de détail
' Paramètre :
'   actif : True pour activer, False pour désactiver
' -----------------------------------------------------------------------------
Private Sub DefinirEtatChamps(actif As Boolean)
    TxtNom.Enabled = actif
    TxtDate.Enabled = actif
    TxtHeureDebut.Enabled = actif
    TxtHeureFin.Enabled = actif
    CboTheme.Enabled = actif
    TxtAnimePar.Enabled = actif
    BtnSauvegarder.Enabled = actif
    BtnSupprimerAtelier.Enabled = actif
    BtnSupprimerPresence.Enabled = actif
End Sub

' -----------------------------------------------------------------------------
' ViderChampsDetail : Vide tous les champs de détail de l'atelier
' -----------------------------------------------------------------------------
Private Sub ViderChampsDetail()
    TxtNom.Value = ""
    TxtDate.Value = ""
    TxtHeureDebut.Value = ""
    TxtHeureFin.Value = ""
    TxtDuree.Value = ""
    CboTheme.ListIndex = -1
    TxtNbParticipants.Value = ""
    TxtNbParticipantsPro.Value = ""
    TxtAnimePar.Value = ""
End Sub

' -----------------------------------------------------------------------------
' UserForm_KeyDown : Fermeture du formulaire avec la touche Échap
' -----------------------------------------------------------------------------
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub
