' =============================================================================
' UserForm : FrmGererParticipants
' Description : Formulaire de recherche et modification des participants
'
' Contrôles requis (à créer dans l'éditeur VBA) :
'   Zone de recherche :
'   - TxtRecherche   : TextBox — Critère de recherche (Nom ou Prénom)
'   - BtnRechercher  : CommandButton — Lancer la recherche
'   - LstResultats   : ListBox — Résultats (ColumnCount=4 : ID, Nom, Prénom, Statut)
'
'   Zone d'édition (Enabled=False par défaut) :
'   - TxtENom        : TextBox — Nom (éditable)
'   - TxtEPrenom     : TextBox — Prénom (éditable)
'   - CboEStatut     : ComboBox — Statut (éditable)
'   - TxtEDateContact: TextBox — Date premier contact (éditable)
'   - TxtEEntreprise : TextBox — Nom entreprise (éditable)
'   - TxtECommune    : TextBox — Commune (éditable)
'   - TxtECodePostal : TextBox — Code postal (éditable)
'   - TxtEMail       : TextBox — Mail (éditable)
'   - TxtETelephone  : TextBox — Téléphone (éditable)
'   - TxtEActivite   : TextBox — Activité (éditable)
'
'   Boutons :
'   - BtnModifier    : CommandButton — Activer les champs d'édition
'   - BtnSauvegarder : CommandButton — Sauvegarder les modifications
'   - BtnFermer      : CommandButton — Fermer le formulaire
'
' Voir formulaires.md pour les propriétés détaillées de chaque contrôle.
' =============================================================================

' Variable pour stocker l'ID du participant en cours de modification
Private idParticipantSelectionne As Long

' -----------------------------------------------------------------------------
' UserForm_Initialize : Initialisation du formulaire à l'ouverture
' -----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    idParticipantSelectionne = 0
    
    ' Configurer la ListBox des résultats
    LstResultats.ColumnCount = 4
    LstResultats.ColumnWidths = "40;150;120;100"
    
    ' Remplir le ComboBox des statuts
    Dim statuts As String()
    Dim statut As Variant
    statuts = ObtenirListeStatuts()
    CboEStatut.Clear
    For Each statut In statuts
        CboEStatut.AddItem statut
    Next statut
    
    ' Désactiver les champs d'édition par défaut
    Call DefinirEtatEdition(False)
    
    ' Charger tous les participants dans la liste (recherche vide = tous)
    Call LancerRecherche("")
    
    ' Mettre le focus sur la recherche
    TxtRecherche.SetFocus
End Sub

' -----------------------------------------------------------------------------
' BtnRechercher_Click : Lance la recherche
' -----------------------------------------------------------------------------
Private Sub BtnRechercher_Click()
    Call LancerRecherche(TxtRecherche.Value)
End Sub

' -----------------------------------------------------------------------------
' TxtRecherche_KeyDown : Lancement de la recherche avec la touche Entrée
' -----------------------------------------------------------------------------
Private Sub TxtRecherche_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then ' Touche Entrée
        Call LancerRecherche(TxtRecherche.Value)
    End If
End Sub

' -----------------------------------------------------------------------------
' LancerRecherche : Effectue la recherche et charge les résultats
' Paramètre :
'   critere : Le texte à rechercher (vide = tous les participants)
' -----------------------------------------------------------------------------
Private Sub LancerRecherche(critere As String)
    Dim resultats As Variant
    Dim i As Integer
    
    ' Vider la liste et réinitialiser la sélection
    LstResultats.Clear
    idParticipantSelectionne = 0
    Call DefinirEtatEdition(False)
    Call ViderChampsEdition
    
    ' Rechercher les participants
    resultats = RechercherParticipants(critere)
    
    ' Afficher les résultats dans la ListBox
    On Error Resume Next
    If UBound(resultats, 1) < 0 Then Exit Sub
    On Error GoTo 0
    
    For i = 0 To UBound(resultats, 1)
        If Not IsEmpty(resultats(i, 0)) And resultats(i, 0) <> "" Then
            LstResultats.AddItem resultats(i, 0)  ' ID
            LstResultats.List(LstResultats.ListCount - 1, 1) = resultats(i, 1)  ' Nom
            LstResultats.List(LstResultats.ListCount - 1, 2) = resultats(i, 2)  ' Prenom
            LstResultats.List(LstResultats.ListCount - 1, 3) = resultats(i, 3)  ' Statut
        End If
    Next i
End Sub

' -----------------------------------------------------------------------------
' LstResultats_Click : Sélection d'un participant dans la liste
' -----------------------------------------------------------------------------
Private Sub LstResultats_Click()
    ' Récupérer l'ID du participant sélectionné
    If LstResultats.ListIndex < 0 Then Exit Sub
    
    On Error Resume Next
    idParticipantSelectionne = CLng(LstResultats.List(LstResultats.ListIndex, 0))
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------
' BtnModifier_Click : Active les champs d'édition et charge les données
' -----------------------------------------------------------------------------
Private Sub BtnModifier_Click()
    ' Vérifier qu'un participant est sélectionné
    If idParticipantSelectionne <= 0 Then
        MsgBox "Veuillez sélectionner un participant dans la liste.", vbExclamation, "Sélection manquante"
        Exit Sub
    End If
    
    ' Charger les données du participant dans les champs d'édition
    If Not ChargerDonneesParticipant(idParticipantSelectionne) Then
        Exit Sub
    End If
    
    ' Activer les champs d'édition
    Call DefinirEtatEdition(True)
    
    ' Mettre le focus sur le premier champ
    TxtENom.SetFocus
End Sub

' -----------------------------------------------------------------------------
' ChargerDonneesParticipant : Charge les données d'un participant dans les champs
' Paramètre :
'   idParticipant : L'ID du participant à charger
' Retourne True si succès
' -----------------------------------------------------------------------------
Private Function ChargerDonneesParticipant(idParticipant As Long) As Boolean
    Dim wsParticipants As Worksheet
    Dim tblParticipants As ListObject
    Dim ligneParticipant As ListRow
    
    ChargerDonneesParticipant = False
    
    On Error GoTo ErrChargement
    Set wsParticipants = ThisWorkbook.Sheets("PARTICIPANTS")
    Set tblParticipants = wsParticipants.ListObjects("TblParticipants")
    On Error GoTo 0
    
    If tblParticipants.DataBodyRange Is Nothing Then Exit Function
    
    ' Rechercher le participant
    For Each ligneParticipant In tblParticipants.ListRows
        If IsNumeric(ligneParticipant.Range.Cells(1, 1).Value) Then
            If CLng(ligneParticipant.Range.Cells(1, 1).Value) = idParticipant Then
                ' Remplir les champs d'édition
                TxtENom.Value = CStr(ligneParticipant.Range.Cells(1, 2).Value)
                TxtEPrenom.Value = CStr(ligneParticipant.Range.Cells(1, 3).Value)
                
                ' Sélectionner le bon statut dans le ComboBox
                Dim statutVal As String
                statutVal = CStr(ligneParticipant.Range.Cells(1, 4).Value)
                Dim j As Integer
                For j = 0 To CboEStatut.ListCount - 1
                    If CboEStatut.List(j) = statutVal Then
                        CboEStatut.ListIndex = j
                        Exit For
                    End If
                Next j
                
                ' Date au format JJ/MM/AAAA
                Dim dateVal As String
                dateVal = ""
                On Error Resume Next
                dateVal = Format(CDate(ligneParticipant.Range.Cells(1, 5).Value), "DD/MM/YYYY")
                On Error GoTo 0
                TxtEDateContact.Value = dateVal
                
                TxtEEntreprise.Value = CStr(ligneParticipant.Range.Cells(1, 6).Value)
                TxtECommune.Value = CStr(ligneParticipant.Range.Cells(1, 7).Value)
                TxtECodePostal.Value = CStr(ligneParticipant.Range.Cells(1, 8).Value)
                TxtEMail.Value = CStr(ligneParticipant.Range.Cells(1, 9).Value)
                TxtETelephone.Value = CStr(ligneParticipant.Range.Cells(1, 10).Value)
                TxtEActivite.Value = CStr(ligneParticipant.Range.Cells(1, 11).Value)
                
                ChargerDonneesParticipant = True
                Exit Function
            End If
        End If
    Next ligneParticipant
    
    MsgBox "Participant introuvable.", vbExclamation, "Erreur"
    Exit Function
    
ErrChargement:
    MsgBox "Erreur lors du chargement des données.", vbCritical, "Erreur"
End Function

' -----------------------------------------------------------------------------
' BtnSauvegarder_Click : Sauvegarde les modifications du participant
' -----------------------------------------------------------------------------
Private Sub BtnSauvegarder_Click()
    ' Vérifier qu'un participant est en cours de modification
    If idParticipantSelectionne <= 0 Then
        MsgBox "Aucun participant sélectionné pour modification.", vbExclamation, "Erreur"
        Exit Sub
    End If
    
    ' Appel de la fonction de modification
    Dim succes As Boolean
    succes = ModifierParticipant( _
        idParticipant:=idParticipantSelectionne, _
        nom:=TxtENom.Value, _
        prenom:=TxtEPrenom.Value, _
        statut:=CboEStatut.Value, _
        dateContact:=TxtEDateContact.Value, _
        nomEntreprise:=TxtEEntreprise.Value, _
        commune:=TxtECommune.Value, _
        codePostal:=TxtECodePostal.Value, _
        mail:=TxtEMail.Value, _
        telephone:=TxtETelephone.Value, _
        activite:=TxtEActivite.Value)
    
    If succes Then
        MsgBox "Les modifications ont été sauvegardées avec succès !", vbInformation, "Succès"
        
        ' Désactiver les champs d'édition
        Call DefinirEtatEdition(False)
        
        ' Rafraîchir la liste des résultats
        Call LancerRecherche(TxtRecherche.Value)
    End If
End Sub

' -----------------------------------------------------------------------------
' DefinirEtatEdition : Active ou désactive les champs d'édition
' Paramètre :
'   actif : True pour activer, False pour désactiver
' -----------------------------------------------------------------------------
Private Sub DefinirEtatEdition(actif As Boolean)
    TxtENom.Enabled = actif
    TxtEPrenom.Enabled = actif
    CboEStatut.Enabled = actif
    TxtEDateContact.Enabled = actif
    TxtEEntreprise.Enabled = actif
    TxtECommune.Enabled = actif
    TxtECodePostal.Enabled = actif
    TxtEMail.Enabled = actif
    TxtETelephone.Enabled = actif
    TxtEActivite.Enabled = actif
    BtnSauvegarder.Enabled = actif
End Sub

' -----------------------------------------------------------------------------
' ViderChampsEdition : Vide tous les champs d'édition
' -----------------------------------------------------------------------------
Private Sub ViderChampsEdition()
    TxtENom.Value = ""
    TxtEPrenom.Value = ""
    CboEStatut.ListIndex = -1
    TxtEDateContact.Value = ""
    TxtEEntreprise.Value = ""
    TxtECommune.Value = ""
    TxtECodePostal.Value = ""
    TxtEMail.Value = ""
    TxtETelephone.Value = ""
    TxtEActivite.Value = ""
End Sub

' -----------------------------------------------------------------------------
' BtnFermer_Click : Fermeture du formulaire
' -----------------------------------------------------------------------------
Private Sub BtnFermer_Click()
    Unload Me
End Sub

' -----------------------------------------------------------------------------
' UserForm_KeyDown : Gestion des touches clavier
' -----------------------------------------------------------------------------
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Fermer le formulaire avec la touche Échap
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub
