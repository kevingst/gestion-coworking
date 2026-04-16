' =============================================================================
' UserForm : FrmNouveauParticipant
' Description : Formulaire de création d'un nouveau participant
'
' Contrôles requis (à créer dans l'éditeur VBA) :
'   - TxtNom         : TextBox — Nom (obligatoire)
'   - TxtPrenom      : TextBox — Prénom
'   - CboStatut      : ComboBox — Statut ("Projet pro" / "Lancé")
'   - TxtDateContact : TextBox — Date premier contact (JJ/MM/AAAA)
'   - TxtEntreprise  : TextBox — Nom de l'entreprise
'   - TxtCommune     : TextBox — Commune
'   - TxtCodePostal  : TextBox — Code postal
'   - TxtMail        : TextBox — Adresse mail
'   - TxtTelephone   : TextBox — Numéro de téléphone
'   - TxtActivite    : TextBox — Description de l'activité
'   - BtnEnregistrer : CommandButton — Enregistrer
'   - BtnAnnuler     : CommandButton — Annuler
'
' Voir formulaires.md pour les propriétés détaillées de chaque contrôle.
' =============================================================================

' -----------------------------------------------------------------------------
' UserForm_Initialize : Initialisation du formulaire à l'ouverture
' -----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Dim statuts As Variant
    Dim statut As Variant
    
    ' Remplir la liste des statuts dans le ComboBox
    statuts = ObtenirListeStatuts()
    CboStatut.Clear
    For Each statut In statuts
        CboStatut.AddItem statut
    Next statut
    
    ' Sélectionner le premier statut par défaut ("Projet pro")
    If CboStatut.ListCount > 0 Then
        CboStatut.ListIndex = 0
    End If
    
    ' Pré-remplir la date de premier contact avec la date du jour
    TxtDateContact.Value = Format(Now(), "DD/MM/YYYY")
    
    ' Mettre le focus sur le champ Nom
    TxtNom.SetFocus
End Sub

' -----------------------------------------------------------------------------
' BtnEnregistrer_Click : Validation et enregistrement du participant
' -----------------------------------------------------------------------------
Private Sub BtnEnregistrer_Click()
    Dim succes As Boolean
    
    ' Appel de la fonction d'enregistrement avec les valeurs des champs
    succes = EnregistrerParticipant( _
        nom:=TxtNom.Value, _
        prenom:=TxtPrenom.Value, _
        statut:=CboStatut.Value, _
        dateContact:=TxtDateContact.Value, _
        nomEntreprise:=TxtEntreprise.Value, _
        commune:=TxtCommune.Value, _
        codePostal:=TxtCodePostal.Value, _
        mail:=TxtMail.Value, _
        telephone:=TxtTelephone.Value, _
        activite:=TxtActivite.Value)
    
    ' Si l'enregistrement a réussi, fermer le formulaire
    If succes Then
        MsgBox "Le participant a été enregistré avec succès !", vbInformation, "Succès"
        Unload Me
    End If
    ' Sinon, le message d'erreur a déjà été affiché par EnregistrerParticipant
End Sub

' -----------------------------------------------------------------------------
' BtnAnnuler_Click : Fermeture du formulaire sans enregistrement
' -----------------------------------------------------------------------------
Private Sub BtnAnnuler_Click()
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

