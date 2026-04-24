' =============================================================================
' UserForm : FrmNouvelAtelier
' Description : Formulaire de création d'un nouvel atelier
'
' Contrôles requis (à créer dans l'éditeur VBA) :
'   - TxtNom        : TextBox — Nom de l'atelier
'   - TxtDate       : TextBox — Date (JJ/MM/AAAA)
'   - TxtHeureDebut : TextBox — Heure de début (HH:MM)
'   - TxtHeureFin   : TextBox — Heure de fin (HH:MM)
'   - CboTheme      : ComboBox — Thème de l'atelier
'   - TxtAnimePar   : TextBox — Animé par (saisie libre)
'   - BtnEnregistrer: CommandButton — Enregistrer
'   - BtnAnnuler    : CommandButton — Annuler
'
' Voir formulaires.md pour les propriétés détaillées de chaque contrôle.
' =============================================================================

' -----------------------------------------------------------------------------
' UserForm_Initialize : Initialisation du formulaire à l'ouverture
' -----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Dim themes As Variant
    Dim theme As Variant
    
    ' Remplir la liste des thèmes dans le ComboBox
    themes = ObtenirListeThemes()
    CboTheme.Clear
    For Each theme In themes
        CboTheme.AddItem theme
    Next theme
    
    ' Sélectionner le premier thème par défaut
    If CboTheme.ListCount > 0 Then
        CboTheme.ListIndex = 0
    End If
    
    ' Pré-remplir la date avec la date du jour
    TxtDate.Value = Format(Now(), "DD/MM/YYYY")
    
    ' Mettre le focus sur le champ Nom
    TxtNom.SetFocus
End Sub

' -----------------------------------------------------------------------------
' BtnEnregistrer_Click : Validation et enregistrement de l'atelier
' -----------------------------------------------------------------------------
Private Sub BtnEnregistrer_Click()
    Dim succes As Boolean
    
    ' Appel de la fonction d'enregistrement avec les valeurs des champs
    succes = EnregistrerAtelier( _
        nom:=TxtNom.Value, _
        dateStr:=TxtDate.Value, _
        heureDebut:=TxtHeureDebut.Value, _
        heureFin:=TxtHeureFin.Value, _
        theme:=CboTheme.Value, _
        animepar:=TxtAnimePar.Value)
    
    ' Si l'enregistrement a réussi, fermer le formulaire
    If succes Then
        MsgBox "L'atelier a été enregistré avec succès !", vbInformation, "Succès"
        Unload Me
    End If
    ' Sinon, le message d'erreur a déjà été affiché par la fonction EnregistrerAtelier
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

