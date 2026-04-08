' =============================================================================
' UserForm : FrmSaisirPresences
' Description : Formulaire de saisie des présences pour un atelier
'
' Contrôles requis (à créer dans l'éditeur VBA) :
'   - LstAteliers     : ListBox — Liste des ateliers (ColumnCount=3)
'   - LstParticipants : ListBox — Liste des participants (MultiSelect, ColumnCount=3)
'   - BtnValider      : CommandButton — Valider les présences
'   - BtnAnnuler      : CommandButton — Annuler
'
' Voir formulaires.md pour les propriétés détaillées de chaque contrôle.
' =============================================================================

' Variable pour stocker l'ID de l'atelier actuellement sélectionné
Private idAtelierSelectionne As Long

' -----------------------------------------------------------------------------
' UserForm_Initialize : Initialisation du formulaire à l'ouverture
' -----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    idAtelierSelectionne = 0
    
    ' Charger la liste des ateliers
    Call ChargerListeAteliers
    
    ' Vider la liste des participants (sera chargée à la sélection d'un atelier)
    LstParticipants.Clear
End Sub

' -----------------------------------------------------------------------------
' ChargerListeAteliers : Charge tous les ateliers dans LstAteliers
' -----------------------------------------------------------------------------
Private Sub ChargerListeAteliers()
    Dim wsAteliers As Worksheet
    Dim tblAteliers As ListObject
    Dim ligneAtelier As ListRow
    
    LstAteliers.Clear
    
    On Error GoTo ErrChargement
    Set wsAteliers = ThisWorkbook.Sheets("ATELIERS")
    Set tblAteliers = wsAteliers.ListObjects("TblAteliers")
    On Error GoTo 0
    
    ' Configurer les colonnes de la ListBox
    LstAteliers.ColumnCount = 3
    LstAteliers.ColumnWidths = "40;180;100"
    
    ' Ajouter chaque atelier (ID, Nom, Date)
    If Not tblAteliers.DataBodyRange Is Nothing Then
        For Each ligneAtelier In tblAteliers.ListRows
            LstAteliers.AddItem ligneAtelier.Range.Cells(1, 1).Value  ' ID_Atelier
            LstAteliers.List(LstAteliers.ListCount - 1, 1) = ligneAtelier.Range.Cells(1, 2).Value  ' Nom
            
            ' Formatage de la date
            Dim dateVal As String
            dateVal = ""
            On Error Resume Next
            dateVal = Format(CDate(ligneAtelier.Range.Cells(1, 3).Value), "DD/MM/YYYY")
            On Error GoTo 0
            LstAteliers.List(LstAteliers.ListCount - 1, 2) = dateVal  ' Date
        Next ligneAtelier
    End If
    
    Exit Sub
    
ErrChargement:
    MsgBox "Erreur lors du chargement des ateliers.", vbCritical, "Erreur"
End Sub

' -----------------------------------------------------------------------------
' LstAteliers_Click : Chargement des participants quand un atelier est sélectionné
' -----------------------------------------------------------------------------
Private Sub LstAteliers_Click()
    ' Récupérer l'ID de l'atelier sélectionné (première colonne)
    If LstAteliers.ListIndex < 0 Then Exit Sub
    
    On Error Resume Next
    idAtelierSelectionne = CLng(LstAteliers.List(LstAteliers.ListIndex, 0))
    On Error GoTo 0
    
    If idAtelierSelectionne <= 0 Then Exit Sub
    
    ' Charger la liste des participants pour cet atelier
    Call ChargerListeParticipants(idAtelierSelectionne)
End Sub

' -----------------------------------------------------------------------------
' ChargerListeParticipants : Charge tous les participants dans LstParticipants
' en pré-cochant ceux déjà présents à l'atelier
' Paramètre :
'   idAtelier : L'ID de l'atelier sélectionné
' -----------------------------------------------------------------------------
Private Sub ChargerListeParticipants(idAtelier As Long)
    Dim wsParticipants As Worksheet
    Dim tblParticipants As ListObject
    Dim tblPresences As ListObject
    Dim wsPresences As Worksheet
    Dim ligneParticipant As ListRow
    Dim idPart As Long
    
    LstParticipants.Clear
    
    On Error GoTo ErrChargementPart
    Set wsParticipants = ThisWorkbook.Sheets("PARTICIPANTS")
    Set tblParticipants = wsParticipants.ListObjects("TblParticipants")
    Set wsPresences = ThisWorkbook.Sheets("PRESENCES")
    Set tblPresences = wsPresences.ListObjects("TblPresences")
    On Error GoTo 0
    
    ' Configurer les colonnes
    LstParticipants.ColumnCount = 3
    LstParticipants.ColumnWidths = "150;120;100"
    
    ' Obtenir la liste des participants déjà présents à cet atelier
    Dim presencesExistantes() As Long
    presencesExistantes = ObtenirPresencesAtelier(idAtelier)
    
    ' Ajouter chaque participant (Nom, Prénom, Statut)
    If Not tblParticipants.DataBodyRange Is Nothing Then
        Dim i As Integer
        i = 0
        For Each ligneParticipant In tblParticipants.ListRows
            idPart = 0
            On Error Resume Next
            idPart = CLng(ligneParticipant.Range.Cells(1, 1).Value)
            On Error GoTo 0
            
            If idPart > 0 Then
                LstParticipants.AddItem ligneParticipant.Range.Cells(1, 2).Value  ' Nom
                LstParticipants.List(i, 1) = ligneParticipant.Range.Cells(1, 3).Value  ' Prenom
                LstParticipants.List(i, 2) = ligneParticipant.Range.Cells(1, 4).Value  ' Statut
                
                ' Pré-cocher si déjà présent
                If EstDansTableau(presencesExistantes, idPart) Then
                    LstParticipants.Selected(i) = True
                End If
                
                i = i + 1
            End If
        Next ligneParticipant
    End If
    
    Exit Sub
    
ErrChargementPart:
    MsgBox "Erreur lors du chargement des participants.", vbCritical, "Erreur"
End Sub

' -----------------------------------------------------------------------------
' EstDansTableau : Vérifie si une valeur est présente dans un tableau Long
' -----------------------------------------------------------------------------
Private Function EstDansTableau(tableau() As Long, valeur As Long) As Boolean
    Dim i As Integer
    EstDansTableau = False
    
    On Error Resume Next
    If UBound(tableau) < 0 Then Exit Function
    On Error GoTo 0
    
    For i = 0 To UBound(tableau)
        If tableau(i) = valeur Then
            EstDansTableau = True
            Exit Function
        End If
    Next i
End Function

' -----------------------------------------------------------------------------
' BtnValider_Click : Validation et enregistrement des présences
' -----------------------------------------------------------------------------
Private Sub BtnValider_Click()
    Dim wsParticipants As Worksheet
    Dim tblParticipants As ListObject
    Dim ligneParticipant As ListRow
    Dim idsSelectionnes() As Long
    Dim nb As Integer
    Dim i As Integer
    
    ' Vérifier qu'un atelier est sélectionné
    If idAtelierSelectionne <= 0 Then
        MsgBox "Veuillez sélectionner un atelier.", vbExclamation, "Sélection manquante"
        Exit Sub
    End If
    
    ' Vérifier qu'au moins un participant est sélectionné
    Dim nbSelectionnes As Integer
    nbSelectionnes = 0
    For i = 0 To LstParticipants.ListCount - 1
        If LstParticipants.Selected(i) Then
            nbSelectionnes = nbSelectionnes + 1
        End If
    Next i
    
    If nbSelectionnes = 0 Then
        MsgBox "Veuillez sélectionner au moins un participant.", vbExclamation, "Sélection manquante"
        Exit Sub
    End If
    
    ' Récupérer les IDs des participants sélectionnés
    ' On doit naviguer dans TblParticipants dans le même ordre que LstParticipants
    On Error GoTo ErrValider
    Set wsParticipants = ThisWorkbook.Sheets("PARTICIPANTS")
    Set tblParticipants = wsParticipants.ListObjects("TblParticipants")
    On Error GoTo 0
    
    ReDim idsSelectionnes(0 To nbSelectionnes - 1)
    nb = 0
    i = 0
    
    If Not tblParticipants.DataBodyRange Is Nothing Then
        For Each ligneParticipant In tblParticipants.ListRows
            Dim idPart As Long
            idPart = 0
            On Error Resume Next
            idPart = CLng(ligneParticipant.Range.Cells(1, 1).Value)
            On Error GoTo 0
            
            If idPart > 0 Then
                If i < LstParticipants.ListCount Then
                    If LstParticipants.Selected(i) Then
                        idsSelectionnes(nb) = idPart
                        nb = nb + 1
                    End If
                    i = i + 1
                End If
            End If
        Next ligneParticipant
    End If
    
    ' Enregistrer les présences
    Dim succes As Boolean
    succes = EnregistrerPresences(idAtelierSelectionne, idsSelectionnes)
    
    If succes Then
        MsgBox "Les présences ont été enregistrées avec succès !", vbInformation, "Succès"
        Unload Me
    End If
    
    Exit Sub
    
ErrValider:
    MsgBox "Erreur lors de la validation des présences.", vbCritical, "Erreur"
End Sub

' -----------------------------------------------------------------------------
' BtnAnnuler_Click : Fermeture du formulaire sans enregistrement
' -----------------------------------------------------------------------------
Private Sub BtnAnnuler_Click()
    Unload Me
End Sub
