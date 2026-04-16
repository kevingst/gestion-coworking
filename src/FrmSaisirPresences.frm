' =============================================================================
' UserForm : FrmSaisirPresences
' Description : Formulaire de saisie des présences pour un atelier
'
' Contrôles requis (à créer dans l'éditeur VBA) :
'   - TxtRechercheAtelier     : TextBox — Recherche filtrée en temps réel dans LstAteliers
'   - LstAteliers             : ListBox — Liste des ateliers (ColumnCount=3)
'   - TxtRechercheParticipant : TextBox — Recherche filtrée en temps réel dans LstParticipants
'   - LstParticipants         : ListBox — Liste des participants (MultiSelect, ColumnCount=4)
'   - BtnValider              : CommandButton — Valider les présences
'   - BtnAnnuler              : CommandButton — Annuler
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

    ' Charger la liste des ateliers (triés par date décroissante)
    Call ChargerListeAteliers("")

    ' Vider la liste des participants (sera chargée à la sélection d'un atelier)
    LstParticipants.Clear
End Sub

' -----------------------------------------------------------------------------
' TxtRechercheAtelier_Change : Filtrage en temps réel de la liste des ateliers
' -----------------------------------------------------------------------------
Private Sub TxtRechercheAtelier_Change()
    Call ChargerListeAteliers(TxtRechercheAtelier.Value)
    idAtelierSelectionne = 0
    LstParticipants.Clear
End Sub

' -----------------------------------------------------------------------------
' TxtRechercheParticipant_Change : Filtrage en temps réel de la liste des participants
' -----------------------------------------------------------------------------
Private Sub TxtRechercheParticipant_Change()
    If idAtelierSelectionne > 0 Then
        Call ChargerListeParticipants(idAtelierSelectionne, TxtRechercheParticipant.Value)
    End If
End Sub

' -----------------------------------------------------------------------------
' ChargerListeAteliers : Charge les ateliers dans LstAteliers (tri par date décroissante)
' Paramètre :
'   filtre : Texte de filtre sur le nom (vide = tous)
' -----------------------------------------------------------------------------
Private Sub ChargerListeAteliers(filtre As String)
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

    If tblAteliers.DataBodyRange Is Nothing Then Exit Sub

    ' Collecter les lignes correspondant au filtre
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
    ReDim donnees(0 To nb - 1, 0 To 2)
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

            ' Stocker la date brute pour le tri
            On Error Resume Next
            dates(idx) = CDate(ligneAtelier.Range.Cells(1, 3).Value)
            On Error GoTo 0

            idx = idx + 1
        End If
    Next ligneAtelier

    ' Tri à bulles par date décroissante
    Dim tempDate As Date
    Dim tempDonnees(0 To 2) As Variant
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
                For col = 0 To 2
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
    Next idx

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
    Call ChargerListeParticipants(idAtelierSelectionne, "")
End Sub

' -----------------------------------------------------------------------------
' ChargerListeParticipants : Charge les participants dans LstParticipants
' en pré-cochant ceux déjà présents à l'atelier, avec filtre optionnel
' Paramètres :
'   idAtelier : L'ID de l'atelier sélectionné
'   filtre    : Texte de filtre sur Nom ou Prénom (vide = tous)
' -----------------------------------------------------------------------------
Private Sub ChargerListeParticipants(idAtelier As Long, filtre As String)
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

    ' Configurer les colonnes (ID masqué en col 0, Nom, Prénom, Statut)
    LstParticipants.ColumnCount = 4
    LstParticipants.ColumnWidths = "0;150;120;100"

    ' Obtenir la liste des participants déjà présents à cet atelier
    Dim presencesExistantes() As Long
    presencesExistantes = ObtenirPresencesAtelier(idAtelier)

    ' Ajouter chaque participant filtré (ID masqué, Nom, Prénom, Statut)
    If Not tblParticipants.DataBodyRange Is Nothing Then
        Dim i As Integer
        i = 0
        For Each ligneParticipant In tblParticipants.ListRows
            idPart = 0
            On Error Resume Next
            idPart = CLng(ligneParticipant.Range.Cells(1, 1).Value)
            On Error GoTo 0

            If idPart > 0 Then
                ' Appliquer le filtre sur Nom ou Prénom
                Dim nomPart As String
                Dim prenomPart As String
                nomPart = CStr(ligneParticipant.Range.Cells(1, 2).Value)
                prenomPart = CStr(ligneParticipant.Range.Cells(1, 3).Value)

                Dim inclure As Boolean
                If filtre = "" Then
                    inclure = True
                Else
                    inclure = InStr(1, nomPart, filtre, vbTextCompare) > 0 Or _
                              InStr(1, prenomPart, filtre, vbTextCompare) > 0
                End If

                If inclure Then
                    LstParticipants.AddItem idPart                                          ' ID (masqué)
                    LstParticipants.List(i, 1) = nomPart                                    ' Nom
                    LstParticipants.List(i, 2) = prenomPart                                 ' Prénom
                    LstParticipants.List(i, 3) = ligneParticipant.Range.Cells(1, 4).Value  ' Statut

                    ' Pré-cocher si déjà présent
                    If EstDansTableau(presencesExistantes, idPart) Then
                        LstParticipants.Selected(i) = True
                    End If

                    i = i + 1
                End If
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
    Dim idsSelectionnes() As Long
    Dim nb As Integer
    Dim i As Integer
    Dim idx As Integer

    ' Vérifier qu'un atelier est sélectionné
    If idAtelierSelectionne <= 0 Then
        MsgBox "Veuillez sélectionner un atelier.", vbExclamation, "Sélection manquante"
        Exit Sub
    End If

    ' Compter les participants sélectionnés
    nb = 0
    For i = 0 To LstParticipants.ListCount - 1
        If LstParticipants.Selected(i) Then
            nb = nb + 1
        End If
    Next i

    If nb = 0 Then
        MsgBox "Veuillez sélectionner au moins un participant.", vbExclamation, "Sélection manquante"
        Exit Sub
    End If

    ' Récupérer les IDs depuis la colonne 0 (masquée) de LstParticipants
    ReDim idsSelectionnes(0 To nb - 1)
    idx = 0
    For i = 0 To LstParticipants.ListCount - 1
        If LstParticipants.Selected(i) Then
            On Error Resume Next
            idsSelectionnes(idx) = CLng(LstParticipants.List(i, 0))
            On Error GoTo 0
            idx = idx + 1
        End If
    Next i

    ' Enregistrer les présences
    Dim succes As Boolean
    succes = EnregistrerPresences(idAtelierSelectionne, idsSelectionnes)

    If succes Then
        MsgBox "Les présences ont été enregistrées avec succès !", vbInformation, "Succès"
        Unload Me
    End If
End Sub

' -----------------------------------------------------------------------------
' BtnAnnuler_Click : Fermeture du formulaire sans enregistrement
' -----------------------------------------------------------------------------
Private Sub BtnAnnuler_Click()
    Unload Me
End Sub

