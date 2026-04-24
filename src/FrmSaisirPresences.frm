' =============================================================================
' UserForm : FrmSaisirPresences
' Description : Formulaire de saisie des présences pour un atelier
'
' Contrôles requis (à créer dans l'éditeur VBA) :
'   - TxtRechercheAtelier     : TextBox — Recherche filtrée en temps réel dans LstAteliers
'   - LstAteliers             : ListBox — Liste des ateliers (ColumnCount=3)
'   - TxtRechercheParticipant : TextBox — Recherche filtrée en temps réel dans LstParticipants
'   - LstParticipants         : ListBox — Liste des participants (ColumnCount=4)
'                               Col 0 : ☐/☑  Col 1 : ID (masqué)  Col 2 : Nom  Col 3 : Prénom  Col 4: Statut
'   - BtnValider              : CommandButton — Valider les présences
'   - BtnAnnuler              : CommandButton — Annuler
'
' Voir formulaires.md pour les propriétés détaillées de chaque contrôle.
' =============================================================================

' ID de l'atelier sélectionné
Private idAtelierSelectionne As Long

' Tableau des IDs participants cochés (persiste entre les filtrages)
Private idsCoches() As Long
Private nbCoches As Long

' Tableau parallèle à LstParticipants : mappe index d'affichage → ID participant
' Utilisé à la place de LstParticipants.List(idx, 1) pour éviter les problèmes
' de colonne 0-width (largeur nulle) dans certaines versions de VBA/Excel.
Private idsParticipantsAffiches() As Long
Private nbAffiches As Long

' Garde-fou contre la récursion : LstParticipants_Click se redéclenche quand
' on restaure ListIndex après rechargement — ce flag coupe la boucle.
Private bCharging As Boolean

' -----------------------------------------------------------------------------
' UserForm_Initialize
' -----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    idAtelierSelectionne = 0
    nbCoches = 0
    ReDim idsCoches(0)
    nbAffiches = 0
    ReDim idsParticipantsAffiches(0)
    Call ChargerListeAteliers("")
    LstParticipants.Clear
End Sub

' -----------------------------------------------------------------------------
' TxtRechercheAtelier_Change
' -----------------------------------------------------------------------------
Private Sub TxtRechercheAtelier_Change()
    Call ChargerListeAteliers(TxtRechercheAtelier.Value)
    idAtelierSelectionne = 0
    nbCoches = 0
    ReDim idsCoches(0)
    nbAffiches = 0
    ReDim idsParticipantsAffiches(0)
    LstParticipants.Clear
End Sub

' -----------------------------------------------------------------------------
' TxtRechercheParticipant_Change
' -----------------------------------------------------------------------------
Private Sub TxtRechercheParticipant_Change()
    If idAtelierSelectionne > 0 Then
        Call ChargerListeParticipants(idAtelierSelectionne, TxtRechercheParticipant.Value)
    End If
End Sub

' -----------------------------------------------------------------------------
' ChargerListeAteliers : tri par date décroissante, filtre optionnel
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

    LstAteliers.ColumnCount = 3
    LstAteliers.ColumnWidths = "40;180;100"

    If tblAteliers.DataBodyRange Is Nothing Then Exit Sub

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

    Dim donnees() As Variant
    ReDim donnees(0 To nb - 1, 0 To 2)
    Dim dates() As Date
    ReDim dates(0 To nb - 1)

    Dim idx As Long
    idx = 0
    For Each ligneAtelier In tblAteliers.ListRows
        nomAtelier = CStr(ligneAtelier.Range.Cells(1, 2).Value)
        If filtre = "" Or InStr(1, nomAtelier, filtre, vbTextCompare) > 0 Then
            donnees(idx, 0) = ligneAtelier.Range.Cells(1, 1).Value
            donnees(idx, 1) = nomAtelier
            Dim dateVal As String
            dateVal = ""
            On Error Resume Next
            dateVal = Format(CDate(ligneAtelier.Range.Cells(1, 3).Value), "DD/MM/YYYY")
            dates(idx) = CDate(ligneAtelier.Range.Cells(1, 3).Value)
            On Error GoTo 0
            donnees(idx, 2) = dateVal
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
                tempDate = dates(m)
                dates(m) = dates(m + 1)
                dates(m + 1) = tempDate
                Dim col As Long
                For col = 0 To 2
                    tempDonnees(col) = donnees(m, col)
                    donnees(m, col) = donnees(m + 1, col)
                    donnees(m + 1, col) = tempDonnees(col)
                Next col
            End If
        Next m
    Next k

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
' LstAteliers_Click
' -----------------------------------------------------------------------------
Private Sub LstAteliers_Click()
    If LstAteliers.ListIndex < 0 Then Exit Sub

    On Error Resume Next
    idAtelierSelectionne = CLng(LstAteliers.List(LstAteliers.ListIndex, 0))
    On Error GoTo 0

    If idAtelierSelectionne <= 0 Then Exit Sub

    ' Réinitialiser les coches et pré-cocher les présences existantes
    nbCoches = 0
    ReDim idsCoches(0)

    ' Pré-cocher les participants déjà présents
    Dim presencesExistantes() As Long
    presencesExistantes = ObtenirPresencesAtelier(idAtelierSelectionne)
    Dim p As Integer
    For p = 0 To UBound(presencesExistantes)
        If presencesExistantes(p) > 0 Then
            Call AjouterCoche(presencesExistantes(p))
        End If
    Next p

    Call ChargerListeParticipants(idAtelierSelectionne, "")
End Sub

Private Sub ChargerListeParticipants(idAtelier As Long, filtre As String)
    Dim wsParticipants As Worksheet
    Dim tblParticipants As ListObject
    Dim ligneParticipant As ListRow
    Dim idPart As Long

    LstParticipants.Clear

    ' Réinitialiser le tableau parallèle
    nbAffiches = 0
    ReDim idsParticipantsAffiches(0)

    On Error GoTo ErrChargementPart
    Set wsParticipants = ThisWorkbook.Sheets("PARTICIPANTS")
    Set tblParticipants = wsParticipants.ListObjects("TblParticipants")
    On Error GoTo 0

    ' 5 colonnes : ☐/☑ | ID masqué | Nom | Prénom | Statut
    LstParticipants.ColumnCount = 5
    LstParticipants.ColumnWidths = "20;0;150;120;100"

    If Not tblParticipants.DataBodyRange Is Nothing Then
        Dim i As Integer
        i = 0
        For Each ligneParticipant In tblParticipants.ListRows
            idPart = 0
            On Error Resume Next
            idPart = CLng(ligneParticipant.Range.Cells(1, 1).Value)
            On Error GoTo 0

            If idPart > 0 Then
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
                    ' Afficher ☑ si coché, ☐ sinon
                    Dim caseChar As String
                    If EstCoche(idPart) Then
                        caseChar = ChrW(9745)  ' ☑
                    Else
                        caseChar = ChrW(9744)  ' ☐
                    End If

                    LstParticipants.AddItem caseChar          ' Col 0 : case
                    LstParticipants.List(i, 1) = idPart       ' Col 1 : ID (masqué)
                    LstParticipants.List(i, 2) = nomPart      ' Col 2 : Nom
                    LstParticipants.List(i, 3) = prenomPart   ' Col 3 : Prénom
                    LstParticipants.List(i, 4) = ligneParticipant.Range.Cells(1, 4).Value  ' Col 4 : Statut

                    ' Stocker l'ID dans le tableau parallèle (source fiable pour _Click)
                    ReDim Preserve idsParticipantsAffiches(0 To i)
                    idsParticipantsAffiches(i) = idPart

                    i = i + 1
                End If
            End If
        Next ligneParticipant
        nbAffiches = i
    End If

    Exit Sub

ErrChargementPart:
    MsgBox "Erreur lors du chargement des participants.", vbCritical, "Erreur"
End Sub

' -----------------------------------------------------------------------------
' LstParticipants_Click : bascule ☐/☑ au clic
'
' Pourquoi on recharge la liste entière ?
' Excel/VBA ne rafraîchit pas visuellement la colonne 0 d'une ListBox quand on
' écrit LstParticipants.List(idx, 0) = "..." (bug connu sur la première colonne
' remplie via AddItem). La seule façon fiable de mettre à jour le symbole ☐/☑
' est de reconstruire la liste. L'état des coches est dans idsCoches() qui
' persiste, donc ChargerListeParticipants affichera le bon symbole.
'
' bCharging évite la récursion : restaurer ListIndex déclenche un nouveau _Click.
' -----------------------------------------------------------------------------
Private Sub LstParticipants_Click()
    If bCharging Then Exit Sub

    Dim idx As Long
    idx = LstParticipants.ListIndex
    If idx < 0 Or idx >= nbAffiches Then Exit Sub

    Dim idPart As Long
    idPart = idsParticipantsAffiches(idx)
    If idPart <= 0 Then Exit Sub

    If EstCoche(idPart) Then
        Call RetirerCoche(idPart)
    Else
        Call AjouterCoche(idPart)
    End If

    ' Recharger la liste pour que le symbole ☐/☑ soit bien mis à jour
    Dim topIdx As Long
    topIdx = LstParticipants.TopIndex

    bCharging = True
    Call ChargerListeParticipants(idAtelierSelectionne, TxtRechercheParticipant.Value)
    If topIdx < LstParticipants.ListCount Then
        LstParticipants.TopIndex = topIdx
    End If
    If idx < LstParticipants.ListCount Then
        LstParticipants.ListIndex = idx
    End If
    bCharging = False
End Sub

' -----------------------------------------------------------------------------
' EstCoche : vérifie si un ID est dans idsCoches
' -----------------------------------------------------------------------------
Private Function EstCoche(idPart As Long) As Boolean
    EstCoche = False
    If nbCoches = 0 Then Exit Function
    Dim i As Long
    For i = 0 To nbCoches - 1
        If idsCoches(i) = idPart Then
            EstCoche = True
            Exit Function
        End If
    Next i
End Function

' -----------------------------------------------------------------------------
' AjouterCoche : ajoute un ID dans idsCoches
' -----------------------------------------------------------------------------
Private Sub AjouterCoche(idPart As Long)
    If EstCoche(idPart) Then Exit Sub
    ReDim Preserve idsCoches(0 To nbCoches)
    idsCoches(nbCoches) = idPart
    nbCoches = nbCoches + 1
End Sub

' -----------------------------------------------------------------------------
' RetirerCoche : retire un ID de idsCoches
' -----------------------------------------------------------------------------
Private Sub RetirerCoche(idPart As Long)
    If nbCoches = 0 Then Exit Sub
    Dim newIds() As Long
    Dim newNb As Long
    newNb = 0
    ReDim newIds(0 To nbCoches - 1)
    Dim i As Long
    For i = 0 To nbCoches - 1
        If idsCoches(i) <> idPart Then
            newIds(newNb) = idsCoches(i)
            newNb = newNb + 1
        End If
    Next i
    nbCoches = newNb
    If newNb > 0 Then
        ReDim Preserve newIds(0 To newNb - 1)
        idsCoches = newIds
    Else
        ReDim idsCoches(0)
    End If
End Sub

' -----------------------------------------------------------------------------
' BtnValider_Click
' -----------------------------------------------------------------------------
Private Sub BtnValider_Click()
    If idAtelierSelectionne <= 0 Then
        MsgBox "Veuillez sélectionner un atelier.", vbExclamation, "Sélection manquante"
        Exit Sub
    End If

    If nbCoches = 0 Then
        MsgBox "Veuillez sélectionner au moins un participant.", vbExclamation, "Sélection manquante"
        Exit Sub
    End If

    ' Construire le tableau des IDs cochés
    Dim idsSelectionnes() As Long
    ReDim idsSelectionnes(0 To nbCoches - 1)
    Dim i As Long
    For i = 0 To nbCoches - 1
        idsSelectionnes(i) = idsCoches(i)
    Next i

    Dim succes As Boolean
    succes = EnregistrerPresences(idAtelierSelectionne, idsSelectionnes)

    If succes Then
        MsgBox "Les présences ont été enregistrées avec succès !", vbInformation, "Succès"
        Unload Me
    End If
End Sub

' -----------------------------------------------------------------------------
' BtnAnnuler_Click
' -----------------------------------------------------------------------------
Private Sub BtnAnnuler_Click()
    Unload Me
End Sub
