' =============================================================================
' Module : ModuleAteliers
' Description : Gestion des ateliers â€” crÃ©ation et enregistrement
' =============================================================================

' Liste des thÃ¨mes disponibles pour les ateliers
Public Const THEMES_ATELIERS As String = "Administration,Réseautage,Création,Numérique,Bien-être"

' -----------------------------------------------------------------------------
' EnregistrerAtelier : Enregistre un nouvel atelier dans la feuille ATELIERS
' ParamÃ¨tres :
'   nom        : Nom de l'atelier
'   dateStr    : Date au format JJ/MM/AAAA
'   heureDebut : Heure de dÃ©but au format HH:MM
'   heureFin   : Heure de fin au format HH:MM
'   theme      : ThÃ¨me de l'atelier
'   animepar   : Nom de l'animateur de l'atelier
' Retourne True si succÃ¨s, False si erreur
' -----------------------------------------------------------------------------
Public Function EnregistrerAtelier(nom As String, dateStr As String, _
                                   heureDebut As String, heureFin As String, _
                                   theme As String, animepar As String) As Boolean
    Dim wsAteliers As Worksheet
    Dim tblAteliers As ListObject
    Dim nouvelleDate As Date
    Dim heureDebutVal As Date
    Dim heureFinVal As Date
    Dim dureeMinutes As Long
    Dim dureeFormatee As String
    Dim nouvelID As Long
    Dim nouvelleLigne As ListRow
    
    EnregistrerAtelier = False
    
    ' Validation des champs obligatoires
    If Trim(nom) = "" Then
        MsgBox "Le nom de l'atelier est obligatoire.", vbExclamation, "Champ obligatoire"
        Exit Function
    End If
    
    If Trim(dateStr) = "" Then
        MsgBox "La date de l'atelier est obligatoire.", vbExclamation, "Champ obligatoire"
        Exit Function
    End If
    
    ' Validation et conversion de la date
    On Error GoTo ErrDate
    nouvelleDate = CDate(dateStr)
    On Error GoTo 0
    
    ' Validation et conversion des heures
    If Trim(heureDebut) <> "" And Trim(heureFin) <> "" Then
        On Error GoTo ErrHeure
        heureDebutVal = CDate(heureDebut)
        heureFinVal = CDate(heureFin)
        On Error GoTo 0
        
        ' VÃ©rification que l'heure de fin est aprÃ¨s l'heure de dÃ©but
        If heureFinVal <= heureDebutVal Then
            MsgBox "L'heure de fin doit Ãªtre postÃ©rieure Ã  l'heure de dÃ©but.", vbExclamation, "Heure invalide"
            Exit Function
        End If
        
        ' Calcul de la durÃ©e en minutes
        dureeMinutes = DateDiff("n", heureDebutVal, heureFinVal)
        
        ' Formatage de la durÃ©e en HH:MM
        dureeFormatee = Format(dureeMinutes \ 60, "00") & ":" & Format(dureeMinutes Mod 60, "00")
    Else
        dureeFormatee = "00:00"
    End If
    
    ' AccÃ¨s Ã  la feuille et au tableau
    On Error GoTo ErrFeuille
    Set wsAteliers = ThisWorkbook.Sheets("ATELIERS")
    Set tblAteliers = wsAteliers.ListObjects("TblAteliers")
    On Error GoTo 0
    
    ' DÃ©sprotÃ©ger la feuille pour Ã©crire
    wsAteliers.Unprotect Password:=MOT_DE_PASSE
    
    ' Calcul du prochain ID (auto-incrÃ©mentÃ©)
    nouvelID = 1
    If Not tblAteliers.DataBodyRange Is Nothing Then
        Dim ligneMax As ListRow
        For Each ligneMax In tblAteliers.ListRows
            If IsNumeric(ligneMax.Range.Cells(1, 1).Value) Then
                If CLng(ligneMax.Range.Cells(1, 1).Value) >= nouvelID Then
                    nouvelID = CLng(ligneMax.Range.Cells(1, 1).Value) + 1
                End If
            End If
        Next ligneMax
    End If
    
    ' Ajout d'une nouvelle ligne dans le tableau
    Set nouvelleLigne = tblAteliers.ListRows.Add
    
    With nouvelleLigne.Range
        .Cells(1, 1).Value = nouvelID           ' ID_Atelier
        .Cells(1, 2).Value = Trim(nom)          ' Nom
        .Cells(1, 3).Value = nouvelleDate        ' Date
        .Cells(1, 3).NumberFormat = "DD/MM/YYYY" ' Formatage de la date
        .Cells(1, 4).Value = heureDebut          ' Heure_Debut
        .Cells(1, 5).Value = heureFin            ' Heure_Fin
        .Cells(1, 6).Value = dureeFormatee       ' Duree
        .Cells(1, 7).Value = Trim(theme)         ' Theme
        .Cells(1, 8).Value = 0                   ' Nb_Participants (initialisÃ© Ã  0)
        .Cells(1, 9).Value = 0                   ' Nb_Participants_Pro (initialisÃ© Ã  0)
        .Cells(1, 10).Value = Trim(animepar)     ' Anime_Par
    End With
    
    ' ReprotÃ©ger la feuille
    wsAteliers.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True
    
    ' Mettre Ã  jour les statistiques
    Call MettreAJourStats
    
    EnregistrerAtelier = True
    Exit Function
    
ErrDate:
    MsgBox "Format de date invalide. Utilisez le format JJ/MM/AAAA." & vbCrLf & _
           "Exemple : 25/03/2025", vbExclamation, "Date invalide"
    Exit Function

ErrHeure:
    MsgBox "Format d'heure invalide. Utilisez le format HH:MM." & vbCrLf & _
           "Exemple : 09:30", vbExclamation, "Heure invalide"
    Exit Function

ErrFeuille:
    MsgBox "Erreur d'accÃ¨s Ã  la feuille ATELIERS ou au tableau TblAteliers.", _
           vbCritical, "Erreur"
    Exit Function
End Function

' -----------------------------------------------------------------------------
' ObtenirListeThemes : Retourne le tableau des thèmes disponibles
' Lit les thèmes depuis la feuille CONFIG (colonne A, à partir de A2)
' Si la feuille CONFIG est introuvable ou vide, retourne des thèmes par défaut
' -----------------------------------------------------------------------------
Public Function ObtenirListeThemes() As String()
    Dim wsConfig As Worksheet
    Dim themes() As String
    Dim nb As Integer
    Dim i As Integer
    
    On Error GoTo ThemesParDefaut
    Set wsConfig = ThisWorkbook.Sheets("CONFIG")
    On Error GoTo 0
    
    nb = 0
    i = 2
    Do While wsConfig.Cells(i, 1).Value <> ""
        nb = nb + 1
        i = i + 1
    Loop
    
    If nb = 0 Then GoTo ThemesParDefaut
    
    ReDim themes(0 To nb - 1)
    For i = 0 To nb - 1
        themes(i) = CStr(wsConfig.Cells(i + 2, 1).Value)
    Next i
    
    ObtenirListeThemes = themes
    Exit Function
    
ThemesParDefaut:
    ObtenirListeThemes = Split(THEMES_ATELIERS, ",")
End Function

' -----------------------------------------------------------------------------
' RecalculerNbParticipants : Recalcule et met Ã  jour Nb_Participants et
' Nb_Participants_Pro pour un atelier donnÃ©
' ParamÃ¨tre :
'   idAtelier : L'ID de l'atelier Ã  recalculer
' -----------------------------------------------------------------------------
Public Sub RecalculerNbParticipants(idAtelier As Long)
    Dim wsAteliers As Worksheet
    Dim wsPresences As Worksheet
    Dim tblAteliers As ListObject
    Dim tblPresences As ListObject
    Dim ligneAtelier As ListRow
    Dim lignePresence As ListRow
    Dim nbTotal As Long
    Dim nbPro As Long
    
    On Error GoTo ErrRecalcul
    Set wsAteliers = ThisWorkbook.Sheets("ATELIERS")
    Set wsPresences = ThisWorkbook.Sheets("PRESENCES")
    Set tblAteliers = wsAteliers.ListObjects("TblAteliers")
    Set tblPresences = wsPresences.ListObjects("TblPresences")
    On Error GoTo 0
    
    ' Compter les prÃ©sences pour cet atelier
    nbTotal = 0
    nbPro = 0
    
    If Not tblPresences.DataBodyRange Is Nothing Then
        For Each lignePresence In tblPresences.ListRows
            ' Colonne 2 = ID_Atelier
            If CLng(lignePresence.Range.Cells(1, 2).Value) = idAtelier Then
                nbTotal = nbTotal + 1
                ' Colonne 6 = Statut_Participant
                If lignePresence.Range.Cells(1, 6).Value = "Lancé" Then
                    nbPro = nbPro + 1
                End If
            End If
        Next lignePresence
    End If
    
    ' Mettre Ã  jour la ligne de l'atelier dans TblAteliers
    wsAteliers.Unprotect Password:=MOT_DE_PASSE
    
    If Not tblAteliers.DataBodyRange Is Nothing Then
        For Each ligneAtelier In tblAteliers.ListRows
            If CLng(ligneAtelier.Range.Cells(1, 1).Value) = idAtelier Then
                ligneAtelier.Range.Cells(1, 8).Value = nbTotal  ' Nb_Participants
                ligneAtelier.Range.Cells(1, 9).Value = nbPro    ' Nb_Participants_Pro
                Exit For
            End If
        Next ligneAtelier
    End If
    
    wsAteliers.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True
    
    Exit Sub
    
ErrRecalcul:
    wsAteliers.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub
