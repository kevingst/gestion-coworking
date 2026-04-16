' =============================================================================
' Module : ModulePresences
' Description : Gestion des prÃ©sences â€” enregistrement et recalcul
' =============================================================================

' -----------------------------------------------------------------------------
' EnregistrerPresences : Enregistre les prÃ©sences sÃ©lectionnÃ©es pour un atelier
' ParamÃ¨tres :
'   idAtelier        : L'ID de l'atelier
'   idsParticipants  : Tableau des IDs des participants sÃ©lectionnÃ©s
' Retourne True si succÃ¨s, False si erreur
' -----------------------------------------------------------------------------
Public Function EnregistrerPresences(idAtelier As Long, idsParticipants() As Long) As Boolean
    Dim wsPresences As Worksheet
    Dim wsParticipants As Worksheet
    Dim tblPresences As ListObject
    Dim tblParticipants As ListObject
    Dim nouvelleLigne As ListRow
    Dim nouvelID As Long
    Dim i As Integer
    
    EnregistrerPresences = False
    
    If Not IsArray(idsParticipants) Then
        MsgBox "Aucun participant sÃ©lectionnÃ©.", vbExclamation, "SÃ©lection vide"
        Exit Function
    End If
    
    ' AccÃ¨s aux feuilles et tableaux
    On Error GoTo ErrFeuille
    Set wsPresences = ThisWorkbook.Sheets("PRESENCES")
    Set wsParticipants = ThisWorkbook.Sheets("PARTICIPANTS")
    Set tblPresences = wsPresences.ListObjects("TblPresences")
    Set tblParticipants = wsParticipants.ListObjects("TblParticipants")
    On Error GoTo 0
    
    ' DÃ©sprotÃ©ger la feuille PRESENCES
    wsPresences.Unprotect Password:=MOT_DE_PASSE
    
    ' Calcul du prochain ID de prÃ©sence
    nouvelID = 1
    If Not tblPresences.DataBodyRange Is Nothing Then
        Dim ligneExist As ListRow
        For Each ligneExist In tblPresences.ListRows
            If IsNumeric(ligneExist.Range.Cells(1, 1).Value) Then
                If CLng(ligneExist.Range.Cells(1, 1).Value) >= nouvelID Then
                    nouvelID = CLng(ligneExist.Range.Cells(1, 1).Value) + 1
                End If
            End If
        Next ligneExist
    End If
    
    ' Pour chaque participant sÃ©lectionnÃ©
    For i = 0 To UBound(idsParticipants)
        Dim idPart As Long
        idPart = idsParticipants(i)
        
        ' VÃ©rifier que ce participant n'est pas dÃ©jÃ  enregistrÃ© pour cet atelier
        If Not EstDejaPresent(tblPresences, idAtelier, idPart) Then
            ' Rechercher les infos du participant
            Dim nomPart As String
            Dim prenomPart As String
            Dim statutPart As String
            Call ObtenirInfosParticipant(tblParticipants, idPart, nomPart, prenomPart, statutPart)
            
            ' Ajouter la ligne de prÃ©sence
            Set nouvelleLigne = tblPresences.ListRows.Add
            With nouvelleLigne.Range
                .Cells(1, 1).Value = nouvelID    ' ID_Presence
                .Cells(1, 2).Value = idAtelier   ' ID_Atelier
                .Cells(1, 3).Value = idPart      ' ID_Participant
                .Cells(1, 4).Value = nomPart     ' Nom_Participant
                .Cells(1, 5).Value = prenomPart  ' Prenom_Participant
                .Cells(1, 6).Value = statutPart  ' Statut_Participant
            End With
            
            nouvelID = nouvelID + 1
        End If
    Next i
    
    ' ReprotÃ©ger la feuille PRESENCES
    wsPresences.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True
    
    ' Recalculer les compteurs de l'atelier
    Call RecalculerNbParticipants(idAtelier)
    
    ' Mettre Ã  jour les statistiques globales
    Call MettreAJourStats
    
    EnregistrerPresences = True
    Exit Function
    
ErrFeuille:
    wsPresences.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True
    MsgBox "Erreur d'accÃ¨s aux feuilles. VÃ©rifiez que PRESENCES et PARTICIPANTS existent.", _
           vbCritical, "Erreur"
    Exit Function
End Function

' -----------------------------------------------------------------------------
' EstDejaPresent : VÃ©rifie si un participant est dÃ©jÃ  enregistrÃ© pour un atelier
' ParamÃ¨tres :
'   tblPresences  : Le tableau de prÃ©sences
'   idAtelier     : L'ID de l'atelier
'   idParticipant : L'ID du participant
' Retourne True si dÃ©jÃ  prÃ©sent
' -----------------------------------------------------------------------------
Public Function EstDejaPresent(tblPresences As ListObject, idAtelier As Long, _
                                idParticipant As Long) As Boolean
    Dim ligne As ListRow
    EstDejaPresent = False
    
    If tblPresences.DataBodyRange Is Nothing Then Exit Function
    
    For Each ligne In tblPresences.ListRows
        If IsNumeric(ligne.Range.Cells(1, 2).Value) And _
           IsNumeric(ligne.Range.Cells(1, 3).Value) Then
            If CLng(ligne.Range.Cells(1, 2).Value) = idAtelier And _
               CLng(ligne.Range.Cells(1, 3).Value) = idParticipant Then
                EstDejaPresent = True
                Exit Function
            End If
        End If
    Next ligne
End Function

' -----------------------------------------------------------------------------
' ObtenirInfosParticipant : RÃ©cupÃ¨re les informations d'un participant par son ID
' ParamÃ¨tres :
'   tblParticipants : Le tableau des participants
'   idParticipant   : L'ID recherchÃ©
'   nom             : (sortie) Le nom du participant
'   prenom          : (sortie) Le prÃ©nom du participant
'   statut          : (sortie) Le statut du participant
' -----------------------------------------------------------------------------
Private Sub ObtenirInfosParticipant(tblParticipants As ListObject, idParticipant As Long, _
                                     ByRef nom As String, ByRef prenom As String, _
                                     ByRef statut As String)
    Dim ligne As ListRow
    nom = ""
    prenom = ""
    statut = ""
    
    If tblParticipants.DataBodyRange Is Nothing Then Exit Sub
    
    For Each ligne In tblParticipants.ListRows
        If IsNumeric(ligne.Range.Cells(1, 1).Value) Then
            If CLng(ligne.Range.Cells(1, 1).Value) = idParticipant Then
                nom = CStr(ligne.Range.Cells(1, 2).Value)
                prenom = CStr(ligne.Range.Cells(1, 3).Value)
                statut = CStr(ligne.Range.Cells(1, 4).Value)
                Exit Sub
            End If
        End If
    Next ligne
End Sub

' -----------------------------------------------------------------------------
' ObtenirPresencesAtelier : Retourne la liste des IDs participants prÃ©sents
' pour un atelier donnÃ©
' ParamÃ¨tre :
'   idAtelier : L'ID de l'atelier
' Retourne un tableau d'IDs (Long) ou un tableau vide
' -----------------------------------------------------------------------------
Public Function ObtenirPresencesAtelier(idAtelier As Long) As Long()
    Dim wsPresences As Worksheet
    Dim tblPresences As ListObject
    Dim lignePresence As ListRow
    Dim ids() As Long
    Dim nb As Integer
    
    nb = 0
    ReDim ids(0)
    
    On Error GoTo ErrObtenirPresences
    Set wsPresences = ThisWorkbook.Sheets("PRESENCES")
    Set tblPresences = wsPresences.ListObjects("TblPresences")
    On Error GoTo 0
    
    If tblPresences.DataBodyRange Is Nothing Then
        ObtenirPresencesAtelier = ids
        Exit Function
    End If
    
    ' Compter d'abord
    For Each lignePresence In tblPresences.ListRows
        If IsNumeric(lignePresence.Range.Cells(1, 2).Value) Then
            If CLng(lignePresence.Range.Cells(1, 2).Value) = idAtelier Then
                nb = nb + 1
            End If
        End If
    Next lignePresence
    
    If nb = 0 Then
        ObtenirPresencesAtelier = ids
        Exit Function
    End If
    
    ReDim ids(0 To nb - 1)
    Dim i As Integer
    i = 0
    
    ' Remplir le tableau
    For Each lignePresence In tblPresences.ListRows
        If IsNumeric(lignePresence.Range.Cells(1, 2).Value) Then
            If CLng(lignePresence.Range.Cells(1, 2).Value) = idAtelier Then
                ids(i) = CLng(lignePresence.Range.Cells(1, 3).Value)
                i = i + 1
            End If
        End If
    Next lignePresence
    
    ObtenirPresencesAtelier = ids
    Exit Function
    
ErrObtenirPresences:
    ObtenirPresencesAtelier = ids
End Function
