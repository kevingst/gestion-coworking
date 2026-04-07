Attribute VB_Name = "ModuleStats"
' =============================================================================
' Module : ModuleStats
' Description : Calcul et mise à jour des statistiques de l'année en cours
'               affichées sur la feuille ACCUEIL
' =============================================================================

' -----------------------------------------------------------------------------
' MettreAJourStats : Calcule et affiche les statistiques de l'année en cours
' Appelée à l'ouverture du fichier et après chaque enregistrement
' -----------------------------------------------------------------------------
Public Sub MettreAJourStats()
    Dim wsAccueil As Worksheet
    Dim wsAteliers As Worksheet
    Dim tblAteliers As ListObject
    Dim anneeEnCours As Integer
    
    Dim nbAteliers As Long
    Dim dureeTotale As Double
    Dim nbParticipants As Long
    Dim nbParticipantsPro As Long
    
    Dim ligneAtelier As ListRow
    Dim dateAtelier As Date
    Dim dureeAtelier As String
    Dim heuresAtelier As Integer
    Dim minutesAtelier As Integer
    
    ' Récupérer les feuilles
    On Error GoTo ErrStats
    Set wsAccueil = ThisWorkbook.Sheets("ACCUEIL")
    Set wsAteliers = ThisWorkbook.Sheets("ATELIERS")
    Set tblAteliers = wsAteliers.ListObjects("TblAteliers")
    On Error GoTo 0
    
    ' Année en cours
    anneeEnCours = Year(Now())
    
    ' Initialiser les compteurs
    nbAteliers = 0
    dureeTotale = 0 ' En minutes
    nbParticipants = 0
    nbParticipantsPro = 0
    
    ' Parcourir chaque atelier
    If Not tblAteliers.DataBodyRange Is Nothing Then
        For Each ligneAtelier In tblAteliers.ListRows
            ' Vérifier que la date est dans l'année en cours
            On Error Resume Next
            dateAtelier = CDate(ligneAtelier.Range.Cells(1, 3).Value) ' Colonne Date
            On Error GoTo 0
            
            If Year(dateAtelier) = anneeEnCours And dateAtelier <> 0 Then
                ' Incrémenter le nombre d'ateliers
                nbAteliers = nbAteliers + 1
                
                ' Calculer la durée (colonne Duree = colonne 6)
                dureeAtelier = CStr(ligneAtelier.Range.Cells(1, 6).Value)
                If InStr(dureeAtelier, ":") > 0 Then
                    heuresAtelier = CInt(Split(dureeAtelier, ":")(0))
                    minutesAtelier = CInt(Split(dureeAtelier, ":")(1))
                    dureeTotale = dureeTotale + (heuresAtelier * 60) + minutesAtelier
                End If
                
                ' Additionner les participants (colonne Nb_Participants = colonne 8)
                If IsNumeric(ligneAtelier.Range.Cells(1, 8).Value) Then
                    nbParticipants = nbParticipants + CLng(ligneAtelier.Range.Cells(1, 8).Value)
                End If
                
                ' Additionner les participants pro (colonne Nb_Participants_Pro = colonne 9)
                If IsNumeric(ligneAtelier.Range.Cells(1, 9).Value) Then
                    nbParticipantsPro = nbParticipantsPro + CLng(ligneAtelier.Range.Cells(1, 9).Value)
                End If
            End If
        Next ligneAtelier
    End If
    
    ' Convertir la durée totale (minutes) en format HH:MM
    Dim heuresTotal As Long
    Dim minutesTotal As Long
    heuresTotal = dureeTotale \ 60
    minutesTotal = dureeTotale Mod 60
    
    Dim dureeTotaleFormatee As String
    dureeTotaleFormatee = Format(heuresTotal, "00") & ":" & Format(minutesTotal, "00")
    
    ' Écrire les statistiques dans la feuille ACCUEIL
    ' (Cellules C5 à C8 selon le SETUP.md)
    wsAccueil.Cells(5, 3).Value = nbAteliers
    wsAccueil.Cells(6, 3).Value = dureeTotaleFormatee
    wsAccueil.Cells(7, 3).Value = nbParticipants
    wsAccueil.Cells(8, 3).Value = nbParticipantsPro
    
    Exit Sub
    
ErrStats:
    ' En cas d'erreur, ne pas bloquer l'utilisateur
    ' Les cellules restent avec leurs valeurs précédentes
    On Error GoTo 0
End Sub
