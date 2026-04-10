Attribute VB_Name = "FeuilleStats"
' =============================================================================
' Module : FeuilleStats (Microsoft Excel Objects — feuille STATS)
' Description : Détecte le changement du sélecteur d'année (cellule B1)
'               et recalcule les statistiques
' =============================================================================

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Détecter le changement de la cellule B1 (sélecteur d'année)
    If Not Intersect(Target, Me.Range("B1")) Is Nothing Then
        If IsNumeric(Me.Range("B1").Value) And Me.Range("B1").Value > 2000 Then
            ' Synchroniser le sélecteur d'année sur ACCUEIL
            On Error Resume Next
            ThisWorkbook.Sheets("ACCUEIL").Range("B2").Value = Me.Range("B1").Value
            On Error GoTo 0
            ' Recalculer les stats pour la nouvelle année
            Call RecalculerStatsAnnee(CInt(Me.Range("B1").Value))
            Call MettreAJourAccueil
        End If
    End If
End Sub
