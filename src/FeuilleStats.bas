' =============================================================================
' Module : FeuilleStats (Microsoft Excel Objects â€” feuille STATS)
' Description : DÃ©tecte le changement du sÃ©lecteur d'annÃ©e (cellule B1)
'               et recalcule les statistiques
' =============================================================================

Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("B1")) Is Nothing Then
        If IsNumeric(Me.Range("B1").Value) And Me.Range("B1").Value > 2000 Then
            Application.EnableEvents = False   ' ? STOPPE la boucle
            On Error GoTo Nettoyage
            
            ThisWorkbook.Sheets("ACCUEIL").Range("B2").Value = Me.Range("B1").Value
            Call RecalculerStatsAnnee(CInt(Me.Range("B1").Value))
            Call MettreAJourAccueil

Nettoyage:
            Application.EnableEvents = True    ' ? TOUJOURS réactiver
        End If
    End If
End Sub
