' =============================================================================
' Module : FeuilleAccueil (Microsoft Excel Objects â€” feuille ACCUEIL)
' Description : DÃ©tecte les changements des sÃ©lecteurs de mois (B1) et
'               d'annÃ©e (B2) sur la feuille ACCUEIL
' =============================================================================

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Changement du sÃ©lecteur de mois (B1) ou d'annÃ©e (B2)
    If Not Intersect(Target, Me.Range("B1:B2")) Is Nothing Then
        Application.EnableEvents = False   ' ? STOPPE la boucle
        On Error GoTo Nettoyage
        ' Synchroniser STATS!B1 si l'annÃ©e change
        If Not Intersect(Target, Me.Range("B2")) Is Nothing Then
            If IsNumeric(Me.Range("B2").Value) And Me.Range("B2").Value > 2000 Then
                On Error Resume Next
                ThisWorkbook.Sheets("STATS").Range("B1").Value = Me.Range("B2").Value
                On Error GoTo 0
                Call RecalculerStatsAnnee(CInt(Me.Range("B2").Value))
            End If
        End If
        Call MettreAJourAccueil
Nettoyage:
        Application.EnableEvents = True    ' ? TOUJOURS réactiver
    End If
End Sub
