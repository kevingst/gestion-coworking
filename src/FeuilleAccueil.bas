Attribute VB_Name = "FeuilleAccueil"
' =============================================================================
' Module : FeuilleAccueil (Microsoft Excel Objects — feuille ACCUEIL)
' Description : Détecte les changements des sélecteurs de mois (B1) et
'               d'année (B2) sur la feuille ACCUEIL
' =============================================================================

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Changement du sélecteur de mois (B1) ou d'année (B2)
    If Not Intersect(Target, Me.Range("B1:B2")) Is Nothing Then
        ' Synchroniser STATS!B1 si l'année change
        If Not Intersect(Target, Me.Range("B2")) Is Nothing Then
            If IsNumeric(Me.Range("B2").Value) And Me.Range("B2").Value > 2000 Then
                On Error Resume Next
                ThisWorkbook.Sheets("STATS").Range("B1").Value = Me.Range("B2").Value
                On Error GoTo 0
                Call RecalculerStatsAnnee(CInt(Me.Range("B2").Value))
            End If
        End If
        Call MettreAJourAccueil
    End If
End Sub
