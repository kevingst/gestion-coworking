' =============================================================================
' Module : ModuleRecalcul
' Description : Utilitaires de recalcul en masse des compteurs
'
' Ce module expose des Sub directement exécutables depuis la boîte de dialogue
' Macros (Alt+F8) ou depuis un bouton sur une feuille Excel.
' =============================================================================

' -----------------------------------------------------------------------------
' RecalculerTousLesParticipants : Recalcule la colonne Nb_Ateliers_Participes
' pour TOUS les participants présents dans TblParticipants.
'
' Utilisation :
'   - Ouvrir la boîte de dialogue Macros (Alt+F8)
'   - Sélectionner "RecalculerTousLesParticipants"
'   - Cliquer sur Exécuter
'
' Quand l'utiliser :
'   - Après un import en masse de participants ou de présences
'   - Après une modification manuelle directe dans les feuilles
'   - Pour corriger des compteurs devenus incohérents
' -----------------------------------------------------------------------------
Public Sub RecalculerTousLesParticipants()
    Dim wsParticipants As Worksheet
    Dim tblParticipants As ListObject
    Dim ligne As ListRow
    Dim nb As Long
    Dim nbErreurs As Long
    
    nb = 0
    nbErreurs = 0
    
    ' Accès à la feuille et au tableau
    On Error GoTo ErrFeuille
    Set wsParticipants = ThisWorkbook.Sheets("PARTICIPANTS")
    Set tblParticipants = wsParticipants.ListObjects("TblParticipants")
    On Error GoTo 0
    
    ' Vérifier qu'il y a des données
    If tblParticipants.DataBodyRange Is Nothing Then
        MsgBox "Aucun participant trouvé dans TblParticipants.", _
               vbInformation, "Recalcul"
        Exit Sub
    End If
    
    ' Parcourir chaque participant et recalculer son compteur
    For Each ligne In tblParticipants.ListRows
        If IsNumeric(ligne.Range.Cells(1, 1).Value) Then
            Dim idPart As Long
            idPart = CLng(ligne.Range.Cells(1, 1).Value)
            
            On Error GoTo ErrRecalcul
            Call RecalculerNbAteliers(idPart)
            On Error GoTo 0
            
            nb = nb + 1
        End If
    Next ligne
    
    MsgBox "Recalcul terminé." & vbCrLf & _
           nb & " participant(s) mis à jour.", _
           vbInformation, "Recalcul terminé"
    Exit Sub

ErrFeuille:
    MsgBox "Erreur : impossible d'accéder à la feuille PARTICIPANTS ou au tableau TblParticipants." & vbCrLf & _
           "Vérifiez que la feuille et le tableau existent.", _
           vbCritical, "Erreur d'accès"
    Exit Sub

ErrRecalcul:
    nbErreurs = nbErreurs + 1
    Resume Next
End Sub
