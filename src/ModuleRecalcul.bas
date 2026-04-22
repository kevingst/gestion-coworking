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
'
' Note technique :
'   Les IDs sont d'abord copiés dans un tableau avant de lancer le recalcul.
'   Cela évite le crash causé par la modification de TblParticipants pendant
'   une itération For Each sur ce même tableau.
' -----------------------------------------------------------------------------
Public Sub RecalculerTousLesParticipants()
    Dim wsParticipants As Worksheet
    Dim tblParticipants As ListObject
    Dim ligne As ListRow
    Dim ids() As Long
    Dim nbTotal As Long
    Dim nbOk As Long
    Dim i As Long
    
    nbTotal = 0
    nbOk = 0
    
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
    
    ' --- Étape 1 : collecter tous les IDs dans un tableau ---
    ' On lit tous les IDs AVANT toute modification, pour éviter que la boucle
    ' For Each ne soit corrompue quand RecalculerNbAteliers modifie la table.
    ReDim ids(0 To tblParticipants.ListRows.Count - 1)
    For Each ligne In tblParticipants.ListRows
        If IsNumeric(ligne.Range.Cells(1, 1).Value) Then
            ids(nbTotal) = CLng(ligne.Range.Cells(1, 1).Value)
            nbTotal = nbTotal + 1
        End If
    Next ligne
    
    If nbTotal = 0 Then
        MsgBox "Aucun identifiant valide trouvé dans TblParticipants.", _
               vbInformation, "Recalcul"
        Exit Sub
    End If
    
    ' --- Étape 2 : recalculer chaque participant par son ID ---
    For i = 0 To nbTotal - 1
        On Error GoTo ErrRecalcul
        Call RecalculerNbAteliers(ids(i))
        On Error GoTo 0
        nbOk = nbOk + 1
    Next i
    
    MsgBox "Recalcul terminé." & vbCrLf & _
           nbOk & " participant(s) mis à jour.", _
           vbInformation, "Recalcul terminé"
    Exit Sub

ErrFeuille:
    MsgBox "Erreur : impossible d'accéder à la feuille PARTICIPANTS ou au tableau TblParticipants." & vbCrLf & _
           "Vérifiez que la feuille et le tableau existent.", _
           vbCritical, "Erreur d'accès"
    Exit Sub

ErrRecalcul:
    Resume Next
End Sub
