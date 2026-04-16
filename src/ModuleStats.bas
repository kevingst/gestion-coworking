' =============================================================================
' Module : ModuleStats
' Description : Calcul et mise à jour des statistiques mensuelles et annuelles
'               Feuille STATS (tableau mensuel) + feuille ACCUEIL (résumé)
' =============================================================================

' Noms des mois pour les sélecteurs et le tableau STATS
Public Const NOMS_MOIS As String = "Janvier,Février,Mars,Avril,Mai,Juin,Juillet,Août,Septembre,Octobre,Novembre,Décembre"

' -----------------------------------------------------------------------------
' MettreAJourStats : Point d'entrée unique appelé depuis partout dans le projet
' Lit l'année depuis STATS!B1 puis recalcule STATS et ACCUEIL
' -----------------------------------------------------------------------------
Public Sub MettreAJourStats()
    Dim annee As Integer
    
    ' Lire l'année depuis STATS!B1
    On Error Resume Next
    Dim valAnnee As Variant
    valAnnee = ThisWorkbook.Sheets("STATS").Range("B1").Value
    On Error GoTo 0
    
    If IsNumeric(valAnnee) And valAnnee > 2000 Then
        annee = CInt(valAnnee)
    Else
        annee = Year(Now())
    End If
    
    Call RecalculerStatsAnnee(annee)
    Call MettreAJourAccueil
End Sub

' -----------------------------------------------------------------------------
' RecalculerStatsAnnee : Remplit la feuille STATS avec les données mensuelles
' Paramètre :
'   annee : L'année à calculer
' -----------------------------------------------------------------------------
Public Sub RecalculerStatsAnnee(annee As Integer)
    Dim wsStats As Worksheet
    Dim wsAteliers As Worksheet
    Dim tblAteliers As ListObject
    
    On Error GoTo ErrRecalcul
    Set wsStats = ThisWorkbook.Sheets("STATS")
    Set wsAteliers = ThisWorkbook.Sheets("ATELIERS")
    Set tblAteliers = wsAteliers.ListObjects("TblAteliers")
    On Error GoTo 0
    
    ' Déprotéger la feuille STATS
    wsStats.Unprotect Password:=MOT_DE_PASSE
    
    ' Titre en A1 (ne pas écraser si déjà présent)
    If wsStats.Cells(1, 1).Value = "" Then
        wsStats.Cells(1, 1).Value = "Statistiques"
    End If
    
    ' Sélecteur année en B1 (ne pas écraser la valeur existante)
    If wsStats.Cells(1, 2).Value = "" Then
        wsStats.Cells(1, 2).Value = annee
    End If
    
    ' En-têtes en ligne 2
    wsStats.Cells(2, 1).Value = "Mois"
    wsStats.Cells(2, 2).Value = "Nb Ateliers"
    wsStats.Cells(2, 3).Value = "Durée Totale"
    wsStats.Cells(2, 4).Value = "Nb Participants"
    wsStats.Cells(2, 5).Value = "Nb Participants Pro"
    
    Dim nomsMois() As String
    nomsMois = Split(NOMS_MOIS, ",")
    
    Dim totalAteliers As Long
    Dim totalDureeMin As Long
    Dim totalParticipants As Long
    Dim totalParticipantsPro As Long
    totalAteliers = 0
    totalDureeMin = 0
    totalParticipants = 0
    totalParticipantsPro = 0
    
    Dim mois As Integer
    For mois = 1 To 12
        Dim nbAteliers As Long
        Dim dureeTotaleMin As Long
        Dim nbParticipants As Long
        Dim nbParticipantsPro As Long
        nbAteliers = 0
        dureeTotaleMin = 0
        nbParticipants = 0
        nbParticipantsPro = 0
        
        ' Parcourir chaque atelier
        If Not tblAteliers.DataBodyRange Is Nothing Then
            Dim ligneAtelier As ListRow
            For Each ligneAtelier In tblAteliers.ListRows
                Dim dateAtelier As Date
                On Error Resume Next
                dateAtelier = CDate(ligneAtelier.Range.Cells(1, 3).Value) ' Colonne Date
                On Error GoTo 0
                
                If dateAtelier <> 0 And Year(dateAtelier) = annee And Month(dateAtelier) = mois Then
                    nbAteliers = nbAteliers + 1
                    
                    ' Calculer la durée (colonne Duree = colonne 6)
                    Dim dureeRaw As Variant
                    dureeRaw = ligneAtelier.Range.Cells(1, 6).Value
                    Dim dureeMin As Long
                    dureeMin = 0
                    
                    If Not IsEmpty(dureeRaw) And IsNumeric(dureeRaw) Then
                        ' Valeur décimale Excel (fraction de journée) ? convertir en minutes
                        dureeMin = CLng(CDbl(dureeRaw) * 24 * 60)
                    ElseIf InStr(CStr(dureeRaw), ":") > 0 Then
                        ' Texte au format HH:MM
                        Dim parties() As String
                        parties = Split(CStr(dureeRaw), ":")
                        dureeMin = CInt(parties(0)) * 60 + CInt(parties(1))
                    End If
                    
                    dureeTotaleMin = dureeTotaleMin + dureeMin
                    
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
        
        ' Formater la durée du mois en HH:MM
        Dim heuresMois As Long
        Dim minutesMois As Long
        heuresMois = dureeTotaleMin \ 60
        minutesMois = dureeTotaleMin Mod 60
        Dim dureeMoisFormatee As String
        dureeMoisFormatee = Format(heuresMois, "00") & ":" & Format(minutesMois, "00")
        
        ' Écrire la ligne du mois (lignes 3 à 14)
        Dim ligne As Integer
        ligne = mois + 2
        wsStats.Cells(ligne, 1).Value = nomsMois(mois - 1)
        wsStats.Cells(ligne, 2).Value = nbAteliers
        wsStats.Cells(ligne, 3).Value = dureeMoisFormatee
        wsStats.Cells(ligne, 4).Value = nbParticipants
        wsStats.Cells(ligne, 5).Value = nbParticipantsPro
        
        ' Accumuler les totaux annuels
        totalAteliers = totalAteliers + nbAteliers
        totalDureeMin = totalDureeMin + dureeTotaleMin
        totalParticipants = totalParticipants + nbParticipants
        totalParticipantsPro = totalParticipantsPro + nbParticipantsPro
    Next mois
    
    ' Ligne 15 : TOTAL ANNÉE
    Dim heuresAnnee As Long
    Dim minutesAnnee As Long
    heuresAnnee = totalDureeMin \ 60
    minutesAnnee = totalDureeMin Mod 60
    Dim dureeTotaleFormatee As String
    dureeTotaleFormatee = Format(heuresAnnee, "00") & ":" & Format(minutesAnnee, "00")
    
    wsStats.Cells(15, 1).Value = "TOTAL ANNÉE"
    wsStats.Cells(15, 2).Value = totalAteliers
    wsStats.Cells(15, 3).Value = dureeTotaleFormatee
    wsStats.Cells(15, 4).Value = totalParticipants
    wsStats.Cells(15, 5).Value = totalParticipantsPro
    
    ' Reprotéger la feuille STATS
    wsStats.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True
    
    Exit Sub
    
ErrRecalcul:
    On Error Resume Next
    wsStats.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------
' MettreAJourAccueil : Lit les sélecteurs de ACCUEIL et met à jour les stats
' Utilise les données déjà calculées dans la feuille STATS
' -----------------------------------------------------------------------------
Public Sub MettreAJourAccueil()
    On Error GoTo ErrAccueil
    
    Dim wsAccueil As Worksheet
    Dim wsStats As Worksheet
    Set wsAccueil = ThisWorkbook.Sheets("ACCUEIL")
    Set wsStats = ThisWorkbook.Sheets("STATS")
    
    ' Lire le mois depuis ACCUEIL!B1 (texte "Janvier"..."Décembre")
    Dim nomMois As String
    nomMois = CStr(wsAccueil.Range("B1").Value)
    Dim numMois As Integer
    numMois = NumeroMois(nomMois)
    If numMois = 0 Then numMois = Month(Now())
    
    ' Lire l'année depuis ACCUEIL!B2
    Dim anneeAccueil As Integer
    Dim valAnneeAccueil As Variant
    valAnneeAccueil = wsAccueil.Range("B2").Value
    If IsNumeric(valAnneeAccueil) And valAnneeAccueil > 2000 Then
        anneeAccueil = CInt(valAnneeAccueil)
    Else
        anneeAccueil = Year(Now())
    End If
    
    ' Synchroniser STATS!B1 si l'année diffère
    Dim anneeStats As Integer
    Dim valAnneeStats As Variant
    valAnneeStats = wsStats.Range("B1").Value
    If IsNumeric(valAnneeStats) And valAnneeStats > 2000 Then
        anneeStats = CInt(valAnneeStats)
    Else
        anneeStats = 0
    End If
    
    If anneeStats <> anneeAccueil Then
        wsStats.Unprotect Password:=MOT_DE_PASSE
        wsStats.Range("B1").Value = anneeAccueil
        wsStats.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True
        Call RecalculerStatsAnnee(anneeAccueil)
    End If
    
    ' Lire les données du mois dans STATS (ligne = numMois + 2)
    Dim ligneMois As Integer
    ligneMois = numMois + 2
    
    ' Écrire les stats du mois sur ACCUEIL
    wsAccueil.Cells(5, 3).Value = wsStats.Cells(ligneMois, 2).Value  ' Nb ateliers
    wsAccueil.Cells(6, 3).Value = wsStats.Cells(ligneMois, 3).Value  ' Durée totale
    wsAccueil.Cells(7, 3).Value = wsStats.Cells(ligneMois, 4).Value  ' Nb participants
    wsAccueil.Cells(8, 3).Value = wsStats.Cells(ligneMois, 5).Value  ' Nb participants pro
    
    ' Écrire les totaux annuels sur ACCUEIL (ligne 15 de STATS = TOTAL ANNÉE)
    wsAccueil.Cells(10, 3).Value = wsStats.Cells(15, 2).Value  ' Nb ateliers total
    wsAccueil.Cells(11, 3).Value = wsStats.Cells(15, 3).Value  ' Durée totale année
    wsAccueil.Cells(12, 3).Value = wsStats.Cells(15, 4).Value  ' Nb participants total
    wsAccueil.Cells(13, 3).Value = wsStats.Cells(15, 5).Value  ' Nb participants pro total
    
    ' --- Mise à jour de la plage source du Graphique 1 (mois sélectionné) ---
    ' Colonnes G:H utilisées comme source fixe pour GraphiqueMois
    ' La durée est convertie en minutes (entier) pour l'axe du graphique
    On Error Resume Next
    wsAccueil.Range("G1").Value = "Indicateur"
    wsAccueil.Range("H1").Value = "Valeur"
    wsAccueil.Range("G2").Value = "Nb Ateliers"
    wsAccueil.Range("G3").Value = "Dur" & Chr(233) & "e (min)"
    wsAccueil.Range("G4").Value = "Participants"
    wsAccueil.Range("G5").Value = "Participants Pro"
    
    ' Lire les valeurs du mois depuis STATS
    Dim ligneStats As Integer
    ligneStats = numMois + 2  ' Janvier=mois 1 ? ligne 3, etc.
    
    wsAccueil.Range("H2").Value = wsStats.Cells(ligneStats, 2).Value  ' Nb Ateliers
    
    ' Durée : convertir HH:MM en minutes pour le graphique
    Dim dureeStr As String
    dureeStr = CStr(wsStats.Cells(ligneStats, 3).Value)
    Dim dureeGraphMin As Long
    dureeGraphMin = 0
    If InStr(dureeStr, ":") > 0 Then
        Dim partiesGraph() As String
        partiesGraph = Split(dureeStr, ":")
        If UBound(partiesGraph) >= 1 Then
            dureeGraphMin = CLng(partiesGraph(0)) * 60 + CLng(partiesGraph(1))
        End If
    ElseIf IsNumeric(dureeStr) And dureeStr <> "" Then
        dureeGraphMin = CLng(CDbl(dureeStr) * 24 * 60)
    End If
    wsAccueil.Range("H3").Value = dureeGraphMin
    
    wsAccueil.Range("H4").Value = wsStats.Cells(ligneStats, 4).Value  ' Nb Participants
    wsAccueil.Range("H5").Value = wsStats.Cells(ligneStats, 5).Value  ' Nb Participants Pro
    On Error GoTo 0
    
    ' --- Mise à jour des titres des graphiques ---
    ' Les graphiques doivent être nommés "GraphiqueMois" et "GraphiqueAnnee" dans Excel
    Dim nomMoisAffiche As String
    Dim anneeAffichee As String
    nomMoisAffiche = nomMois   ' variable locale contenant le nom du mois ex: "Avril"
    anneeAffichee = CStr(anneeAccueil)  ' variable locale contenant l'année ex: "2026"
    
    On Error Resume Next
    Dim cht As ChartObject
    For Each cht In wsAccueil.ChartObjects
        If cht.Name = "GraphiqueMois" Then
            cht.Chart.HasTitle = True
            cht.Chart.ChartTitle.Text = nomMoisAffiche & " " & anneeAffichee
        End If
        If cht.Name = "GraphiqueAnnee" Then
            cht.Chart.HasTitle = True
            cht.Chart.ChartTitle.Text = "Bilan " & anneeAffichee
        End If
    Next cht
    On Error GoTo 0
    
    Exit Sub
    
ErrAccueil:
    ' Gestion silencieuse : ne pas bloquer si STATS n'existe pas encore
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------
' NumeroMois : Retourne le numéro (1-12) d'un mois à partir de son nom
' Retourne 0 si le nom n'est pas reconnu
' Paramètre :
'   nomMois : Le nom du mois (ex: "Janvier", "Février", etc.)
' -----------------------------------------------------------------------------
Public Function NumeroMois(nomMois As String) As Integer
    Dim moisArr() As String
    moisArr = Split(NOMS_MOIS, ",")
    
    Dim i As Integer
    For i = 0 To UBound(moisArr)
        If LCase(Trim(nomMois)) = LCase(Trim(moisArr(i))) Then
            NumeroMois = i + 1
            Exit Function
        End If
    Next i
    
    NumeroMois = 0
End Function
