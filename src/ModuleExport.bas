Attribute VB_Name = "ModuleExport"
' =============================================================================
' Module : ModuleExport
' Description : Export des participants newsletter vers un fichier CSV Brevo
' =============================================================================

' -----------------------------------------------------------------------------
' ExporterParticipantsBrevo : Exporte les participants (Newsletter=Oui) en CSV
' Format Brevo : EMAIL;PRENOM;NOM;STATUT
' -----------------------------------------------------------------------------
Public Sub ExporterParticipantsBrevo()
    Dim wsParticipants As Worksheet
    Dim tblParticipants As ListObject
    Dim ligneParticipant As ListRow
    Dim cheminFichier As String
    Dim numFichier As Integer
    Dim nbExportes As Integer
    Dim newsletter As String
    Dim mail As String
    Dim prenom As String
    Dim nom As String
    Dim statut As String
    
    ' Demander le chemin de sauvegarde
    cheminFichier = Application.GetSaveAsFilename( _
        InitialFileName:="participants_brevo_" & Format(Now(), "yyyymmdd") & ".csv", _
        FileFilter:="Fichiers CSV (*.csv), *.csv", _
        Title:="Exporter les participants pour Brevo")
    
    If cheminFichier = "False" Or cheminFichier = "" Then Exit Sub  ' Annulé
    
    On Error GoTo ErrExport
    Set wsParticipants = ThisWorkbook.Sheets("PARTICIPANTS")
    Set tblParticipants = wsParticipants.ListObjects("TblParticipants")
    On Error GoTo 0
    
    numFichier = FreeFile
    Open cheminFichier For Output As #numFichier
    
    ' En-tête CSV compatible Brevo
    Print #numFichier, "EMAIL;PRENOM;NOM;STATUT"
    
    nbExportes = 0
    
    If Not tblParticipants.DataBodyRange Is Nothing Then
        For Each ligneParticipant In tblParticipants.ListRows
            ' Colonne 12 = Newsletter
            newsletter = ""
            On Error Resume Next
            newsletter = CStr(ligneParticipant.Range.Cells(1, 12).Value)
            On Error GoTo 0
            
            If LCase(Trim(newsletter)) = "oui" Then
                mail = Trim(CStr(ligneParticipant.Range.Cells(1, 8).Value))   ' Colonne 8 = Mail
                prenom = Trim(CStr(ligneParticipant.Range.Cells(1, 3).Value)) ' Colonne 3 = Prenom
                nom = Trim(CStr(ligneParticipant.Range.Cells(1, 2).Value))    ' Colonne 2 = Nom
                statut = Trim(CStr(ligneParticipant.Range.Cells(1, 4).Value)) ' Colonne 4 = Statut
                
                ' N'exporter que si l'email est renseigné
                If mail <> "" Then
                    Print #numFichier, mail & ";" & prenom & ";" & nom & ";" & statut
                    nbExportes = nbExportes + 1
                End If
            End If
        Next ligneParticipant
    End If
    
    Close #numFichier
    
    MsgBox nbExportes & " participant(s) exporté(s) avec succès !" & vbCrLf & _
           "Fichier : " & cheminFichier, vbInformation, "Export réussi"
    Exit Sub
    
ErrExport:
    On Error Resume Next
    Close #numFichier
    On Error GoTo 0
    MsgBox "Erreur lors de l'export CSV." & vbCrLf & _
           "Erreur " & Err.Number & " : " & Err.Description, vbCritical, "Erreur"
End Sub
