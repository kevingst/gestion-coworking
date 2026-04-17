' =============================================================================
' Module : ThisWorkbook
' Description : Gestion des événements du classeur (ouverture, fermeture)
' =============================================================================


' -----------------------------------------------------------------------------
' Workbook_Open : Exécuté à l'ouverture du fichier
' -----------------------------------------------------------------------------
Private Sub Workbook_Open()
    ' Activer la feuille d'accueil au démarrage
    ThisWorkbook.Sheets("ACCUEIL").Activate
    
    ' Mettre à jour les statistiques de l'année en cours
    Call MettreAJourStats
    
    ' Informer l'utilisateur si les macros sont bien actives
    ' (Ce message ne s'affiche que si les macros sont activées, ce qui confirme leur bon fonctionnement)
    ' MsgBox "Bienvenue dans le gestionnaire d'ateliers de coworking !", vbInformation, "Gestion Coworking"
End Sub

' -----------------------------------------------------------------------------
' InitialiserFeuilles : Vérifie et initialise les tableaux structurés
' Appelée une seule fois lors de la première utilisation
' -----------------------------------------------------------------------------
Public Sub InitialiserFeuilles()
    Dim wsAteliers As Worksheet
    Dim wsParticipants As Worksheet
    Dim wsPresences As Worksheet
    Dim wsImport As Worksheet
    
    ' Récupérer les feuilles
    On Error GoTo ErrFeuille
    Set wsAteliers = ThisWorkbook.Sheets("ATELIERS")
    Set wsParticipants = ThisWorkbook.Sheets("PARTICIPANTS")
    Set wsPresences = ThisWorkbook.Sheets("PRESENCES")
    Set wsImport = ThisWorkbook.Sheets("IMPORT")
    On Error GoTo 0
    
    ' Vérifier que les tableaux structurés existent (sinon les créer)
    Call VerifierOuCreerTableau(wsAteliers, "TblAteliers", _
        Array("ID_Atelier", "Nom", "Date", "Heure_Debut", "Heure_Fin", "Duree", "Theme", "Nb_Participants", "Nb_Participants_Pro"))
    
    Call VerifierOuCreerTableau(wsParticipants, "TblParticipants", _
        Array("ID_Participant", "Nom", "Prenom", "Statut", "Date_Premier_Contact", "Nom_Entreprise", "Commune", "Code_Postal", "Mail", "Telephone", "Activite"))
    
    Call VerifierOuCreerTableau(wsPresences, "TblPresences", _
        Array("ID_Presence", "ID_Atelier", "ID_Participant", "Nom_Participant", "Prenom_Participant", "Statut_Participant"))
    
    ' Créer la feuille CONFIG si elle n'existe pas encore
    Dim wsConfig As Worksheet
    Dim configExiste As Boolean
    Dim ws As Worksheet
    configExiste = False
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "CONFIG" Then
            configExiste = True
            Exit For
        End If
    Next ws
    If Not configExiste Then
        Set wsConfig = ThisWorkbook.Sheets.Add
        wsConfig.Name = "CONFIG"
        wsConfig.Cells(1, 1).Value = "Themes"
        Dim defaultThemes() As String
        defaultThemes = Split(THEMES_ATELIERS, ",")
        Dim t As Integer
        For t = 0 To UBound(defaultThemes)
            wsConfig.Cells(t + 2, 1).Value = Trim(defaultThemes(t))
        Next t
        wsConfig.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True
    End If
    
    MsgBox "Initialisation terminée avec succès !", vbInformation, "Initialisation"
    Exit Sub
    
ErrFeuille:
    MsgBox "Erreur : Impossible de trouver une feuille requise." & vbCrLf & _
           "Vérifiez que les feuilles ATELIERS, PARTICIPANTS, PRESENCES et IMPORT existent.", _
           vbCritical, "Erreur d'initialisation"
End Sub

' -----------------------------------------------------------------------------
' VerifierOuCreerTableau : Crée un tableau structuré si il n'existe pas
' Paramètres :
'   ws        : La feuille cible
'   nomTable  : Le nom du tableau à créer
'   colonnes  : Tableau des noms de colonnes
' -----------------------------------------------------------------------------
Private Sub VerifierOuCreerTableau(ws As Worksheet, nomTable As String, colonnes As Variant)
    Dim tbl As ListObject
    Dim col As Variant
    Dim i As Integer
    
    ' Vérifier si le tableau existe déjà
    For Each tbl In ws.ListObjects
        If tbl.Name = nomTable Then
            Exit Sub ' Le tableau existe déjà
        End If
    Next tbl
    
    ' Désprotéger la feuille
    ws.Unprotect Password:=MOT_DE_PASSE
    
    ' Écrire les en-têtes en ligne 1
    For i = 0 To UBound(colonnes)
        ws.Cells(1, i + 1).Value = colonnes(i)
    Next i
    
    ' Créer le tableau structuré
    Set tbl = ws.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(colonnes) + 1)), _
        XlListObjectHasHeaders:=xlYes)
    tbl.Name = nomTable
    
    ' Reprotéger la feuille
    ws.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True
End Sub

' -----------------------------------------------------------------------------
' OuvrirFrmNouvelAtelier : Ouvre le formulaire Nouvel Atelier (pour bouton)
' -----------------------------------------------------------------------------
Public Sub OuvrirFrmNouvelAtelier()
    FrmNouvelAtelier.Show
End Sub

' -----------------------------------------------------------------------------
' OuvrirFrmSaisirPresences : Ouvre le formulaire Saisir Présences (pour bouton)
' -----------------------------------------------------------------------------
Public Sub OuvrirFrmSaisirPresences()
    FrmSaisirPresences.Show
End Sub

' -----------------------------------------------------------------------------
' OuvrirFrmNouveauParticipant : Ouvre le formulaire Nouveau Participant (pour bouton)
' -----------------------------------------------------------------------------
Public Sub OuvrirFrmNouveauParticipant()
    FrmNouveauParticipant.Show
End Sub

' -----------------------------------------------------------------------------
' OuvrirFrmGererParticipants : Ouvre le formulaire Gérer Participants (pour bouton)
' -----------------------------------------------------------------------------
Public Sub OuvrirFrmGererParticipants()
    FrmGererParticipants.Show
End Sub

' -----------------------------------------------------------------------------
' OuvrirFrmGererAteliers : Ouvre le formulaire Gérer Ateliers (pour bouton)
' -----------------------------------------------------------------------------
Public Sub OuvrirFrmGererAteliers()
    FrmGererAteliers.Show
End Sub

' -----------------------------------------------------------------------------
' OuvrirFrmGererThemes : Ouvre le formulaire Gérer Thèmes (pour bouton)
' -----------------------------------------------------------------------------
Public Sub OuvrirFrmGererThemes()
    FrmGererThemes.Show
End Sub

