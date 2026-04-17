' =============================================================================
' UserForm : FrmGererThemes
' Description : Formulaire de gestion des themes d'ateliers
'
' Controles requis :
'   - LstThemes       : ListBox — Liste des themes actuels
'   - TxtNouveauTheme : TextBox — Saisie d'un nouveau theme
'   - BtnAjouter      : CommandButton — Ajouter le theme saisi
'   - BtnSupprimer    : CommandButton — Supprimer le theme selectionne
'   - BtnFermer       : CommandButton — Fermer le formulaire
' =============================================================================

Private Sub UserForm_Initialize()
    Call ChargerThemes
    TxtNouveauTheme.SetFocus
End Sub

Private Sub ChargerThemes()
    Dim wsConfig As Worksheet
    LstThemes.Clear
    
    On Error GoTo ErrChargement
    Set wsConfig = ThisWorkbook.Sheets("CONFIG")
    On Error GoTo 0
    
    Dim i As Integer
    i = 2
    Do While wsConfig.Cells(i, 1).Value <> ""
        LstThemes.AddItem CStr(wsConfig.Cells(i, 1).Value)
        i = i + 1
    Loop
    Exit Sub
    
ErrChargement:
    MsgBox "La feuille CONFIG est introuvable.", vbCritical, "Erreur"
End Sub

Private Sub BtnAjouter_Click()
    Dim nouveauTheme As String
    nouveauTheme = Trim(TxtNouveauTheme.Value)
    
    If nouveauTheme = "" Then
        MsgBox "Veuillez saisir un nom de theme.", vbExclamation, "Champ vide"
        Exit Sub
    End If
    
    Dim i As Integer
    For i = 0 To LstThemes.ListCount - 1
        If LCase(LstThemes.List(i)) = LCase(nouveauTheme) Then
            MsgBox "Ce theme existe deja.", vbExclamation, "Doublon"
            Exit Sub
        End If
    Next i
    
    Dim wsConfig As Worksheet
    On Error GoTo ErrAjout
    Set wsConfig = ThisWorkbook.Sheets("CONFIG")
    On Error GoTo 0
    
    wsConfig.Unprotect Password:=MOT_DE_PASSE
    
    Dim ligne As Integer
    ligne = 2
    Do While wsConfig.Cells(ligne, 1).Value <> ""
        ligne = ligne + 1
    Loop
    wsConfig.Cells(ligne, 1).Value = nouveauTheme
    
    wsConfig.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True
    
    TxtNouveauTheme.Value = ""
    Call ChargerThemes
    Exit Sub
    
ErrAjout:
    MsgBox "Erreur lors de l ajout du theme.", vbCritical, "Erreur"
End Sub

Private Sub BtnSupprimer_Click()
    If LstThemes.ListIndex < 0 Then
        MsgBox "Veuillez selectionner un theme a supprimer.", vbExclamation, "Selection manquante"
        Exit Sub
    End If
    
    Dim themeASupprimer As String
    themeASupprimer = LstThemes.List(LstThemes.ListIndex)
    
    Dim reponse As Integer
    reponse = MsgBox("Supprimer le theme """ & themeASupprimer & """ ?", _
                     vbYesNo + vbQuestion, "Confirmation")
    If reponse <> vbYes Then Exit Sub
    
    Dim wsConfig As Worksheet
    On Error GoTo ErrSuppression
    Set wsConfig = ThisWorkbook.Sheets("CONFIG")
    On Error GoTo 0
    
    wsConfig.Unprotect Password:=MOT_DE_PASSE
    
    Dim j As Integer
    j = 2
    Do While wsConfig.Cells(j, 1).Value <> ""
        If CStr(wsConfig.Cells(j, 1).Value) = themeASupprimer Then
            wsConfig.Rows(j).Delete
            Exit Do
        End If
        j = j + 1
    Loop
    
    wsConfig.Protect Password:=MOT_DE_PASSE, UserInterfaceOnly:=True
    
    Call ChargerThemes
    Exit Sub
    
ErrSuppression:
    MsgBox "Erreur lors de la suppression du theme.", vbCritical, "Erreur"
End Sub

Private Sub BtnFermer_Click()
    Unload Me
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub
