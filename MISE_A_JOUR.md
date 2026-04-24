# Guide de mise à jour — Gestion des Ateliers de Coworking

Ce document explique comment appliquer chaque mise à jour du code VBA dans votre fichier Excel `.xlsm`.
Il est mis à jour à chaque nouvelle modification du projet.

---

## Comment ouvrir l'éditeur VBA

1. Ouvrez votre fichier `GestionCoworking.xlsm` dans Excel
2. Appuyez sur **Alt + F11** pour ouvrir l'éditeur Visual Basic for Applications (VBA)
3. Dans le panneau de gauche (Explorateur de projets), vous verrez la liste des modules et formulaires du projet

> Si l'Explorateur de projets n'est pas visible, utilisez le menu **Affichage → Explorateur de projets** (ou **Ctrl + R**).

---

## Procédure générale pour remplacer le code d'un module ou formulaire

### Pour un module standard (fichier `.bas`)

1. Dans l'Explorateur de projets, double-cliquez sur le module concerné (ex : `ModuleStats`)
2. Dans la fenêtre de code qui s'ouvre, sélectionnez tout le contenu : **Ctrl + A**
3. Copiez le nouveau contenu depuis le fichier `.bas` correspondant dans le dossier `src/`
4. Collez-le dans la fenêtre de code : **Ctrl + V**
5. Sauvegardez le fichier Excel : **Ctrl + S** (depuis Excel, pas depuis l'éditeur VBA)

### Pour un UserForm (fichier `.frm`)

1. Dans l'Explorateur de projets, double-cliquez sur le formulaire concerné (ex : `FrmSaisirPresences`)
2. Pour accéder au **code** du formulaire : clic droit → **Visualiser le code** (ou appuyer sur **F7**)
3. Sélectionnez tout le contenu : **Ctrl + A**
4. Copiez le nouveau contenu depuis le fichier `.frm` correspondant dans le dossier `src/`
5. Collez-le dans la fenêtre de code : **Ctrl + V**
6. Sauvegardez le fichier Excel : **Ctrl + S**

> **Important** : La procédure ci-dessus remplace uniquement le **code VBA** du formulaire.
> Si la mise à jour nécessite d'ajouter ou de supprimer des contrôles (boutons, zones de texte, etc.),
> vous devrez également modifier le formulaire visuellement dans le designer VBA.
> Consultez la section correspondante dans `formulaires.md` pour la liste exacte des contrôles attendus.

### Pour `ThisWorkbook`

1. Dans l'Explorateur de projets, double-cliquez sur **ThisWorkbook** (sous « Microsoft Excel Objets »)
2. Sélectionnez tout le contenu : **Ctrl + A**
3. Copiez le contenu du fichier `src/ThisWorkbook.bas`
4. Collez et sauvegardez

---

## Historique des mises à jour

### 2026-04-10 — Création initiale

Mise en place complète du projet. Voir `SETUP.md` pour la procédure d'installation complète.

**Fichiers créés :**
- `src/ThisWorkbook.bas`
- `src/ModuleStats.bas`
- `src/ModuleAteliers.bas`
- `src/ModuleParticipants.bas`
- `src/ModulePresences.bas`
- `src/ModuleImport.bas`
- `src/FrmNouvelAtelier.frm`
- `src/FrmSaisirPresences.frm`
- `src/FrmNouveauParticipant.frm`
- `src/FrmGererParticipants.frm`

---

### 2026-04-10 — Ajout du formulaire FrmGererAteliers

Nouveau formulaire permettant de rechercher, modifier et supprimer des ateliers,
ainsi que de gérer les présences d'un atelier sélectionné.

**Fichiers modifiés :**
- `src/ThisWorkbook.bas` — ajout de la procédure `OuvrirFrmGererAteliers()`

**Fichiers créés :**
- `src/FrmGererAteliers.frm` — nouveau formulaire complet

**Manipulations à effectuer dans Excel :**

1. Ouvrir l'éditeur VBA (Alt + F11)
2. Remplacer le code de `ThisWorkbook` (voir procédure générale ci-dessus)
3. Créer un nouveau UserForm :
   - Menu **Insertion → UserForm**
   - Nommer le formulaire `FrmGererAteliers` dans les propriétés
   - Ajouter tous les contrôles décrits dans `formulaires.md` (section `FrmGererAteliers`)
   - Coller le contenu de `src/FrmGererAteliers.frm` dans la fenêtre de code
4. Sur la feuille `ACCUEIL` dans Excel, ajouter un bouton **🔧 Gérer les Ateliers** et lui assigner la macro `OuvrirFrmGererAteliers`

---

### 2026-04-10 — Correction du bug des heures décimales

Correction d'un bug où les heures s'affichaient sous forme décimale (ex : `0,375` au lieu de `09:00`)
et où la durée totale des ateliers affichait `00:00`.

**Fichiers modifiés :**
- `src/FrmGererAteliers.frm` — conversion correcte des valeurs décimales en format `HH:MM`
- `src/ModuleStats.bas` — détection des durées stockées en décimal et conversion en minutes
- `src/ModuleAteliers.bas` — forçage du format `HH:MM` lors de l'enregistrement des heures

**Manipulations à effectuer dans Excel :**

1. Ouvrir l'éditeur VBA (Alt + F11)
2. Remplacer le code de `ModuleStats` (Ctrl + A, coller, sauvegarder)
3. Remplacer le code de `ModuleAteliers` (Ctrl + A, coller, sauvegarder)
4. Remplacer le code de `FrmGererAteliers` (clic droit → Visualiser le code, Ctrl + A, coller, sauvegarder)

---

### 2026-04-10 — Amélioration de FrmGererParticipants

Suppression du bouton « Rechercher » et du bouton « Modifier » : la recherche se fait désormais
en temps réel à la saisie, et un clic sur un participant charge ses informations directement.

**Fichiers modifiés :**
- `src/FrmGererParticipants.frm` — suppression de `BtnRechercher_Click` et `BtnModifier_Click`,
  ajout de `TxtRecherche_Change`, mise à jour de `LstResultats_Click`

**Manipulations à effectuer dans Excel :**

1. Ouvrir l'éditeur VBA (Alt + F11)
2. Remplacer le code de `FrmGererParticipants` (clic droit → Visualiser le code, Ctrl + A, coller, sauvegarder)
3. Dans le **designer** du formulaire `FrmGererParticipants` (double-clic sur le formulaire dans l'Explorateur),
   supprimer visuellement les boutons `BtnRechercher` et `BtnModifier`

---

### 2026-04-10 — Amélioration de FrmSaisirPresences

Ajout de barres de recherche avec détection automatique dans le formulaire de saisie des présences.
Les ateliers sont désormais triés du plus récent au plus ancien.

**Fichiers modifiés :**
- `src/FrmSaisirPresences.frm` — refonte complète avec tri par date décroissante,
  recherche en temps réel pour les ateliers et pour les participants

**Nouveaux contrôles à ajouter dans le designer du formulaire `FrmSaisirPresences` :**

| Nom | Type | Légende / Texte | Position suggérée |
|---|---|---|---|
| `LblRechercheAtelier` | Label | `Rechercher un atelier :` | Au-dessus de `LstAteliers` |
| `TxtRechercheAtelier` | TextBox | *(vide)* | Au-dessus de `LstAteliers` |
| `LblRechercheParticipant` | Label | `Rechercher un participant :` | Au-dessus de `LstParticipants` |
| `TxtRechercheParticipant` | TextBox | *(vide)* | Au-dessus de `LstParticipants` |

**Manipulations à effectuer dans Excel :**

1. Ouvrir l'éditeur VBA (Alt + F11)
2. Dans le **designer** du formulaire `FrmSaisirPresences` (double-clic dans l'Explorateur),
   ajouter les quatre contrôles décrits dans le tableau ci-dessus (voir aussi `formulaires.md`)
3. Remplacer le code de `FrmSaisirPresences` (clic droit → Visualiser le code, Ctrl + A, coller, sauvegarder)

> **Note** : La propriété `ColumnWidths` de `LstParticipants` passe de `"150;120;100"` à `"0;150;120;100"`
> (4 colonnes, la première masquée contenant l'ID du participant).
> Vérifiez que `ColumnCount = 4` est bien défini sur `LstParticipants`.

---

### [2026-04-10] Statistiques mensuelles — Feuille STATS + sélecteurs ACCUEIL

#### Fichiers modifiés
- `src/ModuleStats.bas` — refactorisé avec `RecalculerStatsAnnee()` et `MettreAJourAccueil()`
- `src/FeuilleStats.bas` — NOUVEAU (module de la feuille STATS)
- `src/FeuilleAccueil.bas` — NOUVEAU (module de la feuille ACCUEIL)

#### Étapes manuelles dans Excel

1. **CRÉER LA FEUILLE STATS**
   - Clic droit sur un onglet → Insérer une feuille → la nommer exactement `STATS`
   - Taper `2026` en cellule B1
   - Ajouter une validation de données sur B1 : Liste = `2024;2025;2026;2027;2028`

2. **COLLER LE CODE DE FEUILLE**
   - Dans l'éditeur VBA (Alt+F11), développer **Microsoft Excel Objects**
   - Double-cliquer sur **STATS** → coller le contenu de `src/FeuilleStats.bas`
   - Double-cliquer sur **ACCUEIL** → coller le contenu de `src/FeuilleAccueil.bas`
   - Double-cliquer sur **ModuleStats** → remplacer le contenu par `src/ModuleStats.bas`

3. **CONFIGURER ACCUEIL**
   - En B1 : taper le mois en cours (ex: `Avril`)
   - Ajouter une validation de données sur B1 : Liste = `Janvier,Février,Mars,Avril,Mai,Juin,Juillet,Août,Septembre,Octobre,Novembre,Décembre`
   - En B2 : taper l'année en cours (ex: `2026`)
   - Ajouter une validation de données sur B2 : Liste = `2024;2025;2026;2027;2028`

4. **INITIALISER LES STATS**
   - Dans l'éditeur VBA, ouvrir la fenêtre Exécution (Ctrl+G)
   - Taper : `MettreAJourStats`
   - Appuyer sur Entrée → les données apparaissent dans STATS et ACCUEIL

5. **CRÉER LE GRAPHIQUE (manuel)**
   - Dans STATS, sélectionner A2:E14 (en-têtes + 12 mois)
   - Insertion → Graphique → Histogramme groupé
   - Placer le graphique sur la feuille ACCUEIL
   - Titre du graphique : `Statistiques mensuelles ` & année

---

### [2026-04-10] Graphiques dynamiques — 2 graphiques sur ACCUEIL

#### Fichiers modifiés
- src/ModuleStats.bas — MettreAJourAccueil() enrichie : plage G:H + titres dynamiques

#### Étapes manuelles dans Excel (à faire une seule fois)

1. CRÉER LA PLAGE SOURCE G:H SUR ACCUEIL
   - Lancer MettreAJourStats depuis la fenêtre Exécution VBA (Ctrl+G puis taper MettreAJourStats)
   - Les cellules G1:H5 sont remplies automatiquement
   - Masquer les colonnes G et H : sélectionner les colonnes G:H, clic droit → Masquer

2. CRÉER LE GRAPHIQUE 1 — Stats du mois (GraphiqueMois)
   - Sélectionner ACCUEIL!G1:H5 (même si masquées, utiliser la zone de nom pour naviguer)
     Astuce : taper G1:H5 dans la zone de nom (en haut à gauche) puis Entrée
   - Insertion → Graphique → Histogramme groupé → OK
   - Placer le graphique sur ACCUEIL à l'emplacement souhaité
   - Cliquer sur le graphique → dans la zone de nom (en haut à gauche), taper : GraphiqueMois → Entrée
   - Le titre se mettra à jour automatiquement au prochain changement de sélecteur

3. CRÉER LE GRAPHIQUE 2 — Bilan annuel (GraphiqueAnnee)
   - Aller sur la feuille STATS
   - Sélectionner A2:E14 (en-têtes de colonnes + les 12 lignes de mois)
   - Insertion → Graphique → Courbe (ou Histogramme groupé) → OK
   - Couper le graphique (Ctrl+X) → aller sur ACCUEIL → Coller (Ctrl+V)
   - Cliquer sur le graphique → dans la zone de nom, taper : GraphiqueAnnee → Entrée
   - Le titre "Bilan 2026" se mettra à jour automatiquement

4. VÉRIFIER LE FONCTIONNEMENT
   - Changer le mois en ACCUEIL!B1 → le titre de GraphiqueMois et ses données changent
   - Changer l'année en ACCUEIL!B2 → les deux graphiques se mettent à jour

---

### [2026-04-16] Newsletter, comptage ateliers par participant, gestion des themes

#### Fichiers modifies
- src/ModuleParticipants.bas — ajout parametre newsletter, RecalculerNbAteliers()
- src/ModuleAteliers.bas — ObtenirListeThemes() lit depuis CONFIG
- src/ModulePresences.bas — appel RecalculerNbAteliers() apres enregistrement
- src/FrmGererParticipants.frm — ajout ChkNewsletter + TxtNbAteliers
- src/FrmNouveauParticipant.frm — ajout ChkNewsletter
- src/FrmGererThemes.frm — NOUVEAU formulaire

#### Etapes manuelles dans Excel

1. AJOUTER LES COLONNES DANS TblParticipants (feuille PARTICIPANTS)
   - Cliquer sur la cellule a droite de la derniere colonne du tableau
   - Ajouter colonne L : nommer l en-tete "Newsletter"
   - Ajouter colonne M : nommer l en-tete "Nb_Ateliers_Participes"
   - Pour les participants existants : remplir Newsletter avec "Non" par defaut
   - Pour Nb_Ateliers_Participes : laisser vide (sera recalcule)

2. CREER LA FEUILLE CONFIG
   - Clic droit sur un onglet -> Inserer -> nommer "CONFIG"
   - En A1 : taper "Themes" (en-tete)
   - En A2, A3, A4... : saisir les themes un par ligne

3. COLLER LE CODE VBA
   - Alt+F11 -> ouvrir chaque fichier modifie et remplacer le code
   - Pour FrmGererThemes : Insertion -> UserForm -> nommer "FrmGererThemes" -> coller le code

4. AJOUTER LES CONTROLES DANS LES FORMULAIRES
   FrmGererParticipants :
   - Ajouter une CheckBox nommee "ChkNewsletter" avec Caption "Accepte la newsletter"
   - Ajouter un TextBox nomme "TxtNbAteliers" avec Enabled=False
   - Ajouter un Label "Nb ateliers participes :" a cote

   FrmNouveauParticipant :
   - Ajouter une CheckBox nommee "ChkNewsletter" avec Caption "Accepte la newsletter"

   FrmGererThemes (nouveau formulaire) :
   - ListBox : "LstThemes"
   - TextBox : "TxtNouveauTheme"
   - CommandButton : "BtnAjouter" (Caption "Ajouter")
   - CommandButton : "BtnSupprimer" (Caption "Supprimer")
   - CommandButton : "BtnFermer" (Caption "Fermer")

5. RECALCULER LES DONNEES EXISTANTES
   - Ouvrir la fenetre Execution VBA (Ctrl+G)
   - Executer MettreAJourStats pour recalculer les stats globales
   - Pour recalculer Nb_Ateliers_Participes de tous les participants existants,
     ajouter temporairement cette macro dans un module et l executer :
     Sub RecalculerTousLesParticipants()
         Dim ws As Worksheet
         Dim tbl As ListObject
         Dim ligne As ListRow
         Set ws = ThisWorkbook.Sheets("PARTICIPANTS")
         Set tbl = ws.ListObjects("TblParticipants")
         For Each ligne In tbl.ListRows
             If IsNumeric(ligne.Range.Cells(1,1).Value) Then
                 Call RecalculerNbAteliers(CLng(ligne.Range.Cells(1,1).Value))
             End If
         Next ligne
     End Sub

6. AJOUTER LE BOUTON SUR ACCUEIL (optionnel)
   - Insertion -> Formes -> Rectangle -> Caption "Gerer les themes"
   - Clic droit -> Affecter une macro -> taper : OuvrirGererThemes
   - Dans un module standard, ajouter :
     Public Sub OuvrirGererThemes()
         FrmGererThemes.Show
     End Sub

---

### [2026-04-22] Export CSV participants newsletter pour Brevo

#### Fichiers modifiés
- src/ModuleExport.bas — NOUVEAU module d'export CSV

#### Étapes manuelles dans Excel (à faire une seule fois)

1. COLLER LE CODE VBA
   - Alt+F11 -> Dans un module standard existant ou nouveau, coller le contenu de src/ModuleExport.bas
   - Ou : Insertion -> Module -> coller le code -> renommer le module "ModuleExport"

2. AJOUTER LE BOUTON SUR ACCUEIL
   - Sur la feuille ACCUEIL, Insertion -> Formes -> Rectangle
   - Saisir le texte : "Exporter contacts Brevo"
   - Clic droit sur le bouton -> Affecter une macro -> sélectionner "ExporterParticipantsBrevo"
   - Mettre en forme le bouton selon le style de la feuille

3. UTILISATION
   - Cliquer sur le bouton "Exporter contacts Brevo"
   - Choisir l'emplacement de sauvegarde du fichier CSV
   - Le fichier généré contient les colonnes : EMAIL;PRENOM;NOM;STATUT
   - Seuls les participants avec Newsletter = "Oui" ET une adresse email renseignée sont exportés
   - Importer le CSV dans Brevo : Contacts -> Importer des contacts -> Importer un fichier CSV


### [2026-04-23] Suppression participant, cases à cocher présences, recalcul auto nb ateliers

#### Fichiers modifiés
- src/ModuleParticipants.bas — ajout SupprimerParticipant()
- src/FrmGererParticipants.frm — ajout BtnSupprimer + BtnSupprimer_Click()
- src/FrmSaisirPresences.frm — réécriture complète avec cases ☐/☑ persistantes
- src/ModulePresences.bas — vérification appel RecalculerNbAteliers()

#### Étapes manuelles dans Excel

1. COLLER LE CODE VBA
   - Alt+F11 → remplacer le code de chaque fichier modifié

2. AJOUTER LE BOUTON SUPPRIMER DANS FrmGererParticipants
   - Ouvrir FrmGererParticipants dans l'éditeur de formulaire
   - Ajouter un CommandButton nommé "BtnSupprimer"
   - Caption : "Supprimer"
   - Enabled : False (sera activé automatiquement à la sélection d'un participant)
   - Placer le bouton à côté de BtnSauvegarder

3. MODIFIER LstParticipants DANS FrmSaisirPresences
   - Ouvrir FrmSaisirPresences dans l'éditeur de formulaire
   - Sélectionner LstParticipants
   - Propriété MultiSelect : 0 - fmMultiSelectSingle (les cases ☐/☑ gèrent la sélection multiple)
   - Propriété ColumnCount : 5
   - Propriété ColumnWidths : "20;0;150;120;100"

### [2026-04-24] Champ Animé par + mail dans présences

#### Fichiers modifiés
- src/ModuleAteliers.bas — ajout paramètre `animepar` dans `EnregistrerAtelier()`, écriture en colonne J (Cells(1,10))
- src/FrmNouvelAtelier.frm — ajout contrôle `TxtAnimePar`, transmission à `EnregistrerAtelier()`
- src/FrmGererAteliers.frm — chargement/sauvegarde/activation/vidage de `TxtAnimePar`, `LstPresences` passe à 5 colonnes avec affichage du mail

#### Étapes manuelles dans Excel

1. AJOUTER LA COLONNE ANIME_PAR DANS TblAteliers (feuille ATELIERS)
   - Cliquer sur la cellule à droite de la dernière colonne du tableau TblAteliers
   - Ajouter la colonne J : nommer l'en-tête "Anime_Par"
   - Pour les ateliers existants, laisser vide (ou renseigner si connu)

2. AJOUTER TxtAnimePar DANS FrmNouvelAtelier
   - Ouvrir FrmNouvelAtelier dans l'éditeur de formulaire (Alt+F11)
   - Ajouter un TextBox nommé "TxtAnimePar"
   - Ajouter un Label "Animé par :" à côté
   - Remplacer le code VBA (clic droit → Visualiser le code, Ctrl+A, coller src/FrmNouvelAtelier.frm)

3. AJOUTER TxtAnimePar DANS FrmGererAteliers
   - Ouvrir FrmGererAteliers dans l'éditeur de formulaire
   - Ajouter un TextBox nommé "TxtAnimePar" dans la zone de détail de l'atelier
   - Ajouter un Label "Animé par :" à côté
   - Remplacer le code VBA (clic droit → Visualiser le code, Ctrl+A, coller src/FrmGererAteliers.frm)

4. METTRE À JOUR LstPresences DANS FrmGererAteliers
   - Dans le designer de FrmGererAteliers, sélectionner LstPresences
   - Propriété ColumnCount : 5
   - Propriété ColumnWidths : "40;150;120;100;180"
   - (Le code VBA définit ces valeurs automatiquement à l'initialisation)
