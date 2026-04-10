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
