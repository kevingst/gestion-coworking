# Structure des Formulaires VBA — Gestion des Ateliers de Coworking

Ce document décrit la structure exacte de chaque UserForm à créer dans l'éditeur VBA d'Excel. Pour chaque formulaire, vous trouverez les noms des contrôles, leurs propriétés et leur disposition.

> **Comment créer un UserForm** : Dans l'éditeur VBA (Alt+F11), menu **Insertion → UserForm**. Utilisez la Boîte à outils (menu **Affichage → Boîte à outils**) pour glisser-déposer les contrôles.

---

## Formulaire 1 : `FrmNouvelAtelier`

**Propriétés du formulaire :**
| Propriété | Valeur |
|---|---|
| Name | `FrmNouvelAtelier` |
| Caption | `Nouvel Atelier` |
| Width | `350` |
| Height | `320` |
| StartUpPosition | `1 - CenterOwner` |

### Contrôles

| Name | Type | Caption/Text | Left | Top | Width | Height |
|---|---|---|---|---|---|---|
| `LblNom` | Label | `Nom de l'atelier :` | 12 | 18 | 120 | 18 |
| `TxtNom` | TextBox | *(vide)* | 140 | 15 | 180 | 22 |
| `LblDate` | Label | `Date (JJ/MM/AAAA) :` | 12 | 48 | 120 | 18 |
| `TxtDate` | TextBox | *(vide)* | 140 | 45 | 180 | 22 |
| `LblHeureDebut` | Label | `Heure début (HH:MM) :` | 12 | 78 | 120 | 18 |
| `TxtHeureDebut` | TextBox | *(vide)* | 140 | 75 | 180 | 22 |
| `LblHeureFin` | Label | `Heure fin (HH:MM) :` | 12 | 108 | 120 | 18 |
| `TxtHeureFin` | TextBox | *(vide)* | 140 | 105 | 180 | 22 |
| `LblTheme` | Label | `Thème :` | 12 | 138 | 120 | 18 |
| `CboTheme` | ComboBox | *(vide)* | 140 | 135 | 180 | 22 |
| `BtnEnregistrer` | CommandButton | `Enregistrer` | 60 | 220 | 100 | 30 |
| `BtnAnnuler` | CommandButton | `Annuler` | 180 | 220 | 100 | 30 |

**Propriétés du ComboBox `CboTheme` :**
- `Style` : `2 - fmStyleDropDownList` (liste déroulante non éditable)
- Les valeurs sont chargées par le code VBA dans `UserForm_Initialize`

---

## Formulaire 2 : `FrmSaisirPresences`

**Propriétés du formulaire :**
| Propriété | Valeur |
|---|---|
| Name | `FrmSaisirPresences` |
| Caption | `Saisir les Présences` |
| Width | `480` |
| Height | `500` |
| StartUpPosition | `1 - CenterOwner` |

### Contrôles

| Name | Type | Caption/Text | Left | Top | Width | Height |
|---|---|---|---|---|---|---|
| `LblSelAtelier` | Label | `1. Sélectionnez un atelier :` | 12 | 12 | 200 | 18 |
| `LblRechercheAtelier` | Label | `Rechercher un atelier :` | 12 | 32 | 140 | 18 |
| `TxtRechercheAtelier` | TextBox | *(vide)* | 158 | 29 | 294 | 22 |
| `LstAteliers` | ListBox | *(vide)* | 12 | 56 | 440 | 110 |
| `LblSelPart` | Label | `2. Sélectionnez les participants présents :` | 12 | 178 | 280 | 18 |
| `LblRechercheParticipant` | Label | `Rechercher un participant :` | 12 | 198 | 150 | 18 |
| `TxtRechercheParticipant` | TextBox | *(vide)* | 168 | 195 | 284 | 22 |
| `LstParticipants` | ListBox | *(vide)* | 12 | 220 | 440 | 170 |
| `BtnValider` | CommandButton | `Valider les présences` | 60 | 405 | 160 | 30 |
| `BtnAnnuler` | CommandButton | `Annuler` | 280 | 405 | 100 | 30 |

**Propriétés de `LstAteliers` :**
- `MultiSelect` : `0 - fmMultiSelectSingle`
- `ColumnCount` : `3`
- `ColumnWidths` : `40;180;100`

**Propriétés de `LstParticipants` :**
- `MultiSelect` : `1 - fmMultiSelectMulti`
- `ColumnCount` : `4`
- `ColumnWidths` : `0;150;120;100`

> **Fonctionnement** : La saisie dans `TxtRechercheAtelier` filtre la liste des ateliers en temps réel (tri par date décroissante). La sélection d'un atelier dans `LstAteliers` déclenche le chargement de `LstParticipants`. La saisie dans `TxtRechercheParticipant` filtre la liste des participants en temps réel. Les participants déjà présents à l'atelier sélectionné sont automatiquement pré-cochés. La première colonne de `LstParticipants` (largeur 0) contient l'ID masqué du participant.

---

## Formulaire 3 : `FrmNouveauParticipant`

**Propriétés du formulaire :**
| Propriété | Valeur |
|---|---|
| Name | `FrmNouveauParticipant` |
| Caption | `Nouveau Participant` |
| Width | `380` |
| Height | `470` |
| StartUpPosition | `1 - CenterOwner` |

### Contrôles

| Name | Type | Caption/Text | Left | Top | Width | Height |
|---|---|---|---|---|---|---|
| `LblNom` | Label | `Nom :` | 12 | 18 | 110 | 18 |
| `TxtNom` | TextBox | *(vide)* | 130 | 15 | 220 | 22 |
| `LblPrenom` | Label | `Prénom :` | 12 | 48 | 110 | 18 |
| `TxtPrenom` | TextBox | *(vide)* | 130 | 45 | 220 | 22 |
| `LblStatut` | Label | `Statut :` | 12 | 78 | 110 | 18 |
| `CboStatut` | ComboBox | *(vide)* | 130 | 75 | 220 | 22 |
| `LblDateContact` | Label | `Date premier contact :` | 12 | 108 | 110 | 18 |
| `TxtDateContact` | TextBox | *(date du jour)* | 130 | 105 | 220 | 22 |
| `LblEntreprise` | Label | `Nom entreprise :` | 12 | 138 | 110 | 18 |
| `TxtEntreprise` | TextBox | *(vide)* | 130 | 135 | 220 | 22 |
| `LblCommune` | Label | `Commune :` | 12 | 168 | 110 | 18 |
| `TxtCommune` | TextBox | *(vide)* | 130 | 165 | 220 | 22 |
| `LblCodePostal` | Label | `Code postal :` | 12 | 198 | 110 | 18 |
| `TxtCodePostal` | TextBox | *(vide)* | 130 | 195 | 220 | 22 |
| `LblMail` | Label | `Mail :` | 12 | 228 | 110 | 18 |
| `TxtMail` | TextBox | *(vide)* | 130 | 225 | 220 | 22 |
| `LblTelephone` | Label | `Téléphone :` | 12 | 258 | 110 | 18 |
| `TxtTelephone` | TextBox | *(vide)* | 130 | 255 | 220 | 22 |
| `LblActivite` | Label | `Activité :` | 12 | 288 | 110 | 18 |
| `TxtActivite` | TextBox | *(vide)* | 130 | 285 | 220 | 22 |
| `ChkNewsletter` | CheckBox | `Accepte la newsletter` | 130 | 318 | 220 | 22 |
| `BtnEnregistrer` | CommandButton | `Enregistrer` | 60 | 400 | 100 | 30 |
| `BtnAnnuler` | CommandButton | `Annuler` | 210 | 400 | 100 | 30 |

**Propriétés du ComboBox `CboStatut` :**
- `Style` : `2 - fmStyleDropDownList`
- Valeurs chargées par `UserForm_Initialize` : `Projet pro`, `Lancé`

**Propriétés de `ChkNewsletter` :**
- `Value` : `False` par défaut

> **Note** : Le champ `TxtDateContact` est pré-rempli avec la date du jour à l'ouverture du formulaire.

---

## Formulaire 4 : `FrmGererParticipants`

**Propriétés du formulaire :**
| Propriété | Valeur |
|---|---|
| Name | `FrmGererParticipants` |
| Caption | `Gérer les Participants` |
| Width | `600` |
| Height | `620` |
| StartUpPosition | `1 - CenterOwner` |

### Contrôles — Zone de recherche

| Name | Type | Caption/Text | Left | Top | Width | Height |
|---|---|---|---|---|---|---|
| `LblRecherche` | Label | `Rechercher (Nom ou Prénom) :` | 12 | 12 | 180 | 18 |
| `TxtRecherche` | TextBox | *(vide)* | 200 | 9 | 380 | 22 |
| `LstResultats` | ListBox | *(vide)* | 12 | 42 | 560 | 120 |

**Propriétés de `LstResultats` :**
- `MultiSelect` : `0 - fmMultiSelectSingle`
- `ColumnCount` : `4`
- `ColumnWidths` : `40;150;120;100`

### Contrôles — Zone d'édition

| Name | Type | Caption/Text | Left | Top | Width | Height |
|---|---|---|---|---|---|---|
| `LblEdit` | Label | `Modifier le participant sélectionné :` | 12 | 175 | 250 | 18 |
| `LblENom` | Label | `Nom :` | 12 | 200 | 100 | 18 |
| `TxtENom` | TextBox | *(vide)* | 120 | 197 | 200 | 22 |
| `LblEPrenom` | Label | `Prénom :` | 12 | 228 | 100 | 18 |
| `TxtEPrenom` | TextBox | *(vide)* | 120 | 225 | 200 | 22 |
| `LblEStatut` | Label | `Statut :` | 12 | 256 | 100 | 18 |
| `CboEStatut` | ComboBox | *(vide)* | 120 | 253 | 200 | 22 |
| `LblEDateContact` | Label | `Date contact :` | 12 | 284 | 100 | 18 |
| `TxtEDateContact` | TextBox | *(vide)* | 120 | 281 | 200 | 22 |
| `LblEEntreprise` | Label | `Entreprise :` | 12 | 312 | 100 | 18 |
| `TxtEEntreprise` | TextBox | *(vide)* | 120 | 309 | 200 | 22 |
| `LblECommune` | Label | `Commune :` | 12 | 340 | 100 | 18 |
| `TxtECommune` | TextBox | *(vide)* | 120 | 337 | 200 | 22 |
| `LblECodePostal` | Label | `Code postal :` | 12 | 368 | 100 | 18 |
| `TxtECodePostal` | TextBox | *(vide)* | 120 | 365 | 200 | 22 |
| `LblEMail` | Label | `Mail :` | 12 | 396 | 100 | 18 |
| `TxtEMail` | TextBox | *(vide)* | 120 | 393 | 200 | 22 |
| `LblETelephone` | Label | `Téléphone :` | 12 | 424 | 100 | 18 |
| `TxtETelephone` | TextBox | *(vide)* | 120 | 421 | 200 | 22 |
| `LblEActivite` | Label | `Activité :` | 12 | 452 | 100 | 18 |
| `TxtEActivite` | TextBox | *(vide)* | 120 | 449 | 200 | 22 |
| `ChkNewsletter` | CheckBox | `Accepte la newsletter` | 120 | 480 | 200 | 22 |
| `LblNbAteliers` | Label | `Nb ateliers participés :` | 12 | 508 | 100 | 18 |
| `TxtNbAteliers` | TextBox | *(vide)* | 120 | 505 | 80 | 22 |

### Contrôles — Boutons d'action

| Name | Type | Caption/Text | Left | Top | Width | Height |
|---|---|---|---|---|---|---|
| `BtnSauvegarder` | CommandButton | `Sauvegarder` | 340 | 245 | 100 | 30 |
| `BtnFermer` | CommandButton | `Fermer` | 340 | 555 | 100 | 30 |

**Propriétés du ComboBox `CboEStatut` :**
- `Style` : `2 - fmStyleDropDownList`
- Valeurs : `Projet pro`, `Lancé`

**Propriétés des TextBox d'édition :**
- Par défaut, tous les champs d'édition ont `Enabled = False`
- Ils sont activés (`Enabled = True`) automatiquement au clic sur un participant dans la liste

**Propriétés de `ChkNewsletter` :**
- `Enabled = False` par défaut (activé au clic sur un participant)

**Propriétés de `TxtNbAteliers` :**
- `Enabled = False`, `Locked = True` (lecture seule — calculé automatiquement)

> **Fonctionnement** : La saisie dans `TxtRecherche` filtre la liste des participants en temps réel. Un clic direct sur un participant dans `LstResultats` charge ses informations dans les champs éditables et les active immédiatement (sans bouton "Modifier"). `BtnSauvegarder` met à jour les données dans la feuille `PARTICIPANTS` et recalcule `Nb_Ateliers_Participes`.

---

## Formulaire 5 : `FrmGererAteliers`

**Propriétés du formulaire :**
| Propriété | Valeur |
|---|---|
| Name | `FrmGererAteliers` |
| Caption | `Gérer les Ateliers` |
| Width | `640` |
| Height | `620` |
| StartUpPosition | `1 - CenterOwner` |

### Contrôles — Zone de recherche et liste des ateliers

| Name | Type | Caption/Text | Left | Top | Width | Height |
|---|---|---|---|---|---|---|
| `LblRecherche` | Label | `Rechercher par nom :` | 12 | 12 | 130 | 18 |
| `TxtRecherche` | TextBox | *(vide)* | 148 | 9 | 200 | 22 |
| `LstAteliers` | ListBox | *(vide)* | 12 | 40 | 600 | 110 |

**Propriétés de `LstAteliers` :**
- `MultiSelect` : `0 - fmMultiSelectSingle`
- `ColumnCount` : `4`
- `ColumnWidths` : `40;200;80;120`

### Contrôles — Zone de détail de l'atelier

| Name | Type | Caption/Text | Left | Top | Width | Height |
|---|---|---|---|---|---|---|
| `LblNom` | Label | `Nom de l'atelier :` | 12 | 165 | 120 | 18 |
| `TxtNom` | TextBox | *(vide)* | 140 | 162 | 200 | 22 |
| `LblDate` | Label | `Date (JJ/MM/AAAA) :` | 12 | 193 | 120 | 18 |
| `TxtDate` | TextBox | *(vide)* | 140 | 190 | 120 | 22 |
| `LblHeureDebut` | Label | `Heure début (HH:MM) :` | 12 | 221 | 120 | 18 |
| `TxtHeureDebut` | TextBox | *(vide)* | 140 | 218 | 80 | 22 |
| `LblHeureFin` | Label | `Heure fin (HH:MM) :` | 12 | 249 | 120 | 18 |
| `TxtHeureFin` | TextBox | *(vide)* | 140 | 246 | 80 | 22 |
| `LblDuree` | Label | `Durée :` | 12 | 277 | 120 | 18 |
| `TxtDuree` | TextBox | *(vide)* | 140 | 274 | 80 | 22 |
| `LblTheme` | Label | `Thème :` | 12 | 305 | 120 | 18 |
| `CboTheme` | ComboBox | *(vide)* | 140 | 302 | 180 | 22 |
| `LblNbPart` | Label | `Nb participants :` | 360 | 165 | 120 | 18 |
| `TxtNbParticipants` | TextBox | *(vide)* | 488 | 162 | 60 | 22 |
| `LblNbPartPro` | Label | `Nb participants pro :` | 360 | 193 | 120 | 18 |
| `TxtNbParticipantsPro` | TextBox | *(vide)* | 488 | 190 | 60 | 22 |

**Propriétés importantes :**
- `TxtDuree` : `Enabled = False`, `Locked = True` (lecture seule)
- `TxtNbParticipants` : `Enabled = False`, `Locked = True` (lecture seule)
- `TxtNbParticipantsPro` : `Enabled = False`, `Locked = True` (lecture seule)
- `CboTheme` : `Style = 2 - fmStyleDropDownList`
- Les champs de détail (`TxtNom`, `TxtDate`, `TxtHeureDebut`, `TxtHeureFin`, `CboTheme`) ont `Enabled = False` par défaut ; ils s'activent à la sélection d'un atelier

### Contrôles — Zone des présences

| Name | Type | Caption/Text | Left | Top | Width | Height |
|---|---|---|---|---|---|---|
| `LblPresences` | Label | `Participants présents à l'atelier :` | 12 | 340 | 220 | 18 |
| `LstPresences` | ListBox | *(vide)* | 12 | 362 | 600 | 130 |

**Propriétés de `LstPresences` :**
- `MultiSelect` : `0 - fmMultiSelectSingle`
- `ColumnCount` : `4`
- `ColumnWidths` : `40;150;120;100`

### Contrôles — Boutons d'action

| Name | Type | Caption/Text | Left | Top | Width | Height |
|---|---|---|---|---|---|---|
| `BtnSauvegarder` | CommandButton | `💾 Sauvegarder` | 12 | 510 | 130 | 30 |
| `BtnSupprimerAtelier` | CommandButton | `🗑️ Supprimer l'atelier` | 160 | 510 | 160 | 30 |
| `BtnSupprimerPresence` | CommandButton | `❌ Retirer la présence` | 340 | 510 | 160 | 30 |
| `BtnFermer` | CommandButton | `Fermer` | 520 | 510 | 90 | 30 |

**Propriétés des boutons d'action :**
- `BtnSauvegarder` : `Enabled = False` par défaut
- `BtnSupprimerAtelier` : `Enabled = False` par défaut
- `BtnSupprimerPresence` : `Enabled = False` par défaut

### Disposition suggérée

```
┌─────────────────────────────────────────────────────────────┐
│ Gérer les Ateliers                                          │
├─────────────────────────────────────────────────────────────┤
│ Rechercher par nom : [TxtRecherche          ]               │
│ ┌─────────────────────────────────────────────────────────┐ │
│ │ LstAteliers (ID | Nom | Date | Thème)                   │ │
│ └─────────────────────────────────────────────────────────┘ │
│ ─────────────────── Détail de l'atelier ────────────────── │
│ Nom de l'atelier : [TxtNom            ]  Nb part. :  [   ] │
│ Date (JJ/MM/AAAA): [TxtDate  ]           Nb part.pro:[   ] │
│ Heure début      : [TxtHDbut ]                              │
│ Heure fin        : [TxtHFin  ]                              │
│ Durée            : [TxtDuree ]                              │
│ Thème            : [CboTheme            ]                   │
│ ───────────────── Participants présents ───────────────── │
│ ┌─────────────────────────────────────────────────────────┐ │
│ │ LstPresences (ID | Nom | Prénom | Statut)               │ │
│ └─────────────────────────────────────────────────────────┘ │
│ [Sauvegarder] [Supprimer l'atelier] [Retirer présence] [X] │
└─────────────────────────────────────────────────────────────┘
```

### Instructions pour créer le formulaire dans l'éditeur VBA

1. Dans l'éditeur VBA (Alt+F11), menu **Insertion → UserForm**
2. Dans la fenêtre **Propriétés**, définir :
   - `Name` : `FrmGererAteliers`
   - `Caption` : `Gérer les Ateliers`
   - `Width` : `640`
   - `Height` : `620`
3. Ajouter les contrôles depuis la **Boîte à outils** en respectant le tableau ci-dessus
4. Pour `TxtDuree`, `TxtNbParticipants`, `TxtNbParticipantsPro` : définir `Enabled = False` et `Locked = True` dans les propriétés
5. Pour `CboTheme` : définir `Style = 2 - fmStyleDropDownList`
6. Pour les boutons `BtnSauvegarder`, `BtnSupprimerAtelier`, `BtnSupprimerPresence` : définir `Enabled = False`
7. Double-cliquer sur le formulaire pour ouvrir la fenêtre de code
8. Copier-coller le contenu du fichier `src/FrmGererAteliers.frm`

---

## Formulaire 6 : `FrmGererThemes`

**Propriétés du formulaire :**
| Propriété | Valeur |
|---|---|
| Name | `FrmGererThemes` |
| Caption | `Gérer les Thèmes` |
| Width | `320` |
| Height | `340` |
| StartUpPosition | `1 - CenterOwner` |

### Contrôles

| Name | Type | Caption/Text | Left | Top | Width | Height |
|---|---|---|---|---|---|---|
| `LblThemes` | Label | `Thèmes disponibles :` | 12 | 12 | 180 | 18 |
| `LstThemes` | ListBox | *(vide)* | 12 | 34 | 280 | 160 |
| `LblNouveauTheme` | Label | `Nouveau thème :` | 12 | 208 | 100 | 18 |
| `TxtNouveauTheme` | TextBox | *(vide)* | 120 | 205 | 172 | 22 |
| `BtnAjouter` | CommandButton | `Ajouter` | 12 | 240 | 90 | 28 |
| `BtnSupprimer` | CommandButton | `Supprimer` | 116 | 240 | 90 | 28 |
| `BtnFermer` | CommandButton | `Fermer` | 220 | 240 | 72 | 28 |

**Propriétés de `LstThemes` :**
- `MultiSelect` : `0 - fmMultiSelectSingle`

> **Fonctionnement** : Les thèmes sont lus depuis la feuille `CONFIG` (colonne A, à partir de A2). "Ajouter" écrit le nouveau thème à la première ligne vide de CONFIG (avec vérification des doublons). "Supprimer" supprime la ligne sélectionnée dans CONFIG. La liste est rechargée automatiquement après chaque modification. La touche Échap ferme le formulaire.

### Instructions pour créer le formulaire dans l'éditeur VBA

1. Dans l'éditeur VBA (Alt+F11), menu **Insertion → UserForm**
2. Dans la fenêtre **Propriétés**, définir :
   - `Name` : `FrmGererThemes`
   - `Caption` : `Gérer les Thèmes`
   - `Width` : `320`
   - `Height` : `340`
3. Ajouter les contrôles depuis la **Boîte à outils** en respectant le tableau ci-dessus
4. Double-cliquer sur le formulaire pour ouvrir la fenêtre de code
5. Copier-coller le contenu du fichier `src/FrmGererThemes.frm`

---

## Notes générales sur les UserForms

1. **Toujours vérifier les noms des contrôles** : Le code VBA référence les contrôles par leur nom exact. Une erreur de nom provoquera une erreur à l'exécution.

2. **Ordre de tabulation** : Configurez l'ordre de tabulation (TabIndex) pour faciliter la saisie au clavier : numérotez les contrôles dans l'ordre logique de saisie.

3. **Validation des champs obligatoires** :
   - `FrmNouvelAtelier` : Nom et Date sont obligatoires
   - `FrmNouveauParticipant` : Nom est obligatoire
   - La validation est gérée dans le code du bouton "Enregistrer"

4. **Touche Entrée** : Pour que la touche Entrée soumette le formulaire, définissez `Default = True` sur le bouton "Enregistrer".

5. **Touche Échap** : Pour que la touche Échap ferme le formulaire, définissez `Cancel = True` sur le bouton "Annuler" ou "Fermer".
