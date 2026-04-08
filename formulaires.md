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
| `LstAteliers` | ListBox | *(vide)* | 12 | 32 | 440 | 120 |
| `LblSelPart` | Label | `2. Sélectionnez les participants présents :` | 12 | 165 | 280 | 18 |
| `LstParticipants` | ListBox | *(vide)* | 12 | 185 | 440 | 200 |
| `BtnValider` | CommandButton | `Valider les présences` | 60 | 400 | 160 | 30 |
| `BtnAnnuler` | CommandButton | `Annuler` | 280 | 400 | 100 | 30 |

**Propriétés de `LstAteliers` :**
- `MultiSelect` : `0 - fmMultiSelectSingle`
- `ColumnCount` : `3`
- `ColumnWidths` : `40;180;100`

**Propriétés de `LstParticipants` :**
- `MultiSelect` : `1 - fmMultiSelectMulti`
- `ColumnCount` : `3`
- `ColumnWidths` : `150;120;100`

> **Fonctionnement** : La sélection d'un atelier dans `LstAteliers` déclenche le chargement de `LstParticipants`. Les participants déjà présents à l'atelier sélectionné sont automatiquement pré-cochés et grisés.

---

## Formulaire 3 : `FrmNouveauParticipant`

**Propriétés du formulaire :**
| Propriété | Valeur |
|---|---|
| Name | `FrmNouveauParticipant` |
| Caption | `Nouveau Participant` |
| Width | `380` |
| Height | `440` |
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
| `BtnEnregistrer` | CommandButton | `Enregistrer` | 60 | 370 | 100 | 30 |
| `BtnAnnuler` | CommandButton | `Annuler` | 210 | 370 | 100 | 30 |

**Propriétés du ComboBox `CboStatut` :**
- `Style` : `2 - fmStyleDropDownList`
- Valeurs chargées par `UserForm_Initialize` : `Projet pro`, `Lancé`

> **Note** : Le champ `TxtDateContact` est pré-rempli avec la date du jour à l'ouverture du formulaire.

---

## Formulaire 4 : `FrmGererParticipants`

**Propriétés du formulaire :**
| Propriété | Valeur |
|---|---|
| Name | `FrmGererParticipants` |
| Caption | `Gérer les Participants` |
| Width | `600` |
| Height | `560` |
| StartUpPosition | `1 - CenterOwner` |

### Contrôles — Zone de recherche

| Name | Type | Caption/Text | Left | Top | Width | Height |
|---|---|---|---|---|---|---|
| `LblRecherche` | Label | `Rechercher (Nom ou Prénom) :` | 12 | 12 | 180 | 18 |
| `TxtRecherche` | TextBox | *(vide)* | 200 | 9 | 200 | 22 |
| `BtnRechercher` | CommandButton | `Rechercher` | 410 | 8 | 80 | 24 |
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

### Contrôles — Boutons d'action

| Name | Type | Caption/Text | Left | Top | Width | Height |
|---|---|---|---|---|---|---|
| `BtnModifier` | CommandButton | `Modifier` | 340 | 200 | 100 | 30 |
| `BtnSauvegarder` | CommandButton | `Sauvegarder` | 340 | 245 | 100 | 30 |
| `BtnFermer` | CommandButton | `Fermer` | 340 | 490 | 100 | 30 |

**Propriétés du ComboBox `CboEStatut` :**
- `Style` : `2 - fmStyleDropDownList`
- Valeurs : `Projet pro`, `Lancé`

**Propriétés des TextBox d'édition :**
- Par défaut, tous les champs d'édition ont `Enabled = False`
- Ils sont activés (`Enabled = True`) uniquement lorsqu'on clique sur le bouton "Modifier"

> **Fonctionnement** : La recherche s'effectue sur Nom et Prénom. Un clic sur "Modifier" charge les informations du participant sélectionné dans les champs éditables. "Sauvegarder" met à jour les données dans la feuille `PARTICIPANTS`.

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
