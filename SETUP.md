# Guide d'installation — Gestion des Ateliers de Coworking

Ce guide explique pas à pas comment créer le fichier Excel `.xlsm` et y intégrer le code VBA pour obtenir le logiciel de gestion fonctionnel.

---

## Prérequis

- Microsoft Excel 2016 ou version ultérieure (Windows)
- Les macros VBA doivent être activées

---

## Étape 1 — Créer le fichier Excel

1. Ouvrez **Excel**
2. Créez un nouveau classeur vide : **Fichier → Nouveau → Classeur vide**
3. Enregistrez-le immédiatement au format `.xlsm` :
   - **Fichier → Enregistrer sous**
   - Choisissez l'emplacement souhaité
   - Dans "Type de fichier", sélectionnez **Classeur Excel (prenant en charge les macros) (\*.xlsm)**
   - Nommez-le `GestionCoworking.xlsm`
   - Cliquez sur **Enregistrer**

---

## Étape 2 — Créer les feuilles

Renommez les feuilles existantes et créez-en de nouvelles :

1. **Clic droit sur l'onglet "Feuil1"** → Renommer → tapez `ACCUEIL`
2. **Clic droit sur un onglet vide** (ou cliquer sur le `+`) → ajouter et nommer `ATELIERS`
3. Répétez pour : `PARTICIPANTS`, `PRESENCES`, `IMPORT`

À la fin, vous devez avoir **5 onglets** dans cet ordre : `ACCUEIL`, `ATELIERS`, `PARTICIPANTS`, `PRESENCES`, `IMPORT`

---

## Étape 3 — Créer les en-têtes des tableaux structurés

### Feuille `ATELIERS`

Dans la feuille `ATELIERS`, saisissez les en-têtes suivantes en **ligne 1** :

| A | B | C | D | E | F | G | H | I |
|---|---|---|---|---|---|---|---|---|
| ID_Atelier | Nom | Date | Heure_Debut | Heure_Fin | Duree | Theme | Nb_Participants | Nb_Participants_Pro |

Ensuite, convertissez en tableau structuré :
- Sélectionnez la plage A1:I1
- **Insertion → Tableau** → cochez "Mon tableau comporte des en-têtes" → OK
- Dans l'onglet "Création du tableau" (ou "Table Design"), nommez le tableau : `TblAteliers`

### Feuille `PARTICIPANTS`

En-têtes en **ligne 1** :

| A | B | C | D | E | F | G | H | I | J | K |
|---|---|---|---|---|---|---|---|---|---|---|
| ID_Participant | Nom | Prenom | Statut | Date_Premier_Contact | Nom_Entreprise | Commune | Code_Postal | Mail | Telephone | Activite |

Convertissez en tableau structuré nommé `TblParticipants`.

### Feuille `PRESENCES`

En-têtes en **ligne 1** :

| A | B | C | D | E | F |
|---|---|---|---|---|---|
| ID_Presence | ID_Atelier | ID_Participant | Nom_Participant | Prenom_Participant | Statut_Participant |

Convertissez en tableau structuré nommé `TblPresences`.

### Feuille `IMPORT`

En-têtes en **ligne 1** (identiques à PARTICIPANTS) :

| A | B | C | D | E | F | G | H | I | J | K |
|---|---|---|---|---|---|---|---|---|---|---|
| ID_Participant | Nom | Prenom | Statut | Date_Premier_Contact | Nom_Entreprise | Commune | Code_Postal | Mail | Telephone | Activite |

Ajoutez un bouton "Importer les données" sur cette feuille (voir Étape 6 pour les boutons).

---

## Étape 4 — Préparer la feuille ACCUEIL

Dans la feuille `ACCUEIL`, saisissez le texte suivant pour créer l'interface :

- **Cellule B2** : `Gestion des Ateliers de Coworking` (titre — mettez en forme : police grande, gras)
- **Cellule B4** : `Statistiques de l'année en cours`
- **Cellule B5** : `Nombre d'ateliers :`
- **Cellule C5** : *(laisser vide — sera remplie par la macro)*
- **Cellule B6** : `Durée totale des ateliers :`
- **Cellule C6** : *(laisser vide — sera remplie par la macro)*
- **Cellule B7** : `Nombre de participants :`
- **Cellule C7** : *(laisser vide — sera remplie par la macro)*
- **Cellule B8** : `Nombre de participants pro :`
- **Cellule C8** : *(laisser vide — sera remplie par la macro)*

---

## Étape 5 — Ouvrir l'éditeur VBA

1. Appuyez sur **Alt + F11** pour ouvrir l'éditeur VBA (Visual Basic for Applications)
2. Dans le menu **Outils → Références**, vérifiez que les références suivantes sont cochées :
   - `Microsoft Excel xx.x Object Library` (déjà coché)
   - `Microsoft Forms 2.0 Object Library` (nécessaire pour les UserForms)
3. Cliquez sur **OK**

---

## Étape 6 — Importer les modules `.bas`

Pour chaque fichier `.bas` du dossier `src/` :

1. Dans l'éditeur VBA, menu **Fichier → Importer un fichier...**
2. Naviguez jusqu'au dossier `src/` du projet
3. Importez les fichiers dans cet ordre :
   - `ModuleStats.bas`
   - `ModuleAteliers.bas`
   - `ModuleParticipants.bas`
   - `ModulePresences.bas`
   - `ModuleImport.bas`

### Pour `ThisWorkbook.bas`

Ce fichier correspond au module `ThisWorkbook` (pas un module standard) :

1. Dans l'explorateur de projets VBA (à gauche), double-cliquez sur **ThisWorkbook**
2. Ouvrez le fichier `src/ThisWorkbook.bas` avec un éditeur de texte (Bloc-notes, VS Code...)
3. **Copiez tout le contenu** et **collez-le** dans la fenêtre de code de `ThisWorkbook`

---

## Étape 7 — Créer et importer les UserForms

Pour chaque formulaire (`.frm`), vous devez :

1. **Créer le UserForm** dans l'éditeur VBA :
   - Menu **Insertion → UserForm**
   - Dans la fenêtre Propriétés, changez le **Name** du formulaire selon le tableau ci-dessous
2. **Ajouter les contrôles** selon la description dans `formulaires.md`
3. **Coller le code VBA** depuis le fichier `.frm` correspondant :
   - Double-cliquez sur le formulaire pour accéder au code
   - Copiez-collez le contenu du fichier `.frm`

| Fichier `.frm` | Name du UserForm |
|---|---|
| `FrmNouvelAtelier.frm` | `FrmNouvelAtelier` |
| `FrmSaisirPresences.frm` | `FrmSaisirPresences` |
| `FrmNouveauParticipant.frm` | `FrmNouveauParticipant` |
| `FrmGererParticipants.frm` | `FrmGererParticipants` |

> **Note** : Le détail des contrôles à ajouter sur chaque formulaire est décrit dans `formulaires.md`.

---

## Étape 8 — Ajouter les boutons sur la feuille ACCUEIL

1. Dans Excel, activez l'onglet **Développeur** si ce n'est pas déjà fait :
   - **Fichier → Options → Personnaliser le ruban** → cochez `Développeur`
2. Allez sur la feuille `ACCUEIL`
3. Dans l'onglet **Développeur → Insérer → Contrôles de formulaire → Bouton**
4. Dessinez 4 boutons sur la feuille et assignez les macros suivantes :

| Texte du bouton | Macro à assigner |
|---|---|
| ➕ Nouvel Atelier | `OuvrirFrmNouvelAtelier` |
| 📋 Saisir les Présences | `OuvrirFrmSaisirPresences` |
| 👤 Nouveau Participant | `OuvrirFrmNouveauParticipant` |
| 👁️ Gérer les Participants | `OuvrirFrmGererParticipants` |

Pour assigner une macro à un bouton : **clic droit sur le bouton → Affecter une macro → choisir la macro → OK**

---

## Étape 9 — Ajouter le bouton sur la feuille IMPORT

1. Allez sur la feuille `IMPORT`
2. Ajoutez un bouton "Importer les données"
3. Assignez la macro : `ImporterDonnees`

---

## Étape 10 — Protection des feuilles

Pour protéger les feuilles (les données ne doivent se saisir que via les formulaires) :

1. Allez sur la feuille `ATELIERS`
2. **Révision → Protéger la feuille**
3. Décochez tout sauf "Sélectionner les cellules verrouillées"
4. Définissez un mot de passe si souhaité (notez-le précieusement)
5. Cliquez **OK**
6. Répétez pour `PARTICIPANTS`, `PRESENCES`, `IMPORT`

> **Note** : Le code VBA déprotège et reprotège automatiquement les feuilles lors des opérations d'écriture. Si vous définissez un mot de passe, vous devrez le mettre à jour dans le code VBA (variable `MOT_DE_PASSE` dans `ThisWorkbook.bas`).

---

## Étape 11 — Test final

1. Fermez et rouvrez le fichier `.xlsm`
2. Acceptez d'activer les macros si demandé
3. La feuille `ACCUEIL` doit s'afficher avec les statistiques (toutes à 0 si aucune donnée)
4. Testez chaque bouton :
   - ➕ Nouvel Atelier → le formulaire s'ouvre
   - 📋 Saisir les Présences → le formulaire s'ouvre
   - 👤 Nouveau Participant → le formulaire s'ouvre
   - 👁️ Gérer les Participants → le formulaire s'ouvre

---

## Structure des fichiers du projet

```
gestion-coworking/
├── SETUP.md                    ← Ce guide
├── formulaires.md              ← Structure des UserForms
└── src/
    ├── ThisWorkbook.bas        ← Code du module ThisWorkbook
    ├── ModuleStats.bas         ← Calcul des statistiques
    ├── ModuleAteliers.bas      ← Gestion des ateliers
    ├── ModuleParticipants.bas  ← Gestion des participants
    ├── ModulePresences.bas     ← Gestion des présences
    ├── ModuleImport.bas        ← Import (préparé pour usage futur)
    ├── FrmNouvelAtelier.frm    ← Formulaire Nouvel Atelier
    ├── FrmSaisirPresences.frm  ← Formulaire Saisir Présences
    ├── FrmNouveauParticipant.frm ← Formulaire Nouveau Participant
    └── FrmGererParticipants.frm  ← Formulaire Gérer Participants
```

---

## En cas de problème

- **Les macros ne s'exécutent pas** : Vérifiez que le fichier est bien enregistré en `.xlsm` et que les macros sont activées (message de sécurité en haut du fichier)
- **Erreur "Objet introuvable"** : Vérifiez que les noms des tableaux structurés sont bien `TblAteliers`, `TblParticipants`, `TblPresences`
- **Erreur sur un formulaire** : Vérifiez que les noms des contrôles correspondent exactement à ceux décrits dans `formulaires.md`
