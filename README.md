# 🚀 Projet VBA : Gestion de Stock Informatique

## Introduction
Ce projet VBA est une solution simple et efficace pour la gestion du stock d'un service informatique. Il permet le suivi des équipements, matériels et consommables, depuis leur entrée jusqu'à leur utilisation ou sortie.

Le système est structuré autour de trois feuilles de calcul principales :
| Feuille          | Rôle                                                                                     |
|------------------|------------------------------------------------------------------------------------------|
| **stock**        | Inventaire permanent de tous les articles : visualisation des quantités et seuils critiques. |
| **mouvement**    | Historique des transactions : entrées et sorties de matériel.                            |
| **configuration**| Tables de référence pour standardiser les données (catégories, sous-catégories).         |

---

## Guide de Configuration du Classeur Excel

### 1. Création du classeur et des feuilles
1. Créez un nouveau classeur Excel.
2. Enregistrez-le au format **.xlsm** (macros activées).
3. Renommez les trois premières feuilles :
   - `stock`
   - `mouvement`
   - `configuration`

---

### 2. Configuration de la feuille `stock`
- **Cellule A1** : `=AUJOURDHUI()`
- Sélectionnez la plage **A2:I2**, mettez-la sous forme de tableau et nommez-le **`stock`**.
- **En-têtes de colonne** :
  | libellé          | stock | catégorie        | maj         | seuil | sous-catégorie   | commentaire | ligne_tableau | ligne_feuille |
  |------------------|-------|------------------|-------------|-------|------------------|-------------|---------------|---------------|
  | (Texte)          | (Nombre) | (Texte)          | (Date courte) | (Nombre) | (Texte)          | (Texte)     | (Nombre)      | (Nombre)      |
- Masquez la **ligne 1**.
---

### 3. Configuration de la feuille `mouvement`
- Sélectionnez la plage **A2:E2**, mettez-la sous forme de tableau et nommez-le **`movement`**.
- **En-têtes de colonne** :
  | date         | type  | valeur | description | matériel      |
  |--------------|-------|--------|-------------|---------------|
  | (Date courte)| (Texte)| (Nombre)| (Texte)     | (Texte)       |

---

### 4. Configuration de la feuille `configuration`

#### Catégorie
- **Nom du tableau** : `category`
- **Plage** : `A1:A12`
- **En-tête** : `Catégorie`
- **Contenu** :
  - accessoire
  - composant interne
  - connectique/câblage
  - consommable
  - imprimante/scanner
  - logiciel/licence
  - matériel de bureau
  - matériel mobile
  - matériel réseau
  - périphérique
  - stockage

#### matériel de bureau
- **Nom du tableau** : `office_equipment`
- **Plage** : `C1:C6`
- **En-tête** : `matériel de bureau`
- **Contenu** :
  - écran et moniteur
  - ordinateur fixe
  - ordinateur portable
  - station de travail
  - vidéoprojecteur

#### imprimante et scanner
- **Nom du tableau** : `printer_scanner`
- **Plage** : `E1:E5`
- **En-tête** : `imprimante et scanner`
- **Contenu** :
  - imprimante jet d'encre
  - imprimante laser
  - imprimante multifonction
  - scanner

#### composant interne
- **Nom du tableau** : `internal_component`
- **Plage** : `G1:G7`
- **En-tête** : `composant interne`
- **Contenu** :
  - alimentation électrique (PSU)
  - boîtier
  - carte graphique (GPU)
  - carte mère
  - mémoire vive (RAM)
  - processeur (CPU)

#### périphérique
- **Nom du tableau** : `peripheral`
- **Plage** : `I1:I6`
- **En-tête** : `périphérique`
- **Contenu** :
  - casque
  - clavier
  - microphone
  - souris
  - webcam

#### matériel réseau
- **Nom du tableau** : `network_hardware`
- **Plage** : `K1:K5`
- **En-tête** : `mtériel réseau`
- **Contenu** :
  - carte réseau
  - commutateur
  - point d'accès
  - routeur

#### stockage
- **Nom du tableau** : `storage`
- **Plage** : `M1:M7`
- **En-tête** : `stockage`
- **Contenu** :
  - carte mémoire
  - clé USB
  - disque dur interne (HDD)
  - disque externe
  - disque SSD interne
  - serveur NAS

#### connectique et câblage
- **Nom du tableau** : `connector_cabling`
- **Plage** : `O1:O5`
- **En-tête** : `connectique et câblage`
- **Contenu** :
  - Aaaptateur et convertisseur
  - câble de donnée
  - câble réseau
  - câble vidéo

#### accessoire
- **Nom du tableau** : `accessorie`
- **Plage** : `Q1:Q6`
- **En-tête** : `accessoire`
- **Contenu** :
  - batterie et chargeur
  - onduleur (UPS)
  - outil et kit de nettoyage
  - pile et accumulateur
  - station d'accueil

#### consommable
- **Nom du tableau** : `consumable`
- **Plage** : `S1:S4`
- **En-tête** : `consommable`
- **Contenu** :
  - cartouche d'encre
  - papier
  - toner

#### logiciel et licence
- **Nom du tableau** : `software_licence`
- **Plage** : `U1:U5`
- **En-tête** : `logiciel et licence`
- **Contenu** :
  - logiciel de sécurité
  - logiciel métier
  - suite bureautique
  - système d'exploitation

#### matériel mobile
- **Nom du tableau** : `mobile_hardware`
- **Plage** : `W1:W4`
- **En-tête** : `matériel mobile`
- **Contenu** :
  - smartphone
  - smartwatch
  - tablette

---
**Note** : Tous les tableaux de cette feuille "configuration" doivent être de type **Texte** et triés par ordre croissant (A → Z).

## Roadmap

- Validation des données des champs : Mise en place de contrôles pour garantir que les informations saisies (nombres, dates, etc.) sont correctes et au bon format.

- Fonctionnalités d'impression et d'export : Création de macros pour générer des rapports imprimables ou des documents PDF à partir des données de stock et de mouvement, ainsi que des graphiques.

- Gestion des thèmes : Implémentation d'un thème clair pour l'interface utilisateur (formulaires VBA), pour offrir une meilleure lisibilité par défaut.

- Export de graphiques : Ajout d'une fonctionnalité pour exporter les graphiques de suivi de stock en tant qu'images ou en format PDF.
---

## Tech Stack

* **Logiciel** : Microsoft Excel 2021
* **Langage** : Visual Basic for Applications (VBA)

---

### Assistance et Méthodologie

Ce projet a bénéficié de l'assistance de **Gemini** (assistant conversationnel développé par Google) pour :

* **Améliorer la gestion des erreurs** : Robustesse accrue grâce à des procédures de gestion des erreurs affinées et testées.
* **Documentation du code** : Création de commentaires clairs et de la documentation pour améliorer la lisibilité et la maintenabilité du programme.

---

### Création des procédures

* **Enregistreur de macros Excel** : L'enregistreur de macros a été utilisé comme base pour certaines procédures automatisées, ce qui a permis d'accélérer le développement initial. Ces procédures ont ensuite été affinées et adaptées manuellement pour répondre aux besoins spécifiques du projet.