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

#### Catégories
- **Nom du tableau** : `category`
- **Plage** : `A1:A12`
- **En-tête** : `Catégorie`
- **Contenu** :
  - Accessoire
  - Composant Interne
  - Connectique/Câblage
  - Consommable
  - Imprimante/Scanner
  - Logiciel/Licence
  - Matériel de Bureau
  - Matériel Mobile
  - Matériel Réseau
  - Périphérique
  - Stockage

#### Matériel de Bureau
- **Nom du tableau** : `office_equipment`
- **Plage** : `C1:C6`
- **En-tête** : `Matériel de bureau`
- **Contenu** :
  - Écran et moniteur
  - Ordinateur fixe
  - Ordinateur portable
  - Station de travail
  - Vidéoprojecteur

#### Imprimante et Scanner
- **Nom du tableau** : `printer_scanner`
- **Plage** : `E1:E5`
- **En-tête** : `Imprimante et scanner`
- **Contenu** :
  - Imprimante jet d'encre
  - Imprimante laser
  - Imprimante multifonction
  - Scanner

#### Composant interne
- **Nom du tableau** : `internal_component`
- **Plage** : `G1:G7`
- **En-tête** : `Composant interne`
- **Contenu** :
  - Alimentation électrique (PSU)
  - Boîtier
  - Carte graphique (GPU)
  - Carte mère
  - Mémoire vive (RAM)
  - Processeur (CPU)

#### Périphérique
- **Nom du tableau** : `peripheral`
- **Plage** : `I1:I6`
- **En-tête** : `Périphérique`
- **Contenu** :
  - Casque
  - Clavier
  - Microphone
  - Souris
  - Webcam

#### Matériel réseau
- **Nom du tableau** : `network_hardware`
- **Plage** : `K1:K5`
- **En-tête** : `Matériel réseau`
- **Contenu** :
  - Carte réseau
  - Commutateur
  - Point d'accès
  - Routeur

#### Stockage
- **Nom du tableau** : `storage`
- **Plage** : `M1:M7`
- **En-tête** : `Stockage`
- **Contenu** :
  - Carte mémoire
  - Clé USB
  - Disque dur interne (HDD)
  - Disque externe
  - Disque SSD interne
  - Serveur NAS

#### Connectique et Câblage
- **Nom du tableau** : `connector_cabling`
- **Plage** : `O1:O5`
- **En-tête** : `Connectique et Câblage`
- **Contenu** :
  - Adaptateur et convertisseur
  - Câble de donnée
  - Câble réseau
  - Câble vidéo

#### Accessoire
- **Nom du tableau** : `accessorie`
- **Plage** : `Q1:Q6`
- **En-tête** : `Accessoire`
- **Contenu** :
  - Batterie et chargeur
  - Onduleur (UPS)
  - Outil et kit de nettoyage
  - Pile et accumulateur
  - Station d'accueil

#### Consommable
- **Nom du tableau** : `consumable`
- **Plage** : `S1:S4`
- **En-tête** : `Consommable`
- **Contenu** :
  - Cartouche d'encre
  - Papier
  - Toner

#### Logiciel et Licence
- **Nom du tableau** : `software_licence`
- **Plage** : `U1:U5`
- **En-tête** : `Logiciel et Licence`
- **Contenu** :
  - Logiciel de sécurité
  - Logiciel métier
  - Suite bureautique
  - Système d'exploitation

#### Matériel Mobile
- **Nom du tableau** : `mobile_hardware`
- **Plage** : `W1:W4`
- **En-tête** : `Matériel mobile`
- **Contenu** :
  - Smartphone
  - Smartwatch
  - Tablette

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