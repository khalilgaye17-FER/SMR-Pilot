# Analyse Détaillée du Code - SMR PILOT 2F

## Vue d'Ensemble du Projet

**Type de fichier**: Microsoft Excel 2007+ avec macros (`.xlsm`)
**Taille**: 476 KB
**Auteur original**: pgaye (Keolis)
**Domaine**: Gestion de Maintenance des Services Réseau (SMR - Service Maintenance Réseau)

---

## 1. ARCHITECTURE GLOBALE

### 1.1 Structure des Feuilles Excel (21 feuilles)

| Catégorie | Feuilles | Fonction |
|-----------|----------|----------|
| **Interface Principale** | POV | Point de Vue principal / Dashboard |
| **Procédures** | Procedures_Master, P001_Logigramme, Historique_Procedures | Gestion des procédures opérationnelles |
| **Données de base** | Data, Voies, Itineraires, Voie_Compositions, Liste_Elements | Configuration et données de référence |
| **Opérations** | Remise_Service, Log_NotesService, Log_Derangements | Suivi opérationnel |
| **Ressources** | Agents, Dependances, Catalogue_Travaux | Gestion des ressources |
| **Historique** | Historique_OL, Historique_Travaux, Historique_Manoeuvres | Traçabilité des opérations |
| **Archives** | Archive_Manoeuvres, Archive_Travaux, Archive_OL | Conservation des données |

### 1.2 Modules VBA (Architecture)

```
SMR_PILOT/
├── ThisWorkbook (Module de Classeur)
│   └── Workbook_Open() - Point d'entrée de l'application
│
├── Modules de Classe
│   ├── CSegment - Représentation d'un segment d'itinéraire
│   └── Feuilles (Feuil1-Feuil15) - Code événementiel par feuille
│
├── Modules Métier
│   ├── Mod_Global - Variables et constantes globales
│   ├── Mod_LogiqueApp - Logique applicative principale
│   ├── Mod_Dashboard - Gestion du tableau de bord
│   ├── GestionDonnees - CRUD sur les données
│   ├── GestionAffichage - Manipulation visuelle
│   ├── GestionSecurite - Règles de sécurité et validations
│   ├── MoteurItineraires - Calcul d'itinéraires
│   └── Manoeuvres - Gestion des manœuvres de rames
│
└── UserForms (Interfaces utilisateur)
    ├── UserForm_Connexion - Authentification
    ├── UserForm_CreerItineraire - Création/modification d'itinéraires
    └── UserForm_PlacerRame - Placement initial des rames
```

---

## 2. ANALYSE DES MODULES CLÉS

### 2.1 ThisWorkbook - Point d'Entrée

**Flux de démarrage**:
```vba
Workbook_Open()
  1. Application.Visible = False (Masque Excel)
  2. Réinitialisation des variables globales
  3. Affichage du formulaire de connexion (UserForm_Connexion.Show)
  4. Si connexion réussie → LancerPriseDeService()
```

**Points positifs**:
- Processus de démarrage sécurisé
- Gestion de l'authentification avant l'accès aux fonctionnalités

**Points d'attention**:
- Pas de gestion d'erreur visible
- Application.Visible = False peut dérouter les utilisateurs

### 2.2 Module GestionDonnees - Couche d'Accès aux Données

**Fonctions principales**:

| Fonction | Rôle | Colonne Data |
|----------|------|--------------|
| `LireEtatElement(nomElement)` | Lit l'état d'un élément (voie/aiguille) | B |
| `LireEtatOccupation(nomElement)` | Lit spécifiquement l'état d'occupation | B |
| `ModifierEtatElement(nom, etat)` | Modifie l'état d'un élément | B |
| `MettreAJourOccupation(voie, rame)` | Associe une rame à une voie | C |
| `MettreAJourPositionRame(rame, position)` | Met à jour la position d'une rame | F |
| `AjouterProtection(element, type)` | Ajoute une protection (Travaux, Dérangement, etc.) | D |
| `RetirerProtection(element, type)` | Retire une protection | D |
| `CompterProtections(element)` | Compte les protections actives | D |
| `LireProtections(element)` | Retourne la chaîne complète des protections | D |

**Qualité du code**:
✅ **Points forts**:
- Nettoyage systématique avec `Trim()` sur toutes les entrées
- Recherche insensible à la casse (`MatchCase:=False`)
- Gestion des protections multiples avec séparateur virgule
- Évite les doublons lors de l'ajout de protections

⚠️ **Améliorations possibles**:
- Pas de gestion d'erreur si la feuille "Data" n'existe pas
- Constante `NOM_FEUILLE_DATA` définie mais pourrait être centralisée
- Aucune validation des valeurs d'état avant écriture

### 2.3 Module GestionSecurite - Logique de Sécurité Complexe

**Fonctions de validation**:

#### A) `ValiderCheminLibre(chemin)`
**Rôle**: Valide qu'un itinéraire est praticable
```vba
Pour chaque élément (SAUF LE PREMIER):
  1. Vérification Etat = "Libre"
  2. Vérification NbProtections = 0
  → Bloque si une condition échoue
```

#### B) `VerifierDependancesZone(origine, destination)`
**Rôle**: Applique les règles métier d'interdiction de zone (Règle N°2)
- Lit la feuille "Dependances"
- Bloque si un élément déclencheur interdit l'accès à la destination

#### C) `VerifierDependancesDirectionnelles(origineMouvement, voiePrincipale, destinationFinale)`
**Rôle**: Gère les règles directionnelles (ex: VSMR1 → interdiction sur certaines voies)
```
1. Détermine la direction d'origine (VSMR1, VTB_VTC, etc.)
2. Parcourt les règles de type "Interdire_Depuis_XXX"
3. Bloque si la destination finale est dans la liste des éléments impactés
```

#### D) `EstCheminLineaireLibre(origine, destination)`
**Rôle**: Validation silencieuse pour mouvements internes sur une même voie
- Vérifie que tous les éléments intermédiaires sont "Libre"
- Vérifie qu'il n'y a pas de protections

#### E) `VerifierDisponibiliteElements(elementsStr)`
**Rôle**: Validation pour mise en travaux
- Bloque si éléments en état "Verrouille" ou "Transit"
- Autorise si "Libre" ou "Occupé"

#### F) `EstZoneAccessible(destination, origineMouvement)`
**Rôle**: Version silencieuse de VerifierDependancesZone
- **Innovation**: Ne bloque PAS si le déclencheur est l'origine elle-même (évite l'auto-blocage)

**Qualité**:
✅ **Excellente logique métier**:
- Séparation claire des responsabilités
- Gestion de cas complexes (mouvements linéaires vs avec transit)
- Protection contre l'auto-blocage

⚠️ **Complexité élevée**:
- Couplage fort avec la structure de la feuille "Dependances"
- Difficile à tester unitairement
- Messages d'erreur en dur (internationalisation impossible)

### 2.4 Module MoteurItineraires - Moteur de Calcul

**Fonction phare**: `TrouverTransitsPossibles(origine, destinationFinale)`

**Algorithme** (simplifié):
```
1. Identifier la voie d'origine et de destination
2. Pour chaque itinéraire sortant de la voie d'origine:
   a. Vérifier chemin interne vers point de sortie
   b. Vérifier disponibilité du point de sortie
   c. Vérifier disponibilité des aiguilles de sortie
   d. Chercher une connexion vers la voie de destination
   e. Vérifier chemin d'entrée et aiguilles d'entrée
   f. Vérifier destination finale accessible
   g. Si tout OK → Ajouter le transit à la collection
3. Retourner la collection de transits valides
```

**Autres fonctions clés**:
- `GetVoieInfoFromEmplacement(ws, emplacement, nomVoie, sequence)`: Extrait infos d'une voie
- `FindSegmentItineraire(ws, nomVoie, transit)`: Trouve un segment dans la table Itinéraires
- `GetSegmentInterne(origine, dest, sequenceVoie)`: Calcule le chemin interne sur une voie
- `ExisteUneSortieValide(position)`: Vérifie si une rame peut bouger

**Qualité**:
✅ **Architecture solide**:
- Séparation logique vs données
- Utilisation de classe `CSegment` pour encapsuler les chemins

⚠️ **Maintenabilité**:
- Fonction `TrouverTransitsPossibles` très longue (>100 lignes)
- Logique imbriquée difficile à suivre
- Dépendance forte sur la structure de la feuille "Itineraires"

### 2.5 Module Manoeuvres - Orchestration des Opérations

**Fonction principale**: `DemarrerManoeuvre_Selectionnee()`

**Machine à états**:
```
VERS_TRANSIT
  → [CLIC 1] Verrouillage itinéraire
  → EN_ROUTE_VERS_TRANSIT

EN_ROUTE_VERS_TRANSIT
  → [CLIC 2] Déplacement + Libération origine + Recalcul chemin
  → VERS_DESTINATION_FINALE ou VERS_DESTINATION_LINÉAIRE

VERS_DESTINATION_FINALE / VERS_DESTINATION_LINÉAIRE
  → [CLIC 3] Verrouillage final
  → EN_ROUTE_VERS_DESTINATION

EN_ROUTE_VERS_DESTINATION
  → [CLIC 4] Arrivée finale + Libération complète
  → COMPLETE
```

**Points positifs**:
- Utilisation d'une machine à états claire
- Séparation verrouillage / déplacement
- Recalcul dynamique des chemins

**Points d'attention**:
- Code répétitif (verrouillage/déverrouillage)
- Pas de mécanisme de rollback en cas d'erreur
- Dépendance sur les clics utilisateur (pas d'automatisation possible)

### 2.6 Module GestionAffichage - Interface Visuelle

**Constantes de couleurs**:
```vba
COULEUR_LIBRE = RGB(211, 211, 211)      ' Gris
COULEUR_OCCUPE = RGB(255, 0, 0)         ' Rouge
COULEUR_TRANSIT = RGB(255, 255, 0)      ' Jaune
COULEUR_ITINERAIRE = vbBlue             ' Bleu
COULEUR_TRAVAUX = RGB(255, 0, 255)      ' Magenta
```

**Fonctions**:
- `ChangerCouleurVoie(nom, couleur)`: Modifie la couleur d'une forme Excel
- `AfficherTexteSurVoie(nom, texte)`: Affiche le nom d'une rame sur la voie
- `RestaurerCouleurApresProtection(nomElement)`: Logique intelligente de restauration

**Innovation**:
La fonction `RestaurerCouleurApresProtection` gère correctement la superposition de protections multiples (ex: Travaux + Dérangement).

### 2.7 UserForms - Interfaces Utilisateur

#### UserForm_CreerItineraire

**Modes**:
1. **Mode Création**: Préparer une nouvelle manœuvre
2. **Mode Modification**: Changer la destination d'une manœuvre existante

**Logique de validation complexe**:
```vba
RemplirDestinationsLibres():
  Pour chaque emplacement:
    1. Est "Libre" ET non protégé ET non-Aiguille?
    2. Zones accessibles (règles de dépendance)?
    3. SI même voie:
         → Vérifier chemin linéaire libre
       SINON:
         → Vérifier qu'il existe AU MOINS un transit valide
    4. SI OK → Ajouter à la liste
```

**Points forts**:
- Filtrage intelligent des destinations impossibles
- Évite que l'utilisateur sélectionne une destination invalide
- Utilisation d'un interrupteur `IsInitializing` pour éviter les événements en cascade

**Points faibles**:
- Code très long (>300 lignes)
- Mélange logique métier et interface
- Difficulté de maintenance

---

## 3. PATTERNS ET PRATIQUES DE DÉVELOPPEMENT

### 3.1 Patterns Identifiés

| Pattern | Utilisation | Exemple |
|---------|-------------|---------|
| **Module Pattern** | Séparation des responsabilités | GestionDonnees, GestionSecurite, etc. |
| **State Machine** | Gestion des manœuvres | États: VERS_TRANSIT, EN_ROUTE, COMPLETE |
| **Data Transfer Object** | Classe CSegment | Encapsule origine, destination, elements |
| **Flag Guards** | Éviter événements en cascade | `IsInitializing` dans UserForms |
| **Soft Delete** | Archivage au lieu de suppression | Archive_Manoeuvres, Archive_Travaux |

### 3.2 Conventions de Code

✅ **Bonnes pratiques observées**:
- `Option Explicit` dans tous les modules
- Nettoyage systématique avec `Trim()`
- Recherches insensibles à la casse
- Noms de variables descriptifs (français)
- Commentaires explicatifs dans les sections complexes

⚠️ **Conventions incohérentes**:
- Mélange de français et anglais (`Set ws`, `derniereLigne`)
- Indentation variable
- Pas de gestion d'erreur globale (`On Error Goto`)

---

## 4. ANALYSE DE SÉCURITÉ

### 4.1 Sécurité Opérationnelle (Métier)

✅ **Excellente couverture**:
- Multiples couches de validation avant chaque opération
- Règles de dépendance configurables (feuille Dependances)
- Gestion des protections multiples simultanées
- Validation en amont pour éviter les erreurs utilisateur

### 4.2 Sécurité Informatique

⚠️ **Risques identifiés**:

1. **Pas de gestion d'erreur**
   - Aucun `On Error Resume Next` ou `On Error Goto`
   - Un crash peut laisser les données dans un état incohérent

2. **Pas de transactions**
   - Opérations multi-étapes sans rollback
   - Exemple: Si le verrouillage échoue à mi-chemin, état corrompu

3. **Authentification basique**
   - UserForm_Connexion (code non fourni, mais probablement simple)
   - Pas de chiffrement visible

4. **Injection possible**
   - Utilisation de `Range().Find()` sans validation
   - Risque faible mais existe

5. **Macros non signées**
   - Fichier .xlsm nécessite d'activer les macros
   - Risque de malware si fichier modifié

### 4.3 Intégrité des Données

⚠️ **Points de vigilance**:
- Aucune validation de type (un texte peut être mis dans une colonne numérique)
- Pas de contraintes d'intégrité référentielle
- Dépendance sur les noms de feuilles (renommer "Data" casse l'app)

---

## 5. PERFORMANCE

### 5.1 Points de Performance

**Optimisations présentes**:
- Utilisation de `Set ws = ...` pour éviter les recherches répétées
- `Exit For` dès qu'un élément est trouvé
- Constantes pour les noms de feuilles

**Problèmes de performance**:

1. **Boucles imbriquées dans TrouverTransitsPossibles**
   - Complexité O(n²) voire O(n³)
   - Peut être lent avec beaucoup d'itinéraires

2. **Recherches linéaires répétées**
   - `Range().Find()` sur toute la colonne à chaque appel
   - Solution: Charger en mémoire dans un Dictionary

3. **Pas de cache**
   - Recalcul complet à chaque fois
   - Exemple: Liste des destinations recalculée à chaque ouverture du formulaire

### 5.2 Estimation de Scalabilité

| Dimension | Limite estimée | Raison |
|-----------|---------------|--------|
| Nombre de rames | ~50 | Performance des boucles |
| Nombre de voies | ~100 | Complexité de recherche |
| Nombre d'itinéraires | ~200 | TrouverTransitsPossibles |
| Manœuvres simultanées | ~20 | Interface utilisateur |

---

## 6. MAINTENABILITÉ

### 6.1 Score de Maintenabilité: ⭐⭐⭐ (3/5)

**Forces**:
- Architecture modulaire claire
- Séparation des responsabilités
- Code bien commenté dans les parties complexes

**Faiblesses**:
- Couplage fort entre modules
- Fonctions très longues (>100 lignes)
- Logique métier mélangée avec l'UI
- Pas de tests unitaires

### 6.2 Dette Technique Identifiée

| Catégorie | Dette | Impact | Effort |
|-----------|-------|--------|--------|
| **Code Dupliqué** | Logique de verrouillage répétée 3-4 fois | Moyen | Faible |
| **Fonctions Longues** | `TrouverTransitsPossibles` (>150 lignes) | Élevé | Moyen |
| **Couplage** | Dépendance sur structure exacte des feuilles | Élevé | Élevé |
| **Pas de Tests** | Aucun test automatisé | Critique | Élevé |
| **Gestion d'Erreur** | Absente dans 90% du code | Critique | Moyen |

---

## 7. ANALYSE FONCTIONNELLE

### 7.1 Cas d'Usage Principaux

1. **Prise de service**
   - Connexion agent
   - Initialisation du POV
   - Affichage du dashboard

2. **Placer une rame**
   - Sélection rame au garage
   - Sélection voie libre
   - Mise à jour affichage

3. **Créer une manœuvre**
   - Sélection rame positionnée
   - Calcul destinations possibles
   - Calcul chemins et transits
   - Préparation de la manœuvre

4. **Exécuter une manœuvre** (machine à 4 états)
   - Clic 1: Verrouillage vers transit
   - Clic 2: Déplacement au transit
   - Clic 3: Verrouillage vers destination
   - Clic 4: Arrivée finale

5. **Gestion des travaux/dérangements**
   - Protection d'éléments
   - Blocage des itinéraires affectés
   - Restauration après fin

### 7.2 Règles Métier Complexes

**Règle N°1**: Mouvements linéaires vs avec transit
- Si origine et destination sur même voie → Linéaire (direct)
- Sinon → Nécessite un transit

**Règle N°2**: Dépendances de zone
- Exemple: "Si V17B est Occupé, alors V17A est interdit comme DESTINATION"
- Configurable dans feuille "Dependances"

**Règle N°3**: Dépendances directionnelles
- Exemple: "Depuis VSMR1, interdit d'aller vers V17A si V17B est occupé"
- Basé sur la direction d'origine (VSMR1, VTB_VTC, etc.)

**Règle N°4**: Protections multiples
- Un élément peut avoir plusieurs protections simultanées
- La couleur change seulement quand TOUTES les protections sont levées

---

## 8. POINTS FORTS DU PROJET

1. **Modélisation métier excellente**
   - Capture fidèle des règles ferroviaires/de dépôt
   - Sécurité opérationnelle prioritaire

2. **Architecture modulaire**
   - Séparation claire des couches
   - Réutilisabilité du code

3. **Expérience utilisateur**
   - Filtrage intelligent (pas de choix invalides)
   - Retour visuel immédiat (couleurs)
   - Guidage par machine à états

4. **Flexibilité**
   - Règles configurables via feuille "Dependances"
   - Itinéraires configurables via feuille "Itineraires"

5. **Traçabilité**
   - Historiques et archives
   - Logs d'événements

---

## 9. FAIBLESSES ET RISQUES

### 9.1 Risques Critiques

1. **Absence de gestion d'erreur**
   - Risque: Corruption de données en cas de crash
   - Impact: Élevé
   - Probabilité: Moyenne

2. **Pas de sauvegarde/rollback**
   - Risque: Impossible d'annuler une opération
   - Impact: Élevé
   - Probabilité: Faible

3. **Couplage aux noms de feuilles/colonnes**
   - Risque: Renommer une feuille casse l'application
   - Impact: Critique
   - Probabilité: Faible

### 9.2 Risques Moyens

4. **Performance dégradée à grande échelle**
   - Risque: Lenteur avec >100 voies
   - Impact: Moyen
   - Probabilité: Moyenne

5. **Difficulté de maintenance**
   - Risque: Modification introduit des bugs
   - Impact: Moyen
   - Probabilité: Élevée

6. **Pas de versioning des données**
   - Risque: Perte d'historique en cas de corruption
   - Impact: Moyen
   - Probabilité: Faible

---

## 10. RECOMMANDATIONS

### 10.1 Corrections Urgentes (Priorité 1)

1. **Ajouter gestion d'erreur globale**
   ```vba
   On Error GoTo GestionErreur
   ' Code métier
   Exit Sub

   GestionErreur:
       MsgBox "Erreur: " & Err.Description
       ' Rollback si possible
       Err.Clear
   ```

2. **Centraliser les constantes**
   - Créer module `Mod_Constantes` avec noms de feuilles, colonnes, couleurs
   - Un seul point de modification

3. **Valider l'existence des feuilles au démarrage**
   ```vba
   Function VerifierStructure() As Boolean
       ' Vérifier que toutes les feuilles requises existent
   End Function
   ```

### 10.2 Améliorations Court Terme (Priorité 2)

4. **Refactoriser les fonctions longues**
   - Extraire `TrouverTransitsPossibles` en sous-fonctions
   - Maximum 50 lignes par fonction

5. **Implémenter un système de log**
   ```vba
   Sub EcrireLog(message As String)
       ' Ajouter dans feuille LOG avec timestamp
   End Sub
   ```

6. **Créer une couche d'abstraction pour les données**
   ```vba
   Class CDataAccess
       Function GetElement(nom) As CElement
       Function SaveElement(element As CElement)
   End Class
   ```

### 10.3 Évolutions Moyen Terme (Priorité 3)

7. **Migrer vers une base de données**
   - Remplacer feuilles Data/Itineraires par Access ou SQLite
   - Intégrité référentielle
   - Transactions

8. **Implémenter un cache**
   ```vba
   Private dictCache As New Dictionary
   Function GetElementCached(nom) As String
       If Not dictCache.Exists(nom) Then
           dictCache(nom) = GetElementFromSheet(nom)
       End If
       GetElementCached = dictCache(nom)
   End Function
   ```

9. **Ajouter des tests unitaires**
   - Framework: [Rubberduck](https://rubberduckvba.com/)
   - Tester au minimum les fonctions de GestionSecurite

### 10.4 Évolutions Long Terme (Priorité 4)

10. **Réécrire en application moderne**
    - Option 1: Excel + C# VSTO
    - Option 2: Web App (React + Node.js + PostgreSQL)
    - Option 3: Application desktop (Python + PyQt + SQLite)

11. **Séparation UI / Logique**
    - Modèle: Clean Architecture
    - Interface indépendante de la logique métier

12. **Gestion des utilisateurs robuste**
    - Base utilisateurs centralisée
    - Permissions par rôle
    - Audit trail

---

## 11. MÉTRIQUES DE CODE

| Métrique | Valeur | Évaluation |
|----------|--------|------------|
| **Nombre de modules** | ~15 | ✅ Correct |
| **Nombre de UserForms** | ~3 | ✅ Correct |
| **Lignes de code (estimées)** | ~3000-4000 | ⚠️ Conséquent |
| **Fonction la plus longue** | ~300 lignes | ❌ Trop long |
| **Modules avec Option Explicit** | 100% | ✅ Excellent |
| **Couverture de tests** | 0% | ❌ Critique |
| **Gestion d'erreur** | <10% | ❌ Insuffisant |
| **Commentaires** | ~20% | ✅ Acceptable |
| **Complexité cyclomatique max** | ~30+ | ❌ Élevée |

---

## 12. CONCLUSION

### 12.1 Synthèse

Ce projet représente une **application métier mature et fonctionnelle** qui modélise fidèlement des processus opérationnels complexes (gestion de dépôt ferroviaire/tramway).

**Forces principales**:
- Architecture modulaire
- Logique métier riche et pertinente
- Sécurité opérationnelle (métier) excellente

**Faiblesses principales**:
- Absence de gestion d'erreur
- Maintenabilité limitée par la longueur des fonctions
- Couplage fort à la structure Excel

### 12.2 Verdict Global

**Note globale**: ⭐⭐⭐⭐ (4/5)

| Critère | Note | Commentaire |
|---------|------|-------------|
| **Fonctionnalité** | ⭐⭐⭐⭐⭐ | Couvre tous les besoins métier |
| **Architecture** | ⭐⭐⭐⭐ | Bonne séparation, peut être améliorée |
| **Qualité du code** | ⭐⭐⭐ | Correct mais manque de rigueur |
| **Sécurité (métier)** | ⭐⭐⭐⭐⭐ | Excellente validation |
| **Sécurité (info)** | ⭐⭐ | Nombreuses lacunes |
| **Performance** | ⭐⭐⭐ | Acceptable pour usage actuel |
| **Maintenabilité** | ⭐⭐⭐ | Difficile mais faisable |
| **Evolutivité** | ⭐⭐ | Limitée par Excel |

### 12.3 Recommandation Finale

**Pour un usage de production immédiat**:
- ✅ **Acceptable** avec les corrections Priorité 1 (gestion d'erreur)

**Pour un usage long terme**:
- ⚠️ **Nécessite refactoring** (Priorités 1-3)

**Pour une évolution majeure**:
- 🔄 **Envisager migration** vers technologie moderne (Priorité 4)

---

## ANNEXES

### A. Glossaire des Termes Métier

- **SMR**: Service Maintenance Réseau
- **POV**: Point de Vue (Dashboard principal)
- **Rame**: Véhicule ferroviaire (train, tramway)
- **Manœuvre**: Déplacement d'une rame entre deux positions
- **Transit**: Point intermédiaire lors d'un changement de voie
- **Itinéraire**: Chemin défini entre deux voies (avec aiguilles)
- **Dérangement**: Incident bloquant un élément
- **Protection**: Verrouillage temporaire d'un élément

### B. Structure des Données Principales

**Feuille "Data"**:
```
Colonne A: Nom_Element (voies, aiguilles)
Colonne B: Etat_Occupation (Libre, Occupé, Transit, Verrouille)
Colonne C: Rame_Presente
Colonne D: Protections (CSV: "ZEP1,DER")
Colonne E: Nom_Rame
Colonne F: Position_Rame
```

**Feuille "Itineraires"**:
```
Colonne A: Voie_Origine
Colonne B: Voie_Destination
Colonne C: Chemin (CSV d'éléments)
```

**Feuille "Dependances"**:
```
Colonne A: Element_Declencheur
Colonne B: Etat_Declencheur
Colonne C: Elements_Impactes (CSV)
Colonne D: Action (ex: "Interdire_Mouvement", "Interdire_Depuis_VSMR1")
```

### C. Machine à États des Manœuvres

```mermaid
INIT
  ↓
VERS_TRANSIT
  ↓ (Clic 1: Verrouillage)
EN_ROUTE_VERS_TRANSIT
  ↓ (Clic 2: Déplacement)
VERS_DESTINATION_FINALE / VERS_DESTINATION_LINÉAIRE
  ↓ (Clic 3: Verrouillage final)
EN_ROUTE_VERS_DESTINATION
  ↓ (Clic 4: Arrivée)
COMPLETE
```

---

**Analyse réalisée le**: 2025-11-15
**Outil**: olevba 0.60.2
**Analyseur**: Claude Code (Sonnet 4.5)
