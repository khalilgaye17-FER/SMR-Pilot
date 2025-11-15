# GUIDE DE MIGRATION - Corrections VBA
## Application des Corrections Critiques

**Fichier**: SMR_PILOT 2F.xlsm
**Date**: 2025-11-15
**Durée estimée**: 1h30

---

## 🎯 OBJECTIF

Appliquer les 8 corrections critiques identifiées dans le rapport de bugs pour:
- ✅ Éliminer les risques de crash
- ✅ Rendre l'application portable (fonctionne sur n'importe quelle machine)
- ✅ Améliorer la robustesse et la traçabilité des erreurs

---

## ⚡ QUICK START (Version Rapide)

**Vous êtes pressé? Voici le strict minimum:**

### 1. Sauvegarde (2 min)
```
1. Ouvrir SMR_PILOT 2F.xlsm
2. Fichier > Enregistrer sous > SMR_PILOT 2F - BACKUP.xlsm
3. Fermer le backup
```

### 2. Correction #4 - CRITIQUE (5 min)
**Le bug qui empêche le PDF de fonctionner sur d'autres machines**

```vba
Alt+F11 → Ouvrir module "Mod_Session" → Chercher "C:\Users\pgaye"

REMPLACER:
cheminBase = "C:\Users\pgaye\OneDrive - Keolis\Bureau\personnel PKG\GOSMR\SMR Pilot\SMRCC\Rapports_CTC\"

PAR:
cheminBase = ThisWorkbook.Path & "\Rapports_CTC\"
```

### 3. Test (2 min)
```
1. F5 → Lancer l'application
2. Créer un rapport PDF
3. Vérifier que le dossier "Rapports_CTC" est créé à côté du fichier Excel
```

**✅ C'est fait! L'application fonctionne maintenant sur n'importe quelle machine.**

Pour une correction complète, continuez ci-dessous.

---

## 📋 PRÉPARATION

### Étape 0: Vérifications Préalables

**Avant de commencer**:
- [ ] Vous avez Microsoft Excel 2016 ou supérieur
- [ ] Les macros sont activées
- [ ] Vous avez les droits d'administrateur (pour créer des dossiers)
- [ ] Le fichier n'est pas ouvert en lecture seule

### Étape 1: Sauvegarde (OBLIGATOIRE)

```
1. Fermer Excel complètement
2. Copier "SMR_PILOT 2F.xlsm"
3. Renommer la copie: "SMR_PILOT 2F - BACKUP - 2025-11-15.xlsm"
4. Mettre le backup dans un dossier sûr (ex: OneDrive, USB)
5. Ouvrir "SMR_PILOT 2F.xlsm" (l'original)
```

### Étape 2: Ouvrir l'Éditeur VBA

```
1. Dans Excel: Alt+F11
2. Affichage > Explorateur de projet (Ctrl+R)
3. Affichage > Fenêtre Exécution (Ctrl+G)
```

Vous devriez voir:
```
┌─────────────────────────────┐
│ VBAProject (SMR_PILOT 2F)  │
│  ├─ Microsoft Excel Objets  │
│  │  ├─ ThisWorkbook         │
│  │  └─ Feuil1, Feuil2...    │
│  ├─ Modules                 │
│  │  ├─ Mod_Global           │
│  │  ├─ Mod_Session          │
│  │  └─ ...                  │
│  └─ UserForms               │
│     └─ UserForm_Connexion   │
└─────────────────────────────┘
```

---

## 🔧 CORRECTIONS (Par ordre de priorité)

### CORRECTION #4 - Chemin Absolu (PRIORITÉ MAXIMALE)

**Temps**: 5 minutes
**Impact**: L'application ne fonctionne que sur l'ordinateur de pgaye

#### Instructions:

1. **Localiser** le module `Mod_Session`
2. **Chercher** la fonction `CreerRapportPDF` (Ctrl+F → "CreerRapportPDF")
3. **Trouver** la ligne contenant `"C:\Users\pgaye\"`

**Code actuel**:
```vba
Dim cheminBase As String
cheminBase = "C:\Users\pgaye\OneDrive - Keolis\Bureau\personnel PKG\GOSMR\SMR Pilot\SMRCC\Rapports_CTC\"

On Error Resume Next
If Dir(cheminBase, vbDirectory) = "" Then MkDir cheminBase
On Error GoTo 0
```

**REMPLACER PAR** (version simple):
```vba
Dim cheminBase As String
cheminBase = ThisWorkbook.Path & "\Rapports_CTC\"

On Error Resume Next
If Dir(cheminBase, vbDirectory) = "" Then
    MkDir cheminBase
    If Err.Number <> 0 Then
        MsgBox "Impossible de créer le dossier: " & cheminBase & vbCrLf & Err.Description, vbCritical
        Exit Function
    End If
End If
On Error GoTo 0
```

4. **Compiler**: Debug > Compile VBAProject
5. **Sauvegarder**: Ctrl+S

✅ **Vérifié**: Le chemin est maintenant relatif au fichier Excel

---

### CORRECTION #1 - Conversion CLng (PRIORITÉ HAUTE)

**Temps**: 10 minutes
**Impact**: Crash si la dernière manœuvre a un ID invalide

#### Instructions:

1. **Localiser** le module `UserForm_CreerItineraire`
2. **Chercher** la fonction `cmdPreparer_Click`
3. **Trouver** la section "Génération de l'ID"

**Code actuel**:
```vba
Dim newID As String
Dim idNum As Long
If prochaineLigne = 2 Then
    idNum = 1
Else
    idNum = CLng(Mid(wsPOV.Cells(prochaineLigne - 1, "Q").Value, 4)) + 1
End If
newID = "MAN" & Format(idNum, "00")
```

**REMPLACER PAR**:
```vba
Dim newID As String
Dim idNum As Long
Dim tempVal As String

If prochaineLigne = 2 Then
    idNum = 1
Else
    tempVal = Trim(Mid(wsPOV.Cells(prochaineLigne - 1, "Q").Value, 4))
    If IsNumeric(tempVal) And tempVal <> "" Then
        idNum = CLng(tempVal) + 1
    Else
        MsgBox "ID précédent invalide. Réinitialisation à 1.", vbExclamation
        idNum = 1
    End If
End If
newID = "MAN" & Format(idNum, "00")
```

4. **Compiler**: Debug > Compile VBAProject
5. **Sauvegarder**: Ctrl+S

✅ **Vérifié**: Plus de crash si l'ID est corrompu

---

### CORRECTION #2 - Boucle While (PRIORITÉ HAUTE)

**Temps**: 8 minutes
**Impact**: Boucle infinie possible lors de la recherche de ligne libre

#### Instructions:

1. **Dans la même fonction** `cmdPreparer_Click`
2. **Trouver** la section "Trouver la prochaine ligne"

**Code actuel**:
```vba
Dim prochaineLigne As Long
prochaineLigne = 2

While wsPOV.Cells(prochaineLigne, "Q").Value <> ""
    prochaineLigne = prochaineLigne + 1
    If prochaineLigne = 12 Or prochaineLigne = 28 Then
        MsgBox "ERREUR: Le tableau des manœuvres semble plein.", vbCritical
        Exit Sub
    End If
Wend
```

**REMPLACER PAR**:
```vba
Dim prochaineLigne As Long
Dim compteurSecurite As Long

prochaineLigne = 2
compteurSecurite = 0

While wsPOV.Cells(prochaineLigne, "Q").Value <> "" And compteurSecurite < 1000
    prochaineLigne = prochaineLigne + 1
    compteurSecurite = compteurSecurite + 1

    If prochaineLigne = 12 Or prochaineLigne = 28 Or compteurSecurite >= 1000 Then
        MsgBox "ERREUR: Le tableau des manœuvres semble plein.", vbCritical
        Exit Sub
    End If
Wend
```

3. **Compiler**: Debug > Compile VBAProject
4. **Sauvegarder**: Ctrl+S

✅ **Vérifié**: La boucle s'arrêtera après 1000 itérations max

---

### CORRECTION #3 - On Error Resume Next (PRIORITÉ MOYENNE)

**Temps**: 15 minutes
**Impact**: Les erreurs sont masquées silencieusement

**Cette correction nécessite de chercher TOUS les `On Error Resume Next` dans le code.**

#### Méthode Rapide:

1. **Édition** > **Rechercher** (Ctrl+F)
2. **Rechercher**: `On Error Resume Next`
3. **Dans**: Projet en cours
4. **Cocher**: Rechercher dans tout le projet

Pour chaque occurrence trouvée:

**Si le code ressemble à**:
```vba
On Error Resume Next
ws.Shapes(nomElement).Fill.ForeColor.RGB = RGB(255, 0, 0)
' PAS DE "On Error GoTo 0" ici <<<< PROBLÈME
```

**AJOUTER après le bloc**:
```vba
On Error Resume Next
ws.Shapes(nomElement).Fill.ForeColor.RGB = RGB(255, 0, 0)
On Error GoTo 0  ' <<<< AJOUTÉ
```

**Locations typiques**:
- Module `Manoeuvres` (lignes ~731, 802, 839)
- Module `GestionAffichage` (partout où il y a des Shapes)
- Module `Travaux` (gestion visuelle)

**Compteur**: Vous devriez en trouver environ 10-15

✅ **Vérifié**: Chaque `On Error Resume Next` a son `On Error GoTo 0`

---

### CORRECTION #6 - Variable Redéclarée (PRIORITÉ BASSE)

**Temps**: 3 minutes
**Impact**: Confusion entre variable locale et publique

#### Instructions:

1. **Localiser** le module `Travaux`
2. **Chercher** `Dim agentValidateur As String` dans la fonction `TerminerTravaux_Selectionnes`

**REMPLACER**:
```vba
Dim agentValidateur As String
agentValidateur = wsPOV.Cells(ligne, "W").Value
```

**PAR**:
```vba
Dim agentValidateurLocal As String
agentValidateurLocal = Trim(wsPOV.Cells(ligne, "W").Value)
```

3. **Remplacer** toutes les utilisations de `agentValidateur` par `agentValidateurLocal` dans cette fonction
4. **Compiler**: Debug > Compile VBAProject
5. **Sauvegarder**: Ctrl+S

✅ **Vérifié**: Plus de conflit de variable

---

### CORRECTIONS #5, #7, #8 - Améliorations Avancées (OPTIONNEL)

**Temps**: 45 minutes
**Impact**: Robustesse améliorée

Ces corrections nécessitent d'ajouter des fonctions helpers et de modifier de nombreux endroits.

**Voir le document** `CODE_VBA_CORRIGE.md` sections 3, 6 et 7 pour les détails complets.

**Résumé**:
- **#5**: Créer des fonctions sécurisées pour accéder aux Shapes
- **#7**: Valider toutes les chaînes avant `Split()`
- **#8**: Vérifier l'existence des feuilles au démarrage

**Recommandation**: Appliquez ces corrections après avoir testé les 4 premières.

---

## ✅ TESTS

### Test 1: Compilation

```
1. Dans l'éditeur VBA: Debug > Compile VBAProject
2. Résultat attendu: Aucune erreur
3. Si erreur: Noter la ligne, corriger la syntaxe
```

### Test 2: Démarrage

```
1. Fermer l'éditeur VBA (Alt+Q)
2. Fermer Excel
3. Réouvrir SMR_PILOT 2F.xlsm
4. Activer les macros si demandé
5. Se connecter normalement
6. Résultat attendu: Aucun message d'erreur
```

### Test 3: Création Manœuvre

```
1. Menu > Créer Manœuvre
2. Sélectionner une rame
3. Sélectionner une destination
4. Cliquer "Préparer"
5. Résultat attendu: Manœuvre créée avec ID correct
```

### Test 4: Rapport PDF

```
1. Menu > Créer Rapport
2. Résultat attendu:
   - Dossier "Rapports_CTC" créé à côté du fichier Excel
   - PDF créé dans ce dossier
   - Nom: Rapport_CTC_YYYY-MM-DD_HH-MM-SS.pdf
```

**Si tous les tests passent**: ✅ Migration réussie!

---

## 🐛 DÉPANNAGE

### Erreur de Compilation

**Symptôme**: "Compile error: Variable not defined"
**Solution**: Vérifier que vous avez bien copié tout le code, y compris les `Dim`

**Symptôme**: "Compile error: Sub or Function not defined"
**Solution**: Vous avez probablement oublié de créer une fonction helper (ex: `ObtenirFeuilleSecurisee`)

### Erreur d'Exécution

**Symptôme**: "Run-time error '9': Subscript out of range"
**Solution**: Une feuille est manquante ou mal nommée. Vérifier les noms exacts.

**Symptôme**: "Run-time error '76': Path not found"
**Solution**: Le dossier `Rapports_CTC` ne peut pas être créé. Vérifier les permissions du dossier parent.

### Restauration

**Si tout va mal**:
```
1. Fermer Excel SANS SAUVEGARDER (Ctrl+Q → Non)
2. Ouvrir le backup: SMR_PILOT 2F - BACKUP.xlsm
3. Sauvegarder sous: SMR_PILOT 2F.xlsm (écraser)
4. Recommencer avec plus de prudence
```

---

## 📊 CHECKLIST FINALE

Après migration complète:

- [ ] Bug #1 corrigé: Conversion CLng protégée
- [ ] Bug #2 corrigé: Boucle While sécurisée
- [ ] Bug #3 corrigé: Tous les `On Error Resume Next` restaurés
- [ ] Bug #4 corrigé: Chemin relatif au lieu de `C:\Users\pgaye\`
- [ ] Bug #6 corrigé: Variable renommée
- [ ] Le code compile sans erreur
- [ ] L'application démarre normalement
- [ ] Une manœuvre peut être créée
- [ ] Un PDF peut être généré
- [ ] Le dossier `Rapports_CTC` est créé automatiquement

**Si tous les points sont cochés**: 🎉 **Migration terminée avec succès!**

---

## 📈 AMÉLIORATIONS FUTURES

Une fois ces corrections appliquées, vous pouvez envisager:

1. **Court terme**:
   - Ajouter un système de logging dans un fichier texte
   - Créer une feuille "Configuration" pour les chemins personnalisables

2. **Moyen terme**:
   - Implémenter les corrections #5, #7, #8 (fonctions helpers)
   - Ajouter des tests unitaires avec Rubberduck VBA

3. **Long terme**:
   - Migrer vers une base de données Access ou SQLite
   - Créer une application moderne (Web ou Desktop)

---

## 📞 RESSOURCES

**Fichiers de référence**:
- `RAPPORT_BUGS_VBA.md` - Liste complète des 37 bugs
- `CODE_VBA_CORRIGE.md` - Code source corrigé complet
- `ANALYSE_CODE.md` - Architecture de l'application

**En cas de problème**:
1. Consulter la section Dépannage ci-dessus
2. Vérifier que le backup existe
3. Tester chaque correction individuellement

---

**Version**: 1.0
**Dernière mise à jour**: 2025-11-15
**Temps total**: ~1h30 (4 corrections critiques) ou ~2h30 (toutes corrections)
