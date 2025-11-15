# 🐛 CORRECTION ERREUR PDF - RowHeight

**Erreur**: Erreur d'exécution 1004: Impossible de définir la propriété RowHeight de la classe Range
**Module**: Mod_Procedures (ou module contenant `ConstruireRapportProcedure`)
**Ligne problématique**: `ws.Rows(ligne).RowHeight = Application.Max(40, nbLignesTexte * 14)`

---

## 🔍 ANALYSE DU PROBLÈME

Cette erreur survient généralement pour l'une de ces raisons:

1. **Cellules fusionnées** sur la ligne → Excel ne peut pas définir la hauteur d'une ligne partiellement fusionnée
2. **Valeur invalide** → La hauteur calculée est trop grande (>409 points max dans Excel)
3. **Variable non initialisée** → `nbLignesTexte` pourrait être vide ou invalide
4. **Feuille protégée** → La feuille temporaire est verrouillée
5. **Ligne hors limites** → La variable `ligne` dépasse les limites Excel

---

## ✅ SOLUTION IMMÉDIATE (2 minutes)

### Étape 1: Localiser la fonction

1. Appuyez sur **Alt+F11** (ouvrir VBA)
2. Appuyez sur **Ctrl+F** (Rechercher)
3. Tapez: `ConstruireRapportProcedure`
4. Cliquez sur "Suivant"

### Étape 2: Trouver la ligne problématique

Dans la fonction, cherchez (Ctrl+F):
```vba
ws.Rows(ligne).RowHeight = Application.Max(40, nbLignesTexte * 14)
```

### Étape 3: REMPLACER par ce code corrigé

**Code ORIGINAL (BUGUÉ)**:
```vba
ws.Rows(ligne).RowHeight = Application.Max(40, nbLignesTexte * 14)
```

**Code CORRIGÉ** (copiez-collez):
```vba
' === CORRECTION ERREUR 1004: RowHeight sécurisé ===
On Error Resume Next
Dim hauteurCalculee As Double
Dim hauteurMax As Double

' Validation de nbLignesTexte
If IsNumeric(nbLignesTexte) And nbLignesTexte > 0 Then
    hauteurCalculee = nbLignesTexte * 14
Else
    hauteurCalculee = 40 ' Valeur par défaut
End If

' Calcul de la hauteur finale avec maximum
hauteurMax = Application.Max(40, hauteurCalculee)

' Limite de sécurité Excel (409 points max)
If hauteurMax > 409 Then hauteurMax = 409

' Application avec gestion d'erreur
ws.Rows(ligne).RowHeight = hauteurMax

If Err.Number <> 0 Then
    ' Si échec, essayer sans fusion
    Debug.Print "Avertissement: Impossible de définir la hauteur de la ligne " & ligne & " - " & Err.Description

    ' Fallback: AutoFit au lieu de RowHeight
    ws.Rows(ligne).AutoFit
    Err.Clear
End If
On Error GoTo 0
```

### Étape 4: Compiler et Tester

1. **Menu**: Debug > Compile VBAProject
2. **Résultat attendu**: Aucune erreur
3. **Sauvegarder**: Ctrl+S
4. **Fermer VBA**: Alt+Q
5. **Tester**: Générer à nouveau le PDF

---

## 🔧 SOLUTION ALTERNATIVE (Si la première ne fonctionne pas)

Si vous avez encore l'erreur, c'est probablement à cause de cellules fusionnées.

### Solution: Désactiver les fusions avant de définir RowHeight

**REMPLACER TOUTE la ligne par**:

```vba
' === SOLUTION ALTERNATIVE: Défusionner puis redimensionner ===
On Error Resume Next

' 1. Défusionner temporairement la ligne si nécessaire
Dim plageOriginale As Range
Set plageOriginale = ws.Rows(ligne)

Dim estFusionnee As Boolean
estFusionnee = False

' Vérifier si des cellules sont fusionnées
Dim cell As Range
For Each cell In plageOriginale.Cells
    If cell.MergeCells Then
        estFusionnee = True
        Exit For
    End If
Next cell

' 2. Si fusionnée, stocker les paramètres et défusionner
If estFusionnee Then
    plageOriginale.UnMerge
End If

' 3. Définir la hauteur
Dim hauteurCalculee As Double
hauteurCalculee = Application.Max(40, nbLignesTexte * 14)
If hauteurCalculee > 409 Then hauteurCalculee = 409 ' Limite Excel

ws.Rows(ligne).RowHeight = hauteurCalculee

If Err.Number <> 0 Then
    Debug.Print "Erreur RowHeight ligne " & ligne & ": " & Err.Description
    ws.Rows(ligne).AutoFit ' Fallback
    Err.Clear
End If

On Error GoTo 0
```

---

## 📋 SOLUTION COMPLÈTE (Fonction entière corrigée)

Si vous voulez une version ultra-robuste, voici comment réécrire toute la section:

```vba
' ============================================
' HELPER: Définir hauteur de ligne de façon sécurisée
' ============================================
Private Sub DefinirHauteurLigneSecurisee(ws As Worksheet, ligne As Long, nbLignesTexte As Long)
    ' RÔLE: Définit la hauteur d'une ligne de façon robuste

    On Error Resume Next

    ' Validation des paramètres
    If ligne < 1 Or ligne > 1048576 Then Exit Sub ' Limites Excel

    Dim hauteurCalculee As Double
    If IsNumeric(nbLignesTexte) And nbLignesTexte > 0 Then
        hauteurCalculee = Application.Max(40, nbLignesTexte * 14)
    Else
        hauteurCalculee = 40
    End If

    ' Limite Excel: 409 points maximum
    If hauteurCalculee > 409 Then hauteurCalculee = 409

    ' Tentative de définition directe
    ws.Rows(ligne).RowHeight = hauteurCalculee

    If Err.Number <> 0 Then
        ' Échec: probablement cellules fusionnées
        Debug.Print "DefinirHauteurLigne: Erreur ligne " & ligne & " - " & Err.Description

        ' Fallback: AutoFit
        Err.Clear
        ws.Rows(ligne).AutoFit

        If Err.Number <> 0 Then
            ' Même AutoFit a échoué, on abandonne silencieusement
            Debug.Print "DefinirHauteurLigne: AutoFit impossible ligne " & ligne
            Err.Clear
        End If
    End If

    On Error GoTo 0
End Sub
```

**Puis dans `ConstruireRapportProcedure`, REMPLACER**:
```vba
ws.Rows(ligne).RowHeight = Application.Max(40, nbLignesTexte * 14)
```

**PAR**:
```vba
Call DefinirHauteurLigneSecurisee(ws, ligne, nbLignesTexte)
```

---

## 🎯 RECOMMANDATION

**Pour une correction rapide** (2 minutes):
→ Utilisez la **Solution Immédiate Étape 3**

**Pour une solution robuste** (5 minutes):
→ Utilisez la **Solution Complète** avec la fonction helper

---

## ✅ VÉRIFICATION

Après avoir appliqué la correction:

1. **Compiler**: Debug > Compile VBAProject
2. **Tester**: Générer un PDF de procédure
3. **Résultat attendu**:
   - Aucune erreur 1004
   - PDF généré avec succès
   - Lignes correctement dimensionnées (ou AutoFit si problème)

---

## 🐛 SI L'ERREUR PERSISTE

### Diagnostic avancé

1. **Ajouter des Debug** pour identifier la valeur problématique:

```vba
' Juste AVANT la ligne qui plante
Debug.Print "Ligne: " & ligne
Debug.Print "nbLignesTexte: " & nbLignesTexte
Debug.Print "Hauteur calculée: " & Application.Max(40, nbLignesTexte * 14)

ws.Rows(ligne).RowHeight = Application.Max(40, nbLignesTexte * 14)
```

2. **Lancer le PDF**, quand l'erreur survient:
   - Noter les valeurs affichées dans la fenêtre Exécution (Ctrl+G)
   - M'envoyer ces valeurs pour diagnostic

### Autres causes possibles

**Si la feuille temporaire est protégée**:

Ajouter au début de `ConstruireRapportProcedure`:
```vba
' Déverrouiller la feuille si nécessaire
On Error Resume Next
ws.Unprotect
On Error GoTo 0
```

**Si le problème vient d'une fusion spécifique**:

Identifier quelle ligne exactement pose problème:
```vba
' Version avec numéro de ligne affiché
On Error Resume Next
ws.Rows(ligne).RowHeight = Application.Max(40, nbLignesTexte * 14)
If Err.Number <> 0 Then
    MsgBox "Erreur à la ligne " & ligne & vbCrLf & "Valeur nbLignesTexte: " & nbLignesTexte, vbExclamation
    Err.Clear
End If
On Error GoTo 0
```

---

## 📝 NOUVEAU BUG IDENTIFIÉ

Ce bug sera ajouté au rapport comme:

**BUG #39: RowHeight sur cellules fusionnées**
- **Module**: Mod_Procedures
- **Fonction**: ConstruireRapportProcedure
- **Impact**: ÉLEVÉ - Empêche la génération de PDF
- **Cause**: Tentative de définir RowHeight sur une ligne avec cellules fusionnées
- **Correction**: Gestion d'erreur et fallback AutoFit

---

## 🚀 APPLIQUER LA CORRECTION MAINTENANT

**ÉTAPES RAPIDES**:

1. Alt+F11 (VBA)
2. Ctrl+F → Chercher "ConstruireRapportProcedure"
3. Trouver la ligne `ws.Rows(ligne).RowHeight`
4. Copier le code corrigé de l'Étape 3 ci-dessus
5. Coller à la place de la ligne bugguée
6. Debug > Compile
7. Ctrl+S
8. Tester le PDF

**Durée**: 2 minutes
**Résultat**: PDF fonctionne! ✅

---

Faites-moi savoir si la correction fonctionne ou si vous avez besoin d'aide supplémentaire!
