# 🔧 CORRECTION : Rame bloquée par sa propre position

**Problème** : V17B (et autres rames) affichent "Indisponible" alors qu'elles ne sont bloquées que par leur propre occupation.

**Cause** : `EstCheminDisponible()` vérifie TOUS les éléments du chemin, y compris la position actuelle de la rame.

---

## ✅ SOLUTION COMPLÈTE

### ÉTAPE 1 : Modifier la signature de `EstCheminDisponible()`

**Ouvrir** : Module `GestionSecurite`

**TROUVER** :
```vba
Public Function EstCheminDisponible(chemin As String) As Boolean
```

**REMPLACER PAR** :
```vba
Public Function EstCheminDisponible(chemin As String, Optional origine As String = "") As Boolean
```

---

### ÉTAPE 2 : Modifier la logique de vérification

**TROUVER** (lignes 20-26 environ) :
```vba
    For i = LBound(elements) To UBound(elements)
        nomElem = Trim(elements(i))

        ' 1. Vérifie l'occupation (doit être "Libre")
        ' NOTE : Si l'origine est incluse dans le chemin, il faudra peut-être l'ignorer ici.
        ' Adaptez si nécessaire : If nomElem <> OrigineEnCours Then ...
        If GestionDonnees.LireEtatElement(nomElem) <> "Libre" Then
            EstCheminDisponible = False
            Exit Function
        End If

        ' 2. Vérifie les protections (doit être 0)
        If GestionDonnees.CompterProtections(nomElem) > 0 Then
            EstCheminDisponible = False
            Exit Function
        End If
    Next i
```

**REMPLACER PAR** :
```vba
    For i = LBound(elements) To UBound(elements)
        nomElem = Trim(elements(i))

        ' IGNORER LA POSITION D'ORIGINE (la rame y est déjà)
        Dim estOrigine As Boolean
        estOrigine = False

        If origine <> "" Then
            ' Vérifier si cet élément correspond à l'origine
            ' (comparaison insensible à la casse et aux espaces)
            If InStr(1, UCase(Trim(origine)), UCase(nomElem), vbTextCompare) > 0 Or _
               InStr(1, UCase(nomElem), UCase(Trim(origine)), vbTextCompare) > 0 Then
                estOrigine = True
            End If
        End If

        ' 1. Vérifie l'occupation (doit être "Libre")
        ' SAUF si c'est l'origine (la rame peut être sur sa propre position)
        If Not estOrigine Then
            If GestionDonnees.LireEtatElement(nomElem) <> "Libre" Then
                EstCheminDisponible = False
                Exit Function
            End If
        End If

        ' 2. Vérifie TOUJOURS les protections (même pour l'origine)
        If GestionDonnees.CompterProtections(nomElem) > 0 Then
            EstCheminDisponible = False
            Exit Function
        End If
    Next i
```

---

### ÉTAPE 3 : Mettre à jour les appels dans `MoteurItineraires`

**Ouvrir** : Module `MoteurItineraires`

**TROUVER** dans `ExisteUneSortieValide` :
```vba
Public Function ExisteUneSortieValide(origine As String) As Boolean
    Dim wsItin As Worksheet
    Set wsItin = ThisWorkbook.Worksheets("Itineraires")

    Dim i As Long, derniereLigne As Long
    Dim origineItin As String, cheminItin As String

    derniereLigne = wsItin.Cells(wsItin.Rows.count, "A").End(xlUp).Row
    ExisteUneSortieValide = False

    For i = 2 To derniereLigne
        origineItin = Trim(wsItin.Cells(i, 1).Value)

        If InStr(1, origine, origineItin, vbTextCompare) = 1 Then
             cheminItin = Trim(wsItin.Cells(i, 3).Value)
             If GestionSecurite.EstCheminDisponible(cheminItin) Then  ' ← ICI
                ExisteUneSortieValide = True
                Exit Function
             End If
        End If
    Next i
End Function
```

**REMPLACER PAR** :
```vba
Public Function ExisteUneSortieValide(origine As String) As Boolean
    Dim wsItin As Worksheet
    Set wsItin = ThisWorkbook.Worksheets("Itineraires")

    Dim i As Long, derniereLigne As Long
    Dim origineItin As String, cheminItin As String

    derniereLigne = wsItin.Cells(wsItin.Rows.count, "A").End(xlUp).Row
    ExisteUneSortieValide = False

    For i = 2 To derniereLigne
        origineItin = Trim(wsItin.Cells(i, 1).Value)

        If InStr(1, origine, origineItin, vbTextCompare) = 1 Then
             cheminItin = Trim(wsItin.Cells(i, 3).Value)
             If GestionSecurite.EstCheminDisponible(cheminItin, origine) Then  ' ← PASSER ORIGINE
                ExisteUneSortieValide = True
                Exit Function
             End If
        End If
    Next i
End Function
```

---

### ÉTAPE 4 : Mettre à jour `AnalyserDisponibiliteRoutes`

**TROUVER** dans `AnalyserDisponibiliteRoutes` :
```vba
            ' Vérifier si cette route est disponible
            If GestionSecurite.EstCheminDisponible(cheminItin) Then  ' ← ICI
                routesDisponibles = routesDisponibles + 1
            Else
                routesBloquees = routesBloquees + 1
            End If
```

**REMPLACER PAR** :
```vba
            ' Vérifier si cette route est disponible
            If GestionSecurite.EstCheminDisponible(cheminItin, origine) Then  ' ← PASSER ORIGINE
                routesDisponibles = routesDisponibles + 1
            Else
                routesBloquees = routesBloquees + 1
            End If
```

---

### ÉTAPE 5 : Mettre à jour `ObtenirDetailRoutes`

**TROUVER** dans `ObtenirDetailRoutes` :
```vba
            If GestionSecurite.EstCheminDisponible(cheminItin) Then  ' ← ICI
                rapport = rapport & "✅ Vers " & transit & " : LIBRE" & vbCrLf
                routesOK = routesOK + 1
            Else
```

**REMPLACER PAR** :
```vba
            If GestionSecurite.EstCheminDisponible(cheminItin, origine) Then  ' ← PASSER ORIGINE
                rapport = rapport & "✅ Vers " & transit & " : LIBRE" & vbCrLf
                routesOK = routesOK + 1
            Else
```

---

### ÉTAPE 6 : Vérifier les autres appels (IMPORTANT)

Recherchez TOUS les autres endroits où `EstCheminDisponible` est appelé :

**Comment chercher** :
1. Ctrl+F dans VBA
2. Rechercher : `EstCheminDisponible(`
3. Pour chaque occurrence trouvée :
   - Si l'appel a accès à la position d'origine → Ajouter le paramètre
   - Sinon → Laisser sans paramètre (comportement par défaut)

**Exemple d'autres modules à vérifier** :
- `MoteurManoeuvres` - lors de la validation d'itinéraire
- `GestionSecurite` - autres fonctions de vérification
- Tout module appelant cette fonction

---

## 🧪 TEST DE VALIDATION

### Test 1 : V17B seule sur sa voie
```
Position : V17B
État : V17B = Occupée
Itinéraires depuis V17B : 2 routes libres

ATTENDU : [OK] Disponible (vert)
```

### Test 2 : V17B avec une aiguille bloquée
```
Position : V17B
État : V17B = Occupée, Ag5 = Dérangement
Itinéraires : 1 route bloquée par Ag5, 1 route libre

ATTENDU : [~] Partiel (orange)
```

### Test 3 : V17B totalement bloquée
```
Position : V17B
État : V17B = Occupée, Ag5 = Dérangement, V18 = Occupée par autre rame
Itinéraires : Toutes les routes bloquées

ATTENDU : [!] Indisponible (rouge)
```

---

## 📋 CHECKLIST

- [ ] Modifié signature `EstCheminDisponible(chemin, Optional origine)`
- [ ] Ajouté logique d'exclusion de l'origine dans la boucle
- [ ] Mis à jour `ExisteUneSortieValide()`
- [ ] Mis à jour `AnalyserDisponibiliteRoutes()`
- [ ] Mis à jour `ObtenirDetailRoutes()`
- [ ] Vérifié TOUS les autres appels avec Ctrl+F
- [ ] Compilé sans erreur (Debug > Compile)
- [ ] Testé V17B → Doit afficher "Disponible"
- [ ] Testé diagnostic interactif (clic sur rame)

---

## ⏱️ TEMPS D'APPLICATION

- Modification `EstCheminDisponible()` : **5 min**
- Mise à jour des 3 appels dans `MoteurItineraires` : **3 min**
- Recherche d'autres appels : **5 min**
- Compilation et tests : **5 min**

**Total** : ~18 minutes

---

## 🎯 RÉSUMÉ

**Avant** : La rame vérifie tous les éléments du chemin → se bloque elle-même
**Après** : La rame ignore sa propre position → se débloque correctement

Cette correction permet aux rames d'analyser correctement leurs routes de sortie sans être bloquées par leur propre présence !
