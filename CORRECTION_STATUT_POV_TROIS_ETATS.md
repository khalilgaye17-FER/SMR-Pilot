# 🚆 CORRECTION STATUT POV - SYSTÈME À TROIS ÉTATS

**Problème**: Le tableau POV n'affiche que "Piégée" ou "Disponible", sans nuance.

**Solution**: Implémenter un système à 3 états :
- ✅ **Disponible** - Toutes les routes sont libres
- ⚠️ **Partiel** - Au moins une route est bloquée, mais pas toutes
- ❌ **Indisponible** - Toutes les routes sont bloquées

---

## 📊 ANALYSE DU PROBLÈME ACTUEL

### Code source actuel (MettreAJourDashboard) :

```vba
' Cas 2 : La rame est piégée (aucune sortie)
ElseIf Not MoteurItineraires.ExisteUneSortieValide(posRame) Then
    icone = "[!]"
    statut = "Piégée (Sorties bloquées)"

' Cas 3 : Tout va bien
Else
    icone = "[OK]"
    statut = "Disponible"
End If
```

**Problème** : `ExisteUneSortieValide` retourne `True` dès qu'UNE seule sortie est disponible, ignorant les autres sorties bloquées.

---

## ✅ SOLUTION COMPLÈTE

### ÉTAPE 1 : Nouvelle fonction d'analyse (MoteurItineraires)

**Ouvrir** : Alt+F11 → Module `MoteurItineraires` (ou `GestionSecurite`)

**AJOUTER** cette nouvelle fonction :

```vba
' ============================================================
' ANALYSE DÉTAILLÉE DE LA DISPONIBILITÉ DES ROUTES
' ============================================================
Public Function AnalyserDisponibiliteRoutes(origine As String) As String
    ' RETOURNE:
    ' "DISPONIBLE" - Toutes les routes sont libres
    ' "PARTIEL" - Au moins une route bloquée, mais pas toutes
    ' "INDISPONIBLE" - Toutes les routes sont bloquées
    ' "AUCUNE_ROUTE" - Pas d'itinéraire défini pour cette origine

    Dim wsItin As Worksheet
    On Error Resume Next
    Set wsItin = ThisWorkbook.Worksheets("Itineraires")
    On Error GoTo 0

    If wsItin Is Nothing Then
        AnalyserDisponibiliteRoutes = "ERREUR"
        Exit Function
    End If

    Dim i As Long, derniereLigne As Long
    Dim origineItin As String, cheminItin As String
    Dim totalRoutes As Long
    Dim routesDisponibles As Long
    Dim routesBloquees As Long

    totalRoutes = 0
    routesDisponibles = 0
    routesBloquees = 0

    derniereLigne = wsItin.Cells(wsItin.Rows.Count, "A").End(xlUp).Row

    ' Parcourir tous les itinéraires possibles depuis cette origine
    For i = 2 To derniereLigne
        origineItin = Trim(wsItin.Cells(i, 1).Value)

        ' Vérifier si cet itinéraire part de notre position
        If InStr(1, origine, origineItin, vbTextCompare) = 1 Then
            totalRoutes = totalRoutes + 1
            cheminItin = Trim(wsItin.Cells(i, 3).Value)

            ' Vérifier si cette route est disponible
            If GestionSecurite.EstCheminDisponible(cheminItin) Then
                routesDisponibles = routesDisponibles + 1
            Else
                routesBloquees = routesBloquees + 1
            End If
        End If
    Next i

    ' Déterminer le statut final
    If totalRoutes = 0 Then
        AnalyserDisponibiliteRoutes = "AUCUNE_ROUTE"
    ElseIf routesDisponibles = totalRoutes Then
        AnalyserDisponibiliteRoutes = "DISPONIBLE"
    ElseIf routesDisponibles = 0 Then
        AnalyserDisponibiliteRoutes = "INDISPONIBLE"
    Else
        AnalyserDisponibiliteRoutes = "PARTIEL"
    End If
End Function

' ============================================================
' DÉTAIL DES ROUTES POUR DIAGNOSTIC
' ============================================================
Public Function ObtenirDetailRoutes(origine As String) As String
    ' Retourne un rapport détaillé de chaque route possible

    Dim wsItin As Worksheet
    On Error Resume Next
    Set wsItin = ThisWorkbook.Worksheets("Itineraires")
    On Error GoTo 0

    If wsItin Is Nothing Then
        ObtenirDetailRoutes = "Erreur: Feuille Itineraires introuvable"
        Exit Function
    End If

    Dim i As Long, derniereLigne As Long
    Dim origineItin As String, transit As String, cheminItin As String
    Dim rapport As String
    Dim raisonBlocage As String
    Dim totalRoutes As Long, routesOK As Long

    rapport = ""
    totalRoutes = 0
    routesOK = 0

    derniereLigne = wsItin.Cells(wsItin.Rows.Count, "A").End(xlUp).Row

    For i = 2 To derniereLigne
        origineItin = Trim(wsItin.Cells(i, 1).Value)

        If InStr(1, origine, origineItin, vbTextCompare) = 1 Then
            totalRoutes = totalRoutes + 1
            transit = Trim(wsItin.Cells(i, 2).Value)
            cheminItin = Trim(wsItin.Cells(i, 3).Value)

            If GestionSecurite.EstCheminDisponible(cheminItin) Then
                rapport = rapport & "✅ Vers " & transit & " : LIBRE" & vbCrLf
                routesOK = routesOK + 1
            Else
                raisonBlocage = GestionSecurite.ObtenirRaisonBlocage(cheminItin)
                rapport = rapport & "❌ Vers " & transit & " : " & raisonBlocage & vbCrLf
            End If
        End If
    Next i

    If totalRoutes = 0 Then
        rapport = "Aucun itinéraire défini pour cette position."
    Else
        rapport = "ROUTES DISPONIBLES: " & routesOK & "/" & totalRoutes & vbCrLf & vbCrLf & rapport
    End If

    ObtenirDetailRoutes = rapport
End Function
```

---

### ÉTAPE 2 : Modifier MettreAJourDashboard()

**Ouvrir** : Module `Mod_Dashboard`

**TROUVER** ce bloc de code :

```vba
            ' --- DIAGNOSTIC DE LA RAME ---
            If estBlocageTotal Then
                 icone = "[STOP"
                 statut = "Site Fermé"

            ' Cas 1 : La rame est sur une voie protégée
            ElseIf GestionDonnees.CompterProtections(posRame) > 0 Then
                icone = "[Px]"
                statut = "Consignée (" & GestionDonnees.LireProtections(posRame) & ")"

            ' Cas 2 : La rame est piégée (aucune sortie)
            ElseIf Not MoteurItineraires.ExisteUneSortieValide(posRame) Then
                icone = "[!]"
                statut = "Piégée (Sorties bloquées)"

            ' Cas 3 : Tout va bien
            Else
                icone = "[OK]"
                statut = "Disponible"
            End If
```

**REMPLACER PAR** :

```vba
            ' --- DIAGNOSTIC DE LA RAME (SYSTÈME À 3 ÉTATS) ---
            If estBlocageTotal Then
                 icone = "[STOP]"
                 statut = "Site Fermé"

            ' Cas 1 : La rame est sur une voie protégée
            ElseIf GestionDonnees.CompterProtections(posRame) > 0 Then
                icone = "[Px]"
                statut = "Consignée (" & GestionDonnees.LireProtections(posRame) & ")"

            ' Cas 2, 3, 4 : Analyse détaillée des routes disponibles
            Else
                Dim analyseBLocage As String
                analyseBLocage = MoteurItineraires.AnalyserDisponibiliteRoutes(posRame)

                Select Case analyseBLocage
                    Case "DISPONIBLE"
                        ' Toutes les routes sont libres
                        icone = "[OK]"
                        statut = "Disponible"

                    Case "PARTIEL"
                        ' Certaines routes bloquées, mais au moins une disponible
                        icone = "[~]"
                        statut = "Partiel (Sorties limitées)"

                    Case "INDISPONIBLE"
                        ' Toutes les routes sont bloquées
                        icone = "[!]"
                        statut = "Indisponible (Bloquée)"

                    Case "AUCUNE_ROUTE"
                        ' Pas d'itinéraire défini
                        icone = "[?]"
                        statut = "Non configuré"

                    Case Else
                        ' Erreur inattendue
                        icone = "[ERR]"
                        statut = "Erreur analyse"
                End Select
            End If
```

---

### ÉTAPE 3 : Améliorer le diagnostic interactif

**TROUVER** dans `DiagnostiquerRameSelectionnee` :

```vba
    ' 2. Si c'est un blocage "Piégée" ("[!]"), on lance le diagnostic détaillé
    If InStr(celluleRame, "[!]") > 0 Then
```

**REMPLACER PAR** (gérer aussi "[~]" et "[!]") :

```vba
    ' 2. Si c'est un blocage partiel ou total, on lance le diagnostic détaillé
    If InStr(celluleRame, "[!]") > 0 Or InStr(celluleRame, "[~]") > 0 Then
        Dim rapport As String
        rapport = MoteurItineraires.ObtenirDetailRoutes(posRame)

        Dim titreRapport As String
        If InStr(celluleRame, "[!]") > 0 Then
            titreRapport = "RAME INDISPONIBLE"
        Else
            titreRapport = "DISPONIBILITÉ PARTIELLE"
        End If

        MsgBox titreRapport & vbCrLf & vbCrLf & _
               "Rame : " & nomRame & " (sur " & posRame & ")" & vbCrLf & vbCrLf & _
               rapport, vbInformation, "Analyse des Routes"
```

---

## 🎨 FORMATAGE VISUEL DU TABLEAU POV

### ÉTAPE 4 (OPTIONNEL) : Couleurs par statut

Ajoutez ce code à la fin de `MettreAJourDashboard()` pour colorer automatiquement :

```vba
            ' --- FORMATAGE COULEUR SELON STATUT ---
            With wsPOV.Cells(ligneDash, "C")
                Select Case analyseBLocage
                    Case "DISPONIBLE"
                        .Font.Color = RGB(0, 128, 0) ' Vert
                        .Font.Bold = False
                    Case "PARTIEL"
                        .Font.Color = RGB(255, 165, 0) ' Orange
                        .Font.Bold = True
                    Case "INDISPONIBLE"
                        .Font.Color = RGB(192, 0, 0) ' Rouge
                        .Font.Bold = True
                    Case Else
                        .Font.Color = RGB(128, 128, 128) ' Gris
                        .Font.Bold = False
                End Select
            End With
```

**Note** : Insérez ce bloc JUSTE APRÈS l'assignation de `statut` et AVANT l'écriture dans les cellules.

---

## 📋 CODE COMPLET DE MettreAJourDashboard() CORRIGÉ

Voici la fonction complète avec toutes les modifications :

```vba
Public Sub MettreAJourDashboard()
    On Error Resume Next

    Application.ScreenUpdating = False

    Dim wsData As Worksheet, wsPOV As Worksheet
    Set wsData = ThisWorkbook.Worksheets("Data")
    Set wsPOV = ThisWorkbook.Worksheets(NOM_FEUILLE_POV)

    If wsData Is Nothing Or wsPOV Is Nothing Then
        MsgBox "Erreur: Feuilles Data ou POV introuvables.", vbCritical
        Exit Sub
    End If

    ' Nettoyer la zone du tableau (colonnes A, B, C à partir de la ligne du tableau)
    ' Adapter ces valeurs selon votre mise en page POV
    Dim ligneDashDebut As Long
    ligneDashDebut = 25 ' Ligne où commence le tableau d'état

    Dim ligneDash As Long
    ligneDash = ligneDashDebut

    ' Effacer les anciennes données
    wsPOV.Range("A" & ligneDashDebut & ":C" & (ligneDashDebut + 50)).ClearContents
    wsPOV.Range("A" & ligneDashDebut & ":C" & (ligneDashDebut + 50)).Font.Color = RGB(0, 0, 0)
    wsPOV.Range("A" & ligneDashDebut & ":C" & (ligneDashDebut + 50)).Font.Bold = False

    ' Vérifier s'il y a un blocage total du site
    Dim estBlocageTotal As Boolean
    estBlocageTotal = GestionDonnees.EstSiteFerme()

    ' Parcourir toutes les rames
    Dim i As Long, derniereLigneData As Long
    derniereLigneData = wsData.Cells(wsData.Rows.Count, "E").End(xlUp).Row

    Dim nomRame As String, posRame As String
    Dim icone As String, statut As String
    Dim analyseBlocage As String

    For i = 2 To derniereLigneData
        nomRame = Trim(wsData.Cells(i, "E").Value)
        posRame = Trim(wsData.Cells(i, "F").Value)

        ' On n'affiche que les rames présentes sur site (pas au Garage)
        If nomRame <> "" And posRame <> "Garage" And posRame <> "" Then

            ' --- DIAGNOSTIC DE LA RAME (SYSTÈME À 3 ÉTATS) ---
            If estBlocageTotal Then
                 icone = "[STOP]"
                 statut = "Site Fermé"
                 analyseBlocage = "STOP"

            ' Cas 1 : La rame est sur une voie protégée
            ElseIf GestionDonnees.CompterProtections(posRame) > 0 Then
                icone = "[Px]"
                statut = "Consignée (" & GestionDonnees.LireProtections(posRame) & ")"
                analyseBlocage = "CONSIGNEE"

            ' Cas 2, 3, 4 : Analyse détaillée des routes disponibles
            Else
                analyseBlocage = MoteurItineraires.AnalyserDisponibiliteRoutes(posRame)

                Select Case analyseBlocage
                    Case "DISPONIBLE"
                        icone = "[OK]"
                        statut = "Disponible"

                    Case "PARTIEL"
                        icone = "[~]"
                        statut = "Partiel (Sorties limitées)"

                    Case "INDISPONIBLE"
                        icone = "[!]"
                        statut = "Indisponible (Bloquée)"

                    Case "AUCUNE_ROUTE"
                        icone = "[?]"
                        statut = "Non configuré"

                    Case Else
                        icone = "[ERR]"
                        statut = "Erreur analyse"
                End Select
            End If

            ' --- ÉCRITURE DANS LE TABLEAU POV ---
            wsPOV.Cells(ligneDash, "A").Value = icone & " " & nomRame
            wsPOV.Cells(ligneDash, "B").Value = posRame
            wsPOV.Cells(ligneDash, "C").Value = statut

            ' --- FORMATAGE COULEUR SELON STATUT ---
            With wsPOV.Cells(ligneDash, "C")
                Select Case analyseBlocage
                    Case "DISPONIBLE"
                        .Font.Color = RGB(0, 128, 0) ' Vert
                        .Font.Bold = False
                    Case "PARTIEL"
                        .Font.Color = RGB(255, 165, 0) ' Orange
                        .Font.Bold = True
                    Case "INDISPONIBLE"
                        .Font.Color = RGB(192, 0, 0) ' Rouge
                        .Font.Bold = True
                    Case "CONSIGNEE"
                        .Font.Color = RGB(128, 0, 128) ' Violet
                        .Font.Bold = True
                    Case "STOP"
                        .Font.Color = RGB(192, 0, 0) ' Rouge
                        .Font.Bold = True
                    Case Else
                        .Font.Color = RGB(128, 128, 128) ' Gris
                        .Font.Bold = False
                End Select
            End With

            ligneDash = ligneDash + 1
        End If
    Next i

    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub
```

---

## 🧪 CAS DE TEST

### Scénario 1 : Rame complètement disponible
- **Situation** : Rame R8050 sur Voie 1, toutes les aiguilles libres
- **Attendu** : `[OK] R8050 | Voie 1 | Disponible` (vert)

### Scénario 2 : Rame partiellement bloquée
- **Situation** : Rame R8051 sur Voie 2, aiguille Ag1 en dérangement (bloque sortie Nord), sortie Sud libre
- **Attendu** : `[~] R8051 | Voie 2 | Partiel (Sorties limitées)` (orange)

### Scénario 3 : Rame totalement bloquée
- **Situation** : Rame R8052 sur Voie 3, Ag2 en dérangement ET Voie 4 occupée par R8053
- **Attendu** : `[!] R8052 | Voie 3 | Indisponible (Bloquée)` (rouge)

### Scénario 4 : Clic sur rame partielle
- **Action** : Cliquer sur la ligne "[~] R8051"
- **Attendu** : MsgBox affichant :
  ```
  DISPONIBILITÉ PARTIELLE

  Rame : R8051 (sur Voie 2)

  ROUTES DISPONIBLES: 1/2

  ❌ Vers Transit Nord : Bloqué par Ag1 (Dérangement)
  ✅ Vers Transit Sud : LIBRE
  ```

---

## 📋 CHECKLIST D'APPLICATION

- [ ] Ajouté `AnalyserDisponibiliteRoutes()` dans MoteurItineraires
- [ ] Ajouté `ObtenirDetailRoutes()` dans MoteurItineraires
- [ ] Remplacé la logique de statut dans `MettreAJourDashboard()`
- [ ] Mis à jour `DiagnostiquerRameSelectionnee()` pour gérer `[~]`
- [ ] Compilé sans erreur (Debug > Compile)
- [ ] Testé avec une rame disponible
- [ ] Testé avec une rame partiellement bloquée
- [ ] Testé avec une rame totalement bloquée
- [ ] Vérifié les couleurs dans le tableau POV

---

## ⏱️ TEMPS D'IMPLÉMENTATION

- Ajouter `AnalyserDisponibiliteRoutes()` : **3 min**
- Ajouter `ObtenirDetailRoutes()` : **3 min**
- Modifier `MettreAJourDashboard()` : **5 min**
- Compiler et corriger erreurs : **2 min**
- Tests fonctionnels : **10 min**

**Total estimé** : ~25 minutes

---

## ⚠️ POINTS D'ATTENTION

1. **Variable `analyseBlocage`** : Doit être déclarée au bon endroit (dans le For...Next)

2. **Module correct** :
   - `AnalyserDisponibiliteRoutes` va dans **MoteurItineraires** ou **GestionSecurite**
   - La modification de `MettreAJourDashboard` va dans **Mod_Dashboard**

3. **Appel de fonction** : Assurez-vous que le module appelé est bien Public

4. **Ligne du tableau** : Adaptez `ligneDashDebut = 25` selon votre feuille POV

---

## 🎯 RÉSUMÉ DES CHANGEMENTS

| Avant | Après |
|-------|-------|
| 2 états (Piégée/Disponible) | **3 états** (Disponible/Partiel/Indisponible) |
| Icône `[OK]` ou `[!]` | `[OK]`, `[~]`, `[!]` |
| Texte noir | **Couleurs** (Vert/Orange/Rouge) |
| Diagnostic basique | **Rapport détaillé** avec routes ✅/❌ |

---

Appliquez ces modifications et votre tableau POV affichera maintenant correctement les 3 états de disponibilité ! 🚆✅
