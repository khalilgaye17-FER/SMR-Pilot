# 🔧 AJOUT RAPPORT PROCÉDURES AU PDF GROUPÉ

**Objectif**: Inclure le rapport des procédures (celui avec l'aperçu formaté) dans le PDF groupé lors de la déconnexion.

---

## MODIFICATION DANS GenererPDFGroupe()

### Instructions

1. Ouvrir VBA (Alt+F11)
2. Module **Mod_Session**
3. Trouver la fonction **GenererPDFGroupe()**
4. **AJOUTER** cette section après "3.7 NOTES DE SERVICE" et AVANT "4. SUPPRIMER LES FEUILLES PAR DÉFAUT"

### Code à AJOUTER

Trouvez cette partie dans `GenererPDFGroupe()`:

```vba
    ' 3.7 NOTES DE SERVICE (SI DONNÉES)
    If HistoriqueContientDonnees("Log_NotesService") Then
        ' ... code existant ...
    End If

    ' === 4. SUPPRIMER LES FEUILLES PAR DÉFAUT ===
```

**AJOUTEZ CE CODE** juste avant la section 4:

```vba
    ' 3.8 HISTORIQUE PROCÉDURES (SI DONNÉES)
    If HistoriqueContientDonnees("Historique_Procedures") Then
        On Error Resume Next
        Set wsSource = ThisWorkbook.Worksheets("Historique_Procedures")
        If Not wsSource Is Nothing Then
            wsSource.Copy After:=wbTemp.Sheets(wbTemp.Sheets.Count)
            Set wsTemp = wbTemp.Sheets(wbTemp.Sheets.Count)
            wsTemp.Name = "Historique_Procedures"

            ' Formatage spécial pour le rapport procédures
            FormaterRapportProcedures wsTemp

            ConfigurerMiseEnPagePDF wsTemp, "Historique des Procédures", False ' Portrait
            nbFeuillesCopiees = nbFeuillesCopiees + 1
        End If
        On Error GoTo 0
    End If

    ' === 4. SUPPRIMER LES FEUILLES PAR DÉFAUT DU CLASSEUR TEMPORAIRE ===
```

---

## NOUVELLE FONCTION - FormaterRapportProcedures

### Instructions

Dans le même module **Mod_Session**, **AJOUTER** cette fonction (après `ConfigurerMiseEnPagePDF`):

```vba
' ============================================================
' FORMATAGE RAPPORT PROCÉDURES - Identique à l'aperçu
' ============================================================
Private Sub FormaterRapportProcedures(ws As Worksheet)
    ' RÔLE: Formate le rapport procédures pour qu'il soit identique à l'aperçu

    On Error Resume Next

    ' === 1. LARGEURS DE COLONNES OPTIMISÉES ===
    ws.Columns("A").ColumnWidth = 15 ' ID Exécution
    ws.Columns("B").ColumnWidth = 12 ' Date
    ws.Columns("C").ColumnWidth = 25 ' Procédure
    ws.Columns("D").ColumnWidth = 25 ' Dérangement lié
    ws.Columns("E").ColumnWidth = 20 ' Agent
    ws.Columns("F").ColumnWidth = 12 ' Statut
    ws.Columns("G").ColumnWidth = 15 ' Heure début
    ws.Columns("H").ColumnWidth = 15 ' Heure fin
    ws.Columns("I").ColumnWidth = 12 ' Durée
    ws.Columns("J").ColumnWidth = 30 ' Chemin pris
    ws.Columns("K").ColumnWidth = 35 ' Observations
    ws.Columns("L").ColumnWidth = 35 ' Problèmes rencontrés

    ' === 2. EN-TÊTE FORMATÉ (IDENTIQUE À L'APERÇU) ===
    With ws.Rows(1)
        .Font.Bold = True
        .Font.Size = 11
        .Font.Name = "Calibri"
        .Interior.Color = RGB(47, 85, 151) ' Bleu foncé professionnel
        .Font.Color = RGB(255, 255, 255) ' Texte blanc
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 35
        .WrapText = True
    End With

    ' === 3. ZONE DE DONNÉES ===
    Dim derniereLigne As Long
    derniereLigne = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    If derniereLigne > 1 Then
        ' Bordures complètes
        With ws.Range("A1:L" & derniereLigne)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = RGB(128, 128, 128)
        End With

        ' Bordure épaisse autour du tableau
        With ws.Range("A1:L" & derniereLigne).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
        With ws.Range("A1:L" & derniereLigne).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
        With ws.Range("A1:L" & derniereLigne).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
        With ws.Range("A1:L" & derniereLigne).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With

        ' Lignes alternées pour meilleure lisibilité
        Dim i As Long
        For i = 2 To derniereLigne
            If i Mod 2 = 0 Then
                ws.Range("A" & i & ":L" & i).Interior.Color = RGB(242, 247, 252) ' Bleu très clair
            Else
                ws.Range("A" & i & ":L" & i).Interior.Color = RGB(255, 255, 255) ' Blanc
            End If
        Next i

        ' Formatage des cellules de données
        With ws.Range("A2:L" & derniereLigne)
            .Font.Size = 10
            .Font.Name = "Calibri"
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With

        ' Alignement spécifique par colonne
        ws.Range("A2:A" & derniereLigne).HorizontalAlignment = xlCenter ' ID
        ws.Range("B2:B" & derniereLigne).HorizontalAlignment = xlCenter ' Date
        ws.Range("C2:C" & derniereLigne).HorizontalAlignment = xlLeft   ' Procédure
        ws.Range("D2:D" & derniereLigne).HorizontalAlignment = xlLeft   ' Dérangement
        ws.Range("E2:E" & derniereLigne).HorizontalAlignment = xlCenter ' Agent
        ws.Range("F2:F" & derniereLigne).HorizontalAlignment = xlCenter ' Statut
        ws.Range("G2:H" & derniereLigne).HorizontalAlignment = xlCenter ' Heures
        ws.Range("I2:I" & derniereLigne).HorizontalAlignment = xlCenter ' Durée
        ws.Range("J2:L" & derniereLigne).HorizontalAlignment = xlLeft   ' Textes longs

        ' Formatage conditionnel du statut
        For i = 2 To derniereLigne
            Select Case UCase(Trim(ws.Cells(i, "F").Value))
                Case "TERMINÉ", "TERMINE", "COMPLETED"
                    ws.Cells(i, "F").Font.Color = RGB(0, 128, 0) ' Vert
                    ws.Cells(i, "F").Font.Bold = True
                Case "ABANDONNÉ", "ABANDONNE", "ABANDONED"
                    ws.Cells(i, "F").Font.Color = RGB(192, 0, 0) ' Rouge
                    ws.Cells(i, "F").Font.Bold = True
                Case "EN COURS", "IN PROGRESS"
                    ws.Cells(i, "F").Font.Color = RGB(255, 165, 0) ' Orange
                    ws.Cells(i, "F").Font.Bold = True
            End Select
        Next i

        ' Ajuster automatiquement les hauteurs de ligne
        For i = 2 To derniereLigne
            ws.Rows(i).AutoFit
            ' Minimum 25 pixels de hauteur
            If ws.Rows(i).RowHeight < 25 Then
                ws.Rows(i).RowHeight = 25
            End If
        Next i
    End If

    ' === 4. TITRE DU RAPPORT (OPTIONNEL) ===
    ' Si vous voulez un titre au-dessus du tableau, décommentez:
    '
    ' ws.Rows(1).Insert
    ' With ws.Range("A1:L1")
    '     .Merge
    '     .Value = "RAPPORT DES PROCÉDURES EXÉCUTÉES"
    '     .Font.Bold = True
    '     .Font.Size = 16
    '     .Font.Color = RGB(47, 85, 151)
    '     .HorizontalAlignment = xlCenter
    '     .VerticalAlignment = xlCenter
    '     .RowHeight = 40
    ' End With

    On Error GoTo 0
End Sub
```

---

## ✅ RÉSULTAT FINAL

Après cette modification, le PDF groupé contiendra **8 sections** (au lieu de 7):

| Ordre | Section | Orientation | Condition |
|-------|---------|-------------|-----------|
| 1 | POV | 🗺️ Paysage | Toujours |
| 2 | Remise de Service | 📄 Portrait | Si données |
| 3 | Historique Manœuvres | 🗺️ Paysage | Si données |
| 4 | Historique Travaux | 🗺️ Paysage | Si données |
| 5 | Historique OL | 🗺️ Paysage | Si données |
| 6 | Rapport Dérangements | 📄 Portrait | Si données |
| 7 | Notes de Service | 📄 Portrait | Si données |
| **8** | **Historique Procédures** | **📄 Portrait** | **Si données** |

---

## 🎨 LE FORMATAGE APPLIQUÉ

Le rapport des procédures aura:

✅ **En-tête bleu foncé** avec texte blanc (comme l'aperçu)
✅ **Lignes alternées** bleu clair / blanc (meilleure lisibilité)
✅ **Bordures complètes** sur le tableau
✅ **Colonnes optimisées** pour chaque type de donnée
✅ **Statut coloré**:
   - ✅ Vert pour "Terminé"
   - ❌ Rouge pour "Abandonné"
   - ⏳ Orange pour "En cours"
✅ **Texte enveloppé** pour les descriptions longues
✅ **Hauteur de ligne automatique** (min 25px)

---

## 📋 CHECKLIST D'APPLICATION

- [ ] ✅ Ajouté la section "3.8 HISTORIQUE PROCÉDURES" dans `GenererPDFGroupe()`
- [ ] ✅ Ajouté la fonction `FormaterRapportProcedures()`
- [ ] ✅ Compilé sans erreur (Debug > Compile)
- [ ] ✅ Sauvegardé (Ctrl+S)
- [ ] ✅ Testé avec des procédures dans l'historique

---

## 🧪 TEST

### Pour tester:

1. **Exécuter une procédure** (si vous n'en avez pas déjà dans l'historique)
2. **Se déconnecter**
3. **Ouvrir le PDF groupé**
4. **Vérifier** que la section "Historique des Procédures" est présente
5. **Vérifier** que le formatage est identique à l'aperçu:
   - En-tête bleu foncé
   - Lignes alternées
   - Statuts colorés

---

## ⚠️ SI PROBLÈME

### Le rapport procédures n'apparaît pas

**Vérifier**:
1. La feuille "Historique_Procedures" existe
2. Il y a des données dedans (au moins une ligne après l'en-tête)

### Erreur de compilation

**Vérifier**:
1. La fonction `FormaterRapportProcedures` est bien déclarée comme `Private Sub`
2. Pas d'erreur de syntaxe (parenthèses, guillemets)

---

## 🎯 MODIFICATION OPTIONNELLE - Ajouter un titre

Si vous voulez un titre "RAPPORT DES PROCÉDURES EXÉCUTÉES" au-dessus du tableau dans le PDF:

1. Dans la fonction `FormaterRapportProcedures`
2. **Décommentez** la section "4. TITRE DU RAPPORT (OPTIONNEL)"

---

## ⏱️ TEMPS D'APPLICATION

- Ajouter la section 3.8: **2 min**
- Copier `FormaterRapportProcedures()`: **3 min**
- Compiler et tester: **5 min**

**Total**: ~10 minutes

---

## 📄 RÉSUMÉ DU PDF GROUPÉ FINAL

```
Rapport_Complet_Jean_Dupont_2025-11-15_14-30-00.pdf

📊 Page 1-3:   POV (Paysage)
📋 Page 4:     Remise de Service (Portrait)
🚆 Page 5-7:   Historique Manœuvres (Paysage)
⚠️ Page 8-9:   Rapport Dérangements (Portrait)
📦 Page 10-11: Historique OL (Paysage)
🔧 Page 12-13: Historique Procédures (Portrait) ← NOUVEAU!
📝 Page 14:    Notes de Service (Portrait)
```

---

**Appliquez cette modification et le rapport des procédures sera inclus dans votre PDF groupé avec un formatage professionnel!** 📄✨

Faites-moi savoir si ça fonctionne ou si vous voulez ajuster le formatage!
