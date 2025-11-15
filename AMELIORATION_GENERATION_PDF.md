# 🔧 AMÉLIORATION GÉNÉRATION PDF LORS DE LA DÉCONNEXION

**Objectifs**:
1. ✅ Générer les historiques **seulement s'ils contiennent des données**
2. ✅ Générer la Remise de Service en PDF (si applicable)
3. ✅ Générer le Rapport Dérangement en PDF (si applicable)
4. ✅ **Uniformiser le format** de tous les PDF
5. ✅ Mettre la feuille **POV en paysage**

---

## 📋 ORDRE DES MODIFICATIONS

| Ordre | Module/Fonction | Action | Temps |
|-------|-----------------|--------|-------|
| 1 | Mod_Session | Ajouter fonctions helpers | 15 min |
| 2 | Mod_Session | Modifier Deconnexion_Et_Archivage | 10 min |
| 3 | Mod_Session | Créer GenererTousLesPDF | 15 min |
| 4 | POV (Feuille) | Mettre en paysage | 2 min |

**Total**: ~40 minutes

---

## MODIFICATION 1 - FONCTIONS HELPERS (Mod_Session)

### Instructions

1. Ouvrir VBA (Alt+F11)
2. Trouver le module **Mod_Session**
3. **Aller à la fin** du module
4. **Coller** les fonctions suivantes **AVANT** le dernier `End Sub` ou `End Function`

### Code à AJOUTER à la fin du module Mod_Session

```vba
' ============================================================
' NOUVELLES FONCTIONS HELPERS - GÉNÉRATION PDF INTELLIGENTE
' Date: 2025-11-15
' ============================================================

' Vérifie si un historique contient des données
Private Function HistoriqueContientDonnees(nomFeuille As String) As Boolean
    HistoriqueContientDonnees = False

    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(nomFeuille)

    If ws Is Nothing Then
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    ' Vérifier s'il y a des données (au-delà de la ligne d'en-tête)
    ' On suppose que la ligne 1 contient les en-têtes
    Dim derniereLigne As Long
    derniereLigne = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' S'il y a plus d'une ligne (en-tête + données)
    If derniereLigne > 1 Then
        ' Vérifier qu'il y a vraiment du contenu, pas juste des lignes vides
        Dim i As Long
        For i = 2 To derniereLigne
            If Trim(ws.Cells(i, "A").Value) <> "" Then
                HistoriqueContientDonnees = True
                Exit Function
            End If
        Next i
    End If
End Function

' Génère un PDF d'une feuille avec format standardisé
Private Function GenererPDFFeuille(nomFeuille As String, typeFeuille As String, Optional enPaysage As Boolean = False) As String
    ' RÔLE: Génère un PDF standardisé pour n'importe quelle feuille
    ' RETOURNE: Chemin du fichier PDF créé, ou "" si échec

    GenererPDFFeuille = ""

    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(nomFeuille)

    If ws Is Nothing Then
        Debug.Print "GenererPDFFeuille: Feuille '" & nomFeuille & "' introuvable"
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    ' Construction du chemin (portable)
    Dim cheminBase As String
    Dim nomAgent As String

    nomAgent = Mod_Global.AgentConnecte
    nomAgent = Replace(nomAgent, " ", "_")

    cheminBase = ThisWorkbook.Path & Mod_Global.DOSSIER_RAPPORTS

    ' Créer le dossier si nécessaire
    On Error Resume Next
    If Dir(cheminBase, vbDirectory) = "" Then
        MkDir cheminBase
        If Err.Number <> 0 Then
            Debug.Print "Erreur création dossier: " & Err.Description
            Err.Clear
            On Error GoTo 0
            Exit Function
        End If
    End If
    On Error GoTo 0

    ' Nom du fichier avec timestamp
    Dim dateStr As String
    dateStr = Format(Now, "yyyy-mm-dd_HH-mm-ss")

    Dim nomFichier As String
    nomFichier = cheminBase & typeFeuille & "_" & nomAgent & "_" & dateStr & ".pdf"

    ' === MISE EN PAGE STANDARDISÉE ===
    On Error Resume Next
    With ws.PageSetup
        .PrintTitleRows = ws.Rows(1).Address ' En-tête sur chaque page

        If enPaysage Then
            .Orientation = xlLandscape
        Else
            .Orientation = xlPortrait
        End If

        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False ' Plusieurs pages en hauteur si nécessaire
        .PaperSize = xlPaperA4

        ' Marges standard (en points)
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)

        ' En-tête et pied de page
        .LeftHeader = "&B" & typeFeuille
        .CenterHeader = "&BAgent: " & Mod_Global.AgentConnecte
        .RightHeader = "&BDate: &D"
        .CenterFooter = "Page &P sur &N"

        ' Qualité d'impression
        .PrintQuality = 600
        .PrintErrors = xlPrintErrorsDisplayed
    End With

    If Err.Number <> 0 Then
        Debug.Print "Erreur mise en page: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    ' Export PDF
    On Error GoTo ErreurExport
    ws.ExportAsFixedFormat Type:=xlTypePDF, _
                            Filename:=nomFichier, _
                            Quality:=xlQualityStandard, _
                            IncludeDocProperties:=True, _
                            IgnorePrintAreas:=False, _
                            OpenAfterPublish:=False

    GenererPDFFeuille = nomFichier
    Debug.Print "PDF créé: " & nomFichier
    Exit Function

ErreurExport:
    Debug.Print "Erreur export PDF: " & Err.Description
    GenererPDFFeuille = ""
    Err.Clear
End Function

' Ouvre un fichier PDF (ou affiche un message si impossible)
Private Sub OuvrirPDFSecurise(cheminPDF As String)
    If cheminPDF = "" Or Dir(cheminPDF) = "" Then Exit Sub

    On Error Resume Next
    ' Méthode 1: Utiliser la fonction Windows ShellExecute
    ShellExecute 0, "open", cheminPDF, "", "", 1

    If Err.Number <> 0 Then
        ' Méthode 2: Utiliser Shell
        Err.Clear
        Shell "explorer.exe """ & cheminPDF & """", vbNormalFocus

        If Err.Number <> 0 Then
            ' Fallback: Juste afficher le chemin
            MsgBox "PDF créé:" & vbCrLf & cheminPDF, vbInformation
        End If
    End If
    On Error GoTo 0
End Sub
```

---

## MODIFICATION 2 - FONCTION PRINCIPALE DE GÉNÉRATION

### Code à AJOUTER dans Mod_Session

Ajoutez cette fonction après les helpers ci-dessus:

```vba
' ============================================================
' FONCTION PRINCIPALE: Génération intelligente de tous les PDF
' ============================================================
Public Function GenererTousLesPDF() As Boolean
    ' RÔLE: Génère tous les PDF nécessaires lors de la déconnexion
    ' RETOURNE: True si au moins un PDF a été créé

    GenererTousLesPDF = False

    Dim pdfsCrees As New Collection
    Dim cheminPDF As String
    Dim nbPDFCrees As Long
    nbPDFCrees = 0

    ' === 1. FEUILLE POV (TOUJOURS GÉNÉRER) ===
    Debug.Print "Génération PDF POV..."
    cheminPDF = GenererPDFFeuille(Mod_Global.NOM_FEUILLE_POV, "POV", True) ' PAYSAGE

    If cheminPDF <> "" Then
        pdfsCrees.Add cheminPDF
        nbPDFCrees = nbPDFCrees + 1
    End If

    ' === 2. REMISE DE SERVICE (SI CONTIENT DONNÉES) ===
    If HistoriqueContientDonnees("Remise_Service") Then
        Debug.Print "Génération PDF Remise de Service..."
        cheminPDF = GenererPDFFeuille("Remise_Service", "Remise_Service", False) ' PORTRAIT

        If cheminPDF <> "" Then
            pdfsCrees.Add cheminPDF
            nbPDFCrees = nbPDFCrees + 1
        End If
    Else
        Debug.Print "Remise_Service: Vide, PDF non généré"
    End If

    ' === 3. HISTORIQUE MANOEUVRES (SI CONTIENT DONNÉES) ===
    If HistoriqueContientDonnees("Historique_Manoeuvres") Then
        Debug.Print "Génération PDF Historique Manoeuvres..."
        cheminPDF = GenererPDFFeuille("Historique_Manoeuvres", "Historique_Manoeuvres", True) ' PAYSAGE

        If cheminPDF <> "" Then
            pdfsCrees.Add cheminPDF
            nbPDFCrees = nbPDFCrees + 1
        End If
    Else
        Debug.Print "Historique_Manoeuvres: Vide, PDF non généré"
    End If

    ' === 4. HISTORIQUE TRAVAUX (SI CONTIENT DONNÉES) ===
    If HistoriqueContientDonnees("Historique_Travaux") Then
        Debug.Print "Génération PDF Historique Travaux..."
        cheminPDF = GenererPDFFeuille("Historique_Travaux", "Historique_Travaux", True) ' PAYSAGE

        If cheminPDF <> "" Then
            pdfsCrees.Add cheminPDF
            nbPDFCrees = nbPDFCrees + 1
        End If
    Else
        Debug.Print "Historique_Travaux: Vide, PDF non généré"
    End If

    ' === 5. HISTORIQUE OL (SI CONTIENT DONNÉES) ===
    If HistoriqueContientDonnees("Historique_OL") Then
        Debug.Print "Génération PDF Historique OL..."
        cheminPDF = GenererPDFFeuille("Historique_OL", "Historique_OL", True) ' PAYSAGE

        If cheminPDF <> "" Then
            pdfsCrees.Add cheminPDF
            nbPDFCrees = nbPDFCrees + 1
        End If
    Else
        Debug.Print "Historique_OL: Vide, PDF non généré"
    End If

    ' === 6. LOG DÉRANGEMENTS (SI CONTIENT DONNÉES) ===
    If HistoriqueContientDonnees("Log_Derangements") Then
        Debug.Print "Génération PDF Log Dérangements..."
        cheminPDF = GenererPDFFeuille("Log_Derangements", "Rapport_Derangements", False) ' PORTRAIT

        If cheminPDF <> "" Then
            pdfsCrees.Add cheminPDF
            nbPDFCrees = nbPDFCrees + 1
        End If
    Else
        Debug.Print "Log_Derangements: Vide, PDF non généré"
    End If

    ' === 7. LOG NOTES DE SERVICE (SI CONTIENT DONNÉES) ===
    If HistoriqueContientDonnees("Log_NotesService") Then
        Debug.Print "Génération PDF Notes de Service..."
        cheminPDF = GenererPDFFeuille("Log_NotesService", "Notes_Service", False) ' PORTRAIT

        If cheminPDF <> "" Then
            pdfsCrees.Add cheminPDF
            nbPDFCrees = nbPDFCrees + 1
        End If
    Else
        Debug.Print "Log_NotesService: Vide, PDF non généré"
    End If

    ' === RÉCAPITULATIF ===
    If nbPDFCrees > 0 Then
        Dim message As String
        message = "✅ GÉNÉRATION TERMINÉE" & vbCrLf & vbCrLf & _
                  nbPDFCrees & " PDF créé(s):" & vbCrLf & vbCrLf

        ' Lister les PDF créés
        Dim i As Long
        For i = 1 To pdfsCrees.Count
            message = message & "• " & Dir(pdfsCrees(i)) & vbCrLf
        Next i

        message = message & vbCrLf & "Dossier: " & ThisWorkbook.Path & Mod_Global.DOSSIER_RAPPORTS

        MsgBox message, vbInformation, "PDF Générés"

        ' Ouvrir le dossier des rapports
        On Error Resume Next
        Shell "explorer.exe """ & ThisWorkbook.Path & Mod_Global.DOSSIER_RAPPORTS & """", vbNormalFocus
        On Error GoTo 0

        GenererTousLesPDF = True
    Else
        MsgBox "Aucun PDF généré (aucune donnée dans les historiques).", vbExclamation
        GenererTousLesPDF = False
    End If
End Function
```

---

## MODIFICATION 3 - MODIFIER LA FONCTION DE DÉCONNEXION

### Instructions

Dans le module **Mod_Session**, trouvez la fonction `Deconnexion_Et_Archivage`

### Trouver cette section (ligne ~2900):

```vba
' 4. Générer et ouvrir le PDF
If Not CreerRapportPDF() Then
    MsgBox "ERREUR CRITIQUE: Le rapport PDF n'a pas pu être créé." & vbCrLf & _
           "L'archivage est annulé par sécurité.", vbCritical
    Exit Sub
End If
```

### REMPLACER PAR:

```vba
' === CORRECTION: Génération intelligente de TOUS les PDF ===
If Not GenererTousLesPDF() Then
    Dim reponseErreur As VbMsgBoxResult
    reponseErreur = MsgBox("ATTENTION: Aucun PDF n'a pu être créé." & vbCrLf & vbCrLf & _
                           "Voulez-vous continuer l'archivage malgré tout ?", _
                           vbYesNo + vbExclamation, "Erreur PDF")

    If reponseErreur = vbNo Then
        Exit Sub
    End If
End If
```

---

## MODIFICATION 4 - METTRE POV EN PAYSAGE

### Instructions

1. Dans Excel (pas VBA), cliquez sur l'onglet **POV**
2. Menu **Mise en page** > **Orientation** > **Paysage**
3. Menu **Mise en page** > **Taille** > **A4**
4. Menu **Mise en page** > **Zone d'impression** > **Définir la zone d'impression**
   - Sélectionnez toute la zone utile (ex: A1:Z50)
5. **Ctrl+S** (sauvegarder)

**OU via VBA** (plus rapide):

```vba
' Exécuter une seule fois dans la fenêtre Exécution (Ctrl+G)
With ThisWorkbook.Worksheets("POV").PageSetup
    .Orientation = xlLandscape
    .PaperSize = xlPaperA4
    .Zoom = False
    .FitToPagesWide = 1
    .FitToPagesTall = 1
End With
ThisWorkbook.Save
```

---

## MODIFICATION 5 - SUPPRIMER L'ANCIENNE FONCTION CreerRapportPDF (OPTIONNEL)

Si vous voulez nettoyer, vous pouvez **commenter** l'ancienne fonction `CreerRapportPDF`:

```vba
' ============================================
' ANCIENNE FONCTION - REMPLACÉE PAR GenererTousLesPDF
' Conservée pour compatibilité si utilisée ailleurs
' ============================================
' Function CreerRapportPDF() As Boolean
'     ... (tout le code)
' End Function
```

---

## ✅ RÉSUMÉ DES AMÉLIORATIONS

### PDF Générés Intelligemment

| Document | Condition | Orientation | Format |
|----------|-----------|-------------|--------|
| **POV** | Toujours | 🗺️ Paysage | A4 |
| **Remise de Service** | Si données | 📄 Portrait | A4 |
| **Historique Manoeuvres** | Si données | 🗺️ Paysage | A4 |
| **Historique Travaux** | Si données | 🗺️ Paysage | A4 |
| **Historique OL** | Si données | 🗺️ Paysage | A4 |
| **Rapport Dérangements** | Si données | 📄 Portrait | A4 |
| **Notes de Service** | Si données | 📄 Portrait | A4 |

### Format Standardisé

✅ **En-tête sur chaque page**: Type de document | Agent | Date
✅ **Pied de page**: Numéro de page (Page 1 sur 3)
✅ **Marges**: 0.5" gauche/droite, 0.75" haut/bas
✅ **Qualité**: 600 dpi
✅ **Nom de fichier**: `Type_NomAgent_Date.pdf`

### Résultat Final

À la déconnexion:
1. ✅ L'agent valide avec son mot de passe
2. ✅ Le système génère **seulement les PDF nécessaires**
3. ✅ Un message récapitulatif liste les PDF créés
4. ✅ Le dossier des rapports s'ouvre automatiquement
5. ✅ L'archivage continue normalement

---

## 📋 CHECKLIST D'APPLICATION

- [ ] ✅ Ajouté les 3 fonctions helpers dans Mod_Session
- [ ] ✅ Ajouté la fonction GenererTousLesPDF dans Mod_Session
- [ ] ✅ Modifié Deconnexion_Et_Archivage pour appeler GenererTousLesPDF
- [ ] ✅ Mis POV en paysage
- [ ] ✅ Compilé sans erreur (Debug > Compile)
- [ ] ✅ Sauvegardé (Ctrl+S)
- [ ] ✅ Testé la déconnexion

---

## 🧪 TEST

### Scénario de test

1. **Remplir** quelques données dans:
   - Historique_Manoeuvres (ajouter une manœuvre)
   - Log_Derangements (ajouter un dérangement)
   - Laisser Historique_Travaux **vide**

2. **Déconnecter** (bouton de déconnexion)

3. **Vérifier** dans le dossier `Rapports_CTC`:
   - ✅ POV.pdf (généré)
   - ✅ Historique_Manoeuvres.pdf (généré car contient données)
   - ✅ Rapport_Derangements.pdf (généré car contient données)
   - ❌ Historique_Travaux.pdf (NON généré car vide)

4. **Ouvrir** les PDF:
   - ✅ POV en paysage
   - ✅ Tous au format A4
   - ✅ En-têtes et pieds de page présents

---

## 🐛 DÉPANNAGE

### Erreur: "Sub or Function not defined"

**Cause**: Les helpers n'ont pas été ajoutés

**Solution**: Vérifier que les 3 fonctions helpers sont bien dans Mod_Session:
- `HistoriqueContientDonnees`
- `GenererPDFFeuille`
- `OuvrirPDFSecurise`

### Erreur: "ShellExecute not defined"

**Cause**: Fonction Windows non déclarée

**Solution**: Ajouter en haut du module Mod_Session:

```vba
' Déclaration API Windows
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As LongPtr, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As LongPtr
```

### Aucun PDF n'est généré

**Cause**: Les historiques sont considérés vides

**Solution**: Vérifier dans la fenêtre Exécution (Ctrl+G) les messages Debug.Print

---

## 📊 TEMPS D'APPLICATION

- Copier-coller les helpers: **5 min**
- Copier-coller GenererTousLesPDF: **5 min**
- Modifier Deconnexion_Et_Archivage: **2 min**
- Mettre POV en paysage: **1 min**
- Compiler et tester: **5 min**

**Total**: ~20 minutes

---

**Appliquez ces modifications et testez! Les PDF seront générés intelligemment.** 🚀
