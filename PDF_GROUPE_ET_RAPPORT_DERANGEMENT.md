# 🔧 PDF GROUPÉ ET RAPPORT DÉRANGEMENT

**Objectifs**:
1. ✅ **Fusionner tous les PDF en UN SEUL fichier**
2. ✅ **Sauvegarder dans un dossier spécifique à l'agent**
3. ✅ **Rapport dérangement identique à l'aperçu**

---

## 📋 SOLUTION COMPLÈTE

### Architecture de la solution

```
Déconnexion
    ↓
Créer dossier agent (si n'existe pas)
    ↓
Créer classeur temporaire
    ↓
Copier toutes les feuilles avec données
    ↓
Exporter en UN SEUL PDF
    ↓
Sauvegarder dans dossier agent
    ↓
Archiver
```

---

## MODIFICATION 1 - NOUVELLE FONCTION PDF GROUPÉ

### Instructions

1. Ouvrir VBA (Alt+F11)
2. Module **Mod_Session**
3. **REMPLACER** la fonction `GenererTousLesPDF()` existante par cette nouvelle version

### Code COMPLET de la fonction (REMPLACER)

```vba
' ============================================================
' FONCTION PRINCIPALE: Génération PDF GROUPÉ
' Version 2.0 - Tous les documents en UN SEUL PDF
' ============================================================
Public Function GenererPDFGroupe() As String
    ' RÔLE: Génère UN SEUL PDF contenant tous les documents de la journée
    ' RETOURNE: Chemin du fichier PDF créé, ou "" si échec

    On Error GoTo ErreurGeneration
    GenererPDFGroupe = ""

    ' === 1. CRÉATION DU DOSSIER AGENT ===
    Dim cheminDossierAgent As String
    Dim nomAgent As String

    nomAgent = Mod_Global.AgentConnecte
    If Trim(nomAgent) = "" Then nomAgent = "Agent_Inconnu"
    nomAgent = Replace(nomAgent, " ", "_")
    nomAgent = Replace(nomAgent, "/", "_")
    nomAgent = Replace(nomAgent, "\", "_")

    ' Chemin: Rapports_CTC\NomAgent\
    cheminDossierAgent = ThisWorkbook.Path & Mod_Global.DOSSIER_RAPPORTS & nomAgent & "\"

    ' Créer le dossier agent s'il n'existe pas
    On Error Resume Next
    If Dir(ThisWorkbook.Path & Mod_Global.DOSSIER_RAPPORTS, vbDirectory) = "" Then
        MkDir ThisWorkbook.Path & Mod_Global.DOSSIER_RAPPORTS
    End If
    If Dir(cheminDossierAgent, vbDirectory) = "" Then
        MkDir cheminDossierAgent
    End If

    If Err.Number <> 0 Then
        MsgBox "Impossible de créer le dossier:" & vbCrLf & cheminDossierAgent & vbCrLf & Err.Description, vbCritical
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    ' === 2. CRÉER UN CLASSEUR TEMPORAIRE ===
    Dim wbTemp As Workbook
    Set wbTemp = Workbooks.Add

    Dim wsSource As Worksheet
    Dim wsTemp As Worksheet
    Dim nbFeuillesCopiees As Long
    nbFeuillesCopiees = 0

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' === 3. COPIER LES FEUILLES AVEC DONNÉES ===

    ' 3.1 POV (TOUJOURS)
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets(Mod_Global.NOM_FEUILLE_POV)
    If Not wsSource Is Nothing Then
        wsSource.Copy After:=wbTemp.Sheets(wbTemp.Sheets.Count)
        Set wsTemp = wbTemp.Sheets(wbTemp.Sheets.Count)
        wsTemp.Name = "POV"
        ConfigurerMiseEnPagePDF wsTemp, "POV - Point de Vue Opérationnel", True ' Paysage
        nbFeuillesCopiees = nbFeuillesCopiees + 1
    End If
    On Error GoTo 0

    ' 3.2 REMISE DE SERVICE (SI DONNÉES)
    If HistoriqueContientDonnees("Remise_Service") Then
        On Error Resume Next
        Set wsSource = ThisWorkbook.Worksheets("Remise_Service")
        If Not wsSource Is Nothing Then
            wsSource.Copy After:=wbTemp.Sheets(wbTemp.Sheets.Count)
            Set wsTemp = wbTemp.Sheets(wbTemp.Sheets.Count)
            wsTemp.Name = "Remise_Service"
            ConfigurerMiseEnPagePDF wsTemp, "Remise de Service", False ' Portrait
            nbFeuillesCopiees = nbFeuillesCopiees + 1
        End If
        On Error GoTo 0
    End If

    ' 3.3 HISTORIQUE MANOEUVRES (SI DONNÉES)
    If HistoriqueContientDonnees("Historique_Manoeuvres") Then
        On Error Resume Next
        Set wsSource = ThisWorkbook.Worksheets("Historique_Manoeuvres")
        If Not wsSource Is Nothing Then
            wsSource.Copy After:=wbTemp.Sheets(wbTemp.Sheets.Count)
            Set wsTemp = wbTemp.Sheets(wbTemp.Sheets.Count)
            wsTemp.Name = "Historique_Manoeuvres"
            ConfigurerMiseEnPagePDF wsTemp, "Historique des Manœuvres", True ' Paysage
            nbFeuillesCopiees = nbFeuillesCopiees + 1
        End If
        On Error GoTo 0
    End If

    ' 3.4 HISTORIQUE TRAVAUX (SI DONNÉES)
    If HistoriqueContientDonnees("Historique_Travaux") Then
        On Error Resume Next
        Set wsSource = ThisWorkbook.Worksheets("Historique_Travaux")
        If Not wsSource Is Nothing Then
            wsSource.Copy After:=wbTemp.Sheets(wbTemp.Sheets.Count)
            Set wsTemp = wbTemp.Sheets(wbTemp.Sheets.Count)
            wsTemp.Name = "Historique_Travaux"
            ConfigurerMiseEnPagePDF wsTemp, "Historique des Travaux", True ' Paysage
            nbFeuillesCopiees = nbFeuillesCopiees + 1
        End If
        On Error GoTo 0
    End If

    ' 3.5 HISTORIQUE OL (SI DONNÉES)
    If HistoriqueContientDonnees("Historique_OL") Then
        On Error Resume Next
        Set wsSource = ThisWorkbook.Worksheets("Historique_OL")
        If Not wsSource Is Nothing Then
            wsSource.Copy After:=wbTemp.Sheets(wbTemp.Sheets.Count)
            Set wsTemp = wbTemp.Sheets(wbTemp.Sheets.Count)
            wsTemp.Name = "Historique_OL"
            ConfigurerMiseEnPagePDF wsTemp, "Historique OL", True ' Paysage
            nbFeuillesCopiees = nbFeuillesCopiees + 1
        End If
        On Error GoTo 0
    End If

    ' 3.6 RAPPORT DÉRANGEMENTS (SI DONNÉES)
    If HistoriqueContientDonnees("Log_Derangements") Then
        On Error Resume Next
        Set wsSource = ThisWorkbook.Worksheets("Log_Derangements")
        If Not wsSource Is Nothing Then
            wsSource.Copy After:=wbTemp.Sheets(wbTemp.Sheets.Count)
            Set wsTemp = wbTemp.Sheets(wbTemp.Sheets.Count)
            wsTemp.Name = "Rapport_Derangements"
            ConfigurerMiseEnPagePDF wsTemp, "Rapport des Dérangements", False ' Portrait
            nbFeuillesCopiees = nbFeuillesCopiees + 1
        End If
        On Error GoTo 0
    End If

    ' 3.7 NOTES DE SERVICE (SI DONNÉES)
    If HistoriqueContientDonnees("Log_NotesService") Then
        On Error Resume Next
        Set wsSource = ThisWorkbook.Worksheets("Log_NotesService")
        If Not wsSource Is Nothing Then
            wsSource.Copy After:=wbTemp.Sheets(wbTemp.Sheets.Count)
            Set wsTemp = wbTemp.Sheets(wbTemp.Sheets.Count)
            wsTemp.Name = "Notes_Service"
            ConfigurerMiseEnPagePDF wsTemp, "Notes de Service", False ' Portrait
            nbFeuillesCopiees = nbFeuillesCopiees + 1
        End If
        On Error GoTo 0
    End If

    ' === 4. SUPPRIMER LES FEUILLES PAR DÉFAUT DU CLASSEUR TEMPORAIRE ===
    Dim ws As Worksheet
    For Each ws In wbTemp.Worksheets
        If ws.Name = "Feuil1" Or ws.Name = "Sheet1" Or _
           ws.Name = "Feuil2" Or ws.Name = "Sheet2" Or _
           ws.Name = "Feuil3" Or ws.Name = "Sheet3" Then
            On Error Resume Next
            ws.Delete
            On Error GoTo 0
        End If
    Next ws

    ' === 5. GÉNÉRER LE NOM DU FICHIER PDF ===
    Dim dateStr As String
    dateStr = Format(Now, "yyyy-mm-dd_HH-mm-ss")

    Dim nomFichierPDF As String
    nomFichierPDF = cheminDossierAgent & "Rapport_Complet_" & nomAgent & "_" & dateStr & ".pdf"

    ' === 6. EXPORTER EN UN SEUL PDF ===
    If nbFeuillesCopiees > 0 Then
        ' Exporter TOUT le classeur temporaire en UN SEUL PDF
        wbTemp.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=nomFichierPDF, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False

        GenererPDFGroupe = nomFichierPDF

        ' Message de succès
        MsgBox "✅ RAPPORT COMPLET GÉNÉRÉ" & vbCrLf & vbCrLf & _
               "Fichier: " & Dir(nomFichierPDF) & vbCrLf & _
               "Contient: " & nbFeuillesCopiees & " section(s)" & vbCrLf & vbCrLf & _
               "Dossier: " & cheminDossierAgent, _
               vbInformation, "PDF Créé"

        ' Ouvrir le dossier
        On Error Resume Next
        Shell "explorer.exe """ & cheminDossierAgent & """", vbNormalFocus
        On Error GoTo 0

    Else
        MsgBox "Aucune donnée à exporter dans le rapport.", vbExclamation
    End If

    ' === 7. NETTOYAGE ===
    wbTemp.Close SaveChanges:=False
    Set wbTemp = Nothing

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Exit Function

ErreurGeneration:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "ERREUR lors de la génération du PDF groupé:" & vbCrLf & _
           "Numéro: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbCritical

    If Not wbTemp Is Nothing Then
        wbTemp.Close SaveChanges:=False
    End If

    GenererPDFGroupe = ""
End Function

' ============================================================
' HELPER: Configuration mise en page standardisée
' ============================================================
Private Sub ConfigurerMiseEnPagePDF(ws As Worksheet, titre As String, enPaysage As Boolean)
    ' RÔLE: Configure la mise en page d'une feuille pour export PDF

    On Error Resume Next

    With ws.PageSetup
        ' Orientation
        If enPaysage Then
            .Orientation = xlLandscape
        Else
            .Orientation = xlPortrait
        End If

        ' Format et ajustement
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False ' Plusieurs pages en hauteur si nécessaire

        ' Marges (en pouces)
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)

        ' En-têtes et pieds de page
        .LeftHeader = "&B" & titre
        .CenterHeader = "&BAgent: " & Mod_Global.AgentConnecte
        .RightHeader = "&BDate: &D"
        .CenterFooter = "Page &P sur &N"

        ' Ligne d'en-tête répétée
        On Error Resume Next
        .PrintTitleRows = ws.Rows(1).Address
        On Error GoTo 0

        ' Qualité
        .PrintQuality = 600
        .PrintErrors = xlPrintErrorsDisplayed
    End With

    On Error GoTo 0
End Sub
```

---

## MODIFICATION 2 - MODIFIER LA DÉCONNEXION

### Instructions

Dans **Mod_Session**, fonction `Deconnexion_Et_Archivage`:

**TROUVER** (ligne ~2920):
```vba
If Not GenererTousLesPDF() Then
```

**REMPLACER PAR**:
```vba
' === GÉNÉRATION DU PDF GROUPÉ ===
Dim cheminPDF As String
cheminPDF = GenererPDFGroupe()

If cheminPDF = "" Then
    Dim reponseErreur As VbMsgBoxResult
    reponseErreur = MsgBox("ATTENTION: Le PDF n'a pas pu être créé." & vbCrLf & vbCrLf & _
                           "Voulez-vous continuer l'archivage malgré tout ?", _
                           vbYesNo + vbExclamation, "Erreur PDF")

    If reponseErreur = vbNo Then
        Exit Sub
    End If
End If
```

---

## MODIFICATION 3 - RAPPORT DÉRANGEMENT IDENTIQUE À L'APERÇU

### Instructions

Le rapport de dérangement utilise maintenant directement la feuille `Log_Derangements` telle qu'elle apparaît à l'écran.

**Si vous voulez personnaliser l'apparence**, modifiez directement la feuille `Log_Derangements`:

1. Cliquez sur l'onglet **Log_Derangements**
2. Ajustez:
   - Largeurs de colonnes
   - Formatage des cellules
   - Couleurs d'en-tête
   - Bordures
3. **Ctrl+S** (sauvegarder)

**La feuille sera exportée exactement comme vous la voyez!**

### Si vous voulez un rapport formaté différemment

Créez une fonction qui génère une version formatée:

```vba
' À ajouter dans Mod_Session
Private Sub FormaterRapportDerangements(ws As Worksheet)
    ' RÔLE: Formate le rapport dérangements pour l'export PDF

    On Error Resume Next

    ' Largeurs de colonnes
    ws.Columns("A").ColumnWidth = 15 ' ID
    ws.Columns("B").ColumnWidth = 25 ' Type
    ws.Columns("C").ColumnWidth = 20 ' Installation
    ws.Columns("D").ColumnWidth = 30 ' Description
    ws.Columns("E").ColumnWidth = 15 ' Agent
    ws.Columns("F").ColumnWidth = 18 ' Date/Heure
    ws.Columns("G").ColumnWidth = 12 ' Statut

    ' En-tête
    With ws.Rows(1)
        .Font.Bold = True
        .Font.Size = 11
        .Interior.Color = RGB(68, 114, 196) ' Bleu
        .Font.Color = RGB(255, 255, 255) ' Blanc
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With

    ' Bordures sur toute la zone de données
    Dim derniereLigne As Long
    derniereLigne = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    If derniereLigne > 1 Then
        With ws.Range("A1:G" & derniereLigne)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With

        ' Alterner les couleurs de lignes
        Dim i As Long
        For i = 2 To derniereLigne
            If i Mod 2 = 0 Then
                ws.Rows(i).Interior.Color = RGB(242, 242, 242) ' Gris très clair
            End If
        Next i
    End If

    ' Ajuster hauteur des lignes avec texte long
    ws.Rows.AutoFit

    On Error GoTo 0
End Sub
```

**Puis dans `GenererPDFGroupe`, section 3.6**:

```vba
' 3.6 RAPPORT DÉRANGEMENTS (SI DONNÉES)
If HistoriqueContientDonnees("Log_Derangements") Then
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets("Log_Derangements")
    If Not wsSource Is Nothing Then
        wsSource.Copy After:=wbTemp.Sheets(wbTemp.Sheets.Count)
        Set wsTemp = wbTemp.Sheets(wbTemp.Sheets.Count)
        wsTemp.Name = "Rapport_Derangements"

        ' === FORMATAGE PERSONNALISÉ ===
        FormaterRapportDerangements wsTemp

        ConfigurerMiseEnPagePDF wsTemp, "Rapport des Dérangements", False
        nbFeuillesCopiees = nbFeuillesCopiees + 1
    End If
    On Error GoTo 0
End If
```

---

## MODIFICATION 4 - SUPPRIMER L'ANCIENNE FONCTION (Optionnel)

Vous pouvez **commenter** ou **supprimer** l'ancienne fonction `GenererTousLesPDF()` car elle est remplacée par `GenererPDFGroupe()`.

---

## ✅ RÉSULTAT FINAL

### Structure des dossiers

```
SMR_Pilot/
└── Rapports_CTC/
    ├── Jean_Dupont/
    │   ├── Rapport_Complet_Jean_Dupont_2025-11-15_14-30-00.pdf
    │   ├── Rapport_Complet_Jean_Dupont_2025-11-15_20-15-00.pdf
    │   └── ...
    ├── Marie_Martin/
    │   ├── Rapport_Complet_Marie_Martin_2025-11-15_08-00-00.pdf
    │   └── ...
    └── ...
```

### Contenu du PDF groupé

Le PDF contient **dans l'ordre**:

1. **POV** (Paysage) - Toujours présent
2. **Remise de Service** (Portrait) - Si données
3. **Historique Manœuvres** (Paysage) - Si données
4. **Historique Travaux** (Paysage) - Si données
5. **Historique OL** (Paysage) - Si données
6. **Rapport Dérangements** (Portrait) - Si données
7. **Notes de Service** (Portrait) - Si données

**Chaque section** commence sur une nouvelle page avec:
- En-tête: Titre | Agent | Date
- Pied de page: Numéro de page

---

## 🎬 APERÇU DU RÉSULTAT

### Message à la déconnexion

```
✅ RAPPORT COMPLET GÉNÉRÉ

Fichier: Rapport_Complet_Jean_Dupont_2025-11-15_14-30-00.pdf
Contient: 5 section(s)

Dossier: C:\...\SMR_Pilot\Rapports_CTC\Jean_Dupont\

[Le dossier s'ouvre automatiquement]
```

### PDF généré

```
Rapport_Complet_Jean_Dupont_2025-11-15_14-30-00.pdf (15 pages)

Page 1-3:   POV (Paysage)
Page 4:     Remise de Service (Portrait)
Page 5-8:   Historique Manœuvres (Paysage)
Page 9-10:  Rapport Dérangements (Portrait)
Page 11-15: Historique OL (Paysage)
```

---

## 📋 CHECKLIST D'APPLICATION

- [ ] ✅ Ajouté la fonction `GenererPDFGroupe()` dans Mod_Session
- [ ] ✅ Ajouté la fonction helper `ConfigurerMiseEnPagePDF()` dans Mod_Session
- [ ] ✅ (Optionnel) Ajouté `FormaterRapportDerangements()` si formatage personnalisé
- [ ] ✅ Modifié `Deconnexion_Et_Archivage` pour appeler `GenererPDFGroupe()`
- [ ] ✅ Formaté la feuille `Log_Derangements` comme souhaité
- [ ] ✅ Compilé sans erreur (Debug > Compile)
- [ ] ✅ Sauvegardé (Ctrl+S)
- [ ] ✅ Testé la déconnexion

---

## 🧪 TEST

1. **Ajouter des données** dans plusieurs historiques
2. **Se déconnecter**
3. **Vérifier**:
   - ✅ Un dossier avec votre nom est créé dans `Rapports_CTC\`
   - ✅ UN SEUL fichier PDF est créé
   - ✅ Le PDF contient toutes les sections avec données
   - ✅ POV est en paysage
   - ✅ Dérangements est en portrait et identique à l'aperçu
   - ✅ Chaque section a en-tête et pied de page

---

## 🐛 DÉPANNAGE

### Erreur: "Subscript out of range"

**Cause**: Une feuille n'existe pas

**Solution**: Vérifier que toutes les feuilles existent (POV, Remise_Service, etc.)

### PDF vide ou incomplet

**Cause**: Les feuilles n'ont pas été copiées

**Solution**: Vérifier dans la fenêtre Exécution (Ctrl+G) les messages Debug.Print

### Rapport dérangement pas comme à l'écran

**Cause**: Le formatage n'est pas sauvegardé

**Solution**:
1. Formater manuellement la feuille Log_Derangements
2. Sauvegarder le classeur
3. OU utiliser la fonction `FormaterRapportDerangements()`

---

## ⏱️ TEMPS D'APPLICATION

- Copier `GenererPDFGroupe()`: **5 min**
- Copier `ConfigurerMiseEnPagePDF()`: **2 min**
- Modifier `Deconnexion_Et_Archivage`: **2 min**
- (Optionnel) Formater rapport dérangement: **5 min**
- Compiler et tester: **5 min**

**Total**: ~20 minutes

---

## 🎯 AVANTAGES DE CETTE SOLUTION

✅ **UN SEUL FICHIER** au lieu de 7 fichiers séparés
✅ **Organisation par agent** (un dossier par personne)
✅ **Historique complet** de toutes les sessions
✅ **Navigation facile** entre les sections du PDF
✅ **Formatage cohérent** sur tous les documents
✅ **Portable** (pas besoin de logiciel de fusion PDF)

---

**Appliquez ces modifications et vous aurez un PDF groupé professionnel!** 📄✨

Dites-moi si vous voulez des ajustements sur le formatage du rapport dérangement ou d'autres sections.
