# 🔧 REMPLACEMENT DES MODULES VBA - ÉTAPE PAR ÉTAPE

**Fichier**: SMR_PILOT 2F.xlsm
**Date**: 2025-11-15
**Durée totale**: ~1h30

---

## ⚠️ AVANT DE COMMENCER - OBLIGATOIRE

### Étape 0: Sauvegarde (5 minutes)

**TRÈS IMPORTANT**: Ne JAMAIS modifier sans backup!

```
1. Fermer Excel complètement
2. Aller dans le dossier contenant "SMR_PILOT 2F.xlsm"
3. Clic droit sur le fichier > Copier
4. Clic droit > Coller
5. Renommer la copie: "SMR_PILOT 2F - BACKUP - 2025-11-15.xlsm"
6. Mettre ce backup dans un endroit sûr (autre dossier, USB, OneDrive)
7. Ouvrir "SMR_PILOT 2F.xlsm" (l'original)
```

✅ **Vérification**: Vous avez bien 2 fichiers maintenant?
- `SMR_PILOT 2F.xlsm` (celui qu'on va modifier)
- `SMR_PILOT 2F - BACKUP - 2025-11-15.xlsm` (sauvegarde)

---

## 📋 ORDRE DE REMPLACEMENT

Nous allons remplacer les modules dans cet ordre (du plus critique au moins critique):

| Ordre | Module | Bugs corrigés | Temps | Difficulté |
|-------|--------|---------------|-------|------------|
| 1 | **Mod_Global** | Ajout constantes | 10 min | Facile |
| 2 | **Mod_Session** | #4 Chemin absolu | 5 min | Facile |
| 3 | **UserForm_CreerItineraire** | #1 CLng, #2 While | 15 min | Moyen |
| 4 | **ThisWorkbook** | #8 Vérification feuilles | 8 min | Facile |
| 5 | **GestionAffichage** | #5 Accès Shapes | 12 min | Moyen |
| 6 | **Manoeuvres** | #3 On Error | 20 min | Moyen |
| 7 | **GestionSecurite** | #7 Split() | 15 min | Moyen |
| 8 | **Travaux** | #6 Variable | 5 min | Facile |

**Total**: ~1h30

---

## 🎯 COMMENT UTILISER CE GUIDE

Pour chaque module, vous allez:

1. ✅ **Ouvrir l'éditeur VBA** (Alt+F11)
2. ✅ **Localiser le module** dans l'explorateur de projet (à gauche)
3. ✅ **Sélectionner TOUT le code** (Ctrl+A)
4. ✅ **Copier le nouveau code** depuis ce document
5. ✅ **Coller** dans le module VBA (Ctrl+V)
6. ✅ **Compiler** pour vérifier (Debug > Compile VBAProject)
7. ✅ **Sauvegarder** (Ctrl+S)

---

## MODULE 1 - MOD_GLOBAL (NOUVEAU MODULE OU EXISTANT)

### Instructions

**Si le module existe déjà**:
1. Double-cliquer sur "Mod_Global" dans l'explorateur de projet
2. Sélectionner TOUT le code (Ctrl+A)
3. Copier le code ci-dessous et coller (Ctrl+V)

**Si le module n'existe PAS**:
1. Insertion > Module
2. Dans la fenêtre Propriétés (F4), changer le nom en "Mod_Global"
3. Copier-coller le code ci-dessous

### Code Complet du Module Mod_Global (CORRIGÉ)

```vba
' ===============================================
' Module : Mod_Global (VERSION CORRIGÉE)
' Date correction: 2025-11-15
' Bugs corrigés: Ajout constantes centralisées
' ===============================================
Option Explicit

' ============================================
' CONSTANTES POUR LES NOMS DE FEUILLES (NOUVEAU)
' ============================================
Public Const NOM_FEUILLE_POV As String = "POV"
Public Const NOM_FEUILLE_DATA As String = "Data"
Public Const NOM_FEUILLE_VOIES As String = "Voies"
Public Const NOM_FEUILLE_ITINERAIRES As String = "Itineraires"
Public Const NOM_FEUILLE_DEPENDANCES As String = "Dependances"
Public Const NOM_FEUILLE_HIST_MANOEUVRES As String = "Historique_Manoeuvres"
Public Const NOM_FEUILLE_HIST_TRAVAUX As String = "Historique_Travaux"
Public Const NOM_FEUILLE_HIST_OL As String = "Historique_OL"
Public Const NOM_FEUILLE_AGENTS As String = "Agents"

' ============================================
' CONSTANTES POUR LES CHEMINS (NOUVEAU)
' ============================================
Public Const DOSSIER_RAPPORTS As String = "\Rapports_CTC\"

' ============================================
' CONSTANTES MAGIQUES (NOUVEAU)
' ============================================
Public Const LIGNE_DEBUT_MANOEUVRES As Long = 2
Public Const LIGNE_FIN_MANOEUVRES_ZONE1 As Long = 12
Public Const LIGNE_FIN_MANOEUVRES_ZONE2 As Long = 28
Public Const LIGNE_DEBUT_TRAVAUX As Long = 13
Public Const MAX_ITERATIONS_SECURITE As Long = 1000

' ============================================
' VARIABLES GLOBALES (EXISTANTES - NE PAS MODIFIER)
' ============================================
' Ajoutez ici vos variables globales existantes
' Par exemple:
Public AgentConnecte As String
Public IsQuitting As Boolean

' ============================================
' NOUVELLES FONCTIONS HELPERS (BUG #8)
' ============================================

' Obtient une référence sécurisée à une feuille
Public Function ObtenirFeuilleSecurisee(nomFeuille As String) As Worksheet
    On Error Resume Next
    Set ObtenirFeuilleSecurisee = ThisWorkbook.Worksheets(nomFeuille)

    If Err.Number <> 0 Then
        MsgBox "ERREUR CRITIQUE: Feuille '" & nomFeuille & "' introuvable." & vbCrLf & _
               "L'application ne peut pas continuer." & vbCrLf & vbCrLf & _
               "Vérifiez que la feuille existe et n'a pas été renommée.", _
               vbCritical, "Feuille Manquante"
        Set ObtenirFeuilleSecurisee = Nothing
        Err.Clear
    End If
    On Error GoTo 0
End Function

' Vérifie au démarrage que toutes les feuilles essentielles existent
Public Function VerifierStructureFeuilles() As Boolean
    VerifierStructureFeuilles = True

    Dim feuilles() As String
    feuilles = Split( _
        NOM_FEUILLE_POV & "," & _
        NOM_FEUILLE_DATA & "," & _
        NOM_FEUILLE_VOIES & "," & _
        NOM_FEUILLE_ITINERAIRES & "," & _
        NOM_FEUILLE_DEPENDANCES & "," & _
        NOM_FEUILLE_AGENTS, _
        "," _
    )

    Dim i As Long
    Dim nomFeuille As String
    Dim feuille As Worksheet
    Dim feuillesManquantes As String

    feuillesManquantes = ""

    For i = LBound(feuilles) To UBound(feuilles)
        nomFeuille = Trim(feuilles(i))

        On Error Resume Next
        Set feuille = ThisWorkbook.Worksheets(nomFeuille)

        If Err.Number <> 0 Or feuille Is Nothing Then
            feuillesManquantes = feuillesManquantes & "- " & nomFeuille & vbCrLf
            VerifierStructureFeuilles = False
        End If

        Set feuille = Nothing
        Err.Clear
        On Error GoTo 0
    Next i

    If Not VerifierStructureFeuilles Then
        MsgBox "ERREUR CRITIQUE DE STRUCTURE" & vbCrLf & vbCrLf & _
               "Les feuilles suivantes sont manquantes:" & vbCrLf & vbCrLf & _
               feuillesManquantes & vbCrLf & _
               "L'application ne peut pas démarrer." & vbCrLf & vbCrLf & _
               "Restaurez le fichier depuis une sauvegarde.", _
               vbCritical, "Structure Invalide"
    End If
End Function

' ============================================
' (CONSERVEZ ICI TOUTES VOS AUTRES FONCTIONS EXISTANTES)
' ============================================
```

### ✅ Vérification Module 1

Après avoir collé le code:

1. **Menu**: Debug > Compile VBAProject
2. **Résultat attendu**: Aucune erreur
3. **Si erreur**: Vérifiez que vous avez bien copié TOUT le code
4. **Sauvegarder**: Ctrl+S

✅ **Module 1 terminé!**

---

## MODULE 2 - MOD_SESSION (BUG #4 - CRITIQUE)

### Localisation

1. Chercher "Mod_Session" dans l'explorateur de projet
2. Si n'existe pas, chercher une fonction "CreerRapportPDF" dans d'autres modules

### Instructions

**ATTENTION**: Ne remplacez PAS tout le module, seulement la fonction concernée!

### Trouver la fonction à modifier

1. Dans le module, appuyez sur **Ctrl+F** (Rechercher)
2. Tapez: `CreerRapportPDF`
3. Cliquez sur "Suivant"

### Code à REMPLACER

**Cherchez ces lignes** (environ ligne 2980-2990):

```vba
Dim cheminBase As String
cheminBase = "C:\Users\pgaye\OneDrive - Keolis\Bureau\personnel PKG\GOSMR\SMR Pilot\SMRCC\Rapports_CTC\"

On Error Resume Next
If Dir(cheminBase, vbDirectory) = "" Then MkDir cheminBase
On Error GoTo 0
```

### Code CORRIGÉ à mettre à la place

```vba
' === CORRECTION BUG #4: Chemin portable ===
Dim cheminBase As String
Dim cheminComplet As String

' Construction du chemin relatif au fichier Excel
cheminBase = ThisWorkbook.Path & Mod_Global.DOSSIER_RAPPORTS

' Vérification et création du dossier avec gestion d'erreur
If Dir(cheminBase, vbDirectory) = "" Then
    On Error Resume Next
    MkDir cheminBase

    If Err.Number <> 0 Then
        MsgBox "ERREUR: Impossible de créer le dossier de rapports." & vbCrLf & _
               "Chemin: " & cheminBase & vbCrLf & _
               "Erreur: " & Err.Description & vbCrLf & vbCrLf & _
               "Vérifiez les permissions du dossier.", vbCritical, "Création Dossier"
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
End If
```

### ✅ Vérification Module 2

1. **Compiler**: Debug > Compile VBAProject
2. **Sauvegarder**: Ctrl+S

✅ **Bug #4 corrigé! L'application fonctionne maintenant sur n'importe quelle machine.**

---

## MODULE 3 - USERFORM_CREERITINERAIRE (BUGS #1 et #2)

### Localisation

1. Dans l'explorateur de projet, chercher "UserForm_CreerItineraire"
2. Double-cliquer dessus
3. Clic droit sur le formulaire > Afficher le code

### CORRECTION #1 - Bug CLng (ligne ~344)

**Chercher** (Ctrl+F): `Génération de l'ID`

**Code ORIGINAL à remplacer**:

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

**Code CORRIGÉ**:

```vba
' === CORRECTION BUG #1: Conversion CLng sécurisée ===
Dim newID As String
Dim idNum As Long
Dim tempVal As String

If prochaineLigne = 2 Then
    idNum = 1
Else
    ' Extraction sécurisée de la partie numérique
    tempVal = Trim(Mid(wsPOV.Cells(prochaineLigne - 1, "Q").Value, 4))

    ' Validation avant conversion
    If IsNumeric(tempVal) And tempVal <> "" Then
        idNum = CLng(tempVal) + 1
    Else
        ' Valeur par défaut si la cellule précédente est invalide
        MsgBox "Attention: L'ID précédent est invalide. Réinitialisation à 1.", vbExclamation
        idNum = 1
    End If
End If

newID = "MAN" & Format(idNum, "00")
```

### CORRECTION #2 - Bug Boucle While (ligne ~330)

**Chercher** (Ctrl+F): `Trouver la prochaine ligne`

**Code ORIGINAL à remplacer**:

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

**Code CORRIGÉ**:

```vba
' === CORRECTION BUG #2: Boucle While sécurisée ===
Dim prochaineLigne As Long
Dim compteurSecurite As Long

prochaineLigne = Mod_Global.LIGNE_DEBUT_MANOEUVRES
compteurSecurite = 0

While wsPOV.Cells(prochaineLigne, "Q").Value <> "" And compteurSecurite < Mod_Global.MAX_ITERATIONS_SECURITE
    prochaineLigne = prochaineLigne + 1
    compteurSecurite = compteurSecurite + 1

    ' Vérification des limites de zone
    If prochaineLigne = Mod_Global.LIGNE_FIN_MANOEUVRES_ZONE1 Or _
       prochaineLigne = Mod_Global.LIGNE_FIN_MANOEUVRES_ZONE2 Or _
       compteurSecurite >= Mod_Global.MAX_ITERATIONS_SECURITE Then

        MsgBox "ERREUR: Le tableau des manœuvres semble plein." & vbCrLf & _
               "Ligne atteinte: " & prochaineLigne & vbCrLf & _
               "Itérations: " & compteurSecurite, vbCritical, "Tableau Plein"
        Exit Sub
    End If
Wend

' Vérification supplémentaire de sécurité
If compteurSecurite >= Mod_Global.MAX_ITERATIONS_SECURITE Then
    MsgBox "ERREUR CRITIQUE: Boucle infinie détectée." & vbCrLf & _
           "Contactez l'administrateur.", vbCritical, "Erreur Système"
    Exit Sub
End If
```

### ✅ Vérification Module 3

1. **Compiler**: Debug > Compile VBAProject
2. **Sauvegarder**: Ctrl+S

✅ **Bugs #1 et #2 corrigés!**

---

## MODULE 4 - THISWORKBOOK (BUG #8)

### Localisation

1. Dans "Microsoft Excel Objets"
2. Double-cliquer sur "ThisWorkbook"

### Code à MODIFIER

**Chercher la fonction**: `Workbook_Open`

**Code ORIGINAL**:

```vba
Private Sub Workbook_Open()
    ' 1. Cacher l'application et réinitialiser
    Application.Visible = False
    Mod_Global.AgentConnecte = ""
    Mod_Global.IsQuitting = False

    ' 2. Afficher le formulaire de connexion
    UserForm_Connexion.Show

    ' 3. L'utilisateur a fermé le formulaire
    If Mod_Global.AgentConnecte = "" Then
        Unload UserForm_Connexion
        Exit Sub
    End If

    ' 4. CONNEXION RÉUSSIE
    Call Mod_LogiqueApp.LancerPriseDeService
End Sub
```

**Code CORRIGÉ** (REMPLACER COMPLÈTEMENT):

```vba
Private Sub Workbook_Open()
    ' === CORRECTION BUG #8: Vérification de structure ===
    If Not Mod_Global.VerifierStructureFeuilles() Then
        ' Les feuilles essentielles sont manquantes
        Application.Visible = True
        ThisWorkbook.Close SaveChanges:=False
        Exit Sub
    End If

    ' === CODE EXISTANT (inchangé) ===
    ' 1. Cacher l'application et réinitialiser
    Application.Visible = False
    Mod_Global.AgentConnecte = ""
    Mod_Global.IsQuitting = False

    ' 2. Afficher le formulaire de connexion
    UserForm_Connexion.Show

    ' 3. L'utilisateur a fermé le formulaire
    If Mod_Global.AgentConnecte = "" Then
        Unload UserForm_Connexion
        Exit Sub
    End If

    ' 4. CONNEXION RÉUSSIE
    Call Mod_LogiqueApp.LancerPriseDeService
End Sub
```

### ✅ Vérification Module 4

1. **Compiler**: Debug > Compile VBAProject
2. **Sauvegarder**: Ctrl+S

✅ **Bug #8 corrigé! L'application vérifie maintenant la structure au démarrage.**

---

## MODULE 5 - GESTIONAFFICHAGE (BUG #5)

### Localisation

1. Chercher "GestionAffichage" dans les modules
2. Double-cliquer dessus

### Instructions

**AJOUTER** ces deux nouvelles fonctions à la fin du module (avant `End Sub` final s'il y en a un):

### Code à AJOUTER

```vba
' ===============================================
' NOUVELLES FONCTIONS SÉCURISÉES (BUG #5)
' Date ajout: 2025-11-15
' ===============================================

' Change la couleur d'une voie avec gestion d'erreur robuste
Public Function ChangerCouleurVoieSecurisee(nomVoie As String, couleur As Long) As Boolean
    ChangerCouleurVoieSecurisee = False

    Dim wsPOV As Worksheet
    Dim voieShape As Shape
    Dim nomNettoye As String

    nomNettoye = Trim(nomVoie)
    If nomNettoye = "" Then Exit Function

    On Error Resume Next
    Set wsPOV = ThisWorkbook.Worksheets(Mod_Global.NOM_FEUILLE_POV)
    If Err.Number <> 0 Then
        Debug.Print "ERREUR: Feuille POV introuvable"
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    ' Tentative d'accès au Shape
    Set voieShape = wsPOV.Shapes(nomNettoye)
    If Err.Number <> 0 Then
        Debug.Print "Avertissement: Shape '" & nomNettoye & "' introuvable"
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    ' Si on arrive ici, le Shape existe
    voieShape.Fill.ForeColor.RGB = couleur
    ChangerCouleurVoieSecurisee = True
End Function

' Affiche du texte sur une voie avec gestion d'erreur robuste
Public Function AfficherTexteSurVoieSecurise(nomVoie As String, texte As String) As Boolean
    AfficherTexteSurVoieSecurise = False

    Dim wsPOV As Worksheet
    Dim voieShape As Shape
    Dim nomNettoye As String

    nomNettoye = Trim(nomVoie)
    If nomNettoye = "" Then Exit Function

    On Error Resume Next
    Set wsPOV = ThisWorkbook.Worksheets(Mod_Global.NOM_FEUILLE_POV)
    If Err.Number <> 0 Then
        Debug.Print "ERREUR: Feuille POV introuvable"
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    Set voieShape = wsPOV.Shapes(nomNettoye)
    If Err.Number <> 0 Then
        Debug.Print "Avertissement: Shape '" & nomNettoye & "' introuvable"
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    ' Modification du texte
    With voieShape.TextFrame.Characters
        .Text = texte
        .Font.Bold = msoTrue
        .Font.Fill.ForeColor.RGB = vbBlack
    End With

    AfficherTexteSurVoieSecurise = True
End Function
```

### ✅ Vérification Module 5

1. **Compiler**: Debug > Compile VBAProject
2. **Sauvegarder**: Ctrl+S

✅ **Bug #5: Fonctions sécurisées ajoutées!**

**Note**: Pour utiliser ces fonctions dans le code existant, vous devrez remplacer les appels directs aux Shapes. Cela sera fait dans le Module 6.

---

## MODULE 6 - MANOEUVRES (BUG #3)

### Localisation

1. Chercher "Manoeuvres" dans les modules
2. Double-cliquer dessus

### Instructions - Bug #3

**ATTENTION**: Ce module nécessite plusieurs modifications. Nous allons chercher tous les `On Error Resume Next` et ajouter `On Error GoTo 0`.

### Méthode de correction

1. Appuyez sur **Ctrl+F**
2. Tapez: `On Error Resume Next`
3. Cliquez sur "Rechercher dans le module en cours"
4. Pour **chaque occurrence trouvée**:

**Si vous voyez**:

```vba
On Error Resume Next
ws.Shapes(nomElement).Fill.ForeColor.RGB = RGB(255, 0, 0)
' Pas de "On Error GoTo 0"
```

**Ajoutez après**:

```vba
On Error Resume Next
ws.Shapes(nomElement).Fill.ForeColor.RGB = RGB(255, 0, 0)
On Error GoTo 0  ' <<< AJOUTER CETTE LIGNE
```

### Exemples de modifications typiques

**Exemple 1** (ligne ~731):

```vba
' AVANT:
For Each elem1 In elementsChemin1
    GestionDonnees.ModifierEtatElement Trim(CStr(elem1)), "Verrouille"
    ws.Shapes(Trim(CStr(elem1))).TextFrame.Characters.Font.Color = vbBlue
Next elem1

' APRÈS:
For Each elem1 In elementsChemin1
    GestionDonnees.ModifierEtatElement Trim(CStr(elem1)), "Verrouille"
    On Error Resume Next
    ws.Shapes(Trim(CStr(elem1))).TextFrame.Characters.Font.Color = vbBlue
    On Error GoTo 0  ' <<< AJOUTÉ
Next elem1
```

**Exemple 2** (ligne ~802):

```vba
' AVANT:
On Error Resume Next
ws.Shapes(nomAiguille).TextFrame.Characters.Font.Color = vbBlack

' APRÈS:
On Error Resume Next
ws.Shapes(nomAiguille).TextFrame.Characters.Font.Color = vbBlack
On Error GoTo 0  ' <<< AJOUTÉ
```

### Compteur

Vous devriez trouver environ **8-10 occurrences** dans ce module.

### ✅ Vérification Module 6

1. **Vérifier**: Chaque `On Error Resume Next` a son `On Error GoTo 0`
2. **Compiler**: Debug > Compile VBAProject
3. **Sauvegarder**: Ctrl+S

✅ **Bug #3 corrigé!**

---

## MODULE 7 - GESTIONSECURITE (BUG #7)

### Localisation

1. Chercher "GestionSecurite" dans les modules
2. Double-cliquer dessus

### Instructions

**Chercher la fonction**: `ValiderCheminLibre`

### Code COMPLET de la fonction à REMPLACER

**REMPLACER TOUTE la fonction** par cette version corrigée:

```vba
Function ValiderCheminLibre(chemin As String) As Boolean
    ' RÔLE: Vérifie si un chemin est libre (CORRIGÉ - Bug #7)

    ValiderCheminLibre = True  ' OK par défaut

    ' === VALIDATION ROBUSTE DU CHEMIN (CORRIGÉ) ===
    Dim cheminNettoye As String
    cheminNettoye = Trim(chemin)

    ' Vérification 1: Chaîne vide ou erreur
    If cheminNettoye = "" Or cheminNettoye = "ERREUR: CHEMIN NON TROUVÉ" Or cheminNettoye = "ERREUR" Then
        Debug.Print "ValiderCheminLibre: Chemin invalide ou vide"
        Exit Function
    End If

    ' Vérification 2: Au moins un élément
    If InStr(cheminNettoye, ",") = 0 And Len(cheminNettoye) = 0 Then
        Debug.Print "ValiderCheminLibre: Aucun élément dans le chemin"
        Exit Function
    End If

    ' Split sécurisé
    Dim elementsChemin() As String
    elementsChemin = Split(cheminNettoye, ",")

    ' Vérification 3: Tableau non vide
    If UBound(elementsChemin) < LBound(elementsChemin) Then
        Debug.Print "ValiderCheminLibre: Tableau d'éléments vide après Split"
        Exit Function
    End If

    ' Vérification 4: Au moins 2 éléments pour valider un chemin
    If UBound(elementsChemin) < LBound(elementsChemin) + 1 Then
        ' Chemin avec un seul élément = origine = destination
        ' C'est valide, on sort
        Exit Function
    End If

    ' === BOUCLE DE VALIDATION (CODE EXISTANT AMÉLIORÉ) ===
    Dim i As Long
    Dim nomNettoye As String
    Dim etatElement As String
    Dim nbProtections As Long

    ' On commence à partir du DEUXIÈME élément
    For i = LBound(elementsChemin) + 1 To UBound(elementsChemin)
        nomNettoye = Trim(CStr(elementsChemin(i)))

        ' Skip les éléments vides (par sécurité)
        If nomNettoye = "" Then
            Debug.Print "ValiderCheminLibre: Élément vide ignoré à l'index " & i
            GoTo NextElement
        End If

        ' Vérification 1: État "Libre"
        etatElement = GestionDonnees.LireEtatElement(nomNettoye)
        If etatElement <> "Libre" Then
            MsgBox "VALIDATION ÉCHOUÉE (Chemin)" & vbCrLf & vbCrLf & _
                     "L'itinéraire est impossible car l'élément '" & nomNettoye & "' n'est pas libre." & vbCrLf & _
                     "(État actuel: " & etatElement & ")", _
                     vbCritical, "Sécurité Itinéraire"
            ValiderCheminLibre = False
            Exit Function
        End If

        ' Vérification 2: Aucune protection
        nbProtections = GestionDonnees.CompterProtections(nomNettoye)
        If nbProtections > 0 Then
            MsgBox "VALIDATION ÉCHOUÉE (Chemin)" & vbCrLf & vbCrLf & _
                   "L'itinéraire est impossible car l'élément '" & nomNettoye & "' est indisponible." & vbCrLf & _
                   "(Cause: Protection ou Dérangement actif)", _
                   vbCritical, "Sécurité Itinéraire"
            ValiderCheminLibre = False
            Exit Function
        End If

NextElement:
    Next i
End Function
```

### ✅ Vérification Module 7

1. **Compiler**: Debug > Compile VBAProject
2. **Sauvegarder**: Ctrl+S

✅ **Bug #7 corrigé!**

---

## MODULE 8 - TRAVAUX (BUG #6)

### Localisation

1. Chercher "Travaux" dans les modules (ou le code pourrait être ailleurs)
2. Double-cliquer dessus

### Instructions

**Chercher la fonction**: `TerminerTravaux_Selectionnes`

**Chercher dans cette fonction** (Ctrl+F): `Dim agentValidateur`

**Code ORIGINAL**:

```vba
Dim agentValidateur As String
agentValidateur = wsPOV.Cells(ligne, "W").Value
```

**Code CORRIGÉ**:

```vba
' === CORRECTION BUG #6: Renommage variable locale ===
Dim agentValidateurLocal As String
agentValidateurLocal = Trim(wsPOV.Cells(ligne, "W").Value)

' Si la variable publique doit aussi être mise à jour:
agentValidateur = agentValidateurLocal
```

**Ensuite**, dans le reste de la fonction, **remplacer** toutes les utilisations de `agentValidateur` par `agentValidateurLocal`.

### ✅ Vérification Module 8

1. **Compiler**: Debug > Compile VBAProject
2. **Sauvegarder**: Ctrl+S

✅ **Bug #6 corrigé!**

---

## ✅ TESTS FINAUX

Après avoir modifié TOUS les modules:

### Test 1: Compilation

```
1. Dans l'éditeur VBA: Debug > Compile VBAProject
2. Résultat attendu: "Compilation réussie" (aucun message)
3. Si erreur: Noter le module et la ligne, corriger
```

### Test 2: Démarrage

```
1. Fermer l'éditeur VBA (Alt+Q)
2. Fermer Excel complètement
3. Rouvrir "SMR_PILOT 2F.xlsm"
4. Activer les macros
5. Résultat attendu: Aucune erreur, formulaire de connexion s'affiche
```

### Test 3: Connexion

```
1. Se connecter avec vos identifiants
2. Résultat attendu: Accès au tableau de bord
```

### Test 4: Manœuvre

```
1. Menu > Créer Manœuvre
2. Sélectionner une rame
3. Sélectionner une destination
4. Cliquer "Préparer"
5. Résultat attendu: Manœuvre créée avec ID correct (ex: MAN01)
```

### Test 5: Rapport PDF (TEST CRITIQUE)

```
1. Menu > Créer Rapport
2. Résultat attendu:
   - Un dossier "Rapports_CTC" est créé à côté du fichier Excel
   - Un PDF est créé dans ce dossier
   - Nom du fichier: Rapport_CTC_2025-11-15_XX-XX-XX.pdf
```

**Si TOUS les tests passent**: 🎉 **MIGRATION RÉUSSIE!**

---

## 🐛 DÉPANNAGE

### Erreur: "Compile error: Variable not defined"

**Cause**: Vous avez oublié de déclarer une variable ou de modifier Mod_Global

**Solution**:
1. Vérifier que le Module 1 (Mod_Global) a bien été modifié
2. Vérifier que toutes les constantes sont déclarées

### Erreur: "Compile error: Sub or Function not defined"

**Cause**: Une fonction appelée n'existe pas

**Solution**:
1. Vérifier que le Module 5 (GestionAffichage) a bien les nouvelles fonctions
2. Vérifier que le Module 1 (Mod_Global) a les fonctions helpers

### Erreur: "Run-time error '9': Subscript out of range"

**Cause**: Une feuille est introuvable

**Solution**:
1. Vérifier que toutes les feuilles existent (POV, Data, Voies, etc.)
2. Vérifier que le Module 4 (ThisWorkbook) a bien été modifié

### L'application ne démarre pas

**Solution**:
1. Fermer Excel
2. Restaurer depuis le backup
3. Recommencer module par module
4. Compiler après CHAQUE module

---

## 📋 CHECKLIST FINALE

Cochez au fur et à mesure:

### Modules modifiés

- [ ] Module 1: Mod_Global (constantes et helpers)
- [ ] Module 2: Mod_Session (chemin PDF)
- [ ] Module 3: UserForm_CreerItineraire (CLng et While)
- [ ] Module 4: ThisWorkbook (vérification feuilles)
- [ ] Module 5: GestionAffichage (fonctions sécurisées)
- [ ] Module 6: Manoeuvres (On Error GoTo 0)
- [ ] Module 7: GestionSecurite (Split validé)
- [ ] Module 8: Travaux (variable renommée)

### Tests

- [ ] Compilation réussie
- [ ] Démarrage OK
- [ ] Connexion OK
- [ ] Création manœuvre OK
- [ ] Rapport PDF OK (dossier créé à côté du fichier Excel)

### Sauvegarde

- [ ] Fichier sauvegardé (Ctrl+S)
- [ ] Backup conservé en lieu sûr

**Si tout est coché**: ✅ **TERMINÉ!**

---

## 📞 BESOIN D'AIDE?

**Si vous êtes bloqué**:

1. **Vérifier** que le backup existe
2. **Restaurer** le backup si nécessaire
3. **Recommencer** en suivant les étapes une par une
4. **Compiler** après CHAQUE modification

**Rappel important**: Ne jamais modifier plusieurs modules à la fois. Toujours compiler et sauvegarder après chaque module.

---

**Temps total**: ~1h30
**Niveau**: Intermédiaire (copier-coller principalement)
**Résultat**: Application robuste et portable

Bon courage! 🚀
