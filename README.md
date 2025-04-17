# PURH TypoCleaner

Un module VBA pour Microsoft Word destiné aux **Presses universitaires de Rouen et du Havre (PURH)**.  
Il automatise le toilettage typographique : guillemets, apostrophes, ligatures, ponctuation, siècles en petites capitales, notes de bas de page…

---

## 📋 Fonctionnalités

- Nettoyage des doubles paragraphes, espaces et tabulations  
- Remplacement :
  - des guillemets droits (`"`) et des “smart quotes” anglaises (`“”`) par des chevrons français (`« … »`) avec espaces insécables  
  - des apostrophes droites (`'`) par des apostrophes typographiques (`’`)  
  - des triples points (`...`) par points de suspension (`…`)  
  - des doubles tirets (`--`) par tirets cadratins (`—`)  
  - insertion d’une espace insécable avant `: ; ! ?`  
  - des ligatures (`oeuvre`, `voeu[x]`, `soeur[s]`, `oeuf[s]`) en `œ`, `vœu[x]`, `sœur[s]`, `œuf[s]`, toutes variantes singulier/pluriel et minuscule/majuscule  
- Mise en petites capitales + exposant des siècles **Iᵉ → XXIᵉ**  
- Traitement complet **dans le corps** et **dans les notes de bas de page** :  
  - ajout d’un point final si manquant  
  - insécable après `p.` pour numéros de page  

---

## 🚀 Installation

1. Ouvrez Word et appuyez sur **Alt + F11** pour ouvrir l’éditeur VBA.  
2. Dans le projet **Normal** (ou votre modèle `.dotm`), `Insertion > Module`.  
3. Copiez‑collez les deux routines suivantes :

   ```vba
   Sub Cleaner()
       Dim doc As Document, fn As Footnote
       Set doc = ActiveDocument
       Application.ScreenUpdating = False

       ' 1) Tous les nettoyages dans le corps…
       Call CleanTypoRange(doc.Content)

       ' 2) …et dans chaque note de bas de page
       For Each fn In doc.Footnotes
           Call CleanTypoRange(fn.Range)
           ' point final et insécable "p. X"
           With fn.Range.Find
               .ClearFormatting: .Replacement.ClearFormatting
               .MatchWildcards = False: .Wrap = wdFindContinue
               .Text = "p. ": .Replacement.Text = "p." & Chr(160)
               .Execute Replace:=wdReplaceAll
               If Right(Trim(fn.Range.Text),1) <> "." Then fn.Range.InsertAfter "."
           End With
       Next fn

       Application.ScreenUpdating = True
       MsgBox "Toilettage typographique PURH terminé !", vbInformation
   End Sub

   Sub CleanTypoRange(rng As Range)
       Dim ponct As Variant, i As Integer, toggle As Boolean
       Dim fullRng As Range

       ' 0) Nettoyage de base
       With rng.Find
           .ClearFormatting: .Replacement.ClearFormatting
           .Wrap = wdFindStop: .Forward = True
           .Text = "^p^p":   .Replacement.Text = "^p"
           Do While .Execute(Replace:=wdReplaceAll): Loop
           .Text = "  ":     .Replacement.Text = " "
           Do While .Execute(Replace:=wdReplaceAll): Loop
           .Text = "^t":     .Replacement.Text = ""
           .Execute Replace:=wdReplaceAll
       End With

       ' 1a) “smart quotes” anglaises → guillemets français + insécables
       With rng.Find
           .ClearFormatting: .MatchWildcards = False: .Wrap = wdFindContinue
           .Text = ChrW(&H201C): .Replacement.Text = "«" & Chr(160)
           .Execute Replace:=wdReplaceAll
           .Text = ChrW(&H201D): .Replacement.Text = Chr(160) & "»"
           .Execute Replace:=wdReplaceAll
       End With

       ' 1b) guillemets droits → chevrons alternés
       toggle = True
       With rng.Find
           .ClearFormatting: .Replacement.ClearFormatting
           .Text = Chr(34): .Replacement.Text = ""
           .MatchWildcards = False: .Wrap = wdFindStop
       End With
       Do While rng.Find.Execute
           rng
