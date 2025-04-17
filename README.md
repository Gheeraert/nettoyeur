# PURH TypoCleaner

Un module VBA pour Microsoft Word conçu pour les **Presses universitaires de Rouen et du Havre (PURH)**.  
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
  - ajout d’un point final si manquant dans les notes
  - insécable après `p.` pour numéros de page  

---

## 🚀 Installation

1. Ouvrez Word et appuyez sur **Alt + F11** pour ouvrir l’éditeur VBA.  
2. Dans le projet **Normal** (ou votre modèle `.dotm`), `Insertion > Module`.  
3. Copiez‑collez les deux routines du modèle.
