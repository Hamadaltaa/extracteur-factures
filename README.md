# Extracteur de factures

Transforme un dossier de factures PDF en un fichier Excel propre (fournisseur,
date, HT, TVA, TTC), et signale automatiquement les lignes douteuses.

## Pourquoi
La saisie manuelle de factures prend du temps et provoque des erreurs.
Cet outil fait la saisie, et ne laisse à l'humain que les cas à vérifier.

## Installation
```
pip install -r requirements.txt
export GEMINI_API_KEY="votre_clé"         # Windows : set GEMINI_API_KEY=votre_clé
```

## Utilisation
Placez vos PDF dans un dossier `factures/`, puis :
```
python extract.py factures resultat.xlsx
```

## Contrôles automatiques
- Champ manquant (fournisseur, date, montants)
- HT + TVA différent du TTC
- PDF scanné sans texte lisible

## Limites actuelles
PDF avec texte uniquement. Photos de reçus, interface web et multi-devises :
prévus pour une version 2.
