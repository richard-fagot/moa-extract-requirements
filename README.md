# moa-extract-requirements
Script VBA pour extraires les spécifications en conservant la hiérarchie des titres qui sont représentatifs du contenu fonctionnel décrit par les exigences.

Voir le fichier `exemple.docx` pour un exemple de spécification avec des ecigences conforme à ce qu'attend le script.

Voir le fichier `SFG - Exigences.xlsx` pour un exemple d'extraction basée sur le fichier précedent.

## Prérequis
 - Word 2013
 - Excel 2013

## Installation
1. Activer l'onglet développeur dans Word : **Fichier/Options/Personnaliser le ruban**. cocher **Développeur** dans la zone de droite ;
1. Ouvrir VBA : **Développeur/Visual Basic** ;
1. Dans la fenêtre qui s'ouvre double-cliquer sur **Normal/ThisDocument** et copier-coller le contenu du fichier `ThisDocument.cls` à partir de `Sub Main()` OU ;
1. Faire **Fichier/Importer** et importer le fichier `ThisDocument.cls`.

## Exécution
1. Mettre le curseur de la souris dans `Sub Main()` ;
1. Faire **Exécution/Exécuter** ou `F5`ou cliquer sur le symbole "lecture".

## Attention
Le script peut être très long si le document est volumineux.
