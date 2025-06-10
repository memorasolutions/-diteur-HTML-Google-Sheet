# Éditeur HTML pour Google Sheets

Ce script ajoute un éditeur WYSIWYG pour modifier le contenu HTML d'une cellule. L'éditeur s'ouvre toujours dans une boîte de dialogue modale pour offrir plus d'espace d'édition.

## Bouton d'accès direct

Vous pouvez insérer un dessin ou une image dans la feuille puis lui attribuer la fonction `editActiveCellDialog`. Ainsi, un simple clic sur cette image ouvrira l'éditeur pour la cellule active.

Pour automatiser l'ajout du bouton, exécutez la fonction `insertEditorButton` depuis l'éditeur Apps Script. Une petite icône sera placée en haut à gauche de la feuille et lancera `editActiveCellDialog` lorsqu'on clique dessus.
