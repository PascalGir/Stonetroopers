# Les StoneTroopers — Planning

Page web (lecture seule) des matchs de curling et des disponibilités de l'équipe.

En ligne : https://pascalgir.github.io/stonetroopers/

## Mettre à jour le planning

1. Modifie l'Excel comme d'habitude.
2. Lance : `python3 maj_planning.py Planning26-27_Curling-Stonetroopes.xlsx` (nécessite `pip install openpyxl`). Ça régénère `data.js`.
3. Sur GitHub, remplace `data.js` par le nouveau fichier (Add file → Upload files → Commit). La page se met à jour en 1 à 2 minutes.

Tu peux aussi modifier `data.js` directement sur GitHub (icône crayon) : chaque match a sa date, son heure, sa piste, son adversaire et la réponse de chaque joueur (`Oui`, `Non`, `Incertain` ou vide). Les vacances se mettent dans `vacances` sous la forme `{"joueur": "Pascal", "debut": "2026-12-20", "fin": "2027-01-03"}`.
