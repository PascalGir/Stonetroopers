# Les StoneTroopers — Planning

Matchs de curling et disponibilités de l'équipe : https://pascalgir.github.io/stonetroopers/

Chaque joueur choisit son nom dans **Je suis**, puis clique sur ses cases pour changer sa réponse (vide → Oui → Incertain → Non). Les réponses sont enregistrées dans le Google Sheet de l'équipe.

## Relier la page au Google Sheet (une seule fois)

1. Sur Google Drive : **Nouveau → Importer un fichier** → `Planning26-27_Curling-Stonetroopes.xlsx`, puis ouvre-le et fais **Fichier → Enregistrer au format Google Sheets**.
2. Dans ce Google Sheet : **Extensions → Apps Script**. Efface le contenu, colle tout le fichier `Code.gs`, puis 💾 Enregistrer.
3. **Déployer → Nouveau déploiement** → roue dentée → **Application Web** :
   - Exécuter en tant que : **Moi**
   - Qui a accès : **Tout le monde**
   - **Déployer**, puis **Autoriser l'accès** avec ton compte Google (écran « Google n'a pas validé cette application » → Paramètres avancés → Accéder au projet).
4. Copie l'**URL de l'application Web** (elle se termine par `/exec`).
5. Sur GitHub, ouvre `config.js` → crayon ✏️ → colle l'URL entre les guillemets de `API_URL` → **Commit changes**.

Après 1 à 2 minutes, la page lit et écrit directement dans le Google Sheet. Tant que `API_URL` est vide, la page affiche `data.js` en lecture seule.

## Ajouter / modifier des matchs

Directement dans le Google Sheet (onglet Planning, colonnes A à E ; vacances dans l'onglet Vacances). La page se met à jour toute seule.

Si tu modifies `Code.gs` plus tard : **Déployer → Gérer les déploiements → ✏️ → Version : Nouvelle version** (l'URL reste la même).
