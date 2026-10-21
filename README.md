# Les StoneTroopers — Planning

Matchs de curling et disponibilités de l'équipe : https://pascalgir.github.io/stonetroopers/

Chaque joueur choisit son nom dans **Je suis**, puis clique sur ses cases pour changer sa réponse (vide → Oui → Incertain → Non). Les données sont stockées dans un Google Sheet.

## Mise en place (une seule fois)

1. Google Drive → **Nouveau → Importer un fichier** → `StoneTroopers_Base.xlsx`. Ouvre-le, puis **Fichier → Enregistrer au format Google Sheets**.
2. Dans ce Google Sheet : **Extensions → Apps Script**. Efface le contenu, colle tout `Code.gs`, puis 💾.
3. **Déployer → Nouveau déploiement** → ⚙️ **Application Web** — Exécuter en tant que : **Moi** · Qui a accès : **Tout le monde** → **Déployer** → **Autoriser l'accès** (si Google affiche « application non validée » : Paramètres avancés → Accéder au projet).
4. Copie l'URL qui finit par `/exec` et colle-la dans `config.js` (entre les guillemets de `API_URL`).

## Le Google Sheet

- **Matchs** : ID, Date, Heure, Piste, Adversaire, Score nous, Score eux, Notes, puis une colonne par joueur.
- **Joueurs** : Nom, Actif (Non = masqué sur la page), Poste.
- **Vacances** : ID, Joueur, Début, Fin, Commentaire.

On peut le modifier à la main ; garder les noms d'onglets et de colonnes.

Après une modification de `Code.gs` : **Déployer → Gérer les déploiements → ✏️ → Nouvelle version** (l'URL ne change pas).

`data.js` ne sert que de secours si le Google Sheet ne répond pas.
