# Korecto : consignes de travail

## Rythme de publication
- Une mise à jour par semaine : les corrections s'accumulent sur la branche de travail, la release est décidée par Maxime après usage.
- Ne jamais publier (étiquette `vX.Y.Z`, mise sur `main` pour une release) sans demande explicite. Avant la release : numéro de version dans `src/AppSettings.cs`, ligne en tête de `src/NOUVEAUTES.txt`, nombre de contrôles dans `README.md` et `.github/workflows/tests.yml`.
- Les tests Windows se lancent par GitHub Actions (`tests.yml`, workflow_dispatch) : on ne peut pas compiler sur Linux.

## Liste des corrections en attente
Ne pas la répéter à chaque message. La donner seulement quand Maxime la demande, ou quand une correction est ajoutée ou que la release se prépare (une ligne par correction : description, commit, résultat des tests Windows).

## Corrections en attente de release (prochaine : 1.0.9)
Ajouter ici chaque correction après tests Windows, et vider la liste à chaque release.
- Sécurité : plus d'installation de Claude Code par PowerShell (seulement sa connexion), mise à jour refusée sans empreinte SHA-256, consignes toujours intégrées à l'exe, IA lancée dans le dossier de données. Commit f0dc35c, tests Windows 321/321.
