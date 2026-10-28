# Feuille de route

## Fait
- Numéro de version unique (`AppBrand.Version` dans `src/AppSettings.cs`), affiché dans Paramètres et dans les Applications de Windows.
- Installation, désinstallation et mise à jour par l'exe lui-même (`src/Installation.cs`).
- Mise à jour depuis l'application (`src/Updater.cs`) : recherche de la dernière Release GitHub au démarrage, bouton « Mise à jour disponible » dans la fenêtre, bouton dans Paramètres, téléchargement contrôlé (taille, SHA-256, numéro de version), remplacement après fermeture, relance.

## Publier une nouvelle version
Tout est automatique : GitHub compile et publie (`.github/workflows/release.yml`).
1. Corriger ou ajouter, puis augmenter `AppBrand.Version` dans `src/AppSettings.cs` (ex. `1.0.2`).
2. Vérifier en local : `src\tests\LANCER_TESTS.ps1`.
3. Commiter et pousser sur `main`.
4. Poser l'étiquette et la pousser : `git tag v1.0.2` puis `git push origin v1.0.2`.
5. GitHub compile `Korecto.exe`, vérifie que l'étiquette correspond à la version de l'exe (sinon la Release est annulée), et publie la Release. Les Korecto installés proposent la mise à jour à leur prochain démarrage.

## Idées
- Intégration au-delà d'Outlook classique : raccourci global avec presse-papiers, champs de texte d'autres applications, extension navigateur pour les messageries web (clic droit « Corriger avec Korecto » ; un premier essai existe dans l'historique git, à publier sur le Chrome Web Store pour une installation en un clic), complément pour le nouvel Outlook.
- Signature de code de l'exe (supprime l'avertissement SmartScreen ; environ 100 à 400 € par an).
- Licence du dépôt (à choisir).

## À valider en conditions réelles
- Lecture et application dans un vrai Outlook : signature, listes, tableaux, réponses intégrées.
- Une mise à jour réelle d'une version à la suivante, avec les fenêtres à l'écran.
- Essai réel de Luna et de Haiku, consommation par correction.
