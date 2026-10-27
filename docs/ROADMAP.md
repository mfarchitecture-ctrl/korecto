# Feuille de route

## Fait
- Numéro de version unique (`AppBrand.Version` dans `src/AppSettings.cs`), affiché dans Paramètres et dans les Applications de Windows.
- Installation, désinstallation et mise à jour par l'exe lui-même (`src/Installation.cs`).
- Mise à jour depuis l'application (`src/Updater.cs`) : recherche de la dernière Release GitHub au démarrage, notification, bouton dans Paramètres, téléchargement contrôlé (taille, SHA-256, numéro de version), remplacement après fermeture, relance.

## Publier une nouvelle version
1. Corriger ou ajouter, puis augmenter `AppBrand.Version` (ex. `1.0.2`).
2. `src\tests\LANCER_TESTS.ps1`, puis `src\CONSTRUIRE.ps1`.
3. Sur GitHub : nouvelle Release, étiquette `v1.0.2` (même numéro que dans l'exe, sinon la mise à jour est refusée), joindre `Korecto.exe`.
4. Les Korecto installés proposent la mise à jour à leur prochain démarrage.

## Idées
- Intégration au-delà d'Outlook classique : raccourci global avec presse-papiers, champs de texte d'autres applications, extension navigateur pour les messageries web, complément pour le nouvel Outlook.
- Signature de code de l'exe (supprime l'avertissement SmartScreen ; environ 100 à 400 € par an).
- Licence du dépôt (à choisir).

## À valider en conditions réelles
- Lecture et application dans un vrai Outlook : signature, listes, tableaux, réponses intégrées.
- Une mise à jour réelle d'une version à la suivante, avec les fenêtres à l'écran.
- Essai réel de Luna et de Haiku, consommation par correction.
