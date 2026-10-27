# Feuille de route : vers un logiciel propre

Objectif (Maxime, 27/10/2026) : numéro de version, mise à jour depuis l'application, le plus automatique possible. Rien de ceci n'est fait ; l'ordre ci-dessous évite de refaire du travail.

## 1. Numéro de version
- Une seule source de vérité : `[assembly: AssemblyVersion/AssemblyInformationalVersion]` dans `CorrecteurMail.cs`, lue par `CONSTRUIRE.ps1` et par l'installateur.
- Affichée dans Paramètres et dans l'entrée « Applications » de Windows (`DisplayVersion`).
- Étiquettes Git `v1.0.0`, `v1.1.0`… une par version livrée (versionnage sémantique).

## 2. Fabrication d'une livraison
- Script unique : compile, lance les tests, fabrique `Korecto.exe` de la version et son empreinte SHA-256. La Release ne contient que cet exe : il s'installe lui-même.
- Notes de version (`CHANGELOG.md`).

## 3. Diffusion (décision à prendre)
- Releases GitHub : exigent un dépôt public, ou un second dépôt public `korecto-releases` si le code reste privé.
- Alternative simple pour quelques proches : lien de partage vers le ZIP, sans détection de version.

## 4. Mise à jour dans l'application
- Étape A : vérification au lancement (et bouton « Rechercher une mise à jour ») de la dernière Release ; message « Version X disponible » avec lien. Aucun téléchargement automatique.
- Étape B : téléchargement, vérification SHA-256, puis remplacement de l'exe installé (la logique d'installation et de mise à jour est déjà dans Installation.cs) après fermeture de l'appli, et relance. Réglages conservés.
- Contraintes : désactivable dans Paramètres, aucun envoi de données personnelles, échec silencieux sans réseau.

## 5. Confiance Windows
- L'exe n'est pas signé : SmartScreen avertit au premier lancement. Une signature de code (environ 100 à 400 € par an) supprime l'avertissement et est conseillée avant une diffusion large.

## 6. Qualité restante (indépendante de ce qui précède)
- Valider dans un vrai Outlook : aperçu, changements de brouillon, Appliquer avec signature et mise en forme, réponses intégrées.
- Essai réel de Luna et de Haiku ; mesure de la consommation par correction.
- Installation réelle chez un second utilisateur Windows.
