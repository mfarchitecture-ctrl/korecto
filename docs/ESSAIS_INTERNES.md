# Essayer Korecto sans publier ni mettre à jour la version installée

## Voir la version interne

Dans le dossier `dist`, double-cliquer sur `VOIR_KORECTO.cmd`.

La fenêtre porte la mention « TEST INTERNE ». Elle affiche un texte fictif et le bouton « Mise à jour disponible » simulé. Ce bouton ne télécharge rien dans ce mode. Le programme installé, ses réglages, son démarrage Windows et la Release GitHub ne sont pas remplacés. Les réglages de test sont conservés dans `dist/DONNEES_TEST`. Le choix IA existant est lu au premier lancement ; changer ce choix dans la fenêtre de test ne modifie pas les réglages installés.

Les corrections IA ne partent que sur clic et utilisent le compte du fournisseur choisi. Une correction réelle consomme les limites de ce compte, même en mode de test.

## Alt + K

Garder cette version interne ouverte. Dans une autre application, sélectionner du texte, puis appuyer sur Alt + K et relâcher les touches. La sélection arrive dans l’onglet « Texte d’origine », en mode Manuel, sans appel IA. Choisir ensuite Corriger ou Reformuler. Copier le résultat pour le recoller soi-même.

Le raccourci simule une copie de la sélection puis tente de rétablir le contenu précédent du presse-papiers. Il n’utilise pas Ctrl + A et n’effectue pas de collage automatique. L’application source doit accepter Ctrl + C. Certains champs protégés, applications exécutées comme administrateur ou raccourcis déjà enregistrés peuvent empêcher cette capture. Le menu du navigateur est une autre voie d’importation et ne passe pas par le presse-papiers.

En cas de conflit, fermer l’autre version de Korecto ou l’application qui réserve Alt + K. Fermer la fenêtre avec la croix masque Korecto ; utiliser son icône près de l’horloge → Quitter pour libérer le raccourci.

## Extension Chrome et Brave — installation locale de test

1. Dans Korecto, ouvrir les réglages puis cliquer sur « Extension du navigateur ». Ce clic identifie le navigateur par défaut de Windows, prépare les fichiers et la liaison locale, copie le chemin et ouvre sa page des extensions. Aucun second bouton ni choix de navigateur.
2. Dans le navigateur ouvert, activer « Mode développeur », puis « Charger l’extension non empaquetée ». Coller le chemin avec Ctrl + V dans la sélection de dossier, puis valider.
3. Sélectionner du texte sur une page : clic droit → « Corriger avec Korecto ». Le texte arrive dans Korecto sans appel IA ; Corriger ou Reformuler restent des choix explicites.

La version interne reconnaît Chrome, Brave, Microsoft Edge, Vivaldi et Chromium. Un navigateur par défaut incompatible ou non identifié est signalé, sans basculer automatiquement sur un autre. Firefox nécessite une version adaptée non encore fournie. Le chargement local reste une étape imposée par le navigateur pour une extension de test non publiée.
Le script `dist/ACTIVER_EXTENSION.cmd` et le dossier `extension` du dépôt restent une solution de test manuelle facultative. L’installation depuis les réglages prépare ses propres fichiers dans le dossier de données de la copie utilisée ; une liaison appartenant à un autre exécutable Korecto est conservée et signalée.
L’extension ne demande que les permissions de menu contextuel et de messagerie native. Elle ne demande aucun accès permanent aux sites et ne transmet pas l’URL de la page. Si la liaison échoue, son icône affiche « ! » ; ouvrir l’extension pour lire le message. Épingler l’extension dans la barre du navigateur facilite cet accès.

Pour retirer l’essai : supprimer l’extension depuis la page des extensions et exécuter `dist/DESACTIVER_EXTENSION.cmd`. Ce script ne retire que les clés pointant vers cette liaison de test. Ne pas déplacer le dossier de test après activation sans relancer le script d’activation.

## Vérifications réalisées

291 contrôles C# locaux réussis, dont import sans appel IA, rejet des réponses périmées, reformulation sur clic, limites de taille, échanges Windows entre processus et protocole de l’hôte navigateur. Six groupes de contrôles JavaScript réussis : permissions, menu sans envoi automatique, sélection Unicode seule, limites, erreurs et identité de l’extension.

La capture physique Alt + K dans les applications réelles et le chargement de l’extension dans Chrome/Brave restent à valider sur le poste. Les tests automatisés n’actionnent pas le vrai clavier et ne lisent pas de mails personnels.

Aucune intégration Outlook supplémentaire, aucun push, aucune nouvelle Release.
